# Outlook Assistant — 專案 Memory

> 輸出時間：2026/07/02 11:08 (Asia/Taipei)

---

## 一、用途與背景

Simon 是 **Outlook Assistant** 的唯一開發者，這是一個 VB.NET (.NET 10 WinForms) 應用程式，用來分析大型本地 Outlook PST 封存（約 350,000 封信、27 個 PST 檔），透過 Outlook LTSC 2021 COM interop (OOM) 與 Redemption RDO library 存取。專案是 partial class 架構，跨多個檔案：`Module_Outlook.vb`（實際是 `Partial Class Form1`，不是 VB Module——已確認的架構事實）、`Form1.vb`、`Form1_MainTab12.vb`、`Form1_Maintab56.vb`、`Module_SQLite2.vb`、`Module_Win32API.vb` 等。

**核心架構：** 分層 L1 (UI) → L2 (流程/分派) → L2.5 (cache proxy) → L3 (直接 COM/RDO)。快取層級：memory (`ConcurrentDictionary`) → SQLite/SSD (`OLAcache.db` + `OLAcacheMail.db`) → COM/RDO fallback。RDO (`_rdo2`，獨立 session) 是主要讀取路徑；OOM 是 fallback。`_rdo`（piggyback session）是舊版，正在淘汰中。`GetRdoStore` 刻意用一般 `Dictionary` 而非 `ConcurrentDictionary`——COM apartment-affinity 讓 `ConcurrentDictionary` 的執行緒安全是假象。

**關鍵人物：** 程式碼註解中會標註 Simon 與 Claude（偶爾 Gemini）的貢獻。`_rdo` 與 `_rdo2` 永遠成對初始化/釋放，這是不變式。

**程式碼寫作守則（Simon 明確規定，全程遵守）：**
- 動手前先拆解計畫給 Simon 確認，不自行假設
- 方向性/多重解讀問題必須列選項讓 Simon 選，不可猜測後貿然執行
- 看到更精簡做法要主動提出
- 誠實評估，不過度附和，方向有誤要明講
- 不知道的事要老實說不知道，不可幻覺出一個聽起來合理的答案
- 對話貼的程式碼 > 檔案區 > memory.md，優先順序不可顛倒
- 只動該動的部分，不順手動旁邊的程式碼；重要的可以先問
- 不可無故刪除原有註解/Debug記錄/日期標記
- 疑似 dead code 用說的，不自行清除
- 命名/風格要跟現有程式碼一致，不確定要問
- 呼叫鏈盡量要有規則性、對稱
- 能 50 行解決就不寫 200 行，不加不必要的彈性與 error handling
- 大檔案分段讀，不一次讀完
- 同一問題修兩次沒解決，要建立多假設 + 要求放探針收集線索，不可繼續憑單一推論亂改
- 小改動直接貼 diff 在對話；大改動或多處散落才整檔下載；沒讀入完整檔案時直接貼 diff（不要為了輸出整檔而浪費 token 讀入）

---

## 二、目前狀態

### 2.1 Stage 2 subtree 契約重構（既有主線工作，本輪未觸碰）

`GetSubtree` 回傳契約已從 `(folder As Folder, fPath)` 改成純資料 tuple `(eid, sid, fPath)`，省去 Phase 2 OOM 物化（~950ms）。Stage 1 元件已到位：`GetFolderById(eid, sid)` 作為集中物化 helper；`CcIsMail(cc)` 規則（IPF.Note/Post/Imap/空 → mail；IPF.Configuration → 非mail，已驗證 796 個資料夾）；三個免-folder L2.5 proxy 多載 `GetMailCount/GetFolderCount/GetFolderSize(fPath, eid, sid, skipCache)`。

**下一步 R2a（尚未執行，非本輪範圍）：** 在三個免-folder 多載中用 `DbGetFolderStats(fPath)` 直接還原 `② DB lazy`（信任 DB 不做 snap 驗證）。修正唯一已確認的 regression：暖快取重啟時，DB 直讀 = 99ms vs RDO 逐夾 = 611ms（慢 6 倍），原因是免-folder 路徑跳過了 `② DB lazy`。

### 2.2 COM 釋放順序修正（pending，本輪未觸碰）

`Form1_Closing()` 應改成 LIFO 釋放順序：`_pstStoreList` → `_rdo` → `ReleaseRdoSession()`（對應 `_rdo2`）→ `_olNS` → `_olApp`。現況是 `_olApp` 在 `_olNS` 之前釋放（子先於父的違規）。外層 `If _rdo2 IsNot Nothing` 應移除（`ReleaseRdoSession()` 內部已有 null check，外層 guard 有跳過快取清理的風險）。

### 2.3 Form1_Outlook.vb 程式碼組織（規劃中，本輪未觸碰）

全部 64 個函式已分類進五層模型（L1/L2/L2.5/L3/helpers）。三個標 `AllL3` 的函式其實是 L2 層聚合器（層級錯置）。純 COM 邏輯抽到 `Outlook_COM.vb` 預估搬動 ~743–933 行。判斷準則：任何函式碰 `_cacheXXX` 或呼叫非 L3 函式，就不屬於 `Outlook_COM.vb`。Simon 計畫先手動整理過一輪再繼續。全域宣告區塊（~85行）放哪裡也還沒決定。

### 2.4 RefreshMailInfoL3（pending，本輪未觸碰）

生產讀取路徑中最後一個還在用 `_rdo` 的函式——遷移到 `_rdo2` store-scoped 模式（`GetRdoStore(folderPath)` → `store.GetMessageFromID(entryID)`）的 diff 已備妥，等待套用。

### 2.5 ★ Task 2 — GetBasicMailInfo 重複邏輯處理（本輪主線，優先順序已更正）

**背景：** `GetFolderBasicByEntryIDL3`（Form1_Outlook.vb 928行）與 `GetBasicMailInfoOOM`（2397行）都是整夾 OOM table-scan，回傳形狀不同（Dictionary vs List(Mail,Topic)），有共用邏輯可抽但非同構（folder 取得/釋放責任不同、讀取欄位數不同 5 vs 7）。

**優先順序更正（Simon 2026/07/02 定案）：** 原規劃先做 2a（OOM dedup）再做 2b（RDO 加速），Simon 改成 **2b 優先，2a 延後**——「骨架整體可以動了再看 OOM 版本如何合併與優化」。

**Task 2a（OOM dedup，設計已定案但延後執行，尚未動手）：**
- 選項 B 定案：只共用「開 table → GetArray 迴圈 → 節流 → 取消處理」這段分頁掃描骨架（暫擬名 `ScanFolderTableL3`），欄位解析各自留在呼叫端（各自傳 row-parser lambda）
- 選項 A（統一讀 7 欄 superset）被否決：會讓 `GetFolderBasicByEntryIDL3`（方法B）每列多讀 2 個 COM 屬性造成不必要開銷，是行為變更不是純重構，Claude 判斷不該自己吞下這個 trade-off
- folder 釋放責任維持現狀不變：`GetFolderBasicByEntryIDL3` 自己解析+釋放，`GetBasicMailInfoOOM` 呼叫端持有+不釋放，helper 不介入
- 待確認事項（尚未問過 Simon）：helper 命名 `ScanFolderTableL3` 是否OK

**Task 2b（RDO 加速骨架，本輪主要產出）：**

設計定案：
```
Private Function GetBasicMailInfoRdo(fPath As String, eid As String, sid As String) As List(Of MailItemInfo)
```
- 純資料掃描骨架：不算 Topic、不轉容器，轉換責任下放呼叫端（903版 `GetBasicMailInfo` 轉 `List(Of (Mail, Topic))`；未來 `GetFolderBasicByEntryIDL3` 轉 `Dictionary(EntryID→MailItemInfo)`）
- 失敗訊號：`Nothing`（含中途例外，丟棄已累積的部分結果，不可回傳掃一半）；空夾回空 `List`（合法狀態非失敗）
- 同步、無 cToken/節流（比照既有 `GetMailCountRdo`/`GetFolderCountRdo`/`GetFolderSizeRdo`/`GetSubtreeRdoBatch` 全部都是 sync 的風格）
- 批次讀取：`tbl.Columns = COLS` → `GoToFirst()` → `GetRows(5000)` chunk 迴圈（比照 `GetFolderSizeRdo`，非 `GetSubtreeRdoBatch` 的一次 `GetRows(rc)`，因量級可能達單夾 5 萬封）
- 讀 7 欄：`EntryID, PR_SUBJECT, PR_MESSAGE_SIZE(0x0E080003), PR_MESSAGE_DELIVERY_TIME(0x0E060040), PR_SENDER_NAME(0x0C1A001E), PR_INTERNET_MESSAGE_ID(0x1035001E), PR_SENDER_EMAIL(0x0C1F001E)`，並填 `.FolderPath = fPath`
- **⚠️ EntryID 是 Byte()，須經既有 `RdoTableEidToHex()` 轉字串**（風險教訓來自讀 `GetSubtreeRdoBatch` 才發現，不是原本就知道的）
- **Claude 自己的設計取捨（已告知 Simon）：** `Subject`/`ReceivedTime`/`SenderName` 改用明確 proptag，而非沿用 OOM 版的具名屬性字串（`"Subject"`/`"ReceivedTime"`/`"SenderName"`）——因為既有代碼只證明過 `Folders.MAPITable` 認得具名屬性（`GetSubtreeRdoBatch` 用過 `"Name"`/`"EntryID"`），`Items.MAPITable` 是否也認得未經驗證，改用無歧義的 proptag 可減少一個未知變因
- 插入點：只做 903 版免-folder `GetBasicMailInfo`，`GetFolderBasicByEntryIDL3` 這輪不動
- 回傳型別定案討論：曾考慮 List vs Dictionary，最終選 **List(Of MailItemInfo)**（最原始掃描結果，不含 Topic），因為兩個未來呼叫端（903版要 List(Mail,Topic)、未來 `GetFolderBasicByEntryIDL3` 要 Dictionary）需求形狀不同，讓核心函式維持中立最省事

**探針 `TestProbeBasicInfoRdoParity`（PROBE_BASICINFO_RDO，程式碼已產出）：**
- 位置：`Form1_Maintab56.vb`，緊接 `TestProbeS1SubfolderSql` 之後
- 走選中 root 子樹（經 `GetSubtree(root, includeSubF:=True)` 拿 eid/sid/fPath 清單），逐夾用 `GetBasicMailInfoOOM` 當 baseline
- 比對維度：RDO EntryID 查無對應數（優先看，若非0代表EntryID轉換有問題）、6 個欄位逐筆 parity（Subject/Size/RcvTime/SenderName/MsgIDhash/SenderEmail）
- 同時 benchmark：Stopwatch 累加 OOM vs RDO 耗時，算倍率
- 已查證 `_dbg` 實作（`Form1.vb` 47行）：輸出走 `DebugForm.ActiveInstance.AddMessage3(...)`，是 in-app GUI 視窗，非 console/log 檔

**⚠️ 目前真實狀態（誠實記錄，避免下次誤判進度）：**
- `GetBasicMailInfoRdo` 本體與探針函式的完整程式碼**已在對話中產出並提供給 Simon**
- **是否已實際貼入本地檔案，未確認**（Claude 只能讀專案區唯讀副本，無法寫入 Simon 的實際工作檔案）
- **探針從未被實際執行過，沒有任何真實 parity/benchmark 數據**——上面所有「⚠️風險點」都只是設計時的預判疑慮，不是已驗證結論
- 除錯觸發（`Await TestProbeBasicInfoRdoParity()` 那行）尚未接上，需 Simon 手動加

**Claude Code 協作可行性（本輪額外討論，已查證）：**
- Claude Code 有本機檔案編輯 + bash 執行能力，可自己寫程式碼、跑 `dotnet build`
- 但**無法自動跑通這個探針拿數據**：探針觸發依賴 Tab1 TreeView 手動選取 root 節點 + 手動按除錯按鈕，輸出目的地是 `DebugForm` 這個 in-app GUI 視窗（非 console/檔案），Claude Code 沒有內建 Windows GUI automation 能力去自動點選/讀取
- 若要讓探針能被自動化執行，需先把觸發方式改成不依賴 GUI（例如 root 用參數/設定檔指定，輸出額外寫一份到文字檔）——這是否要做、怎麼做，**尚未定案**，需 Simon 決定
- memory 帶到 Claude Code 的做法：claude.ai 的 memory 系統與 Claude Code 不共用，需將內容整理進專案的 `CLAUDE.md`，Claude Code 每次啟動會自動讀取當上下文

---

## 三、未來規劃 / 待辦（依優先順序）

1. **[進行中] Task 2b 驗證**：Simon 套用 `GetBasicMailInfoRdo` + 探針程式碼 → 手動接除錯觸發 → 選中小型子樹跑探針 → 看 parity 結果分岔（過關才接 L2.5 ③ Fallback；不過關先排查 proptag/型別問題再重跑）
2. **[延後] Task 2a**：RDO 骨架穩定接上 L2.5 之後才做，選項 B 共用骨架已定案，待 Simon 確認 helper 命名後動手；同時讓 `GetFolderBasicByEntryIDL3` 未來能複用 `GetBasicMailInfoRdo` 核心（轉 Dictionary）
3. **R2a 實作**：三個免-folder 多載還原 `② DB lazy`
4. **ExecSQL 實驗**（Simon 2026/06/27）：比較 RDO `ExecSQL COUNT(*)` vs 逐年迴圈做 Tab2 年/月分桶。`ExecSQL GROUP BY` 不支援；`SUM()` 會讓 Redemption DLL crash（已確認）；`COUNT(*)` 驗證安全。待做：Tab56 獨立探針 `SpikeYearCountExecSql`
5. **`PR_CONTENT_COUNT` 併入 subtree batch**：可加進 `GetSubtreeRdoBatch` 既有 `COLS` 常數，零額外 COM 成本預熱 `_cacheMailCount`——需先做探針驗證跟 OOM `Items.Count` 的 parity
6. **效能剖析**（Simon 2026/06/23）：RDO 單執行緒 pipeline 的 Stopwatch 分段——RDO parse vs 屬性讀取（`.Body`/`.Attachments`）、body `.Body` 讀取 vs `NormalizeMailBody` regex、L2.5 dispatch 開銷、SimHash vs body-read 比例。Body 讀取 ~150/s vs attach ~3,000/s（20倍差距，推測是 body payload + regex 造成）
7. **L3 程式碼重組**進 `Outlook_COM.vb`（等 Simon 手動驗證層級邊界後）
8. **RDO 讀取平行化 Phase C**（延後——單執行緒目前判斷已足夠）
9. **[未定案] 探針自動化**：是否要把探針改成不依賴 GUI 的觸發方式，讓 Claude Code 之類的工具能無人值守執行

---

## 四、關鍵學習與原則

**架構：**
- 一致、對稱、解耦、統一呼叫介面。這讓插入新層級時呼叫端幾乎零改動（例如把 RDO 層插入 L2.5 proxy，Tab3/4/5 消費端幾乎沒動）
- `Module_Outlook.vb` 是 `Partial Class Form1`——所有 Private 成員 Tab56 等其他 partial class 檔案都能存取。不可未讀該檔頂層宣告就斷言成員不可存取
- PST 單檔 I/O：`PR_CONTENT_COUNT` 是快取的資料夾屬性（跨程序 COM 開銷才是瓶頸，不是 I/O 鎖）。`PR_MESSAGE_SIZE_EXTENDED (PT_I8)` 透過 `GetRows` 會回傳垃圾值；改用 `PR_MESSAGE_SIZE (PT_LONG 0x0E080003)`——比 OOM 快 3–10 倍，parity 已確認
- `ExecSQL SUM(size)` 會讓 Redemption DLL access violation crash——永久禁止。`ExecSQL COUNT(*)` 安全
- `GetRdoStore` 快取用一般 `Dictionary` 是刻意的（COM apartment-affinity 才是真正限制，`ConcurrentDictionary` 是假安全）
- 用獨立 `Logon(ProfileName, NewSession:=True)` 開的 `RDOSession` 只看得到自動設定檔載入清單裡的 store——手動掛載的 PST 對全新 session 是隱形的
- **[本輪新增] RDO table 讀出的 `EntryID` 是 `Byte()`，需經 `RdoTableEidToHex()` 轉字串，不可直接當字串用**（來自 `GetSubtreeRdoBatch` 既有寫法，本輪讀代碼才發現）
- **[本輪新增] `Items.MAPITable` 是否認得具名屬性字串（如 `"Subject"`/`"ReceivedTime"`）尚未驗證**——只證明過 `Folders.MAPITable` 認得（`GetSubtreeRdoBatch` 用 `"Name"`/`"EntryID"`），跨集合類型不可假設通用

**效能：**
- `Task.Yield()` 每次呼叫成本 5–15ms；快速 COM 操作應每 10 個資料夾節流一次
- `GetSubtreeRdoBatch` 比 OOM BFS 全 PST 估計快 135 倍，正確率 100%
- 跨 session 用單一 EntryID 參數的 `GetMessageFromID` 不可靠；改用 store-scoped `store.GetMessageFromID(eid)`
- RDO 讀取比 OOM 快 9 倍（254ms vs 2319ms）。「物化一次重複用」的假設是錯的——OOM 每次呼叫都會重讀 MAPI 屬性
- 探針必須同時測 parity（正確性）跟 benchmark（速度），且要在完整使用情境下測，不能只測單一操作讀取

**除錯紀律：**
- 假設要有量測證據才能行動。動手改前先收集診斷計數器
- 任何「不可達／不存在／必須妥協」的說法，講之前要先查證源碼。不可未查就斷言限制
- 一次 grep 完整列出所有呼叫點，逐一分類 region/存活狀態/語意角色，再動手。死路徑（註解掉的代碼、■99 region）要標出來跳過，不用問

---

## 五、做法與模式

**先探針再定案：** 所有實驗做成 Tab56 scaffolding region 裡的獨立完整函式，標 grep-able 清理 token（如 `PROBE_XXX`）。絕不為了探針修改正在運作的代碼。若一定要動正式代碼，要提供 grep token 清單供事後清理確認。

**探針既有命名慣例（本輪核對過）：** `' PROBE_XXX  ↓↓↓ 整塊可刪 ↓↓↓ ----`...`' PROBE_XXX  ↑↑↑ 整塊可刪 ↑↑↑ ----` 包住整段；函式內第一行註解 `' YYYY/MM/DD by Simon/Claude [PROBE_XXX]: <目的>`；用 `SimTree1.SelectedNode?.Tag` 轉型拿使用者選的 root 資料夾；用 `_dbg(header, msg)` 輸出結果；用 `Stopwatch` 量測。

**Diff 格式（Simon 2026/06/30 確認，永久）：** 修改用 `改前/改後` 成對區塊包在 ` ```vbnet ` 內，各半放乾淨代碼。不用 `+/-` 符號，不用行尾標記。多處零散改動 = 多個代碼區塊依序排列。大量/零散改動 = 整檔下載。**純新增（無修改既有代碼）不需要 改前/改後 格式，直接貼新代碼並註明插入位置即可**（本輪確認的補充規則）。

**改動紀律：** 不動既有可運作的代碼，除非明確取得同意。改動前要把架構分岔跟設計決策攤開來講。改別人（含 Claude 自己先前寫的）函式前要先標出來取得同意。

**呼叫點稽核規則：** 任何實作動手前，先 grep 全部呼叫點，一次分類完（region、存活/死亡、語意角色）。`Module_Win32API.vb` 的 ■99 region 永久排除在所有分析之外——grep 打到當死路徑跳過，不用評論。

**程式碼風格：** 繁體中文註解，`' YYYY/MM/DD by Simon/Claude:` 歸屬格式。適當使用壓縮的 VB.NET 單行風格。`System.Exception` 永遠完整限定名（Outlook namespace 衝突）。`_dbg` 做除錯記錄（輸出到 in-app DebugForm 視窗，非 console/檔案）。原則：「若能50行不寫200行」，把 blast radius 降到最小。

**檔案大小意識：** 超過 ~2,500 行的檔案該在架構邊界處拆分。只上傳當下任務相關的單一檔案。

---

## 六、工具與資源

- **核心技術棧：** VB.NET (.NET 10 WinForms)、Outlook LTSC 2021 OOM、Redemption RDO 6.7（`_rdo2` 獨立 session 為主，`_rdo` 淘汰中）、Microsoft.Data.Sqlite (SQLitePCLRaw)
- **資料庫：** `OLAcache.db`（資料夾統計、郵件 metadata）、`OLAcacheMail.db`（attach_filenames + mail_simhash——`ZipAndRebuildDB` 後仍保留）
- **關鍵檔案：** `Module_Outlook.vb` (~3,000行, Partial Class Form1)、`Form1_MainTab12.vb`、`Form1_Maintab56.vb`、`Module_SQLite2.vb`、`Module_Win32API.vb`
- **永久禁區：** `Module_Win32API.vb` → `#Region "■ 99 舊版備用 (勿刪)"`——絕不檢視、修改、或列入任何重構考量
