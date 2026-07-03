# memory — RDO 列舉 vs OOM GetTable 實測 / 計數類 RDO 化評估 交接

> 用途：交給下一個對話接手。承接 2026/06/23 下午「導入 rdo2 到計數類」那一份。
> 撰寫：2026/06/23 by Simon/Claude Opus 4.8。專案：Outlook Assistant（VB.NET .NET10 WinForms，Outlook LTSC 2021 OOM + Redemption RDO 6.7.0.6412，本地 PST，olNotExchange）。
> ★ 本份原則：**只寫有看到實測數字的結論，不寫推論**。推論性內容單獨標 [推論] 並降權。

> ## ⚠️ 2026/07/02 更正（Simon/Claude）：§1.2/§2.1/§2.2 的「RDO 化是負優化」結論已過時，不可再引用
> **問題根源**：§1.2 探針二 B 段量的是 `For Each Items`（`.Columns` 設定失敗、退回逐列列舉），本質上是量到「RDO 最差讀法」的耗時，不是 RDO 真正的批次能力。
> **推翻證據**：2026/07/02 `GetMailInfoRdo`（`Module_Outlook.vb` ~1435 行，`memory_20260702_1254`）改用 `Items.MAPITable` + `.Columns=COLS` + `.GoToFirst()` + `.GetRows(5000)` chunk 批次讀（跟 OOM 的 `GetTable`+`GetArray` 同屬「一次 COM 往返抓一批列」），在 15,865 封真實信(含中日文、跨年份)上驗證：**parity 全過，RDO 比 OOM 快 2.0 倍**（306ms vs 609ms）。
> **§2.2 判準要更正為**：「OOM 逐封開信 → RDO 會贏」不變；但「OOM 已用 GetTable 批次 → RDO 會輸」這句是錯的 — 只要 RDO 也改用 `MAPITable+GetRows` 批次(而非 `For Each Items` 逐列列舉)，兩邊同屬批次讀，RDO 實測反而更快。真正決定勝負的是「兩邊讀法是否都是批次」，不是「OOM 有沒有用批次」。
> **§2.1 的 `GetAttachMailListL3`(現名 `GetAttachMailListOOM`)結論要重新評估**：這個函式屬於「OOM 用 `@SQL=hasattachment=True` 過濾後批次讀」，跟 `GetMailInfoRdo` 的「兩邊都讀全部列」基準不同(OOM 這裡篩選後通常只剩極少數列)，**不能直接套用 `GetMailInfoRdo` 的 2 倍結論**，但也不能沿用本份的舊結論。已在 `Form1_MainTab56.vb` `PROBE_ATTACHLIST_RDO` 區塊(2026/07/02)寫了新探針 `GetAttachMailListRdo_Probe`+`TestProbeAttachMailListRdoParity` 重測(命令列 `/autoprobeattach:Store|Folder` 觸發)。
>
> **2026/07/02 由 Claude Code 實際跑通(Outlook 已在跑,直接 `Start-Process` 帶命令列參數,無 GUI 手動觸發)**，`寄件備份, 2013~2018, Toshiba` 整個 store(10 個資料夾)實測：
> - **Parity 沒過**：有附件封數 OOM=3077、RDO=7450，RDO漏掉=0、**RDO多出=4373**(全部集中在 `2017` 這一個資料夾)。
> - 時間：OOM 累加=458ms、RDO 累加=370ms、倍率=1.2x — 但因為兩邊回傳的集合大小不同(RDO 多算了 4373 筆)，**這個時間比較目前無效**，不能拿來下結論。
> - **[推論] 根因**：候選 RDO 版用 `PR_HASATTACH`(0x0E1B000B) 當篩選欄位，這個 MAPI 屬性只要訊息有「任何」附件(含隱藏的內嵌圖片，例如 HTML 簽名檔的 embedded image)就是 True；而 OOM 的 `@SQL="urn:schemas:httpmail:hasattachment"=True` 篩選疑似排除了隱藏附件(`PR_ATTACHMENT_HIDDEN`)。全部差異集中在單一資料夾(2017)，符合「某年之後開始用帶內嵌圖片的簽名檔」這類猜測，但**這只是推論，尚未實際驗證**(需抽樣幾筆多出的 EntryID，實際列出 Attachments 看 Hidden 旗標才能確認)。
> - **下一步**：驗證上述假說 → 若成立，設計修正(例如篩選後的候選再檢查 `PR_ATTACHMENT_HIDDEN` 或只算 `olByValue` 型附件，可能需要對篩出的子集額外開一次 Attachments 集合，不能純靠 `Items.MAPITable` 一次批次搞定) → 修正後重跑本探針拿新的 parity+耗時數字，屆時才能對「該不該加 RDO 版本」下結論。**目前狀態：候選函式尚不能拿來用，此輪先如實記錄真數據，不下推薦。**
>
> **2026/07/02 假說驗證(`DiagnoseAttachmentsRdo`，抽樣 5 筆「多出」的信)：假說部分成立，但比預期更複雜。**
> 5 筆樣本分兩類：(a) 3 筆是 `Type=olByValue` 但 `Hidden=True`(HTML 簽名內嵌圖) (b) 2 筆是 `Type=olOLE`(RTF 內嵌物件，根本不是 `olByValue`)。統一規則：OOM 認定「有附件」≈「至少一個 `olByValue` 且非 `Hidden`」的附件——跟專案既有函式(`RefreshMailInfoRdo`/`OOM`、`GetAttachFilenameOOM`)「僅算 olByValue(1) 型附件」的既有慣例一致，只是這次多發現要再排除 `Hidden`。
>
> **2026/07/02 兩階段版本實測(加 Stage2 `HasVisibleByValueAttachmentRdo`，對 Stage1 候選逐封開信檢查 olByValue+非Hidden)：Parity 沒有變好，反而多了新問題，效能崩潰式變差。**
> - 精準度：多出從 4373 降到 43(大幅改善)，但**新增 RDO漏掉=542**(Stage1 只用 `PR_HASATTACH` 時漏掉是 0，說明 `Type=1`/`Fields(PR_ATTACHMENT_HIDDEN)` 這個判斷法本身還有沒抓對的邊界情況，尚未查明)。
> - **時間：OOM 累加=878ms、RDO 累加=15,637ms，倍率=0.1x（RDO 慢 OOM 18 倍）**。
>
> **✅ 結論已經夠清楚，可以下判斷了（不必再精修 Stage2 的判定邊界）**：即使把 Stage2 判定邏輯修到 100% 精準，瓶頸也不在「判斷邏輯對不對」，而在於「Stage2 對每個候選都要重新開一次 `store.GetMessageFromID` + 枚舉 `Attachments`」這個結構本身——OOM 的 `hasattachment` 是 Jet 引擎在 `GetTable` 階段就算好、伺服端一次批次篩完；RDO 沒有等價的單一批次屬性可以同時表達「至少一個 olByValue 且非 Hidden」，只能候選子集逐封開信驗證，這正好是「OOM 已批次、RDO 被迫逐封開信」的場景，跟 memory_20260702 的判準完全吻合：**OOM 逐封開信 → RDO 可能贏；RDO 被迫逐封開信而 OOM 已批次 → RDO 會輸。18 倍的實測數字印證了這點。**
>
> **📌 對 Q1「GetAttachMailList 該不該加 RDO 版本」的最終回答：不要。** 原因不是舊版「RDO 列舉天生慢」的錯誤結論，而是「這個函式的過濾語意(hasattachment 且排除隱藏/非 olByValue 附件)本質上需要逐封開信才能精準判定，RDO 在這種場景下沒有結構性優勢」。`GetAttachMailListOOM` 維持現狀即可。(PROBE_ATTACHLIST_RDO 探針已刪除。)
>
> **2026/07/02 深夜補測(PROBE_RESTRICT_RDO)：最後一條路「Restrict 下推」也實測封死,此題蓋棺。** 同 store 實測：
> - `RDOItems.Restrict` API 存在(晚綁定可呼叫)。語法注意:不吃 OOM 的 `@SQL=` 前綴(拋 You cannot set more than one operator),也不吃 `= True` 字面值(true 被當屬性名),要寫成 `"urn:schemas:httpmail:hasattachment" = 1`(無前綴+數字字面值)。ExecSQL 的布林條件同理,`WHERE "…0x0E1B000B" <> 0` 可用、`= true` 不行。
> - **語意判決**：`Restrict(hasattachment DASL)` = `Restrict(raw PR_HASATTACH)` = `ExecSQL COUNT` = **7450**,OOM 迴紋針 = 3077,差就是那 4373 筆——**Redemption 把 `urn:schemas:httpmail:hasattachment` 單純映射成 raw `PR_HASATTACH`,沒有 Outlook 查詢引擎的迴紋針加工(排除 Hidden/olOLE)**。
> - **速度判決(更關鍵的補刀)**：Restrict+批次讀 = 1725ms vs OOM = 258ms,**慢 6.7 倍**——即使日後找到能表達迴紋針語意的 subrestriction 語法,Restrict 評估路徑本身就輸,語意問題已無關緊要。
> - **Option B(ExecSQL 剪枝器)也順便量死**：12ms/夾,而 OOM 掃描平均 ~26ms/夾(無附件夾更便宜,只剩 GetTable 固定成本)——花 12ms 省 ~15ms,無實質收益,不做。
> - **可留用的正面副產品**：Redemption Restrict/ExecSQL 的可用語法形狀(上面第一點)已驗證,未來別的場景要用 Restrict 時直接照抄,不用再猜。
> - PROBE_RESTRICT_RDO 探針(`Form1_MainTab56.vb` + `Form1.vb` `/autoproberestrict` 掛勾)使命完成,可刪。
>
> §1.1(IPMRootFolder 解 visibility)、§1.3(共用 session 慢 13 倍)兩段數字不受影響，仍然有效。

---

---

## 0. 鐵則（每次修改都要守，沿用前份）
1. 動既有 working code 前先問 Simon。
2. 一致、對稱、解耦、呼叫路徑統一。
3. **不猜 API**——本輪嚴重違反多次（見 §4），務必查證或先寫最小驗證再寫正式。
4. 對話貼/傳的程式碼最新 > 檔案區 > memory。讀大檔分段讀。
5. 既有註解/debug/日期不可遺失；dead code 用說的不自刪。
6. 小改動貼 diff；回覆繁中+英文。
7. ■99 禁區（Module_Win32API.vb 舊版備用 Region）永遠不看不動。
8. 檢視呼叫端一次到位，不分批冒問題。

---

## 1. 本輪兩支探針的實測數字（這是本份核心，全部有截圖根據）

### 1.1 探針一 SpikeFolderVisibilityCompare — visibility 已用 IPMRootFolder 解決（數字確證）
- 第一次（RDO 端用 `st.RootFolder`）：OOM 可見夾 = **822**；RDO 枚舉 = **1028**；RDO-only = **206**；OOM-only = **0**。
- 那 206 個 RDO-only 全部是 search folder / 系統夾：`IPM_COMMON_VIEWS`、`IPM_VIEWS`、`ItemProcSearch`、`SPAM Search Folder 2`、`Search Folders`、`Contact Search`、`Drizzle`、`Freebusy Data`、`MS-OLK-AllMailItems`、`MS-OLK-FGPooledSearchFolder*`、`MS-OLK-BGPooledSearchFolder*`、`MS-OLK-AllOutlookItems1/2`、`~MAPISP(Internal)`、`待辦事項列搜尋`、`追蹤的郵件處理`、`提醒`、`搜尋資料夾1~30` 等。
- 第二次（RDO 端改 `st.IPMRootFolder`）：RDO 枚舉 = **822**；RDO-only = **0**；OOM-only = **0**。**兩集合完全相等。**
- **確證結論**：RDO 子樹枚舉只要從 `st.IPMRootFolder` 出發（不是 `st.RootFolder`），枚舉集合天生 = OOM 可見集合，那 206 個系統夾在 IPM 樹外天生撈不到。**零黑名單、零白名單、零快取、零 isRDO 過濾、新增夾照樣枚舉得到。**
- 附帶事實：探針一裡 `Kind`/`PR_ATTR_HIDDEN` 都讀成 `?`（讀法錯，見 §4）。但 IPMRootFolder 已解 visibility，這兩個判據用不上了。[推論] 屬性名可能該用 `FolderKind`（rdoFolderKind.fkSearch）而非 `Kind`，HotExamples 有此寫法，未在本機驗證。

### 1.2 探針二 SpikeFolderTableBenchmark — RDO 列舉比 OOM GetTable 慢一個數量級（數字確證）
標的：8 個 >= 500 封的夾，Work profile。

**A 段 OOM GetTable（基準，逐夾 Table+Array ms）：**
- Inbox 1156 封：Table 24 + Array 22 = 46ms
- 2019  1740 封：Table 15 + Array 33 = 48ms
- Finance 936 封：Table 14 + Array 19 = 33ms
- TEVP  782 封：Table 12 + Array 14 = 26ms
- （行事曆夾因無 ReceivedTime 欄位拋例外，正常，略過）

**B 段 RDO 列舉 Items（設 Columns 失敗但列舉照跑，Read ms）：**
- Inbox 1156：Read 470ms（對 OOM 46ms，約 10×慢）
- 2019  1740：Read 725ms（對 48ms，約 15×慢）
- Finance 936：Read 460ms（對 33ms，約 14×慢）
- TEVP  782：Read 544ms（對 26ms，約 21×慢）
- Resolve 段（ResolveFolderOnSession BFS 解夾）3~47ms，Cols 段 7~10ms。

**確證結論**：對「整夾逐封多欄位」場景，RDO `For Each Items`（即使設 Columns）比 OOM GetTable/GetArray **慢約 10~20 倍**。OOM GetTable 一次 COM 往返抓一批列；RDO 列舉是 per-row COM 往返，marshalling 成本累積。

### 1.3 探針二 C 段 平行 K=1/2/4（數字確證，此次讀法已修對）
| 模式 | K=1 | K=2 | K=4 |
|---|---|---|---|
| 共用 _rdo2 | 30847ms | 26957ms | 24524ms |
| 各自獨立 session | 2269ms | 1699ms | 1587ms |

- **共用 _rdo2 比各自獨立 session 慢約 13 倍**（K=1: 30847 vs 2269），且幾乎不隨 K 加速。→ 多 worker 共用同一條 session 被嚴重序列化，**確定不可行**。
- 各自獨立 session 平行有效但微弱：K=1→K=4 只快 1.43×（2269→1587），K=2→K=4 幾乎無進步（1699→1587）。→ 瓶頸在 PST 實體 I/O / MSPST provider，與前兩天計數結論方向一致。
- 各自獨立 session 即使 K=1（2269ms），仍比 A 段 OOM 8 夾總和慢很多——再次印證 RDO 列舉本身就慢。

---

## 2. 由數字導出的下一步決策

### 2.1 Tab3 / Tab4 兩條進度條：維持 OOM，不做 RDO 化
- Tab3「載入資料夾清單」底層 `GetAttachMailListL3`、Tab4「掃描系列郵件」底層 `GetBasicMailInfoL3`，兩者都走 OOM `SafeGetTable`+`SafeGetArray` 批次讀。
- 實測 RDO 列舉慢 10~20 倍，平行只追回 1.4×，補不回來。**RDO 化是負優化，維持 OOM GetTable。**

### 2.2 [可寫進長期記憶的判準] 「能不能用 RDO 加速」取決於 OOM 在該場景是否逐封開信
- attach/body（phase 1）RDO 贏：因 OOM 要逐封 GetItemFromID 開信，RDO 繞過開信。
- count / GetTable 場景 RDO 輸：因 OOM 本來就用 table 批次，RDO 沒有對應批次能力（MAPITable 無 GetArray，列舉仍 per-row）。
- 判準：**OOM 逐封開信 → RDO 可能贏；OOM 已用 GetTable 批次 → RDO 會輸。**

### 2.3 之前沒寫好、確定有問題的 RDO 段落（Module_Outlook.vb）
- `_rdoFastPath`（宣告第 52 行，恆 False）+ 三個使用點（約 1562/1583/1729）+ `GetSubtreeToListL3_Rdo`（約 2410）：這條子樹枚舉 RDO 平行路徑**從未真正執行過**（開關恆 False）。
- 已知兩個確定問題：(a) visibility——用 `RootFolder` 會多撈 206 系統夾（本輪已證實，解法=改 IPMRootFolder）；(b) `GetSubtreeToListL3_Rdo` 內 2414 行註解宣稱「Redemption free-threaded 可直接 Parallel.ForEach 共用」——本輪 C 段數字證實**共用 session 慢 13 倍**，此註解的前提錯誤。
- 1741 段舊平行碼用 `Parallel.ForEach` 跨執行緒碰同一批 `rdoF.Folders.Count`（共用模式），按 C 段數字屬於慢 13× 的那種寫法。
- **處置選項（待 Simon 決定，本輪未動任何 production 碼）**：
  - 選項甲：整段清除這條 RDO 平行路徑（_rdoFastPath + 三使用點 RDO 區塊 + GetSubtreeToListL3_Rdo），因實測證明此場景 RDO 負優化、共用 session 不可行，留著是死碼+錯誤註解。
  - 選項乙：改寫完整（IPMRootFolder + 各自獨立 session），但 §1.2/§1.3 數字顯示即使改對也比 OOM 慢，[推論] 不值得。
  - Claude 傾向：[推論] 選項甲清除為主，因為改寫完整也贏不了 OOM。但這是動既有碼，須 Simon 拍板。

### 2.4 GetMailCountAll / GetFolderCountAll 能否用 RDO 一舉換掉？— 查無捷徑
- 本輪查證（dimastr 官方 + Outlook-Redemption groups.io）：**RDO 沒有「整樹彙總 total item count / total folder count」的單一屬性**。微軟 `TotalItemCount` 是 FeedFolder(RSS) 專屬，非通用 MAPI 夾屬性。
- 唯一有根據可試方向：felix（groups.io 2021）用 `RDOFolder.Folders.MAPITable.ExecSQL("SELECT Name, EntryID, ...0x360A000B AS PR_SUBFOLDERS")` 一次撈某夾所有子夾結構、只開有子夾者遞迴。這是「枚舉夾結構」加速，**不是 count 彙總**。
- 純計數（只要一個數字）可用 `RDOFolder.Items.Count`（MAPI 快取屬性，不必 GetTable）。但要換 GetMailCountAll 仍須逐夾加總，與現行 OOM 做法同質，無質變優勢。
- **下一對話若要試**：只能試 `Items.Count` 逐夾加總 vs OOM 現行的速度對照，或 `ExecSQL` 撈子夾結構 vs OOM .Folders 枚舉。**[推論] 鑑於本輪 RDO 列舉已慢 10-20×，勿預設樂觀，要先小實驗驗證再投入。**

---

## 3. 本輪唯一確定可採用的正向成果
- **IPMRootFolder 解 visibility**（§1.1）：若未來任何 RDO 子樹枚舉要恢復，root 一律從 `st.IPMRootFolder` 出發，不需任何過濾。這是乾淨還債方案，已被 822=822 數字證實。

---

## 4. 本輪 Claude 表現檢討（Simon 嚴正不滿，記取）
1. **連續猜 API 不查證**（最嚴重，違反鐵則 3）：`Kind` 用 CallByName+CStr（錯）；`PR_ATTR_HIDDEN` 用 .Fields(DASL)（search folder 拋例外）；`MAPITable.GetArray`（不存在，B/C 段第一次全爆）；`MAPITable.Columns.Clear`（不存在，第二次仍報錯）。每次都讓 Simon 多貼一次、多測一次。
2. **ask_user_input 互動框在 Simon 端未正常渲染**：連續多次「先問你一個方向問題：」後就斷句，問題內容卡在沒顯示的元件裡，浪費數輪。教訓：**改用純文字把選項寫完，不要用 ask_user_input 工具問方向。**
3. **為保護自己而轉嫁 effort**：提議「先寫最小探針」表面是謹慎，實則把「我怕再猜錯」的風險轉成 Simon 多貼幾次。Simon 明確指出：查清楚再一次寫對，比讓他重貼多次重要。
4. 整個下午幾乎零有效產出，全靠 Simon 一直提示方向、指正、找錯、重複解釋。下一對話務必：查證到位再出手、一次到位、用文字問選項。

---

## 5. 待清理（測完的拋棄式探針）
- `SpikeFolderVisibilityCompare`（探針一）、`SpikeFolderTableBenchmark`（探針二）、`ResolveFolderOnSession`（探針二專用 helper）——皆在 Form1_Maintab56.vb Debug 測試區，使命已達成，可整段刪除。
- 上一份提到的 `SpikeResolveFolderOnRdo2` 等舊探針若仍在，一併視情況清理。
