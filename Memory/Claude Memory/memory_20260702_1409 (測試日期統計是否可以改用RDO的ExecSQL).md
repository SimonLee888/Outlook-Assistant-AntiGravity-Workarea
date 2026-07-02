# Outlook Assistant — ExecSQL 年份/月份計數加速 專題進度記錄

> 本檔案記錄單一 session 的完整討論脈絡：從 Redemption 官方文件查詢 → 探針設計 → 實測驗證 → 架構決策。
> 專案整體背景（VB.NET/.NET 10 WinForms、Redemption RDO、L1~L3 分層、快取架構等）請見既有的 project memory，本檔案不重複，只記錄本次 session 新增的內容。

---

## 一、本次 session 的起點：四個 Redemption 官方文件查證問題

Simon 提出四個問題，逐一查證 dimastr.com 官方文件、Outlook-Redemption groups.io 論壇、Microsoft Learn MAPI 文件後的結論：

### Q1. `GetSubtreeRdo` 可以順手抓 count 跟 content class 嗎？
- **content class：已經在抓。** 現有 `GetSubtreeRdoBatch` 的 `COLS` 已包含 `PR_CONTAINER_CLASS (0x3613001E)`。
- **count：目前沒抓，但理論上可以免費加進同一次 `GetRows`。** MS 官方文件證實：message-store hierarchy table 的 required column set 包含 `PR_CONTENT_COUNT (0x36020003, PT_LONG)`，跟現有已使用的 `PR_SUBFOLDERS`/`PR_CONTAINER_CLASS` 是同一張表上的欄位。
- **具體機會（尚未實作）：** 若在 `COLS` 加入 `PR_CONTENT_COUNT`，可以在 `GetSubtreeRdoBatch` 現有的批次 `GetRows` 中，順手把整支子樹每個節點的 `_cacheMailCount(nd.path)` 一併填好（目前該函數只回填 `_cacheFolderCount`），零額外 COM call。**這是獨立於本次 ExecSQL 年份計數專題的另一個優化機會，還沒有動手，也還沒開探針驗證 `PR_CONTENT_COUNT` 在你的 27 個 PST 上是否與現有 `GetMailCount` COM 路徑的值完全一致。**

### Q2. 可以用 RDO 的 `GetRows()` 直接抓整個 folder 下的 folder count 和 content class 嗎？
- 本質跟 Q1 同一件事，答案相同：folder count(子夾數) 用 `Folders.MAPITable.RowCount` 就有；content class 已經在抓；子夾各自的郵件數要加 `PR_CONTENT_COUNT`（同 Q1 的未實作機會）。論壇上也有其他開發者驗證過同樣的 `Folders.MAPITable` + `Columns`/`GetRows` 撈 `PR_SUBFOLDERS` 的模式。

### Q3. `GetRdoStore` 為何用一般 `Dictionary` 而不是 `ConcurrentDictionary`？
- **結論：現在的設計是對的，不建議只是把型別換成 `ConcurrentDictionary`。**
- 原因：Redemption/RDO 物件是 apartment／MAPI session 綁定的。官方論壇明確說明，背景執行緒若要用 Extended MAPI，必須自己呼叫 `MAPIInitialize`（每個 creatable 的 Redemption 物件如 `RDOSession` 建構時會自動做），且**背景執行緒不能直接共用 UI 執行緒建立的 RDO 物件**，正確做法是在背景執行緒自建一個新的 `RDOSession`，透過 `MAPIOBJECT` 屬性接上既有 session。
- 換句話說：`ConcurrentDictionary` 只解決「字典本身的 race」，不解決「字典裡裝的 COM 物件本身跨執行緒不安全」這個更根本的問題。**若未來 Phase C 真的要做背景執行緒讀取 RDO，需要的是「每執行緒各自一個 `RDOSession`」，不是換字典型別。** 這點留給 Phase C 再處理，目前現狀（單 UI 執行緒 + 一般 `Dictionary`）不需要改。

### Q4. `ExecSQL` 怎麼用在 `GetYearCount`/`GetMonthCount`？
→ 這是本次 session 後續花最多篇幅深入驗證的主題，見下方第二~五節。

**官方文件確認的 `ExecSQL` 能力邊界（重要，之後寫 production code 要記住）：**
- **支援：** `SELECT`、`WHERE`、`ORDER BY`、`TOP n`、`COUNT(*)`、`AS`、`LIKE`、比較運算子、`AND`/`OR`/`IN`/`NOT`、`IS NULL`/`IS NOT NULL`
- **不支援：** `GROUP BY`、`JOIN`、`INSERT` 等 → **沒有 `GROUP BY` 代表不可能一條 SQL 做完「年份分組計數」，必須逐年（或逐月）各發一條 `COUNT(*)`。**
- `FROM Folder` 是固定關鍵字（虛擬 pseudo-table 名稱），不管操作對象是 `Items.MAPITable` 還是 `Folders.MAPITable` 都用這個名稱，不是真的指涉「資料夾」語意。
- `ReceivedTime` 是文件列出的合法 OOM 屬性名稱，`WHERE` 子句可直接用，不需要 DASL 寫法（不用 `[ReceivedTime]` 中括號）。
- **已知地雷（先前既有驗證，非本次新發現）：** `SUM()` 在 `ExecSQL` 上會讓 `REDEMP~2.DLL` 產生 Access Violation 崩潰，只有 `COUNT(*)` 是安全的。

---

## 二、日期字面值格式實測結果（已完成，8 種候選全測過）

**背景：** `Date.ToString()` 預設會用系統語系（zh-TW），產生 `"2009/1/1 上午 12:00:00"` 這種字串，`ExecSQL` 完全看不懂，丟出 `Could not convert ... to a datetime value` 例外。

**8 種候選格式實測結果（用 `InvariantCulture` 產生，排除在地化問題）：**

| 格式 | 範例 | 結果 |
|---|---|---|
| A `yyyy-MM-dd` | `'2006-01-01'` | ✓ 可解析 |
| B `yyyy-MM-dd HH:mm:ss` | `'2006-01-01 00:00:00'` | ✓ 可解析 |
| C `yyyy-MM-ddTHH:mm:ss` | `'2006-01-01T00:00:00'` | ✓ 可解析 |
| D `yyyyMMdd` | `'20060101'` | ✗ Could not convert |
| E `MM/dd/yyyy` | `'01/01/2006'` | ✗ Could not convert |
| F `MM/dd/yyyy HH:mm:ss` | `'01/01/2006 00:00:00'` | ✗ Could not convert |
| G `#MM/dd/yyyy#` | `#01/01/2006#` | ✗ Unsupported operator: `/` |
| H `#MM/dd/yyyy HH:mm:ss#` | `#01/01/2006 00:00:00#` | ✗ Unsupported operator: `/` |

**結論：Redemption 的 `ExecSQL` 只吃 ISO 8601 字串格式，完全不認 `#...#`（Jet/Access 風格）也不認美式斜線日期。**

**採用哪一種？重要的邊界風險提醒：** 一開始探針邏輯是「第一個成功就採用」，結果自動選中 A（`yyyy-MM-dd`，無時間部分）。但年份查詢右邊界原本是 `New Date(y, 12, 31, 23, 59, 59)`，格式 A 會把時間部分截斷，若 Redemption 把裸日期字面值解讀成當天 `00:00:00`，**12/31 當天的信會被漏算**。已修正挑選規則為「優先選帶時間的成功格式」，最終自動升級選中 **B（`yyyy-MM-dd HH:mm:ss`）**。

**⚠️ 尚未在大量真實資料上驗證 12/31 邊界是否真的會漏算**——目前的 parity 測試都是用格式 B（有時間），沒有刻意測格式 A 在邊界日的行為差異。若未來真的要用格式 A 圖方便，這個邊界風險要先驗證過。目前的建議與後續探針都固定用 **格式 B `yyyy-MM-dd HH:mm:ss`**，不使用格式 A。

---

## 三、子樹級別實測結果（已完成，6 個真實子樹，17 萬封信）

**探針設計演進：**
1. 第一版探針只測「單一資料夾」，被 Simon 指出不切實際（現實中沒人會把好幾萬封信塞在同一資料夾）。
2. 改版後**沿用 production 既有合約**：用 `GetSubtreeRdo(rootFolder, rootPath)` 展開整支子樹，得到 `List(Of (eid, sid, fPath))`——這跟 Tab2 `CollectYearCounts(fList As List(Of (eid,sid,fPath)), ...)` 吃的參數是同一種 tuple，**不是另外兜一套走樹邏輯**。
3. A（基準）逐夾呼叫 production 的 `GetYearCountsForFolderL3`，物化時機比照 `GetYearCountsForFolder`（L2.5）：`GetFolderById` 延後到真正要算才建。
4. B（RDO ExecSQL）逐夾 × 逐年發 `COUNT(*)`，年份範圍**借用 A 算出來的 min~max**（探針階段的簡化，production 不能這樣做，見第四節①）。
5. 兩邊都用 `GetMailCount(fPath, eid, sid) <= 0` 提前過濾空夾，比照 `CollectYearCounts` 的既有慣例。

**實測結果彙總（6 個子樹，暖機一次後才計時）：**

| 子樹 | 資料夾數 | 郵件數 | A (GetTable) | B (ExecSQL) | 查詢數 | 加速比 | Parity |
|---|---|---|---|---|---|---|---|
| 工作郵件存檔 2006~2009 | 92 | 31,754 | 954ms | 240ms | 320 | 4.0x | ✓ |
| 工作郵件存檔 2010~2011 | 59 | 38,183 | 642ms | 122ms | 159 | 5.3x | ✓ |
| 工作郵件存檔 2015~2016 | 57 | 41,878 | 614ms | 101ms | 98 | 6.1x | ✓ |
| 工作郵件存檔 2017~2018 | 77 | 31,103 | 811ms | 80ms | 138 | 10.1x | ✓ |
| 寄件備份 2013~2018 | 10 | 15,865 | 133ms | 28ms | 36 | 4.8x | ✓ |
| 寄件備份 2006~11 | 68 | 11,255 | 530ms | 103ms | 275 | 5.2x | ✓ |
| **合計** | — | **170,038** | **3,684ms** | **674ms** | — | **5.5x** | **全部一致** |

之後 Simon 又用相同探針在更多子樹（工作郵件存檔、寄件備份、過期ePaper 等，共約 10+ 個子樹）重跑過一輪，並且**先暖機兩次、只截圖第三次結果**，確認數字穩定，但該次截圖右側被裁切，實際 ms 數字未能記錄（Claude 有明確告知看不到完整數字，Simon 表示可以，不影響決策）。

---

## 四、已知限制與尚未驗證的地方（誠實列出，決策前必讀）

1. **只驗證了「年」，沒驗證「月」。** `GetMonthCountsForYearL3` 若比照做法，會是「每夾 × 12 個月」而非「每夾 × 該夾橫跨的年數(通常 2~6 年)」，查詢數可能暴增到 3~6 倍（例如 92 夾 × 12 月 = 1,104 次）。**Simon 已明確決定：不用另外測月份，就算加速比只有 1~2 倍也值得做，直接照年份的架構套用到月份。**
2. **A/B 兩邊都沒有分開的冷/暖機嚴謹測試**，時間裡混了「COM 物件第一次接觸」的成本，但兩邊都沒暖機所以相對公平；5.5x 差距夠大，冷啟動雜訊不太可能翻盤，但嚴謹度上有此註記。
3. **A 刻意繞過①②快取（只測純 COM cache miss 情境）。** 5.5x 優勢主要體現在**第一次算（cache miss）**的情境，例如冷開機、快取被 invalidate、或第一次展開新選取範圍；重複查看同一資料夾在現有架構下本來就走①記憶體命中，不受此優化影響。
4. **年份範圍探測（TOP1+ORDER BY）尚未驗證。** 子樹探針的 B 迴圈是「借用」A 算出來的 min~max 年份，如果 ExecSQL 路徑要在 production 完全獨立運作（不依賴先跑一次 OOM），需要自己的範圍偵測手段。Simon 選定用 `ExecSQL SELECT TOP 1 ReceivedTime ... ORDER BY ReceivedTime ASC/DESC` 各查一次，**但這個語法組合（`TOP n` + `ORDER BY` 用在非 `COUNT(*)` 的一般 SELECT）目前一次都还没測過**，只有文件上說支援，需要開探針驗證。→ **這是目前最新、還沒執行的下一步（見第六節）。**
5. **`PR_CONTENT_COUNT` 加入 `GetSubtreeRdoBatch` 的機會（Q1/Q2 提到）完全獨立於本次 ExecSQL 年份計數專題，尚未實作也尚未驗證。**

---

## 五、Production 架構決策（Simon 已確認，尚未動手寫 production code）

透過 `ask_user_input_v0` 三選項確認，Simon 的選擇：

**① minYear/maxYear 範圍怎麼決定？**
→ **選 A：用 `ExecSQL TOP 1 + ORDER BY` 查真實範圍**（每夾多付 2 次查詢），**明確標注「未驗證，需先開探針測」**。

**② RDO 失敗時的 fallback 策略？**
→ **逐夾 fallback**：某夾 RDO 解析失敗（`GetRdoStore` 失敗／`GetFolderFromID` 失敗／`ExecSQL` 拋例外），該夾單獨退回 `GetFolderById` + 舊的 OOM `GetYearCountsForFolderL3`，其餘夾不受影響。**不可以像探針一樣直接跳過該夾**（production 跳過等於漏數字）。

**③ 新函數怎麼跟舊 L3 共存？**
→ **新增平行函數 `GetYearCountsForFolderL3Rdo(eid, sid, fPath)`**（免-folder 簽章，符合 Stage2 的一貫風格），`GetYearCountsForFolder`（L2.5）在③這一步**先試新的 RDO 版，失敗才退回現有的「`GetFolderById` + 舊 OOM `GetYearCountsForFolderL3`」**。舊函數完全不用改，風險最小，也符合專案既有「RDO 為主、OOM 為 fallback」的慣例。

**Production 實作時記得（跟探針的差異）：**
- 日期格式**直接硬編碼用格式 B (`yyyy-MM-dd HH:mm:ss`)**，不要把探針裡「跑時多格式測試挑第一個成功」的邏輯搬進 production（那是探索階段用的，production 已經知道答案，不需要每次都重新測）。
- 年份範圍偵測要用 `TOP1+ORDER BY`（見①），不能像探針一樣借用 A 的答案。

---

## 六、下一步待辦事項（依序）

### 目前最新、緊接著要做的（探針，尚未執行）
Claude 已經設計好 `SpikeYearRangeExecSql` 探針函數與三處小改動（插入 A 迴圈記錄 ground truth 範圍、B 迴圈裡驗證 `TOP1+ORDER BY` 抓到的範圍跟 A 是否一致、統計這兩條查詢的額外耗時），**程式碼已經在對話中貼出，但 Simon 明確表示「還沒有修改這些探針」**——也就是說 `SpikeYearCountExecSql` 的子樹版本（含格式測試 `PickWorkingDateFormat`/`BuildDateLiteral`/`SpikeYearExecSqlLoop`）以及最新的範圍驗證 helper `SpikeYearRangeExecSql`，都還只存在於對話紀錄裡，**尚未實際貼進 `Form1_Maintab56.vb`，也還沒掛上 `DebugButton_Click` 執行過。**

**接手 Claude Code 後的第一步：** 把這些探針程式碼實際套用到專案檔案，跑一次，重點看：
1. `TOP1+ORDER BY` 抓到的年份範圍是否跟 A(OOM) 的真實範圍**相符數 = 測試夾數**（100% 相符）。
2. 範圍驗證這兩條查詢的 `ms/次`，加上原本逐年 `COUNT(*)` 的 `ms/次`，兩者相加後跟純 OOM(A) 相比**還划不划算**（因為 production 用這個範圍偵測法，每個夾都要多付 2 次查詢的固定成本）。

### 驗證通過之後
3. 依第五節的三個架構決策，正式撰寫 production 函數：
   - 新增 `GetYearCountsForFolderL3Rdo(eid, sid, fPath)`（RDO ExecSQL 版本，內部：先 `TOP1+ORDER BY` 找範圍 → 逐年 `COUNT(*)` → 任何一步失敗就整個函數失敗，交給呼叫端 fallback）
   - 修改 `GetYearCountsForFolder`（L2.5）：③這一步先試新函數，失敗才退回舊路徑（`GetFolderById` + 舊 `GetYearCountsForFolderL3`）
   - 確認 `_cacheYearCounts` 寫入邏輯不用改（不管③走哪條路徑，寫回快取的介面一致）
4. 同樣的架構套用到月份：新增 `GetMonthCountsForYearL3Rdo`，修改 `GetMonthCountsForYear`（L2.5）——**Simon 已表示不用先另外測，直接做**，但要留意月份的查詢量結構跟年份不同（每夾固定 12 次起跳，資料夾數量本身是主要風險，不是年份跨度）。
5. 實際跑過一輪真實情境（不只是探針），確認 production 路徑（含 fallback 分支）行為正確。

### 尚未排入但已知的獨立機會（不屬於本次 ExecSQL 年份計數專題，先記錄）
- Q1/Q2 提到的 `PR_CONTENT_COUNT` 加入 `GetSubtreeRdoBatch`，讓 `_cacheMailCount` 在子樹展開時順手填好——這跟 R2a、Task 1 等既有 backlog 項目一樣，屬於獨立待辦，尚未排入本次專題的執行順序。

---

## 七、探針程式碼現況總覽（供對照，避免 Claude Code 重新設計一遍）

以下函數目前**只存在於對話文字紀錄中，尚未貼入任何專案檔案**，全部位於預定的 `Form1_Maintab56.vb`、`#Region` 內、`PROBE_YEARSQL` 可刪除區塊：

- `SpikeYearCountExecSql()` — 主探針，`Private Async Sub`，走子樹版：`SimTree3.SelectedNodes` 取 root → `GetSubtreeRdo` 展開 → A/B 對拍迴圈 → 彙總 log + `MessageBox.Show`
- `BuildDateLiteral(d As Date, fmtStr As String, useHash As Boolean) As String` — 共用的日期字面值產生器
- `PickWorkingDateFormat(rdoFolder, testYear, log) As (fmtStr, useHash)?` — 8 種格式全測，挑選規則「優先選帶時間的成功格式」
- `SpikeYearExecSqlLoop(rdoFolder, minYear, maxYear, fmtStr, useHash) As Dictionary(Of Integer, Integer)` — 逐年 `COUNT(*)` 迴圈
- `SpikeYearRangeExecSql(rdoFolder) As (minDate As Date?, maxDate As Date?)` — **最新設計，尚未整合進主探針測試過**，測 `TOP1+ORDER BY ASC/DESC`
- `DebugButton_Click` 的掛載方式：比照專案既有慣例，舊探針呼叫註解保留、新探針呼叫加上去（`Await SpikeYearCountExecSql()`）

若要在 Claude Code 重新取得完整程式碼內容，建議直接請 Claude 回顧本對話紀錄逐段複製，或由 Simon 重新貼一次目前 `Form1_Maintab56.vb` 的內容以確認最新基準。
