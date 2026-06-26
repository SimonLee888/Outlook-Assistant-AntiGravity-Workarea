# memory — RDO 列舉 vs OOM GetTable 實測 / 計數類 RDO 化評估 交接

> 用途：交給下一個對話接手。承接 2026/06/23 下午「導入 rdo2 到計數類」那一份。
> 撰寫：2026/06/23 by Simon/Claude Opus 4.8。專案：Outlook Assistant（VB.NET .NET10 WinForms，Outlook LTSC 2021 OOM + Redemption RDO 6.7.0.6412，本地 PST，olNotExchange）。
> ★ 本份原則：**只寫有看到實測數字的結論，不寫推論**。推論性內容單獨標 [推論] 並降權。

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
