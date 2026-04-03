# Fallback 鏈與底層 COM 函數優化總結

## 優化概覽

我們已經成功將所有與「郵件數量」、「資料夾數量」、「郵件/資料夾大小」相關的底層 COM 存取函數，依照嚴格的 **「由快到慢」Fallback 機制** 進行了重構。同時徹底移除了所有可能會引發 OOM STA (Single-Threaded Apartment) 違規的安全隱患。

### 核心變更總結

1. **新增 `GetSubFolderList_RDO()`**：
   - 專門為 Redemption 平行路徑設計的 BFS。
   - 透過 `ConcurrentQueue` 搭配 `Parallel.ForEach` 展開樹狀結構，回傳 `List(Of Redemption.RDOFolder)`。
   - 完全繞過 OOM 物件在背景執行緒的存取限制。

2. **Parallel.ForEach + Interlocked.Add**：
   - `GetMailCountAll`, `GetFolderCountAll`, `GetFolderSizeAll` 的平行遍歷全面採用此架構，取代大量的 `Task.Run()`，大幅降低記憶體與 GC 負擔。

3. **統一且詳細的 Dbg 輸出**：
   - 保留了您所有的原始註解與除錯紀錄，並在每一次 Fallback 的進入點與失敗點加入明確的 `Dbg` 訊息（標示 `⓪ RDO`、`① OOM` 等層級），讓日後除錯一目了然。

---

## 函數 Fallback 鏈實作細節

### `GetMailCount()`
- ⓪ `rdoFolder.Items.Count` (底層讀取 PR_CONTENT_COUNT)
- ① MAPI: `PR_CONTENT_COUNT` (0x36020003)
- ② OOM: `folder.Items.Count`

### `GetMailCountAll()`
- ⓪ `rdoFolder.TotalItemCount` (MAPI 快取，瞬間取得整棵樹總數)
- ① 平行 BFS (RDO): `GetSubFolderList_RDO()` + `Parallel.ForEach` + `Interlocked.Add(rdoF.Items.Count)`
- ② 循序 BFS: `GetSubFolderList()` (OOM)+ 逐一 `GetMailCount()`
- ③ 遞迴 Fallback (OOM)

### `GetFolderCount()`
- ⓪ `rdoFolder.Folders.Count`
- ① MAPI: PR_FOLDER_CHILD_COUNT (現已註解備用)
- ② OOM: `folder.Folders.Count`

### `GetFolderCountAll()`
- ⓪ RDO 平行: `GetSubFolderList_RDO()` + `Parallel.ForEach` + `Interlocked.Add(rdoF.Folders.Count)`
- ① OOM BFS 循序累加
- ② 全部失敗: `Return -1`

### `GetFolderSize()`
- ⓪ RDO 屬性: `rdoFolder.Fields(PR_MESSAGE_SIZE_EXTENDED)` (0x0E080014)
- ① OOM 批次 (最快): `folder.GetTable()` + `GetArray(1000)` (讀取 PR_MESSAGE_SIZE_EXTENDED)
- ② OOM 循序 (保險): `folder.GetTable()` + `GetNextRow()`

### `GetFolderSizeAll()`
- ⓪ RDO 平行: `Parallel.ForEach` 取得每個資料夾的 `rdoF.Fields(PR_SIZE_EXTENDED)` 並 `Interlocked.Add`
- ① OOM 循序: 嚴格跑 `For` 迴圈呼叫 `Await GetFolderSize(L3)`，**避免破壞 STA**
- ② 全部失敗: `Return -1`

### Tab2 日期統計 (`GetYearCountsForFolder` 等)
- 先前已優化為單次 `GetTable("[ReceivedTime]...")` + `GetArray(BATCH_SIZE)` 並在記憶體中分組，無須再更動。

---

## 驗證建議
請實際運行應用程式，並透過 DebugForm 觀察以下訊息是否如預期出現：
1. 正常環境下應大量出現 `⓪ RDO平行成功`、`⓪ RDO 成功取得 TotalItemCount`。
2. Size 的部分若 PST 檔不支援 `PR_MESSAGE_SIZE_EXTENDED`，應該會看到退回 `① OOM GetTable.GetArray 成功`。
