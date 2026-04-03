# Fallback Chain Optimization Plan

## User Review Required
> [!IMPORTANT]
> 請確認以下微調後的 Fallback 鏈是否符合預期。因為 OOM (Outlook Object Model) 有 STA 單執行緒限制，而且 Redemption 的 `RDOFolder` 不支援 `.GetArray()` 和 `.Size`，所以做了一些細微調整。

### 1. 關於 Parallel.ForEach 與 Task.Run 的差別
你問到 `Parallel.ForEach` 搭配 `Interlocked.Add` 會不會比 `Task.Run` 快？
**答案是：會。** `Parallel.ForEach` 在底層會根據 CPU 核心數自動最佳化 Thread 建立數量，借用 ThreadPool 重複利用；而對每個資料夾產生一個 `Task.Run` 若資料夾破千，會產生大量 Task 物件增加 GC 負擔。因此你提議的 **ConcurrentQueue + Parallel.ForEach 搭配 Interlocked.Add 是最理想的解法**。

### 2. GetSubFolderList 的平行化限制
原版 `GetSubFolderList()` 回傳的是 `Outlook.Folder` (OOM)。因為 OOM 物件有 STA 限制，我們**不能**在背景平行執行緒 (Parallel.ForEach) 裡建立並回傳它們，否則必定觸發異常。

**👉 調整方案：新增 RDO 專屬版**
我會新增一個 `GetSubFolderList_RDO()` 專門給 Redemption 平行路徑使用：
1. **[RDO 版]** _rdo 存在時，用 ConcurrentQueue + Parallel.ForEach 快速走訪，回傳 `List(Of Redemption.RDOFolder)`。
2. **[OOM 版]** _rdo 失敗或不存在時，Fallback 原本的 OOM 版 BFS 循序產生 `List(Of Outlook.Folder)`。

---

## 各層級函數 Fallback 鏈實作細節

### GetMailCount()
1. (RDO) `rdoFolder.Items.Count` (底層其實就是讀寫 `PR_CONTENT_COUNT`)。
2. (MAPI) `folder.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x36020003")`
3. (OOM) `folder.Items.Count`

### GetMailCountAll()
1. (RDO) `rdoFolder.TotalItemCount` (瞬間拿到整棵子樹)。
2. (RDO 平行) `GetSubFolderList_RDO()` + `Parallel.ForEach` + `Interlocked.Add(rdoF.Items.Count)`。
3. (OOM 循序) `GetSubFolderList()` (OOM版) + 循序迴圈呼叫 `GetMailCount` 加總。
4. (OOM 遞迴) 最後退回現有的遞迴安全網。

### GetFolderCount()
1. (RDO) `rdoFolder.Folders.Count`.
2. (OOM) `folder.Folders.Count`.

### GetFolderCountAll()
1. (RDO) `rdoFolder.Folders.Count` (如果可行)。否則跳 2。
2. (RDO 平行) `GetSubFolderList_RDO()` + `Parallel.ForEach` + `Interlocked.Add(rdoF.Folders.Count)`。
3. (OOM 循序) `GetSubFolderList()` (OOM版) + 循序迴圈呼叫 `GetFolderCount` 加總。

### GetFolderSize()
**限制說明：** `RDOFolder` 本身沒有 `.Size` 屬性，且 Redemption 的 Table 不支援微軟特有 `GetArray()` 方法。
1. (RDO) 嘗試讀取資料夾本身的 MAPI 屬性 `rdoFolder.Fields(PR_MESSAGE_SIZE_EXTENDED)` (Exchange偶爾會有，但本機端 PST 八成會報錯跳過)。
2. (OOM 批次) `GetTable(PR_MESSAGE_SIZE_EXTENDED)` + **`GetArray(1000)`** 批次讀取加總 (目前最快、最安全)。
3. (OOM 逐行) `GetTable(PR_MESSAGE_SIZE_EXTENDED)` + **`GetNextRow()`**。

### GetFolderSizeAll()
**限制說明：** 最快的 Size 計算來自上面 OOM `GetTable().GetArray()`。因為 OOM 只能循序執行，所以這層不能做 `Parallel.ForEach`，平行反而會崩潰。
1. (RDO 嘗試) 先試圖平行呼叫每個資料夾的 `rdoFolder.Fields(Size)`。
2. (OOM 循序 + 快速批次) 透過 BFS 列出循序 OOM 資料夾，內部呼叫上面的 `GetFolderSize()` (會走 `GetTable().GetArray()` 最快路徑)。

### Tab2 年份/月份統計 (已完成設計)
- 從多次 `Restrict` 變更為單次 `GetTable` 讀取 `ReceivedTime` + `GetArray(1000)` 記憶體內 `GroupBy Year` 和 `GroupBy Month`。

> 註：所有函數都會加上統一格式的 `Dbg` 紀錄，清楚標示「成功用到哪一層」與「在哪一層失敗而觸發 Fallback」。舊版 `GetMailCountAll(1/2)` 會移到備用區，加上 `#Region` 不去更動。

請檢視 [implementation_plan.md](file:///C:/Users/Simon/.gemini/antigravity/brain/12c7974d-c969-4d9c-b393-7698c9cb579c/implementation_plan.md) 並告知是否能以此邏輯開始動工。
