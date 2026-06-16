# Form1.vb Tab1 效能與快取未命中問題分析與修正計畫

## 發現問題的原因

### 問題 1: 為何點開 `+` 號時會重複統計 `subfolder.count`？
> **回應您的疑問**：「當我點了 A 資料夾... 應該已經有 A1~A4 的 subfolder count 在快取裡面了，為何點開 + 號時還會再去讀一次？」
> 
> 您的觀察非常精確！的確，當您點選 A 時，右側 ListView 已經靠底層 BFS 展開並算出了 A1~A4 的「包含子孫的資料夾總數」(`TotalSubCount`)，這些數字也確實存進了 `_folderCountCache` 裡。
> 
> **問題出在 `LoadSubFolderToTreeView` 呼叫的函數**：
> 當您點開 `+` 號載入 A1~A4 作為真實節點時，它依賴 `GetFolderCount(folder)` 去檢查「這一個直屬子資料夾本身底下有沒有資料夾」。而目前的 `GetFolderCount(folder)` 底層函數**完全沒有去讀任何快取**的邏輯。它被設計成直接去問 RDO：`rdoFolder.Folders.Count`。
> 另外，`_folderCountCache` 存的是「包含子孫」的資料夾總數，但 `LoadSubFolderToTreeView` 需要的是「第一層直屬子資料夾數量」，所以無法直接拿現有快取套用。這就導致雖然系統已經遍歷過這些資料夾，但是點開加號時，沒有快取保護的 `GetFolderCount(folder)` 就會被迫再發出一連串的 RDO COM 呼叫了。

### 問題 2: 為何在 A/B 目錄之間切換時，仍會看見 RDO `GetMailCount` 成功，而沒有讀取快取？
**原因**：雖然 L2 的 `ComputeFolderStatsAsync` 有將「包含子孫的郵件總計」(`TotalMailCount`) 存入 `_mailCountCache`，但是 ListView1 的第一欄需要顯示的是「**本層**郵件數」(`DirectMailCount`)。
在 `ComputeFolderStatsAsync` 的 Step 5 (組裝回傳清單) 中，程式發現如果該節點是從快取命中的 (`entry.IsFromCache = True`)，它就沒有在 Step 2 讀取本層郵件數，因此會**強制再呼叫一次** `GetMailCount(entry.Folder)` 以取得 `DirectMailCount` 供 ListView 顯示。這會對您點選的目標資料夾及其**所有直屬子資料夾**各發起一次 `GetMailCount` 呼叫，導致明明命中快取了依然反覆觸發 RDO COM 呼叫。

---

## 建議的程式碼修正 (Proposed Changes)

為徹底解決以上問題，我建議進行以下三項修正。回應您對於分層架構與快取鍵值 (Cache Key) 的精彩提問：

> **關於快取應該放在 L2 還是 L3 的討論**：
> 您說得完全正確！我們當初定下的架構是：**L3 應該是純粹的資料取得與 Fallback 處理，不該牽涉狀態保存與快取**。
> 因此，快取字典與判斷邏輯 **必須放在 L2 (流程控制層)**。
> - **優點**：架構乾淨、職責分明。L3 (`GetFolderCount`, `GetMailCount`) 永遠回傳最真實、最新的 COM 數據。L2 (`ComputeFolderStatsAsync`, 或新建的 L2 封裝函數如 `GetCachedFolderCount`) 則負責「決定是否要打 L3」或「直接從記憶體拿」。
> - **實作方式**：我們會在 L2 層級宣告這些 Dictionary，並在 `LoadSubFolderToTreeView` 呼叫 L3 之前，先在 L2 進行攔截與快取讀取。

> **關於為何不用 `Outlook.Folder` 物件而改用字串作為 Cache Key 的討論**：
> 為什麼我們之前的快取會「看得到吃不到」(Cache Miss)？因為在 Outlook COM (RCW 機制) 中，即使是同一個實體資料夾，每次透過 COM 屬性讀取時，.NET 可能會產生一個**記憶體位址不同**的 Wrapper 物件。這導致 `Dictionary.ContainsKey(folderObj)` 永遠回傳 `False`。
> 我們必須改用**字串**來當 Key。那麼該用 `EntryID` 還是 `FolderPath`？
> 1. **EntryID**：由 Server 核發的唯一識別碼。優點是絕對唯一，缺點是：如果信件/資料夾被「移動」(Move) 到另一個信箱 (Store)，EntryID **會改變**。
> 2. **FolderPath**：例如 `\\user@domain\Inbox\Test`。優點是直觀且不會像 COM 物件位址那樣亂跳。缺點是：如果在 Outlook 裡「重新命名」或「移動資料夾」，路徑就會變，此時快取就對不上了 (但這剛好符合我們需要重新掃描的需求)。
> **結論**：對於**資料夾統計快取**而言，使用 `folder.FolderPath` (字串) 會比 `EntryID` 或 `Outlook.Folder` 物件更清晰、穩定且好除錯。我們將全面改用 `FolderPath` 作為鍵值！

### 1. 修正 `GetFolderCount` 重複讀取問題 (於 L2 實作)
- **作法**：在 L2 層級新增快取字典 `_directFolderCountCache As New ConcurrentDictionary(Of String, Integer)`，鍵值改為 `FolderPath` 字串。
- **修改點**：在 `LoadSubFolderToTreeView` (L1/L2邊界) 中，不直接呼叫 `GetFolderCount`，而是先檢查 `_directFolderCountCache` 是否存在該 `FolderPath`。若無，才呼叫 L3 `GetFolderCount` 並寫入快取。這樣 L3 依然保持純潔。

### 2. 修正 `ComputeFolderStatsAsync` 中 `DirectMailCount` 未快取的問題 (於 L2 實作)
- **作法**：新增快取字典 `_directMailCountCache As New ConcurrentDictionary(Of String, Integer)`，鍵值同樣改為 `FolderPath` 字串。
- **修改點**：在 L2 的 `ComputeFolderStatsAsync` Step 2，當真正呼叫 L3 打撈出 `DirectMailCount` 時，將其寫入 `_directMailCountCache`。
- **修改點**：在 Step 5 中，當 `entry.IsFromCache` 為 `True` 時，改為透過 `FolderPath` 優先從 `_directMailCountCache` 讀取，如果真的沒有才 fallback 到 L3 的 `GetMailCount`。

### 3. 全面更新現有快取的鍵值設計
- **作法**：將現有的 `_mailCountCache` 與 `_folderCountCache`，從原本以 `Outlook.Folder` (物件) 為 Key，全面修正為以 `FolderPath` (字串) 為 Key。
- **目的**：徹底根除長久以來因 COM 物件位址飄移而導致的隨機性快取未命中問題。

---

## 驗證計畫 (Verification Plan)

### 手動驗證 (Manual Verification)
1. **問題 1 驗證**：套用修改後，點選並展開任何有許多子資料夾的節點。觀察 Debug 視窗，確認除了第一次掃描外，不會再湧現大量的 `GetFolderCount ⓪ RDO 成功` 訊息，且展開動作不再有卡頓感。
2. **問題 2 驗證**：於兩個已掃描過的資料夾 (A 與 B) 之間來回切換點選。觀察 Debug 視窗，確認不再出現 `GetMailCount ⓪ RDO 成功` 的迴圈訊息。切換速度應近乎瞬間完成，確實命中快取。

> **User Review Required**: 
> 請問您是否同意此修正方向？同意的話，我會開始執行修改，並依照您的規則加上 `by AntiGravity, 2026/xx/xx` 的註解。
