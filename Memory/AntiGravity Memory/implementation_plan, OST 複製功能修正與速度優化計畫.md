# OST 複製功能修正與速度優化計畫

這個計畫旨在解決在 Tab7 執行 `CopyFolder` 時，第二次複製會出現 `Parent folder NID not found` 錯誤的問題，並優化匯出速度。

## 使用者評論回應
關於您的意見：「OST 內容在第一次讀的時候應該已經有大部分在記憶體裡面了吧？」
是的，`ost2pst` 在 `LoadOST` 時會將 **NBT (Node B-Tree)** 與 **BBT (Block B-Tree)** 讀入記憶體清單中。然而，目前的 `CopySourceDatablocksToPST` 實作會「破壞」這些記憶體中的 NBT 狀態（為了對齊 PST 格式），導致第二次複製時找不到原本的父資料夾 NID。

## 待解決問題
1. **錯誤原因**：`CopySourceDatablocksToPST` 內部的 `CheckFoldersToExport` 會修改 NBT 節點的 `nidParent`，且匯出過程會清除或過濾掉不符合條件的節點，導致原本載入的 `FM.folders` 關係斷層。
2. **速度瓶頸**：目前的匯出邏輯會走訪 **所有的 NBT 節點**（可能數萬個），即使我們只需要匯出特定資料夾。對於第二次複製，因為 NBT 狀態已亂掉，速度反而變慢甚至報錯。

## 提出的變更

### 1. [MODIFY] [Form1_OST.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_OST.vb)

#### 修正 NID 遺失問題 (解決錯誤 1 & 2)
在呼叫 `CopySourceDatablocksToPST` 前，先備份當前的 `ost2pst.FM.srcFile.NBTs`。匯出完成後立即還原，確保 `FM.folders` 的父子關係不會因為 `nidParent` 被修改而失效。

#### 優化匯出速度 (解決速度問題 3)
實作「精準掃描」邏輯。目前的 `CopySourceDatablocksToPST` 是全檔案掃描，我們可以改為：
- 只針對標記為 `toBeExported` 的資料夾及其內容進行處理。
- 既然 NBT 已經在記憶體中，我們可以利用現有的 `folders` 清單來精確導航，而不是遍歷整個 `NBTs` 陣列。
- *註：由於底層庫限制，我們先採用「狀態保護」模式，並優化過濾判斷。*

#### 資源清理
確保 `CloseOutputFile` 被正確呼叫，並在結束後清理 `MessagesToExportNIDs`。

## 驗證計畫

### 手動測試
1. 啟動程式，載入 OST。
2. 選擇第一個資料夾進行 `Copy Folder`，確認成功。
3. **不要重新 Load OST**，直接選擇第二個資料夾再次 `Copy Folder`。
4. 確認不會出現 `Parent folder NID not found` 錯誤。
5. 觀察 ProgressBar2 顯示的速度，預期第二次應該維持穩定速度或更快。

### 程式碼複檢
- [ ] 確認 `NBTs` 的備份與還原邏輯在 `SyncLock` 內。
- [ ] 確認 `tempPstPath` 的 GUID 命名避免衝突。
- [ ] 檢查 `ost2pst.FM.srcFile` 的 null check。
