# 任務清單

- [x] 分析 TreeView1 展開節點 (`BeforeExpand` 或相關事件) 時的邏輯，找出為何 `selectednode` 已經統計過，點開 `+` 號時還會再統計一次 `subfolder.count` (只有第一次會)。
- [x] 分析切換目錄時，RDO 成功呼叫 `GetMailCount` 而沒有讀取快取的邏輯原因。

### 執行階段 (Execution)
- [x] 將 [Form1.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1.vb) 中的資料夾相關快取字典 (如 `_mailCountCache`, `_folderCountCache`, `_folderTreeCache` 等) 的鍵值型別改為 `String` (使用 `FolderPath`)。
- [x] 於 L2 層級新增 `_directFolderCountCache` 與 `_directMailCountCache` 兩個字串鍵值的字典。
- [x] 修改 `LoadSubFolderToTreeView` 邏輯，使其查閱 `_directFolderCountCache`，若命中快取則直接使用，否則呼叫 `GetFolderCount` 並寫入快取。
- [x] 修改 `ComputeFolderStatsAsync` 邏輯，在 Step 2 寫入 `_directMailCountCache`，並在 Step 5 讀取它，徹底避免切換資料夾時二次呼叫 `GetMailCount`。
- [x] 配合快取鍵值型别變更，修改 `ComputeFolderStatsAsync` 內部讀取/寫入 `_mailCountCache` 和 `_folderCountCache` 的邏輯。
