# Outlook Assistant 修正計畫 - 2026/04/22

## 使用者需求回報
1. **Lv1 的 `EnterSelectedFolder` 進不去**：根據 debug message，搜尋目標為 `'銀行 OR 帳單'`，但 `parentNode.Nodes` 中的節點名稱可能因為 UI 格式化（前綴 ` - ` 或後綴空格）而無法精確匹配。
2. **Lv3 folderpath 與快取 OK**：此部分已確認正常，不需改動。
3. **Tab 6 (tabpage6) databaseStat 內容更新**：需要將新增的 `basic_maillist` 統計資訊加入顯示清單中。

## 待修復問題分析

### 1. Lv1 `EnterSelectedFolder` 節點搜尋邏輯
- **問題根因**：在 `EnterSelectedFolder` 中，我們使用 `t.SubFolder.Name`（例如 `"銀行 OR 帳單"`) 去匹配 `parentNode.Nodes` 裡的 `node.Text`。
- **現狀**：為了 UI 美觀，子資料夾在 ListView 顯示時會加上 `" - "` 前綴；而在 TreeView 節點中，可能也會因為斜體或其他格式化而有微小差異。
- **解決方案**：
  - 改用節點的 `Tag` 進行比對（`node.Tag` 存的是實體 `Outlook.Folder` 物件，我們可以比對其 `EntryID`）。
  - **最保險做法**：遍歷 `parentNode.Nodes` 時，直接比對 `node.Tag` (Outlook.Folder) 的 `EntryID` 是否與目標 `t.SubFolder.EntryID` 相同。

### 2. Tab 6 `databaseStat` 顯示內容
- **問題根因**：`RefreshDatabaseStats` 目前只顯示了 `attach_maillist` 等舊有的資料表。
- **解決方案**：
  - 修改 `GetDatabaseSummary` 函數，增加 `basic_maillist` 的計數查詢。
  - 修改 `RefreshDatabaseStats` 函數，在 ListView 中加入一列顯示 `basic_maillist` 的筆數。

---

## 預計變更檔案

### [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_MainTabs.vb)

#### [MODIFY] `EnterSelectedFolder`
- **修改內容**：
  - 在 `For Each node As TreeNode In parentNode.Nodes` 迴圈中，改用 `EntryID` 比對。

#### [MODIFY] `RefreshDatabaseStats`
- **修改內容**：
  - 在 SQLite 數據填充區段，加入 `basic_maillist` 的統計列。

### [Form1_SQLite2.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_SQLite2.vb)

#### [MODIFY] `GetDatabaseSummary`
- **修改內容**：
  - 增加一個回傳欄位（或修改現有 tuple）來包含 `basic_maillist` 的筆數。
  - 加入 SQL 查詢：`SELECT COUNT(*) FROM basic_maillist`。

---

## 驗證計畫

### 自動化測試 (透過 Debug 日誌)
1. **Lv1 導航驗證**：在 Lv1 雙擊資料夾，觀察 DebugForm 輸出。
   - 預期：`尋找節點` 步驟應能成功匹配 `foundNode`。
   - 預期：畫面應正確跳轉至該資料夾並展開其子資料夾。
2. **Tab 6 統計驗證**：切換至 Setting 頁面。
   - 預期：`databaseStat` (ListView) 應顯示 `basic_maillist` 及其正確筆數。

### 手動測試
- 確認 Lv1 鍵盤 Enter 鍵也能正常進入資料夾。
- 確認點擊 Save Cache 後，Tab 6 的統計數據有立即更新。
