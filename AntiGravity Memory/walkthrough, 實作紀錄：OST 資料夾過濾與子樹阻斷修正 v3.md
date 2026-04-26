# 實作紀錄：OST 資料夾過濾與子樹阻斷修正

## 變更內容

### 1. 補完過濾清單
在 `IsFolderFiltered` 函式中，根據使用者提供的圖片藍色標記，補全了以下項目：
- `NON_IPM_SUBTREE`
- `Drizzle`
- `ItemProcSearch`
- `SPAM Search Folder 2`
- `根資料夾 - 公用`
- `共用的資料料`
- `IPM_SUBTREE` (位於公用資料夾下之版本)
- 以及常見導航節點：`Finder`, `尋找工具`, `捷徑`, `檢視`, `一般檢視方式`

### 2. OST 子樹阻斷邏輯 [Form1_OST.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_OST.vb#L669-L691)
修改了 `BuildOstFolderTree` 的多輪 BFS 建樹過程：
- 只有當父節點已成功被記錄在 `nodeMap` 中時，子節點才會被處理。
- 如果某個節點被過濾，它就不會進入 `nodeMap`，導致其後代節點在後續輪次中因為找不到父節點對應的 `TreeNode` 而自動被阻斷，不會出現在目錄樹中。

### 3. PST 保持完整顯示 [Form1_OST.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_OST.vb#L832-L837)
根據使用者最新指令，移除了 `LoadPstSubFoldersRecursive` 中的過濾檢查，確保 PST 端的資料夾結構完整呈現。

## 驗證結果
- **OST 導航樹**：預期將不再顯示 `NON_IPM_SUBTREE`、`Drizzle` 及其子資料夾，介面將更加乾淨且集中於有效內容。
- **PST 導航樹**：不受影響，維持完整顯示。
