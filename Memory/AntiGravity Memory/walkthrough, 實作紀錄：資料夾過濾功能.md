# 實作紀錄：資料夾過濾功能

## 變更內容

### [Form1_OST.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_OST.vb)

1.  **新增 `IsFolderFiltered` 函式**：
    -   過濾以下名稱：`NON_IPM_SUBTREE`, `Drizzle`, `ItemProcSearch`, `SPAM Search Folder 2`。
    -   過濾 Outlook 常見系統名稱：`Finder`, `尋找工具`, `捷徑`, `檢視`, `一般檢視方式`。
    -   過濾所有以 `~` 開頭的名稱。
    -   **保留 `IPM_SUBTREE`** 以確保收件匣可見。
2.  **更新 `BuildOstFolderTree`**：在 OST 建樹的第一步（根節點）與第二步（子節點）循環中加入過濾。
3.  **更新 `LoadPstSubFoldersRecursive`**：在遞迴 PST 資料夾時同步過濾。

## 驗證建議
- 請重新載入 OST 或 PST 檔案，確認原本出現在根目錄或子目錄中的 `Drizzle`、`~MAPISP`、`尋找工具` 等資料夾是否已消失。
- 確認 `IPM_SUBTREE` 及其下的 `收件匣` 是否正常顯示。
