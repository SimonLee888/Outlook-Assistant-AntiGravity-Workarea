# 實作紀錄：資料夾過濾功能

## 變更內容

### [Form1_OST.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_OST.vb)

1.  **更新 `IsFolderFiltered` 函式**：
    -   新增過濾：`根資料夾 - 公用`、`共用的資料料`。
    -   保留原有：`NON_IPM_SUBTREE`、`Drizzle`、`~` 開頭等。
2.  **實作整棵子樹過濾 (Subtree Filtering)**：
    -   **OST (`BuildOstFolderTree`)**：
        -   當父資料夾被過濾時，不將其加入 `nodeMap`。
        -   後續的 BFS 輪次中，子資料夾會因為 `nodeMap.ContainsKey(f.parent)` 為 False 而被判定為掛載失敗，最終落入 `pending`。
        -   在第三步處理 `pending` (孤兒) 時，同樣執行 `IsFolderFiltered` 檢查，確保整棵被阻斷的子樹都不會被掛載到目錄樹中。
    -   **PST (`LoadPstSubFoldersRecursive`)**：
        -   若目前資料夾符合過濾條件，直接 `Continue For`。這會阻斷後續對該資料夾子目錄的遞迴呼叫，實現整棵子樹的過濾。

## 驗證建議
- 請重新載入 OST 或 PST 檔案，確認原本出現在根目錄或子目錄中的 `Drizzle`、`~MAPISP`、`尋找工具` 等資料夾是否已消失。
- 確認 `IPM_SUBTREE` 及其下的 `收件匣` 是否正常顯示。
