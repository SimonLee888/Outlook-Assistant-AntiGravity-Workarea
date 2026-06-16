# OST/PST 目錄樹過濾邏輯修正計畫 (v3)

目前的過濾邏輯僅針對資料夾名稱進行過濾，若某個父資料夾被過濾，但其子資料夾不在過濾清單內，子資料夾仍可能出現在錯誤的位置（例如 OST 的孤兒處理邏輯）。
本計畫將修正為「整棵子樹阻斷」模式，確保被過濾節點及其後代皆不顯示。

## 使用者需求確認

> [!IMPORTANT]
> 1.  **子節點連帶過濾**：如果 Drizzle 不要顯示，則其下的所有子節點也都不要加入。
> 2.  **特定系統資料夾阻斷**：`Non_IPM_Subtree` 及其以下的所有資料夾都不要顯示。
> 3.  **過濾清單更新**：根據圖片，藍色標記的節點（如 `根資料夾 - 公用`、`~MAPISP`、`Drizzle`、`SPAM Search Folder 2`、`ItemProcSearch`、`檢視` 下的子項等）皆應移除。

## 擬定的變更

### 1. 更新 `IsFolderFiltered` [MODIFY] [Form1_OST.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_OST.vb)

-   補完過濾清單，包含 `IPM_SUBTREE`（如果是在 `根資料夾 - 公用` 下面，需要特別處理或統一過濾）。
-   根據圖片，`根目錄 - 信箱` 下面的 `一般檢視方式`、`尋找工具`、`捷徑`、`檢視` 也都應該過濾。

### 2. 修正 `BuildOstFolderTree` (OST) [MODIFY] [Form1_OST.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_OST.vb)

-   **根節點階段**：若根節點被過濾，不加入 `nodeMap`。
-   **子節點階段 (BFS)**：修改邏輯，只有當 `f.parent` **存在於 `nodeMap` 中** 且 **`f` 本身不被過濾** 時，才建立節點並加入 `nodeMap`。
-   這能保證：如果父節點被過濾（未進入 `nodeMap`），其所有子節點也會因為找不到父節點對應的 `TreeNode` 而無法被加入。

### 3. 修正 `LoadPstSubFoldersRecursive` (PST) [MODIFY] [Form1_OST.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_OST.vb)

-   目前的遞迴邏輯已經在 `For Each` 中執行 `IsFolderFiltered`。
-   因為是遞迴調用，一旦父資料夾被 `Continue For` 跳過，其遞迴子路徑自然就不會執行，已具備「阻斷子樹」的效果。
-   主要任務是確保 `IsFolderFiltered` 的清單足夠完整。

## 驗證計畫

### 1. OST 驗證
-   載入 OST 檔案。
-   檢查 TreeView 是否不再出現 `NON_IPM_SUBTREE`。
-   檢查 `Drizzle` 及其內容是否消失。
-   檢查 `根資料夾 - 公用` 是否完全消失。

### 2. PST 驗證
-   載入 PST 檔案。
-   檢查導航樹是否乾淨，無系統隱藏資料夾。
