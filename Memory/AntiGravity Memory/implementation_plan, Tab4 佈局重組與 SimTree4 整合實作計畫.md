# Tab4 佈局重組與 SimTree4 整合實作計畫

此計畫旨在優化 Tab4 (系列郵件) 的操作流程，將選取器 (`SimTree4`) 與結果顯示 (`TreeView4`, `ListView4`) 進行物理分離，建立直觀的三欄式佈置。

## 擬議變更

### [Component] Form1.vb

#### [MODIFY] [InitTab4UI](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1.vb)

- **UI 結構調整**：
    1. **初始化 SimTree4**：呼叫 `InitTreeView(SimTree4)`。
    2. **重新分配 SplitContainer4 內容**：
        - `SplitContainer4.Panel1.Controls.Clear()`。
        - 將 `SimTree4` 放入 `SplitContainer4.Panel1` 並設為 `Dock = Fill`。
    3. **建立巢狀分欄 (`scnrResults`)**：
        - 在 `SplitContainer4.Panel2` 中建立一個全屏的 `SplitContainer` (水平分割)。
        - **左側 (Panel1)**：放入 `TreeView4` (系列主題)，`Dock = Fill`。
        - **右側 (Panel2)**：放入原本的 `ListView4` 與 `pnlOptions_tab4`。
- **資料初始化 (Form1_Shown)**：
    - 將 `LoadStoreToTreeView` 的對象從 `TreeView4` 改為 `SimTree4`。

### [Component] Form1_MainTabs.vb

#### [MODIFY] [Button4_Click](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_MainTabs.vb)

- **搜尋來源切換**：
    - `rootFolder` 的來源從 `SimTree1.SelectedNode` 改為 `SimTree4.SelectedNode`。

#### [MODIFY] [TreeView4_KeyDown] (快捷鍵)

- **ESC 邏輯更新**：
    - 按下 ESC 時，不再重新載入目錄樹到 `TreeView4`。
    - 改為執行：`TreeView4.Nodes.Clear()`、`ListView4.Items.Clear()`，並將焦點還給 `SimTree4`。

## 待確認問題 (Open Questions)

> [!IMPORTANT]
> **佈局方式確認**
> 我目前規劃採用 **三欄式 (資料夾樹 | 系列主題 | 郵件清單)**。
> 目前 `SplitContainer4` 只有兩欄，我會透過在右側 Panel 中再塞入一個 SplitContainer 來達成。
> 請問這樣的橫向三欄排列是否符合你的預期？還是你希望主題與郵件清單是上下排列？

## 驗證計畫

### 手動驗證
1. **佈局檢查**：開啟 Tab4，觀察左側是否出現目錄樹，右側是否有兩塊空白區域。
2. **交互測試**：
    - 在左側 `SimTree4` 選取資料夾。
    - 點擊「搜尋系列郵件」。
    - 確認中間的 `TreeView4` 顯示搜尋出的主題。
    - 點選主題，確認右側 `ListView4` 顯示該主題的所有郵件。
3. **ESC 測試**：
    - 在搜尋結果狀態按下 ESC，確認搜尋結果被清除，且選取焦點回到左方的資料夾樹。
