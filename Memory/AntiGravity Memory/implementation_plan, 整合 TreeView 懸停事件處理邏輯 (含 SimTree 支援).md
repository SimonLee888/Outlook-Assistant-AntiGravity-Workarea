# 整合 TreeView 懸停事件處理邏輯 (含 SimTree 支援)

本計畫將 `TreeView` 的 `MouseMove` 指向與 `MouseLeave` 整合。這不但能精簡程式碼，還能解決 `MouseLeave` 可能未正確呈現 `SimTree` 選取色的問題。

## 使用者評論要求
> [!IMPORTANT]
> - **保留註解與歷程**：必須完整保留 2026-03-17 的最終版開發註解，並在邏輯整合處標註修改歷程。
> - **標記規範**：使用 `by AntiGravity, 2026/04/03` 進行標記。

## 擬議變更

### Form1.vb

#### [MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1.vb)

1.  **修改 `InitTreeView` (L301-302)**：
    *   將 `HandleTreeViewMouseMoveShared` 與 `HandleTreeViewMouseLeaveShared` 統一指向 `HandleTreeViewMouseHover`。

2.  **整合並新增 `HandleTreeViewMouseHover`**：
    *   **核心邏輯**：
        *   判斷事件類型（MouseMove 或 MouseLeave）。
        *   **還原舊節點顏色**：保留原有的 `SimTree` 選取判斷邏輯，確保離開節點（或離開 TreeView）時顏色正確還原。
        *   **套用新節點顏色**：套用 Hover 色。
    *   **完整保留原始註解**：將原本 L846-L858 的規劃註解移入新函式中。

3.  **移除舊函式**。

## 驗證計畫

### 自動化測試
*   檢查編譯。

### 手動驗證
1.  **普通 TreeView 測試**：驗證 TreeView1, 3, 4, 5 的 Hover 效果與離開效果。
2.  **SimTree2 測試**：
    *   檢查「選取中」的節點移動滑鼠時，是否不會被 Hover 色蓋過。
    *   檢查滑鼠移動到「選取中」的節點後離開，選取色是否正確保持。
    *   檢查滑鼠直接離開控制項時，選取色是否正常。
