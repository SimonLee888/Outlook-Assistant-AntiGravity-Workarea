# 修正過濾切換後 TreeView 焦點重置與 ListView 重複顯示計畫 (修訂版)

## 問題深層分析

### 1. 為什麼不能直接存下 Node 物件？
在 WinForms 中，當 `tv.Nodes.Clear()` 被執行後，所有的 `TreeNode` 物件都會被標記為銷毀或脫離狀態。即便變數引用還在，它們也無法被重新加入新生成的樹狀結構中。重新載入樹後，所有的節點都是全新的記憶體物件。因此，我們必須記錄**路徑 (String)** 或是 **EntryID**，在新的樹中重新搜尋並換取「新的節點物件」。

### 2. ListView 重複顯示的根源
`SimTree` 的私有清單 `_selectedNodes` 存儲的是舊節點的引用。當新樹建立並選取新節點時，舊引用若未清除，`SelectedNodes` 屬性會回傳「舊+新」的混合，導致統計邏輯重複執行。

---

## 擬定修復方案

### [Component] Form1_SimTree.vb (自訂控制項強化)

#### [MODIFY] [Form1_SimTree.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_SimTree.vb)
*   **強化 `SelectedNodes` 屬性**：
    *   [ ] 增加過濾邏輯：回傳清單前，先篩選掉 `node.TreeView IsNot Me` 的無效節點。這可以從根本上防止「已分離節點」干擾外部統計。
*   **同步清理邏輯**：
    *   [ ] 修改 `ClearSelectedNodes`，確保徹底清除 `_lastClickedNode` 與私有清單。

### [Component] Form1_Win32API.vb (導航工具層)

#### [NEW] `SelectNodeByPath(tv, path)` 工具函數
*   [ ] 實作一個遞迴搜尋函數，根據傳入的 `FolderPath` 字串，在 TreeView 中逐層展開並尋找對應節點。
*   [ ] 找到後呼叫 `st.SetSelectedNode(node)`。

### [Component] Form1.vb (流程控制層)

#### [MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1.vb)
*   **`CheckShowAllFolders_CheckedChanged`**:
    1.  [ ] `Dim oldPath As String = SimTree1.SelectedNode?.Tag?.ToString` (假設 Tag 存有路徑，或由 node.Tag.FolderPath 取得)。
    2.  [ ] 呼叫 `SimTree1.ClearSelectedNodes()`。
    3.  [ ] 執行 `Nodes.Clear()` 與 `LoadStoreToTreeView`。
    4.  [ ] 呼叫 `SelectNodeByPath(SimTree1, oldPath)`。
    5.  [ ] **Fallback**: 若 `SelectNodeByPath` 回傳 False（路徑被過濾了），則執行 `ExpandTreeToDefaultInbox`。

## 使用者問題回覆摘要
*   **為何不自動 Clear？** 因為原生 `Nodes.Clear` 無法簡單攔截，我們改在 `SimTree` 讀取端增加「歸屬性檢查」最安全。
*   **多選還原？** 考慮到效能與複雜度，切換模式時僅還原「最後焦點資料夾」。
*   **路徑消失？** 若還原失敗，會自動導向「收件匣」，不會讓 UI 處於無選取狀態。

## 驗證計畫
1.  **情境 A**：選取「寄件備份」後切換模式 -> 應停留在「寄件備份」。
2.  **情境 B**：選取一個「純資料夾」（會被過濾掉的）後切換模式 -> 應跳回「收件匣」。
3.  **情境 C**：觀察 ListView 標題行 -> 不應再出現重複的收件匣統計。
