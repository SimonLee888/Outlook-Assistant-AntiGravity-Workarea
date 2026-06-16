# 修正過濾切換後 TreeView 焦點重置與 ListView 重複顯示計畫

## 問題分析

### 1. TreeView 焦點重置 (Issue 1)
當勾選「顯示全部資料夾」時，程式會執行以下流程：
1. `CheckShowAllFolders_CheckedChanged` 觸發。
2. `tv.Nodes.Clear()` 清空所有樹。
3. `LoadStoreToTreeView` 重新載入。
4. `ExpandTreeToDefaultInbox` 被呼叫，固定選取「收件匣」。
這導致使用者原本正在查看的資料夾（如：寄件備份）會被強行跳回收件匣。

### 2. ListView 重複顯示 (Issue 2)
這是 `SimTree` 自定義控制項的一個隱藏 Bug：
1. `SimTree` 內部維護一個私有的 `_selectedNodes` 清單，用來支援多選。
2. 當外部執行 `tv.Nodes.Clear()` 時，原生 TreeView 會移除節點，但 `SimTree` 的 `_selectedNodes` **仍然保留著對舊節點物件的引用**。
3. 重新載入樹後，如果新的「收件匣」被選取並呼叫 `AddSelectedNode`，`_selectedNodes` 內會同時存在『舊的』和『新的』收件匣節點。
4. `SimTree1_AfterSelect` 遍歷 `SelectedNodes` 時，會對同一個路徑跑兩次統計並重複加入 `ListView1`。

## 擬定修復方案

### [Component] Form1.vb (事件協調層)

#### [MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1.vb)
*   **`CheckShowAllFolders_CheckedChanged`**:
    *   [ ] 在 `Nodes.Clear()` 之前，先利用 `GetSelectedFolderPath(SimTree1)` 存起當前路徑。
    *   [ ] 對所有 `SimTree` 顯式呼叫 `.ClearSelectedNodes()`，強制清空內部 stale 引用。
    *   [ ] 重新載入樹後，優先呼叫 `SelectNodeByPath(path)`；若失敗才走 `ExpandTreeToDefaultInbox`。

### [Component] Form1_Win32API.vb (UI 工具層)

#### [NEW] `SelectNodeByPath` 與 `GetSelectedFolderPath` 工具函數
為了還原焦點，需要一套不依賴節點物件（物件在 Clear 後就失效了）而是依賴「路徑字串」的查找機制。

### [Component] Form1_SimTree.vb (自訂控制項層)

#### [MODIFY] [Form1_SimTree.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_SimTree.vb)
*   **`AddSelectedNode` / `ToggleSingleNode`**: 
    *   [ ] 增加安全檢查 `If node.TreeView IsNot Me Then Return`，防止已分離的舊節點被加入選取清單。
    *   [ ] 考慮在 `ClearSelectedNodes` 中增加更多防呆。

## 待確認事項
> [!IMPORTANT]
> 「路徑字串」在 `_showAllFolders` 不同時可能會有細微差異嗎？
> 原則上 Outlook FolderPath 是唯一的。只要模式切換後該資料夾依然存在（沒被過濾掉），就能根據路徑還原。

## 驗證計畫

### 手動測試 (由 USER 配合或模擬行為)
1.  選取一個子資料夾（非收件匣）。
2.  切換「顯示全部資料夾」。
3.  **預期 1**：TreeView 選中項應停留在原本的資料夾（若該資料夾符合顯示規則）。
4.  **預期 2**：ListView 不應出現重複的資料夾統計行。
