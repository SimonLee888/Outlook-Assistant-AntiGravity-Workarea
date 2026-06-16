# UI 回歸修復完成：TreeView 焦點還原與 ListView 重複行消除

本任務成功解決了在切換「顯示全部資料夾」過濾器後，UI 出現的行為異常與資料重複問題。

## 解決的主要問題

### 1. TreeView 點選位置自動還原
**現象**：原本正在點選「寄件備份」或其它子資料夾，一勾選「顯示全部資料夾」，樹狀圖就會自動跳回最頂端的「收件匣」。
**修復方案**：
*   實作了 `GetSelectedFolderPath` 與 `SelectNodeByPath` 工具函數。
*   在樹狀圖被重載（Nodes.Clear）前，先記憶當前路徑字串。
*   重載後透過字串搜尋重新選回該資料夾，並自動處裡動態載入。

### 2. ListView 資料行重複累加
**現象**：切換過濾後，ListView 同時顯示了多個重複的資料夾統計行（例如看到兩次「收件匣」）。
**修復方案**：
*   **SimTree 自我修復**：修改 `SelectedNodes` 屬性，在回傳前會自動過濾掉不屬於當前 TreeView 的無效節點（Stale Nodes）。這解決了 `Nodes.Clear()` 無法主動通知自訂清單清理的架構問題。
*   **顯式狀態重置**：在事件處理程序中，明確對所有樹狀圖呼叫 `ClearSelectedNodes()`，確保內部變數清空。

---

## 關鍵程式碼異動

### [MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1.vb)
重構過濾切換事件，整合「路徑備份 -> 完整清理 -> 重載 -> 路徑還原」流程：

```vb
' 修改後的核心逻辑片段
Dim oldPath As String = GetSelectedFolderPath(SimTree1)
For Each tv In GetAllTreeViews(Me)
    Dim st = TryCast(tv, SimTree)
    st?.ClearSelectedNodes() ' 解決重複行問題的關鍵
    tv.Nodes.Clear()
Next
LoadStoreToTreeView(_pstStoreList, SimTree1)
If Not SelectNodeByPath(SimTree1, oldPath) Then
    ExpandTreeToDefaultInbox(SimTree1)
End If
```

### [MODIFY] [Form1_SimTree.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_SimTree.vb)
提升自訂控制項的魯棒性：

```vb
Public ReadOnly Property SelectedNodes As List(Of TreeNode)
    Get
        ' 過濾掉因為單純從 Nodes 集合移除但引用仍殘存在 List 中的無效節點
        _selectedNodes.RemoveAll(Function(n) n.TreeView IsNot Me)
        Return _selectedNodes
    End Get
End Property
```

---

## 驗證結果

### 預期行為確認
*   [x] **焦點還原**：在任何資料夾點選後切換過濾，UI 焦點均能準確留在原處。
*   [x] **資料正確性**：切換後 ListView 僅顯示一組資料夾統計，重複行現象消失。
*   [x] **Fallback 機制**：若原本選取的資料夾在切換模式後被過濾掉，系統能正確自動選回「收件匣」。

> [!TIP]
> 此次修改特別優化了 `SelectNodeByPathRecursive`，它能自動偵測並觸發未展開資料夾的動態載入，確保隱藏路徑也能被正確選回。
