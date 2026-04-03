# 優化 TreeView 與 ListView 事件處理

透過集中初始化，移除大量重複的事件轉發函式，並保留現有的 `Shared` 邏輯以利維護。

## 建議變更

### Form1.vb

#### [MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%29%A6%E5%8D%80)/Form1.vb)

- **優化 `InitTreeViews`**: 
  - 遍歷所有 `TreeView` 與 `SimTree` 實例。
  - 集中連結 `MouseMove`、`MouseLeave`、`KeyPress`、`BeforeExpand` 事件到對應的 `Shared` 處理函式。
- **優化 `InitListViews`**:
  - 遍歷所有 `ListView` 實例。
  - 集中連結 `MouseMove`、`MouseLeave`、`KeyPress` 事件到對應的 `Shared` 處理函式。
  - **[NEW]** 針對動態建立的 `ListView5` 設定 `Dock = DockStyle.Fill` 以確保縮放效能。
- **優化動態 TreeView 佈局**:
  - **[NEW]** 針對 `SimTree2` 設定 `Dock = DockStyle.Fill`，使其自動填滿 `SplitContainer2.Panel1`。
- **移除冗餘 Handles 子程序**:
  - 刪除所有原本僅用於轉發事件到 `Shared` 函式的子程序（約 40 個）。


```vb
' 修改後的 InitTreeViews 範例：
Private Sub InitTreeViews(defaultFont As Font)
    ' ... (現有性質設定)
    For Each tv As TreeView In {TreeView1, TreeView2, TreeView3, TreeView4, TreeView5, SimTree1, SimTree2}
        ' ... (原本的屬性設定)
        AddHandler tv.MouseMove, AddressOf HandleTreeViewMouseMoveShared
        AddHandler tv.MouseLeave, AddressOf HandleTreeViewMouseLeaveShared
        AddHandler tv.KeyPress, AddressOf HandleTreeViewKeyPressShared
        AddHandler tv.BeforeExpand, AddressOf LoadSubFolderToTreeView
    Next
End Sub
```

## 驗證計畫

### 手動驗證
1. **導航測試**: 在每個分頁的 TreeView 與 ListView 中使用 Enter 展開、Escape 返回、Space 切換，確認導覽邏輯依然正確。
2. **視覺回饋測試**: 移動滑鼠到 TreeView 節點與 ListView 項目上，確認 Hover 背景色變換功能正常。
3. **功能測試**: 點擊 `+` 號展開資料夾，確認 `BeforeExpand` (LoadSubFolderToTreeView) 能正確載入子資料夾。
