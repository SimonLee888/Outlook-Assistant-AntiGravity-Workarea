# 自動選取搜尋結果實作計畫

本計畫旨在優化 Tab4 (系列郵件) 的使用者體驗。在按下 `Button4` 搜尋完成後，若搜尋結果不為空，程式將自動選取第一個結果節點並將焦點移至該 TreeView 控制項，讓使用者可以立即開始瀏覽郵件系列。

## 使用者評論要求
無特殊設計變更，此為 UI 互動流暢度優化。

## 擬議變更

### Form1_MainTabs

#### [MODIFY] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_MainTabs.vb)

修改 `RenderTab4Groups` 函數，在結束 `EndUpdate()` 後加入選取邏輯。

```vb
    Private Sub RenderTab4Groups(topicDict As Dictionary(Of String, List(Of MailItemInfo)))
        ' ... 前略 ...
        SimTree4.EndUpdate()
        
        ' ✅ by Gemini 3.0 flash, 2026/04/21: 搜尋完成後，自動選取第一個結果並 Focus
        If SimTree4.Nodes.Count > 0 Then
            SimTree4.SelectedNode = SimTree4.Nodes(0)
            SimTree4.Focus()
        End If
        
        ProgressBar1.Text = $"找到 {SimTree4.Nodes.Count} 個系列 (排序: {If(_tab4SortGroupsByCount, "數量", "主旨")})"
    End Sub
```

## 開放問題
無。

## 驗證計畫

### 手動測試
1. 切換至 Tab4。
2. 點選資料夾並按下搜尋 (Button4)。
3. 確認搜尋完成後，`SimTree4` 的第一個節點是否被藍色背景選取。
4. 確認焦點是否已在 `SimTree4` (可直接用鍵盤上下鍵移動)。
