# Tab3 & Tab4 互動邏輯詳細整合計畫 (AddHandler 重構版)

## 修改前

### Tab3 (虛擬模式)
- `ListView3_KeyPress`: 手動檢查 Enter 和 ESC，手動讀取 `_lv3MailList` 索引。
- `ListView3_MouseDoubleClick`: 手動呼叫開啟。
- `ListView3_MouseClick`: 僅複製主旨，無路徑同步。

### Tab4 (實體模式)
- `ListView4_KeyPress`: 手動檢查 Enter，手動確認開啟數，手動 TryCast `Tag`。
- `ListView4_MouseDoubleClick`: 同上。
- `ListView4_SelectedIndexChanged` / `MouseClick`: 手動更新路徑。

---

## 修改後

### 1. `Form1_Outlook.vb` 或 `Form1_MainTabs.vb` (底層)
#### [MODIFY] `OpenMailByEntryID`
將數量檢查邏輯內聚：
```vb
Private Sub OpenMailByEntryID(entryIDs As List(Of String))
    If entryIDs Is Nothing OrElse entryIDs.Count = 0 Then Return
    ' 📢 整合數值確認 (by Gemini 3.0 flash, 2026/04/21)
    If entryIDs.Count > 10 Then
        If MessageBox.Show($"確定要同時開啟 {entryIDs.Count} 封郵件嗎？", "確認", MessageBoxButtons.YesNo) = DialogResult.No Then Return
    End If
    ' ... (原本的執行緒邏輯) ...
End Sub
```

### 2. `Form1.vb` (初始化層)
#### [MODIFY] `InitListView`
統一掛載通用事件：
```vb
Private Sub InitListView(lv As ListView)
    ' ... (外觀設定保持不變) ...
    AddHandler lv.KeyPress, AddressOf CommonListViewKeyPress
    AddHandler lv.MouseDoubleClick, AddressOf CommonListViewDoubleClick
    AddHandler lv.SelectedIndexChanged, AddressOf CommonListViewSyncPath
    AddHandler lv.MouseClick, AddressOf CommonListViewSyncPath
End Sub
```

### 3. `Form1_MainTabs.vb` (邏輯調度層)
#### [NEW] 通用輔助函數與分發器
```vb
' 提取選中的 EntryID 清單
Private Function GetSelectedEntryIDs(lv As ListView) As List(Of String)
    Dim ids As New List(Of String)
    If lv.VirtualMode Then
        For Each idx As Integer In lv.SelectedIndices
            If idx >= 0 AndAlso idx < _lv3MailList.Count Then ids.Add(_lv3MailList(idx).EntryID)
        Next
    Else
        For Each item As ListViewItem In lv.SelectedItems
            If TypeOf item.Tag Is MailItemInfo Then ids.Add(DirectCast(item.Tag, MailItemInfo).EntryID)
        Next
    End If
    Return ids
End Function

' 共通 KeyPress 處理
Private Sub CommonListViewKeyPress(sender As Object, e As KeyPressEventArgs)
    Dim lv = DirectCast(sender, ListView)
    If e.KeyChar = ChrW(Keys.Enter) Then
        OpenMailByEntryID(GetSelectedEntryIDs(lv)) : e.Handled = True
    ElseIf e.KeyChar = ChrW(Keys.Escape) Then
        ' 清除選取
        If lv.VirtualMode Then lv.SelectedIndices.Clear() Else lv.SelectedItems.Clear()
        ' Tab4 專屬：回退焦點至 SimTree4
        If lv Is ListView4 Then SimTree4.Focus()
        e.Handled = True
    End If
End Sub

' 共通雙擊開啟
Private Sub CommonListViewDoubleClick(sender As Object, e As MouseEventArgs)
    If e.Button = MouseButtons.Left Then OpenMailByEntryID(GetSelectedEntryIDs(DirectCast(sender, ListView)))
End Sub

' 共通路徑同步 (Tab3與Tab4皆受惠)
Private Sub CommonListViewSyncPath(sender As Object, e As EventArgs)
    Dim lv = DirectCast(sender, ListView)
    If lv.SelectedItems.Count > 0 Then
        Dim path As String = ""
        If lv.VirtualMode Then
            Dim idx = lv.SelectedIndices(0)
            If idx >= 0 AndAlso idx < _lv3MailList.Count Then path = _lv3MailList(idx).FolderPath
        Else
            If TypeOf lv.SelectedItems(0).Tag Is MailItemInfo Then path = DirectCast(lv.SelectedItems(0).Tag, MailItemInfo).FolderPath
        End If
        If path <> "" Then ProgressBar2.Text = path
    End If
End Sub
```

### 4. 移除冗餘
- 刪除 `ListView3_KeyPress`, `ListView4_KeyPress`, `ListView4_MouseDoubleClick` 等 `Handles` 方法。

---

## 驗證計畫
1. **全選(Ctrl+A)**: 在 Tab3 中全選，按 Enter，檢查是否彈出開啟數十封信的確認窗。
2. **Tab3 點擊同步**: 原本 Tab3 點擊沒反應，現在應與 Tab4 一樣能在下方看到資料夾路徑。
3. **ESC 回歸**: 在 Tab4 的列表按 ESC，焦點是否確實彈回 `SimTree4`。
