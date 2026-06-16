# DebugForm ListView 項目消失修復計畫

## 目的
解決當 `DebugForm` 的高度拉高、導致右側縱向卷軸 (Vertical Scrollbar) 消失時，`lvwDebug` 裡面的所有項目 (ListViewItems) 也會跟著全部消失無法顯示的嚴重 Bug。

## 根本原因分析 (Root Cause)
這是一個 WinForms `ListView` (底層 SysListView32) 廣為人知的地雷：
當視窗高度增加，項目總高度小於視窗高度時，**縱向卷軸會自動隱藏**。
卷軸隱藏的瞬間，`ListView` 的可用寬度 (`ClientSize.Width`) 會變大（多出了原本卷軸佔用的十幾像素）。
這會立即觸發 `ClientSizeChanged` 事件，進而呼叫我們的 `RecalcColumnWidths` 函數。

致命點在於：在 `RecalcColumnWidths` 中呼叫了 `lvwDebug.BeginUpdate()` 與 `EndUpdate()`。
在 **「卷軸狀態切換的生命週期」** 中（底層正在重算 Bounds 與 Scrollbars）插入 `BeginUpdate()`/`EndUpdate()`，會導致控制項內部的 GDI 繪製狀態與 Layout 進入死鎖或損毀狀態，讓 `ListView` 誤以為不需要繪製內容，結果就是所有項目視覺上全部消失（Vanish）。

`BeginUpdate`/`EndUpdate` 的設計初衷是用於「批次新增/刪除大量項目」以暫停重繪，**絕不應該**在 `Resize` 或變更欄寬的事件中使用。

## 變動事項

### [DebugForm.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/DebugForm.vb)

#### [MODIFY] `RecalcColumnWidths`
- 移除 `lvwDebug.BeginUpdate()`
- 移除 `lvwDebug.EndUpdate()`
- 移除 `Try...Finally` 區塊（因為已不需要 Ensure EndUpdate）
- 直接賦值 `lvwDebug.Columns(0).Width = newWidth` 即可。WinForms 內部在改變欄寬時本身就會觸發適當的重繪。

```vb
    Private Sub RecalcColumnWidths(sender As Object, e As EventArgs)
        ' 2026/04/01 by AntiGravity: 修正 ListView 項目在卷軸消失時跟著消失的致命 Bug
        If lvwDebug.Columns.Count < 2 Then Return
        If Math.Abs(lvwDebug.ClientSize.Width - _lastRecalcWidth) < 2 Then Return
        _lastRecalcWidth = lvwDebug.ClientSize.Width

        Dim reservedWidth As Integer = 0
        For i As Integer = 1 To lvwDebug.Columns.Count - 1
            reservedWidth += lvwDebug.Columns(i).Width
        Next

        Dim newWidth As Integer = lvwDebug.ClientSize.Width - reservedWidth - 4

        ' 絕不可以在 ClientSizeChanged 期間呼叫 BeginUpdate/EndUpdate，否則會導致 ListView 重繪機制崩潰
        If newWidth > 100 AndAlso lvwDebug.Columns(0).Width <> newWidth Then
            lvwDebug.Columns(0).Width = newWidth
        End If
    End Sub
```

## 驗證計畫

### 手動驗證
1. 啟動 `DebugForm`。
2. 讓視窗內產生數十筆資料（產生縱向卷軸）。
3. 用滑鼠將視窗向下大幅拉長，直到高度超過所有資料的總長度。
4. 觀察縱向卷軸消失的瞬間，ListViewItems 是否能**安然無恙地繼續顯示且寬度正確自動展延**。
