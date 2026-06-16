# DebugForm 搜尋標題 (Caption) 修正計畫

## 目的
解決 `DebugForm` 在搜尋功能中，AND/OR 切換時標題未同步更新，以及搜尋框清除後標題未回復預設值的問題。

## 使用者確認事項
- 預設標題目前設定為「執行期除錯視窗」(依據檔案註解)。

## 變動事項

### [DebugForm.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/DebugForm.vb)

#### [MODIFY] Update Caption Logic
- 新增私有方法 `UpdateSearchCaption()`，集中處理標題更新邏輯。
- 修正 `txtDebug_TextChanged`：呼叫 `UpdateSearchCaption()`。
- 修正 `chkSearchLogic_CheckedChanged`：呼蓋 `UpdateSearchCaption()`。

```vb
Private Sub UpdateSearchCaption()
    ' by AntiGravity, 2026/03/31: 集中化標題管理，解決切換 logic 時標題沒跳動，以及清除後不回復的問題
    If String.IsNullOrWhiteSpace(txtDebug.Text) Then
        Me.Text = "執行期除錯視窗"
    Else
        Dim logic As String = If(checkAndOr.Checked, "AND", "OR")
        Me.Text = $"除錯視窗 - 搜尋比對 ({logic}): {txtDebug.Text}"
    End If
End Sub
```

## 驗證計畫

### 手動驗證 (代碼審查)
- 確認 `txtDebug_TextChanged` 有呼叫 `UpdateSearchCaption()`。
- 確認 `checkAndOr_CheckedChanged` 有呼叫 `UpdateSearchCaption()`。
- 確認搜尋框空白時標題能正確回歸「執行期除錯視窗」。
- 確認切換 AND/OR 時標題能即時顯示正確的邏輯標示。
