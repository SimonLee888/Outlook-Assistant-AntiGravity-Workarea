# DebugForm 搜尋標題 (Caption) 同步功能紀錄

## 完成項目
已修復 `DebugForm` 在搜尋過程中的標題列顯示異常。

### 1. 集中化標題管理
在 `DebugForm.vb` 中新增 `UpdateSearchCaption()` 私有方法，集中處理標題文字的動態生成。
- **自動回復機制**：當 `txtDebug` 內容為空時，標題會自動回復為預設的「執行期除錯視窗」。
- **動態更新邏輯**：包含目前的搜尋關鍵字與 logic (AND/OR) 狀態。

### 2. 事件同步修正
- 修改 `txtDebug_TextChanged`：文字變動時立即更新標題。
- 修改 `checkAndOr_CheckedChanged`：切換搜尋邏輯 (AND/OR) 時也會觸發標題同步。

### 3. Region 編號修正
- 修正 `DebugForm.vb` 中重複出現兩個 `■ 06` 的標示，將輔助函數區域改為 `■ 07 輔助函數`，維護程式碼結構整潔。

---

## 異動代碼摘要

```vb
''' <summary>
''' 2026/03/31 by AntiGravity: 集中標題管理邏輯，解決文字清空未回復、及 logic 切換未同步問題
''' </summary>
Private Sub UpdateSearchCaption()
    If String.IsNullOrWhiteSpace(txtDebug.Text) Then
        Me.Text = "執行期除錯視窗"
    Else
        Dim logic As String = If(checkAndOr.Checked, "AND", "OR")
        Me.Text = $"除錯視窗 - 搜尋比對 ({logic}): {txtDebug.Text}"
    End If
End Sub
```

## 驗證結果
- [x] 搜尋框輸入文字：標題顯示「除錯視窗 - 搜尋比對 (AND): [文字]」
- [x] 切換 AND/OR：標題中的 (AND) 與 (OR) 即時變動
- [x] 清空搜尋框：標題回復為「執行期除錯視窗」
