# 非同步函數 Stack Trace 追蹤優化

成功修正了 `DebugForm` 無法正確識別 `Async` 函數呼叫來源的問題。

## 變更項目

### 1. 增強 WhoCallsMe 邏輯 (DebugForm.vb & Form1.vb)
更新後的邏輯不再只是跳過 `MoveNext`，而是會主動偵測是否處於 Async 狀態機環境。

```vb
' 關鍵解析邏輯範例
If typeName.StartsWith("<") AndAlso typeName.Contains(">") Then
    ' 提取 < 與 > 之間的原始函數名
    Dim methodName As String = typeName.Substring(1, typeName.IndexOf(">") - 1)
    ' 取得狀態機所屬的原始類別 (例如 Form1)
    Dim parentType = m.DeclaringType.DeclaringType
    Return If(parentType IsNot Nothing, $"{parentType.Name}.{methodName}", $"{typeName}.{methodName}")
End If
```

### 2. 效能優化
將 `StackTrace` 修改為不抓取檔案行號與列號，這能大幅減少每次呼叫 `Dbg()` 時的系統開銷：
- 舊版：`New StackTrace(skipLevels + 1, True)` (慢)
- 新版：`New StackTrace(skipLevels + 1, False)` (快)

## 驗證細節

- **[x] DebugForm 更新**: 已完成，支援解析狀態機。
- **[x] Form1 同步更新**: 已完成，邏輯保持一致。
- **[x] 效能確認**: 已確認 Release 模式下無任何開銷，Debug 模式下追蹤速度提升。

現在您在 `DebugForm` 的「呼叫者」欄位中，應該能看到清晰的 `Form1.ComputeYearCounts` 等路徑，而非先前的 "Unknown Method"。

> [!TIP]
> 如果未來有更多自定義的 Wrapper 函數需要跳過，只需在 `WhoCallsMe` 的過濾條件中加入對應的關鍵字即可。
