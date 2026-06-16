# 修正非同步函數 (Async) 的 Stack Trace 追蹤問題 [v2]

解決 `DebugForm` 中 `WhoCallsMe()` 無法識別 `ComputeYearCounts`、`ShowYearView`、`GetMonthCountsForYear` 等非同步函數的問題。

## User Review Required

> [!IMPORTANT]
> 此變更對 Release 版本**效能影響為 0** (屬 Conditional("DEBUG"))。在 Debug 模式下，因 `st.GetFrame` 不需要捕捉行號 (False)，速度會比目前的程式碼更快。

## Core Implementation Code

```vb
    Private Function WhoCallsMe(Optional skipLevels As Integer = 1) As String
        ' ================================================================
        ' WhoCallsMe: 解析呼叫者字串，支援 Async 狀態機
        ' 2026-03-30 by AntiGravity: 增強對 Async 狀態機的解析，效能優化
        ' ================================================================
        Dim st As New StackTrace(skipLevels + 1, False) ' False 不抓行號，速度更快
        For i As Integer = 0 To st.FrameCount - 1
            Dim m = st.GetFrame(i)?.GetMethod()
            If m Is Nothing OrElse m.DeclaringType Is Nothing Then Continue For
            
            Dim typeName As String = m.DeclaringType.Name
            
            ' 排除 Debug 相關類別與函數
            If m.DeclaringType Is GetType(DebugForm) OrElse m.Name.Contains("Dbg") Then Continue For
            
            ' ✅ Async 狀態機偵測: 狀態機類別通常命名為 "<FunctionName>d__XX"
            If typeName.StartsWith("<") AndAlso typeName.Contains(">") Then
                 ' 提取 < 與 > 之間的原始函數名
                Dim methodName As String = typeName.Substring(1, typeName.IndexOf(">") - 1)
                ' 狀態機是巢狀類別，其父類別 DeclaringType 就是原來的 Form1
                Dim parentType = m.DeclaringType.DeclaringType
                Return If(parentType IsNot Nothing, $"{parentType.Name}.{methodName}", $"{typeName}.{methodName}")
            End If

            ' 排除一般的底層 MoveNext
            If m.Name = "MoveNext" Then Continue For

            Return $"{m.DeclaringType.Name}.{m.Name}"
        Next
        Return "Unknown Method"
    End Function
```

## Proposed Changes

### [DebugForm]

#### [MODIFY] [DebugForm.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/DebugForm.vb)
更新 `WhoCallsMe` 函數。

---

### [Form1]

#### [MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/Form1.vb)
同步更新 `Form1` 內的 `WhoCallsMe` 輔助函數。

## Verification Plan

### Manual Verification
1. 啟動程式並進入 Tab2。
2. 選取資料夾觸發 `ComputeYearCounts`。
3. 觀察 `DebugForm` 中的日誌，確認呼叫者欄位是否從 "Unknown Method" 變為 "Form1.ComputeYearCounts"。
