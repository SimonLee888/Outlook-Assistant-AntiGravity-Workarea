# 抽出 ETA 預估剩餘秒數邏輯

目標是將專案中多次重複出現的「處理速度與預估剩餘秒數 (ETA)」計算邏輯，抽離成一個獨立且可重複使用的輔助函數。

## 思考過程 (Thinking Process)
by Gemini 3.1 Pro, 2026/05/11

1. **分析重複的程式碼結構**:
   觀察 `Form1_MainTab345.vb` 與 `Form1_Outlook.vb` 總計五處的程式碼，發現它們都執行了完全相同的步驟：
   - 取得經過的安全時間 `elapsedSec` (避免除以零)。
   - 計算處理速度 `speed` = `已處理數量 / elapsedSec`。
   - 判斷是否大於某個門檻值 (`total > 500` 或 `totalFolders > 10`)。
   - 計算剩餘時間 `remainingSec` 並格式化為 `mm:ss`。
   - 將結果輸出為字串 `etaString`，同時也需要用到 `speed` 變數來顯示目前的每秒處理量。
   
2. **設計可共用的函數**:
   - **輸入參數**: `totalItems` (總數), `processedItems` (已處理數), `elapsedSec` (經過時間), `minTotalThreshold` (啟動 ETA 計算的最低門檻，例如 500 或 10)。
   - **輸出 (修正使用 Tuple)**: 捨棄傳統的 `ByRef`，改用更現代且優雅的 `ValueTuple`，直接回傳 `(Speed As Double, EtaString As String)`。
   - **放置位置**: 建議放在 `Form1_MainTab345.vb` 的共通輔助函數區域 (例如 `#Region "  └ 輔助函數"`)。

## User Review Required

請確認改用 Tuple 後的語法與結構是否符合您的期望！如果沒問題，我們就可以開始動手寫入。

## Proposed Changes

### [MODIFY] Form1_MainTab345.vb

**1. 新增共用函數 (建議放在 `#Region "  └ 輔助函數"` 底部)**
```vbnet
    ''' <summary>
    ''' by Gemini 3.1 Pro, 2026/05/11
    ''' 計算處理速度與預估剩餘時間 (ETA)
    ''' </summary>
    ''' <param name="totalItems">總項目數</param>
    ''' <param name="processedItems">已處理項目數</param>
    ''' <param name="elapsedSec">經過時間 (秒)</param>
    ''' <param name="minTotalThreshold">計算 ETA 的最低總項目數門檻 (例如 10 或 500)</param>
    ''' <returns>Tuple 包含: Speed (項目數/秒) 與 EtaString (格式化字串)</returns>
    Private Function CalculateSpeedAndETA(totalItems As Integer, processedItems As Integer, elapsedSec As Double, minTotalThreshold As Integer) As (Speed As Double, EtaString As String)
        Dim safeElapsedSec As Double = Math.Max(elapsedSec, 0.001)
        Dim currentSpeed As Double = If(processedItems > 0, processedItems / safeElapsedSec, 0)
        
        Dim etaStr As String = ""
        If totalItems > minTotalThreshold AndAlso currentSpeed > 0 Then
            Dim remainingSec As Integer = CInt(Math.Max(0, (totalItems - processedItems) / currentSpeed))
            If remainingSec > 3 Then 
                etaStr = $"，預估剩餘 {remainingSec \ 60:D2}:{remainingSec Mod 60:D2}"
            End If
        End If
        Return (currentSpeed, etaStr)
    End Function
```

**2. 修改點 1 (Tab3 - 約 Line 381)**
**修改前:**
```vbnet
Dim elapsedSec As Double = Math.Max(swTotal.Elapsed.TotalSeconds, 0.001)
Dim speed As Double = processed / elapsedSec
Dim etaString As String = ""
If total > 500 AndAlso speed > 0 Then
    Dim remainingSec As Integer = CInt(Math.Max(0, (total - processed) / speed))
    If remainingSec > 3 Then etaString = $"，預估剩餘 {remainingSec \ 60:D2}:{remainingSec Mod 60:D2}"
End If
```
**修改後:**
```vbnet
Dim eta = CalculateSpeedAndETA(total, processed, swTotal.Elapsed.TotalSeconds, 500)
Dim speed As Double = eta.Speed
Dim etaString As String = eta.EtaString
```

**3. 修改點 2 (Tab4 - 約 Line 650)**
**修改前:**
```vbnet
Dim elapsedSec As Double = Math.Max(sw.Elapsed.TotalSeconds, 0.001)
Dim speed As Double = If(processed > 0, processed / elapsedSec, 0)
Dim etaString As String = ""
If targetFolderList.Count > 10 AndAlso speed > 0 Then
    Dim remainingSec As Integer = CInt(Math.Max(0, (targetFolderList.Count - processed) / speed))
    If remainingSec > 3 Then etaString = $"，預估剩餘 {remainingSec \ 60:D2}:{remainingSec Mod 60:D2}"
End If
```
**修改後:**
```vbnet
Dim eta = CalculateSpeedAndETA(targetFolderList.Count, processed, sw.Elapsed.TotalSeconds, 10)
Dim speed As Double = eta.Speed
Dim etaString As String = eta.EtaString
```

**4. 修改點 3 (Tab5 - 約 Line 1378)**
**修改前:**
```vbnet
Dim elapsedSec As Double = Math.Max(swTotal.Elapsed.TotalSeconds, 0.001)
Dim speed As Double = If(totalProcessed > 0, totalProcessed / elapsedSec, 0)
Dim etaString As String = ""
If totalFolders > 10 AndAlso speed > 0 Then
    Dim remainingSec As Integer = CInt(Math.Max(0, (totalFolders - totalProcessed) / speed))
    If remainingSec > 3 Then etaString = $"，預估剩餘 {remainingSec \ 60:D2}:{remainingSec Mod 60:D2}"
End If
```
**修改後:**
```vbnet
Dim eta = CalculateSpeedAndETA(totalFolders, totalProcessed, swTotal.Elapsed.TotalSeconds, 10)
Dim speed As Double = eta.Speed
Dim etaString As String = eta.EtaString
```

### [MODIFY] Form1_Outlook.vb

**5. 修改點 4 (Form1_Outlook.vb - 約 Line 831)**
**修改前:**
```vbnet
Dim elapsedSec As Double = Math.Max(swTotal.Elapsed.TotalSeconds, 0.001)
Dim speed As Double = curProcessed / elapsedSec
Dim etaString As String = ""
If total > 500 AndAlso speed > 0 Then
    Dim remainingSec As Integer = CInt(Math.Max(0, (total - curProcessed) / speed))
    If remainingSec > 3 Then etaString = $"，預估剩餘 {remainingSec \ 60:D2}:{remainingSec Mod 60:D2}"
End If
```
**修改後:**
```vbnet
Dim eta = CalculateSpeedAndETA(total, curProcessed, swTotal.Elapsed.TotalSeconds, 500)
Dim speed As Double = eta.Speed
Dim etaString As String = eta.EtaString
```

**6. 修改點 5 (Form1_Outlook.vb - 約 Line 893)**
**修改前:**
```vbnet
Dim elapsedSec As Double = Math.Max(swTotal.Elapsed.TotalSeconds, 0.001)
Dim speed As Double = curProcessed / elapsedSec
Dim etaString As String = ""
If total > 500 AndAlso speed > 0 Then
    Dim remainingSec As Integer = CInt(Math.Max(0, (total - curProcessed) / speed))
    If remainingSec > 3 Then etaString = $"，預估剩餘 {remainingSec \ 60:D2}:{remainingSec Mod 60:D2}"
End If
```
**修改後:**
```vbnet
Dim eta = CalculateSpeedAndETA(total, curProcessed, swTotal.Elapsed.TotalSeconds, 500)
Dim speed As Double = eta.Speed
Dim etaString As String = eta.EtaString
```
