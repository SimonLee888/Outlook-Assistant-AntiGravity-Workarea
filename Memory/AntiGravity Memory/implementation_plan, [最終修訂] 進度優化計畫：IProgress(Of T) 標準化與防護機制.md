# [最終修訂] 進度優化計畫：IProgress(Of T) 標準化與防護機制

本計畫已納入您的最終建議，以下是實作前的最後確認。

## 1. 代碼模式對比：改了什麼？

### [以前] 傳統 Action 模式 (繁瑣且有執行緒風險)
> 需要手動處理 `Invoke`，且 L3 必須攜帶具體數值參數，邏輯較硬。

```vb
' --- L2 呼叫層 ---
Await GetMailCountAll(folder, Sub(curr, total)
    ' 必須檢查是否需要 Invoke，否則背景執行緒會崩潰
    If Me.InvokeRequired Then
        Me.BeginInvoke(Sub() lblStatus1.Text = $"掃描中: {curr}/{total}")
    Else
        lblStatus1.Text = $"掃描中: {curr}/{total}"
    End If
End Sub)

' --- L3 資料層 ---
Public Sub GetMailCountAll(folder, onProgress As Action(Of Integer, Integer))
    ' ... 迴圈 ...
    onProgress(processed, total) ' 呼叫端必須自己負責 UI 安全
End Sub
```

### [現在] IProgress 模式 (簡潔且執行緒安全)
> `Progress` 物件在 UI 執行緒建立後，會自動處理調度。L3 只管回報資料，不管 UI 怎麼畫。

```vb
' --- L1/L2 UI與流程層 ---
Dim progressHandler = New Progress(Of L3ProgressReport)(Sub(report)
    ' ⚠️ 自動進入 UI 執行緒，不需 Invoke！
    lblStatus1.Text = report.Message 
End Sub)
Await GetMailCountAll(folder, progressHandler)

' --- L3 資料層 ---
Public Async Function GetMailCountAll(folder, progress As IProgress(Of L3ProgressReport))
    ' ... 迴圈 ...
    ' 直接回報結構體，擴充性強 (可帶訊息、百分比、不確定狀態)
    progress?.Report(New L3ProgressReport With {.Message = "...", .CurrentCount = n})
End Function
```

---

## 2. 效能與防護守則 (依照您的要求)

> [!IMPORTANT]
> **1. 禁發 Debug Message**
> 為了避免 `DebugForm` 在數萬次掃描時爆炸，**嚴格禁止**在 100ms 的進度回報區塊中呼叫 `Dbg()`。
> 
> **2. 100ms 節流閥**
> 固定設定節流間隔為 **100ms**。確保 UI 既有流暢動態，又不會鎖死老舊的 WinForm 訊息幫浦。
> 
> **3. 小塊寫入與鎖定回報**
> - 修改代碼時，我會將檔案切為小塊進行 `replace_file_content`，避免大規模覆蓋。
> - **若遇到檔案被鎖定 (File Locked)**：我會立即中斷並告知您，請您協助關閉對應視窗或檔案，不會盲目重試。

---

## 3. 實作路徑 (順序)

1.  **[New]** 在 `Form1_ComL3.vb` 定義通用結構 `L3ProgressReport`。
2.  **[Modify]** 重構 `Form1_ComL3.vb` 的底層函數 (加入節流器與 IProgress)。
3.  **[Modify]** 同步修改 `Form1_Main.vb` 的 L2 協調函數與各 Tab 的 Button 點擊事件。
4.  **[Verify]** 進行橫跨 Tab 1 到 Tab 5 的效能驗證。

---

## 4. 保留歷史註解
我將嚴格遵守 `user_global` 規則，保留所有演進歷史、思考過程以及您的原始註解。

> 您可以放心，我已準備好動手了。請批准此最終版計畫。
