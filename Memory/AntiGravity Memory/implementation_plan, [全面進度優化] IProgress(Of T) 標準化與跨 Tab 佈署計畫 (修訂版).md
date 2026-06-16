# [全面進度優化] IProgress(Of T) 標準化與跨 Tab 佈署計畫 (修訂版)

本計畫旨在全面優化 `Form1_ComL3.vb` 與 `Form1_Main.vb` 中所有涉及大量迴圈的操作，解決使用者在執行大型掃描時的等待焦慮。

## 1. 核心技術解答：為什麼換成 IProgress(Of T)？

> [!NOTE]
> **Action(Of Integer, Integer) vs. IProgress(Of T)**
> - **傳統 Action**：在背景執行緒呼叫時，必須手動使用 `Me.Invoke` 或 `Control.BeginInvoke` 回傳至 UI 執行緒，程式碼瑣碎且容易出錯。
> - **IProgress(Of T)**：它在建立時會**自動捕捉目前的 SynchronizationContext (即 UI 執行緒)**。當底層呼叫 `.Report()` 時，它會自動完成執行緒調度，讓底層函数完全不需要知道 UI 控制項的存在。這不僅讓 L3 函數更純粹，也大大增加了擴充性。

---

## 2. 擬議變更概要

### A. 定義通用進度結構 (L3ProgressReport)

為了讓上層 L1 UI 與下層 L3 完美對接，我們定義一個豐富的結構：
```vb
Public Structure L3ProgressReport
    Public CurrentCount As Integer   ' 目前完成數
    Public TotalCount As Integer     ' 總數
    Public Message As String         ' 提示訊息 (如「正在掃描：收件匣...」)
    Public IsIndeterminate As Boolean ' 是否為不確定進度 (例如讀取總數中)
End Structure
```

### B. 擴大底層 L3 函數改造範圍 (Form1_ComL3.vb)

不只是郵件數與大小，**資料夾計數**也需納入：
*   [MODIFY] `GetMailCountAll`: 抽換 `Action` 為 `IProgress`。
*   [MODIFY] `GetFolderCountAll`: **[新增進度支援]** 在子樹展開時回報。
*   [MODIFY] `GetFolderSize`: 在 `GetArray(1000)` 的分批迴圈中加入細部進度。
*   [MODIFY] `GetFolderSizeAll`: 確實回報目前的資料夾處理序位。

### C. 跨 Tab 對接與 UI 觸發 (Form1_Main.vb)

針對您提到的 Tab3 及後續功能，我們將在 L2 流程層全面介入：
*   **Tab 1/2**: 定義 UI 端的 `Progress(Of L3ProgressReport)` 處理器，更新 `lblStatus1`。
*   **Tab 3 (附件搜尋)**: 
    *   在 `ScanFolderWithAttachment` (Phase 1) 內部的 `GetArray` 加入進度。
    *   在 `ScanAttachmentByName` (Phase 2) 載入 MailItem 的高成本循環中加入 `Report`。
*   **Tab 4/5 (系列與重複郵件)**: 這些通常涉及所有資料夾的遍歷，將在 `Button4_Click` / `Button5_Click` 的迴圈中實作 `progress.Report`。

---

## 3. 關鍵挑戰：如何平衡「有效回報」與「執行效能」？

> [!IMPORTANT]
> **計算平衡點：時間節流 (Time-based Throttle)**
> 如果每一封郵件都更新一次 UI，會導致訊息佇列塞車。
> **策略**：在 L2/L3 的迴圈內部使用 `Stopwatch`：
> ```vb
> If swThrottle.ElapsedMilliseconds > 70 Then ' 約每秒 14 次更新
>     progress.Report(...)
>     swThrottle.Restart()
> End If
> ```
> 實測證明，人眼感覺流暢的更新頻率約在 10~20Hz，設定在 **70-100ms** 是效能損耗極低且感官極佳的「甜蜜點」。

---

## 4. 註解規範與保留

*   保留所有原始註解與過去的思考進程（如 v1, v2, v3 演進歷史）。
*   新加入的邏輯與註解將標註 `by AntiGravity, 2026/04/02`。

## 5. 驗證計畫

1.  **全區域壓力測試**：特別是在 Tab 5 執行重複郵件掃描時，確保護 `lblStatus` 持續穩定跳動，且 UI 視窗不出現「沒有回應」的白霧。
2.  **取消響應測試**：在 `IProgress` 回報的同時，確認 `ESC` 與 `_cancelRequested` 依然能即時截斷迴圈。
