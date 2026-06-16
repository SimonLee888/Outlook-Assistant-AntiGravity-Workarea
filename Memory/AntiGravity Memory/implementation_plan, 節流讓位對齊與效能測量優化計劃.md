# 節流讓位對齊與效能測量優化計劃

此計劃將針對 `Form1_Outlook.vb` 進行兩項主要調整：
1. **節流邏輯對齊**：將 `GetFolderBasicMailInfosL3` (或其他類似位置) 的舊式手動節流邏輯替換為全域統一的 `ThrottledYieldAsync`。
2. **微觀 Log 優化**：移除或移動那些在快取命中情況下（耗時 < 1ms）依然會觸發的 `_dbg` 紀錄，減少不必要的字串運算開銷。

## Proposed Changes

### [Component] Outlook 資料存取層 (Form1_Outlook.vb)

---

#### [MODIFY] [Form1_Outlook.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_Outlook.vb)

##### 1. 統一節流讓位 (L2011 附近)
將手動檢查 `ElapsedMilliseconds` 與 `Task.Yield()` 的區塊替換為 `ThrottledYieldAsync`。

- **原本**:
  ```vb
  If swYield.ElapsedMilliseconds >= ThrottleFreq.Mid Then
      swYield.Restart()
      Await Task.Yield()
      ct.ThrowIfCancellationRequested()
  End If
  ```
- **更換後**:
  ```vb
  ' ✅ 使用統一節流讓位，內部 Task.Delay(1) 確保 ESC 中斷靈敏度 (by Gemini 3.0 flash, 2026/04/19)
  Await ThrottledYieldAsync(swYield, ct, ThrottleFreq.Mid)
  ```

##### 2. 微觀 Log 調整
搜尋並優化 `GetSortedSubFolders` 或 L2.5 快取層中的 `_dbg` 呼叫。

- **原則**: 如果 `TryGetValue` 成功（0ms 命中），則跳過「開始」與「結束」的 Debug Log 輸出。

## Verification Plan

### Automated Tests
- 使用 `ripgrep` 檢查是否還有殘留的手動 `Task.Yield` 節流模式。

### Manual Verification
- 執行程式並切換 Tab，確認 UI 依然靈敏，且點擊 ESC 鍵能正確中斷執行。
- 觀察 Debug Output，確認在快取命中時不再產生過於頻繁的微觀 Log。
