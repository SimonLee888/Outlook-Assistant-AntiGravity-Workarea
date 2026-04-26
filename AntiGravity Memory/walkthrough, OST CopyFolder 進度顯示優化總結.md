# OST CopyFolder 進度顯示優化總結

本次更新為 OST 複製資料夾功能加入了專業的進度回饋機制，仿照 Tab3 的效能監控邏輯，提供即時速度與 ETA 預估。

## 變更內容

### [Form1_OST.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_OST.vb)

1. **計時器整合**：
   - 引入了 `_tab7StatusSw` (Stopwatch) 用於精確追蹤 `CopySourceDatablocksToPST` 的執行時間。
2. **進度訊息解析**：
   - 在 `CopyFolder_Click` 中注入了增強型的 `StatusMsg` 委派。
   - 使用 Regex 解析核心庫回傳的 `"Converting ... X out of Y"` 字串。
   - 自動計算 `速度 = 已處理筆數 / 流逝時間`。
3. **預估時間 (ETA)**：
   - 基於當前速度推算剩餘筆數所需的秒數，並以 `(剩餘 mm:ss)` 格式附加在進度訊息後方。
4. **自動還原機制**：
   - 使用 `Try...Finally` 結構，確保無論複製成功、失敗或中斷，`StatusMsg` 都會還原為簡約版（僅顯示原始文字），避免對「解析 OST」等其他輕量操作造成干擾。

## 驗證結果

- **解析 OST**：維持原樣，顯示 `正在解析 OST...` 等基本文字。
- **複製資料夾**：
  - `ProgressBar2` 開始跳動顯示細部進度。
  - 當處理量超過 0.1 秒後，出現例如 `(156 筆/秒 (剩餘 01:12))` 的額外資訊。
- **結束還原**：操作完成後，PB2 內容清空，委派邏輯回復。

> [!TIP]
> 目前設定為剩餘時間大於 2 秒時才顯示 ETA 字串，以避免在處理即將結束時文字劇烈跳動影響閱讀。
