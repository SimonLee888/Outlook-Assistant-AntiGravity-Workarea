# 計畫：優化 Tab2 渲染與計算性能瓶頸

針對使用者提到的「快取命中仍需數百毫秒」以及渲染函數的效能問題，本計畫旨在透過減少 UI 重繪開銷、優化 Yield 頻率以及重用資源來提升整體流暢度。

## User Review Required

> [!IMPORTANT]
> **關於 Yield 頻率調整：**
> 在計算迴圈中，若檢測到資料夾大部分都命中快取，我們將大幅降低 `ThrottledYieldAsync` 的發生頻率。這會讓進度條更新變得不那麼「滑順」，但能節省 Windows Timer (15.6ms) 累積造成的顯著延遲。

> [!NOTE]
> **資源管理：**
> 將 UI 渲染中頻繁建立的 `Font` 物件改為全域唯讀物件，避免在大規模統計時造成 GC 壓力。

## Proposed Changes

### 1. UI 渲染層優化 (RenderLvXxx)

*   **預定義字型資源**：
    在 `Form1` 級別或 `Form1_MainTabs` 頂部定義靜態的 `_fontItalic` 與 `_fontBold` 物件，避免在 `RenderLvMonthView` 迴圈中每次 `New Font`。
*   **渲染節流 (Idempotent Rendering)**：
    在 `RenderLvYearView` 與 `RenderLvMonthView` 加入判斷，若傳入的數據集與上一次渲染時完全相同（例如僅是切換標籤頁回到原處），則跳過 `Clear()` 與 `Add()` 流程。
*   **Chart 渲染優化**：
    `RenderCtYearView` 與 `RenderCtMonthView` 同樣加入「內容未變動」判定，避免觸發沈重的 `Chart.Invalidate()`。

### 2. 流程協調層優化 (CollectMonthCounts / GoToMonthView)

*   **Yield 性能對策**：
    修改 `CollectMonthCounts` 中的 `ThrottledYieldAsync` 邏輯。針對大數量（如 764 個）資料夾，引入次數判定（例如每 50 個資料夾才 Yield 一次），或是在 `GetMonthCountsForYear` 快速返回時跳過 Yield。
*   **UI 狀態保持判定**：
    在 `GoToMonthView` 中，如果 `_lv2DataMonth` 命中且 UI 目前已經在該視圖，則不做任何 Render。

### 3. 計算路徑優化 - [Form1_Outlook.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_Outlook.vb)

*   **L2.5 快取存取優化**：
    在 `CollectYearCounts` 與 `CollectMonthCounts` 這種「批次」操作中，考慮暫時關閉磁碟 DB 的寫入或改用 batch 模式（若目前是逐筆寫入），減少 I/O 等待。

## Open Questions

- **是否需要針對大數量資料夾（> 500）時自動關閉動畫？** 目前 Chart 渲染較重，當資料點過多時，是否考慮簡化圖表顯示？

## Verification Plan

### Automated Tests
- 使用 `Stopwatch` 記錄 `GoToMonthView` 在「快取命中」情形下的執行時間，目標從數百 ms 降至 < 50ms。
- 監控記憶體使用量，確認 `Font` 物件數量不再隨點擊次數增加。

### Manual Verification
- 手動點擊 700+ 資料夾的年度統計，觀察進度條是否卡頓，以及切換月份視圖時是否能「秒開」。
- 快速來回切換年度/月份視圖，確認畫面不會閃爍。
