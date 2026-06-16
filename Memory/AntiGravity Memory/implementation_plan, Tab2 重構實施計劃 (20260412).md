# Tab2 重構實施計劃 (2026/04/12)

本計劃旨在根據 `tab2_refactor_changes_04121400.md` 的規範，對 `Form1_MainTabs.vb` 中的 Tab2 功能進行深度重構。這將實現職責分離（UI/Render/Logic），優化快取機制，並提升使用者在年度與月份視圖切換時的流暢度。

## 使用者校閱 (User Review Required)

> [!IMPORTANT]
> 本次修改涉及 Tab2 核心邏輯的完整替換。原本的 `ShowYearView` 與 `ShowMonthView` 等函數將被刪除，功能將被整合（Inline）至事件處理器中，並配合新的 Render 層函數執行。這是一次破壞性的重構，旨在提升代碼質量，但需確保所有原本的標註與歷史記錄被正確保留。

## 擬定變更 (Proposed Changes)

### [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_MainTabs.vb)

#### [MODIFY] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_MainTabs.vb)

1.  **成員變數更新 (Step ①)**:
    *   刪除 `_tab2CachedYearCounts`。
    *   新增 `_lv2DataYear` (年度快取) 與 `_lv2DataMonth` (月份快取)。

2.  **Tab2 Region 註解替換 (Step ②)**:
    *   替換 line 636 附近的結構化說明註解，明確定義 Layer 1, 1.5, 2, 3 的職責。

3.  **SimTree2_AfterSelect 邏輯替換 (Step ③)**:
    *   將結果處理邏輯由原本的 `ShowResultTab2` / `ShowProgressTab2` 替換為新的 Render 函數與 Inline 進度顯示。
    *   將結果存入新快取 `_lv2DataYear`。

4.  **ListView2_MouseDoubleClick 邏輯替換 (Step ④)**:
    *   實作新版「返回年度統計」與「進入月份視圖」的邏輯。
    *   導入「方案 A」快取檢查：若進入相同年份，直接使用 `_lv2DataMonth` 渲染而不重新計算。

5.  **SelectedIndexChanged 註解更新 (Step ⑤)**:
    *   修正對 Render 函數的參照註解。

6.  **函數清理 (Step ⑥)**:
    *   刪除舊有的過時函數：`ShowYearView`, `ShowMonthView`, `ShowResultTab2`, `UpdateChart2forYearView`, `UpdateChart2forMonthView`, `ShowProgressTab2`。

7.  **新增 Render 與計算函數 (Step ⑦)**:
    *   新增 `RenderLvwYearView`, `RenderChart2YearView` (年度 Render)。
    *   新增 `RenderLvwMonthView`, `RenderChart2MonthView` (月份 Render)。
    *   新增 `CollectMonthCounts` (Layer 2 月份計算邏輯)。

## 驗證計畫 (Verification Plan)

### 自動化測試 (Automated Tests)
*   完成關鍵修改後，主動使用 `view_file` 讀取受影響的程式碼行，確認變數名稱對齊（例如 `_lv2DataYear` 的使用）與邏輯一致。

### 手動驗證 (Manual Verification)
*   請使用者依照以下情境操作 Tab2：
    1.  點選資料夾節點，確認年度統計結果正確顯示且圖表同步。
    2.  雙擊單一年份，確認能快速切換至月份視圖，且包含「← 返回年度統計」按鈕。
    3.  雙擊返回按鈕，確認回到年度視圖且選取游標停留在剛才進入的年份上。
    4.  快速切換年份（方案 A 測試）：進入 A 年後返回，再進入 A 年，觀察是否瞬間渲染（無進度條）。
    5.  ESC 中斷測試：在計算年份或月份分佈時按下 ESC，確認能優雅停止且不發生例外錯誤。

## 開放性問題 (Open Questions)
*   無。 md 文件的重構藍圖非常明確。

by Gemini 3 Flash, 2026/04/12
