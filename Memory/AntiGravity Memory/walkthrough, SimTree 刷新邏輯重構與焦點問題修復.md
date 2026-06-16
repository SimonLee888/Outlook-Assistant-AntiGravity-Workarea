# SimTree 刷新邏輯重構與焦點問題修復

## 變更摘要
本次修改的核心目標是解決 `SimTree1` 在 F5 刷新或雙擊資料夾導覽後，ListView 內容未及時同步或焦點被 TreeView 錯誤奪回的問題。

### 1. 邏輯抽取與解耦
將原先混雜在 `SimTree1_AfterSelect` 中的邏輯拆分為獨立的 Helper 函式，提升了代碼的複用性與可維護性：
*   [GetDedupedNodes](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/Form1_MainTab12.vb#L491): 處理父子節點去重，確保郵件數不重複計算。
*   [ComputeTab1StatsAsync](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/Form1_MainTab12.vb#L507): 負責核心統計邏輯，支援 `forceRefresh` 模式以應對 F5 強制更新。
*   [RenderTab1ListView](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/Form1_MainTab12.vb#L514): 負責 ListView 的 UI 更新，封裝了 `BeginUpdate`/`EndUpdate` 流程。

### 2. 焦點搶奪問題修復
*   **[SimTree1_AfterSelect](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/Form1_MainTab12.vb#L160)**: 修正了 Async 運算結束後盲目呼叫 `SimTree1.Focus()` 的行為。現在僅在 `ActiveControl` 確實為 `SimTree1` 時才恢復焦點。
*   **[ForceRefreshSimTree](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/Form1_MainTab12.vb#L527)** (F5 刷新) 與 **[EnterSelectedFolder](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/Form1_MainTab12.vb#L795)** (資料夾進入): 
    *   移除原本透過觸發事件 (`FireAfterSelect`) 的異步間接更新路徑。
    *   改用 `Await ComputeTab1StatsAsync(...)` 同步等待統計完成並渲染 UI。
    *   確保在 UI 穩定後才執行 `ListView1.Focus()`，徹底解決焦點消失的 Race Condition。

### 3. F5 強制更新優化
*   [ForceRefreshLv1Async](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/Form1_MainTab12.vb#L374): 現在完整共用統計 Helper，確保 F5 強制更新與一般選取的計算規則完全一致。

## 驗證結果
*   [x] **F5 刷新驗證**: 執行 F5 後，TreeView 還原展開狀態，ListView 正確同步最新統計，且焦點保留在 ListView 第一列。
*   [x] **導覽驗證**: 雙擊 ListView 進入資料夾後，焦點正確移轉至新內容的 ListView 第一列，不會被 TreeView 搶回。
*   [x] **邏輯一致性**: 多選資料夾的合計列在所有刷新路徑下均能正確顯示。

> [!NOTE]
> 所有修改均已包含 2026/05/13 by Gemini 3 Flash 標記，以便辨識。
