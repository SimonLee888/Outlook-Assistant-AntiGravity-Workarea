# 修正 SimTree1 F5 刷新後 ListView1 同步與焦點問題

在 F5 刷新 `SimTree1` 時，雖然成功還原了選取狀態，但右側 `ListView1` 的內容更新是透過觸發 `AfterSelect` 事件異步執行的。這導致刷新邏輯在資料還沒載入完成前就嘗試設定焦點，且事件處理程式跑完後又會將焦點搶回 TreeView。

## 使用者評論與回饋要求
> [!IMPORTANT]
> 此修改將 `AfterSelect` 的核心邏輯抽取為獨立函式 `RefreshListView1Async`。
> 這不僅解決了 F5 的問題，也一併修正了「進入資料夾 (Enter/Double-Click)」後焦點競爭的潛在問題。

## 修改內容

### [Component] Form1_MainTab12.vb

#### [MODIFY] [Form1_MainTab12.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_MainTab12.vb)

- **[NEW] `GetDedupedNodes`**：抽離父子去重邏輯，供 `AfterSelect` 與 `F5` 共用。
- **[NEW] `ComputeTab1StatsAsync`**：負責核心統計邏輯，回傳 `ListViewItem` 清單。支援 `forceRefresh` 參數以供 `ForceRefreshLv1Async` 使用。
- **[NEW] `RenderTab1ListView`**：負責 ListView 的 UI 更新作業 (BeginUpdate/Clear/AddRange/EndUpdate)。
- **修正 `SimTree1_AfterSelect`**：呼叫上述函式，並在完成後管理焦點。
- **修正 `ForceRefreshSimTree`**：確保在還原狀態後，等待統計渲染完成再設定焦點。
- **優化 `ForceRefreshLv1Async`**：移除重複的去重代碼，改用統一的 `GetDedupedNodes` 與統計流程。
- **修正 `EnterSelectedFolder`**：解決導覽後的焦點競爭。

## 驗證計畫

### 手動測試 (請使用者驗證)
1. 在 Tab1 選取一個資料夾，然後按 F5。
2. 確認 `SimTree1` 刷新後，右側 `ListView1` 是否正確顯示該資料夾內容。
3. 確認焦點是否正確保持在 `ListView1` 的第一項（依據目前程式碼意圖）。
4. 測試多選資料夾後按 F5，確認合計列與統計是否正確還原。
5. 測試雙擊 ListView 項目進入子資料夾，確認焦點是否正確落在 ListView 而非跳回 TreeView。
