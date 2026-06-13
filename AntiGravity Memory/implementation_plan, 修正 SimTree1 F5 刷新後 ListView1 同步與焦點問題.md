# 修正 SimTree1 F5 刷新後 ListView1 同步與焦點問題

在 F5 刷新 `SimTree1` 時，雖然成功還原了選取狀態，但右側 `ListView1` 的內容更新是透過觸發 `AfterSelect` 事件異步執行的。這導致刷新邏輯在資料還沒載入完成前就嘗試設定焦點，且事件處理程式跑完後又會將焦點搶回 TreeView。

## 使用者評論與回饋要求
> [!IMPORTANT]
> 此修改將 `AfterSelect` 的核心邏輯抽取為獨立函式 `RefreshListView1Async`。
> 這不僅解決了 F5 的問題，也一併修正了「進入資料夾 (Enter/Double-Click)」後焦點競爭的潛在問題。

## 修改內容

### [Component] Form1_MainTab12.vb

#### [MODIFY] [Form1_MainTab12.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_MainTab12.vb)

- **抽取統計邏輯**：將 `SimTree1_AfterSelect` 內部的 BFS 統計與 ListView 填充邏輯抽取至新函式 `RefreshListView1Async()`。
- **解耦焦點控制**：從 `RefreshListView1Async` 中移除最後一行的 `SimTree1.Focus()`，改由各個呼叫端自行決定焦點去向。
- **修正 `ForceRefreshSimTree`**：
    - 使用 `Await RefreshListView1Async()` 取代 `FireAfterSelect`。
    - 確保在統計完成後才執行 `ListView1.Focus()`。
    - 移除原本失效的 TODO 程式碼。
- **修正 `EnterSelectedFolder`**：
    - 同樣使用 `Await RefreshListView1Async()` 確保導覽後 ListView 內容與選取正確顯示。

## 驗證計畫

### 手動測試 (請使用者驗證)
1. 在 Tab1 選取一個資料夾，然後按 F5。
2. 確認 `SimTree1` 刷新後，右側 `ListView1` 是否正確顯示該資料夾內容。
3. 確認焦點是否正確保持在 `ListView1` 的第一項（依據目前程式碼意圖）。
4. 測試多選資料夾後按 F5，確認合計列與統計是否正確還原。
5. 測試雙擊 ListView 項目進入子資料夾，確認焦點是否正確落在 ListView 而非跳回 TreeView。
