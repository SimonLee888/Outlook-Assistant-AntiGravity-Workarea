# 單位優化與 ListView4 按鍵功能修復計畫

本計畫旨在完成資料單位轉換，並修復 `ListView4` 控制項中失去作用的按鍵功能 (`F5`, `ESC` 等)。

## Proposed Changes

### [Component] 資料夾大小單位調整 (ListView1)

#### [MODIFY] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/Form1_MainTabs.vb)
- 修改 `ComputeFolderSize` 函數 (大約 L496)。
- 計算公式由 `/ 1024` (KB) 改為 `/ 1024.0 / 1024.0` (MB)。
- 輸出格式調整為 `N2` 並加上 `" MB"` 字尾。

---

### [Component] ListView4 按鍵功能修復與增強

#### [MODIFY] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/Form1_MainTabs.vb)
- **新增 ESC 處理**：
    - 在 `ListView4_KeyDown` 中捕獲 `Keys.Escape`。
    - 觸發 `ReturnToFolderView()` (假設功能：清除搜尋框、切換 `SimTree4` 回資料夾模式)。
- **強化 F5 刷新**：
    - 確認 `RefreshListView4MailsAsync` 是否能正確取得最新的 `Outlook.MailItem` 並更新 `info.Subject` 等欄位。
- **補全 TreeView4 與 ListView4 的聯合同步**：
    - 確保在 `ListView4` 獲得焦點時，鍵盤事件不會被父容器攔截。

---

## Open Questions

- **ESC 的具體行為**：
  您希望按 `ESC` 時是「清除搜尋結果回歸資料夾樹」，還是「取消目前的計算進度」？（按照先前開發脈絡，應是回歸資料夾視圖）。
- **F5 失效現象**：
  失效是指「按下沒反應」還是「轉圈圈後數字沒變」？如果是沒反應，可能是掛載問題；如果是數字沒變，可能是 Outlook 快取尚未更新。

## Verification Plan

### 自動與手動測試
- 切換至 Tab1 計算資料夾大小，確認單位為 `xx.xx MB`。
- 切換至 Tab4，執行搜尋後，按 `ESC` 確認是否能快速回到資料夾樹狀視圖。
- 在郵件清單上按 `F5`，確認狀態列是否顯示「同步完成」，且主旨若有變動應即時反映。
