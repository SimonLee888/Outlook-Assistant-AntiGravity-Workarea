# 優化視窗縮放與搜尋結果呈現邏輯

針對使用者反應的 `AutoResizeLvColumns` 觸發過於頻繁的問題，已完成以下優化：

## 主要改動

### 1. 視窗狀態偵測與節流機制整合 (`Form1.vb`)
-   **優化 `Form1_Resize`**：現在偵測到 `WindowState` 改變（最大化/還原）時，會透過 `HandleLvResize` 將事件轉發給 100ms 節流計時器處理，而不是直接插隊呼叫。
-   **強化 `HandleLvResize`**：加入 `Static Dictionary` 記錄各個 ListView 的上一次執行寬度。若寬度未變（例如僅改變高度），則直接忽略事件。
-   **終極防護 `AutoResizeLvColumns`**：在真正執行昂貴的 UI 重算前，再次確認寬度是否有實質變化，避免 `_dbg` 紀錄中出現無效的執行紀錄。

### 2. 搜尋流程精簡 (`Form1_MainTab345.vb`)
-   **移除冗餘重繪**：在 `ShowLv3Result` 中移除手動的 `ListView3.Invalidate()`。虛擬模式下設定 `VirtualListSize` 本身就會引發重繪，移除多餘呼叫可避免誘發不必要的佈局事件。

## 驗證結果
-   **最大化操作**：預期 `AutoResizeLvColumns` 只會顯示一次「開始/結束」紀錄。
-   **滑鼠拉動視窗**：在連續拉動過程中不觸發，放手後 100ms 執行一次。
-   **搜尋操作**：搜尋結束後不再伴隨多次 Resize 紀錄。

## 程式碼複檢
-   已確認 `lastLvWidths` 與 `lastProcessedWidths` 使用 `Static` 變數正確持久化狀態。
-   已確認修改點前後無遺留多餘的註解或調試程式碼。
