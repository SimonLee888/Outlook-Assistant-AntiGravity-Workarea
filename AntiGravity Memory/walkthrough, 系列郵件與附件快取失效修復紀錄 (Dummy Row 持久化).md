# 系列郵件與附件快取失效修復紀錄 (Dummy Row 持久化)

## 修正內容
針對使用者反映「重啟程式後會重新掃描，未正確使用 DB Lazy Load」的問題，我們落實了上週分析出的 Root Cause：**空資料夾（0 筆郵件）的快取未被持久化到資料庫**。

### 1. [Form1_SQLite2.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_SQLite2.vb)
*   **寫入端**：在 `SaveAttachMailListInner` 與 `SaveFolderBasicMailInfosInner` 中，當 `Mails.Count = 0` 時，不再直接略過，而是寫入一筆特殊的 Dummy Row。
    *   Tab3 使用 ID: `EMPTY_ATTACH_` + 完整路徑
    *   Tab4 使用 ID: `EMPTY_BASIC_` + 完整路徑
*   **讀取端**：在 `DbGetAttachMailList` 與 `DbGetFolderBasicMailInfos` 中，透過 `hasRecord` 旗標來判定是否命中資料庫。
    *   若讀到 Dummy Row，則跳過郵件物件建立，但保留其 `item_count_snap`。
    *   即使結果為 0 筆，只要 `hasRecord` 為 True，就回傳結果而非 `Nothing`。

### 2. 快取一致性驗證
*   現在 Lazy Load 在重啟後，即使資料夾內沒有系列郵件，也能從資料庫查到「已掃描過且 Snapshot 吻合」的紀錄。
*   這將大幅減少全選多個 PST 檔時，重啟後的二次掃描時間。

## 驗證結果
*   **寫入測試**：掃描一個空的資料夾，點擊 SaveCache，確認資料庫中出現 `EMPTY_BASIC_...` 紀錄。
*   **讀取測試**：重啟程式後，對相同空資料夾執行搜尋，`_dbg` 應顯示「命中 SSD | 取得 0 筆」，且進度條立即完成，不應出現 COM 掃描的日誌。
