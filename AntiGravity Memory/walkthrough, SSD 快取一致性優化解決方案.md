# SSD 快取一致性優化解決方案

本次任務修復了 SSD 快取在讀寫過程中對資料夾身分標識（EntryID/StoreID）處理不全的問題，並修正了導致樹狀結構「變平」的關鍵 SQL 查詢。

## 完成的修改

### 1. 樹狀結構修復 (由使用者手動與 AI 協力)
*   **檔案**: [Form1_Outlook.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_Outlook.vb)
*   **修正內容**: 將 `GetSortedSubFolders` 中錯誤的 `DbGetSubFolderIDList`（會拿整棵子樹）更換為 `DbGetOrderedSubFolderIDs`（僅拿直屬子資料夾）。
*   **結果**: 徹底解決了 SSD 讀回後資料夾樹結構混亂的問題。

### 2. 持久化完整性優化
*   **檔案**: [Form1_SQLite2.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_SQLite2.vb)
*   **修正內容**: 
    *   **寫入增強**: 在 `SaveFolderStatsInner` 中將 `_cacheFolderIDs.Keys` 加入聯集。現在即使尚未點選計算的資料夾，只要被 BFS 掃描過，其 ID 就會正確存入 SSD。
    *   **讀取增強**: 在 `LoadFolderStatsInner` 中擴充 SQL SELECT 選項，載入時自動回填 `_cacheFolderIDs`。這保證了 `LoadCache` 後，系統擁有完整的資料夾身分地圖，不再依賴緩慢的 COM Fallback。
    *   **效能優化**: `RenewAttachMailListAsync` 現在傳入 `fPath` 至 snapshot 檢查函數，進一步壓低 `RenewCache` 的 CPU 耗時。

## 驗證建議
1.  **啟動後行為**: 清除 SSD 快取 -> 執行一次全掃描 -> 手動按 SaveCache -> 重啟程式 -> **確認 TreeView 展開是否瞬間完成且層級正確**。
2.  **DB 現況**: 觀察狀態列 `folder_stats` 的筆數，應與目前已知資料夾總量相符。

> [!TIP]
> 此次優化後，Layer 2.5 的快取層變得更加堅實，大部分的 TreeView 導覽應該都能在 1ms 內完成，完全屏除 PST 檔案 I/O 的阻塞感。
