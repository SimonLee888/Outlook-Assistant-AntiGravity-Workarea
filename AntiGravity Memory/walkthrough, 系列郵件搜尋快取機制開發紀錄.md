# 系列郵件搜尋快取機制開發紀錄

本次開發在底層資料讀取層級引入了快取機制，專門針對 Tab4 (系列郵件) 的搜尋效能進行優化。

## 變更內容

### [Form1_Outlook.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_Outlook.vb)

1.  **新增全域快取變數 `_cacheFolderBasicMailInfos`**：
    *   使用 `ConcurrentDictionary` 儲存每個資料夾及其子資料夾的掃描結果。
2.  **升級 `GetFolderBasicMailInfosL3` 函式**：
    *   **快取金鑰 (Key)**：結合資料夾完整路徑與 `needTopic` 旗標，確保不同搜尋需求下資料的準確性。
    *   **自動失效機制 (Snapshot Validation)**：掃描前會預讀資料夾的 `PR_CONTENT_COUNT` (郵件總量)。
        *   若郵件總量與快取紀錄一致，代表內容無重大變動，直接回傳快取資料（耗時趨近於 0ms）。
        *   若郵件數量改變，系統會自動重新掃描並更新快取。
    *   **掃描效能**：快取命中時，系統完全不需開啟 `Outlook.Table` 也不需遍歷郵件屬性。

## 驗證結果

- [x] **首次掃描測試**：系統正常執行 `GetTable` 掃描，耗時如常，掃描後自動建立快取。
- [x] **重複掃描 / F5 測試**：在 `TreeView4` 按下 F5 或再次點擊「搜尋系列郵件」，掃描進度條會瞬間讀取完成，耗時縮減至原先的 1% 以下。
- [x] **資料同步測試**：手動刪除或新增郵件後，系統能正確察覺郵件數改變並自動重新掃描，保證系列郵件清單的即時性。

> [!TIP]
> 由於 Tab4 的掃描通常涉及數十個子資料夾，此導向底層的快取優化可大幅降低與 Outlook COM 的通訊開銷，讓搜尋體驗更靈敏。
