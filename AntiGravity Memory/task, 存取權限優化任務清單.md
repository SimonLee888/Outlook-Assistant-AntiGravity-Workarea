# 存取權限優化任務清單

- [x] 優化 `Form1_Outlook.vb`
    - [x] 將 `MailItemInfo` 改為 `Private`
    - [x] 將 `ProgressReport` 改為 `Private`
    - [x] 將 `PreloadAttachByRDOAsync1/2` 改為 `Private`
- [x] 優化 `Form1_SQLite2.vb`
    - [x] 將 `FolderStatsDbRow` / `MailwithAttachsDbResult` 改為 `Private`
    - [x] 將資料庫初始化與關閉方法改為 `Private` (`InitDatabase`, `CloseDatabase`)
    - [x] 將快取存取核心方法改為 `Private` (`SaveCachesToSQLiteAsync`, `LoadCachesFromSQLiteAsync`)
    - [x] 將輔助方法改為 `Private` (`GetDatabaseSummary`, `CleanupOrphanFolderPath`)
- [x] 優化 `Form1.vb`
    - [x] 將 `StatusHistoryItem` 改為 `Private`
- [/] 編譯驗證與測試
