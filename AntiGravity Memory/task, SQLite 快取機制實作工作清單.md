# SQLite 快取機制實作工作清單

- [ ] `[ ]` 添加 SQLite NuGet 依賴項 (`Microsoft.Data.Sqlite`)
- [ ] `[ ]` 建立 `Form1_SQLite.vb` 並實作核心資料庫邏輯
    - [ ] 實作 `InitDatabase()`: 自動建立或開啟 `cache.db` 並初始化 Table
    - [ ] 實作 `SaveCachesToSQLiteAsync()`: 將 ConcurrentDictionary 併入資料庫
    - [ ] 實作 `LoadCachesFromSQLiteAsync()`: 從資料庫讀取並填回字典
- [ ] `[ ]` 修改 `Form1.vb` 以掛載 UI 按鈕事件
    - [ ] 尋找 `SaveCache` 與 `LoadCache` 按鈕並綁定點擊事件
    - [ ] 加入 `Stopwatch` 效能追蹤與 `Dbg()` 輸出
- [ ] `[ ]` 驗證與調適
    - [ ] 編譯測試
    - [ ] 手動測試存取功能
