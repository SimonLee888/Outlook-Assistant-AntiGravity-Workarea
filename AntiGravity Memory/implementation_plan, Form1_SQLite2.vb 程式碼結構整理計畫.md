# Form1_SQLite2.vb 程式碼結構整理計畫

本計畫旨在透過 `#Region` 與 `#End Region` 標籤，將 `Form1_SQLite2.vb` 模組化的程式碼結構顯性化，並優化部分函數的排列順序，以提升長期維護的便利性。

## 使用者評論請求

> [!NOTE]
> 本次修改主要為結構整理，不涉及核心業務邏輯的變更。

## 擬議變更

### Form1_SQLite2.vb

#### [MODIFY] [Form1_SQLite2.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_SQLite2.vb)

1.  **加入 Region 標籤**：
    *   **Region 1: 基礎結構與資料庫生命週期管理**：包含變數、Row 類別定義、`InitDatabase`、`CloseDatabase`、`GetCreateTablesSql`、`GetDatabaseSummary`、`ResetSSDatabaseAsync`。
    *   **Region 2: 快取主控流程**：包含 `SaveCachesToSQLiteAsync`、`LoadCachesFromSQLiteAsync`、`RenewCacheAsync`、`RenewAttachMailListAsync`、`CleanupOrphanFolderPath`。
    *   **Region 3: 批次寫入核心**：包含所有 `Save...Inner` 方法。
    *   **Region 4: 批次載入核心**：包含所有 `Load...Inner` 方法，並將 `LoadFolderBasicMailInfosInner` 移至 `LoadAttachFilenamesInner` 之前，以對齊寫入端的順序。
    *   **Region 5: Layer 2.5 即時查詢 (Lazy SELECT Helpers)**：包含所有 `DbGet...` 方法。
    *   **Region 6: Layer 2.5 即時寫入 (Lazy UPSERT Helpers)**：包含 `DbSaveMonthCountsSingle`。

2.  **調整說明註解**：
    *   在每個 Region 開頭標記修改紀錄：`by Gemini 3 Flash, 2026/04/25`。

## 驗證計畫

### 自動化測試
*   使用 `view_file` 複檢程式碼行數與 `#Region` 配對是否正確。

### 手動驗證
*   確認專案可以正常編譯。
*   確認 IDE 的折疊功能（Folding）能正確運作。
