# SQLite 快取機制實作計畫

本計畫旨在為 Outlook Assistant 實作 SQLite 持久化快取機制，讓 `_cacheMailCount` 等四個核心 ConcurrentDictionary 可以透過 "SaveCache" 與 "LoadCache" 按鈕存入與讀回，避免重開機後需重新掃描大資料量。

## User Review Required

> [!IMPORTANT]
> **依賴項添加**：實作 SQLite 需要引入 `Microsoft.Data.Sqlite` 或 `System.Data.SQLite`。由於目前專案為 .NET 10 (net10.0-windows)，建議使用 `Microsoft.Data.Sqlite`。
> 我將會嘗試使用 `dotnet add package Microsoft.Data.Sqlite` 來安裝。

> [!NOTE]
> **資料庫結構**：我們將建立一個名為 `cache.db` 的檔案，內含一張 `FolderCache` 表，欄位包含 `FolderPath` (主鍵), `MailCount`, `MailCountAll`, `FolderCount`, `FolderCountAll`, `UpdateTime`。

## Proposed Changes

### 專案組態 (Outlook Assistant.vbproj)

#### [MODIFY] [Outlook Assistant.vbproj](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Outlook%20Assistant.vbproj)
- 添加 `Microsoft.Data.Sqlite` PackageReference。

---

### 新增資料庫管理類別 (Form1_SQLite.vb)

#### [NEW] [Form1_SQLite.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_SQLite.vb)
- 建立 `Partial Class Form1`，負責 SQLite 的初始化、儲存與讀取邏輯。
- `InitDatabase()`: 檢查並建立資料表。
- `SaveCachesToSQLite()`: 將四個 Dictionary 的資料整合存入資料庫。
- `LoadCachesFromSQLite()`: 從資料庫讀取資料並填回 Dictionary。

---

### UI 事件掛載 (Form1.vb)

#### [MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1.vb)
- 在 `InitTab5UI` 或適當位置掛載 `SaveCache` 與 `LoadCache` 按鈕的點擊事件。

## Open Questions

- **資料庫路徑**：預設存放在程式執行目錄 (`AppDomain.CurrentDomain.BaseDirectory`) 下的 `cache.db` 是否合適？
- **同步/非同步**：讀寫資料庫是否需要使用 `Async` 以避免大資料量時卡住 UI？（建議使用 Async）

## Verification Plan

### Automated Tests
- 執行 `dotnet build` 確保編譯成功。

### Manual Verification
- 點擊 `SaveCache`，檢查是否產生 `cache.db`。
- 重開程式（或手動清空介面字典後），點擊 `LoadCache`，觀察快取是否恢復（可透過 Tab1 統計驗證是否有快取命中）。
