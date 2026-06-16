# 統一郵件快照總數 (Item/Content Count Snapshot) 名稱重構計畫

目前在 `Form1_SQLite2.vb` 和 `Form1_Outlook.vb` 中，用來記錄「資料夾內郵件總數的快照值」的名稱非常不一致。這個值主要是對應 Outlook MAPI 屬性 `PR_CONTENT_COUNT`，但在程式碼和資料庫中使用了多種不同的命名。

## User Review Required

> [!IMPORTANT]
> 請確認您偏好使用的**統一名稱**。本計畫預設將所有相關變數、屬性及資料庫欄位統一為以 **`ItemCountSnap` (變數/屬性) / `item_count_snap` (資料庫)** 為主，但如果您更喜歡 `ContentCountSnap / content_count_snap` 也可以修改。
> 另外，因為這涉及修改 SQLite 資料庫的欄位名稱 (如 `folder_info` 的 `content_count_snap` -> `item_count_snap`)，舊的資料庫在修改 schema 和程式碼後，如果無法用 SQLite 輕易重構欄位，是否能接受直接重建該資料表（或需要幫您撰寫 `ALTER TABLE` 的資料遷移邏輯）？

## 現行名稱盤點

以下是目前散落在各處的相關名稱：

### 1. Outlook MAPI 常數
- `PR_CONTENT_COUNT` (此為微軟 MAPI 屬性的標準名稱對應常數，內容為 URL `"http://schemas.microsoft.com/mapi/proptag/0x36020003"`，**建議不改**)

### 2. 資料庫欄位 (SQLite)
- **`folder_info` 資料表**：`content_count_snap`
- **`attach_maillist` 資料表**：`item_count_snap`
- **`basic_maillist` 資料表**：`item_count_snap`

### 3. VB.NET 結構 / 類別屬性 (Properties)
- `FolderStatsDbRow.snap` / `FolderStatsDbRow.Snap`
- `FolderCacheTab3.ItemCountSnap`
- Tuple 回傳值常被定義為 `(Mails As ..., Snap As Integer)` 或 `(Mails As ..., Snap As Long)`

### 4. 區域變數 (Local Variables)
- `snap` (用來接收資料庫的 snapshot 值)
- `liveSnap` (用來接收當下從 Outlook 取得的 snapshot 值)
- `content_count_snap` (註解或某些函數參數)

---

## Proposed Changes (預定修改計畫)

若您同意，我將會執行以下修改，將名稱全面統一為 **`ItemCountSnap` (駝峰式)** 及 **`item_count_snap` (底線式)**。

### 1. 資料庫層面 (Form1_SQLite2.vb)

- **SQL 建立表句法重構**：
  - 將 `folder_info` 表的 `content_count_snap` 更改為 `item_count_snap`。
- **SQL CRUD 指令修改**：
  - 將所有 `INSERT`, `UPDATE`, `SELECT` 中對 `folder_info` 表使用到 `content_count_snap` 的地方，全部改為 `item_count_snap`。
  - 將所有 SQLite Parameter 如 `@snap`，統一改為 `@itemCountSnap`。

### 2. 類別與屬性層面 (Form1_SQLite2.vb / Form1_Outlook.vb)

- **FolderStatsDbRow 等結構**：
  - 將 `Public snap As Long = -1` 與 `Public Snap As Long = -1` 重新命名為 `Public ItemCountSnap As Long = -1`。
- **匿名 Tuple 回傳值**：
  - 例如 `DbGetBasicMailInfo` 的回傳型別 `(Mails As List(...), Snap As Integer)`，改為 `(Mails As List(...), ItemCountSnap As Integer)`。
- **內部屬性**：
  - `FolderCacheTab3.ItemCountSnap` (已是標準名稱，維持不變，並將其餘向此看齊)。

### 3. 區域變數與註解層面

- 將代表 Outlook 當下快照的變數 `liveSnap` 統一改為 `liveItemCountSnap`。
- 將代表已儲存快照的變數 `snap` 統一改為 `savedItemCountSnap` 或 `itemCountSnap`。
- 修正相關的註解，例如把 `' content_count_snap = ...'` 改為 `' item_count_snap = ...'`，確保程式碼與註解一致。

## Verification Plan

### Automated Tests
- 方案儲存後，編譯專案確認所有型別與變數名稱參考是否正確（Visual Studio 應當無編譯錯誤）。

### Manual Verification
- 請在系統中手動刪除舊的 `sync.ffs_db` (若無需保留舊快照) 以讓系統以新的 Schema 重建資料庫，或是由我為您加入 `ALTER TABLE` 來相容舊資料庫。
- 啟動 Outlook Assistant，確認載入資料夾時不報資料庫結構錯誤，且信件同步快照 (Snapshot) 功能運作正常。

> [!TIP]
> 如果您想把統一的名稱改成 `ContentCountSnap` / `content_count_snap` 也是可以的，請您在這份計畫中做決定。決定後，我會按照您的指示自動對所有檔案進行替換。
