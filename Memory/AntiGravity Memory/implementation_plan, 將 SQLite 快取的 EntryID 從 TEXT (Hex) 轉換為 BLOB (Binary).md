# 將 SQLite 快取的 EntryID 從 TEXT (Hex) 轉換為 BLOB (Binary)

這項修改旨在減少 SQLite 資料庫空間佔用並提升 I/O 效能。由於 `EntryID` 通常長達 140 多個字元，以十六進位字串（TEXT）儲存極佔空間，將其轉換為二進位（BLOB）儲存，長度將縮減一半，同時能大幅減小 B-Tree Index 的體積。這項修改只會封裝在資料存取層 (`Form1_SQLite2.vb`)，對上層業務邏輯完全透明。

> [!TIP]
> **低風險高收益**：
> 由於我們只在存入 SQLite 之前和從 SQLite 取出之後進行「轉碼」，所有存在記憶體字典 (`_cache...`) 和 UI 層面的 `EntryID` 依然保持為原本的 `String` (Hex)，因此上層程式碼一行都不用改。

## Proposed Changes

### Form1_SQLite2.vb

這份檔案將會有幾個主要修改：

#### 1. 新增轉碼輔助函數
[NEW] 新增 `HexStringToByteArray(hex As String) As Byte()`
[NEW] 新增 `ByteArrayToHexString(bytes As Byte()) As String`

#### 2. 修改 DB Schema (`BuildSQLiteTableString`)
[MODIFY] 找到 `folder_stats`, `attach_maillist`, `basic_maillist`, `attach_filenames` 這四個資料表的 `CREATE TABLE` 語法。
將其中的 `entry_id TEXT` 更改為 `entry_id BLOB`。

#### 3. 修改寫入函數 (Save*Inner)
[MODIFY] 修改 `SaveFolderStatsInner`, `SaveAttachMailListInner`, `SaveBasicMailInfoInner`, `SaveAttachFilenamesInner` 函數。
寫入時：將 `cmd.Parameters.AddWithValue("@eid", folder.EntryID)` 改為傳入 `HexStringToByteArray(folder.EntryID)`。

#### 4. 修改讀取函數 (Load*Inner 及 DbGet*)
[MODIFY] 修改 `LoadFolderStatsInner`, `LoadAttachMailListInner`, `LoadBasicMailInfoInner`, `LoadAttachFilenamesInner` 以及所有的 `DbGet*` 單點查詢函數。
讀出時：將 `reader.GetString(x)` 或 `reader.GetValue(x)` 取出的 `BLOB` 資料轉回 `String`。
```vbnet
' 示意
Dim bytes As Byte() = DirectCast(reader("entry_id"), Byte())
Dim entryIdStr As String = ByteArrayToHexString(bytes)
```

## Verification Plan

### 安全的驗證與轉移策略
由於 SQLite 不支援直接使用 `ALTER TABLE` 更改現有欄位的型別（從 TEXT 變成 BLOB），直接套用新程式碼讀取舊的 `OLAcache.db` 可能會發生轉型錯誤 (InvalidCastException)。

**我們的驗證計畫如下：**
1. **備份舊資料**：我們將依賴你現有的 `ZipAndRebuildDB()` 機制。在部署此修改後，請先**不要**直接載入舊的快取。
2. **強制重建快取**：首次執行時，可以主動刪除 `%LocalAppData%\OutlookAssistant\Cache\...` 底下的 `OLAcache.db`，或者透過 UI 上的按鈕觸發重設。
3. **驗證檔案大小**：讓系統跑一次完整的 `RenewCacheToDB` 從 Outlook 掃描資料。觀察新的 `OLAcache.db` 檔案大小是否如預期縮小（通常會縮小 30% ~ 50%）。
4. **驗證讀寫一致性**：點擊 UI 上的其他分頁（例如 Tab3 尋找附件、Tab4 尋找重複郵件），確認各功能均能正常抓出資料並顯示正確的郵件項目，證明 BLOB 和 String 之間的轉換完全無損。

## User Review Required

> [!IMPORTANT]
> **舊有快取清除確認**
> 因為這個改動會改變底層資料庫 Schema，舊的 `.db` 檔無法直接相容。你是否同意我們在實作完成後，由你手動刪除舊的快取檔（或透過你的除錯按鈕觸發重建）來進行測試？這代表下一次開啟程式時會需要花一些時間重新快取。如果同意，請 Approve 此計畫，我會立即開始進行修改。
