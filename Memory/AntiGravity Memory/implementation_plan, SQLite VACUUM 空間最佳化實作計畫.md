# SQLite VACUUM 空間最佳化實作計畫

這個計畫旨在專案中實作 SQLite 的 `VACUUM` 指令，用以清除資料庫中被刪除資料的殘留冗餘空間，藉此有效縮減資料庫檔案的實際大小。

## 為什麼需要 VACUUM？

> [!NOTE]
> 當 SQLite 刪除資料或表格時，預設只會將空間標記為可用，並不會真的把檔案縮小。`VACUUM` 會重新建構整個資料庫檔案，把那些冗餘的空白完全釋放，對於長期頻繁更新或刪除的資料庫非常有效。

## User Review Required / Open Questions

> [!IMPORTANT]
> 關於 UI 的呈現，有以下問題需要您確認，以便後續實作：
> 1. **觸發方式**：您希望透過介面上新增一個按鈕來手動執行（例如：在「設定」或「資料庫管理」相關的頁籤），還是希望在程式啟動/關閉等特定時間點自動執行？
> 2. **按鈕位置**：如果採用手動按鈕，目前設計放在哪個 Form 或哪個 Tab 頁面最為合適？

## Proposed Changes

### 資料庫核心模組

#### [MODIFY] [Form1_SQLite2.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_SQLite2.vb)
- 新增 `CompactDatabase()` 方法。
- 在這個方法中實作 `SqliteCommand("VACUUM", _db).ExecuteNonQuery()`。
- （可選功能）計算執行前與執行後的 `sync.ffs_db` 檔案大小，計算出共清理了多少 MB 的空間，回傳給 UI 顯示，提供明確的成效回饋。

### UI 介面模組 (視您的回覆而定)

#### [MODIFY] [Form1_DebugForm.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_DebugForm.vb) 或其他主要的 UI 檔案
- 加入按鈕事件處理，呼叫上述新增的 `CompactDatabase()` 方法，並用 MessageBox 呈現壓縮結果。

## Verification Plan

### Manual Verification
- 寫入並刪除大量測試資料，使資料庫產生空隙。
- 觀察執行 `VACUUM` 功能前後，資料夾中的 `sync.ffs_db` 檔案大小變化。
- 確保壓縮完成後，原有的資料讀取及操作皆正常無誤。
