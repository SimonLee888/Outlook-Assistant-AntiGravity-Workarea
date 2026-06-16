# 變更說明：遷移 SQLite 資料庫路徑

為了解決 SQLite 讀寫與 Dropbox 同步機制衝突導致的「檔案總管卡頓」、「游標繞圈圈」以及「檔案鎖定」問題，我已將快取資料庫的位置遷移。

## 變更項目

### [Form1_SQLite2.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/Form1_SQLite2.vb)

1.  **修改 `_dbPath` 定義**：
    *   從原本的 `Application.StartupPath` (屬於 Dropbox 目錄) 移至 `Environment.SpecialFolder.LocalApplicationData` (`AppData\Local`)。
    *   新路徑為：`C:\Users\<User>\AppData\Local\OutlookAssistant\OLAcache.db`。

2.  **增強 `InitDatabase` 初始化邏輯**：
    *   在開啟連線前，主動檢查並建立 `OutlookAssistant` 資料夾，確保不會因為目錄不存在而引發錯誤。

## 預期效果
- **消除 Explorer 卡頓**：Dropbox 再也不會監控資料庫的讀寫與暫存檔。
- **解決連線鎖定**：避免 Dropbox 在背景偷偷讀取資料庫導致程式無法獨佔寫入。

> [!TIP]
> 舊的 `OLAcache.db` 檔案仍留在原本的專案 `bin\Debug` 目錄下，雖然程式不會再使用它，但若有需要舊資料，可以手動將其刪除以節省空間。
