# F5 重新整理 SSD 快取遞迴清理實作計畫

此計畫旨在解決使用者在 Tv1/Lv1 按下 F5 重新整理時，因為 SSD 快取（SQLite 資料庫）沒有被同步清除，導致後續讀取的深層子資料夾依然命中舊快取，無法完全更新至最新資料的問題。

## 使用者審閱要求
> [!IMPORTANT]
> 1. 我們將在 SQLite 快取層 (`Form1_SQLite2.vb`) 新增一個遞迴刪除函數 `DbDeleteCachesByPathRecursive(rootPath)`，在交易 (Transaction) 內利用 `LIKE` 條件一次清除該資料夾及其所有子孫資料夾在資料庫中的所有快取記錄。
> 2. 我們會在 F5 強制重新整理的業務邏輯 (`Form1_MainTab12.vb` 中的 `ForceLv1Refresh`) 中，針對去重後的每一個選定節點，在清除記憶體快取的同時，調用該遞迴刪除函數。

## Proposed Changes

---

### SQLite 持久化快取層

#### [MODIFY] [Form1_SQLite2.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_SQLite2.vb)
- 新增 `DbDeleteCachesByPathRecursive(rootPath As String)`。
- 該函數將會以 `@p` 及 `@p & "\%"` 的 LIKE 查詢，在單一 Transaction 中刪除 `folder_stats`, `basic_maillist`, `year_counts`, `month_counts`, `attach_maillist`, `attach_filenames` 這六個資料表中，該資料夾及其所有子孫的記錄。

---

### Tab1/2 業務邏輯層

#### [MODIFY] [Form1_MainTab12.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_MainTab12.vb)
- 在 `ForceLv1Refresh()` 函數的「② 清除快取」區塊中，針對選定的去重節點調用 `DbDeleteCachesByPathRecursive(rootPath)`。

## Verification Plan

### Manual Verification
1. 在 Tab1 選取一個具有多層子資料夾的資料夾。
2. 在 Outlook 中變更該資料夾的深層子資料夾（如孫子資料夾）的郵件數量（例如刪除或移入郵件）。
3. 回到程式的 TreeView1/ListView1 焦點上按下 **F5**。
4. 觀察 Debug Log，確認是否呼叫了 `DbDeleteCachesByPathRecursive` 並成功刪除了複數筆資料。
5. 展開子資料夾至深層資料夾，確認顯示的郵件數量是否能 100% 更新為最新數量。
