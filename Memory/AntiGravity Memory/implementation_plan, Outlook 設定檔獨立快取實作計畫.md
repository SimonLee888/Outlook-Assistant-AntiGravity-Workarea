# Outlook 設定檔獨立快取實作計畫

目前 Outlook Assistant 使用單一的 `OLAcache.db` 檔案來儲存所有快取資料。當使用者切換不同的 Outlook 設定檔（如「工作」與「個人」）時，由於資料夾路徑不同，系統會將另一設定檔的資料視為「孤兒（Orphan）」並將其刪除，且不同設定檔間的統計數據也可能因路徑重複而混淆。

## 使用者審閱確認

> [!IMPORTANT]
> **變更核心：資料庫路徑動態化**
> 我們將把資料庫檔案名稱從 `OLAcache.db` 改為 `OLAcache_[ProfileName].db`。
> 這樣做的好處是：
> 1. **完全隔離**：不同設定檔的資料互不干擾。
> 2. **自動保留**：切換設定檔時，舊設定檔的資料會保留在原檔案中，不會被 `CleanupOrphanFolderPath` 誤刪。
> 3. **無需手動清理**：解決使用者提到的「每次切換就要清除快取」的問題。

## 擬定變更範圍

### [Component] 快取持久化層 (Form1_SQLite2.vb)

#### [MODIFY] [Form1_SQLite2.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_SQLite2.vb)

1.  **修改 `_dbPath` 成員**：
    *   移除 `ReadOnly` 並取消在宣告時賦值。
    *   改為在 `InitDatabase` 中根據 `_olNS.CurrentProfileName` 動態計算路徑。
2.  **更新 `InitDatabase()`**：
    *   在開啟連線前，從 `_olNS` 取得目前 Profile 名稱。
    *   組合出具備 Profile 名稱的檔案路徑（例如：`OLAcache_Work.db`）。
    *   確保舊有的 `OLAcache.db` 如果存在，可以考慮保留或提醒使用者（或者直接讓新機制生效，舊檔案就留在那當作歷史）。

---

### [Component] Outlook 初始化流程 (Form1.vb / Form1_Outlook.vb)

#### [MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1.vb)

1.  **確認 `InitDatabase()` 呼叫時機**：
    *   目前在 `Form1_Shown` 中，`InitOutlookNamespace()` 會先執行，這保證了 `_olNS` 已經就緒，可以取得 Profile 名稱。這個順序是正確的，不需要更動。

---

## 影響範圍評估 (Impact Analysis)

1.  **快取檔案量**：磁碟中會出現多個 `.db` 檔案（每個 Profile 一個）。
2.  **資料遷移**：現有的 `OLAcache.db` 資料將不會被自動轉移到新命名的 `OLAcache_[Profile].db` 中。使用者在第一次使用新版時，會需要重新執行一次 `RenewCache` 來建立該 Profile 的專屬快取。
3.  **安全性**：由於不同 Profile 的路徑完全隔離，`CleanupOrphanFolderPath` 的邏輯將變得很安全，只會刪除當前 Profile 中真正消失的資料夾。

## 驗證計畫

### 手動測試
1. 開啟 Profile A，執行 `RenewCache`，觀察是否生成 `OLAcache_A.db`。
2. 關閉程式，切換至 Profile B，執行 `RenewCache`，觀察是否生成 `OLAcache_B.db`。
3. 再次切換回 Profile A，觀察載入的是否為 Profile A 的資料，且資料並未遺失。
4. 檢查 `OLAcache_A.db` 與 `OLAcache_B.db` 的檔案大小與內容是否正確區隔。
