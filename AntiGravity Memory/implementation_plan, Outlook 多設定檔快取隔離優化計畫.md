# Outlook 多設定檔快取隔離優化計畫

## 1. 問題分析
目前 Outlook Assistant 的 SQLite 快取系統（SSD 快取）使用固定的資料庫路庫 `LocalAppData\...\OLAcache.db`。當使用者切換不同的 Outlook Profile（如「工作」與「個人」）時，會產生以下問題：
- **主鍵碰撞**：若不同 Profile 的 PST 檔內有相同路徑的資料夾，舊資料會被新 Profile 覆蓋。
- **統計誤算**：聚合統計欄位（如 `mail_count_all`）會混入不同 Profile 的數據。
- **快取一致性**：現有的 `content_count_snapshot` 僅比對郵件數，不足以區分跨 Profile 的資料庫行。

## 2. 解決方案：實體檔案隔離 (方案 A)
採取最穩健的「一設定檔一資料庫」策略。根據 `NameSpace.CurrentProfileName` 動態決定資料庫路徑。

### 優點：
1. **零碰撞風險**：不同 Profile 的資料完全物理隔離。
2. **開發成本低**：不需修改現有的 SQL Table Schema 或複雜的 Query 語法。
3. **維護方便**：`ZipAndRebuildDB` (重設快取) 將僅影響當前 Profile。

---

## 3. 預計變更範圍

### [Component] 快取初始化與路徑管理

#### [MODIFY] [Form1_SQLite2.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_SQLite2.vb)
- **`_dbPath` 成員變數**：從 `ReadOnly` 改為可變動，移除靜態初始化。
- **`InitDatabase()`**：
    - 從 `_olNS.CurrentProfileName` 取得 Profile 名稱。
    - 實作安全過濾函數，確保 Profile 名稱不含非法路徑字元。
    - 將路徑改為 `...\Cache\[ProfileName]\OLAcache.db`。
    - **重要**：必須確保在 `InitOutlookNamespace` 成功執行後才呼叫 `InitDatabase`。

#### [MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1.vb)
- **`Form1_Shown`**：確認 `InitDatabase()` 的呼叫順序位於 `InitOutlookNamespace()` 之後，以確保能獲取 Profile 名稱。

---

## 4. 實施步驟
1.  **修改變數宣告**：將 `_dbPath` 的初始化邏輯移至 `InitDatabase` 內部或一個新的 `GetDatabasePath()` 函數。
2.  **路徑動態化**：實作 Profile 名稱的安全過濾與子目錄建立邏輯。
3.  **掛載順序校對**：檢查 `Form1_Shown` 流程，確保資料庫開啟時已知 Profile。
4.  **快取清理強化**：在 `LoadCachesFromDB` 之前，明確執行 `ClearMemoryCachesCore`，防止記憶體中殘留數據。

---

## 5. 驗證計畫

### 自動化/手動驗證
1.  **環境切換測試**：
    - 以 Profile A 啟動，存入快取。
    - 關閉程式，以 Profile B 啟動。
    - **預期結果**：`OLAcache.db` 應出現在新的子目錄中，且快取計數為 0（全新資料庫）。
2.  **資料完整性測試**：
    - 再次切換回 Profile A。
    - **預期結果**：正確讀回 Profile A 的所有快取資料，且總計數字與之前一致。
3.  **安全性測試**：
    - 若 Profile 名稱包含特殊字元（如 `Simon's Mail`），確認目錄建立不會失敗。

## 6. 使用者確認事項
> [!IMPORTANT]
> 實施此計畫後，系統會為每個 Profile 建立獨立的資料庫。
> 1. 您目前放在 `...\Cache\OLAcache.db` 的**舊快取資料將不會被自動移動**，系統會重新為每個 Profile 建立新的快取。
> 2. 如果您希望保留舊資料，請告知目前的 Profile 名稱，我可以協助將其移動到正確的子目錄。

請問是否同意此計畫？
