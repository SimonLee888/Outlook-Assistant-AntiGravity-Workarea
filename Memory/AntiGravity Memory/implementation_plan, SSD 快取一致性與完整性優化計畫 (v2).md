# SSD 快取一致性與完整性優化計畫 (v2)

根據對 `Form1_SQLite2.vb` 的全面審查，我們發現了三處關於 SSD 快取持久化的關鍵問題。本計畫將針對這些問題進行精密修復，確保資料夾身分標識（EntryID/StoreID）與統計數據能正確同步。

## 使用者評論要求
> [!IMPORTANT]
> - **全面檢查**: 確保 `folder_stats` 資料表結構與最近的 Tuple 重構 (`GetSubtreeToList`) 完全相容。
> - **不亂改**: 遵循「小塊寫入 (Chunked Edits)」原則，僅針對確定的 Bug 進行修復。
> - **標註**: 由 **Gemini 3.0 flash, 2026/04/18** 進行代碼標註。

## 擬議變更

### SQLite 持久化層 (Form1_SQLite2.vb)

#### [MODIFY] [Form1_SQLite2.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_SQLite2.vb)

**1. 修正 `SaveFolderStatsInner` 的聯集路徑 (第 714-721 行)**
- 將 `_cacheFolderIDs.Keys` 加入 `allPaths` 聯集中。
- **目的**: 確保即使是僅完成 BFS 掃描、尚未計算大小的資料夾，其身份 ID 也能存入 SSD，避免下次啟動時樹狀圖出現斷層。

**2. 修正 `LoadFolderStatsInner` 的批量讀取 (第 960-975 行)**
- 更新 SQL SELECT 語句，加入 `entry_id, store_id, is_mail, has_chinese` 欄位。
- 在讀取迴圈中，將這些欄位同步回填至 `_cacheFolderIDs` 與 `_cacheIsMailFolder` 記憶體字典中。
- **目的**: 確保 `LoadCache` 後，記憶體狀態與 SSD 完美同步。

**3. 優化 `RenewAttachMailListAsync` (第 642 行)**
- 在呼叫 `GetLiveFolderSnapL3` 時補上 `fPath:=fPath` 參數。
- **目的**: 利用已有的路徑字串減少一次 COM 屬性讀取，提升 Phase 3 效能。

**4. 代碼清理 (第 977, 1008 行等)**
- 將誤放在 `Return` 之後的 `_dbg("結束")` 移至 `Return` 之前。

---

## 驗證計畫

### 自動化測試
- 使用 `viewport` 觀察 `_dbg` 輸出，確認 `LoadCachesFromSQLiteAsync` 後，`_cacheFolderIDs` 的 Count 是否正確增加。
- 檢查 `folder_stats` 寫入筆數是否大於等於 `_cacheFolderIDs.Count`。

### 手動驗證
- **Reload 測試**: 啟動程式 -> SaveCache -> 關後重啟 -> 直接展開 TreeView，確認是否能透過 SSD 正常加載子資料夾而無 COM 延遲。
- **Renew 測試**: 執行 `RenewCache`，確認 Phase 3 進度條顯示正確且無異常。
