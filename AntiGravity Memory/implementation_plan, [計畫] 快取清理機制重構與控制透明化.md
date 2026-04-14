# [計畫] 快取清理機制重構與控制透明化

使用者希望能精確控制快取的清理，並避免背景自動遷移 Schema 導致的不可控變數。本計畫將實作三路清理邏輯，並強化清理時的視覺提示。

## User Review Required

> [!IMPORTANT]
> **Schema 策略變更：**
> 我將移除 `InitDatabase` 中的自動 `ALTER TABLE` 與 `DROP TABLE` 邏輯。這意味著如果未來 Schema 變更，使用者必須手動執行「清除 SSD 並重建」來對齊新版程式。這能保證開發期資料庫狀態的絕對確定性。

> [!TIP]
> **清理選項設計：**
> 點擊「清除快取」時，將會彈出對話框提供以下選項：
> 1. **僅記憶體 (Memory)**: `.Clear()` 所有字典，不影響檔案。
> 2. **僅 SSD (SSD Only)**: 關閉連線並刪除 `.db` 檔，然後重新建立 Schema。
> 3. **兩者皆清 (Full Clean)**: 以上兩者同時執行。

## Proposed Changes

### [Form1_SQLite2.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_SQLite2.vb)

#### [MODIFY] `InitDatabase`
- 移除所有 `ALTER TABLE` 與 `DROP TABLE` 的自動偵測邏輯。
- 僅保留 `CREATE TABLE IF NOT EXISTS`。

#### [NEW] `GetLastSaveTime()`
- 執行 `SELECT MAX(updated_at) FROM folder_stats` 並回傳字串。
- 用於在清理提醒中顯示最後儲存時間。

#### [NEW] `ResetSSDatabase()`
- 1. 呼叫 `CloseDatabase()`。
- 2. 徹底刪除 `OLAcache.db`。
- 3. 呼叫 `InitDatabase()` 重建。

---

### [Form1.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1.vb)

#### [MODIFY] `ClearCache_Click`
- 1. 抓取 `GetLastSaveTime()`。
- 2. 彈出客製化 Messagebox 或 `TaskDialog`。
- 3. 根據選擇執行清理路徑。

#### [NEW] `ClearMemoryCachesInternal()`
- 將原本散落在 `ClearCache_Click` 中的 `Clear()` 邏輯封裝，方便重複調用。

## Verification Plan

### Manual Verification
1. **驗證警告資訊**：按下按鈕，確認是否正確顯示了資料庫最後儲存的時間。
2. **驗證 SSD 清理**：選擇清理 SSD，檢查編譯目錄下的 `OLAcache.db` 是否被刪除並重新產生。
3. **驗證記憶體清理**：選擇清理記憶體，確認 TreeView 展開是否變慢（因為必須重走 SSD 或 COM）。
