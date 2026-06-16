# 修正 Tab4 SSD 快取載入缺失計畫 (v2)

目前發現 `Tab4` (系列郵件搜尋) 在重啟程式後會重新掃描，是因為雖然有「存入」SQLite (`basic_maillist`)，但在啟動後的 `LoadCachesFromSQLiteAsync` 流程中漏掉了「讀取」該表的邏輯。本計畫將補齊此環節。

## User Review Required

> [!IMPORTANT]
> **載入策略調整**
> 由於 `basic_maillist` 資料量可能較大（包含所有搜尋過資料夾的郵件摘要），我將比照 `attach_maillist` 的模式，在 `LoadCache` 時一次性載入。若資料量達到數萬封以上，載入時間可能會增加約 0.5~1 秒，但這能確保搜尋時完全不需要再動到 COM 介面。

## Proposed Changes

### [Form1_SQLite2.vb]

#### [MODIFY] [Form1_SQLite2.vb](file:///D:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_SQLite2.vb)

- **新增 `LoadFolderBasicMailInfosInner()`**：從 `basic_maillist` 表中讀取資料並填充回記憶體快取 `_cacheFolderBasicMailInfos`。
- **更新 `LoadCachesFromSQLiteAsync()`**：
  - 在 `Task.Run` 內部加入 `LoadFolderBasicMailInfosInner()` 的呼叫。
  - 將讀取到的數量顯示在最後的統計訊息中。

---

## Verification Plan

### Automated Tests
- 執行程式並點擊 **Setting -> 讀取 SSD 快取**。
- 觀察 Debug 視窗是否出現 `[LoadFolderBasicMailInfosInner] 成功載入 XXX 筆` 字樣。

### Manual Verification
1. 在 **Tab 4** 搜尋一組資料夾（確認資料夾被掃描並顯示結果）。
2. 到 **Setting** 點擊 **Save Cache**。
3. 關閉程式並重啟。
4. 到 **Setting** 點擊 **Load Cache**。
5. 回到 **Tab 4** 再次對「同樣的資料夾範圍」點擊搜尋。
6. **預期結果**：Debug 視窗應顯示 `② DB lazy load 命中`，且搜尋應在瞬間完成（不應再次出現 `L3 COM 掃描` 耗時數秒的情況）。
