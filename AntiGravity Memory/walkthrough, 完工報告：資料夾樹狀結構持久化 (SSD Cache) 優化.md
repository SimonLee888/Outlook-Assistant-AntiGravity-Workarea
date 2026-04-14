# 完工報告：資料夾樹狀結構持久化 (SSD Cache) 優化

這是一項重大的效能提升工程，現在你的 Outlook Assistant 已經具備了「資料夾身分證」記錄功能，重啟程式後再也不需要重複耗時的掃描。

## 🛠️ 修改內容總結

### 1. 資料庫層 (Form1_SQLite2.vb)
- **Schema 自動升級**：`folder_stats` 表格現在新增了 `entry_id`, `store_id`, `is_mail`, `has_chinese` 四個欄位。程式啟動時會自動偵測並執行 `ALTER TABLE` 遷移。
- **持久化寫入**：`SaveFolderStatsInner` 現在會一併存入資料夾的身分證與排序標記。
- **極速查詢**：
    - 新增 `DbGetSubFolderIDList`: 使用 `LIKE` 語法一次抓出整棵子樹。
    - 新增 `DbGetOrderedSubFolderIDs`: 支援「英文優先」的直屬子目錄排序查詢。

### 2. Outlook 資料層 (Form1_Outlook.vb)
- **身分證快取**：新增 `_cacheFolderIDs` 作為記憶體中的中轉站，銜接 OOM 物件與 SQL 資料。
- **DB Lazy Load 攔截**：
    - `GetSortedSubFolders` (TreeView)：優先從 SSD 讀取已排序的 ID 清單。
    - `GetSubFolderList` (BFS)：優先從 SSD 利用 `LIKE` 抓取整棵樹。

## 🧪 驗證與測試指引

> [!IMPORTANT]
> **初次運行建議：**
> 由於資料庫剛升級，裡面的 ID 欄位目前是空的。請先執行一次 **Renew Cache**，這會讓程式跑一次完整的掃描並將所有身分證寫入 SQLite。

### 測試步驟：
1. **建立地圖**：點擊 **Renew Cache**。完成後，點擊 **Save Cache**。
2. **模擬重啟**：關閉程式並重新開啟。
3. **驗證 TreeView**：
    - 點選一個從未展開過的資料夾 `[+]` 號。
    - 觀察 Dbg 視窗，應出現 `SSD Hit: ... 已從資料庫載入 ... 個子目錄`。
    - 感受展開速度（預期應為瞬間彈出）。
4. **驗證統計掃描**：
    - 執行年度統計或月份統計。
    - 觀察 Dbg 視窗，應出現 `SSD Hit (Tree): ...`。
    - 預期不再看到「正在展開資料夾結構」的長條圖緩慢跳動。

## 🛡️ 安全機制說明
- **英文優先**：排序邏輯嚴格遵循 `ORDER BY has_chinese ASC`。
- **ID 變更防護**：若 `GetFolderFromID` 失敗（例如移動跨 Store），程式會自動 Catch 並走回傳統 BFS 模式，不會當機。
- **小塊修改**：所有修改均以多個 ReplacementChunk 完成，確保檔案結構完整性。
