# Outlook Assistant — Session Memory　2026/06/21 03:58

本檔兩大部分：
- **Part 1**：本次已完成的 `RenewCacheToDB` 死列(失效 entryID) purge 修復 — 檢查結論 + 完整 task list + 可重套用 changeset。
- **Part 2**：交給「另一個對話」執行的 DB 切分 handoff — 把 `attach_filenames` 從 `OLAcache.db` 搬入 `OLAsimhash.db`（與 SimHash 表同檔），含完整接觸點地圖與和 Part 1 的耦合點。

兩件事**強耦合**：Part 1 的 purge 會 DELETE `attach_filenames`；Part 2 搬走該表後，那一行 DELETE 必須改連線。務必先套 Part 1、再做 Part 2（理由見 §2.4）。

---

# Part 1 — RenewCacheToDB 死列 purge 修復【已完成】

## 1.1 檢查結論（問題本質）

整個快取有效性設計押在一個不變量上：**「同一資料夾的 basic/attach 列必須共用單一 snap」**（reader 端 `DbGetBasicMailInfo`/`DbGetAttachMailList` 不以 snap 過濾 SQL，只 `WHERE folder_hash=?` 全讀，回傳的 snap = SQLite 疊代到的**最後一列**之 snap，見 1641/1717 行「每行都一樣，讀最後一次即可」假設）。

**死列來源（破壞不變量的兩個入口）：**
- 四個逐封郵件寫入器全是 `INSERT OR REPLACE`（只增不刪），唯一 DELETE 是「整夾消失」才觸發的 `CleanupOrphanPath`。
- 入口 (a)：`RenewCacheToDB` 狀況 A 清了記憶體卻沒清 DB。
- 入口 (b)：任何 proxy ③ 重算後 `SaveCache` 只 `INSERT OR REPLACE`，死列原地不動。

**後果（真 bug，非僅佔空間）：** 資料夾還在、刪掉幾封信後，DB 出現「fresh 列(新 snap) + 死列(舊 snap)」混 snap 狀態。下次讀取依 SQLite 列序非決定性地二選一：
- 最後一列是新列 → snap 命中 → **採信整包 → 幽靈/重複郵件流進 Tab4/Tab5**；
- 最後一列是死列 → snap 不符 → **每次都 COM 重算，快取對該夾永久空轉**。

## 1.2 完整改動清單（Task List）— 8 處，全部位於 `Module_SQLite2.vb`

| # | 位置（搜尋錨點） | 動作 | 狀態 |
|---|---|---|---|
| 1 | `RenewCacheToDB` 檔頭 `' 流程：` | 改寫流程/設計邊界，補 06/20 註解（保留 04/09 等舊署名） | ✅ |
| 2 | 狀況 A `_cacheBasicMailInfo.TryRemove(fPath` 之後 | 補清 attach/month 記憶體 + 呼叫 `DbPurgeFolderMailRows` | ✅ |
| 3 | step 7 `For Each ancestor In GetAncestors(p)` | 改 `GetAncestors(p).Prepend(p)`、`ancestor`→`node`（補入自身 All 聚合） | ✅ |
| 4 | `Private Async Function RenewAttachMailList`（死碼） | 整段刪除，原位置改放新 helper `DbPurgeFolderMailRows` | ✅ |
| 5 | `SaveBasicMailInfoInner` 的 `_dbg("開始")` 後 | 加寫入邊界 purge（INSERT 前先刪 basic_maillist 該夾舊列） | ✅ |
| 6 | `SaveAttachMailListInner` 的 `_dbg("開始")` 後 | 加寫入邊界 purge（attach_maillist） | ✅ |
| 7a/7b | `SaveMonthCountsInner` 的 `Using cmd...sql` / `For Each mo In kvp.Value` | 加 delCmd、逐 `(folder_hash, year)` 先刪後插 | ✅ |
| — | 檔頭索引註解 line 21/22 | RenewCacheToDB 改「精確打擊」描述；刪 `RenewAttachMailList` 索引行 | ✅ |

> 套用順序無關緊要（依內容比對），但建議**先做 #4（新增 `DbPurgeFolderMailRows`）**，#2 才有對象可呼叫。
> Hash 一致性已驗證：`FolderPathToHash64(p)` 內部即 `StringToXxHash64(p)`，與 basic/attach/month/attach_filenames 寫入/讀取/CleanupOrphanPath 全部同值 → purge DELETE 必對得到列。

## 1.3 可重套用 changeset

> 因 `/mnt/project/` 為唯讀副本、編輯不回寫，以下為各處 old→new，供在 Visual Studio 手動套用。署名日期統一用 `2026/06/20`（與當時對話一致；如要改 06/21 自行全域替換即可）。

### #4（先做）— 刪 `RenewAttachMailList`，原位置改放 `DbPurgeFolderMailRows`
**移除整段** `Private Async Function RenewAttachMailList(folder As Outlook.Folder, fPath As String) As Task ... End Function`，**換成：**
```vbnet
    Private Sub DbPurgeFolderMailRows(fPath As String, Optional includeAttachFilenames As Boolean = False)
        ' ---------------------------------------------------------------
        ' DbPurgeFolderMailRows — 刪除單一資料夾在逐封郵件表的全部列 (basic_maillist/attach_maillist/month_counts，
        '   選擇性含 attach_filenames)。用於「資料夾還在但內含郵件有增刪」時清掉死列(失效 entryID)，
        '   維持「同一資料夾的 basic/attach 列共用單一 snap」不變量，根除讀取端混 snap 幽靈郵件。
        '   與 CleanupOrphanPath 的差異：那個是整夾消失才連 folder_stats 全表一起刪；本函式只清逐封郵件列，不動 folder_stats。
        ' 2026/06/20 by Simon/Claude: 取代原 RenewAttachMailList(三路比對) 死碼
        ' ---------------------------------------------------------------
        If _db Is Nothing Then Return
        Dim fh = FolderPathToHash64(fPath)
        Try
            Using txn As SqliteTransaction = _db.BeginTransaction()
                For Each tbl In {"basic_maillist", "attach_maillist", "month_counts"}
                    Using c As New SqliteCommand($"DELETE FROM {tbl} WHERE folder_hash=@fh", _db, txn)
                        c.Parameters.AddWithValue("@fh", fh) : c.ExecuteNonQuery()
                    End Using
                Next
                If includeAttachFilenames Then
                    Using c As New SqliteCommand("DELETE FROM attach_filenames WHERE folder_hash=@fh", _db, txn)
                        c.Parameters.AddWithValue("@fh", fh) : c.ExecuteNonQuery()
                    End Using
                End If
                txn.Commit()
            End Using
        Catch ex As System.Exception
            _dbg("DbPurgeFolderMailRows 錯誤", $"{fPath}: {ex.Message}")
        End Try
    End Sub
```
> ⚠️ **Part 2 改點所在**：上面 `includeAttachFilenames` 那段 DELETE 用的是 `_db`。attach_filenames 搬到 OLAsimhash.db 後，這段要改連 `_dbSim`、且必須**獨立交易**（不能掛在 `_db` 的 txn 上）。詳見 §2.4。

### #2 — 狀況 A 補齊失效集合
**old：**
```vbnet
                    _cacheYearCounts.TryRemove(fPath, Nothing)
                    _cacheBasicMailInfo.TryRemove(fPath, Nothing)
                    _cacheFolderIDs(fPath) = (folder.EntryID, folder.StoreID, IsMailFolder(folder, fPath), True)
```
**new：**
```vbnet
                    _cacheYearCounts.TryRemove(fPath, Nothing)
                    _cacheBasicMailInfo.TryRemove(fPath, Nothing)
                    ' 2026/06/20 by Simon/Claude: 補齊失效集合 — 清 attach/month 記憶體 + purge 該夾全部逐封郵件 DB 列，
                    '   修復「資料夾還在但郵件有刪」殘留死列(失效 entryID)，使下次 lazy 重算後該夾維持單一 snap、不再出現幽靈郵件
                    _cacheAttachMailList.TryRemove(fPath, Nothing)
                    For Each mk In _cacheMonthCounts.Keys.Where(Function(k) k.StartsWith(fPath & "_")).ToList()
                        _cacheMonthCounts.TryRemove(mk, Nothing)
                    Next
                    DbPurgeFolderMailRows(fPath, includeAttachFilenames:=True)
                    _cacheFolderIDs(fPath) = (folder.EntryID, folder.StoreID, IsMailFolder(folder, fPath), True)
```

### #3 — step 7 補入自身 All 聚合
**old：**
```vbnet
                For Each p In updatedPaths
                    For Each ancestor In GetAncestors(p)
                        _cacheMailCountAll.TryRemove(ancestor, Nothing)
                        _cacheMailCountAll.TryRemove(ancestor & "|True", Nothing)
                        _cacheMailCountAll.TryRemove(ancestor & "|False", Nothing)
                        _cacheFolderCountAll.TryRemove(ancestor, Nothing)
                        _cacheFolderSizeAll.TryRemove(ancestor, Nothing)
                    Next
                Next
```
**new：**
```vbnet
                For Each p In updatedPaths
                    ' 2026/06/20 by Simon/Claude: GetAncestors 不含自身，但變動夾「自己」的 All 聚合也會 stale，故 Prepend(p) 一併失效
                    For Each node In GetAncestors(p).Prepend(p)
                        _cacheMailCountAll.TryRemove(node, Nothing)
                        _cacheMailCountAll.TryRemove(node & "|True", Nothing)
                        _cacheMailCountAll.TryRemove(node & "|False", Nothing)
                        _cacheFolderCountAll.TryRemove(node, Nothing)
                        _cacheFolderSizeAll.TryRemove(node, Nothing)
                    Next
                Next
```

### #5 — SaveBasicMailInfoInner 寫入邊界 purge
在 `SaveBasicMailInfoInner` 的 `_dbg("開始")` 與 `Dim sql = "INSERT OR REPLACE INTO basic_maillist" &` 之間插入：
```vbnet
        ' 2026/06/20 by Simon/Claude: 寫入邊界 purge — 對本次要寫入的每個資料夾，INSERT 前先刪 basic_maillist 全部舊列，
        '   保證該夾最終只含記憶體當前清單(單一 snap)，根除死列(失效 entryID)殘留造成的混 snap 幽靈郵件
        Using delCmd As New SqliteCommand("DELETE FROM basic_maillist WHERE folder_hash=@fh", _db, txn)
            delCmd.Parameters.Add("@fh", SqliteType.Integer)
            For Each fpDel In _cacheBasicMailInfo.Keys
                delCmd.Parameters("@fh").Value = FolderPathToHash64(fpDel) : delCmd.ExecuteNonQuery()
            Next
        End Using
```

### #6 — SaveAttachMailListInner 寫入邊界 purge
在 `SaveAttachMailListInner` 的 `_dbg("開始")` 與 `Dim sql = "INSERT OR REPLACE INTO attach_maillist" &` 之間插入：
```vbnet
        ' 2026/06/20 by Simon/Claude: 寫入邊界 purge — INSERT 前先刪 attach_maillist 全部舊列 (理由同 SaveBasicMailInfoInner)
        Using delCmd As New SqliteCommand("DELETE FROM attach_maillist WHERE folder_hash=@fh", _db, txn)
            delCmd.Parameters.Add("@fh", SqliteType.Integer)
            For Each fpDel In _cacheAttachMailList.Keys
                delCmd.Parameters("@fh").Value = FolderPathToHash64(fpDel) : delCmd.ExecuteNonQuery()
            Next
        End Using
```

### #7a — SaveMonthCountsInner 多行 Using + delCmd 參數
**old：**
```vbnet
        Using cmd As New SqliteCommand(sql, _db, txn)
            cmd.Parameters.Add("@fh", SqliteType.Integer)
            cmd.Parameters.Add("@yr", SqliteType.Integer)
            cmd.Parameters.Add("@mo", SqliteType.Integer)
            cmd.Parameters.Add("@cnt", SqliteType.Integer)
```
**new：**
```vbnet
        Using cmd As New SqliteCommand(sql, _db, txn),
              delCmd As New SqliteCommand("DELETE FROM month_counts WHERE folder_hash=@fh AND year=@yr", _db, txn)
            cmd.Parameters.Add("@fh", SqliteType.Integer)
            cmd.Parameters.Add("@yr", SqliteType.Integer)
            cmd.Parameters.Add("@mo", SqliteType.Integer)
            cmd.Parameters.Add("@cnt", SqliteType.Integer)
            ' 2026/06/20 by Simon/Claude: 逐 (folder_hash, year) 先刪後插，避免整個 month bucket 被清空後留 phantom count
            delCmd.Parameters.Add("@fh", SqliteType.Integer)
            delCmd.Parameters.Add("@yr", SqliteType.Integer)
```

### #7b — SaveMonthCountsInner 內層迴圈前先刪
**old：**
```vbnet
                For Each mo In kvp.Value
                    cmd.Parameters("@fh").Value = FolderPathToHash64(fPath)
```
**new：**
```vbnet
                ' 2026/06/20 by Simon/Claude: 本 (folder,year) 先刪舊列(下面再插當前月值)，確保整體覆蓋而非殘留
                delCmd.Parameters("@fh").Value = FolderPathToHash64(fPath)
                delCmd.Parameters("@yr").Value = yearVal
                delCmd.ExecuteNonQuery()
                For Each mo In kvp.Value
                    cmd.Parameters("@fh").Value = FolderPathToHash64(fPath)
```

### #1 與索引註解
- `RenewCacheToDB` 檔頭：把舊「Phase 1~6 / BFS」流程描述改寫成「精確打擊 + purge」現況，保留 `2026/04/09 by Claude`、`04/16`、`05/17`、`6/7` 等所有舊署名，文末新增一行 `2026/06/20 by Simon/Claude` 記錄本次 purge 修正。
- 檔頭索引註解：line 21 `RenewCacheToDB` 由「Phase1~6 完整更新」改為「精確打擊更新 (2026-04-09 新增，2026-05-17 改 GetFolderFromID 模式)」；line 22 `RenewAttachMailList` 索引行整行刪除。

## 1.4 已確認的設計邊界 / 不修項目
- **attach_filenames 寫入邊界刻意不 purge**：它無讀取端幽靈 bug（Tab3 先從已清乾淨的 attach_maillist 取 entryID，再 by entryID 查檔名，死 entryID 永不被查到），且重建最貴（逐封開信）。僅在狀況 A（`includeAttachFilenames:=True`）與整夾消失（CleanupOrphanPath）時清。
- **snap 偵測盲區**：`GetLiveFolderSnapL3` = PR_CONTENT_COUNT。「刪一封又收一封」淨值不變 → snap 不變 → 狀況 B → 不重算。已知限制，不修，僅註解標明。

---

# Part 2 — DB 切分 handoff：attach_filenames → OLAsimhash.db【交另一對話執行】

## 2.1 目標與理由
把昂貴的 `attach_filenames`（逐封開信枚舉附件才能重建）從 `OLAcache.db` 搬入既有的 `OLAsimhash.db`，與同樣昂貴（要讀內文算 SimHash）的 `mail_simhash` 表同檔共存。動機一致：**兩者都「內容不變即永久有效」、重建成本極高**，應與「會被清快取 / ZipAndRebuildDB 重置」的 `OLAcache.db` 物理隔離，讓它們在 rebuild 後存活。

## 2.2 現況：兩個 db 的連線與結構（皆在 `Module_SQLite2.vb`）

**主 db（OLAcache.db）**
- 連線：`_db As SqliteConnection`，路徑 `_dbPath`（line 112）。
- `InitDatabase()` 建所有主表，含 `attach_filenames`（line 268）。
- `SaveCachesToDB()` 把多表包在**單一 `_db` 交易**內批次寫入。

**SimHash 獨立 db（OLAsimhash.db）** — `#Region "■ Fuzzy 模糊比對專用區塊"`（line 2249 起）
- 連線：`_dbSim As SqliteConnection`，路徑 `_dbSimPath`（line 2259，與 `_dbPath` 同目錄、不同檔）。
- `InitSimDatabase()`（2256）：開連線、WAL、建 `mail_simhash (entry_id BLOB PK, simhash INTEGER, bigram_count INTEGER)`。在 `InitDatabase` 末段呼叫（line 163）。
- 記憶體快取 `_cacheSimHash`（2255），`LoadSimHashCache()`（2284）整表載入一次。
- `SaveSimHashBatch()`（2300）：自開 `_dbSim.BeginTransaction()` upsert。
- `DeleteSimDatabase()`（2272）：關連線→刪檔→`_cacheSimHash.Clear()`→`InitSimDatabase()` 重建。供「清快取」對話框 checkbox 勾選時呼叫。
- `ZipAndRebuildDB()`（349）**只壓縮/刪 OLAcache.db，完全不碰 OLAsimhash.db**（line 2251 註解明載）。

**attach_filenames 現有 schema（line 268-274，要原樣搬到 InitSimDatabase）：**
```
entry_id BLOB PRIMARY KEY, folder_hash INTEGER NOT NULL, filenames TEXT, msg_size INTEGER
+ INDEX idx_ma_folder ON attach_filenames(folder_hash)
```

## 2.3 attach_filenames 完整接觸點盤點（搬移地圖）

> 所有「SQL 直接打到 attach_filenames」的點都必須從 `_db` 改成 `_dbSim`。以下行號以**原始檔**為準（套 Part 1 後 #4/CleanupOrphan 行號會位移，請以錨點搜尋）。

| 類別 | 函式 / 位置 | 原連線 | 搬移後動作 |
|---|---|---|---|
| **建表** | `InitDatabase` line 268-274 | _db | 從主 db 移除 → 改在 `InitSimDatabase` 建表+索引 |
| **寫入** | `SaveAttachFilenamesInner` line ~1141（`INSERT OR REPLACE INTO attach_filenames (entry_id,folder_hash,filenames)`，收 `txn As SqliteTransaction`） | _db, txn | 改 `_dbSim` + **獨立 `_dbSim` 交易**（不能再吃外部 `_db` txn）。呼叫端 `SaveCachesToDB` 對此表的交易模型要拆出來 |
| **讀取(單筆)** | `DbGetAttachFilenames` line ~1700（`SELECT filenames WHERE entry_id=@eid`） | _db | 改 `_dbSim` |
| **讀取(整表)** | `LoadAttachFilenamesInner` line ~1441（`SELECT entry_id,filenames`，重建 `_cacheAttachFilename`） | _db | 改 `_dbSim` |
| **DELETE(本次新增)** | `DbPurgeFolderMailRows`（Part 1 #4，`includeAttachFilenames` 段） | _db, txn | 改 `_dbSim` + 獨立交易（**§2.4 耦合點**） |
| **DELETE(孤兒)** | `CleanupOrphanPath` line ~841（`c3 = DELETE ... attach_filenames`，與 c1/c2/c4 同掛 `_db` txn） | _db, txn | 把 c3 拆出，改 `_dbSim` 獨立交易 |
| **統計 COUNT** | `GetDbStats`(或同名) line ~331（`SELECT COUNT(*) FROM attach_filenames`） | _db | 改 `_dbSim`；另注意 attach_filenames 的位元組現在計入 OLAsimhash.db 檔案大小，非 OLAcache.db |
| **UI 顯示** | Form1_Maintab56.vb 1077/1234/1274/1275、Module_SQLite2.vb 2079 displayOrder | — | 純顯示字串/排序，不碰 SQL；視 Tab6「各表落在哪個檔」是否要標示而定 |

`_cacheAttachFilename`（記憶體）本身不變（key 仍是 EntryID），只是其 DB 來源換成 `_dbSim`。

## 2.4 與 Part 1 的耦合點（**這是「兩件事一起做」的關鍵**）

Part 1 把 attach_filenames 的 DELETE 寫進了**兩個 `_db` 交易**裡：
1. `DbPurgeFolderMailRows`：`If includeAttachFilenames Then DELETE ... attach_filenames ... _db, txn`。
2. `CleanupOrphanPath`：c3 `DELETE ... attach_filenames ... _db, txn`。

attach_filenames 一旦搬到 `_dbSim`，這兩處的 attach_filenames DELETE **不能再掛在 `_db` 交易上**（跨檔交易在不同 connection 上是兩回事），必須抽出來用 `_dbSim` 自己的交易執行。建議重構成一支小 helper（例如 `SimDbDeleteAttachFilenamesByFolder(fPath)`），在這兩處呼叫。

> 因此**順序固定**：先套 Part 1（attach_filenames DELETE 落在 `_db`）→ 再做 Part 2（把這兩處 + 寫入/讀取/建表 re-point 到 `_dbSim`）。若先做 Part 2 再套 Part 1，Part 1 的 changeset 會把 DELETE 又寫回 `_db`，得再改一次。

## 2.5 待 Simon 拍板的設計決策（另一對話開工前必問）

1. **跨檔交易模型**：`SaveCachesToDB` 原本「多表單一 `_db` 交易」。attach_filenames 移走後，它的寫入要獨立 `_dbSim` 交易。是否接受「主 db 交易成功、sim db 交易獨立成敗」（兩檔不再原子一致）？實務上可接受（兩者都是可重建快取），但要確認。
2. **「清快取」coupling**：attach_filenames 進 OLAsimhash.db 後，`DeleteSimDatabase()`（清 SimHash 的 checkbox）會**連 attach_filenames 一起砍**。可接受嗎？若可，`DeleteSimDatabase` 需補 `_cacheAttachFilename.Clear()`（目前只 Clear `_cacheSimHash`）。若不可接受，需給 attach_filenames 獨立清除選項。
3. **ZipAndRebuildDB 行為變更**：搬移後 attach_filenames **不再被 rebuild 清掉**（這正是目的）。確認這是期望行為，並更新相關 UI 文案/說明。
4. **舊資料遷移**：既有 OLAcache.db 裡已累積的 attach_filenames 列要不要一次性搬進 OLAsimhash.db（`ATTACH DATABASE` + `INSERT INTO ... SELECT`），還是放棄舊資料、之後 lazy 重建？前者省一次昂貴重掃，建議做。
5. **主 db 是否 DROP 舊表**：搬走後主 db 的 `attach_filenames` 要 `DROP TABLE` 回收空間，或留空表不管？建議 DROP（但需在資料遷移之後）。

## 2.6 建議執行順序（給另一對話）
1. 先確認 §2.5 五個決策。
2. （若決策 4 = 遷移）寫一次性遷移：`ATTACH OLAcache.db` → `INSERT INTO sim.attach_filenames SELECT * FROM main.attach_filenames` → 驗證筆數 → （決策 5）`DROP TABLE main.attach_filenames`。
3. `InitSimDatabase` 補建 attach_filenames 表 + 索引；`InitDatabase` 移除其建表。
4. re-point §2.3 全部 SQL 接觸點到 `_dbSim`；把 §2.4 兩處 DELETE 抽成 sim 專用 helper。
5. `SaveAttachFilenamesInner` 改獨立 `_dbSim` 交易；調整 `SaveCachesToDB` 呼叫方式。
6. （決策 2）`DeleteSimDatabase` 補 `_cacheAttachFilename.Clear()`。
7. 統計/Tab6 顯示校正（COUNT 來源、檔案大小歸屬）。
8. 驗證：Tab3 附件搜尋仍命中；清 SimHash checkbox 行為符合決策 2；ZipAndRebuildDB 後 attach_filenames 仍在。

---

# Part 3 — 給另一個對話的起手指示

> 直接把以下貼給新對話即可：

「我要把 `attach_filenames` 表從 `OLAcache.db` 搬進 `OLAsimhash.db`（與 `mail_simhash` 同檔）。請先讀本 memory 的 **Part 2**（接觸點地圖 §2.3、耦合點 §2.4、決策 §2.5、順序 §2.6）。**前提：Part 1 的 RenewCacheToDB purge 修復已套用**（attach_filenames 的 DELETE 目前落在 `_db` 的 `DbPurgeFolderMailRows` 與 `CleanupOrphanPath` c3）。請先就 §2.5 的五個決策列選項給我確認，再動手；不要直接改。所有改動集中在 `Module_SQLite2.vb`（少數 UI 字串在 `Form1_Maintab56.vb`）。」

---

## 附：本次新增/變更的符號速查
- 新增：`DbPurgeFolderMailRows(fPath, Optional includeAttachFilenames)` — Module_SQLite2.vb。
- 刪除：`RenewAttachMailList(folder, fPath)` — 死碼，三路比對已被通用 purge 取代。
- 行為變更：`SaveBasicMailInfoInner` / `SaveAttachMailListInner`（寫入邊界 purge）、`SaveMonthCountsInner`（逐年先刪後插）、`RenewCacheToDB` 狀況 A（補失效集合）與 step 7（含自身 All 聚合）。
- 不變量：**同一資料夾的 basic/attach 列共用單一 snap**（全套修復的核心目標）。
