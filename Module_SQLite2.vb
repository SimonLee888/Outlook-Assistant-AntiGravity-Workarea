Imports System.Collections.Concurrent
Imports System.IO.Hashing
Imports System.Text
Imports System.Text.Json
Imports Microsoft.Data.Sqlite
Imports Microsoft.Office.Interop

' ==============================================================
' Form1_SQLite2.vb  —  SQLite 持久化快取層
' ==============================================================
' 目的: 把記憶體 ConcurrentDictionary 快取持久化到 SSD，下次啟動可快速重建
' 架構:
'   往上往下串接在Layer1/Layer2/Layer2.5和Layer3之間 (其他層呼叫完全不知道 SQLite 的存在)
'   Form1_SQLite2.vb  (本檔)
'   - InitDatabase()                            ' Form1_Load 呼叫，建 connection + CREATE TABLE IF NOT EXISTS
'   - CloseDatabase()                           ' FormClosing 呼叫
'   - LoadCachesFromDB()                        ' LoadCache 按鈕手動讀出：Bulk Load，輸出詳細 _dbg 分項
'   - SaveCachesToDB()                          ' SaveCache 按鈕手動存入：① CleanupOrphanFolderPath → ② 批次寫入四張表
'   - timerSaveCache_Tick()                     ' 每 60 秒自動呼叫 SaveCachesToDB()，僅在有 dirty 資料夾時才真正寫入 (2026-07-05 新增)
'   - CleanupOrphanFolderPath(livePaths)        ' 清除 DB 中已不存在的 folder_path (原 PurgeStaleFolders)，SaveCache 時順帶呼叫
'   - RenewCacheToDB(includeSize As Boolean)    ' RenewCache 按鈕：Phase1~6 完整更新 (2026-04-09 新增) 
'   - RenewAttMailList(folder, fPath:=fPath)    ' 三路比對更新 att_maillist (2026-04-09 新增) 
'
'   - LazyGetFolderInfo(fPath)                  ' folder_info 單行查詢
'   - DbGetMailBasic(fPath)                     ' mail_basic WHERE folder_path=? 全部行
'   - LazyGetAttFilenames(entryId)              ' att_filenames 單行查詢
'   - LazyGetYearCount(fPath)                   ' year_count WHERE folder_path=? 全部行
'   - LazyGetMonthCount(cacheKey)               ' 2026-04-09 新增，cacheKey = FolderPath_year
'   - GetDBSummary() → (fc, mb, at, yc, mc, basic, kb) ' DB 統計摘要 (六張表行數 + 檔案 KB) 
' ---------------------------------------------------------------
'
'   七張表結構 合一個 cache.db (LocalAppData)
'       2026-04-09 新增 month_count
'       2026-04-22 新增 mail_info
'       2026-06-12 新增 senders；mail_info 移除 topic/sender_email/updated_at，received_time 改 INTEGER
'       folder_info     (folder_path PK, mail_count, mail_count_all, folder_count, folder_count_all,
'                        folder_size, folder_size_all, pr_count_snap, commit_max, updated_at)  ← updated_at 僅此表保留
'                        commit_max = PR_LOCAL_COMMIT_TIME_MAX Ticks，RenewCache 第二 dirty 訊號 (2026-07-04 新增)
'       year_count      (folder_hash+year PK, count)
'       month_count     (folder_hash+year+month PK, count)
'       att_maillist    (entry_id PK, folder_hash, subject, msg_size, received_time INTEGER, sender_name,
'                        att_count, pr_count_snap)           ← 專供 Tab3 尋找附件使用
'       att_filenames   (entry_id PK, folder_hash, filenames TEXT JSON, msg_size)
'       mail_info       (entry_id PK, folder_hash, subject, msg_size, received_time INTEGER, sender_name,
'                        sender_id, msgid_hash, pr_count_snap)  ← 專供 Tab4/Tab5 系列與重複郵件使用
'       senders         (sender_id PK AUTOINCREMENT, sender_email UNIQUE) ← email 正規化，2026-06-12 新增
'                           
' 設計決策 (2026-04-06):
'   1. 跨表 Transaction 保證原子性，一個 Connection 管理最簡單
'   2. 手動控制 (SaveCache / LoadCache 按鈕)，Debug 階段方便目視確認正確性
'      正式版再切換成 Layer2.5 lazy SELECT + 增量寫入
'   3. pr_count_snap 存 _cacheMailCount[path] 的值 (即 PR_CONTENT_COUNT 的讀取結果)
'      Load 後可快速判斷快取是否仍有效，完全不需要呼叫任何 COM
'   4. MailItemInfo 欄位以文字儲存；List(Of String) 附件名稱序列化為 JSON array
'   5. _cacheFolderTree / _cacheSubTreeList 含 COM 物件，永遠不寫入 SQLite
'   6. LoadFolderInfoBatch 使用 TryAdd：若記憶體已有值 (Layer2.5 已讀過)，保留記憶體版本
'      若想強制以 DB 為準 (完整重置)，改用直接賦值 _cacheMailCount(path) = ...
'   7. (2026-04-22) 拆分 att_maillist 與 mail_info：保持 Tab3 與 Tab4/5 邏輯與資料邊界獨立。
' ==============================================================

Partial Class Form1

#Region "■ 私有成員"
    ' by Gemini, 2026/04/10: 將資料庫路徑移至 LocalAppData，避免與 Dropbox 同步衝突導致檔案鎖定與 Explorer 卡頓
    'Private ReadOnly _dbCachePath As String = IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OutlookAssistant\Cache", "OLAcache.db")
    Private _dbCache As SqliteConnection = Nothing
    Private _dbCachePath As String           ' 移除 ReadOnly 與靜態初始化
    ' 2026/06/17 by Simon/Claude Opus 4.8: SimHash 專用獨立 db。body 全文 COM 讀取極貴、simhash 又「內容不變即永久有效」，
    '   故與「會被清快取/重置 SSD 清掉」的 OLAcache.db 物理隔離。ZipAndRebuildDB 只刪 OLAcache.db，完全不碰本檔 → 結構上保證存活。
    Private _dbMail As SqliteConnection
    Private _dbMailPath As String

    Private Shared _dictHashToPath As New ConcurrentDictionary(Of Long, String) ' 路徑雜湊壓縮機制 (2026/06/11 新增)
    Private Shared _dictEmailToSenderId As New Dictionary(Of String, Integer)   ' 2026/06/12 by Simon/Claude Opus 4.8: sender 正規化寫入側：lowercase email → sender_id
    Private Shared _dictSenderIdToEmail As New Dictionary(Of Integer, String)   ' 2026/06/12 by Simon/Claude Opus 4.8: sender 正規化讀取側：sender_id → lowercase email
    Private _simHashLoaded As Boolean = False                                   ' 2026/06/18 by Simon/Claude Opus 4.8: Fuzzy 模糊比對專用區塊 (SimHash + bigram Jaccard)
    Private _cacheSimHash As New Concurrent.ConcurrentDictionary(Of String, (SimHash As Long, BigramCount As Integer))

    ' 2026/07/03 by Simon/Claude Fable 5: dirty 追蹤 — 記錄「自上次 SaveCache 後，此資料夾的 mail_info/att_maillist/
    '   year_count/month_count 曾被 COM(或 RDO)重新計算過」的資料夾路徑。SaveCache 的四張逐封/逐年表批次寫入只處理這個
    '   集合裡的資料夾，不再每次把記憶體快取全部照單全收(實測 307k 列 mail_info 全量重寫要 2.4~2.7 秒，即使完全沒異動)。
    '   只用 Byte 當 Value 純粹借 ConcurrentDictionary 當執行緒安全的 Set 用，Value 本身無意義。
    Private Shared _dirtyMailFolders As New ConcurrentDictionary(Of String, Byte)
    Private _isAutoSavingCache As Boolean = False  ' 2026/07/05 by Simon/Claude: timerSaveCache_Tick 重入防護，避免上一輪存檔還沒跑完就疊加下一輪
    ' 2026/07/11 by Simon/Claude Fable 5: Lv6 統計查詢互斥旗標 — DbShowDbFileStat/DbShowTableStat/DbShowBigramSetStat/RefreshLv6DbStats
    '   都會把查詢丟進 Task.Run，同一條 SqliteConnection 不允許並發使用；旗標僅在 UI 執行緒讀寫(所有入口都是 UI 事件)，
    '   忙碌中直接跳過並提示，比 SemaphoreSlim 排隊簡單且不會累積點擊積壓。
    Private _lv6StatBusy As Boolean = False
    ' 2026/07/04 by Simon/Claude Fable 5: PR_LOCAL_COMMIT_TIME_MAX 快照 (UTC Ticks)，RenewCache 第二 dirty 訊號。
    '   純 count 快照抓不到「數量不變但內容已置換」(copy→修改→放回→刪原始 = 淨零變動)，commit_max 任何增/刪/改都會推高。
    '   僅作寫入暫存 (RenewCache 掃描時填入 → SaveFolderInfoBatch 落 DB)，啟動時不需載回記憶體，比對一律以 DB row 為準。
    Private _cacheFolderCommitMax As New Concurrent.ConcurrentDictionary(Of String, Long)

    ' DB Row 結構 (供 Form1_Outlook.vb 的 Layer2.5 函數使用)
    Private Class FolderInfoDbRow
        ' folder_info 一行的讀出結果；-1 代表該欄位在 DB 中為 NULL 或尚未寫入
        Public mc As Long = -1          ' mail_count
        Public mca As Long = -1         ' mail_count_all
        Public fc As Long = -1          ' folder_count
        Public fca As Long = -1         ' folder_count_all
        Public fs As Long = -1          ' folder_size
        Public fsa As Long = -1         ' folder_size_all
        Public snap As Long = -1        ' pr_count_snap (= PR_CONTENT_COUNT at save time)
        Public cmx As Long = -1         ' commit_max (= PR_LOCAL_COMMIT_TIME_MAX UTC Ticks at save time; -1 = NULL/未知) ' 2026/07/04 by Simon/Claude Fable 5
        Public path As String = ""      ' folder_path        ' by Gemini 3.0 flash, 2026/04/16: 新增路徑標識，供 GetSubtree Tuple 重建使用

        ' by Gemini, 2026/04/10: 新增身分標識與排序標籤，供 TreeView/BFS 持久化優化使用
        Public eid As String = ""       ' entry_id  ' 2026/06/10 by Gemini/Simon: 優化SQLite儲存空間把Entry_id改成BLOB二進位儲存
        Public sid As String = ""       ' store_id
        Public isMail As Integer = -1   ' is_mail (0/1)
        Public hasCh As Integer = -1    ' has_chinese (0/1)
    End Class
    Private Class AttMailListDbResult
        Public Snap As Long = -1                        ' att_maillist WHERE folder_path=? 的讀出結果
        Public Mails As New List(Of MailItemInfo)(1024) ' 預分配容量為 1024，降低自 SQLite 載入大量郵件快取時的 Resize 開銷 (by Gemini 3 Flash, 2026/05/04)
    End Class
#End Region

#Region "■ 基礎結構與資料庫生命週期管理 (Lifecycle & Schema)"
    Private Sub InitDatabase()
        ' ---------------------------------------------------------------
        ' InitDatabase — 建立或開啟 cache.db，確保三張表與索引存在
        ' 在 UI 執行緒呼叫 (Form1_Load)，SQLite DDL 量小，不需要 Async
        ' ---------------------------------------------------------------
        _dbg("開始")

        ' by Claude Sonnet 4.6, 2026/05/03: 修正兩個錯誤：
        '   1. CurrentProfileName? 結尾的 ? 在 VB.NET 是非法語法（?.是 Null-Conditional，結尾? 不存在）
        '   2. SanitizeProfileName 只接受一個參數，不能傳入第二個 "Default" 引數
        '   → 改為先用 If() 提供 Null fallback，再傳入 SanitizeProfileName
        Dim profileName As String = SanitizeProfileName(If(_olNS?.CurrentProfileName, "Default"))
        Dim cacheDir = IO.Path.Combine(Environment.GetFolderPath(
        Environment.SpecialFolder.LocalApplicationData), "OutlookAssistant\Cache", profileName)
        _dbCachePath = IO.Path.Combine(cacheDir, "OLAcache.db")

        Try
            ' by Gemini, 2026/04/10: 確保資料庫目錄存在
            Dim dbDir = IO.Path.GetDirectoryName(_dbCachePath)
            If Not IO.Directory.Exists(dbDir) Then IO.Directory.CreateDirectory(dbDir)
            ' 2026/07/03 by Simon/Claude: 移除 Cache=Shared — WAL 模式下 shared cache 沒有好處，反而引入 table-level lock，
            '   且會讓「專用寫入連線」(SaveCachesToDB) 與本連線共用 page cache 互相牽制，抵銷 WAL 一寫多讀的並行能力
            _dbCache = New SqliteConnection($"Data Source={_dbCachePath};Mode=ReadWriteCreate")
            _dbCache.Open()

            Using cmd As New SqliteCommand("PRAGMA journal_mode=WAL; PRAGMA synchronous=NORMAL;", _dbCache)
                cmd.ExecuteNonQuery()
            End Using
            _dbg("", $"已開啟: {_dbCachePath}")

            Using cmd As New SqliteCommand(BuildSQLiteTableString(), _dbCache)
                cmd.ExecuteNonQuery()
            End Using
            ' 2026/06/10 by Gemini/Simon: 優化SQLite儲存空間把Entry_id改成BLOB二進位儲存
            ' 2026/06/11 by Gemini/Simon: 把 message_id 轉成 xxHash64，並同時改成 BLOB 儲存節省空間
            ' 2026/05/06 by Claude: mail_info 新增 Tab5 去重欄位
            ' 2026/07/04 by Simon/Claude Fable 5: folder_info 新增 commit_max (PR_LOCAL_COMMIT_TIME_MAX Ticks)，RenewCache 淨零置換偵測用
            ' 2026/07/07 by Simon/Claude: 拆除上述歷次的 ALTER TABLE 線上遷移段 — Simon 決策：本專案是單機工具，
            '   schema 有變更時一律直接重建 DB（測試效能本來就要全部重建），不維護線上遷移的冗餘架構。
            '   所有現役欄位已全數內建於 BuildSQLiteTableString 的 CREATE TABLE 本體
            '   （mail_info 的 sender_email 已於 2026/06/12 改為 sender_id 正規化，該句 ALTER 本屬殘留）。

            ' by Claude Sonnet 4.6, 2026/05/06: Root Cause A 一次性資料清理 migration
            ' by Claude Sonnet 4.6, 2026/06/12: 整段刪除 (原本這段 migration 的唯一目的是清理舊 bug 遺留的污染資料)
            ' 2026/06/12 by Claude: mail_info 欄位重排序 migration
            '   目標順序: entry_id, msg_size, subject, topic, msgid_hash, folder_hash,
            '             sender_name, sender_email, pr_count_snap, received_time, updated_at
            '   用 RENAME → CREATE → INSERT SELECT → DROP 流程，因 SQLite 不支援 ALTER COLUMN ORDER

            _dbg("", "資料表確認完成")

            LoadSendersBatch()  ' 2026/06/12 by Simon/Claude Opus 4.8: 載入 senders 表，供 LoadMailInfoCore lazy load 時能解析 sender_id
            InitDbMail()   ' 2026/06/17 by Simon/Claude Opus 4.8: 開啟獨立 SimHash db(OLAsimHash.db)並一次載入記憶體快取(Tab5 Fuzzy 暖快取)
            LoadDbMail()  ' 2026/06/17 by Simon/Claude Opus 4.8: 開啟獨立 SimHash db(OLAsimHash.db)並一次載入記憶體快取(Tab5 Fuzzy 暖快取)

            timerSaveCache.Interval = 1 * 60 * 1000 ' 每60sec自動保存一次快取資料到磁碟
            timerSaveCache.Start()                  ' 啟動定時快取保存

        Catch ex As System.Exception
            _dbg("       ├ 錯誤", ex.Message) ' by Gemini, 2026/04/11: Level 3
            _dbCache = Nothing                     ' 出錯就設 Nothing，後續所有 SQLite 操作因此自動跳過
        Finally : _dbg("結束")                ' by Gemini, 2026/04/11: 修正對應開始層級 Level 0
        End Try

    End Sub
    Private Sub CloseDatabase()
        ' ---------------------------------------------------------------
        ' CloseDatabase — FormClosing 時呼叫，安全關閉 SQLite 連線
        ' ---------------------------------------------------------------
        _dbg("開始")

        If _dbCache Is Nothing Then Return

        Try
            _dbCache.Close() : _dbCache.Dispose() : _dbCache = Nothing

            ' 2026/06/17 by Simon/Claude Opus 4.8: 一併關閉獨立 SimHash db
            If _dbMail IsNot Nothing Then _dbMail.Close() : _dbMail.Dispose() : _dbMail = Nothing

            ' 清除連線池，強制釋放底層的檔案 Handle，避免稍後備份或刪除時發生 IOException
            SqliteConnection.ClearAllPools()

            ' 建議加上 GC 強制回收，因為有些未 Dispose 的 Command 可能仍咬住檔案 (by Gemini, 2026/04/10)
            GC.Collect()
            GC.WaitForPendingFinalizers()

            _dbg("", "SQLite 連線與鎖已安全解除")

        Catch ex As System.Exception
            _dbg("       ├ 錯誤", ex.Message) ' by Gemini, 2026/04/11: Level 3
        Finally : _dbg("結束") ' by Gemini, 2026/04/11: 修正對應開始層級 Level 0
        End Try

    End Sub
    Private Function BuildSQLiteTableString() As String

        ' 2026/06/10 by Gemini/Simon: 優化SQLite儲存空間把Entry_id改成BLOB二進位儲存
        ' 2026/06/11 by Gemini/Simon: 把 message_id 轉成 xxHash64，並同時改成 BLOB 儲存節省空間
        ' 2026/06/12 by Gemini/Simon: 把 folder_path TEXT 改成 folder_hash INTEGER 儲存，並新增一個記憶體字典 _dictHashToPath 來反查雜湊對應的路徑

        ' by Gemini 3.1 Pro (Low), 2026/06/09: 將原本巨大的多行 SQL 字串拆分，使用 StringBuilder 依資料表分類，提升結構清晰度與後續維護性。
        Dim sb As New System.Text.StringBuilder()

        ' 1. folder_info: 資料夾狀態統計 (核心快取層)
        sb.AppendLine("
                        CREATE TABLE IF NOT EXISTS folder_info ( 
                            folder_path         TEXT    PRIMARY KEY,
                            mail_count          INTEGER,
                            mail_count_all      INTEGER,
                            folder_count        INTEGER,
                            folder_count_all    INTEGER,
                            folder_size         INTEGER,
                            folder_size_all     INTEGER,
                            pr_count_snap       INTEGER,
                            commit_max          INTEGER,
                            entry_id            BLOB,
                            store_id            TEXT,
                            is_mail             INTEGER,
                            has_chinese         INTEGER,
                            updated_at          TEXT
                        );")

        ' 2. att_maillist: 附件郵件清單 (專供 Tab3 尋找附件使用)
        ' 2026/06/12 by Simon/Claude Opus 4.8: received_time TEXT→INTEGER (Unix秒)；移除 updated_at (只寫不讀，無用)
        sb.AppendLine("
                        CREATE TABLE IF NOT EXISTS att_maillist (
                            entry_id        BLOB    PRIMARY KEY,
                            folder_hash     INTEGER NOT NULL,
                            subject         TEXT,
                            msg_size        INTEGER,
                            received_time   INTEGER,
                            sender_name     TEXT,
                            att_count       INTEGER,
                            pr_count_snap   INTEGER
                        );
                        CREATE INDEX IF NOT EXISTS idx_mb_folder ON att_maillist(folder_hash);")

        ' 3. mail_info: 基礎郵件清單 (專供 Tab4/Tab5 與重複郵件比對使用)
        ' 2026/06/12 by Simon/Claude Opus 4.8: 移除 topic (改由 GetCleanSubject(subject) 動態計算)
        '   移除 sender_email (改以 sender_id 外鍵指向 senders 表，節省重複儲存)
        '   received_time TEXT→INTEGER (Unix秒)；移除 updated_at (只寫不讀)
        sb.AppendLine("
                        CREATE TABLE IF NOT EXISTS mail_info (
                            entry_id        BLOB    PRIMARY KEY,
                            folder_hash     INTEGER NOT NULL,
                            subject         TEXT,
                            msg_size        INTEGER,
                            received_time   INTEGER,
                            sender_name     TEXT,
                            sender_id       INTEGER,
                            msgid_hash      BLOB,
                            pr_count_snap   INTEGER
                        );
                        CREATE INDEX IF NOT EXISTS idx_mail_info_folder ON mail_info(folder_hash);")

        '' 4. att_filenames: 附件名稱清單
        '' 2026/06/12 by Simon/Claude Opus 4.8: 移除 updated_at (只寫不讀)
        ' 4. att_filenames: 已於 2026/06/21 by Simon/Claude 搬至 OLAcacheMail.db(_dbMail)，改在 InitDbMail 建表。
        '    理由：逐封開信枚舉附件、重建極貴，需隨該檔在 ZipAndRebuildDB 後存活（與 mail_simhash 同策略）。
        'sb.AppendLine("
        '                CREATE TABLE IF NOT EXISTS att_filenames (
        '                    entry_id        BLOB    PRIMARY KEY,
        '                    folder_hash     INTEGER NOT NULL,
        '                    filenames       TEXT,
        '                    msg_size        INTEGER
        '                );
        '                CREATE INDEX IF NOT EXISTS idx_ma_folder ON att_filenames(folder_hash);")

        ' 5. year_count: 年份統計
        ' 2026/06/12 by Simon/Claude Opus 4.8: 移除 updated_at (只寫不讀)
        sb.AppendLine("
                        CREATE TABLE IF NOT EXISTS year_count (
                            folder_hash     INTEGER NOT NULL,
                            year            INTEGER NOT NULL,
                            count           INTEGER NOT NULL,
                            PRIMARY KEY (folder_hash, year)
                        );
                        CREATE INDEX IF NOT EXISTS idx_yc_folder ON year_count(folder_hash);")

        ' 6. month_count: 月份統計
        ' 2026/06/12 by Simon/Claude Opus 4.8: 移除 updated_at (只寫不讀)
        sb.AppendLine("
                        CREATE TABLE IF NOT EXISTS month_count (
                            folder_hash INTEGER NOT NULL,
                            year        INTEGER NOT NULL,
                            month       INTEGER NOT NULL,
                            count       INTEGER NOT NULL,
                            PRIMARY KEY (folder_hash, year, month)
                        );")

        ' 7. senders: 寄件者 email 正規化表 (2026/06/12 by Simon/Claude Opus 4.8 新增)
        '   只存不重複的 sender_email；mail_info 透過 sender_id 外鍵參照
        sb.AppendLine("
                        CREATE TABLE IF NOT EXISTS senders (
                            sender_id       INTEGER PRIMARY KEY AUTOINCREMENT,
                            sender_email    TEXT    UNIQUE NOT NULL
                        );")

        ' 注意：month_count 的舊版 schema 遷移 (cache_key → 三欄 PK) 在 InitDatabase() 中一次性處理，
        ' 不在此處 DROP TABLE，避免每次啟動都清空已存資料。

        Return sb.ToString()

    End Function
    Private Function GetDBSummary() As (fc As Integer, mb As Integer, at As Integer, yc As Integer, mc As Integer, basic As Integer, senders As Integer, kb As Long, lastTs As String, kbMail As Long, sh As Integer, bs As Integer, bsMB As Single)
        ' ---------------------------------------------------------------
        ' GetDBSummary — 取得 DB 統計摘要，供按鈕顯示
        ' 回傳 (folder_info, att_maillist, att_filenames, year_count, month_count, mail_info, senders, KB, lastTs)
        ' 2026/04/09 新增 mc = month_count 行數
        ' 2026/04/10 新增 lastTs = 最後 updated_at 時間
        ' 2026/04/22 by Gemini 3 Flash: 新增 basic = mail_info 行數
        ' 2026/06/14 by Simon/Claude Opus 4.8: 新增 senders = senders 行數 (供 Tab6 Lv6 顯示)
        ' 2026/06/21 by Simon/Claude Opus 4.8: att_filenames 已搬至 _dbMail(OLAcacheMail.db)
        ' ---------------------------------------------------------------
        If _dbCache Is Nothing Then Return (0, 0, 0, 0, 0, 0, 0, 0L, "N/A", 0L, 0, 0, 0F)

        Try
            Dim fc, mb, at, yc, mcount, basicCount, sendersCount As Integer : Dim lastTs As String = "N/A"
            Using cmd As New SqliteCommand("SELECT COUNT(*) FROM folder_info", _dbCache) : fc = Convert.ToInt32(cmd.ExecuteScalar()) : End Using
            Using cmd As New SqliteCommand("SELECT COUNT(*) FROM mail_info", _dbCache) : basicCount = Convert.ToInt32(cmd.ExecuteScalar()) : End Using
            Using cmd As New SqliteCommand("SELECT COUNT(*) FROM senders", _dbCache) : sendersCount = Convert.ToInt32(cmd.ExecuteScalar()) : End Using ' 2026/06/14 by Simon/Claude Opus 4.8
            Using cmd As New SqliteCommand("SELECT COUNT(*) FROM year_count", _dbCache) : yc = Convert.ToInt32(cmd.ExecuteScalar()) : End Using
            Using cmd As New SqliteCommand("SELECT COUNT(*) FROM month_count", _dbCache) : mcount = Convert.ToInt32(cmd.ExecuteScalar()) : End Using
            Using cmd As New SqliteCommand("SELECT COUNT(*) FROM att_maillist", _dbCache) : mb = Convert.ToInt32(cmd.ExecuteScalar()) : End Using

            ' 2026/06/21 by Simon/Claude Opus 4.8: att_filenames 已搬至 _dbMail(OLAcacheMail.db)，COUNT 改打 _dbMail；_dbMail 為 Nothing 時 at=0
            If _dbMail IsNot Nothing Then Using cmd As New SqliteCommand("SELECT COUNT(*) FROM att_filenames", _dbMail) : at = Convert.ToInt32(cmd.ExecuteScalar()) : End Using

            ' 2026/06/21 by Simon/Claude: mail_simhash 筆數同樣讀 _dbMail (供 Tab6 Lv6 雙檔顯示)
            Dim sh As Integer = 0
            If _dbMail IsNot Nothing Then Using cmd As New SqliteCommand("SELECT COUNT(*) FROM mail_simhash", _dbMail) : sh = Convert.ToInt32(cmd.ExecuteScalar()) : End Using

            ' 2026/07/07 by Simon/Claude: bigram_set BLOB 回填統計 — COUNT(欄名) 只算非 NULL(即進過 Tab5 S5 的候選), 另算總 bytes
            Dim bs As Integer = 0 : Dim bsMB As Single = 0F
            If _dbMail IsNot Nothing Then
                Using cmd As New SqliteCommand("SELECT COUNT(bigram_set), IFNULL(SUM(LENGTH(bigram_set)),0) FROM mail_simhash", _dbMail)
                    Using r = cmd.ExecuteReader()
                        If r.Read() Then bs = r.GetInt32(0) : bsMB = CSng(r.GetInt64(1) / 1048576.0)
                    End Using
                End Using
            End If

            ' 抓取最後一次成功的儲存時間字串 (取最大的 updated_at)
            Using cmd As New SqliteCommand("SELECT MAX(updated_at) FROM folder_info", _dbCache)
                Dim val = cmd.ExecuteScalar()
                If val IsNot Nothing AndAlso Not IsDBNull(val) Then lastTs = val.ToString()
            End Using

            ' 2026/06/21 by Simon/Claude: OLAcacheMail.db 檔案大小(KB)；_dbMailPath 空或檔不存在時為 0
            Dim fi As New IO.FileInfo(_dbCachePath)
            Dim kbMail As Long = 0L : If Not String.IsNullOrEmpty(_dbMailPath) AndAlso IO.File.Exists(_dbMailPath) Then kbMail = New IO.FileInfo(_dbMailPath).Length \ 1024L
            Return (fc, mb, at, yc, mcount, basicCount, sendersCount, If(fi.Exists, fi.Length \ 1024L, 0L), lastTs, kbMail, sh, bs, bsMB)

        Catch ex As System.Exception
            _dbg("       ├ 錯誤", ex.Message) ' by Gemini, 2026/04/11: Level 3
            Return (0, 0, 0, 0, 0, 0, 0, 0L, "Err", 0L, 0, 0, 0F)
        Finally : If _iLikeNoisy Then _dbg(" ├ 結束")
        End Try

    End Function
    Private Async Function ZipAndRebuildDB() As Task
        ' ---------------------------------------------------------------
        ' ZipAndRebuildDB — [透明化控制] 徹底清空並重建 SSD 快取檔
        ' 流程: 1. 關閉連線 → 2. 實體刪除 db 檔 → 3. 呼叫 InitDatabase 重建 Schema
        ' ---------------------------------------------------------------
        _dbg(" ├ 開始") ' by Gemini, 2026/04/10: 調整縮排層級為 Level 1
        Try
            ' 1. 關閉連線並釋放檔案
            CloseDatabase()
            Await Task.Delay(200) ' 給系統一點時間釋放 Handle

            ' 2. 備份舊資料庫並壓縮 (by Gemini, 2026/04/10)
            ' ---------------------------------------------------------------
            ' 為了節省空間，改為直接壓縮成 .zip 檔
            If IO.File.Exists(_dbCachePath) Then
                Dim timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss")
                Dim zipName = $"OLAcacheZipBackup_{timestamp}.zip"
                Dim zipPath = IO.Path.Combine(IO.Path.GetDirectoryName(_dbCachePath), zipName)

                _dbg("", $"正要壓縮備份至: {zipName}")

                ' 使用全名 (Fully Qualified Name) 避免 Imports 失敗, 並改用 Stream.CopyTo 避開擴展方法找不到的錯誤
                Using zipFileStream As New System.IO.FileStream(zipPath, System.IO.FileMode.Create)
                    Using archive As New System.IO.Compression.ZipArchive(zipFileStream, System.IO.Compression.ZipArchiveMode.Create)
                        ' 加入 System.IO.Compression.CompressionLevel 參數 (2026/5/8 by simon)使用最佳壓縮等級，減少備份檔案大小
                        AddFileToZipArchive(archive, _dbCachePath, "OLAcache.db")

                        ' 2026/07/02 by Simon/Claude: OLAcacheMail.db 一併備份進同一份 zip (只備份不刪除，rebuild 後仍依原設計存活)
                        Dim mailPath = If(Not String.IsNullOrEmpty(_dbMailPath), _dbMailPath, IO.Path.Combine(IO.Path.GetDirectoryName(_dbCachePath), "OLAcacheMail.db"))
                        If IO.File.Exists(mailPath) Then AddFileToZipArchive(archive, mailPath, "OLAcacheMail.db")
                    End Using
                End Using

                IO.File.Delete(_dbCachePath) ' 壓縮完後刪除原始 db 檔 (OLAcacheMail.db 不刪，維持 rebuild 後存活的原設計)
            End If

            ' 3. 重新建立資料庫與表格
            InitDatabase()
            _dbg(" ├ 結束", "SSD 快取已重設，舊檔案已 Zip 備份") ' by Gemini, 2026/04/11: 修正對應開始層級 Level 1
        Catch ex As System.Exception
            _dbg("       ├ 錯誤", $"無法重置 SSD 資料庫: {ex.Message}")
            Throw
        End Try
    End Function
    Private Sub AddFileToZipArchive(archive As System.IO.Compression.ZipArchive, srcPath As String, entryName As String)
        ' 供 ZipAndRebuildDB 共用：把單一檔案以 SmallestSize 壓縮等級寫入 zip 的一個 entry
        Dim entry = archive.CreateEntry(entryName, System.IO.Compression.CompressionLevel.SmallestSize)
        Using entryStream = entry.Open()
            ' 加上 FileShare.ReadWrite 容許其他可能卡住的唯讀鎖，防止 IOException (by Gemini, 2026/04/10)
            Using fs As New System.IO.FileStream(srcPath, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite)
                fs.CopyTo(entryStream)
            End Using
        End Using
    End Sub
    Private Function SanitizeProfileName(name As String) As String
        ' Profile 名稱安全過濾
        ' CurrentProfileName 可能包含空格、單引號（Simon'st Mail）、甚至斜線等非法路徑字元
        Dim invalid As Char() = IO.Path.GetInvalidFileNameChars()
        Return New String(name.Select(Function(c) If(Array.IndexOf(invalid, c) >= 0, "_"c, c)).ToArray())
    End Function

    Private Sub InitDbMail()
        ' 開啟/建立 OLAcacheMail.db (與 OLAcache.db 同目錄、不同檔)。在 InitDatabase 末段呼叫。
        ' 2026/06/21 by Simon/Claude Opus 4.8: 原 OLAsimhash.db 改名 OLAcacheMail.db；本檔現含 mail_simhash + att_filenames 兩張「逐封讀取極貴」的快取表
        ' 2026/06/21 by Simon/Claude Opus 4.8: 本檔已改名 OLAcacheMail.db，並納入 att_filenames(逐封開信重建極貴)，同享「rebuild 後存活」性質。
        '   注意：rebuild 不清本檔，但 RenewCache 狀況 A(內容有變) 與孤兒清理仍會精確 purge 本檔對應 folder 的 att_filenames 列，避免死列殘留。
        Try
            _dbMailPath = IO.Path.Combine(IO.Path.GetDirectoryName(_dbCachePath), "OLAcacheMail.db")
            _dbMail = New SqliteConnection($"Data Source={_dbMailPath};Mode=ReadWriteCreate")   ' 2026/07/03 by Simon/Claude: 移除 Cache=Shared，理由同 _dbCache
            _dbMail.Open()
            Using cmd As New SqliteCommand("PRAGMA journal_mode=WAL; PRAGMA synchronous=NORMAL;", _dbMail) : cmd.ExecuteNonQuery() : End Using
            EnsureDbMailSchema()
            _dbg("", $"已開啟 OLAcacheMail db: {_dbMailPath}")

        Catch ex As System.Exception
            _dbg("       ├ 錯誤", ex.Message) : _dbMail = Nothing   ' 出錯設 Nothing，後續 sim 讀寫自動跳過 (同主 db 容錯策略)
        End Try
    End Sub
    Private Sub EnsureDbMailSchema()
        ' 2026/07/11 by Simon/Claude: 從 InitDbMail 抽出，供「Tab6 單表刪除」DROP 後重建 schema 共用，避免 CREATE TABLE 語句重複維護兩份
        ' 2026/07/07 by Simon/Claude: 新增 bigram_set BLOB(核心版候選集合快取) — 只有進過 Tab5 S5 精算的候選信才會回填此欄
        '   (每 bigram 4 bytes packed Int32, 實測全庫平均每封僅 ~2.4KB)。舊 DB 無此欄位者直接整檔重建, 不做 ALTER 遷移。
        Using cmd As New SqliteCommand("CREATE TABLE IF NOT EXISTS mail_simhash (entry_id BLOB PRIMARY KEY, simhash INTEGER NOT NULL, bigram_count INTEGER NOT NULL, bigram_set BLOB);", _dbMail)
            cmd.ExecuteNonQuery()
        End Using

        ' 2026/06/21 by Simon/Claude Opus 4.8: att_filenames 由 OLAcache.db 搬入本檔(schema 原樣保留：entry_id BLOB PK / folder_hash / filenames / msg_size)
        Using cmd As New SqliteCommand("CREATE TABLE IF NOT EXISTS att_filenames (entry_id BLOB PRIMARY KEY, folder_hash INTEGER NOT NULL, filenames TEXT, msg_size INTEGER);" &
                                       "CREATE INDEX IF NOT EXISTS idx_ma_folder ON att_filenames(folder_hash);", _dbMail)
            cmd.ExecuteNonQuery()
        End Using
    End Sub
    Private Sub LoadDbMail()
        ' 2026/06/17 by Simon/Claude Opus 4.8: mail_simhash 整表 lazy load 進記憶體 (僅一次)。每列 16B+eid，量小可全載。
        If _simHashLoaded OrElse _dbMail Is Nothing Then Return
        Try
            Using cmd As New SqliteCommand("SELECT entry_id, simhash, bigram_count FROM mail_simhash", _dbMail)
                Using r = cmd.ExecuteReader()
                    While r.Read()
                        _cacheSimHash(ByteArrayToHexString(r.GetFieldValue(Of Byte())(0))) = (r.GetInt64(1), r.GetInt32(2))
                    End While
                End Using
            End Using
            _simHashLoaded = True : _dbg("", $"SimHash 快取載入 {_cacheSimHash.Count} 筆")
        Catch ex As System.Exception
            _dbg("       ├ 錯誤", ex.Message)
        End Try
    End Sub
    Private Sub SaveDbMail(rows As IEnumerable(Of (EntryID As String, SimHash As Long, BigramCount As Integer)))
        ' 批次 upsert 到獨立 db (交易包覆)；entry_id 沿用 HexStringToByteArray 編碼。呼叫端同時更新 _cacheSimHash。
        If _dbMail Is Nothing Then Return
        Try
            Using txn = _dbMail.BeginTransaction()
                Using cmd As New SqliteCommand("INSERT OR REPLACE INTO mail_simhash (entry_id,simhash,bigram_count) VALUES (@eid,@sh,@bc)", _dbMail, txn)
                    ' 2026/07/03 by Simon/Claude Fable 5: 參數物件存區域變數，免去迴圈內名稱線性查找
                    Dim pEid = cmd.Parameters.Add("@eid", SqliteType.Blob) : Dim pSh = cmd.Parameters.Add("@sh", SqliteType.Integer) : Dim pBc = cmd.Parameters.Add("@bc", SqliteType.Integer)
                    For Each row In rows
                        pEid.Value = HexStringToByteArray(row.EntryID)
                        pSh.Value = row.SimHash : pBc.Value = row.BigramCount
                        cmd.ExecuteNonQuery()
                    Next
                End Using
                txn.Commit()
            End Using
        Catch ex As System.Exception
            _dbg("       ├ 錯誤", ex.Message)
        End Try
    End Sub
    Private Sub DeleteDbMail(Optional backupZip As Boolean = True)
        ' 供「清快取」對話框「郵件快取 / 兩者全清」呼叫：關閉連線 → 備份 → 刪檔 → 清記憶體 → 重建空表
        ' 2026/06/21 by Simon/Claude Opus 4.8: 本檔(OLAcacheMail.db)現含 att_filenames，整檔刪除會一併清掉 → 須同步清 _cacheAttFilename
        ' 2026/07/11 by Simon/Claude Fable 5: 加 zip 備份 (本檔重建成本高：att_filenames/SimHash 都要逐封讀取)。
        '   backupZip:=False 供「兩者全清」路徑使用 —— ZipAndRebuildDB 的主 zip 已打包本檔舊檔，不必重複備份第二份。
        Try
            If _dbMail IsNot Nothing Then _dbMail.Close() : _dbMail.Dispose() : _dbMail = Nothing
            SqliteConnection.ClearAllPools()

            If Not String.IsNullOrEmpty(_dbMailPath) AndAlso IO.File.Exists(_dbMailPath) Then
                If backupZip Then
                    Dim zipPath = IO.Path.Combine(IO.Path.GetDirectoryName(_dbMailPath), $"OLAcacheMailZipBackup_{DateTime.Now:yyyyMMdd_HHmmss}.zip")
                    Using zipFileStream As New System.IO.FileStream(zipPath, System.IO.FileMode.Create)
                        Using archive As New System.IO.Compression.ZipArchive(zipFileStream, System.IO.Compression.ZipArchiveMode.Create)
                            AddFileToZipArchive(archive, _dbMailPath, "OLAcacheMail.db")
                        End Using
                    End Using
                    _dbg("", $"OLAcacheMail.db 已壓縮備份至: {IO.Path.GetFileName(zipPath)}")
                End If
                IO.File.Delete(_dbMailPath)
            End If
            _cacheSimHash.Clear() : _cacheAttFilename.Clear() : _simHashLoaded = False

            InitDbMail()   ' 重建空表，後續仍可重新累積

        Catch ex As System.Exception
            _dbg("       ├ 錯誤", ex.Message)
        End Try
    End Sub
#End Region

#Region "■ 快取主控流程 (High-Level Cache Controllers)"
    Private Async Function SaveCachesToDB(Optional quiet As Boolean = False) As Task(Of String)
        ' ---------------------------------------------------------------
        ' SaveCachesToDB — 把記憶體快取全部存入 SQLite
        ' 對應 Setting 頁 SaveCache 按鈕
        ' 流程: ① CleanupOrphanFolderPath (先清孤兒) → ② 批次寫入三張表 → ③ 統計顯示
        ' 2026/07/06 by Simon/Claude: 加 quiet 參數(timerSaveCache 自動存檔用)。背景自動存檔不得碰
        '   PgrsBar1/PgrsBar2/Cursor — 它會蓋掉前景長作業(如 Tab5 掃描)的進度顯示,更會把 Form1_KeyDown
        '   的 ESC 觸發條件(WaitCursor 或 PgrsBar1「正在」開頭)雙雙打掉,造成 ESC 中斷間歇性失效。
        ' ---------------------------------------------------------------
        _dbg(" ├ 開始") ' by Gemini, 2026/04/10: 調整縮排層級為 Level 1
        If _dbCache Is Nothing Then _dbg("", "DB 未初始化") : Return ""

        Dim sw As Diagnostics.Stopwatch = Diagnostics.Stopwatch.StartNew()  ' by Gemini 3.5 Flash, 2026/06/07
        Dim savedFolders, savedAttMailList, savedAttFilenames, savedBasic As Integer
        Dim resultLine As String = ""   ' 2026/07/07 by Simon/Claude: 回傳既有的 statLine3(寫入DB筆數)，供呼叫端(RenewCacheToDB)彙整顯示，不做新統計邏輯
        Try
            If Not quiet Then PgrsBar1.Text = "正在存入快取..." : Cursor = Cursors.WaitCursor

            ' ① 先清孤兒：收集目前記憶體快取中所有仍存在的 folder_path，清除 DB 中已不存在的行
            ' 用記憶體快取的 key 聯集代表「目前已知 live 的資料夾」 (比重新 BFS 掃 COM 快得多)
            ' 2026/07/03 by Simon/Claude: livePaths 建構含 folder_info 全表掃描，原本跑在 UI 執行緒，
            '   存檔開始瞬間會凍結 UI 一下 → 整段移入 Task.Run
            Dim livePaths As HashSet(Of String) = Await Task.Run(
                Function()
                    Dim lp As New HashSet(Of String)(1024)
                    For Each k In _cacheMailCount.Keys : lp.Add(k) : Next
                    For Each k In _cacheFolderCount.Keys : lp.Add(k) : Next
                    For Each k In _cacheAttMailList.Keys : lp.Add(k) : Next

                    ' 2026/06/12 by Simon/Claude Opus 4.8: lazy-load 安全保護
                    ' _cacheMailCount 等字典因 lazy-load 在重啟後可能不完整（記憶體中看不到 ≠ Outlook 中已刪除）
                    ' 把 folder_info 現有路徑全部列為 live，確保 CleanupOrphanPath 不誤刪仍存在的資料夾
                    ' 真正的孤兒清理由 RenewCacheToDB（完整 COM BFS 逐一 GetFolderFromID 確認）負責
                    Using readCmd As New SqliteCommand("SELECT folder_path FROM folder_info", _dbCache)
                        Using reader = readCmd.ExecuteReader()
                            While reader.Read() : lp.Add(reader.GetString(0)) : End While
                        End Using
                    End Using
                    Return lp
                End Function)

            If livePaths.Count > 0 Then Await CleanupOrphanPath(livePaths)

            ' 2026/07/03 by Simon/Claude Fable 5: dirty 追蹤 — 存檔前先拍一張快照。四張逐封/逐年表的 Save*Inner
            ' 只重寫快照裡的資料夾，其餘(記憶體與 DB 早已一致)完全跳過，取代原本「每次全量照單全收」。
            ' 快照獨立於下面的背景寫入：存檔期間若又有新的 COM 讀取進來把某夾標記 dirty，那筆新標記不在這次快照裡，
            ' 不會被本次存檔誤清掉，留給下一次 SaveCache 處理。
            Dim dirtySnapshot As New HashSet(Of String)(_dirtyMailFolders.Keys)

            ' ② SQLite I/O 在背景執行緒，不阻塞 UI
            ' 2026/07/03 by Simon/Claude: 大交易改走「專用寫入連線」dbW，不再與 UI 執行緒共用 _dbCache。
            '   原本寫入的數秒內，UI 的 lazy read (LazyGetMailInfo/LazyGetAttFilenames…) 會在同一條連線的
            '   mutex 上排隊等寫入語句放行 → 卡頓；WAL 天生一寫多讀，分開連線後互不阻塞。
            '   各 Batch 函式已改用 txn.Connection 取得連線，故只需把 txn 開在 dbW 上即可。
            Dim r = Await Task.Run(Function()
                                       Using dbW As New SqliteConnection($"Data Source={_dbCachePath};Mode=ReadWrite")
                                           dbW.Open()
                                           ' journal_mode=WAL 已持久化在 db 檔內，新連線自動繼承；synchronous 是 per-connection 設定必須重設
                                           Using pragmaCmd As New SqliteCommand("PRAGMA synchronous=NORMAL;", dbW) : pragmaCmd.ExecuteNonQuery() : End Using
                                           Using txn As SqliteTransaction = dbW.BeginTransaction()
                                               Try
                                                   Dim f = SaveFolderInfoBatch(txn)
                                                   Dim b = SaveAttMailListBatch(txn, dirtySnapshot)
                                                   'Dim a = SaveAttFilenameBatch(txn)                  ' 2026/06/21 by Simon/Claude Opus 4.8: att_filenames 已搬至 OLAcacheMail.db(_dbMail)，跨檔不掛 _dbCache txn，改自管獨立交易(主 db 已 commit 後再寫)
                                                   Dim y = SaveYearCountBatch(txn, dirtySnapshot)
                                                   Dim m = SaveMonthCountBatch(txn, dirtySnapshot)     ' 2026/04/09 新增
                                                   Dim s = SaveSendersBatch(txn, dirtySnapshot)         ' 2026/06/12 by Simon/Claude Opus 4.8: 先建立 senders 表，供 SaveMailInfoBatch 查 sender_id
                                                   Dim basic = SaveMailInfoBatch(txn, dirtySnapshot)
                                                   txn.Commit()

                                                   Dim a = SaveAttFilenameBatch()                      ' 2026/06/21 by Simon/Claude Opus 4.8: att_filenames 已搬至 OLAcacheMail.db(_dbMail)，跨檔不掛 _dbCache txn，改自管獨立交易(主 db 已 commit 後再寫)
                                                   Return (f, b, a, y, m, s, basic)
                                               Catch ex As System.Exception
                                                   txn.Rollback() : Throw
                                               End Try
                                           End Using
                                       End Using
                                   End Function)

            savedFolders = r.f : savedAttMailList = r.b : savedAttFilenames = r.a
            Dim savedYears As Integer = r.y, savedMonths As Integer = r.m
            Dim savedSenders As Integer = r.s : savedBasic = r.basic
            sw.Stop()

            ' 2026/07/03 by Simon/Claude Fable 5: 上面 Task.Run 成功返回(沒被 Catch 攔截)才會執行到此行，
            ' 代表快照內的資料夾這次確實已寫入 DB，可以安全從 dirty 集合移除；若寫入失敗，例外會跳去下面的 Catch，
            ' 不會執行到這裡，dirty 標記維持不變，下次 SaveCache 會重試。
            For Each p In dirtySnapshot : _dirtyMailFolders.TryRemove(p, Nothing) : Next

            ' ③ 統計：各快取字典目前的 entry 數
            Dim st = GetDBSummary()
            Dim statLine1 = $"① [記憶體] MailCount: {_cacheMailCount.Count} / MailCountAll: {_cacheMailCountAll.Count} / FolderCount: {_cacheFolderCount.Count} / FolderCountAll: {_cacheFolderCountAll.Count}"
            Dim statLine2 = $"② [記憶體] FolderSize: {_cacheFolderSize.Count} / FolderSizeAll: {_cacheFolderSizeAll.Count} / AttPreScan: {_cacheAttMailList.Count} / AttFilename: {_cacheAttFilename.Count}"
            Dim statLine3 = $"③ [寫入DB] folder_info: {savedFolders} 筆 / mail_info: {savedBasic} 筆 / att_maillist: {savedAttMailList} 筆 / att_filenames: {savedAttFilenames} 筆 / year_count: {savedYears} 筆 / month_count: {savedMonths} 筆 / 耗時: {sw.Elapsed.TotalSeconds:0.000} 秒"
            Dim statLine4 = $"④ [DB現況] folder_info: {st.fc} 筆 / att_maillist: {st.mb} 筆 / att_filenames: {st.at} 筆 / year_count: {st.yc} 筆 / month_count: {st.mc} 筆 / mail_simhash: {st.sh} 筆(bigram_set 已回填 {st.bs} 筆/{st.bsMB:F0}MB) / 檔案: {st.kb} KB"   ' 2026/07/07 by Simon/Claude: 補 mail_simhash/bigram_set 統計
            resultLine = statLine3   ' 2026/07/07 by Simon/Claude: 供 RenewCacheToDB 彙整回傳

            If Not quiet Then PgrsBar1.Text = $"SaveCache 完成 — 耗時: {sw.Elapsed.TotalSeconds:0.000} 秒" : PgrsBar2.Text = statLine4
            _dbg(" ├ ", statLine1)
            _dbg(" ├ ", statLine2)
            _dbg(" ├ ", statLine3)
            _dbg(" ├ ", statLine4)

        Catch ex As System.Exception
            If Not quiet Then PgrsBar1.Text = "SaveCache 失敗"
            _dbg("       ├ 錯誤", ex.Message)
        Finally
            If Not quiet Then Cursor = Cursors.Default
            _dbg(" ├ 結束") ' by Gemini, 2026/04/11: 修正對應開始層級 Level 1
        End Try
        Return resultLine

    End Function
    Private Async Function LoadCachesFromDB() As Task
        ' ---------------------------------------------------------------
        ' LoadCachesFromDB — 從 SQLite 讀回所有快取 (Bulk Load) 
        ' 對應 Setting 頁 LoadCache 按鈕，Debug 階段使用
        ' 完成後輸出詳細 _dbg 分項：每個快取字典各自載入了幾筆
        ' ---------------------------------------------------------------
        _dbg("開始")
        If _dbCache Is Nothing Then _dbg("", "DB 未初始化") : Return

        Dim sw As Diagnostics.Stopwatch = Diagnostics.Stopwatch.StartNew()  ' by Gemini 3.5 Flash, 2026/06/07
        Try
            PgrsBar1.Text = "正在載入快取..." : Cursor = Cursors.WaitCursor

            ' 記錄 Load 前各字典的 entry 數，方便比對新增了多少
            Dim beforeMC = _cacheMailCount.Count
            Dim beforeMCA = _cacheMailCountAll.Count
            Dim beforeFC = _cacheFolderCount.Count
            Dim beforeFCA = _cacheFolderCountAll.Count
            Dim beforeFS = _cacheFolderSize.Count
            Dim beforeFSA = _cacheFolderSizeAll.Count
            Dim beforePS = _cacheAttMailList.Count
            Dim beforeAF = _cacheAttFilename.Count

            Dim r = Await Task.Run(Function()
                                       Dim f = LoadFolderInfoBatch()
                                       Dim b = LoadAttMailListBatch()
                                       Dim a = LoadAttFilenamesBatch()
                                       Dim y = LoadYearCountBatch()    ' 2026/07/08 消重：原 LoadYearCountBatch / LoadMonthCountBatch 合併於 LoadDateCountCore
                                       Dim m = LoadMonthCountBatch()   ' 2026/04/09 新增
                                       Dim s = LoadSendersBatch()           ' 2026/06/12 by Simon/Claude Opus 4.8: 先載入 senders 字典，供 LoadMailInfoBatch 解析 sender_id
                                       Dim basic = LoadMailInfoBatch() ' 2026/04/22 by Gemini 3.1 Pro 新增
                                       Return (f, b, a, y, m, s, basic)
                                   End Function)
            sw.Stop()

            ' 詳細 _dbg：各快取字典 Load 後的增量
            Dim statLine1 = $"① [folder_info] 讀入 {r.f} 筆 — " &
                            $"MailCount +{_cacheMailCount.Count - beforeMC} / " &
                            $"MailCountAll +{_cacheMailCountAll.Count - beforeMCA} / " &
                            $"FolderCount +{_cacheFolderCount.Count - beforeFC} / " &
                            $"FolderCountAll +{_cacheFolderCountAll.Count - beforeFCA}"
            Dim statLine2 = $"② [folder_info cont.] " &
                            $"FolderSize +{_cacheFolderSize.Count - beforeFS} / " &
                            $"FolderSizeAll +{_cacheFolderSizeAll.Count - beforeFSA}"
            Dim statLine3 = $"③ [att_maillist] 讀入 {r.b} 筆 → AttPreScan +{_cacheAttMailList.Count - beforePS} 個資料夾"
            Dim statLine4 = $"④ [att_filenames] 讀入 {r.a} 筆 → AttFilename +{_cacheAttFilename.Count - beforeAF} 筆"
            Dim statLine_yc = $"⑤ [year_count] 讀入 {r.y} 筆 → _yearCountCache {_cacheYearCount.Count} 個資料夾"
            Dim statLine_mc = $"⑥ [month_count] 讀入 {r.m} 筆 → _monthCountCache {_cacheMonthCount.Count} 個 cache_key"
            Dim statLine_basic = $"⑦ [mail_info] 讀入 {r.basic} 筆 → BasicPreScan {_cacheMailInfo.Count} 個資料夾" ' 2026/04/22 by Gemini 3.1 Pro
            Dim st = GetDBSummary()
            Dim statLine5 = $"⑧ [DB現況] folder_info: {st.fc} 筆 / mail_info: {st.basic} 筆 / att_maillist: {st.mb} 筆 / att_filenames: {st.at} 筆 / year_count: {st.yc} 筆 / month_count: {st.mc} 筆 / 檔案: {st.kb} KB / 耗時: {sw.Elapsed.TotalSeconds:0.000} 秒" ' 2026/04/22 by Gemini 3.1 Pro: 加入 basic 統計

            PgrsBar1.Text = $"LoadCache 完成 — DB: {st.fc}/{st.basic}/{st.mb}/{st.at}/{st.yc}/{st.mc} 筆，{st.kb} KB，耗時 {sw.Elapsed.TotalSeconds:0.000} 秒" ' 2026/04/22 by Gemini 3.1 Pro
            PgrsBar2.Text = $"記憶體增量 — mailCount+{_cacheMailCount.Count - beforeMC} / attFilename+{_cacheAttFilename.Count - beforeAF} / basicMailInfo:{_cacheMailInfo.Count} 資料夾" ' 2026/04/22 by Gemini 3.1 Pro
            _dbg(" ├ ", statLine1)
            _dbg(" ├ ", statLine2)
            _dbg(" ├ ", statLine3)
            _dbg(" ├ ", statLine4)
            _dbg(" ├ ", statLine_yc)
            _dbg(" ├ ", statLine_mc)
            _dbg(" ├ ", statLine_basic) ' 2026/04/22 by Gemini 3.1 Pro
            _dbg(" ├ ", statLine5)

        Catch ex As System.Exception
            PgrsBar1.Text = "LoadCache 失敗"
            _dbg("錯誤", ex.Message)
        Finally
            Cursor = Cursors.Default
            _dbg("結束")
        End Try

    End Function
    Private Async Function RenewCacheToDB() As Task(Of String)
        ' ---------------------------------------------------------------
        ' RenewCacheToDB — 完整更新 DB 快取 (對應 Setting 頁 RenewCache 按鈕) 
        '
        ' 與 SaveCachesToDB 的差異：
        '   SaveCache  = 把目前記憶體快取照單全收寫入 DB (不更新過期的值) 
        '   RenewCache = 先用 COM 比對 snapshot → 只對有變動的資料夾重新計算 → 寫入 DB
        '
        ' 流程：
        '   Phase 1. BFS 掃出所有 live folders (COM，~1ms/資料夾) 
        '   Phase 2. 每個 folder 讀 PeekFolderTimeSnapOOM (單次 GetProperties 批次讀 count+commit_max) vs DB snapshot → 找 dirty folders
        '            (2026/07/04 三訊號: count 抓增刪、commit_max 抓淨零置換/內容修改、兩者皆淨時再抽樣 GetItemFromID 探活抓純壓縮換ID)
        '   Phase 3. 對每個 dirty folder 重新計算：
        '              mc/fc (快，~1ms) 
        '              year_count (GetTable + GetArray，~10-50ms/資料夾) 
        '              month_count (清記憶體， Phase5 清 DB， 展開時 lazy 重算) 
        '              att_maillist (GetTable 三路比對，~5ms/資料夾) 
        '              folder_size (選擇性，依 includeSize，GetTable 遍歷，~10-30s/大資料夾) 
        '              清除 mca/fca/fsa 聚合快取 (讓下次點選重算) 
        '              清除此 folder 的 month_count 記憶體快取 (不重算，展開年份時 lazy) 
        '   Phase 4. 清除所有 dirty folders 的 ancestor 聚合快取
        '   Phase 5. 批次 DELETE dirty folders 的 month_count DB rows (不是孤兒，不靠 Cleanup) 
        '   Phase 6. CleanupOrphanFolderPath → SaveCachesToDB
        '
        ' 不更新項目 (設計邊界) ：
        '   att_filenames   — 最耗時，留給使用者搜尋附件時 lazy 觸發
        '   month_count     — 清記憶體 + 清 DB，展開年份時 lazy 重算
        ' 2026/04/09 by Claude
        ' ---------------------------------------------------------------
        ' 2026/04/16 by Simon/Claude: 加入 cToken (OkayNowYouHaveToken)，取代 _cancelRequested + GoTo Cancelled 模式
        '   Phase1 改用 Dictionary(Of String, Outlook.Folder) liveDict，每個資料夾只讀一次 FolderPath COM 屬性，
        '   Phase2/3/4 迭代 dict 的 Key/Value，完全省去重複的 folder.FolderPath COM 呼叫（~500 資料夾省 ~250ms）
        '   Phase2/3 節流改用 SmartThrottle(sw, cToken, ThrottleFreq.Low)，取代 Mod N + Task.Delay(1)
        '   GetYearCountOOM / GetFolderSizeOOM 補入 cToken:=cToken
        ' ---------------------------------------------------------------
        ' 2026/05/17 by simon/Gemini: RenewCacheToDB 大幅重構，改為「精確打擊」模式，
        '   不再 BFS 展開每個資料夾的子樹來找對應的 DB row，而是直接從 DB 撈出全部資料夾清單，然後用 GetFolderFromID 精確抓出 COM 物件，比對 snapshot 決定是否 dirty
        ' 2026/6/7: by simon/Gemini: 去除原本函數內的多段計時和狀態顯示, 直接在RenewCache_Click事件中計時顯示整體耗時
        ' ---------------------------------------------------------------
        If _dbCache Is Nothing Then Return ""
        Dim cToken As Threading.CancellationToken = OkayNowYouHaveToken()   ' 2026/07/07 by Simon/Claude: 移到早退檢查後，與下方 Finally 的 OkeyNowByeByeToken 成對
        Dim resultSummary As String = ""   ' 2026/07/07 by Simon/Claude: 彙整既有統計數字供 RenewCache_Click 顯示，不做新統計邏輯

        _dbg("開始")
        Try
            ' 1. 【Rule 1】由專用函數撈出全部追蹤名單，維持主流程乾淨
            Dim dbList = LoadAllFolderInfo()
            If dbList.Count = 0 Then _dbg("RenewCache", "資料庫無快取紀錄，略過") : Return ""

            ' 建立您原本 CleanupOrphanPath 所需的活著名單
            Dim liveFolderPaths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim updatedPaths As New HashSet(Of String)() ' 用於 Rule 6：記錄需要失效父子聚合值的路徑
            Dim orphanFolderCount As Integer = 0     ' 2026/07/07 by Simon/Claude: 累加既有計數，供結尾彙整顯示，不做新統計邏輯
            Dim totalDeadRemoved As Integer = 0      ' 累加 absent.Count(既有值)：既存資料夾內被判定死亡而清除的 entry_id 數
            Dim totalModifiedDetected As Integer = 0 ' 累加 modified.Count(既有值)：既存資料夾內偵測到內容變更(msg_size 改變)的 entry_id 數

            Dim swThrottle As Stopwatch = Stopwatch.StartNew()  ' by Gemini 3.5 Flash, 2026/06/07
            Dim processed = 0

            _dbg("開始逐一檢查資料夾路徑...")
            For Each row In dbList
                cToken.ThrowIfCancellationRequested()
                Dim fPath = row.path

                ' 2. 【Rule 7】精確打擊，用 ID 抓取物件，不使用 BFS 展開
                Dim folder As Outlook.Folder = Nothing
                Try
                    folder = TryCast(_olNS.GetFolderFromID(row.eid, row.sid), Outlook.Folder)
                Catch : End Try

                ' 3. 【Rule 3】Outlook 裡面已經不存在 -> 判定為孤兒
                If folder Is Nothing Then
                    ' 手動清除記憶體字典，避免殘留幽靈快取
                    _cacheMailInfo.TryRemove(fPath, Nothing)
                    _cacheFolderIDs.TryRemove(fPath, Nothing)
                    _cacheMailCount.TryRemove(fPath, Nothing)
                    _cacheFolderCount.TryRemove(fPath, Nothing)
                    _cacheFolderSize.TryRemove(fPath, Nothing)
                    _cacheAttMailList.TryRemove(fPath, Nothing)
                    _cacheYearCount.TryRemove(fPath, Nothing)
                    _cacheFolderCommitMax.TryRemove(fPath, Nothing)   ' 2026/07/04 by Simon/Claude Fable 5: 孤兒一併清 commit 基準
                    ClearMonthCountMemory(fPath)

                    updatedPaths.Add(fPath) ' 標記這條路徑不見了，上層父資料夾的聚合值需要更新
                    orphanFolderCount += 1   ' 2026/07/07 by Simon/Claude: 累加既有判定，供結尾彙整顯示
                    Continue For            ' 死了就不加入 liveFolderPaths，等一下交給您的 CleanupOrphanPath 去刪 DB
                End If

                ' 4. 如果資料夾在 Outlook 還活著，加入活著名單集合
                liveFolderPaths.Add(fPath)

                ' 5. 【Rule 4 & 2】比對 Snap 與更新邏輯
                ' 2026/07/04 by Simon/Claude Fable 5: 雙訊號 dirty 判定 — 純 count 抓不到「數量不變但內容置換」(copy→修改→放回→刪原始 = 淨零變動)，加入 PR_LOCAL_COMMIT_TIME_MAX 比對。
                '   lastUpdated 刻意在重掃「之前」讀取：重掃期間若又有變動，commit 會高於本次存值 → 下次 Renew 再補抓，方向保守安全。
                '   任一端未知 (-1 / DB NULL) 時退回純 count 比對，不誤判 dirty；首次升級後 DB 全為 NULL → 狀況 B 採認現值當基準，防護自下次生效。
                ' 2026/07/11 by Simon/Claude Fable 5 [RenewCache 例外歸零]: 原兩次 GetProperty(PeekLiveFolderSnapOOM+PeekFolderLastUpdateTime)
                '   併為單次 GetProperties 批次 — 缺屬性不再拋例外(舊寫法每夾每次噴一顆 COMException)，並少一次 COM 往返。
                Dim peeked = PeekFolderTimeSnapOOM(folder, fPath)
                Dim liveSnap As Integer = peeked.snap
                Dim lastUpdated As Long = peeked.cmx
                Dim commitDirty As Boolean = lastUpdated >= 0 AndAlso row.cmx >= 0 AndAlso lastUpdated <> row.cmx
                Dim isDirty As Boolean = (liveSnap <> row.snap) OrElse commitDirty

                ' 2026/07/04 by Simon/Claude Sonnet 5: 第三訊號 — 純 PST 壓縮換 entry_id 時，count 和 commit_max 都不會變(壓縮不增減、不改動任何郵件)，
                '   前兩訊號必然雙雙判定「沒變」。只在雙訊號都乾淨時才做，抽一筆已知 entry_id 單次 GetItemFromID 探活(不做整夾表掃，維持狀態B「近乎0 COM」的設計)。
                '   壓縮通常整批 entry_id 一起換掉，抽一筆探測到就能立刻讓這個資料夾就地升級為狀態A全量重讀+surgical清理，不必等使用者剛好對這夾按過 F5 才被動發現。
                If Not isDirty Then
                    Dim sampleEid = PeekLiveFolderId(fPath)
                    If sampleEid <> "" Then
                        Dim probeOk As Boolean = False
                        Try : probeOk = (_olNS.GetItemFromID(sampleEid, row.sid) IsNot Nothing) : Catch : End Try
                        If Not probeOk Then isDirty = True
                    End If
                End If

                If isDirty Then
                    ' 狀況 A：Snap 不一致！代表 Outlook 有變動，進行 Layer 3 COM 讀取更新記憶體
                    _cacheMailCount(fPath) = GetMailCount(folder, fPath, skipCache:=True)       ' 2026/06/23 by Simon/Claude: 狀況A snap 重讀改走 proxy skipCache(RDO 派工)
                    _cacheFolderCount(fPath) = GetFolderCount(folder, fPath, skipCache:=True)   ' 2026/06/23 by Simon/Claude: 同上
                    _cacheFolderSize(fPath) = Await GetFolderSize(folder, fPath:=fPath, skipCache:=True, cToken:=cToken)    ' 2026/6/27 by simon/Claude Opus 4.8: 整合GetFolderSize單一入口再分派RDO/OOM, 加skipCache參數讓DB 重建的強制重讀也吃得到 GetFolderSizeRdo的提速

                    ' 2026/06/22 by Simon/Claude: 缺口1+2 ② Surgical 嚴格清除 —
                    '   (a) 取 live 全郵件 entryID，算「DB 有、live 沒有」的已刪集合 → surgical 清兩張昂貴快取
                    '       (att_filenames/mail_simhash 之 DB 列 + 記憶體)，存活郵件的昂貴快取保留免重讀。
                    '   (b) 便宜逐封表(basic/att_maillist/month_count/year_count)整夾 nuke DB 死列；對應記憶體一併清/重建，
                    '       否則尾端 SaveCachesToDB 會把舊鬼魂列寫回，使 nuke 失效。
                    ' 2026/07/07 by Simon/Claude: 舊清單改用加寬版(帶 msg_size)，供下方就地修改偵測；absent 差集邏輯不變
                    Dim oldSizes = LazyGetFolderIdAsDict(fPath)
                    Dim liveAll = Await GetMailInfoOOM(folder, needTopic:=False, cToken:=cToken, fPath:=fPath)
                    Dim liveSet As New HashSet(Of String)(liveAll.Select(Function(m) m.Mail.EntryID))
                    Dim absent = oldSizes.Keys.Where(Function(e) Not liveSet.Contains(e)).ToList()
                    If absent.Count > 0 Then SimDbDeleteMailRowsByEntryIds(absent, includeAttFilenames:=True)
                    totalDeadRemoved += absent.Count   ' 2026/07/07 by Simon/Claude: 累加既有值，供結尾彙整顯示

                    ' 2026/07/07 by Simon/Claude: (c) 就地修改盲點 — 內容被修改但 entry_id 不變的信，不會落入 absent 差集，
                    '   其 simhash/bigram_set/att_filenames/body 快取全部過期卻永不失效。以「同 entry_id 但 msg_size 改變」
                    '   精準偵測(任何內文/附件修改幾乎必變 size)；新信/刪信不受影響，不會造成整夾重算。
                    '   清掉後下次 Tab5 掃描 S3 對這幾封自動補算指紋，附件名/內文則由各自的 lazy 路徑重讀。
                    Dim modified = liveAll.Where(Function(m) oldSizes.ContainsKey(m.Mail.EntryID) AndAlso
                                                             oldSizes(m.Mail.EntryID) >= 0 AndAlso
                                                             oldSizes(m.Mail.EntryID) <> m.Mail.Size).
                                           Select(Function(m) m.Mail.EntryID).ToList()
                    If modified.Count > 0 Then
                        SimDbDeleteMailRowsByEntryIds(modified, includeAttFilenames:=True)   ' 內建同步清 _cacheSimHash/_cacheAttFilename
                        For Each eid In modified : _cacheMailBody.TryRemove(eid, Nothing) : Next   ' body 快取也已過期，一併清
                        _dbg(" ├ 修改偵測", $"{fPath}: {modified.Count} 封 msg_size 有變，已清指紋/附件名/內文快取")
                    End If
                    totalModifiedDetected += modified.Count   ' 2026/07/07 by Simon/Claude: 累加既有值，供結尾彙整顯示
                    DbPurgeFolderMailRows(fPath, includeAttFilenames:=False) ' 整夾 nuke 便宜表 DB(含 EMPTY_BASIC 哨兵)
                    _cacheAttMailList.TryRemove(fPath, Nothing)              ' 配合 DB nuke，避免 SaveCache 寫回舊 att_maillist
                    ClearMonthCountMemory(fPath)                             ' 月計數同步失效，展開年份時 lazy 重算

                    _cacheYearCount.TryRemove(fPath, Nothing)
                    _cacheMailInfo(fPath) = (liveAll, liveSnap) ' 既已掃描就存回(取代原 TryRemove)，SaveCache 以新 snap 重寫 mail_info，省一次 lazy 重掃
                    MarkMailFolderDirty(fPath)                  ' 2026/07/03 by Simon/Claude: dirty 追蹤 — 剛用 COM 重新掃過，SaveCache 必須重寫此夾
                    _cacheFolderIDs(fPath) = (folder.EntryID, folder.StoreID, IsMailFolder(folder, fPath), True)
                    If lastUpdated >= 0 Then _cacheFolderCommitMax(fPath) = lastUpdated   ' 2026/07/04 by Simon/Claude Fable 5: 重掃完成，以掃前讀到的 commit 當新基準
                    updatedPaths.Add(fPath) ' 標記有異動
                Else
                    ' 狀況 B：Snap 一致！代表 Outlook 沒變。
                    ' 此時若記憶體跟 DB 不符，直接拿 DB 當權威「抄回記憶體」（0次重型 COM 呼叫）
                    Dim memCount As Long = -1
                    If Not _cacheMailCount.TryGetValue(fPath, memCount) OrElse memCount <> row.mc Then
                        _cacheMailCount(fPath) = row.mc
                        _cacheFolderCount(fPath) = row.fc
                        _cacheFolderIDs(fPath) = (row.eid, row.sid, IsMailFolder(folder, fPath), True)
                    End If
                    If lastUpdated >= 0 Then _cacheFolderCommitMax(fPath) = lastUpdated   ' 2026/07/04 by Simon/Claude Fable 5: DB 為 NULL 時採認現值當基準(升級後首跑)；已有值時等值覆寫無害
                End If
                processed += 1
                Await SmartThrottle(swThrottle, ThrottleFreq.Low, Sub() PgrsBar2.Text = $"對帳中 {processed}/{dbList.Count}...", cToken:=cToken)
            Next

            ' 6. 【Rule 3】安全無縫套用：直接呼叫您原本寫好的 CleanupOrphanPath 清理 5 個資料表
            _dbg("清理孤兒資料夾路徑...")
            Dim cleanupSummary = Await CleanupOrphanPath(liveFolderPaths)

            ' 7. 【Rule 6】同步失效相關路徑的上下父子聚合快取（All 結尾的總加總數值）
            If updatedPaths.Count > 0 Then
                _dbg("同步失效相關路徑的上下父子聚合快取...")
                For Each p In updatedPaths
                    For Each ancestor In GetAncestors(p)
                        _cacheMailCountAll.TryRemove(ancestor, Nothing)
                        _cacheMailCountAll.TryRemove(ancestor & "|True", Nothing)
                        _cacheMailCountAll.TryRemove(ancestor & "|False", Nothing)
                        _cacheFolderCountAll.TryRemove(ancestor, Nothing)
                        _cacheFolderSizeAll.TryRemove(ancestor, Nothing)
                    Next
                Next
            End If

            ' 8. 批次將更新後的記憶體快取存回 SSD 資料庫
            _dbg("正在將快取更新至資料庫...")
            Dim saveSummary = Await SaveCachesToDB()

            ' 2026/07/07 by Simon/Claude: 彙整既有數字(不做新統計邏輯，全是上面迴圈/CleanupOrphanPath/SaveCachesToDB 早就算好的值)
            '   對帳/判定異動 = Phase2/3 逐夾迴圈的 dbList.Count / updatedPaths.Count(既有變數)
            '   孤兒資料夾   = orphanFolderCount(既有判定，Rule3 Continue For 分支累加)
            '   夾內死亡/修改 = totalDeadRemoved / totalModifiedDetected(既有 absent.Count / modified.Count 累加)
            '   孤兒清理明細 = cleanupSummary(CleanupOrphanPath 既有 _dbg 字串)
            '   寫入DB明細   = saveSummary(SaveCachesToDB 既有 statLine3)
            resultSummary = $"對帳 {dbList.Count} 夾,{updatedPaths.Count} 夾異動(孤兒 {orphanFolderCount})/ 夾內刪 {totalDeadRemoved} 封,修改 {totalModifiedDetected} 封" &
                            If(cleanupSummary <> "", $" | {cleanupSummary}", "") &
                            If(saveSummary <> "", $" | {saveSummary}", "")

        Catch ex As OperationCanceledException
            PgrsBar1.Text = "RenewCache 已由使用者中斷"
        Catch ex As System.Exception
            PgrsBar1.Text = $"RenewCache 失敗: {ex.Message}"
            _dbg("RenewCache 發生錯誤", ex.Message)
        Finally
            OkeyNowByeByeToken(cToken)   ' 2026/07/07 by Simon/Claude: 歸還 token — 運算中判定 token 化
            Cursor = Cursors.Default
            _dbg("結束")
        End Try
        Return resultSummary
    End Function
    Private Async Function CleanupOrphanPath(liveFolderPaths As HashSet(Of String)) As Task(Of String)
        ' ---------------------------------------------------------------
        ' CleanupOrphanPath — 刪除已不存在於 Outlook 的資料夾孤兒行 (改為非同步 by Gemini 3.1 Pro, 2026/05/05)
        ' liveFolderPaths = 目前仍有效的資料夾路徑集合
        '   呼叫來源 A: SaveCachesToDB 開頭 (用記憶體快取 key 聯集)
        '   呼叫來源 B: RenewCache_Click (用 GetSubtree BFS 掃 COM 取得完整清單)
        ' 2026/07/07 by Simon/Claude: 回傳既有的孤兒清除統計字串，供 RenewCacheToDB 彙整顯示，不做新統計邏輯
        ' ---------------------------------------------------------------
        _dbg("    ├ 開始", $"live 資料夾數: {liveFolderPaths.Count}") ' by Gemini, 2026/04/10: 調整縮排層級為 Level 2
        If _dbCache Is Nothing Then Return ""

        Dim summary As String = Await Task.Run(Function() As String
                                                   Try
                                                       ' 讀出 DB 中所有 folder_path
                                                       Dim dbPaths As New List(Of String)(2048)
                                                       Using cmd As New SqliteCommand("SELECT folder_path FROM folder_info", _dbCache)
                                                           Using reader = cmd.ExecuteReader()
                                                               While reader.Read() : dbPaths.Add(reader.GetString(0)) : End While
                                                           End Using
                                                       End Using
                                                       _dbg("", $"DB 中有 {dbPaths.Count} 個資料夾路徑")

                                                       Dim stalePaths = dbPaths.Where(Function(p) Not liveFolderPaths.Contains(p)).ToList()
                                                       If stalePaths.Count = 0 Then _dbg("", "未發現孤兒快取，略過") : Return ""

                                                       ' 2026/06/22 by Simon/Claude: ② Surgical — 整夾消失 → 該夾全部 entryID 皆失效。
                                                       '   先撈 mail_info 的 entryID(務必在下方 DELETE mail_info 之前)，供稍後清 mail_simhash
                                                       '   (無 folder_hash 只能靠 entryID) 與記憶體 _cacheSimHash/_cacheAttFilename。
                                                       Dim orphanEntryIds As New List(Of String)()
                                                       For Each s In stalePaths : orphanEntryIds.AddRange(LazyGetFolderIdAsList(s)) : Next

                                                       Dim dF, dB, dA, dM, dBasic, dSh, dY As Integer
                                                       Using txn As SqliteTransaction = _dbCache.BeginTransaction()
                                                           Using c1 As New SqliteCommand("DELETE FROM folder_info WHERE folder_path=@fp", _dbCache, txn),
                                       c2 As New SqliteCommand("DELETE FROM att_maillist WHERE folder_hash=@fh", _dbCache, txn),
                                       c4 As New SqliteCommand("DELETE FROM month_count WHERE folder_hash=@fh", _dbCache, txn),
                                       c5 As New SqliteCommand("DELETE FROM mail_info WHERE folder_hash=@fh", _dbCache, txn),
                                       c6 As New SqliteCommand("DELETE FROM year_count WHERE folder_hash=@fh", _dbCache, txn)
                                                               ' 2026/06/21 by Simon/Claude: att_filenames 已搬至 OLAcacheMail.db(_dbMail)，跨檔獨立刪除(原 c3)，dA 由此累計
                                                               ' c3 As New SqliteCommand("DELETE FROM att_filenames WHERE folder_hash=@fh", _dbCache, txn),
                                                               ' 2026/07/06 by Simon/Claude Fable 5: 補 c6 year_count — 孤兒資料夾的年份列原本永遠留在 DB，若日後同路徑資料夾重建(folder_hash 相同)，舊年份分佈會被 GetYearCount ② 原樣復活

                                                               c1.Parameters.Add("@fp", SqliteType.Text)
                                                               c2.Parameters.Add("@fh", SqliteType.Integer) ': c3.Parameters.Add("@fh", SqliteType.Integer)
                                                               c4.Parameters.Add("@fh", SqliteType.Integer) : c5.Parameters.Add("@fh", SqliteType.Integer)
                                                               c6.Parameters.Add("@fh", SqliteType.Integer)

                                                               For Each s In stalePaths
                                                                   Dim h = StringToXxHash64(s) ' 取得孤兒的 Hash
                                                                   c1.Parameters("@fp").Value = s : dF += c1.ExecuteNonQuery()
                                                                   c2.Parameters("@fh").Value = h : dB += c2.ExecuteNonQuery()
                                                                   ' c3.Parameters("@fh").Value = h : dA += c3.ExecuteNonQuery()
                                                                   c4.Parameters("@fh").Value = h : dM += c4.ExecuteNonQuery()
                                                                   c5.Parameters("@fh").Value = h : dBasic += c5.ExecuteNonQuery()
                                                                   c6.Parameters("@fh").Value = h : dY += c6.ExecuteNonQuery()
                                                               Next
                                                           End Using
                                                           txn.Commit()
                                                       End Using
                                                       For Each s In stalePaths : dA += SimDbDeleteAttFilenamesByFolder(s) : Next

                                                       ' 2026/06/22 by Simon/Claude: att_filenames 已由上行按 folder_hash 高效刪，故 includeAttFilenames:=False，僅補 mail_simhash + 兩記憶體快取
                                                       dSh = SimDbDeleteMailRowsByEntryIds(orphanEntryIds, includeAttFilenames:=False)

                                                       Dim line = $"孤兒路徑:{stalePaths.Count} 個 / folder_info:{dF} 行 / mail_info:{dBasic} 行 / att_maillist:{dB} 行 / att_filenames:{dA} 行 / mail_simhash:{dSh} 行 / month_count:{dM} 行 / year_count:{dY} 行"
                                                       _dbg("結束", line)
                                                       Return line

                                                   Catch ex As System.Exception
                                                       _dbg("    ├ 錯誤", ex.Message) ' by Gemini, 2026/04/10
                                                       Return ""
                                                   End Try
                                               End Function)

        Return summary
    End Function
    Private Async Sub timerSaveCache_Tick(sender As Object, e As EventArgs) Handles timerSaveCache.Tick
        ' ---------------------------------------------------------------
        ' timerSaveCache_Tick — 每 60 秒自動存檔一次 (InitDatabase 啟動計時)
        ' 略過條件：正在關閉(FormClosing 已自行呼叫一次)、上一輪自動存檔尚未跑完、或沒有 dirty 資料夾需要寫
        ' ---------------------------------------------------------------
        If _isClosing Then Return
        If _isAutoSavingCache Then Return
        If _cts IsNot Nothing Then Return          ' 2026/07/07 by Simon/Claude: 前景作業進行中(含 ESC 後收尾)先避讓，不跟掃描搶磁碟/CPU；dirty 資料留給下一輪
        If _dirtyMailFolders.IsEmpty Then Return   ' 2026/07/03 dirty 追蹤上線後，沒有異動就不必觸發整段 SaveCachesToDB

        _isAutoSavingCache = True
        Try
            Await SaveCachesToDB(quiet:=True)   ' 2026/07/06 by Simon/Claude: 自動存檔走安靜模式,不碰 PgrsBar/Cursor(詳見 SaveCachesToDB 註解)
        Catch ex As System.Exception
            _dbg("       ├ 錯誤", ex.Message)
        Finally
            _isAutoSavingCache = False
        End Try
    End Sub
#End Region

#Region "■ 批次寫入核心 (Batch Writer Core)"
    Private Function SaveDateCountCore(txn As SqliteTransaction, dirtyPaths As HashSet(Of String), onlyYearCount As Boolean) As Integer
        ' ---------------------------------------------------------------
        ' SaveDateCountCore — year_count / month_count 批次寫入共用核心 (Tab2 年份/月份分布)
        '   onlyYearCount = True    : 寫 year_count，來源 _cacheYearCount (key=folder_path, value={year→count})
        '   onlyYearCount = False   : 寫 month_count，來源 _cacheMonthCount (key=folder_path_year，需解析出 fPath/year, value={month→count})
        ' 2026/07/03 by Simon/Claude Fable 5: 加入 dirtyPaths 過濾 — 只重寫 dirty 的資料夾，其餘(從 DB lazy load 進記憶體、內容跟 DB 一致)整個 Continue For 跳過，不再每次全量重寫
        ' 2026/07/09 by Simon/Claude: 消重 — 原 SaveYearCountBatch / SaveMonthCountBatch 兩函式除欄位數與來源字典 key 格式外幾乎同構，
        '   比照 LoadDateCountCore 的模式合併，呼叫端保留 SaveYearCountBatch/SaveMonthCountBatch 兩個具名薄 wrapper。
        ' ---------------------------------------------------------------
        _dbg(" ├ 開始", If(onlyYearCount, "year_count", "month_count"))
        Dim sql = If(onlyYearCount, "INSERT OR REPLACE INTO year_count (folder_hash,year,count) VALUES (@fh,@yr,@cnt)",
                                    "INSERT OR REPLACE INTO month_count (folder_hash,year,month,count) VALUES (@fh,@yr,@mo,@cnt)")
        Dim source = If(onlyYearCount, _cacheYearCount, _cacheMonthCount)
        Dim count As Integer = 0

        Using cmd As New SqliteCommand(sql, txn.Connection, txn)   ' 2026/07/03 by Simon/Claude: 改跟隨 txn 所在連線
            ' 2026/07/03 by Simon/Claude Fable 5: 參數物件存區域變數，免去迴圈內名稱線性查找
            Dim pFh = cmd.Parameters.Add("@fh", SqliteType.Integer)
            Dim pYr = cmd.Parameters.Add("@yr", SqliteType.Integer)
            Dim pMo As SqliteParameter = Nothing : If Not onlyYearCount Then pMo = cmd.Parameters.Add("@mo", SqliteType.Integer)
            Dim pCnt = cmd.Parameters.Add("@cnt", SqliteType.Integer)

            For Each kvp In source
                Dim fPath As String
                Dim yearVal As Integer = Integer.MinValue
                If onlyYearCount Then
                    fPath = kvp.Key
                Else
                    ' cache_key 格式: "FolderPath_year"，最後一個 "_" 分隔出 year
                    Dim cacheKey = kvp.Key
                    Dim lastUnderscore = cacheKey.LastIndexOf("_"c)
                    If lastUnderscore < 0 Then Continue For
                    fPath = cacheKey.Substring(0, lastUnderscore)
                    If Not Integer.TryParse(cacheKey.Substring(lastUnderscore + 1), yearVal) Then Continue For
                End If
                If Not dirtyPaths.Contains(fPath) Then Continue For   ' 2026/07/03 by Simon/Claude Fable 5: dirty 過濾

                ' 2026/06/12 by Gemini/Simon: 把 folder_path TEXT 改成 folder_hash INTEGER 儲存，並新增一個記憶體字典 _dictHashToPath 來反查雜湊對應的路徑
                Dim fh As Long = FolderPathToHash64(fPath)   ' 2026/07/03 by Simon/Claude: 提到迴圈外每夾算一次，原本每列重算徒增 GC 壓力
                For Each item In kvp.Value
                    pFh.Value = fh
                    pYr.Value = If(onlyYearCount, Item.Key, yearVal)
                    If Not onlyYearCount Then pMo.Value = Item.Key
                    pCnt.Value = Item.Value
                    cmd.ExecuteNonQuery() : count += 1
                Next
            Next
        End Using

        _dbg(" ├ 結束", $"{count} 筆")
        Return count

    End Function
    Private Function SaveYearCountBatch(txn As SqliteTransaction, dirtyPaths As HashSet(Of String)) As Integer
        ' 薄 wrapper — year_count 批次寫入，委派 SaveDateCountCore (2026/07/09 by Simon/Claude)
        Return SaveDateCountCore(txn, dirtyPaths, onlyYearCount:=True)
    End Function
    Private Function SaveMonthCountBatch(txn As SqliteTransaction, dirtyPaths As HashSet(Of String)) As Integer
        ' 薄 wrapper — month_count 批次寫入，委派 SaveDateCountCore (2026/07/09 by Simon/Claude)
        Return SaveDateCountCore(txn, dirtyPaths, onlyYearCount:=False)
    End Function

    Private Function SaveAttMailListBatch(txn As SqliteTransaction, dirtyPaths As HashSet(Of String)) As Integer
        ' ---------------------------------------------------------------
        ' SaveAttMailListBatch — Transaction 內批次寫入 att_maillist (Tab3 Phase1)
        ' 2026/06/12 by Simon/Claude Opus 4.8: received_time TEXT→INTEGER (Unix秒)；移除 updated_at
        ' 2026/07/03 by Simon/Claude Fable 5: 加入 dirtyPaths 過濾，同 SaveYearCountBatch
        ' ---------------------------------------------------------------
        _dbg("開始")
        Dim sql = "INSERT OR REPLACE INTO att_maillist" &
                  " (entry_id,folder_hash,subject,msg_size,received_time,sender_name,att_count,pr_count_snap)" &
                  " VALUES (@eid,@fh,@subj,@sz,@rt,@sn,@ac,@pr)"

        Dim count As Integer = 0
        Using cmd As New SqliteCommand(sql, txn.Connection, txn)   ' 2026/07/03 by Simon/Claude: 改跟隨 txn 所在連線
            ' 2026/07/03 by Simon/Claude Fable 5: 參數物件存區域變數，免去迴圈內名稱線性查找
            Dim pEid = cmd.Parameters.Add("@eid", SqliteType.Blob)     ' 2026/06/10 by Gemini/Simon: 優化SQLite儲存空間把Entry_id改成BLOB二進位儲存
            Dim pFh = cmd.Parameters.Add("@fh", SqliteType.Integer)
            Dim pSubj = cmd.Parameters.Add("@subj", SqliteType.Text)
            Dim pSz = cmd.Parameters.Add("@sz", SqliteType.Integer)
            Dim pRt = cmd.Parameters.Add("@rt", SqliteType.Integer)   ' 2026/06/12 by Simon/Claude Opus 4.8: INTEGER Unix秒
            Dim pSn = cmd.Parameters.Add("@sn", SqliteType.Text)
            Dim pAc = cmd.Parameters.Add("@ac", SqliteType.Integer)
            Dim pPr = cmd.Parameters.Add("@pr", SqliteType.Integer)

            ' _cacheAttMailList: Dictionary(Of String, FolderCacheTab3)
            ' key = folder_path, value.AttMailList = List(Of MailItemInfo)
            For Each kvp In _cacheAttMailList
                Dim fp = kvp.Key
                If Not dirtyPaths.Contains(fp) Then Continue For   ' 2026/07/03 by Simon/Claude Fable 5: dirty 過濾
                Dim snap = kvp.Value.ItemCountSnap
                Dim mails = kvp.Value.AttMailList
                Dim fh As Long = FolderPathToHash64(fp)   ' 2026/07/03 by Simon/Claude: 提到迴圈外每夾算一次，原本每封信重算徒增 GC 壓力

                ' by Gemini 3 Flash, 2026/05/06: 實作「空結果持久化」，記住此資料夾已掃描且為 0 筆
                If mails.Count = 0 Then
                    ' 2026/06/10 by Gemini/Simon: 優化SQLite儲存空間把Entry_id改成BLOB二進位儲存
                    ' 2026/06/12 by Gemini/Simon: 把 folder_path TEXT 改成 folder_hash INTEGER 儲存
                    pEid.Value = HexStringToByteArray("EMPTY_ATT_" & fp)
                    pFh.Value = fh
                    pSubj.Value = ""
                    pSz.Value = 0
                    pRt.Value = 0L   ' 2026/06/12 by Simon/Claude Opus 4.8: sentinel 用 epoch 0
                    pSn.Value = ""
                    pAc.Value = 0
                    pPr.Value = snap
                    cmd.ExecuteNonQuery() : count += 1
                Else
                    ' 2026/06/12 by Simon/Claude Opus 4.8: 本機時間轉 Unix 秒儲存，讀回時 FromUnixTimeSeconds().LocalDateTime 還原
                    For Each mail In mails
                        pEid.Value = HexStringToByteArray(mail.EntryID)
                        pFh.Value = fh
                        pSubj.Value = If(mail.Subject, "")
                        pSz.Value = mail.Size
                        pRt.Value = LocalTimeToUnixSeconds(mail.RcvTime)
                        pSn.Value = If(mail.SenderName, "")
                        pAc.Value = mail.AttCount
                        pPr.Value = snap
                        cmd.ExecuteNonQuery() : count += 1
                    Next
                End If
            Next
        End Using
        Return count
        _dbg("結束")

    End Function
    Private Function SaveAttFilenameBatch() As Integer
        ' ---------------------------------------------------------------
        ' SaveAttFilenameBatch — Transaction 內批次寫入 att_filenames (Tab3 Phase2) 
        ' folder_path 透過反查 _cacheAttMailList 取得 (_cacheAttFilename key 是 EntryID) 
        ' 2026/04/09 修正: 移除 msg_size 欄位 (Phase2 永遠是 NULL，保留在 INSERT 造成
        '   SqliteType.Integer + DBNull.Value 不相容，丟 "Value must be set" InvalidOperationException)
        '   SQLite 未列出的欄位自動填 NULL，不需要明確傳入。
        ' ---------------------------------------------------------------
        _dbg("開始")
        If _dbMail Is Nothing Then Return 0

        Dim sql = "INSERT OR REPLACE INTO att_filenames" & " (entry_id,folder_hash,filenames)" & " VALUES (@eid,@fh,@fn)"
        Dim count As Integer = 0

        ' 反查 EntryID → folder_hash (從 Phase1 快取中建立對應表)
        ' 2026/07/03 by Simon/Claude: 原本存 folder_path、寫入迴圈逐列重算雜湊 → 改成建表時每夾算一次直接存雜湊
        Dim entryToHash As New Dictionary(Of String, Long)()
        For Each kvp In _cacheAttMailList
            Dim fh As Long = FolderPathToHash64(kvp.Key)
            For Each mail In kvp.Value.AttMailList
                If Not entryToHash.ContainsKey(mail.EntryID) Then entryToHash(mail.EntryID) = fh
            Next
        Next

        ' 2026/06/10 by Gemini/Simon: 優化SQLite儲存空間把Entry_id改成BLOB二進位儲存
        ' 2026/06/12 by Gemini/Simon: 把 folder_path TEXT 改成 folder_hash INTEGER 儲存，並新增一個記憶體字典 _dictHashToPath 來反查雜湊對應的路徑
        ' 2026/06/12 by Simon/Claude Opus 4.8: 移除 updated_at（只寫不讀）
        ' 2026/06/21 by Simon/Claude Opus 4.8: att_filenames 搬至 OLAcacheMail.db(_dbMail)，改自管獨立 _dbMail 交易(不再吃外部 _dbCache txn)
        ' 2026/07/03 by Simon/Claude: 改走專用寫入連線 — UI 的附件檔名 lazy read (LazyGetAttFilenames) 也走 _dbMail，
        '   共用連線時寫入交易會讓 UI 讀取在連線 mutex 上排隊 → 卡頓；WAL 一寫多讀，分開連線互不阻塞
        Try
            Using dbW As New SqliteConnection($"Data Source={_dbMailPath};Mode=ReadWrite")
                dbW.Open()
                Using pragmaCmd As New SqliteCommand("PRAGMA synchronous=NORMAL;", dbW) : pragmaCmd.ExecuteNonQuery() : End Using
                Using txnSim = dbW.BeginTransaction()
                    Using cmd As New SqliteCommand(sql, dbW, txnSim)
                        ' 2026/07/03 by Simon/Claude Fable 5: 參數物件存區域變數，免去迴圈內名稱線性查找
                        Dim pEid = cmd.Parameters.Add("@eid", SqliteType.Blob)
                        Dim pFh = cmd.Parameters.Add("@fh", SqliteType.Integer)
                        Dim pFn = cmd.Parameters.Add("@fn", SqliteType.Text)

                        For Each kvp In _cacheAttFilename
                            Dim fh As Long = 0 : entryToHash.TryGetValue(kvp.Key, fh)   ' 查無時 0，與原本 FolderPathToHash64("") 的結果一致
                            pEid.Value = HexStringToByteArray(kvp.Key)
                            pFh.Value = fh
                            pFn.Value = JsonSerializer.Serialize(kvp.Value)
                            cmd.ExecuteNonQuery() : count += 1
                        Next
                    End Using
                    txnSim.Commit()
                End Using
            End Using
        Catch ex As System.Exception
            _dbg("錯誤", ex.Message)   ' _dbMail 寫入失敗不連累主 db；下次掃描自動重建
        End Try
        Return count
        _dbg("結束")

    End Function

    Private Function SaveSendersBatch(txn As SqliteTransaction, dirtyPaths As HashSet(Of String)) As Integer
        ' ---------------------------------------------------------------
        ' SaveSendersBatch — 收集 _cacheMailInfo 中的唯一 email，
        '   批次 INSERT OR IGNORE 進 senders 表，再重建兩個記憶體字典
        ' 2026/06/12 by Simon/Claude Opus 4.8: 配合 sender_email 正規化架構新增
        '   呼叫時機：SaveMailInfoBatch 之前（同一 Transaction）
        '   完成後 _dictEmailToSenderId 可供 SaveMailInfoBatch 直接查 sender_id
        ' 2026/07/03 by Simon/Claude Fable 5: 加入 dirtyPaths 過濾，同 SaveMailInfoBatch —
        '   非 dirty 資料夾的 sender 早在該夾上次為 dirty 時就已寫入 senders 表，不必每次全量重掃 30 萬封信
        ' ---------------------------------------------------------------
        Dim allEmails As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        For Each kvp In _cacheMailInfo
            If Not dirtyPaths.Contains(kvp.Key) Then Continue For   ' 2026/07/03 by Simon/Claude Fable 5: dirty 過濾
            For Each Item In kvp.Value.Mails
                Dim email = Item.Mail.SenderEmail?.Trim()
                If Not String.IsNullOrEmpty(email) Then allEmails.Add(email.ToLower())
            Next
        Next

        If allEmails.Count = 0 Then Return 0

        Using cmd As New SqliteCommand("INSERT OR IGNORE INTO senders (sender_email) VALUES (@se)", txn.Connection, txn)   ' 2026/07/03 by Simon/Claude: 改跟隨 txn 所在連線
            Dim pSe = cmd.Parameters.Add("@se", SqliteType.Text)   ' 2026/07/03 by Simon/Claude Fable 5: 參數物件存區域變數
            For Each email In allEmails
                pSe.Value = email
                cmd.ExecuteNonQuery()
            Next
        End Using

        ' 重建記憶體字典（在同一 Transaction 內可讀到剛寫入的新 rows）
        LoadSendersBatch(txn)   ' 2026/07/03 by Simon/Claude: 必須傳 txn — 專用寫入連線上未 commit 的新 rows，走 _dbCache 讀不到
        Return allEmails.Count   ' 回傳新增 sender 數

    End Function
    Private Function SaveMailInfoBatch(txn As SqliteTransaction, dirtyPaths As HashSet(Of String)) As Integer
        ' ---------------------------------------------------------------
        ' SaveMailInfoBatch — Transaction 內批次寫入 mail_info (Tab4/Tab5)
        ' 2026/06/10 by Gemini/Simon: 優化SQLite儲存空間把Entry_id改成BLOB二進位儲存
        ' 2026/06/11 by Gemini/Simon: 把 message_id 轉成 xxHash64，並同時改成 BLOB 儲存節省空間
        ' 2026/06/12 by Gemini/Simon: 把 folder_path TEXT 改成 folder_hash INTEGER 儲存
        ' 2026/06/12 by Simon/Claude Opus 4.8: 移除 topic (動態計算)、sender_email (改 sender_id)、updated_at
        '   received_time TEXT→INTEGER (Unix秒)；SaveSendersBatch() 必須在此函式前呼叫
        ' 2026/07/03 by Simon/Claude Fable 5: 加入 dirtyPaths 過濾 — 這是四張表裡列數最多的一張(實測 30 萬列)，
        '   原本每次 SaveCache 不論異動與否都全量 INSERT OR REPLACE，耗時 2.4~2.7 秒；改成只重寫 dirty 資料夾後，沒有郵件異動的存檔應降到毫秒等級。
        ' ---------------------------------------------------------------
        _dbg("開始")
        Dim sql = "INSERT OR REPLACE INTO mail_info" &
                  " (entry_id,folder_hash,subject,msg_size,received_time,sender_name,sender_id,msgid_hash,pr_count_snap)" &
                  " VALUES (@eid,@fh,@subj,@sz,@rt,@sn,@sid,@mid,@pr)"

        Dim count As Integer = 0
        Using cmd As New SqliteCommand(sql, txn.Connection, txn)   ' 2026/07/03 by Simon/Claude: 改跟隨 txn 所在連線
            ' 2026/07/03 by Simon/Claude Fable 5: 參數物件存區域變數，免去迴圈內名稱線性查找 (350k 列 × 9 參數的熱迴圈)
            Dim pEid = cmd.Parameters.Add("@eid", SqliteType.Blob)
            Dim pFh = cmd.Parameters.Add("@fh", SqliteType.Integer)
            Dim pSubj = cmd.Parameters.Add("@subj", SqliteType.Text)
            Dim pSz = cmd.Parameters.Add("@sz", SqliteType.Integer)
            Dim pRt = cmd.Parameters.Add("@rt", SqliteType.Integer)   ' 2026/06/12 by Simon/Claude Opus 4.8: INTEGER Unix秒
            Dim pSn = cmd.Parameters.Add("@sn", SqliteType.Text)
            Dim pSid = cmd.Parameters.Add("@sid", SqliteType.Integer)  ' 2026/06/12 by Simon/Claude Opus 4.8: sender_id (NULL 若無 email)
            Dim pMid = cmd.Parameters.Add("@mid", SqliteType.Blob)
            Dim pPr = cmd.Parameters.Add("@pr", SqliteType.Integer)

            For Each kvp In _cacheMailInfo
                ' 2026/05/06 by Claude: key 已是純路徑，不再需 .Split
                Dim fp As String = kvp.Key
                If Not dirtyPaths.Contains(fp) Then Continue For   ' 2026/07/03 by Simon/Claude Fable 5: dirty 過濾 — 最熱的一張表，效果最大
                Dim snap = kvp.Value.Snap
                Dim mails = kvp.Value.Mails
                Dim fh As Long = FolderPathToHash64(fp)   ' 2026/07/03 by Simon/Claude: 提到迴圈外每夾算一次，原本每封信重算徒增 GC 壓力

                If mails.Count = 0 Then
                    ' by Gemini 3 Flash, 2026/05/06: 實作「空結果持久化」，記住此資料夾已掃描且為 0 筆
                    pEid.Value = HexStringToByteArray("EMPTY_BASIC_" & fp)
                    pFh.Value = fh
                    pSubj.Value = ""
                    pSz.Value = 0
                    pRt.Value = 0L
                    pSn.Value = ""
                    pSid.Value = DBNull.Value
                    pMid.Value = DBNull.Value
                    pPr.Value = snap
                    cmd.ExecuteNonQuery() : count += 1
                Else
                    For Each Item In mails
                        ' 2026/06/12 by Simon/Claude Opus 4.8: 查 _dictEmailToSenderId，無 email 時存 NULL
                        Dim emailKey = Item.Mail.SenderEmail?.Trim()?.ToLower()
                        Dim sid As Object
                        Dim foundId As Integer
                        If Not String.IsNullOrEmpty(emailKey) AndAlso _dictEmailToSenderId.TryGetValue(emailKey, foundId) Then
                            sid = foundId
                        Else
                            sid = DBNull.Value
                        End If

                        pEid.Value = HexStringToByteArray(Item.Mail.EntryID)
                        pFh.Value = fh
                        pSubj.Value = If(Item.Mail.Subject, "")
                        pSz.Value = Item.Mail.Size
                        pRt.Value = LocalTimeToUnixSeconds(Item.Mail.RcvTime) ' 2026/06/12 by Simon/Claude Opus 4.8: 本機時間轉 Unix 秒
                        pSn.Value = If(Item.Mail.SenderName, "")
                        pSid.Value = sid
                        pMid.Value = HexStringToByteArray(If(Item.Mail.MsgIDhash, ""))
                        pPr.Value = snap
                        cmd.ExecuteNonQuery() : count += 1
                    Next
                End If
            Next
        End Using
        Return count
        _dbg("結束")
    End Function

    Private Function SaveFolderInfoBatch(txn As SqliteTransaction) As Integer
        ' ---------------------------------------------------------------
        ' SaveFolderInfoBatch — Transaction 內批次寫入 folder_info
        ' 注意: 在 Task.Run 背景執行緒呼叫，不可碰 UI 控制項
        ' pr_count_snap = _cacheMailCount[path]，即 PR_CONTENT_COUNT 讀取結果
        ' ---------------------------------------------------------------
        _dbg("    ├ 開始") ' by Gemini, 2026/04/10: 調整縮排層級為 Level 2
        ' 2026/06/14 by Simon/Claude Opus 4.8: INSERT OR REPLACE → UPSERT。
        '   原 INSERT OR REPLACE 是「整列覆寫」: 當 path 在 allPaths 但 _cacheFolderIDs 查無時(line 941 Else 分支)，
        '   entry_id/store_id/is_mail/has_chinese 會被寫成 NULL，把 DB 既有的身分證洗掉。
        '   而樹載入 LazyGetOrderedSubFolderIDs 帶 "entry_id IS NOT NULL" → 這些資料夾從樹消失
        '   (重現: 第 2 次 do-nothing 關閉存檔 → 第 3 次啟動 Gmail_2022 底下只剩收件匣)。
        '   且會自我延續: eid 被洗→下次 lazy-load 因 .eid 空只補 _cacheFolderCount(進 allPaths) 不補 _cacheFolderIDs→再被洗。
        '   修法: 改 ON CONFLICT DO UPDATE — 統計欄照常覆寫(與原行為一致)；身分欄以 COALESCE(新值, 舊值) 保留，
        '         新值為 NULL(快取沒身分證)時不動 DB 既有值 → entry_id 永不被洗掉，並打斷上述自我延續循環。
        Dim sql = "INSERT INTO folder_info" &
                  " (folder_path,mail_count,mail_count_all,folder_count,folder_count_all,folder_size,folder_size_all,pr_count_snap,commit_max,entry_id,store_id,is_mail,has_chinese,updated_at) " &
                  "VALUES (@fp,@mc,@mca,@fc,@fca,@fs,@fsa,@pr,@cmx,@eid,@sid,@ism,@hasch,@ts) " &
                  "ON CONFLICT(folder_path) DO UPDATE SET " &
                  " mail_count=excluded.mail_count, mail_count_all=excluded.mail_count_all," &
                  " folder_count=excluded.folder_count, folder_count_all=excluded.folder_count_all," &
                  " folder_size=excluded.folder_size, folder_size_all=excluded.folder_size_all," &
                  " pr_count_snap=excluded.pr_count_snap," &
                  " commit_max =COALESCE(excluded.commit_max,  commit_max)," &
                  " entry_id   =COALESCE(excluded.entry_id,    entry_id)," &
                  " store_id   =COALESCE(excluded.store_id,    store_id)," &
                  " is_mail    =COALESCE(excluded.is_mail,     is_mail)," &
                  " has_chinese=COALESCE(excluded.has_chinese, has_chinese)," &
                  " updated_at =excluded.updated_at"

        ' 蒐集六個 dict 中所有出現過的 folder_path 聯集
        Dim allPaths As New HashSet(Of String)()
        For Each k In _cacheMailCount.Keys : allPaths.Add(k) : Next
        For Each k In _cacheMailCountAll.Keys : allPaths.Add(k) : Next
        For Each k In _cacheFolderCount.Keys : allPaths.Add(k) : Next
        For Each k In _cacheFolderCountAll.Keys : allPaths.Add(k) : Next
        For Each k In _cacheFolderSize.Keys : allPaths.Add(k) : Next
        For Each k In _cacheFolderSizeAll.Keys : allPaths.Add(k) : Next
        For Each k In _cacheFolderIDs.Keys : allPaths.Add(k) : Next ' by Gemini 3.0 flash, 2026/04/18: 額外聯集身分標識字典的 Key，確保僅掃描過但未統計的資料夾也能存入 SSD
        For Each k In _cacheFolderCommitMax.Keys : allPaths.Add(k) : Next ' 2026/07/04 by Simon/Claude Fable 5: commit 基準也要落 DB

        ' 2026/06/12 by Simon/Claude Opus 4.8: folder_info 作為主表，確保所有引用 folder_hash 的子表路徑都被收錄
        ' 重啟後 LoadFolderInfoBatch 還原完整 _dictHashToPath，防止 LoadMailInfoBatch 等批次載入 skip 資料
        For Each k In _cacheMailInfo.Keys : allPaths.Add(k) : Next
        For Each k In _cacheAttMailList.Keys : allPaths.Add(k) : Next
        For Each k In _cacheYearCount.Keys : allPaths.Add(k) : Next
        ' _cacheMonthCount key 格式為 "FolderPath_year"，需解析後加入
        For Each k In _cacheMonthCount.Keys
            Dim lastUs = k.LastIndexOf("_"c)
            If lastUs > 0 Then allPaths.Add(k.Substring(0, lastUs))
        Next

        Dim count As Integer = 0
        Dim ts = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
        Using cmd As New SqliteCommand(sql, txn.Connection, txn)   ' 2026/07/03 by Simon/Claude: 改跟隨 txn 所在連線(SaveCachesToDB 的專用寫入連線)，不寫死 _dbCache
            ' 2026/07/03 by Simon/Claude Fable 5: 參數物件存區域變數，免去迴圈內 cmd.Parameters("@x") 逐次的名稱線性查找
            Dim pFp = cmd.Parameters.Add("@fp", SqliteType.Text)
            Dim pMc = cmd.Parameters.Add("@mc", SqliteType.Integer)
            Dim pMca = cmd.Parameters.Add("@mca", SqliteType.Integer)
            Dim pFc = cmd.Parameters.Add("@fc", SqliteType.Integer)
            Dim pFca = cmd.Parameters.Add("@fca", SqliteType.Integer)
            Dim pFs = cmd.Parameters.Add("@fs", SqliteType.Integer)
            Dim pFsa = cmd.Parameters.Add("@fsa", SqliteType.Integer)
            Dim pPr = cmd.Parameters.Add("@pr", SqliteType.Integer)
            Dim pCmx = cmd.Parameters.Add("@cmx", SqliteType.Integer)  ' 2026/07/04 by Simon/Claude Fable 5: commit_max，僅 RenewCache 掃描過的資料夾有值，其餘 NULL 由 COALESCE 保留 DB 基準
            Dim pEid = cmd.Parameters.Add("@eid", SqliteType.Blob)     ' 2026/06/10 by Gemini/Simon: 優化SQLite儲存空間把Entry_id改成BLOB二進位儲存
            Dim pSid = cmd.Parameters.Add("@sid", SqliteType.Text)
            Dim pIsm = cmd.Parameters.Add("@ism", SqliteType.Integer)
            Dim pHasch = cmd.Parameters.Add("@hasch", SqliteType.Integer)
            Dim pTs = cmd.Parameters.Add("@ts", SqliteType.Text)

            For Each path In allPaths
                ' 2026/04/07 修正 v2: 初始值設 -1 仍然不夠，因為 -1 是整數值會被寫入 DB，
                '   LoadFolderInfoBatch 讀回 -1 後直接塞入記憶體快取，
                '   GetFolderCount 命中記憶體回傳 -1 → LoadSubFolderToTreeView 判斷 -1 > 0 為 False → 不顯示 "+"。
                '   正確做法：沒有測量過的欄位一律寫 DBNull.Value (SQL NULL)，這樣 LoadFolderInfoBatch 的 IsDBNull 保護才能正確跳過，不污染記憶體快取。
                Dim mc, mca, fc, fca As Integer : Dim fs, fsa As Long
                Dim hasMc = _cacheMailCount.TryGetValue(path, mc)
                Dim hasMca = _cacheMailCountAll.TryGetValue(path, mca)
                Dim hasFc = _cacheFolderCount.TryGetValue(path, fc)
                Dim hasFca = _cacheFolderCountAll.TryGetValue(path, fca)
                Dim hasFs = _cacheFolderSize.TryGetValue(path, fs)
                Dim hasFsa = _cacheFolderSizeAll.TryGetValue(path, fsa)
                pFp.Value = path
                pMc.Value = If(hasMc, CObj(mc), DBNull.Value)
                pMca.Value = If(hasMca, CObj(mca), DBNull.Value)
                pFc.Value = If(hasFc, CObj(fc), DBNull.Value)
                pFca.Value = If(hasFca, CObj(fca), DBNull.Value)
                pFs.Value = If(hasFs, CObj(fs), DBNull.Value)
                pFsa.Value = If(hasFsa, CObj(fsa), DBNull.Value)
                pPr.Value = If(hasMc, CObj(mc), DBNull.Value)
                Dim cmx As Long
                pCmx.Value = If(_cacheFolderCommitMax.TryGetValue(path, cmx), CObj(cmx), DBNull.Value)   ' 2026/07/04 by Simon/Claude Fable 5

                ' by Gemini, 2026/04/10: 寫入身分標識與標籤 (從 _cacheFolderIDs 提取)
                Dim idInfo As (eid As String, sid As String, isMail As Boolean, hasCh As Boolean) = Nothing
                If _cacheFolderIDs.TryGetValue(path, idInfo) Then
                    pEid.Value = HexStringToByteArray(idInfo.eid) ' 2026/06/10 by Gemini/Simon: 優化SQLite儲存空間把Entry_id改成BLOB二進位儲存
                    pSid.Value = idInfo.sid
                    pIsm.Value = If(idInfo.isMail, 1, 0)
                    pHasch.Value = If(idInfo.hasCh, 1, 0)
                Else
                    pEid.Value = DBNull.Value
                    pSid.Value = DBNull.Value
                    pIsm.Value = DBNull.Value
                    pHasch.Value = DBNull.Value
                End If
                pTs.Value = ts
                cmd.ExecuteNonQuery() : count += 1
            Next
        End Using

        _dbg("    ├ 結束") ' by Gemini, 2026/04/11: 修正與開始對齊
        Return count

    End Function
#End Region

#Region "■ 批次載入核心 (Batch Reader Core)"
    Private Function LoadDateCountCore(onlyYearCount As Boolean) As Integer
        ' ---------------------------------------------------------------
        ' LoadDateCountCore — 全表重建 _cacheYearCount / _cacheMonthCount 共用核心
        '   onlyYearCount = True    : year_count   → key=folder_path, subKey=year
        '   onlyYearCount = False   : month_count  → key=folder_path_year (cacheKey), subKey=month
        ' 先按 key 分組收集，最後一次性 TryAdd (保留記憶體已有版本)
        '   month_count 新增函數群 (2026/04/09 by Claude)
        ' 2026/04/09 修正：改用三欄 PK，從 folder_path + year 重組 cacheKey
        ' 2026/07/08 by Simon/Claude: 消重 — 原 LoadYearCountBatch / LoadMonthCountBatch
        '   兩函式 85% 同構，合併為單一 core 用 onlyYearCount 切換；順手移除 year 版 Return 之後永不執行的 _dbg(" ├ 結束") 死碼。
        ' 2026/07/09 by Simon/Claude: 原名 LoadYearMonthCountBatch，改名 LoadDateCountCore 並拆回 LoadYearCountBatch() / LoadMonthCountBatch() 兩個無參數薄 wrapper，呼叫端維持原本具名語意。
        ' ---------------------------------------------------------------
        _dbg(" ├ 開始", If(onlyYearCount, "year_count", "month_count"))
        Dim sql = If(onlyYearCount, "SELECT folder_hash,year,count FROM year_count",
                                    "SELECT folder_hash,year,month,count FROM month_count")
        Dim target = If(onlyYearCount, _cacheYearCount, _cacheMonthCount)
        Dim count As Integer = 0
        Dim tempDict As New Dictionary(Of String, ConcurrentDictionary(Of Integer, Integer))()

        Using cmd As New SqliteCommand(sql, _dbCache)
            Using reader = cmd.ExecuteReader()
                While reader.Read()
                    Dim fp As String = "" : If Not _dictHashToPath.TryGetValue(reader.GetInt64(0), fp) Then Continue While
                    Dim key = If(onlyYearCount, fp, fp & "_" & reader.GetInt32(1).ToString())   ' month 版重組 cacheKey，與 _cacheMonthCount key 格式一致
                    Dim subKey = reader.GetInt32(If(onlyYearCount, 1, 2))
                    Dim cnt = reader.GetInt32(If(onlyYearCount, 2, 3))

                    Dim value As ConcurrentDictionary(Of Integer, Integer) = Nothing
                    If Not tempDict.TryGetValue(key, value) Then
                        value = New ConcurrentDictionary(Of Integer, Integer)()
                        tempDict(key) = value
                    End If
                    value(subKey) = cnt : count += 1
                End While
            End Using
        End Using

        For Each kvp In tempDict : target.TryAdd(kvp.Key, kvp.Value) : Next
        _dbg(" ├ 結束", $"{count} 筆 → {tempDict.Count} 個 key")
        Return count

    End Function
    Private Function LoadYearCountBatch() As Integer
        ' 薄 wrapper — year_count 全表重建，委派 LoadDateCountCore (2026/07/09 by Simon/Claude)
        Return LoadDateCountCore(onlyYearCount:=True)
    End Function
    Private Function LoadMonthCountBatch() As Integer
        ' 薄 wrapper — month_count 全表重建，委派 LoadDateCountCore (2026/07/09 by Simon/Claude)
        Return LoadDateCountCore(onlyYearCount:=False)
    End Function

    Private Function LoadAttMailListCore(folderPaths As List(Of String)) As Dictionary(Of String, AttMailListDbResult)
        ' ---------------------------------------------------------------
        ' LoadAttMailListCore — att_maillist 共用查詢引擎。folderPaths Is Nothing → 全表(不加 WHERE)；
        '   否則 WHERE folder_hash IN (...)。EMPTY_ATT_ 哨兵只記 Snap、不進 Mails，兩個 LazyGetAttMailList
        '   overload(單夾/全表)與 LoadAttMailListBatch 共用同一份 reader，不再各自維護。
        ' 2026/07/07 by Simon/Claude: 抽出 Core — 順帶修正 LoadAttMailListBatch 原本漏判 EMPTY_ATT_ 哨兵的不一致
        '   (LazyGetAttMailList 單夾版原本有跳過，全表版沒有，兩者不同源導致的分歧)
        ' ---------------------------------------------------------------
        Dim result As New Dictionary(Of String, AttMailListDbResult)()
        Dim isFullTable As Boolean = folderPaths Is Nothing
        If _dbCache Is Nothing OrElse (Not isFullTable AndAlso folderPaths.Count = 0) Then Return result
        If _iLikeNoisy Then _dbg(" ├ 開始", If(isFullTable, "全表載入", $"批次查詢 {folderPaths.Count} 個路徑"))

        Try
            ' 2026/06/12 by Gemini/Simon: 把 folder_path TEXT 改成 folder_hash INTEGER 儲存，並新增一個記憶體字典 _dictHashToPath 來反查雜湊對應的路徑
            Dim sql As String = "SELECT entry_id,folder_hash,subject,msg_size,received_time,sender_name,att_count,pr_count_snap FROM att_maillist"

            ' SQLite 預設 variable 上限 999，300 個路徑絕對安全 (同 LoadMailInfoCore 的作法)
            Dim hashes As List(Of Long) = Nothing
            If Not isFullTable Then
                hashes = New List(Of Long)(folderPaths.Count)
                For Each p In folderPaths : hashes.Add(FolderPathToHash64(p)) : Next
                sql &= " WHERE folder_hash IN (" & String.Join(",", Enumerable.Range(0, hashes.Count).Select(Function(i) "@fh" & i.ToString())) & ")"
            End If

            Using cmd As New SqliteCommand(sql, _dbCache)
                If Not isFullTable Then
                    For i As Integer = 0 To hashes.Count - 1
                        cmd.Parameters.AddWithValue("@fh" & i.ToString(), hashes(i))
                    Next
                End If

                ' 2026/06/10 by Gemini/Simon: 優化SQLite儲存空間把Entry_id改成BLOB二進位儲存
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        Dim fp As String = ""
                        If Not _dictHashToPath.TryGetValue(reader.GetInt64(1), fp) Then Continue While

                        ' 首次見到此 folder_path → 建立 entry（含 sentinel row 也要建，記錄 snap）
                        If Not result.ContainsKey(fp) Then result(fp) = New AttMailListDbResult()
                        result(fp).Snap = If(reader.IsDBNull(7), -1L, reader.GetInt64(7)) ' pr_count_snap 整個 folder 共用同一值，每行都一樣，讀最後一次即可

                        Dim eid = ByteArrayToHexString(reader.GetFieldValue(Of Byte())(0))
                        If eid.StartsWith("EMPTY_ATT_") Then Continue While   ' sentinel row，只記 snap

                        Dim mail As New MailItemInfo With {.EntryID = eid,
                                                           .Subject = If(reader.IsDBNull(2), "", reader.GetString(2)),
                                                           .Size = If(reader.IsDBNull(3), 0L, reader.GetInt64(3)),
                                                           .RcvTime = UnixSecondsToLocalTime(If(reader.IsDBNull(4), 0L, reader.GetInt64(4))),
                                                           .SenderName = If(reader.IsDBNull(5), "", reader.GetString(5)),
                                                           .AttCount = If(reader.IsDBNull(6), 0, reader.GetInt32(6)),
                                                           .FolderPath = fp}
                        result(fp).Mails.Add(mail)
                    End While
                End Using
            End Using
        Catch ex As System.Exception
            _dbg(" ├ 錯誤", ex.Message)
        End Try

        If _iLikeNoisy Then _dbg(" ├ 結束", $"取回 {result.Count} 個資料夾資料")
        Return result
    End Function
    Private Function LoadAttMailListBatch() As Integer
        ' ---------------------------------------------------------------
        ' LoadAttMailListBatch — 重建 _cacheAttMailList (按 folder_path 分組)
        ' 2026/07/07 by Simon/Claude: 消重 — 原 28 行全表 reader 與 LoadAttMailListCore 同構，改委派無參數
        '   全表 overload；本函式只保留「灌進 _cacheAttMailList(TryAdd 記憶體優先) + 回傳筆數統計」的流程身分。
        '   委派後一併修正原本漏判 EMPTY_ATT_ 哨兵的問題(哨兵列不再被當成真信塞入 _cacheAttMailList)。
        ' ---------------------------------------------------------------
        _dbg("開始")
        Dim all = LazyGetAttMailList()   ' 無參數 overload = 全表模式；DB 未初始化/空表時回空 dict
        For Each kvp In all
            _cacheAttMailList.TryAdd(kvp.Key, New FolderCacheTab3 With {.AttMailList = kvp.Value.Mails, .ItemCountSnap = kvp.Value.Snap})
        Next
        Return all.Values.Sum(Function(v) v.Mails.Count)   ' 與舊版 count 同義：只計真信，哨兵列不計
    End Function
    Private Function LoadAttFilenamesBatch() As Integer
        ' ---------------------------------------------------------------
        ' LoadAttFilenamesBatch — 重建 _cacheAttFilename (JSON 反序列化)
        ' ---------------------------------------------------------------
        _dbg("開始")
        Dim count As Integer = 0
        If _dbMail Is Nothing Then Return 0   ' 2026/06/21 by Simon/Claude Opus 4.8: att_filenames 來源改 _dbMail(OLAcacheMail.db)
        Using cmd As New SqliteCommand("SELECT entry_id,filenames FROM att_filenames", _dbMail)
            Using reader = cmd.ExecuteReader()
                ' 2026/06/10 by Gemini/Simon: 優化SQLite儲存空間把Entry_id改成BLOB二進位儲存
                While reader.Read()
                    Dim eidStr = ByteArrayToHexString(reader.GetFieldValue(Of Byte())(0))
                    Dim fnJson = If(reader.IsDBNull(1), "[]", reader.GetString(1))
                    Try
                        Dim list = JsonSerializer.Deserialize(Of List(Of String))(fnJson)
                        _cacheAttFilename.TryAdd(eidStr, list)
                        count += 1   ' 2026/06/23 by Claude Sonnet 4.6: 修正計數遺漏 (原本 count 永遠回傳 0，導致 LoadCache 統計顯示「讀入 0 筆」)
                    Catch ex As System.Exception
                        _dbg("錯誤: 解析失敗", $"{eidStr}: {ex.Message}")
                    End Try
                End While
            End Using
        End Using
        Return count
        _dbg("結束")

    End Function

    Private Function LoadSendersBatch(Optional txn As SqliteTransaction = Nothing) As Integer
        ' ---------------------------------------------------------------
        ' LoadSendersBatch — 載入 senders 表，重建 _dictEmailToSenderId / _dictSenderIdToEmail
        ' 2026/06/12 by Simon/Claude Opus 4.8: 配合 sender_email 正規化架構新增
        '   - 寫入側 (_dictEmailToSenderId): SaveMailInfoBatch 查詢 sender_id
        '   - 讀取側 (_dictSenderIdToEmail): Load/DbGet 函式還原 SenderEmail
        ' 2026/07/03 by Simon/Claude: 新增 Optional txn — SaveCachesToDB 的寫入交易改開在專用連線上後，
        '   交易內剛 INSERT 的 senders 從 _dbCache (另一條連線) 讀不到（未 commit），
        '   必須在同一交易內讀，否則 SaveMailInfoBatch 查 sender_id 全部落空寫成 NULL
        ' ---------------------------------------------------------------
        _dictEmailToSenderId.Clear()
        _dictSenderIdToEmail.Clear()
        Dim conn As SqliteConnection = If(txn IsNot Nothing, txn.Connection, _dbCache)
        If conn Is Nothing Then Return 0

        Try
            Using cmd As New SqliteCommand("SELECT sender_id, sender_email FROM senders", conn, txn)
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        Dim id = reader.GetInt32(0)
                        Dim email = reader.GetString(1)   ' DB 中已是 lowercase
                        _dictEmailToSenderId(email) = id
                        _dictSenderIdToEmail(id) = email
                    End While
                End Using
            End Using
        Catch ex As System.Exception
            _dbg(" ├ DbLoadSendersBatch 錯誤", ex.Message)
        End Try

        ' count 是已在 While 內計數的整數，最後加：
        Return _dictSenderIdToEmail.Count

    End Function
    Private Function LoadMailInfoCore(folderPaths As List(Of String)) As Dictionary(Of String, (Mails As List(Of (Mail As MailItemInfo, Topic As String)), Snap As Long))
        ' ---------------------------------------------------------------
        ' LoadMailInfoCore(List) — 一次 SQL IN 查詢，批次讀回多個資料夾的 mail_info
        ' 取代原來 300 個別查詢的瓶頸；由 PreLoadMailCacheAsync 呼叫。
        ' 回傳 key=folder_path, value=(Mails, Snap)；不做 snap 驗證，由呼叫端決定。
        ' 2026/05/11 by Simon/Claude: 優化B
        ' 2026/06/12 by Simon/Claude Opus 4.8: topic 改由 GetCleanSubject(subject) 動態計算
        ' 2026/07/07 by Simon/Claude: 原名 LazyGetMailInfoBatch，改 Overloads 同名；
        '   folderPaths = Nothing 為全表模式(內部細節，外部一律走無參數 overload)，讓 LoadMailInfoBatch 的全表 reader 併入本函式
        ' ---------------------------------------------------------------
        Dim isFullTable As Boolean = folderPaths Is Nothing
        Dim result As New Dictionary(Of String, (Mails As List(Of (Mail As MailItemInfo, Topic As String)), Snap As Long))(If(isFullTable, 512, folderPaths.Count))
        If _dbCache Is Nothing OrElse (Not isFullTable AndAlso folderPaths.Count = 0) Then Return result
        If _iLikeNoisy Then _dbg(" ├ 開始", If(isFullTable, "全表載入", $"批次查詢 {folderPaths.Count} 個路徑"))

        Try
            ' 2026/06/12 by Simon/Claude Opus 4.8: 更新 SELECT 欄位：移除 topic/sender_email/updated_at，加 sender_id；received_time 改 INTEGER
            ' 新欄位順序 — entry_id(0),folder_hash(1),subject(2),msg_size(3),received_time INTEGER(4),
            '             sender_name(5),sender_id(6),msgid_hash(7),pr_count_snap(8)
            Dim sql As String = "SELECT entry_id,folder_hash,subject,msg_size,received_time,sender_name,sender_id,msgid_hash,pr_count_snap" &
                                "  FROM mail_info"

            ' 建立 Hash 清單與對照表, 2026/06/12 (FolderPathToHash64 同時註冊 _dictHashToPath 反查)
            ' SQLite 預設 variable 上限 999，300 個路徑絕對安全
            Dim hashes As List(Of Long) = Nothing
            If Not isFullTable Then
                hashes = New List(Of Long)(folderPaths.Count)
                For Each p In folderPaths : hashes.Add(FolderPathToHash64(p)) : Next
                sql &= " WHERE folder_hash IN (" & String.Join(",", Enumerable.Range(0, hashes.Count).Select(Function(i) "@fh" & i.ToString())) & ")"
            End If

            Using cmd As New SqliteCommand(sql, _dbCache)
                If Not isFullTable Then
                    For i As Integer = 0 To hashes.Count - 1
                        cmd.Parameters.AddWithValue("@fh" & i.ToString(), hashes(i))
                    Next
                End If

                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        Dim fp As String = ""
                        If Not _dictHashToPath.TryGetValue(reader.GetInt64(1), fp) Then Continue While

                        ' 首次見到此 folder_path → 建立 entry（含 sentinel row 也要建，記錄 snap）
                        Dim snap As Long = If(reader.IsDBNull(8), -1L, reader.GetInt64(8))
                        If Not result.ContainsKey(fp) Then result(fp) = (New List(Of (Mail As MailItemInfo, Topic As String))(256), snap)

                        ' 2026/06/10 by Gemini/Simon: 優化SQLite儲存空間把Entry_id改成BLOB二進位儲存
                        ' 2026/06/11 by Gemini/Simon: 把 message_id 轉成 xxHash64，並同時改成 BLOB 儲存節省空間
                        ' 2026/06/12 by Simon/Claude Opus 4.8: received_time 改為 INTEGER (Unix秒)
                        ' 2026/06/12 by Simon/Claude Opus 4.8: sender_id → email 反查
                        Dim eid As String = ByteArrayToHexString(reader.GetFieldValue(Of Byte())(0))
                        If eid.StartsWith("EMPTY_BASIC_") Then Continue While   ' sentinel row，只記 snap

                        Dim mail As New MailItemInfo With {.EntryID = eid,
                                                           .Subject = If(reader.IsDBNull(2), "", reader.GetString(2)),
                                                           .Size = reader.GetInt64(3),
                                                           .RcvTime = UnixSecondsToLocalTime(If(reader.IsDBNull(4), 0L, reader.GetInt64(4))),
                                                           .SenderName = If(reader.IsDBNull(5), "", reader.GetString(5)),
                                                           .SenderEmail = _dictSenderIdToEmail.GetValueOrDefault(If(reader.IsDBNull(6), 0, reader.GetInt32(6)), ""),
                                                           .FolderPath = fp,
                                                           .MsgIDhash = If(reader.IsDBNull(7), "", ByteArrayToHexString(reader.GetFieldValue(Of Byte())(7)))}
                        result(fp).Mails.Add((mail, GetCleanSubject(mail.Subject)))
                    End While
                End Using
            End Using
        Catch ex As System.Exception
            _dbg(" ├ 錯誤", ex.Message)
        End Try

        If _iLikeNoisy Then _dbg(" ├ 結束", $"取回 {result.Count} 個資料夾資料")
        Return result
    End Function
    Private Function LoadMailInfoBatch() As Integer
        ' ---------------------------------------------------------------
        ' LoadMailInfoBatch — 重建 _cacheMailInfo (Tab4/5 專用)
        ' 2026/04/22 by Gemini 3.1 Pro: 補齊載入邏輯，解決重啟後重複掃描問題
        ' 2026/07/07 by Simon/Claude: 消重 — 原 45 行全表 reader 與 LoadMailInfoCore(List) 同構
        '   (欄位/EMPTY_BASIC_ 哨兵/sender 反查/按夾分組)，改委派無參數全表 overload；
        '   本函式只保留「灌進 _cacheMailInfo(TryAdd 記憶體優先) + 回傳筆數統計」的流程身分。
        '   順帶修正舊版 tempDict Snap 為 Integer、_cacheMailInfo 實為 Long 的型別不一致。
        ' ---------------------------------------------------------------
        _dbg("開始")
        Dim all = LazyGetMailInfo()                     ' 無參數 overload = 全表模式；DB 未初始化/空表時回空 dict
        For Each kvp In all
            _cacheMailInfo.TryAdd(kvp.Key, kvp.Value)   ' TryAdd：記憶體已有值(Layer2.5 已讀過)時保留記憶體版本
        Next
        Return all.Values.Sum(Function(v) v.Mails.Count) ' 與舊版 count 同義：只計真信，哨兵列不計
    End Function

    Private Function LoadFolderInfoCore(folderPaths As List(Of String)) As Dictionary(Of String, FolderInfoDbRow)
        ' ---------------------------------------------------------------
        ' LoadFolderInfoCore — folder_info 共用查詢引擎。
        '   folderPaths Is Nothing → 全表；否則 WHERE folder_path IN (...)。
        '   folder_info 的 PK 本身就是 folder_path TEXT (不像子表要靠 folder_hash 反查)，過濾直接用路徑字串當參數即可。
        '   SELECT 統一補齊成完整 13 欄，取代原本三份各自缺欄的查詢：
        '     LazyGetFolderInfo(fPath)   原缺 path/commit_max
        '     LoadAllFolderInfo()       原缺 has_chinese
        '     LoadFolderInfoBatch()    原缺 pr_count_snap/commit_max
        '   每列統一呼叫 FolderPathToHash64 註冊 _dictHashToPath(TryAdd 冪等)，取代原本只有 LoadFolderInfoBatch 一處負責這件事。
        ' 2026/07/07 by Simon/Claude
        ' ---------------------------------------------------------------
        Dim result As New Dictionary(Of String, FolderInfoDbRow)()
        Dim isFullTable As Boolean = folderPaths Is Nothing
        If _dbCache Is Nothing OrElse (Not isFullTable AndAlso folderPaths.Count = 0) Then Return result
        If _iLikeNoisy Then _dbg(" ├ 開始", If(isFullTable, "全表載入", $"批次查詢 {folderPaths.Count} 個路徑"))

        Try
            Dim sql As String = "SELECT folder_path,mail_count,mail_count_all,folder_count,folder_count_all," &
                                 "folder_size,folder_size_all,pr_count_snap,commit_max,entry_id,store_id,is_mail,has_chinese FROM folder_info"
            Dim paramNames As List(Of String) = Nothing
            If Not isFullTable Then
                paramNames = Enumerable.Range(0, folderPaths.Count).Select(Function(i) "@fp" & i.ToString()).ToList()
                sql &= " WHERE folder_path IN (" & String.Join(",", paramNames) & ")"
            End If

            Using cmd As New SqliteCommand(sql, _dbCache)
                If Not isFullTable Then
                    For i As Integer = 0 To folderPaths.Count - 1
                        cmd.Parameters.AddWithValue(paramNames(i), folderPaths(i))
                    Next
                End If

                ' 2026/06/10 by Gemini/Simon: 優化SQLite儲存空間把Entry_id改成BLOB二進位儲存
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        Dim path = reader.GetString(0)
                        FolderPathToHash64(path)   ' 註冊到字典，供子表反查(原 LoadFolderInfoBatch 行為)
                        result(path) = New FolderInfoDbRow With {.path = path,
                                                                 .mc = If(reader.IsDBNull(1), -1L, reader.GetInt64(1)),
                                                                 .mca = If(reader.IsDBNull(2), -1L, reader.GetInt64(2)),
                                                                 .fc = If(reader.IsDBNull(3), -1L, reader.GetInt64(3)),
                                                                 .fca = If(reader.IsDBNull(4), -1L, reader.GetInt64(4)),
                                                                 .fs = If(reader.IsDBNull(5), -1L, reader.GetInt64(5)),
                                                                 .fsa = If(reader.IsDBNull(6), -1L, reader.GetInt64(6)),
                                                                 .snap = If(reader.IsDBNull(7), -1L, reader.GetInt64(7)),
                                                                 .cmx = If(reader.IsDBNull(8), -1L, reader.GetInt64(8)),
                                                                 .eid = If(reader.IsDBNull(9), "", ByteArrayToHexString(reader.GetFieldValue(Of Byte())(9))),
                                                                 .sid = If(reader.IsDBNull(10), "", reader.GetString(10)),
                                                                 .isMail = If(reader.IsDBNull(11), -1, reader.GetInt32(11)),
                                                                 .hasCh = If(reader.IsDBNull(12), -1, reader.GetInt32(12))}
                    End While
                End Using
            End Using
        Catch ex As System.Exception
            _dbg(" ├ 錯誤", ex.Message)
        End Try

        If _iLikeNoisy Then _dbg(" ├ 結束", $"取回 {result.Count} 個資料夾資料")
        Return result
    End Function
    Private Function LoadFolderInfoBatch() As Integer
        ' ---------------------------------------------------------------
        ' LoadFolderInfoBatch — 讀回六個數字快取
        ' 使用 TryAdd：記憶體已有值時保留記憶體版本 (不覆蓋 Layer2.5 剛讀進來的較新值)
        ' 2026/04/07 修正: 每個欄位加 IsDBNull 保護，NULL 代表「從未測量過」，
        '   不可塞入記憶體快取，否則 GetFolderCount 命中 -1 → LoadSubFolderToTreeView
        '   判斷 -1 > 0 為 False → 不顯示 TreeView "+" 加號 (bug) 。
        ' 2026/07/07 by Simon/Claude: 消重 — 全表 reader 改委派 LazyGetFolderInfo() (→ LoadFolderInfoCore)，
        '   FolderPathToHash64 註冊、-1=NULL 哨兵改由 Core 統一處理；本函式只保留「灌進 6 個 _cache*
        '   字典與 _cacheFolderIDs(TryAdd 記憶體優先) + 回傳筆數統計」的流程身分。row.mc>=0 等同原本的
        '   IsDBNull 判斷(FolderInfoDbRow 本身約定 -1 = DB 中 NULL/未測量，見類別定義註解)。
        ' ---------------------------------------------------------------
        _dbg("開始")
        Dim all = LazyGetFolderInfo()   ' 無參數 overload = 全表模式
        For Each kvp In all
            Dim path = kvp.Key : Dim row = kvp.Value
            ' 只有 NOT NULL(=非 -1)的欄位才塞入記憶體快取；-1 代表「從未測量過」，跳過
            If row.mc >= 0 Then _cacheMailCount.TryAdd(path, row.mc)
            If row.mca >= 0 Then _cacheMailCountAll.TryAdd(path, row.mca)
            If row.fc >= 0 Then _cacheFolderCount.TryAdd(path, row.fc)
            If row.fca >= 0 Then _cacheFolderCountAll.TryAdd(path, row.fca)
            If row.fs >= 0 Then _cacheFolderSize.TryAdd(path, row.fs)
            If row.fsa >= 0 Then _cacheFolderSizeAll.TryAdd(path, row.fsa)

            ' by Gemini 3.0 flash, 2026/04/18: 批量讀取時回填身分標識與標籤字典，確保 LoadCache 後狀態完整
            If row.eid <> "" Then _cacheFolderIDs.TryAdd(path, (row.eid, row.sid, row.isMail = 1, row.hasCh = 1))
        Next
        Return all.Count
    End Function
    Private Function LoadAllFolderInfo() As List(Of FolderInfoDbRow)
        ''' <summary>
        ''' 一次性唯讀撈出 folder_info 全量名單
        ''' </summary>
        ' 2026/07/07 by Simon/Claude: 消重 — 改委派 LazyGetFolderInfo() 全表 overload(含 _dbCache 為 Nothing 的早退)，
        '   取代自己重寫的 SELECT(原本缺 has_chinese，現與另外兩條 folder_info 讀取路徑共用同一份完整欄位查詢)
        Return LazyGetFolderInfo().Values.ToList()
    End Function
#End Region

#Region "■ Db 單筆查詢 (DbGet* Lazy SELECT)"
    ' Phase 2 — Layer2.5 lazy SELECT 用的 DB read helper 群
    ' ==============================================================
    ' 設計原則 (2026-04-07):
    '   1. 只做「讀」，不做「寫」。寫入仍由 SaveCachesToDB (SaveCache 按鈕) 批次處理。
    '   2. 回傳 Nothing 表示 DB 中無此筆資料，呼叫端應繼續往 Layer3 走。
    '   3. 這些函數從 UI 執行緒呼叫，SQLite keyed lookup < 1ms，不需要 Async。
    '   4. FolderInfoDbRow / AttMailListDbResult 定義在本檔，Partial Class 跨檔可見。
    ' ==============================================================
    Private Function LazyGetDateCountCore(sql As String, errCtx As String, fPath As String, Optional year As Integer = Integer.MinValue) As ConcurrentDictionary(Of Integer, Integer)
        ' ---------------------------------------------------------------
        ' LazyGetDateCountCore — year_count / month_count 單鍵查詢共用核心
        ' sql 固定回傳兩欄 (subKey, count)；year 傳 Integer.MinValue 表示不綁 @yr 參數
        ' 回傳 Nothing 表示 DB 中無此鍵的記錄，呼叫端應繼續往 Layer3 走
        ' 2026/06/12 by Gemini/Simon: 把 folder_path TEXT 改成 folder_hash INTEGER 儲存，並新增一個記憶體字典 _dictHashToPath 來反查雜湊對應的路徑
        ' 2026/07/08 by Simon/Claude: 消重 — 原 LazyGetYearCount / LazyGetMonthCount
        '   兩函式 90% 同構(只差 SQL/參數/錯誤訊息)，reader 迴圈與錯誤處理收斂到這裡，
        '   兩個具名 wrapper 保留呼叫端語意。
        ' ---------------------------------------------------------------
        If _iLikeNoisy Then _dbg(" ├ 開始")
        If _dbCache Is Nothing Then Return Nothing

        Try
            Dim result As New ConcurrentDictionary(Of Integer, Integer)()
            Using cmd As New SqliteCommand(sql, _dbCache)
                cmd.Parameters.AddWithValue("@fh", FolderPathToHash64(fPath))
                If year <> Integer.MinValue Then cmd.Parameters.AddWithValue("@yr", year)
                Using reader = cmd.ExecuteReader()
                    While reader.Read() : result(reader.GetInt32(0)) = reader.GetInt32(1) : End While
                End Using
            End Using
            Return If(Not result.IsEmpty, result, Nothing)

        Catch ex As System.Exception
            _dbg(" ├ 錯誤", $"{errCtx}: {ex.Message}")
        Finally : If _iLikeNoisy Then _dbg(" ├ 結束")
        End Try
        Return Nothing

    End Function
    Private Function LazyGetYearCount(fPath As String) As ConcurrentDictionary(Of Integer, Integer)
        ' LazyGetYearCount — 讀取 year_count WHERE folder_hash=? 的所有行 (year → count)
        ' 供 ComputeYearCountsAsync 在記憶體 miss 時先查 DB，避免 COM 呼叫
        Return LazyGetDateCountCore("SELECT year,count FROM year_count WHERE folder_hash=@fh", fPath, fPath)
    End Function
    Private Function LazyGetMonthCount(fPath As String, year As Integer) As ConcurrentDictionary(Of Integer, Integer)
        ' LazyGetMonthCount — 讀取 month_count WHERE folder_hash=? AND year=? (month → count)
        ' 供 GetMonthCount(L2.5) 在記憶體 miss 時先查 DB，避免 COM 呼叫
        '   month_count 新增函數群 (2026/04/09 by Claude)，2026/04/09 改三欄 PK 接收 (fPath, year)
        Return LazyGetDateCountCore("SELECT month,count FROM month_count WHERE folder_hash=@fh AND year=@yr", $"{fPath} {year}", fPath, year)
    End Function

    Private Function LazyGetAttMailList(fPath As String) As AttMailListDbResult
        ' ---------------------------------------------------------------
        ' LazyGetAttMailList(fPath) — 讀取 att_maillist WHERE folder_path=? 的所有行
        ' 回傳 Nothing 表示 DB 中無此資料夾的郵件記錄
        ' 2026/07/07 by Simon/Claude: 消重 — 委派 LoadAttMailListCore 傳單元素清單，reader 迴圈不再維護兩份
        ' ---------------------------------------------------------------
        Dim batch = LoadAttMailListCore(New List(Of String) From {fPath})
        Dim value As AttMailListDbResult = Nothing
        If Not batch.TryGetValue(fPath, value) Then Return Nothing
        Return value
    End Function
    Private Function LazyGetAttMailList() As Dictionary(Of String, AttMailListDbResult)
        ' ---------------------------------------------------------------
        ' LazyGetAttMailList() 無參數 overload — 全表模式(不加 WHERE)，供 LoadAttMailListBatch 批次載入委派。
        '   依賴 _dictHashToPath 已還原(LoadFolderInfoBatch 先跑)，反查不到的 hash 列跳過，與原全表 reader 行為一致。
        ' 2026/07/07 by Simon/Claude: Overloads 兩簽名 — (fPath)單夾 / ()全表，reader 迴圈只有一份(LoadAttMailListCore)
        ' ---------------------------------------------------------------
        Return LoadAttMailListCore(CType(Nothing, List(Of String)))
    End Function
    Private Function LazyGetAttFilenames(entryId As String) As List(Of String)
        ' ---------------------------------------------------------------
        ' LazyGetAttFilenames — 讀取 att_filenames WHERE entry_id=? 的一行
        ' 回傳 Nothing 表示 DB 中無此 EntryID
        ' ---------------------------------------------------------------
        If _iLikeNoisy Then _dbg(" ├ 開始")
        If _dbMail Is Nothing Then Return Nothing   ' 2026/06/21 by Simon/Claude Opus 4.8: att_filenames 已搬至 _dbMail(OLAcacheMail.db)

        Try
            Using cmd As New SqliteCommand("SELECT filenames FROM att_filenames WHERE entry_id=@eid", _dbMail)
                ' 2026/06/23 by Claude Sonnet 4.6: entry_id 欄位是 BLOB，必須先轉 Byte() 再查詢
                ' 原本 AddWithValue(@eid, entryId As String) 用 Text 比對 BLOB，SQLite 型別不符，永遠查不到 → lazy load 永遠 miss
                cmd.Parameters.Add("@eid", SqliteType.Blob).Value = HexStringToByteArray(entryId)
                Using reader = cmd.ExecuteReader()
                    If reader.Read() AndAlso Not reader.IsDBNull(0) Then Return JsonSerializer.Deserialize(Of List(Of String))(reader.GetString(0))
                End Using
            End Using

        Catch ex As System.Exception
            _dbg(" ├ 錯誤", $"{entryId}: {ex.Message}")
        Finally : If _iLikeNoisy Then _dbg(" ├ 結束")
        End Try
        Return Nothing

    End Function

    Private Function LazyGetMailInfo(fPath As String) As (Mails As List(Of (Mail As MailItemInfo, Topic As String)), Snap As Integer)?
        ' ---------------------------------------------------------------
        ' LazyGetMailInfo(fPath) — 讀取 mail_info WHERE folder_path=? 的所有行
        ' 回傳 Nothing 表示 DB 中無此資料夾的郵件記錄
        ' 2026/06/10 by Gemini/Simon: 優化SQLite儲存空間把Entry_id改成BLOB二進位儲存
        ' 2026/06/11 by Gemini/Simon: 把 message_id 轉成 xxHash64，並同時改成 BLOB 儲存節省空間
        ' 2026/06/12 by Simon/Claude Opus 4.8: topic 改動態計算；sender_id→email；received_time 改 INTEGER
        ' 2026/07/07 by Simon/Claude: 消重 — 原本 40 行 reader 迴圈(欄位/EMPTY_BASIC_ 哨兵/sender 反查)與
        '   LoadMailInfoCore(List) 完全同構，改委派 List overload 傳單元素清單。List 版查無此夾時 dict 不含 key，
        '   即原本的 Return Nothing 語意(含 _dbCache Is Nothing 的早退，List 版自帶)。
        ' ---------------------------------------------------------------
        Dim batch = LoadMailInfoCore(New List(Of String) From {fPath})
        Dim value As (Mails As List(Of (Mail As MailItemInfo, Topic As String)), Snap As Long) = Nothing
        If Not batch.TryGetValue(fPath, value) Then Return Nothing
        _dbg(" ├ 命中 SSD", $"{ExtractFolderName(fPath)} | 取得 {value.Mails.Count} 筆 | Snap={value.Snap}")
        Return (value.Mails, CInt(value.Snap))
    End Function
    Private Function LazyGetMailInfo() As Dictionary(Of String, (Mails As List(Of (Mail As MailItemInfo, Topic As String)), Snap As Long))
        ' ---------------------------------------------------------------
        ' LoadMailInfoCore() 無參數 overload — 全表模式(不加 WHERE)，供 LoadMailInfoBatch 批次載入委派。
        '   依賴 _dictHashToPath 已還原(LoadFolderInfoBatch 先跑)，反查不到的 hash 列跳過，與原全表 reader 行為一致。
        '   刻意不讓外部用 List 版傳 Nothing 表達全表 — overload 下 LoadMailInfoCore(Nothing) 有 String/List 歧義。
        ' 2026/07/07 by Simon/Claude: Overloads 三簽名 — (fPath)單夾 / (List)IN 過濾 / ()全表，reader 迴圈只有一份
        ' ---------------------------------------------------------------
        Return LoadMailInfoCore(CType(Nothing, List(Of String)))
    End Function

    Private Function LazyGetFolderInfo(fPath As String) As FolderInfoDbRow
        ' ---------------------------------------------------------------
        ' LazyGetFolderInfo(fPath) — 讀取 folder_info 單一 folder_path 的一行
        ' 回傳 Nothing 表示 DB 中無此路徑
        ' 2026/07/07 by Simon/Claude: 消重 — 委派 LoadFolderInfoCore 傳單元素清單，reader 迴圈不再維護三份
        ' ---------------------------------------------------------------
        Dim batch = LoadFolderInfoCore(New List(Of String) From {fPath})
        Dim value As FolderInfoDbRow = Nothing
        If Not batch.TryGetValue(fPath, value) Then Return Nothing
        Return value
    End Function
    Private Function LazyGetFolderInfo() As Dictionary(Of String, FolderInfoDbRow)
        ' ---------------------------------------------------------------
        ' LazyGetFolderInfo() 無參數 overload — 全表模式(不加 WHERE)，供 LoadFolderInfoBatch/LoadAllFolderInfo 委派。
        ' 2026/07/07 by Simon/Claude: Overloads 兩簽名 — (fPath)單夾 / ()全表，reader 迴圈只有一份(LoadFolderInfoCore)
        ' ---------------------------------------------------------------
        Return LoadFolderInfoCore(CType(Nothing, List(Of String)))
    End Function
    Private Function LazyGetFolderIdAsList(fPath As String) As List(Of String)
        ' ---------------------------------------------------------------
        ' LazyGetFolderIdAsList — 撈出單一資料夾在 mail_info 的全部 entry_id(轉 hex 字串)。
        '   mail_info 是該夾「逐封郵件」的權威清單(att_filenames/mail_simhash 的 entryID 皆其子集)，
        '   供 ② Surgical 差集找已刪郵件。注意：呼叫端若會 DELETE mail_info，務必在 DELETE 之前呼叫本函式。
        ' 2026/06/22 by Simon/Claude: ② Surgical 輔助
        ' 2026/07/07 by Simon/Claude: 消重 — 原 reader 迴圈與 LazyGetFolderIdAsDict 只差一欄 msg_size，
        '   改委派加寬版取 Keys(多讀一欄的成本可忽略)，同一段 SQL/錯誤處理不再維護兩份
        ' ---------------------------------------------------------------
        Return LazyGetFolderIdAsDict(fPath).Keys.ToList()
    End Function
    Private Function LazyGetFolderIdAsDict(fPath As String) As Dictionary(Of String, Long)
        ' ---------------------------------------------------------------
        ' LazyGetFolderIdAsDict — LazyGetFolderIdAsList 的加寬版：一併帶回 msg_size (NULL 存 -1)。
        '   供 RenewCache 狀況A 的「就地修改偵測」：entry_id 相同但 msg_size 改變 = 內容被修改，
        '   simhash/bigram_set/att_filenames 已過期但不會落入 absent 差集，需另行清除。
        '   同 LazyGetFolderIdAsList 注意事項：呼叫端若會 DELETE mail_info，務必在 DELETE 之前呼叫本函式。
        ' 2026/07/07 by Simon/Claude
        ' ---------------------------------------------------------------
        Dim result As New Dictionary(Of String, Long)()
        If _dbCache Is Nothing Then Return result
        Dim fh = FolderPathToHash64(fPath)
        Try
            Using cmd As New SqliteCommand("SELECT entry_id, msg_size FROM mail_info WHERE folder_hash=@fh", _dbCache)
                cmd.Parameters.AddWithValue("@fh", fh)
                Using r = cmd.ExecuteReader()
                    While r.Read()
                        If Not r.IsDBNull(0) Then result(ByteArrayToHexString(r.GetFieldValue(Of Byte())(0))) = If(r.IsDBNull(1), -1L, r.GetInt64(1))
                    End While
                End Using
            End Using
        Catch ex As System.Exception
            _dbg("DbGetFolderEntryIdSizes 錯誤", $"{fPath}: {ex.Message}")
        End Try
        Return result
    End Function
    Private Function LazyGetSubFolderIDAsList(rootPath As String, isIncludeAll As Boolean) As List(Of FolderInfoDbRow)
        ' ---------------------------------------------------------------
        ' LazyGetSubFolderIDAsList — [優化 BFS] 利用 LIKE 一次抓出整棵子樹的所有資料夾身分證
        ' ---------------------------------------------------------------
        ' LazyGetOrderedSubFolderIDs vs LazyGetSubFolderIDAsList 這兩個不建議强行合併——
        ' 雖然長得像（都回傳 List(Of FolderInfoDbRow)、都用 LIKE），但 WHERE 深度限制、是否排序、root 例外、欄位子集四點都不同，
        ' 是「不同查詢意圖」不是「同查詢換參數」。硬合併只會產生一個帶 4 個布林/字串參數的怪函數，可讀性反而變差。
        ' 2026/07/07 by Simon/Claude
        ' ---------------------------------------------------------------
        If _iLikeNoisy Then _dbg(" ├ 開始")
        If _dbCache Is Nothing Then Return Nothing
        Try
            Dim result As New List(Of FolderInfoDbRow)(512)
            ' 過濾條件: 路徑以 rootPath 開頭，且 entry_id 不為空。若沒勾全選，則只抓 is_mail=1 的。
            ' 2026/06/13 by Simon/Claude Opus 4.8: 補 folder_count → 填 row.fc，供 IsSubtreeComplete 的「集合內子夾數 == fc」完整性檢查
            ' 2026/07/04 by Simon/Claude Fable 5 [rootless 骨架未爆彈]: root 列的 entry_id 歷史上多為 NULL(GetSubtreeRdo 不註冊 root
            '   → SaveCache 寫不進 root 身分證),被 entry_id IS NOT NULL 濾掉 → GetSubtree ② 回「缺 root 的骨架」→
            '   FilterSubtreeByMode 從 root 起走直接全滅(Tab1 只剩群組標題列)。改為 root 列豁免 entry_id 過濾(呼叫端以 live rootFolder 補身分證),
            '   is_mail/has_chinese 同步改 NULL 安全讀取(root 列這兩欄也可能為 NULL,原 GetInt32 直接拋例外)。
            Dim filter = If(isIncludeAll, "", " AND is_mail=1")
            Dim sql = $"SELECT folder_path,entry_id,store_id,is_mail,has_chinese,folder_count FROM folder_info " &
                      $"WHERE folder_path LIKE @fp || '%' AND (entry_id IS NOT NULL OR folder_path = @fp)" & filter

            Using cmd As New SqliteCommand(sql, _dbCache)
                cmd.Parameters.AddWithValue("@fp", rootPath)
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        ' 2026/06/10 by Gemini/Simon: 優化SQLite儲存空間把Entry_id改成BLOB二進位儲存
                        result.Add(New FolderInfoDbRow With {.path = reader.GetString(0),
                                                             .eid = If(reader.IsDBNull(1), "", ByteArrayToHexString(reader.GetFieldValue(Of Byte())(1))),
                                                             .sid = If(reader.IsDBNull(2), "", reader.GetString(2)),
                                                             .isMail = If(reader.IsDBNull(3), -1, reader.GetInt32(3)),
                                                             .hasCh = If(reader.IsDBNull(4), 0, reader.GetInt32(4)),
                                                             .fc = If(reader.IsDBNull(5), -1L, reader.GetInt64(5))})
                    End While
                End Using
            End Using
            Return If(result.Count > 0, result, Nothing)
        Catch ex As System.Exception
            _dbg(" ├ 錯誤", rootPath & ": " & ex.Message)
        Finally : If _iLikeNoisy Then _dbg(" ├ 結束")
        End Try
        Return Nothing
    End Function
    Private Function LazyGetOrderedSubFolderIDs(parentPath As String, isIncludeAll As Boolean) As List(Of FolderInfoDbRow)
        ' ---------------------------------------------------------------
        ' LazyGetOrderedSubFolderIDs — [優化 TreeView] 抓出直屬子目錄的身分證，並由 SQL 完成「英文優先」排序
        ' ---------------------------------------------------------------
        ' LazyGetOrderedSubFolderIDs vs LazyGetSubFolderIDAsList 這兩個不建議强行合併——
        ' 雖然長得像（都回傳 List(Of FolderInfoDbRow)、都用 LIKE），但 WHERE 深度限制、是否排序、root 例外、欄位子集四點都不同，
        ' 是「不同查詢意圖」不是「同查詢換參數」。硬合併只會產生一個帶 4 個布林/字串參數的怪函數，可讀性反而變差。
        ' 2026/07/07 by Simon/Claude
        ' ---------------------------------------------------------------
        If _iLikeNoisy Then _dbg(" ├ 開始")
        If _dbCache Is Nothing Then Return Nothing

        Try
            Dim result As New List(Of FolderInfoDbRow)(512)
            ' SQL 邏輯: 
            '   1. 找出 folder_path 以 parentPath + "\" 開頭。
            '   2. 且不包含更深層的 "\" (代表是直屬子項) 。注意: 此邏輯在路徑分隔符不一致時需調整。
            '   3. 按照 has_chinese ASC (0=英, 1=中, 故英優先) 排序。
            Dim filter = If(isIncludeAll, "", " AND is_mail=1")

            ' 精確匹配直屬子目錄：利用 LENGTH + REPLACE 來算出層級
            ' 或是利用路徑字串特性：新的路徑長度應該是在 parent 之後且沒有多餘的層級
            ' 簡化做法：目前專案路徑是用 \ 分隔。
            Dim sql = "SELECT folder_path,entry_id,store_id,is_mail,has_chinese FROM folder_info " &
                      " WHERE folder_path LIKE @fp || '\%' AND entry_id IS NOT NULL " & filter &
                      "   AND folder_path NOT LIKE @fp || '\%\%' " & ' 排除第二層以後的
                      " ORDER BY has_chinese ASC, folder_path ASC"

            Using cmd As New SqliteCommand(sql, _dbCache)
                cmd.Parameters.AddWithValue("@fp", parentPath)
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        ' 2026/06/10 by Gemini/Simon: 優化SQLite儲存空間把Entry_id改成BLOB二進位儲存
                        result.Add(New FolderInfoDbRow With {.path = reader.GetString(0),
                                                              .eid = ByteArrayToHexString(reader.GetFieldValue(Of Byte())(1)),
                                                              .sid = If(reader.IsDBNull(2), "", reader.GetString(2)),
                                                              .isMail = reader.GetInt32(3),
                                                              .hasCh = reader.GetInt32(4)})
                    End While
                End Using
            End Using

            Return If(result.Count > 0, result, Nothing)
        Catch ex As System.Exception
            _dbg(" ├ 錯誤", parentPath & ": " & ex.Message)
        Finally : If _iLikeNoisy Then _dbg(" ├ 結束")
        End Try
        Return Nothing
    End Function
    Private Function PeekLiveFolderId(fPath As String) As String
        ' ---------------------------------------------------------------
        ' PeekLiveFolderId — 只撈「一筆」entry_id(hex)，不是整夾清單。
        ' 2026/07/04 by Simon/Claude Sonnet 5
        ' 供 RenewCacheToDB 狀態B(count/commit_max 都判定沒變)分支做低成本存活探測：
        '   單次 GetItemFromID 遠比整夾 GetTable 表掃便宜，用來抓「純 PST 壓縮換 entry_id」這種
        '   兩個便宜訊號都抓不到的邊界案例 —— 壓縮通常整批 ID 一起換掉，抽一筆探測即可命中。
        ' 2026/07/09 by Simon/Claude: 修正 — 空資料夾在 mail_info 只有 EMPTY_BASIC_ 哨兵列(非真 entry_id)，
        '   之前沒濾掉就直接餵給呼叫端的 GetItemFromID，保證每次都失敗，導致所有空資料夾被誤判 dirty 強制整夾重掃，
        '   RenewCache 因此跳出大量例外且耗時暴增(entry_id/att_maillist 等表其他讀取路徑本來就有濾這個哨兵，這裡漏掉了)。
        ' ---------------------------------------------------------------
        If _dbCache Is Nothing Then Return ""
        Dim fh = FolderPathToHash64(fPath)
        Try
            Using cmd As New SqliteCommand("SELECT entry_id FROM mail_info WHERE folder_hash=@fh LIMIT 1", _dbCache)
                cmd.Parameters.AddWithValue("@fh", fh)
                Dim result = cmd.ExecuteScalar()
                If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                    Dim eid = ByteArrayToHexString(DirectCast(result, Byte()))
                    If Not eid.StartsWith("EMPTY_", StringComparison.OrdinalIgnoreCase) Then Return eid
                End If
            End Using
        Catch ex As System.Exception
            _dbg("DbGetSampleEntryId 錯誤", $"{fPath}: {ex.Message}")
        End Try
        Return ""
    End Function
#End Region

#Region "■ Db 單筆寫入/清除 (Single-Row Save / Delete)"
    Private Sub PoisonFolderSnapDb(fPath As String)
        ' ---------------------------------------------------------------
        ' PoisonFolderSnapDb — 自癒用：強制下次 RenewCacheToDB 判定此資料夾 dirty
        ' 呼叫時機：任何用快取 entry_id 呼叫 GetItemFromID/RDO GetMessageFromID 解析失敗 (NotFound) 時，
        '   代表 DB 裡這個資料夾至少有一顆 entry_id 已死 (常見成因：PST 壓縮換 ID)。
        '   Count/CommitMax 雙訊號仍可能剛好都沒變 (壓縮不改數量也不改修改時間)，
        '   所以直接把 pr_count_snap 寫成 PeekLiveFolderSnapOOM 同款的「不可能值」-999、commit_max 清 NULL，
        '   讓下次 RenewCache 的 liveSnap<>row.snap 必為 True → 強制走狀況A全量重讀 + surgical entry_id 清理。
        ' 只毒化 snapshot，不動 entry_id/store_id/統計欄位 —— 資料夾本身多半還在，只是內容郵件的 ID 換了。
        ' 2026/07/04 by Simon/Claude Fable 5
        ' ---------------------------------------------------------------
        If _dbCache Is Nothing OrElse String.IsNullOrEmpty(fPath) Then Return
        Try
            Using cmd As New SqliteCommand("UPDATE folder_info SET pr_count_snap = -999, commit_max = NULL WHERE folder_path = @fp", _dbCache)
                cmd.Parameters.AddWithValue("@fp", fPath)
                cmd.ExecuteNonQuery()
            End Using
            _dbg("PoisonFolderSnapDb", $"{ExtractFolderName(fPath)} — 已毒化，下次 RenewCache 強制重讀")
        Catch ex As System.Exception
            _dbg("PoisonFolderSnapDb 錯誤", ex.Message)
        End Try
    End Sub
    Private Sub DbSaveMonthCountSingle(fPath As String, year As Integer, monthCount As ConcurrentDictionary(Of Integer, Integer))
        ' ---------------------------------------------------------------
        ' DbSaveMonthCountSingle — 增量寫入單一 (folder_path, year) 的月份分布
        '   在 GetMonthCount 的 ③ COM 計算(RDO 或 OOM)完成後立刻呼叫，不等待 SaveCache 按鈕。
        '   使用獨立 Transaction 包住最多 12 筆，確保原子性。
        ' 2026/04/09 新增 by Claude：解決月份快取只在記憶體、SaveCache 才寫 DB 的問題
        '   根本原因：若該 session 沒點過月份視圖就不 SaveCache，下次仍打 COM
        '   修正後：每次 L3 計算完月份後立刻持久化，下次 DB lazy 直接命中
        ' ---------------------------------------------------------------
        If _iLikeNoisy Then _dbg(" ├ 開始")
        If _dbCache Is Nothing OrElse monthCount Is Nothing OrElse monthCount.IsEmpty Then Return
        Try
            Using txn = _dbCache.BeginTransaction()
                Using cmd As New SqliteCommand(
                    "INSERT OR REPLACE INTO month_count (folder_hash,year,month,count) VALUES (@fh,@yr,@mo,@cnt)", _dbCache, txn)
                    ' 2026/06/12 by Gemini/Simon: 把 folder_path TEXT 改成 folder_hash INTEGER 儲存，並新增一個記憶體字典 _dictHashToPath 來反查雜湊對應的路徑
                    ' 2026/06/12 by Simon/Claude Opus 4.8: 移除 updated_at（只寫不讀）
                    cmd.Parameters.Add("@fh", SqliteType.Integer).Value = FolderPathToHash64(fPath)
                    cmd.Parameters.Add("@yr", SqliteType.Integer).Value = year
                    cmd.Parameters.Add("@mo", SqliteType.Integer)
                    cmd.Parameters.Add("@cnt", SqliteType.Integer)
                    For Each mo In monthCount
                        cmd.Parameters("@mo").Value = mo.Key
                        cmd.Parameters("@cnt").Value = mo.Value
                        cmd.ExecuteNonQuery()
                    Next
                End Using
                txn.Commit()
            End Using
            _dbg(" ├ 更新", $"{fPath} {year} → {monthCount.Count} 個月寫入 DB")

        Catch ex As System.Exception
            _dbg(" ├ 錯誤", $"{fPath} {year}: {ex.Message}")
        Finally : If _iLikeNoisy Then _dbg(" ├ 結束")
        End Try

    End Sub
    Private Sub DbPurgeFolderMailRows(fPath As String, Optional includeAttFilenames As Boolean = False)
        ' ---------------------------------------------------------------
        ' DbPurgeFolderMailRows — 刪除單一資料夾在逐封郵件表的全部列 (mail_info/att_maillist/month_count/year_count，選擇性含 att_filenames)。
        '   用於「資料夾還在但內含郵件有增刪」時清掉死列(失效 entryID)，維持「同一資料夾的 basic/att 列共用單一 snap」不變量，根除讀取端混 snap 幽靈郵件。
        '   與 CleanupOrphanPath 的差異：那個是整夾消失才連 folder_info 全表一起刪；本函式只清逐封郵件列，不動 folder_info。
        ' 2026/06/20 by Simon/Claude: 取代原 RenewAttMailList 三路比對
        ' 2026/07/06 by Simon/Claude Fable 5: 補 year_count — 原本全專案沒有任何地方 DELETE year_count(只進不出)，
        '   而 GetYearCount ② DB lazy 無 snap 驗證，資料夾變動後記憶體清了、DB 舊列卻會被原樣撈回來復活。
        '   各呼叫端本就有清 _cacheYearCount 記憶體，補上 DB 端這一刀失效鏈才算閉合。
        ' ---------------------------------------------------------------
        If _dbCache Is Nothing Then Return
        Dim fh = FolderPathToHash64(fPath)
        Try
            Using txn As SqliteTransaction = _dbCache.BeginTransaction()
                For Each tbl In {"mail_info", "att_maillist", "month_count", "year_count"}
                    Using c As New SqliteCommand($"DELETE FROM {tbl} WHERE folder_hash=@fh", _dbCache, txn)
                        c.Parameters.AddWithValue("@fh", fh) : c.ExecuteNonQuery()
                    End Using
                Next
                txn.Commit()
            End Using
        Catch ex As System.Exception
            _dbg("DbPurgeFolderMailRows 錯誤", $"{fPath}: {ex.Message}")
        End Try
        ' 2026/06/21 by Simon/Claude: att_filenames 已搬至 OLAcacheMail.db(_dbMail)，跨檔不能掛 _dbCache txn，改獨立交易刪除
        If includeAttFilenames Then SimDbDeleteAttFilenamesByFolder(fPath)
    End Sub
    Private Function SimDbDeleteAttFilenamesByFolder(fPath As String) As Integer
        ' ---------------------------------------------------------------
        ' SimDbDeleteAttFilenamesByFolder — 刪除單一資料夾在 att_filenames(OLAcacheMail.db/_dbMail) 的全部列。
        '   獨立 _dbMail 交易，供 DbPurgeFolderMailRows(RenewCache 狀況 A) 與 CleanupOrphanPath(整夾消失) 兩處呼叫。
        '   回傳刪除列數(供 CleanupOrphanPath 統計顯示)。
        ' 2026/06/21 by Simon/Claude: Part 2 拆檔耦合點 — 原本掛在 _dbCache txn 的 att_filenames DELETE 抽到此 helper
        ' ---------------------------------------------------------------
        If _dbMail Is Nothing Then Return 0
        Dim fh = FolderPathToHash64(fPath) : Dim n As Integer = 0
        Try
            Using txn = _dbMail.BeginTransaction()
                Using c As New SqliteCommand("DELETE FROM att_filenames WHERE folder_hash=@fh", _dbMail, txn)
                    c.Parameters.AddWithValue("@fh", fh) : n = c.ExecuteNonQuery()
                End Using
                txn.Commit()
            End Using
        Catch ex As System.Exception
            _dbg("錯誤", $"{fPath}: {ex.Message}")
        End Try
        Return n
    End Function
    Private Function SimDbDeleteMailRowsByEntryIds(entryIds As ICollection(Of String), Optional includeAttFilenames As Boolean = True) As Integer
        ' ---------------------------------------------------------------
        ' SimDbDeleteMailRowsByEntryIds — ② Surgical：依「已刪 entryID 集合」精準清除兩張「逐封讀取極貴」的快取。
        '   記憶體：一律 TryRemove _cacheAttFilename + _cacheSimHash(兩者 key 皆 EntryID 字串)。
        '   DB(_dbMail)：mail_simhash 一律刪(無 folder_hash，只能靠 entryID)；
        '                att_filenames 視 includeAttFilenames —— 狀況 A(夾內增刪)傳 True 逐封刪；
        '                CleanupOrphanPath(整夾消失)傳 False(已由 SimDbDeleteAttFilenamesByFolder 按 folder_hash 高效刪過，免重複)。
        '   只清失效的那幾封，存活郵件的昂貴快取保留(免重讀內文/附件)。回傳 mail_simhash 刪除列數(供 log)。
        ' 2026/06/22 by Simon/Claude: ② Surgical 策略 — 嚴格清除失效 entryID，杜絕昂貴快取死列永久累積
        ' ---------------------------------------------------------------
        If entryIds Is Nothing OrElse entryIds.Count = 0 Then Return 0
        For Each eid In entryIds : _cacheAttFilename.TryRemove(eid, Nothing) : _cacheSimHash.TryRemove(eid, Nothing) : Next
        If _dbMail Is Nothing Then Return 0
        Dim nSh As Integer = 0
        Try
            Using txn = _dbMail.BeginTransaction()
                Using cSh As New SqliteCommand("DELETE FROM mail_simhash WHERE entry_id=@eid", _dbMail, txn),
                      cAf As New SqliteCommand("DELETE FROM att_filenames WHERE entry_id=@eid", _dbMail, txn)
                    cSh.Parameters.Add("@eid", SqliteType.Blob) : cAf.Parameters.Add("@eid", SqliteType.Blob)
                    For Each eid In entryIds
                        Dim blob = HexStringToByteArray(eid)
                        cSh.Parameters("@eid").Value = blob : nSh += cSh.ExecuteNonQuery()
                        If includeAttFilenames Then cAf.Parameters("@eid").Value = blob : cAf.ExecuteNonQuery()
                    Next
                End Using
                txn.Commit()
            End Using
        Catch ex As System.Exception
            _dbg("SimDbDeleteMailRowsByEntryIds 錯誤", ex.Message)
        End Try
        Return nSh
    End Function

    ' ── Tab5 S5 候選 bigram set 核心快取 (2026/07/07 by Simon/Claude) ──────────────────────
    '   只存「通過 S4 篩選、進 S5 精算」的候選信集合(全庫平均每封 ~2.4KB, 數萬候選僅百餘 MB)，
    '   讓下次冷搜尋的 S5 免重讀 body。列與 mail_simhash 同源同進退, 失效跟著既有 purge 路徑走, 不需額外處理。
    Private Function BigramSetToBytes(setB As HashSet(Of Integer)) As Byte()
        ' HashSet(Of Integer) → 每元素 4 bytes 的 BLOB (不排序不壓縮, Buffer.BlockCopy 零解析成本)
        Dim arr = setB.ToArray()
        Dim bytes(arr.Length * 4 - 1) As Byte   ' 空集合時 Dim bytes(-1) 為合法的零長度陣列
        Buffer.BlockCopy(arr, 0, bytes, 0, bytes.Length)
        Return bytes
    End Function
    Private Function BytesToBigramSet(bytes As Byte()) As HashSet(Of Integer)
        Dim ints(bytes.Length \ 4 - 1) As Integer
        Buffer.BlockCopy(bytes, 0, ints, 0, bytes.Length)
        Return New HashSet(Of Integer)(ints)
    End Function
    Private Function SaveDbMailSets(rows As IEnumerable(Of (EntryID As String, SetBytes As Byte()))) As Integer
        ' 批次回填候選信的 bigram_set。列必然已由 S3 SaveDbMail 建立, 故用 UPDATE(打不中就略過, 不補列)。
        If _dbMail Is Nothing Then Return 0
        Dim count As Integer = 0
        Try
            Using txn = _dbMail.BeginTransaction()
                Using cmd As New SqliteCommand("UPDATE mail_simhash SET bigram_set=@bs WHERE entry_id=@eid", _dbMail, txn)
                    Dim pBs = cmd.Parameters.Add("@bs", SqliteType.Blob) : Dim pEid = cmd.Parameters.Add("@eid", SqliteType.Blob)
                    For Each row In rows
                        pEid.Value = HexStringToByteArray(row.EntryID) : pBs.Value = row.SetBytes
                        count += cmd.ExecuteNonQuery()
                    Next
                End Using
                txn.Commit()
            End Using
        Catch ex As System.Exception
            _dbg("       ├ 錯誤", ex.Message)
        End Try
        Return count
    End Function
    Private Function LoadDbMailSets(ids As List(Of String)) As Dictionary(Of String, HashSet(Of Integer))
        ' 依 EntryID 逐筆 PK 點查 bigram_set。純 SQLite+CPU, 呼叫端包 Task.Run 走背景執行緒。
        ' 沒存過的(bigram_set IS NULL)不進字典, 由呼叫端 fallback 讀 body。
        Dim result As New Dictionary(Of String, HashSet(Of Integer))(ids.Count)
        If _dbMail Is Nothing OrElse ids.Count = 0 Then Return result
        Try
            Using cmd As New SqliteCommand("SELECT bigram_set FROM mail_simhash WHERE entry_id=@eid", _dbMail)
                Dim pEid = cmd.Parameters.Add("@eid", SqliteType.Blob)
                For Each id In ids
                    pEid.Value = HexStringToByteArray(id)
                    Dim firstV = cmd.ExecuteScalar()
                    If firstV IsNot Nothing AndAlso Not TypeOf firstV Is DBNull Then result(id) = BytesToBigramSet(CType(firstV, Byte()))
                Next
            End Using
        Catch ex As System.Exception
            _dbg("       ├ 錯誤", ex.Message)   ' 整批失敗就回已載到的部分, 缺的走 body fallback
        End Try
        Return result
    End Function
#End Region

#Region "■ 資料庫維運/診斷 (Vacuum & Stat)"
    Private Async Function DbShowDbFileStat() As Task
        ''' <summary>
        ''' 點擊 ListView6 "DB 檔案大小" 時呼叫。2026/06/21 起改為「依 db 分區塊」：
        '''   OLAcache.db 與 OLAcacheMail.db 各跑一次兩段式統計(per-file 數學自洽)，最後輸出兩檔合計。
        '''   階段一 (UI thread): PRAGMA 頁面資訊 + 各表筆數；階段二 (Task.Run): CAST AS BLOB 算淨重 + 比例估算實體。
        ''' 因 e_sqlite3.dll 預設未啟用 SQLITE_ENABLE_DBSTAT_VTAB，改用「淨重比例分配」估算(相對排名準，絕對值 ±20%)。
        ''' 2026/06/13 by Simon/Claude Opus 4.8 / 2026/06/21 by Simon/Claude: 拆 OLAcacheMail.db 後改雙檔分區塊
        ''' </summary>
        If _dbCache Is Nothing Then Return
        If _lv6StatBusy Then _dbg("略過", "另一個 Lv6 統計查詢進行中，稍後再點 [DB 檔案大小]") : Return   ' 2026/07/11 by Simon/Claude Fable 5: 同連線不允許並發，見 _lv6StatBusy 宣告處
        _lv6StatBusy = True

        ' Task.Run 內呼叫 _dbg 的 stack trace 會抓到編譯器生成的 lambda 名稱，故預先封裝 forwarder
        ' 直接走 DebugForm.AddMessage3 並傳入 forcedCaller，與 _dbg() 的 Release-build 行為等效
        Dim _dbgFwd As Action(Of String, String) = Sub(a, b)
                                                       If _isDebugMode Then DebugForm.AddMessage3(a, b, "DbShowDbFileStat")
                                                   End Sub
        _dbgFwd("開始", "📊DB 檔案大小")

        Try
            ' 2026/06/21 by Simon/Claude: 兩個 db 各一段(淨重佔比/估算實體的分母不可跨檔)，依序輸出後再合計
            Dim r1 = Await DbShowDbFileStatCore(_dbCache, _dbCachePath, "OLAcache.db", _dbgFwd)
            Dim file2 As Single = 0F, net2 As Single = 0F
            If _dbMail IsNot Nothing Then
                Dim r2 = Await DbShowDbFileStatCore(_dbMail, _dbMailPath, "OLAcacheMail.db", _dbgFwd)
                file2 = r2.fileMB : net2 = r2.netMB
            End If

            ' ----- 兩檔合計 -----
            _dbgFwd("", "═══════════ 兩檔合計 ═══════════")
            _dbgFwd(" │", $" 檔案實體合計: {(r1.fileMB + file2):F2} MB    /    純資料淨重合計: {(r1.netMB + net2):F2} MB")
            _dbgFwd("結束", "DB 檔案大小")

        Catch ex As System.Exception
            _dbgFwd(" ├ 錯誤", $"DbShowDbFileStat: {ex.Message}")
        Finally
            _lv6StatBusy = False
        End Try
    End Function
    Private Async Function DbShowDbFileStatCore(conn As SqliteConnection, dbPath As String, dbLabel As String, dbgFwd As Action(Of String, String)) As Task(Of (fileMB As Single, netMB As Single))
        ' 2026/06/21 by Simon/Claude: 由 DbShowDbFileStat 抽出 — 對單一 (conn, dbPath) 跑兩段式統計並輸出，回傳 (fileMB, netMB) 供合計。
        '   原邏輯原樣搬入，僅把寫死的 _dbCache/_dbCachePath 參數化；section header 用粗線標示檔案身分。
        dbgFwd("", $"═══════════ {dbLabel} ═══════════")

        ' ===== 階段一：UI thread 上的瞬時資料 =====
        ' 1. 檔案實體大小
        Dim fpath As New IO.FileInfo(dbPath)
        Dim fileSizeMB As Single = fpath.Length / 1024 / 1024

        ' 2. PRAGMA 頁面資訊：page_size × page_count = DB 主檔內容大小 (不含 -wal/-shm 暫存檔)
        Dim pageSize As Integer = 0, pageCount As Long = 0, freelistPages As Long = 0
        Using cmd As New SqliteCommand("PRAGMA page_size", conn) : pageSize = Convert.ToInt32(cmd.ExecuteScalar()) : End Using
        Using cmd As New SqliteCommand("PRAGMA page_count", conn) : pageCount = Convert.ToInt64(cmd.ExecuteScalar()) : End Using
        Using cmd As New SqliteCommand("PRAGMA freelist_count", conn) : freelistPages = Convert.ToInt64(cmd.ExecuteScalar()) : End Using
        Dim freelistMB As Single = freelistPages * pageSize / 1024 / 1024
        Dim dbPagesMB As Single = pageCount * pageSize / 1024 / 1024

        ' 3. 列出全部 user table (排除 sqlite_ 開頭的系統表)
        Dim tableNames As New List(Of String)
        Using cmd As New SqliteCommand("Select name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%' ORDER BY name", conn)
            Using rd = cmd.ExecuteReader()
                While rd.Read() : tableNames.Add(rd.GetString(0)) : End While
            End Using
        End Using

        ' 2026/06/14 by Simon/Claude Opus 4.8: 依固定邏輯順序重排顯示，與 Tab6 Lv6 一致 (取代上面的字母序)。
        ' 不在清單內的新表自動排到尾端 (保留 ORDER BY name 的字母序)，避免將來新增表時漏顯示。
        ' 2026/06/21 by Simon/Claude: 追加 mail_simhash；各檔只列舉自己 sqlite_master 有的表，缺的自動不顯示，故兩檔共用同一份順序。
        Dim _displayOrder = New String() {"folder_info", "senders", "mail_info", "year_count", "month_count", "att_maillist", "att_filenames", "mail_simhash"}
        tableNames = tableNames.OrderBy(Function(n) If(Array.IndexOf(_displayOrder, n) < 0, Integer.MaxValue, Array.IndexOf(_displayOrder, n))).ToList()

        ' 4. 每張表的筆數 (COUNT(*) 很快，留在 UI thread 內)
        Dim rowCounts As New Dictionary(Of String, Long)
        For Each t In tableNames
            Using cmd As New SqliteCommand($"SELECT COUNT(*) FROM [{t}]", conn)
                rowCounts(t) = Convert.ToInt64(cmd.ExecuteScalar())
            End Using
        Next

        ' 5. 預先收集每張表的欄位清單 (供 Task.Run 內組 SQL，避免 Task.Run 內再讀 PRAGMA)
        Dim colMap As New Dictionary(Of String, List(Of String))
        For Each t In tableNames
            Dim cols As New List(Of String)
            Using cmd As New SqliteCommand($"PRAGMA table_info([{t}])", conn)
                Using rd = cmd.ExecuteReader()
                    While rd.Read() : cols.Add(rd("name").ToString()) : End While
                End Using
            End Using
            colMap(t) = cols
        Next

        ' ----- 階段一輸出 -----
        dbgFwd(" ├", $"資料庫路徑: {dbPath}")
        dbgFwd(" ├", $"檔案實體大小: {fileSizeMB:F2} MB (page_size={pageSize}, page_count={pageCount}, 主檔約 {dbPagesMB:F2} MB)")
        dbgFwd(" ├", $"freelist 碎片: {freelistPages} 頁 ≈ {freelistMB:F3} MB")
        dbgFwd(" ├", "資料表筆數:")
        For Each t In tableNames
            dbgFwd(" │", $" [{t}]".PadRight(22) & rowCounts(t).ToString("N0").PadLeft(10) & " 筆")
        Next

        ' ===== 階段二：背景算各表純資料淨重 =====
        ' Task.Run 內對 conn 操作：第一階段已結束、且兩支 helper 為 await 串接無 overlap，符合 SqliteConnection thread-safety 規範
        Dim netResult = Await Task.Run(Function() As Dictionary(Of String, Single)
                                           Dim result As New Dictionary(Of String, Single)
                                           For Each t In tableNames
                                               Try
                                                   Dim cols = colMap(t)
                                                   If cols.Count = 0 Then result(t) = 0 : Continue For
                                                   ' 2026/06/13 by Simon/Claude Opus 4.8: 改用逗號分隔，對齊 DbShowTableStat 的寫法。
                                                   '    原本用 " + " 串接所有 SUM()，SQLite 中 NULL + 任何值 = NULL，任一欄全為 NULL 時整條表達式變 NULL，外層 IsDBNull 判成 0F (錯誤)。
                                                   '    改為逗號分隔各欄成獨立 SELECT 欄位，每欄各自 IsDBNull 處理，互不干擾。
                                                   Dim parts As New List(Of String)
                                                   For Each c In cols : parts.Add($"SUM(length(CAST([{c}] AS BLOB)))") : Next
                                                   Dim sql = "SELECT " & String.Join(", ", parts) & $" FROM [{t}]"
                                                   Using cmd As New SqliteCommand(sql, conn)
                                                       Using rd = cmd.ExecuteReader()
                                                           If rd.Read() Then
                                                               Dim total As Single = 0
                                                               For i As Integer = 0 To cols.Count - 1
                                                                   If Not rd.IsDBNull(i) Then total += Convert.ToSingle(rd(i)) / 1024.0F / 1024.0F
                                                               Next
                                                               result(t) = total
                                                           Else
                                                               result(t) = 0
                                                           End If
                                                       End Using
                                                   End Using
                                               Catch ex As System.Exception
                                                   result(t) = -1 ' 標記錯誤，輸出時跳過
                                               End Try
                                           Next
                                           Return result
                                       End Function)

        ' ----- 階段二輸出：淨重 + 按比例估算實體佔用 -----
        ' 估算法：(扣除 freelist 後的有效實體) × (本表淨重 / 本檔全部淨重) ≈ 本表實體佔用
        ' 此法相對排名極準，絕對值含 index/row header/b-tree overhead 攤分，誤差 ±20% 內
        Dim totalNetMB As Single = 0
        For Each t In tableNames
            If netResult.ContainsKey(t) AndAlso netResult(t) >= 0 Then totalNetMB += netResult(t)
        Next
        Dim usefulMB As Single = dbPagesMB - freelistMB

        dbgFwd(" ├", "純資料淨重明細 (CAST AS BLOB):")
        For Each t In tableNames
            If Not netResult.ContainsKey(t) OrElse netResult(t) < 0 Then
                dbgFwd(" │", $" [{t}]".PadRight(22) & "讀取失敗")
                Continue For
            End If
            Dim netMB As Single = netResult(t)
            Dim estPhys As Single = If(totalNetMB > 0, usefulMB * (netMB / totalNetMB), 0F)
            dbgFwd(" │", $" [{t}]".PadRight(22) & $"{netMB:F2} MB".PadLeft(12) & $"  估算實體 ≈ {estPhys:F2} MB")
        Next
        dbgFwd(" │", $" 本檔資料表淨重總計: {totalNetMB:F2} MB")
        dbgFwd(" │", $" 檔案 vs 淨重 : {(fileSizeMB - totalNetMB):F2} MB (索引 + header + b-tree + freelist)")

        Return (fileSizeMB, totalNetMB)
    End Function
    Private Async Function DbShowTableStat(tableName As String) As Task
        ''' <summary>
        ''' 深度分析快取表：動態計算精準 Bytes 淨容量 (CAST AS BLOB 破解中文字元數陷阱)
        ''' 2026/07/11 by Simon/Claude Fable 5: 查詢全數移入 Task.Run — mail_info 35萬列全欄位
        '''   SUM(length(CAST AS BLOB)) 原本同步跑在 UI 執行緒會凍結數秒 (BC42356 警告即此)。
        '''   dbstat 分支移除：e_sqlite3.dll 未編譯 SQLITE_ENABLE_DBSTAT_VTAB (DbShowDbFileStat 註解已證實)，
        '''   該查詢永遠拋例外走 fallback 形同死碼；各表實體佔用估算改雙擊「DB 檔案大小」(整檔淨重比例分配法)。
        ''' </summary>

        ' 2026/06/21 by Simon/Claude: att_filenames/mail_simhash 住 OLAcacheMail.db(_dbMail)，依表名路由連線；其餘走 _dbCache
        Dim conn As SqliteConnection = If(tableName = "att_filenames" OrElse tableName = "mail_simhash", _dbMail, _dbCache)
        If conn Is Nothing OrElse String.IsNullOrEmpty(tableName) Then Return
        If _lv6StatBusy Then _dbg("略過", $"另一個 Lv6 統計查詢進行中，稍後再點 [{tableName}]") : Return
        _lv6StatBusy = True

        _dbg("開始", $"[📊{tableName}]")
        Try
            ' 查詢移入背景執行緒；回傳後在 UI context 輸出，_dbg 的呼叫端名稱不受 lambda 影響
            Dim r = Await Task.Run(
                Function()
                    ' 1. 動態取得欄位 MetaData
                    Dim cols As New List(Of (Cid As Integer, Name As String, Type As String, Pk As String, Nn As String))
                    Using cmd As New SqliteCommand($"PRAGMA table_info([{tableName}])", conn)
                        Using rd = cmd.ExecuteReader()
                            While rd.Read()
                                cols.Add((Convert.ToInt32(rd("cid")), rd("name").ToString(), rd("type").ToString(), If(Convert.ToInt32(rd("pk")) > 0, "★", ""), If(Convert.ToInt32(rd("notnull")) > 0, "Y", "N")))
                            End While
                        End Using
                    End Using

                    ' 2. 修正版 SQL：強制 CAST AS BLOB 算真實 Bytes，破解中文字元數陷阱！
                    Dim rowCount As Long = 0
                    Dim totalNetMB As Single = 0
                    Dim colSizes As New Dictionary(Of String, Single)
                    If cols.Count > 0 Then
                        Dim sbSql As New System.Text.StringBuilder("SELECT COUNT(*)")
                        For Each c In cols : sbSql.Append($", SUM(length(CAST([{c.Name}] AS BLOB)))") : Next
                        sbSql.Append($" FROM [{tableName}]")

                        Using cmd As New SqliteCommand(sbSql.ToString(), conn)
                            Using rd = cmd.ExecuteReader()
                                If rd.Read() Then
                                    rowCount = If(rd.IsDBNull(0), 0, rd.GetInt64(0))
                                    For i As Integer = 0 To cols.Count - 1
                                        Dim mb = If(rd.IsDBNull(i + 1), 0, Convert.ToSingle(rd(i + 1))) / 1024 / 1024
                                        colSizes(cols(i).Name) = mb : totalNetMB += mb
                                    Next
                                End If
                            End Using
                        End Using
                    End If
                    Return (Cols:=cols, RowCount:=rowCount, TotalNetMB:=totalNetMB, ColSizes:=colSizes)
                End Function)

            If r.Cols.Count = 0 Then _dbg(" ├ 錯誤", $"找不到表格 [{tableName}]") : Return

            ' 3. 輸出統計
            _dbg(" ├", $"總資料筆數: {r.RowCount} 筆")
            _dbg(" ├", $" {"欄位名稱".PadRight(16)}{"型態".PadRight(12)}欄位資料淨重")
            For Each c In r.Cols : _dbg(" │", $" {$"[{c.Name}]".PadRight(17)}  {c.Type.PadRight(8)}: {r.ColSizes(c.Name).ToString("F2")} MB") : Next
            _dbg(" │", $" 所有欄位純資料淨重 : {r.TotalNetMB.ToString("F2")} MB (程式寫入的真正大小)")
            _dbg(" │ 提醒", "(實體佔用另含 B-Tree/索引/Row Header 開銷，通常為淨重的 1.5~3 倍；各表估算請雙擊「DB 檔案大小」)")
            _dbg("結束", $"[{tableName}]")

        Catch ex As System.Exception : _dbg(" ├ 錯誤", $"分析 {tableName} 失敗: {ex.Message}")
        Finally : _lv6StatBusy = False
        End Try
    End Function
    Private Async Function DbShowBigramSetStat() As Task
        ''' <summary>
        ''' bigram_set 是 mail_simhash 表內的單一欄位(僅 Tab5 S5 精算過的候選信才會回填)，
        ''' 跟全表列數/大小意義不同，故獨立一支統計，不走 DbShowTableStat 的整表路徑。
        ''' </summary>
        If _dbMail Is Nothing Then Return
        If _lv6StatBusy Then _dbg("略過", "另一個 Lv6 統計查詢進行中，稍後再點 [bigram_set]") : Return
        _lv6StatBusy = True

        _dbg("開始", "[📊bigram_set]")
        Try
            ' 2026/07/11 by Simon/Claude Fable 5: SUM(LENGTH(BLOB)) 掃全表，移入 Task.Run 免凍結 UI (同 DbShowTableStat)
            Dim r = Await Task.Run(
                Function()
                    Dim total As Long = 0, filled As Long = 0, bytes As Long = 0
                    Using cmd As New SqliteCommand("SELECT COUNT(*), COUNT(bigram_set), IFNULL(SUM(LENGTH(bigram_set)),0) FROM mail_simhash", _dbMail)
                        Using rd = cmd.ExecuteReader()
                            If rd.Read() Then
                                total = rd.GetInt64(0)
                                filled = rd.GetInt64(1)
                                bytes = rd.GetInt64(2)
                            End If
                        End Using
                    End Using
                    Return (Total:=total, Filled:=filled, Bytes:=bytes)
                End Function)

            Dim netMB As Single = r.Bytes / 1024.0F / 1024.0F
            Dim pct As Single = If(r.Total > 0, r.Filled / CSng(r.Total) * 100.0F, 0F)
            Dim avgKB As Single = If(r.Filled > 0, r.Bytes / 1024.0F / r.Filled, 0F)

            _dbg(" ├", $"已回填(非 NULL): {r.Filled:N0} 筆 / mail_simhash 總筆數: {r.Total:N0} 筆 ({pct:F1}%)")
            _dbg(" │", $" bigram_set 淨重 : {netMB.ToString("F2")} MB (僅計已回填的 BLOB，未含未回填的 NULL 列)")
            _dbg(" │", $" 平均每筆大小   : {avgKB.ToString("F2")} KB")
            _dbg("結束", "[bigram_set]")

        Catch ex As System.Exception : _dbg(" ├ 錯誤", $"分析 bigram_set 失敗: {ex.Message}")
        Finally : _lv6StatBusy = False
        End Try
    End Function
    Private Async Function DbVacuumIfNeeded() As Task
        ' 2026/06/16 by Claude Sonnet 4.6: 檢查碎片比例，超過門檻才執行 VACUUM
        ' VACUUM 會重建整個 DB 檔案，執行期間 DB 鎖定，不可做其他讀寫
        ' 使用 5% 門檻：避免每次 RenewCache 後都強制執行（大多數情況碎片很少）
        If _dbCache Is Nothing Then Return
        Try
            Dim pageCount As Long = 0, freelistPages As Long = 0
            Using cmd As New SqliteCommand("PRAGMA page_count", _dbCache) : pageCount = Convert.ToInt64(cmd.ExecuteScalar()) : End Using
            Using cmd As New SqliteCommand("PRAGMA freelist_count", _dbCache) : freelistPages = Convert.ToInt64(cmd.ExecuteScalar()) : End Using

            Dim ratio As Single = If(pageCount > 0, freelistPages / pageCount, 0)
            _dbg($"freelist={freelistPages}/{pageCount} 頁 ({ratio:P1})")
            If ratio > 0.05 Then
                PgrsBar1.Text = "碎片超過 5%，開始執行 VACUUM..." : PgrsBar2.Text = "VACUUM 整理中..."
                _dbg("碎片超過 5%，開始執行 VACUUM...")
                Await Task.Run(Sub()
                                   Using cmd As New SqliteCommand("VACUUM", _dbCache) : cmd.ExecuteNonQuery() : End Using
                               End Sub)
                PgrsBar1.Text = "DB VACUUM 完成"
            Else
                PgrsBar1.Text = "碎片比例低於 5%，略過 VACUUM"
                _dbg("碎片比例低於 5%，略過 VACUUM")
            End If
        Catch ex As Exception
            _dbg("錯誤", ex.Message)
        End Try
    End Function
#End Region

#Region "■ 轉換用輔助函式"
    Private Shared Function UnixSecondsToLocalTime(unixSec As Long) As DateTime
        ''' <summary>
        ''' Unix 秒 (INTEGER) → 本機 DateTime (遇 0 或負數回傳 DateTime.MinValue)
        ''' </summary>
        ' 2026/06/12 by Simon/Claude Opus 4.8: 配合 received_time TEXT→INTEGER 正規化
        ' unixSec ≤ 0 (含 NULL → 0 的情況) 回傳 DateTime.MinValue，避免顯示 1970-01-01
        Return If(unixSec <= 0, DateTime.MinValue, DateTimeOffset.FromUnixTimeSeconds(unixSec).LocalDateTime)
    End Function
    Private Shared Function LocalTimeToUnixSeconds(dt As DateTime) As Long
        ''' <summary>
        ''' 安全轉換 DateTime -> Unix seconds（遇 MinValue 回傳 0）
        ''' </summary>
        ' 2026/06/12 by Simon/Claude Opus 4.8: 修正 DateTimeOffset.ctor 在 Local-kind DateTime
        If dt = DateTime.MinValue Then Return 0L
        Return New DateTimeOffset(DateTime.SpecifyKind(dt, DateTimeKind.Local)).ToUnixTimeSeconds()
    End Function
    Private Function HexStringToByteArray(idStr As String) As Byte()
        ''' <summary>
        ''' 安全將 EntryID 字串轉為 BLOB (Byte Array)。支援常規 Hex 與 EMPTY_ 哨兵字串。
        ''' 2026/06/10 by Gemini/Simon: 優化SQLite儲存空間把Entry_id改成BLOB二進位儲存
        ''' </summary>
        If String.IsNullOrEmpty(idStr) Then Return Array.Empty(Of Byte)()

        ' 檢查是否為系統內建的哨兵字串 (例如 EMPTY_BASIC_ 或 EMPTY_ATT_)
        If idStr.StartsWith("EMPTY_", StringComparison.OrdinalIgnoreCase) Then
            Return System.Text.Encoding.UTF8.GetBytes(idStr)
        Else
            ' .NET 10.0 原生 SIMD 加速轉換
            Return System.Convert.FromHexString(idStr)
        End If
    End Function
    Private Function ByteArrayToHexString(bytes As Byte()) As String
        ''' <summary>
        ''' 安全將 BLOB (Byte Array) 還原為 EntryID 字串。自動判定 UTF-8 哨兵與常規 Hex。
        ''' 2026/06/10 by Gemini/Simon: 優化SQLite儲存空間把Entry_id改成BLOB二進位儲存
        ''' </summary>
        If bytes Is Nothing OrElse bytes.Length = 0 Then Return ""

        ' 檢查前 6 個 Byte 是否為 "EMPTY_" 的 UTF-8 編碼 (69, 77, 80, 84, 89, 95)
        If bytes.Length >= 6 AndAlso
           bytes(0) = 69 AndAlso bytes(1) = 77 AndAlso bytes(2) = 80 AndAlso
           bytes(3) = 84 AndAlso bytes(4) = 89 AndAlso bytes(5) = 95 Then
            Return System.Text.Encoding.UTF8.GetString(bytes)
        Else
            ' .NET 10.0 原生 SIMD 加速轉換
            Return System.Convert.ToHexString(bytes)
        End If
    End Function
    Private Function FolderPathToHash64(fPath As String) As Long
        ''' <summary>
        ''' 計算路徑的 XxHash64，並自動將 (Hash -> 路徑) 的對應關係註冊到記憶體字典中
        ''' </summary>
        Dim h = StringToXxHash64(fPath)
        _dictHashToPath.TryAdd(h, fPath) ' 確保記憶體隨時能反查
        Return h
    End Function
    Private Function StringToXxHash64(ByVal strRaw As String) As Long
        ''' <summary>
        ''' 使用微軟官方內建的高效能 XxHash64 類別
        ''' 將傳入字串計算後轉為 64位元固定雜湊值 (回傳 Long)
        ''' </summary>
        If String.IsNullOrEmpty(strRaw) Then Return 0
        Dim bytes As Byte() = Encoding.UTF8.GetBytes(strRaw)    ' 將字串轉為位元組陣列
        Dim hashBytes As Byte() = XxHash64.Hash(bytes)
        Return BitConverter.ToInt64(hashBytes, 0)               ' 將 8 位元組的雜湊結果轉換為 VB.NET 的 Long (Int64) 類型, 方便直接存入 SQLite 的 INTEGER 欄位
    End Function
    Private Function StringToXxHash64Hex(strRaw As String) As String
        ''' <summary>
        ''' 使用微軟官方內建的高效能 XxHash64 類別
        ''' 將傳入字串計算後轉為 64位元固定雜湊值 (回傳 HEX)
        ''' </summary>
        If String.IsNullOrEmpty(strRaw) Then Return ""
        Return Convert.ToHexString(XxHash64.Hash(Encoding.UTF8.GetBytes(strRaw)))
    End Function
    Private Sub MarkMailFolderDirty(fPath As String)
        ' 2026/07/03 by Simon/Claude Fable 5: 供各 Layer2.5 快取代理層在「③ 剛用 COM/RDO 重新算出資料」的寫入點呼叫，
        '   標記此資料夾需要在下次 SaveCache 時重寫 mail_info/att_maillist/year_count/month_count。
        '   ①②(記憶體命中/DB lazy load)不可呼叫本函式 — 那些資料本來就跟 DB 一致，標記了只會白白重寫。
        If Not String.IsNullOrEmpty(fPath) Then _dirtyMailFolders(fPath) = 0
    End Sub
#End Region

End Class
