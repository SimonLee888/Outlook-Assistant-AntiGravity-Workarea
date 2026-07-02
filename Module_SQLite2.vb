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
'   - CleanupOrphanFolderPath(livePaths)        ' 清除 DB 中已不存在的 folder_path (原 PurgeStaleFolders)，SaveCache 時順帶呼叫
'   - RenewCacheToDB(includeSize As Boolean)    ' RenewCache 按鈕：Phase1~6 完整更新 (2026-04-09 新增) 
'   - RenewAttachMailList(folder, fPath:=fPath) ' 三路比對更新 attach_maillist (2026-04-09 新增) 
'
'   - DbGetFolderStats(fPath)                   ' folder_stats 單行查詢
'   - DbGetMailBasic(fPath)                     ' mail_basic WHERE folder_path=? 全部行
'   - DbGetAttachFilenames(entryId)             ' mail_attachments 單行查詢
'   - DbGetYearCountsForFolder(fPath)           ' year_counts WHERE folder_path=? 全部行
'   - DbGetMonthCountsForFolder(cacheKey)       ' 2026-04-09 新增，cacheKey = FolderPath_year
'   - GetDBSummary() → (fc, mb, at, yc, mc, basic, kb) ' DB 統計摘要 (六張表行數 + 檔案 KB) 
' ---------------------------------------------------------------
'
'   七張表結構 合一個 cache.db (LocalAppData)
'       2026-04-09 新增 month_counts
'       2026-04-22 新增 mailinfo_list
'       2026-06-12 新增 senders；mailinfo_list 移除 topic/sender_email/updated_at，received_time 改 INTEGER
'       folder_stats        (folder_path PK, mail_count, mail_count_all, folder_count, folder_count_all,
'                            folder_size, folder_size_all, pr_count_snap, updated_at)  ← updated_at 僅此表保留
'       year_counts         (folder_hash+year PK, count)
'       month_counts        (folder_hash+year+month PK, count)
'       attach_maillist     (entry_id PK, folder_hash, subject, msg_size, received_time INTEGER, sender_name,
'                            attach_count, pr_count_snap)           ← 專供 Tab3 尋找附件使用
'       attach_filenames    (entry_id PK, folder_hash, filenames TEXT JSON, msg_size)
'       mailinfo_list      (entry_id PK, folder_hash, subject, msg_size, received_time INTEGER, sender_name,
'                            sender_id, msgid_hash, pr_count_snap)  ← 專供 Tab4/Tab5 系列與重複郵件使用
'       senders             (sender_id PK AUTOINCREMENT, sender_email UNIQUE) ← email 正規化，2026-06-12 新增
'                           
' 設計決策 (2026-04-06):
'   1. 跨表 Transaction 保證原子性，一個 Connection 管理最簡單
'   2. 手動控制 (SaveCache / LoadCache 按鈕)，Debug 階段方便目視確認正確性
'      正式版再切換成 Layer2.5 lazy SELECT + 增量寫入
'   3. pr_count_snap 存 _cacheMailCount[path] 的值 (即 PR_CONTENT_COUNT 的讀取結果)
'      Load 後可快速判斷快取是否仍有效，完全不需要呼叫任何 COM
'   4. MailItemInfo 欄位以文字儲存；List(Of String) 附件名稱序列化為 JSON array
'   5. _cacheFolderTree / _cacheSubTreeList 含 COM 物件，永遠不寫入 SQLite
'   6. LoadFolderStatsInner 使用 TryAdd：若記憶體已有值 (Layer2.5 已讀過)，保留記憶體版本
'      若想強制以 DB 為準 (完整重置)，改用直接賦值 _cacheMailCount(path) = ...
'   7. (2026-04-22) 拆分 attach_maillist 與 mailinfo_list：保持 Tab3 與 Tab4/5 邏輯與資料邊界獨立。
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

    ' DB Row 結構 (供 Form1_Outlook.vb 的 Layer2.5 函數使用)
    Friend Class FolderStatsDbRow
        ' folder_stats 一行的讀出結果；-1 代表該欄位在 DB 中為 NULL 或尚未寫入
        Public mc As Long = -1          ' mail_count
        Public mca As Long = -1         ' mail_count_all
        Public fc As Long = -1          ' folder_count
        Public fca As Long = -1         ' folder_count_all
        Public fs As Long = -1          ' folder_size
        Public fsa As Long = -1         ' folder_size_all
        Public snap As Long = -1        ' pr_count_snap (= PR_CONTENT_COUNT at save time)
        Public path As String = ""      ' folder_path        ' by Gemini 3.0 flash, 2026/04/16: 新增路徑標識，供 GetSubtree Tuple 重建使用

        ' by Gemini, 2026/04/10: 新增身分標識與排序標籤，供 TreeView/BFS 持久化優化使用
        Public eid As String = ""       ' entry_id  ' 2026/06/10 by Gemini/Simon: 優化SQLite儲存空間把Entry_id改成BLOB二進位儲存
        Public sid As String = ""       ' store_id
        Public isMail As Integer = -1   ' is_mail (0/1)
        Public hasCh As Integer = -1    ' has_chinese (0/1)
    End Class
    Friend Class AttachMailListDbResult
        ' attach_maillist WHERE folder_path=? 的讀出結果
        Public Snap As Long = -1
        ' 預分配容量為 1024，降低自 SQLite 載入大量郵件快取時的 Resize 開銷 (by Gemini 3 Flash, 2026/05/04)
        Public Mails As New List(Of MailItemInfo)(1024)
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
            _dbCache = New SqliteConnection($"Data Source={_dbCachePath};Mode=ReadWriteCreate;Cache=Shared")
            _dbCache.Open()

            Using cmd As New SqliteCommand("PRAGMA journal_mode=WAL; PRAGMA synchronous=NORMAL;", _dbCache)
                cmd.ExecuteNonQuery()
            End Using
            _dbg("", $"已開啟: {_dbCachePath}")

            Using cmd As New SqliteCommand(BuildSQLiteTableString(), _dbCache)
                cmd.ExecuteNonQuery()
            End Using
            Try
                ' 2026/06/10 by Gemini/Simon: 優化SQLite儲存空間把Entry_id改成BLOB二進位儲存
                ' 2026/06/11 by Gemini/Simon: 把 message_id 轉成 xxHash64，並同時改成 BLOB 儲存節省空間
                Using cmd As New SqliteCommand(
                    "ALTER TABLE folder_stats ADD COLUMN entry_id BLOB;" &
                    "ALTER TABLE folder_stats ADD COLUMN store_id TEXT;" &
                    "ALTER TABLE folder_stats ADD COLUMN is_mail INTEGER DEFAULT -1;" &
                    "ALTER TABLE folder_stats ADD COLUMN has_chinese INTEGER DEFAULT -1;", _dbCache)
                    cmd.ExecuteNonQuery()
                End Using
            Catch ex As System.Exception
                ' 資料行已存在時會拋出例外，安全忽略 (by Gemini, 2026/04/10)
            End Try

            Try ' 2026/05/06 by Claude: mailinfo_list 新增 Tab5 去重欄位
                Using cmd As New SqliteCommand(
                    "ALTER TABLE mailinfo_list ADD COLUMN msgid_hash BLOB;" &
                    "ALTER TABLE mailinfo_list ADD COLUMN sender_email TEXT;", _dbCache)
                    cmd.ExecuteNonQuery()
                End Using
            Catch ex As System.Exception
                ' 欄位已存在，安全略過
            End Try

            ' by Claude Sonnet 4.6, 2026/05/06: Root Cause A 一次性資料清理 migration
            ' by Claude Sonnet 4.6, 2026/06/12: 整段刪除 (原本這段 migration 的唯一目的是清理舊 bug 遺留的污染資料)
            ' 2026/06/12 by Claude: mailinfo_list 欄位重排序 migration
            '   目標順序: entry_id, msg_size, subject, topic, msgid_hash, folder_hash,
            '             sender_name, sender_email, pr_count_snap, received_time, updated_at
            '   用 RENAME → CREATE → INSERT SELECT → DROP 流程，因 SQLite 不支援 ALTER COLUMN ORDER

            _dbg("", "資料表確認完成")

            LoadSendersInner()  ' 2026/06/12 by Simon/Claude Opus 4.8: 載入 senders 表，供 DbGetMailInfo lazy load 時能解析 sender_id
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

        ' 1. folder_stats: 資料夾狀態統計 (核心快取層)
        sb.AppendLine("
                        CREATE TABLE IF NOT EXISTS folder_stats ( 
                            folder_path         TEXT    PRIMARY KEY,
                            mail_count          INTEGER,
                            mail_count_all      INTEGER,
                            folder_count        INTEGER,
                            folder_count_all    INTEGER,
                            folder_size         INTEGER,
                            folder_size_all     INTEGER,
                            pr_count_snap       INTEGER,
                            entry_id            BLOB,
                            store_id            TEXT,
                            is_mail             INTEGER,
                            has_chinese         INTEGER,
                            updated_at          TEXT
                        );")

        ' 2. attach_maillist: 附件郵件清單 (專供 Tab3 尋找附件使用)
        ' 2026/06/12 by Simon/Claude Opus 4.8: received_time TEXT→INTEGER (Unix秒)；移除 updated_at (只寫不讀，無用)
        sb.AppendLine("
                        CREATE TABLE IF NOT EXISTS attach_maillist (
                            entry_id        BLOB    PRIMARY KEY,
                            folder_hash     INTEGER NOT NULL,
                            subject         TEXT,
                            msg_size        INTEGER,
                            received_time   INTEGER,
                            sender_name     TEXT,
                            attach_count    INTEGER,
                            pr_count_snap   INTEGER
                        );
                        CREATE INDEX IF NOT EXISTS idx_mb_folder ON attach_maillist(folder_hash);")

        ' 3. mailinfo_list: 基礎郵件清單 (專供 Tab4/Tab5 與重複郵件比對使用)
        ' 2026/06/12 by Simon/Claude Opus 4.8: 移除 topic (改由 GetCleanSubject(subject) 動態計算)
        '   移除 sender_email (改以 sender_id 外鍵指向 senders 表，節省重複儲存)
        '   received_time TEXT→INTEGER (Unix秒)；移除 updated_at (只寫不讀)
        sb.AppendLine("
                        CREATE TABLE IF NOT EXISTS mailinfo_list (
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
                        CREATE INDEX IF NOT EXISTS idx_mailinfo_folder ON mailinfo_list(folder_hash);")

        '' 4. attach_filenames: 附件名稱清單
        '' 2026/06/12 by Simon/Claude Opus 4.8: 移除 updated_at (只寫不讀)
        ' 4. attach_filenames: 已於 2026/06/21 by Simon/Claude 搬至 OLAcacheMail.db(_dbMail)，改在 InitDbMail 建表。
        '    理由：逐封開信枚舉附件、重建極貴，需隨該檔在 ZipAndRebuildDB 後存活（與 mail_simhash 同策略）。
        'sb.AppendLine("
        '                CREATE TABLE IF NOT EXISTS attach_filenames (
        '                    entry_id        BLOB    PRIMARY KEY,
        '                    folder_hash     INTEGER NOT NULL,
        '                    filenames       TEXT,
        '                    msg_size        INTEGER
        '                );
        '                CREATE INDEX IF NOT EXISTS idx_ma_folder ON attach_filenames(folder_hash);")

        ' 5. year_counts: 年份統計
        ' 2026/06/12 by Simon/Claude Opus 4.8: 移除 updated_at (只寫不讀)
        sb.AppendLine("
                        CREATE TABLE IF NOT EXISTS year_counts (
                            folder_hash     INTEGER NOT NULL,
                            year            INTEGER NOT NULL,
                            count           INTEGER NOT NULL,
                            PRIMARY KEY (folder_hash, year)
                        );
                        CREATE INDEX IF NOT EXISTS idx_yc_folder ON year_counts(folder_hash);")

        ' 6. month_counts: 月份統計
        ' 2026/06/12 by Simon/Claude Opus 4.8: 移除 updated_at (只寫不讀)
        sb.AppendLine("
                        CREATE TABLE IF NOT EXISTS month_counts (
                            folder_hash INTEGER NOT NULL,
                            year        INTEGER NOT NULL,
                            month       INTEGER NOT NULL,
                            count       INTEGER NOT NULL,
                            PRIMARY KEY (folder_hash, year, month)
                        );")

        ' 7. senders: 寄件者 email 正規化表 (2026/06/12 by Simon/Claude Opus 4.8 新增)
        '   只存不重複的 sender_email；mailinfo_list 透過 sender_id 外鍵參照
        sb.AppendLine("
                        CREATE TABLE IF NOT EXISTS senders (
                            sender_id       INTEGER PRIMARY KEY AUTOINCREMENT,
                            sender_email    TEXT    UNIQUE NOT NULL
                        );")

        ' 注意：month_counts 的舊版 schema 遷移 (cache_key → 三欄 PK) 在 InitDatabase() 中一次性處理，
        ' 不在此處 DROP TABLE，避免每次啟動都清空已存資料。

        Return sb.ToString()

    End Function
    Private Function GetDBSummary() As (fc As Integer, mb As Integer, at As Integer, yc As Integer, mc As Integer, basic As Integer, senders As Integer, kb As Long, lastTs As String, kbMail As Long, sh As Integer)
        ' ---------------------------------------------------------------
        ' GetDBSummary — 取得 DB 統計摘要，供按鈕顯示
        ' 回傳 (folder_stats, attach_maillist, attach_filenames, year_counts, month_counts, mailinfo_list, senders, KB, lastTs)
        ' 2026/04/09 新增 mc = month_counts 行數
        ' 2026/04/10 新增 lastTs = 最後 updated_at 時間
        ' 2026/04/22 by Gemini 3 Flash: 新增 basic = mailinfo_list 行數
        ' 2026/06/14 by Simon/Claude Opus 4.8: 新增 senders = senders 行數 (供 Tab6 Lv6 顯示)
        ' 2026/06/21 by Simon/Claude Opus 4.8: attach_filenames 已搬至 _dbMail(OLAcacheMail.db)
        ' ---------------------------------------------------------------
        If _dbCache Is Nothing Then Return (0, 0, 0, 0, 0, 0, 0, 0L, "N/A", 0L, 0)

        Try
            Dim fc, mb, at, yc, mcount, basicCount, sendersCount As Integer : Dim lastTs As String = "N/A"
            Using cmd As New SqliteCommand("SELECT COUNT(*) FROM folder_stats", _dbCache) : fc = Convert.ToInt32(cmd.ExecuteScalar()) : End Using
            Using cmd As New SqliteCommand("SELECT COUNT(*) FROM mailinfo_list", _dbCache) : basicCount = Convert.ToInt32(cmd.ExecuteScalar()) : End Using
            Using cmd As New SqliteCommand("SELECT COUNT(*) FROM senders", _dbCache) : sendersCount = Convert.ToInt32(cmd.ExecuteScalar()) : End Using ' 2026/06/14 by Simon/Claude Opus 4.8
            Using cmd As New SqliteCommand("SELECT COUNT(*) FROM year_counts", _dbCache) : yc = Convert.ToInt32(cmd.ExecuteScalar()) : End Using
            Using cmd As New SqliteCommand("SELECT COUNT(*) FROM month_counts", _dbCache) : mcount = Convert.ToInt32(cmd.ExecuteScalar()) : End Using
            Using cmd As New SqliteCommand("SELECT COUNT(*) FROM attach_maillist", _dbCache) : mb = Convert.ToInt32(cmd.ExecuteScalar()) : End Using

            ' 2026/06/21 by Simon/Claude Opus 4.8: attach_filenames 已搬至 _dbMail(OLAcacheMail.db)，COUNT 改打 _dbMail；_dbMail 為 Nothing 時 at=0
            If _dbMail IsNot Nothing Then Using cmd As New SqliteCommand("SELECT COUNT(*) FROM attach_filenames", _dbMail) : at = Convert.ToInt32(cmd.ExecuteScalar()) : End Using

            ' 2026/06/21 by Simon/Claude: mail_simhash 筆數同樣讀 _dbMail (供 Tab6 Lv6 雙檔顯示)
            Dim sh As Integer = 0
            If _dbMail IsNot Nothing Then Using cmd As New SqliteCommand("SELECT COUNT(*) FROM mail_simhash", _dbMail) : sh = Convert.ToInt32(cmd.ExecuteScalar()) : End Using

            ' 抓取最後一次成功的儲存時間字串 (取最大的 updated_at)
            Using cmd As New SqliteCommand("SELECT MAX(updated_at) FROM folder_stats", _dbCache)
                Dim val = cmd.ExecuteScalar()
                If val IsNot Nothing AndAlso Not IsDBNull(val) Then lastTs = val.ToString()
            End Using

            ' 2026/06/21 by Simon/Claude: OLAcacheMail.db 檔案大小(KB)；_dbMailPath 空或檔不存在時為 0
            Dim fi As New IO.FileInfo(_dbCachePath)
            Dim kbMail As Long = 0L : If Not String.IsNullOrEmpty(_dbMailPath) AndAlso IO.File.Exists(_dbMailPath) Then kbMail = New IO.FileInfo(_dbMailPath).Length \ 1024L
            Return (fc, mb, at, yc, mcount, basicCount, sendersCount, If(fi.Exists, fi.Length \ 1024L, 0L), lastTs, kbMail, sh)

        Catch ex As System.Exception
            _dbg("       ├ 錯誤", ex.Message) ' by Gemini, 2026/04/11: Level 3
            Return (0, 0, 0, 0, 0, 0, 0, 0L, "Err", 0L, 0)
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
                        Dim entry = archive.CreateEntry("OLAcache.db", System.IO.Compression.CompressionLevel.SmallestSize)

                        Using entryStream = entry.Open()
                            ' 加上 FileShare.ReadWrite 容許其他可能卡住的唯讀鎖，防止 IOException (by Gemini, 2026/04/10)
                            Using fs As New System.IO.FileStream(_dbCachePath, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite)
                                fs.CopyTo(entryStream)
                            End Using
                        End Using
                    End Using
                End Using

                IO.File.Delete(_dbCachePath) ' 壓縮完後刪除原始 db 檔
            End If

            ' 3. 重新建立資料庫與表格
            InitDatabase()
            _dbg(" ├ 結束", "SSD 快取已重設，舊檔案已 Zip 備份") ' by Gemini, 2026/04/11: 修正對應開始層級 Level 1
        Catch ex As System.Exception
            _dbg("       ├ 錯誤", $"無法重置 SSD 資料庫: {ex.Message}")
            Throw
        End Try
    End Function
    Private Function SanitizeProfileName(name As String) As String
        ' Profile 名稱安全過濾
        ' CurrentProfileName 可能包含空格、單引號（Simon'st Mail）、甚至斜線等非法路徑字元
        Dim invalid As Char() = IO.Path.GetInvalidFileNameChars()
        Return New String(name.Select(Function(c) If(Array.IndexOf(invalid, c) >= 0, "_"c, c)).ToArray())
    End Function

    Private Sub InitDbMail()
        ' 開啟/建立 OLAcacheMail.db (與 OLAcache.db 同目錄、不同檔)。在 InitDatabase 末段呼叫。
        ' 2026/06/21 by Simon/Claude Opus 4.8: 原 OLAsimhash.db 改名 OLAcacheMail.db；本檔現含 mail_simhash + attach_filenames 兩張「逐封讀取極貴」的快取表
        ' 2026/06/21 by Simon/Claude Opus 4.8: 本檔已改名 OLAcacheMail.db，並納入 attach_filenames(逐封開信重建極貴)，同享「rebuild 後存活」性質。
        '   注意：rebuild 不清本檔，但 RenewCache 狀況 A(內容有變) 與孤兒清理仍會精確 purge 本檔對應 folder 的 attach_filenames 列，避免死列殘留。
        Try
            _dbMailPath = IO.Path.Combine(IO.Path.GetDirectoryName(_dbCachePath), "OLAcacheMail.db")
            _dbMail = New SqliteConnection($"Data Source={_dbMailPath};Mode=ReadWriteCreate;Cache=Shared")
            _dbMail.Open()
            Using cmd As New SqliteCommand("PRAGMA journal_mode=WAL; PRAGMA synchronous=NORMAL;", _dbMail) : cmd.ExecuteNonQuery() : End Using
            Using cmd As New SqliteCommand("CREATE TABLE IF NOT EXISTS mail_simhash (entry_id BLOB PRIMARY KEY, simhash INTEGER NOT NULL, bigram_count INTEGER NOT NULL);", _dbMail)
                cmd.ExecuteNonQuery()
            End Using

            ' 2026/06/21 by Simon/Claude Opus 4.8: attach_filenames 由 OLAcache.db 搬入本檔(schema 原樣保留：entry_id BLOB PK / folder_hash / filenames / msg_size)
            Using cmd As New SqliteCommand("CREATE TABLE IF NOT EXISTS attach_filenames (entry_id BLOB PRIMARY KEY, folder_hash INTEGER NOT NULL, filenames TEXT, msg_size INTEGER);" &
                                           "CREATE INDEX IF NOT EXISTS idx_ma_folder ON attach_filenames(folder_hash);", _dbMail)
                cmd.ExecuteNonQuery()
            End Using
            _dbg("", $"已開啟 OLAcacheMail db: {_dbMailPath}")

        Catch ex As System.Exception
            _dbg("       ├ 錯誤", ex.Message) : _dbMail = Nothing   ' 出錯設 Nothing，後續 sim 讀寫自動跳過 (同主 db 容錯策略)
        End Try
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
                    cmd.Parameters.Add("@eid", SqliteType.Blob) : cmd.Parameters.Add("@sh", SqliteType.Integer) : cmd.Parameters.Add("@bc", SqliteType.Integer)
                    For Each row In rows
                        cmd.Parameters("@eid").Value = HexStringToByteArray(row.EntryID)
                        cmd.Parameters("@sh").Value = row.SimHash : cmd.Parameters("@bc").Value = row.BigramCount
                        cmd.ExecuteNonQuery()
                    Next
                End Using
                txn.Commit()
            End Using
        Catch ex As System.Exception
            _dbg("       ├ 錯誤", ex.Message)
        End Try
    End Sub
    Private Sub DeleteDbMail()
        ' 供「清快取」對話框那顆 checkbox 勾選時呼叫：關閉連線 → 刪檔。(預設不呼叫；使用者主動勾選才清)
        ' 2026/06/21 by Simon/Claude Opus 4.8: 本檔(OLAcacheMail.db)現含 attach_filenames，整檔刪除會一併清掉 → 須同步清 _cacheAttachFilename
        Try
            If _dbMail IsNot Nothing Then _dbMail.Close() : _dbMail.Dispose() : _dbMail = Nothing
            SqliteConnection.ClearAllPools()

            If Not String.IsNullOrEmpty(_dbMailPath) AndAlso IO.File.Exists(_dbMailPath) Then IO.File.Delete(_dbMailPath)
            _cacheSimHash.Clear() : _cacheAttachFilename.Clear() : _simHashLoaded = False

            InitDbMail()   ' 重建空表，後續仍可重新累積

        Catch ex As System.Exception
            _dbg("       ├ 錯誤", ex.Message)
        End Try
    End Sub
#End Region

#Region "■ 快取主控流程 (High-Level Cache Controllers)"
    Private Async Function SaveCachesToDB() As Task
        ' ---------------------------------------------------------------
        ' SaveCachesToDB — 把記憶體快取全部存入 SQLite
        ' 對應 Setting 頁 SaveCache 按鈕
        ' 流程: ① CleanupOrphanFolderPath (先清孤兒) → ② 批次寫入三張表 → ③ 統計顯示
        ' ---------------------------------------------------------------
        _dbg(" ├ 開始") ' by Gemini, 2026/04/10: 調整縮排層級為 Level 1
        If _dbCache Is Nothing Then _dbg("", "DB 未初始化") : Return

        Dim sw As Diagnostics.Stopwatch = Diagnostics.Stopwatch.StartNew()  ' by Gemini 3.5 Flash, 2026/06/07
        Dim savedFolders, savedAttachMailList, savedAttachFilenames, savedBasic As Integer
        Try
            PgrsBar1.Text = "正在存入快取..." : Cursor = Cursors.WaitCursor

            ' ① 先清孤兒：收集目前記憶體快取中所有仍存在的 folder_path，清除 DB 中已不存在的行
            ' 用記憶體快取的 key 聯集代表「目前已知 live 的資料夾」 (比重新 BFS 掃 COM 快得多) 
            Dim livePaths As New HashSet(Of String)(1024)
            For Each k In _cacheMailCount.Keys : livePaths.Add(k) : Next
            For Each k In _cacheFolderCount.Keys : livePaths.Add(k) : Next
            For Each k In _cacheAttachMailList.Keys : livePaths.Add(k) : Next

            ' 2026/06/12 by Simon/Claude Opus 4.8: lazy-load 安全保護
            ' _cacheMailCount 等字典因 lazy-load 在重啟後可能不完整（記憶體中看不到 ≠ Outlook 中已刪除）
            ' 把 folder_stats 現有路徑全部列為 live，確保 CleanupOrphanPath 不誤刪仍存在的資料夾
            ' 真正的孤兒清理由 RenewCacheToDB（完整 COM BFS 逐一 GetFolderFromID 確認）負責
            Using readCmd As New SqliteCommand("SELECT folder_path FROM folder_stats", _dbCache)
                Using reader = readCmd.ExecuteReader()
                    While reader.Read() : livePaths.Add(reader.GetString(0)) : End While
                End Using
            End Using

            If livePaths.Count > 0 Then Await CleanupOrphanPath(livePaths)
            ' ② SQLite I/O 在背景執行緒，不阻塞 UI
            Dim r = Await Task.Run(Function()
                                       Using txn As SqliteTransaction = _dbCache.BeginTransaction()
                                           Try
                                               Dim f = SaveFolderStatsInner(txn)
                                               Dim b = SaveAttachMailListInner(txn)
                                               'Dim a = SaveAttachFilenamesInner(txn)
                                               Dim y = SaveYearCountsInner(txn)
                                               Dim m = SaveMonthCountsInner(txn)    ' 2026/04/09 新增
                                               Dim s = SaveSendersInner(txn)        ' 2026/06/12 by Simon/Claude Opus 4.8: 先建立 senders 表，供 SaveMailInfoInner 查 sender_id
                                               Dim basic = SaveMailInfoInner(txn)
                                               txn.Commit()

                                               Dim a = SaveAttachFilenamesInner()   ' 2026/06/21 by Simon/Claude Opus 4.8: attach_filenames 已搬至 OLAcacheMail.db(_dbMail)，跨檔不掛 _dbCache txn，改自管獨立交易(主 db 已 commit 後再寫)
                                               Return (f, b, a, y, m, s, basic)
                                           Catch ex As System.Exception
                                               txn.Rollback() : Throw
                                           End Try
                                       End Using
                                   End Function)

            savedFolders = r.f : savedAttachMailList = r.b : savedAttachFilenames = r.a
            Dim savedYears As Integer = r.y, savedMonths As Integer = r.m
            Dim savedSenders As Integer = r.s : savedBasic = r.basic
            sw.Stop()

            ' ③ 統計：各快取字典目前的 entry 數
            Dim st = GetDBSummary()
            Dim statLine1 = $"① [記憶體] MailCount: {_cacheMailCount.Count} / MailCountAll: {_cacheMailCountAll.Count} / FolderCount: {_cacheFolderCount.Count} / FolderCountAll: {_cacheFolderCountAll.Count}"
            Dim statLine2 = $"② [記憶體] FolderSize: {_cacheFolderSize.Count} / FolderSizeAll: {_cacheFolderSizeAll.Count} / AttachPreScan: {_cacheAttachMailList.Count} / AttachFilename: {_cacheAttachFilename.Count}"
            Dim statLine3 = $"③ [寫入DB] folder_stats: {savedFolders} 筆 / mailinfo_list: {savedBasic} 筆 / attach_maillist: {savedAttachMailList} 筆 / attach_filenames: {savedAttachFilenames} 筆 / year_counts: {savedYears} 筆 / month_counts: {savedMonths} 筆 / 耗時: {sw.Elapsed.TotalSeconds:0.000} 秒"
            Dim statLine4 = $"④ [DB現況] folder_stats: {st.fc} 筆 / attach_maillist: {st.mb} 筆 / attach_filenames: {st.at} 筆 / year_counts: {st.yc} 筆 / month_counts: {st.mc} 筆 / 檔案: {st.kb} KB"

            PgrsBar1.Text = $"SaveCache 完成 — 耗時: {sw.Elapsed.TotalSeconds:0.000} 秒"
            PgrsBar2.Text = statLine4
            _dbg(" ├ ", statLine1)
            _dbg(" ├ ", statLine2)
            _dbg(" ├ ", statLine3)
            _dbg(" ├ ", statLine4)

        Catch ex As System.Exception
            PgrsBar1.Text = "SaveCache 失敗"
            _dbg("       ├ 錯誤", ex.Message)
        Finally
            Cursor = Cursors.Default
            _dbg(" ├ 結束") ' by Gemini, 2026/04/11: 修正對應開始層級 Level 1
        End Try

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
            Dim beforePS = _cacheAttachMailList.Count
            Dim beforeAF = _cacheAttachFilename.Count

            Dim r = Await Task.Run(Function()
                                       Dim f = LoadFolderStatsInner()
                                       Dim b = LoadAttachMailListInner()
                                       Dim a = LoadAttachFilenamesInner()
                                       Dim y = LoadYearCountsInner()
                                       Dim m = LoadMonthCountsInner()       ' 2026/04/09 新增
                                       Dim s = LoadSendersInner()           ' 2026/06/12 by Simon/Claude Opus 4.8: 先載入 senders 字典，供 LoadMailInfoInner 解析 sender_id
                                       Dim basic = LoadMailInfoInner() ' 2026/04/22 by Gemini 3.1 Pro 新增
                                       Return (f, b, a, y, m, s, basic)
                                   End Function)
            sw.Stop()

            ' 詳細 _dbg：各快取字典 Load 後的增量
            Dim statLine1 = $"① [folder_stats] 讀入 {r.f} 筆 — " &
                            $"MailCount +{_cacheMailCount.Count - beforeMC} / " &
                            $"MailCountAll +{_cacheMailCountAll.Count - beforeMCA} / " &
                            $"FolderCount +{_cacheFolderCount.Count - beforeFC} / " &
                            $"FolderCountAll +{_cacheFolderCountAll.Count - beforeFCA}"
            Dim statLine2 = $"② [folder_stats cont.] " &
                            $"FolderSize +{_cacheFolderSize.Count - beforeFS} / " &
                            $"FolderSizeAll +{_cacheFolderSizeAll.Count - beforeFSA}"
            Dim statLine3 = $"③ [attach_maillist] 讀入 {r.b} 筆 → AttachPreScan +{_cacheAttachMailList.Count - beforePS} 個資料夾"
            Dim statLine4 = $"④ [attach_filenames] 讀入 {r.a} 筆 → AttachFilename +{_cacheAttachFilename.Count - beforeAF} 筆"
            Dim statLine_yc = $"⑤ [year_counts] 讀入 {r.y} 筆 → _yearCountsCache {_cacheYearCounts.Count} 個資料夾"
            Dim statLine_mc = $"⑥ [month_counts] 讀入 {r.m} 筆 → _monthCountsCache {_cacheMonthCounts.Count} 個 cache_key"
            Dim statLine_basic = $"⑦ [mailinfo_list] 讀入 {r.basic} 筆 → BasicPreScan {_cacheMailInfo.Count} 個資料夾" ' 2026/04/22 by Gemini 3.1 Pro
            Dim st = GetDBSummary()
            Dim statLine5 = $"⑧ [DB現況] folder_stats: {st.fc} 筆 / mailinfo_list: {st.basic} 筆 / attach_maillist: {st.mb} 筆 / attach_filenames: {st.at} 筆 / year_counts: {st.yc} 筆 / month_counts: {st.mc} 筆 / 檔案: {st.kb} KB / 耗時: {sw.Elapsed.TotalSeconds:0.000} 秒" ' 2026/04/22 by Gemini 3.1 Pro: 加入 basic 統計

            PgrsBar1.Text = $"LoadCache 完成 — DB: {st.fc}/{st.basic}/{st.mb}/{st.at}/{st.yc}/{st.mc} 筆，{st.kb} KB，耗時 {sw.Elapsed.TotalSeconds:0.000} 秒" ' 2026/04/22 by Gemini 3.1 Pro
            PgrsBar2.Text = $"記憶體增量 — mailCount+{_cacheMailCount.Count - beforeMC} / attachFilename+{_cacheAttachFilename.Count - beforeAF} / basicMailInfo:{_cacheMailInfo.Count} 資料夾" ' 2026/04/22 by Gemini 3.1 Pro
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
    Private Async Function RenewCacheToDB() As Task
        ' ---------------------------------------------------------------
        ' RenewCacheToDB — 完整更新 DB 快取 (對應 Setting 頁 RenewCache 按鈕) 
        '
        ' 與 SaveCachesToDB 的差異：
        '   SaveCache  = 把目前記憶體快取照單全收寫入 DB (不更新過期的值) 
        '   RenewCache = 先用 COM 比對 snapshot → 只對有變動的資料夾重新計算 → 寫入 DB
        '
        ' 流程：
        '   Phase 1. BFS 掃出所有 live folders (COM，~1ms/資料夾) 
        '   Phase 2. 每個 folder 讀 GetLiveFolderSnapOOM vs DB snapshot → 找 dirty folders
        '   Phase 3. 對每個 dirty folder 重新計算：
        '              mc/fc (快，~1ms) 
        '              year_counts (GetTable + GetArray，~10-50ms/資料夾) 
        '              month_counts (清記憶體， Phase5 清 DB， 展開時 lazy 重算) 
        '              attach_maillist (GetTable 三路比對，~5ms/資料夾) 
        '              folder_size (選擇性，依 includeSize，GetTable 遍歷，~10-30s/大資料夾) 
        '              清除 mca/fca/fsa 聚合快取 (讓下次點選重算) 
        '              清除此 folder 的 month_counts 記憶體快取 (不重算，展開年份時 lazy) 
        '   Phase 4. 清除所有 dirty folders 的 ancestor 聚合快取
        '   Phase 5. 批次 DELETE dirty folders 的 month_counts DB rows (不是孤兒，不靠 Cleanup) 
        '   Phase 6. CleanupOrphanFolderPath → SaveCachesToDB
        '
        ' 不更新項目 (設計邊界) ：
        '   attach_filenames — 最耗時，留給使用者搜尋附件時 lazy 觸發
        '   month_counts     — 清記憶體 + 清 DB，展開年份時 lazy 重算
        ' 2026/04/09 by Claude
        ' ---------------------------------------------------------------
        ' 2026/04/16 by Simon/Claude: 加入 cToken (OkayNowYouHaveToken)，取代 _cancelRequested + GoTo Cancelled 模式
        '   Phase1 改用 Dictionary(Of String, Outlook.Folder) liveDict，每個資料夾只讀一次 FolderPath COM 屬性，
        '   Phase2/3/4 迭代 dict 的 Key/Value，完全省去重複的 folder.FolderPath COM 呼叫（~500 資料夾省 ~250ms）
        '   Phase2/3 節流改用 SmartThrottle(sw, cToken, ThrottleFreq.Low)，取代 Mod N + Task.Delay(1)
        '   GetYearCountsForFolderL3 / GetFolderSizeOOM 補入 cToken:=cToken
        ' ---------------------------------------------------------------
        ' 2026/05/17 by simon/Gemini: RenewCacheToDB 大幅重構，改為「精確打擊」模式，
        '   不再 BFS 展開每個資料夾的子樹來找對應的 DB row，而是直接從 DB 撈出全部資料夾清單，然後用 GetFolderFromID 精確抓出 COM 物件，比對 snapshot 決定是否 dirty
        ' 2026/6/7: by simon/Gemini: 去除原本函數內的多段計時和狀態顯示, 直接在RenewCache_Click事件中計時顯示整體耗時
        ' ---------------------------------------------------------------
        Dim cToken As Threading.CancellationToken = OkayNowYouHaveToken()
        If _dbCache Is Nothing Then Return

        _dbg("開始")
        Try
            ' 1. 【Rule 1】由專用函數撈出全部追蹤名單，維持主流程乾淨
            Dim dbList = DbGetAllFolderStats()
            If dbList.Count = 0 Then _dbg("RenewCache", "資料庫無快取紀錄，略過") : Return

            ' 建立您原本 CleanupOrphanPath 所需的活著名單
            Dim liveFolderPaths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim updatedPaths As New HashSet(Of String)() ' 用於 Rule 6：記錄需要失效父子聚合值的路徑

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
                    _cacheIsMailFolder.TryRemove(fPath, Nothing)
                    _cacheMailCount.TryRemove(fPath, Nothing)
                    _cacheFolderCount.TryRemove(fPath, Nothing)
                    _cacheFolderSize.TryRemove(fPath, Nothing)
                    _cacheAttachMailList.TryRemove(fPath, Nothing)
                    _cacheYearCounts.TryRemove(fPath, Nothing)
                    For Each mk In _cacheMonthCounts.Keys.Where(Function(k) k.StartsWith(fPath & "_")).ToList()
                        _cacheMonthCounts.TryRemove(mk, Nothing)
                    Next

                    updatedPaths.Add(fPath) ' 標記這條路徑不見了，上層父資料夾的聚合值需要更新
                    Continue For            ' 死了就不加入 liveFolderPaths，等一下交給您的 CleanupOrphanPath 去刪 DB
                End If

                ' 4. 如果資料夾在 Outlook 還活著，加入活著名單集合
                liveFolderPaths.Add(fPath)

                ' 5. 【Rule 4 & 2】比對 Snap 與更新邏輯
                Dim liveSnap = GetLiveFolderSnapOOM(folder, fPath)
                If liveSnap <> row.snap Then
                    ' 狀況 A：Snap 不一致！代表 Outlook 有變動，進行 Layer 3 COM 讀取更新記憶體
                    _cacheMailCount(fPath) = GetMailCount(folder, fPath, skipCache:=True)       ' 2026/06/23 by Simon/Claude: 狀況A snap 重讀改走 proxy skipCache(RDO 派工)
                    _cacheFolderCount(fPath) = GetFolderCount(folder, fPath, skipCache:=True)   ' 2026/06/23 by Simon/Claude: 同上
                    _cacheFolderSize(fPath) = Await GetFolderSize(folder, fPath:=fPath, skipCache:=True, cToken:=cToken)    ' 2026/6/27 by simon/Claude Opus 4.8: 整合GetFolderSize單一入口再分派RDO/OOM, 加skipCache參數讓DB 重建的強制重讀也吃得到 GetFolderSizeRdo的提速

                    ' 2026/06/22 by Simon/Claude: 缺口1+2 ② Surgical 嚴格清除 —
                    '   (a) 取 live 全郵件 entryID，算「DB 有、live 沒有」的已刪集合 → surgical 清兩張昂貴快取
                    '       (attach_filenames/mail_simhash 之 DB 列 + 記憶體)，存活郵件的昂貴快取保留免重讀。
                    '   (b) 便宜逐封表(basic/attach_maillist/month_counts)整夾 nuke DB 死列；對應記憶體一併清/重建，
                    '       否則尾端 SaveCachesToDB 會把舊鬼魂列寫回，使 nuke 失效。
                    Dim liveAll = Await GetMailInfoOOM(folder, needTopic:=False, cToken:=cToken, fPath:=fPath)
                    Dim liveSet As New HashSet(Of String)(liveAll.Select(Function(m) m.Mail.EntryID))
                    Dim absent = DbGetFolderEntryIds(fPath).Where(Function(e) Not liveSet.Contains(e)).ToList()
                    If absent.Count > 0 Then SimDbDeleteMailRowsByEntryIds(absent, includeAttachFilenames:=True)
                    DbPurgeFolderMailRows(fPath, includeAttachFilenames:=False) ' 整夾 nuke 便宜表 DB(含 EMPTY_BASIC 哨兵)
                    _cacheAttachMailList.TryRemove(fPath, Nothing)              ' 配合 DB nuke，避免 SaveCache 寫回舊 attach_maillist
                    For Each mk In _cacheMonthCounts.Keys.Where(Function(k) k.StartsWith(fPath & "_")).ToList()
                        _cacheMonthCounts.TryRemove(mk, Nothing)                ' 月計數同步失效，展開年份時 lazy 重算
                    Next

                    _cacheYearCounts.TryRemove(fPath, Nothing)
                    _cacheMailInfo(fPath) = (liveAll, liveSnap)            ' 既已掃描就存回(取代原 TryRemove)，SaveCache 以新 snap 重寫 mailinfo_list，省一次 lazy 重掃
                    _cacheFolderIDs(fPath) = (folder.EntryID, folder.StoreID, IsMailFolder(folder, fPath), True)
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
                End If
                processed += 1
                Await SmartThrottle(swThrottle, cToken, ThrottleFreq.Low, Sub() PgrsBar2.Text = $"對帳中 {processed}/{dbList.Count}...")
            Next

            ' 6. 【Rule 3】安全無縫套用：直接呼叫您原本寫好的 CleanupOrphanPath 清理 5 個資料表
            _dbg("清理孤兒資料夾路徑...")
            Await CleanupOrphanPath(liveFolderPaths)

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
            Await SaveCachesToDB()

        Catch ex As OperationCanceledException
            PgrsBar1.Text = "RenewCache 已由使用者中斷"
        Catch ex As System.Exception
            PgrsBar1.Text = $"RenewCache 失敗: {ex.Message}"
            _dbg("RenewCache 發生錯誤", ex.Message)
        Finally
            Cursor = Cursors.Default
            _dbg("結束")
        End Try
    End Function
    Private Async Function CleanupOrphanPath(liveFolderPaths As HashSet(Of String)) As Task
        ' ---------------------------------------------------------------
        ' CleanupOrphanPath — 刪除已不存在於 Outlook 的資料夾孤兒行 (改為非同步 by Gemini 3.1 Pro, 2026/05/05)
        ' liveFolderPaths = 目前仍有效的資料夾路徑集合
        '   呼叫來源 A: SaveCachesToDB 開頭 (用記憶體快取 key 聯集) 
        '   呼叫來源 B: RenewCache_Click (用 GetSubtree BFS 掃 COM 取得完整清單) 
        ' ---------------------------------------------------------------
        _dbg("    ├ 開始", $"live 資料夾數: {liveFolderPaths.Count}") ' by Gemini, 2026/04/10: 調整縮排層級為 Level 2
        If _dbCache Is Nothing Then Return

        Await Task.Run(Sub()
                           Try
                               ' 讀出 DB 中所有 folder_path
                               Dim dbPaths As New List(Of String)(2048)
                               Using cmd As New SqliteCommand("SELECT folder_path FROM folder_stats", _dbCache)
                                   Using reader = cmd.ExecuteReader()
                                       While reader.Read() : dbPaths.Add(reader.GetString(0)) : End While
                                   End Using
                               End Using
                               _dbg("", $"DB 中有 {dbPaths.Count} 個資料夾路徑")

                               Dim stalePaths = dbPaths.Where(Function(p) Not liveFolderPaths.Contains(p)).ToList()
                               If stalePaths.Count = 0 Then _dbg("", "未發現孤兒快取，略過") : Return

                               ' 2026/06/22 by Simon/Claude: ② Surgical — 整夾消失 → 該夾全部 entryID 皆失效。
                               '   先撈 mailinfo_list 的 entryID(務必在下方 DELETE mailinfo_list 之前)，供稍後清 mail_simhash
                               '   (無 folder_hash 只能靠 entryID) 與記憶體 _cacheSimHash/_cacheAttachFilename。
                               Dim orphanEntryIds As New List(Of String)()
                               For Each s In stalePaths : orphanEntryIds.AddRange(DbGetFolderEntryIds(s)) : Next

                               Dim dF, dB, dA, dM, dBasic, dSh As Integer
                               Using txn As SqliteTransaction = _dbCache.BeginTransaction()
                                   Using c1 As New SqliteCommand("DELETE FROM folder_stats WHERE folder_path=@fp", _dbCache, txn),
                                       c2 As New SqliteCommand("DELETE FROM attach_maillist WHERE folder_hash=@fh", _dbCache, txn),
                                       c4 As New SqliteCommand("DELETE FROM month_counts WHERE folder_hash=@fh", _dbCache, txn),
                                       c5 As New SqliteCommand("DELETE FROM mailinfo_list WHERE folder_hash=@fh", _dbCache, txn)
                                       ' 2026/06/21 by Simon/Claude: attach_filenames 已搬至 OLAcacheMail.db(_dbMail)，跨檔獨立刪除(原 c3)，dA 由此累計
                                       ' c3 As New SqliteCommand("DELETE FROM attach_filenames WHERE folder_hash=@fh", _dbCache, txn),

                                       c1.Parameters.Add("@fp", SqliteType.Text)
                                       c2.Parameters.Add("@fh", SqliteType.Integer) ': c3.Parameters.Add("@fh", SqliteType.Integer)
                                       c4.Parameters.Add("@fh", SqliteType.Integer) : c5.Parameters.Add("@fh", SqliteType.Integer)

                                       For Each s In stalePaths
                                           Dim h = StringToXxHash64(s) ' 取得孤兒的 Hash
                                           c1.Parameters("@fp").Value = s : dF += c1.ExecuteNonQuery()
                                           c2.Parameters("@fh").Value = h : dB += c2.ExecuteNonQuery()
                                           ' c3.Parameters("@fh").Value = h : dA += c3.ExecuteNonQuery()
                                           c4.Parameters("@fh").Value = h : dM += c4.ExecuteNonQuery()
                                           c5.Parameters("@fh").Value = h : dBasic += c5.ExecuteNonQuery()
                                       Next
                                   End Using
                                   txn.Commit()
                               End Using
                               For Each s In stalePaths : dA += SimDbDeleteAttachFilenamesByFolder(s) : Next

                               ' 2026/06/22 by Simon/Claude: attach_filenames 已由上行按 folder_hash 高效刪，故 includeAttachFilenames:=False，僅補 mail_simhash + 兩記憶體快取
                               dSh = SimDbDeleteMailRowsByEntryIds(orphanEntryIds, includeAttachFilenames:=False)

                               _dbg("結束", $"孤兒路徑:{stalePaths.Count} 個 / folder_stats:{dF} 行 / mailinfo_list:{dBasic} 行 / attach_maillist:{dB} 行 / attach_filenames:{dA} 行 / mail_simhash:{dSh} 行 / month_counts:{dM} 行")

                           Catch ex As System.Exception
                               _dbg("    ├ 錯誤", ex.Message) ' by Gemini, 2026/04/10
                           End Try
                       End Sub)

    End Function
#End Region

#Region "■ 批次寫入核心 (Batch Writer Core)"
    Private Function SaveFolderStatsInner(txn As SqliteTransaction) As Integer
        ' ---------------------------------------------------------------
        ' SaveFolderStatsInner — Transaction 內批次寫入 folder_stats
        ' 注意: 在 Task.Run 背景執行緒呼叫，不可碰 UI 控制項
        ' pr_count_snap = _cacheMailCount[path]，即 PR_CONTENT_COUNT 讀取結果
        ' ---------------------------------------------------------------
        _dbg("    ├ 開始") ' by Gemini, 2026/04/10: 調整縮排層級為 Level 2
        ' 2026/06/14 by Simon/Claude Opus 4.8: INSERT OR REPLACE → UPSERT。
        '   原 INSERT OR REPLACE 是「整列覆寫」: 當 path 在 allPaths 但 _cacheFolderIDs 查無時(line 941 Else 分支)，
        '   entry_id/store_id/is_mail/has_chinese 會被寫成 NULL，把 DB 既有的身分證洗掉。
        '   而樹載入 DbGetOrderedSubFolderIDs 帶 "entry_id IS NOT NULL" → 這些資料夾從樹消失
        '   (重現: 第 2 次 do-nothing 關閉存檔 → 第 3 次啟動 Gmail_2022 底下只剩收件匣)。
        '   且會自我延續: eid 被洗→下次 lazy-load 因 .eid 空只補 _cacheFolderCount(進 allPaths) 不補 _cacheFolderIDs→再被洗。
        '   修法: 改 ON CONFLICT DO UPDATE — 統計欄照常覆寫(與原行為一致)；身分欄以 COALESCE(新值, 舊值) 保留，
        '         新值為 NULL(快取沒身分證)時不動 DB 既有值 → entry_id 永不被洗掉，並打斷上述自我延續循環。
        Dim sql = "INSERT INTO folder_stats" &
                  " (folder_path,mail_count,mail_count_all,folder_count,folder_count_all,folder_size,folder_size_all,pr_count_snap,entry_id,store_id,is_mail,has_chinese,updated_at) " &
                  "VALUES (@fp,@mc,@mca,@fc,@fca,@fs,@fsa,@pr,@eid,@sid,@ism,@hasch,@ts) " &
                  "ON CONFLICT(folder_path) DO UPDATE SET " &
                  " mail_count=excluded.mail_count, mail_count_all=excluded.mail_count_all," &
                  " folder_count=excluded.folder_count, folder_count_all=excluded.folder_count_all," &
                  " folder_size=excluded.folder_size, folder_size_all=excluded.folder_size_all," &
                  " pr_count_snap=excluded.pr_count_snap," &
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

        ' 2026/06/12 by Simon/Claude Opus 4.8: folder_stats 作為主表，確保所有引用 folder_hash 的子表路徑都被收錄
        ' 重啟後 LoadFolderStatsInner 還原完整 _dictHashToPath，防止 LoadMailInfoInner 等批次載入 skip 資料
        For Each k In _cacheMailInfo.Keys : allPaths.Add(k) : Next
        For Each k In _cacheAttachMailList.Keys : allPaths.Add(k) : Next
        For Each k In _cacheYearCounts.Keys : allPaths.Add(k) : Next
        ' _cacheMonthCounts key 格式為 "FolderPath_year"，需解析後加入
        For Each k In _cacheMonthCounts.Keys
            Dim lastUs = k.LastIndexOf("_"c)
            If lastUs > 0 Then allPaths.Add(k.Substring(0, lastUs))
        Next

        Dim count As Integer = 0
        Dim ts = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
        Using cmd As New SqliteCommand(sql, _dbCache, txn)
            cmd.Parameters.Add("@fp", SqliteType.Text)
            cmd.Parameters.Add("@mc", SqliteType.Integer)
            cmd.Parameters.Add("@mca", SqliteType.Integer)
            cmd.Parameters.Add("@fc", SqliteType.Integer)
            cmd.Parameters.Add("@fca", SqliteType.Integer)
            cmd.Parameters.Add("@fs", SqliteType.Integer)
            cmd.Parameters.Add("@fsa", SqliteType.Integer)
            cmd.Parameters.Add("@pr", SqliteType.Integer)
            cmd.Parameters.Add("@eid", SqliteType.Blob)     ' 2026/06/10 by Gemini/Simon: 優化SQLite儲存空間把Entry_id改成BLOB二進位儲存
            cmd.Parameters.Add("@sid", SqliteType.Text)
            cmd.Parameters.Add("@ism", SqliteType.Integer)
            cmd.Parameters.Add("@hasch", SqliteType.Integer)
            cmd.Parameters.Add("@ts", SqliteType.Text)

            For Each path In allPaths
                ' 2026/04/07 修正 v2: 初始值設 -1 仍然不夠，因為 -1 是整數值會被寫入 DB，
                '   LoadFolderStatsInner 讀回 -1 後直接塞入記憶體快取，
                '   GetFolderCount 命中記憶體回傳 -1 → LoadSubFolderToTreeView 判斷 -1 > 0 為 False → 不顯示 "+"。
                '   正確做法：沒有測量過的欄位一律寫 DBNull.Value (SQL NULL)，這樣 LoadFolderStatsInner 的 IsDBNull 保護才能正確跳過，不污染記憶體快取。
                Dim mc, mca, fc, fca As Integer : Dim fs, fsa As Long
                Dim hasMc = _cacheMailCount.TryGetValue(path, mc)
                Dim hasMca = _cacheMailCountAll.TryGetValue(path, mca)
                Dim hasFc = _cacheFolderCount.TryGetValue(path, fc)
                Dim hasFca = _cacheFolderCountAll.TryGetValue(path, fca)
                Dim hasFs = _cacheFolderSize.TryGetValue(path, fs)
                Dim hasFsa = _cacheFolderSizeAll.TryGetValue(path, fsa)
                cmd.Parameters("@fp").Value = path
                cmd.Parameters("@mc").Value = If(hasMc, CObj(mc), DBNull.Value)
                cmd.Parameters("@mca").Value = If(hasMca, CObj(mca), DBNull.Value)
                cmd.Parameters("@fc").Value = If(hasFc, CObj(fc), DBNull.Value)
                cmd.Parameters("@fca").Value = If(hasFca, CObj(fca), DBNull.Value)
                cmd.Parameters("@fs").Value = If(hasFs, CObj(fs), DBNull.Value)
                cmd.Parameters("@fsa").Value = If(hasFsa, CObj(fsa), DBNull.Value)
                cmd.Parameters("@pr").Value = If(hasMc, CObj(mc), DBNull.Value)

                ' by Gemini, 2026/04/10: 寫入身分標識與標籤 (從 _cacheFolderIDs 提取)
                Dim idInfo As (eid As String, sid As String, isMail As Boolean, hasCh As Boolean) = Nothing
                If _cacheFolderIDs.TryGetValue(path, idInfo) Then
                    cmd.Parameters("@eid").Value = HexStringToByteArray(idInfo.eid) ' 2026/06/10 by Gemini/Simon: 優化SQLite儲存空間把Entry_id改成BLOB二進位儲存
                    cmd.Parameters("@sid").Value = idInfo.sid
                    cmd.Parameters("@ism").Value = If(idInfo.isMail, 1, 0)
                    cmd.Parameters("@hasch").Value = If(idInfo.hasCh, 1, 0)
                Else
                    cmd.Parameters("@eid").Value = DBNull.Value
                    cmd.Parameters("@sid").Value = DBNull.Value
                    cmd.Parameters("@ism").Value = DBNull.Value
                    cmd.Parameters("@hasch").Value = DBNull.Value
                End If
                cmd.Parameters("@ts").Value = ts
                cmd.ExecuteNonQuery() : count += 1
            Next
        End Using

        _dbg("    ├ 結束") ' by Gemini, 2026/04/11: 修正與開始對齊
        Return count

    End Function
    Private Function SaveYearCountsInner(txn As SqliteTransaction) As Integer
        ' ---------------------------------------------------------------
        ' SaveYearCountsInner — Transaction 內批次寫入 year_counts (Tab2 年份分布) 
        ' _cacheYearCounts: ConcurrentDictionary(Of String, ConcurrentDictionary(Of Integer, Integer))
        '   key = folder_path, value = { year → count }
        ' 每筆寫入 (folder_path, year, count)，PRIMARY KEY = (folder_path, year)
        ' ---------------------------------------------------------------
        _dbg("    ├ 開始") ' by Gemini, 2026/04/10: 調整縮排層級為 Level 2
        Dim sql = "INSERT OR REPLACE INTO year_counts (folder_hash,year,count) VALUES (@fh,@yr,@cnt)"
        Dim count As Integer = 0

        Using cmd As New SqliteCommand(sql, _dbCache, txn)
            cmd.Parameters.Add("@fh", SqliteType.Integer)
            cmd.Parameters.Add("@yr", SqliteType.Integer)
            cmd.Parameters.Add("@cnt", SqliteType.Integer)

            ' 2026/06/12 by Gemini/Simon: 把 folder_path TEXT 改成 folder_hash INTEGER 儲存，並新增一個記憶體字典 _dictHashToPath 來反查雜湊對應的路徑
            ' 2026/06/12 by Simon/Claude Opus 4.8: 移除 updated_at（只寫不讀，折省 ~1MB 儲存）
            For Each kvp In _cacheYearCounts
                Dim fp = kvp.Key
                For Each yr In kvp.Value
                    cmd.Parameters("@fh").Value = FolderPathToHash64(fp)
                    cmd.Parameters("@yr").Value = yr.Key
                    cmd.Parameters("@cnt").Value = yr.Value
                    cmd.ExecuteNonQuery() : count += 1
                Next
            Next
        End Using

        Return count
        _dbg(" ├ 結束")

    End Function
    Private Function SaveMonthCountsInner(txn As SqliteTransaction) As Integer
        ' ---------------------------------------------------------------
        ' SaveMonthCountsInner — Transaction 內批次寫入 month_counts (Tab2 月份分布) 
        ' _cacheMonthCounts key = FolderPath_year，value = { month → count }
        ' 欄位設計: (folder_path, year, month) 三欄 PK，語意清晰且符合 ForFolder() 查詢
        '   ├ month_counts 新增函數群 (2026/04/09 by Claude)
        '   ├ 2026/04/09 修正：改用三欄 PK，移除 cache_key 欄位
        ' ---------------------------------------------------------------
        _dbg(" ├ 開始")
        Dim sql = "INSERT OR REPLACE INTO month_counts (folder_hash,year,month,count) VALUES (@fh,@yr,@mo,@cnt)"
        Dim count As Integer = 0

        Using cmd As New SqliteCommand(sql, _dbCache, txn)
            cmd.Parameters.Add("@fh", SqliteType.Integer)
            cmd.Parameters.Add("@yr", SqliteType.Integer)
            cmd.Parameters.Add("@mo", SqliteType.Integer)
            cmd.Parameters.Add("@cnt", SqliteType.Integer)

            For Each kvp In _cacheMonthCounts
                ' cache_key 格式: "FolderPath_year"，最後一個 "_" 分隔出 year
                Dim cacheKey = kvp.Key
                Dim lastUnderscore = cacheKey.LastIndexOf("_"c)
                If lastUnderscore < 0 Then Continue For
                Dim fPath = cacheKey.Substring(0, lastUnderscore)
                Dim yearVal As Integer
                If Not Integer.TryParse(cacheKey.Substring(lastUnderscore + 1), yearVal) Then Continue For

                ' 2026/06/12 by Gemini/Simon: 把 folder_path TEXT 改成 folder_hash INTEGER 儲存，並新增一個記憶體字典 _dictHashToPath 來反查雜湊對應的路徑
                ' 2026/06/12 by Simon/Claude Opus 4.8: 移除 updated_at（只寫不讀）
                For Each mo In kvp.Value
                    cmd.Parameters("@fh").Value = FolderPathToHash64(fPath)
                    cmd.Parameters("@yr").Value = yearVal
                    cmd.Parameters("@mo").Value = mo.Key
                    cmd.Parameters("@cnt").Value = mo.Value
                    cmd.ExecuteNonQuery() : count += 1
                Next
            Next
        End Using

        _dbg(" ├ 結束", $"{count} 筆")
        Return count

    End Function
    Private Function SaveAttachMailListInner(txn As SqliteTransaction) As Integer
        ' ---------------------------------------------------------------
        ' SaveAttachMailListInner — Transaction 內批次寫入 attach_maillist (Tab3 Phase1) 
        ' 2026/06/12 by Simon/Claude Opus 4.8: received_time TEXT→INTEGER (Unix秒)；移除 updated_at
        ' ---------------------------------------------------------------
        _dbg("開始")
        Dim sql = "INSERT OR REPLACE INTO attach_maillist" &
                  " (entry_id,folder_hash,subject,msg_size,received_time,sender_name,attach_count,pr_count_snap)" &
                  " VALUES (@eid,@fh,@subj,@sz,@rt,@sn,@ac,@pr)"

        Dim count As Integer = 0
        Using cmd As New SqliteCommand(sql, _dbCache, txn)
            cmd.Parameters.Add("@eid", SqliteType.Blob)     ' 2026/06/10 by Gemini/Simon: 優化SQLite儲存空間把Entry_id改成BLOB二進位儲存
            cmd.Parameters.Add("@fh", SqliteType.Integer)
            cmd.Parameters.Add("@subj", SqliteType.Text)
            cmd.Parameters.Add("@sz", SqliteType.Integer)
            cmd.Parameters.Add("@rt", SqliteType.Integer)   ' 2026/06/12 by Simon/Claude Opus 4.8: INTEGER Unix秒
            cmd.Parameters.Add("@sn", SqliteType.Text)
            cmd.Parameters.Add("@ac", SqliteType.Integer)
            cmd.Parameters.Add("@pr", SqliteType.Integer)

            ' _cacheAttachMailList: Dictionary(Of String, FolderCacheTab3)
            ' key = folder_path, value.AttachMailList = List(Of MailItemInfo)
            For Each kvp In _cacheAttachMailList
                Dim fp = kvp.Key : Dim snap = kvp.Value.ItemCountSnap
                Dim mails = kvp.Value.AttachMailList

                ' by Gemini 3 Flash, 2026/05/06: 實作「空結果持久化」，記住此資料夾已掃描且為 0 筆
                If mails.Count = 0 Then
                    ' 2026/06/10 by Gemini/Simon: 優化SQLite儲存空間把Entry_id改成BLOB二進位儲存
                    ' 2026/06/12 by Gemini/Simon: 把 folder_path TEXT 改成 folder_hash INTEGER 儲存
                    cmd.Parameters("@eid").Value = HexStringToByteArray("EMPTY_ATTACH_" & fp)
                    cmd.Parameters("@fh").Value = FolderPathToHash64(fp)
                    cmd.Parameters("@subj").Value = ""
                    cmd.Parameters("@sz").Value = 0
                    cmd.Parameters("@rt").Value = 0L   ' 2026/06/12 by Simon/Claude Opus 4.8: sentinel 用 epoch 0
                    cmd.Parameters("@sn").Value = ""
                    cmd.Parameters("@ac").Value = 0
                    cmd.Parameters("@pr").Value = snap
                    cmd.ExecuteNonQuery() : count += 1
                Else
                    ' 2026/06/12 by Simon/Claude Opus 4.8: 本機時間轉 Unix 秒儲存，讀回時 FromUnixTimeSeconds().LocalDateTime 還原
                    For Each mail In mails
                        cmd.Parameters("@eid").Value = HexStringToByteArray(mail.EntryID)
                        cmd.Parameters("@fh").Value = FolderPathToHash64(fp)
                        cmd.Parameters("@subj").Value = If(mail.Subject, "")
                        cmd.Parameters("@sz").Value = mail.Size
                        cmd.Parameters("@rt").Value = LocalTimeToUnixSeconds(mail.RcvTime)
                        cmd.Parameters("@sn").Value = If(mail.SenderName, "")
                        cmd.Parameters("@ac").Value = mail.AttachCount
                        cmd.Parameters("@pr").Value = snap
                        cmd.ExecuteNonQuery() : count += 1
                    Next
                End If
            Next
        End Using
        Return count
        _dbg("結束")

    End Function
    Private Function SaveAttachFilenamesInner() As Integer
        ' ---------------------------------------------------------------
        ' SaveAttachFilenamesInner — Transaction 內批次寫入 attach_filenames (Tab3 Phase2) 
        ' folder_path 透過反查 _cacheAttachMailList 取得 (_cacheAttachFilename key 是 EntryID) 
        ' 2026/04/09 修正: 移除 msg_size 欄位 (Phase2 永遠是 NULL，保留在 INSERT 造成
        '   SqliteType.Integer + DBNull.Value 不相容，丟 "Value must be set" InvalidOperationException)
        '   SQLite 未列出的欄位自動填 NULL，不需要明確傳入。
        ' ---------------------------------------------------------------
        _dbg("開始")
        If _dbMail Is Nothing Then Return 0

        Dim sql = "INSERT OR REPLACE INTO attach_filenames" & " (entry_id,folder_hash,filenames)" & " VALUES (@eid,@fh,@fn)"
        Dim count As Integer = 0

        ' 反查 EntryID → folder_path (從 Phase1 快取中建立對應表) 
        Dim entryToFolder As New Dictionary(Of String, String)()
        For Each kvp In _cacheAttachMailList
            For Each mail In kvp.Value.AttachMailList
                If Not entryToFolder.ContainsKey(mail.EntryID) Then entryToFolder(mail.EntryID) = kvp.Key
            Next
        Next

        ' 2026/06/10 by Gemini/Simon: 優化SQLite儲存空間把Entry_id改成BLOB二進位儲存
        ' 2026/06/12 by Gemini/Simon: 把 folder_path TEXT 改成 folder_hash INTEGER 儲存，並新增一個記憶體字典 _dictHashToPath 來反查雜湊對應的路徑
        ' 2026/06/12 by Simon/Claude Opus 4.8: 移除 updated_at（只寫不讀）
        ' 2026/06/21 by Simon/Claude Opus 4.8: attach_filenames 搬至 OLAcacheMail.db(_dbMail)，改自管獨立 _dbMail 交易(不再吃外部 _dbCache txn)
        Try
            Using txnSim = _dbMail.BeginTransaction()
                Using cmd As New SqliteCommand(sql, _dbMail, txnSim)
                    cmd.Parameters.Add("@eid", SqliteType.Blob)
                    cmd.Parameters.Add("@fh", SqliteType.Integer)
                    cmd.Parameters.Add("@fn", SqliteType.Text)

                    For Each kvp In _cacheAttachFilename
                        Dim fp = "" : entryToFolder.TryGetValue(kvp.Key, fp)
                        cmd.Parameters("@eid").Value = HexStringToByteArray(kvp.Key)
                        cmd.Parameters("@fh").Value = FolderPathToHash64(fp)
                        cmd.Parameters("@fn").Value = JsonSerializer.Serialize(kvp.Value)
                        cmd.ExecuteNonQuery() : count += 1
                    Next
                End Using
                txnSim.Commit()
            End Using
        Catch ex As System.Exception
            _dbg("SaveAttachFilenamesInner 錯誤", ex.Message)   ' _dbMail 寫入失敗不連累主 db；下次掃描自動重建
        End Try
        Return count
        _dbg("結束")

    End Function
    Private Function SaveSendersInner(txn As SqliteTransaction) As Integer
        ' ---------------------------------------------------------------
        ' SaveSendersInner — 收集 _cacheMailInfo 中的唯一 email，
        '   批次 INSERT OR IGNORE 進 senders 表，再重建兩個記憶體字典
        ' 2026/06/12 by Simon/Claude Opus 4.8: 配合 sender_email 正規化架構新增
        '   呼叫時機：SaveMailInfoInner 之前（同一 Transaction）
        '   完成後 _dictEmailToSenderId 可供 SaveMailInfoInner 直接查 sender_id
        ' ---------------------------------------------------------------
        Dim allEmails As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        For Each kvp In _cacheMailInfo
            For Each item In kvp.Value.Mails
                Dim email = item.Mail.SenderEmail?.Trim()
                If Not String.IsNullOrEmpty(email) Then allEmails.Add(email.ToLower())
            Next
        Next

        If allEmails.Count = 0 Then Return 0

        Using cmd As New SqliteCommand("INSERT OR IGNORE INTO senders (sender_email) VALUES (@se)", _dbCache, txn)
            cmd.Parameters.Add("@se", SqliteType.Text)
            For Each email In allEmails
                cmd.Parameters("@se").Value = email
                cmd.ExecuteNonQuery()
            Next
        End Using

        ' 重建記憶體字典（在同一 Transaction 內可讀到剛寫入的新 rows）
        LoadSendersInner()
        Return allEmails.Count   ' 回傳新增 sender 數

    End Function
    Private Function SaveMailInfoInner(txn As SqliteTransaction) As Integer
        ' ---------------------------------------------------------------
        ' SaveMailInfoInner — Transaction 內批次寫入 mailinfo_list (Tab4/Tab5)
        ' 2026/06/10 by Gemini/Simon: 優化SQLite儲存空間把Entry_id改成BLOB二進位儲存
        ' 2026/06/11 by Gemini/Simon: 把 message_id 轉成 xxHash64，並同時改成 BLOB 儲存節省空間
        ' 2026/06/12 by Gemini/Simon: 把 folder_path TEXT 改成 folder_hash INTEGER 儲存
        ' 2026/06/12 by Simon/Claude Opus 4.8: 移除 topic (動態計算)、sender_email (改 sender_id)、updated_at
        '   received_time TEXT→INTEGER (Unix秒)；SaveSendersInner() 必須在此函式前呼叫
        ' ---------------------------------------------------------------
        _dbg("開始")
        Dim sql = "INSERT OR REPLACE INTO mailinfo_list" &
                  " (entry_id,folder_hash,subject,msg_size,received_time,sender_name,sender_id,msgid_hash,pr_count_snap)" &
                  " VALUES (@eid,@fh,@subj,@sz,@rt,@sn,@sid,@mid,@pr)"

        Dim count As Integer = 0
        Using cmd As New SqliteCommand(sql, _dbCache, txn)
            cmd.Parameters.Add("@eid", SqliteType.Blob)
            cmd.Parameters.Add("@fh", SqliteType.Integer)
            cmd.Parameters.Add("@subj", SqliteType.Text)
            cmd.Parameters.Add("@sz", SqliteType.Integer)
            cmd.Parameters.Add("@rt", SqliteType.Integer)   ' 2026/06/12 by Simon/Claude Opus 4.8: INTEGER Unix秒
            cmd.Parameters.Add("@sn", SqliteType.Text)
            cmd.Parameters.Add("@sid", SqliteType.Integer)  ' 2026/06/12 by Simon/Claude Opus 4.8: sender_id (NULL 若無 email)
            cmd.Parameters.Add("@mid", SqliteType.Blob)
            cmd.Parameters.Add("@pr", SqliteType.Integer)

            For Each kvp In _cacheMailInfo
                ' 2026/05/06 by Claude: key 已是純路徑，不再需 .Split
                Dim fp As String = kvp.Key
                Dim snap = kvp.Value.Snap
                Dim mails = kvp.Value.Mails

                If mails.Count = 0 Then
                    ' by Gemini 3 Flash, 2026/05/06: 實作「空結果持久化」，記住此資料夾已掃描且為 0 筆
                    cmd.Parameters("@eid").Value = HexStringToByteArray("EMPTY_BASIC_" & fp)
                    cmd.Parameters("@fh").Value = FolderPathToHash64(fp)
                    cmd.Parameters("@subj").Value = ""
                    cmd.Parameters("@sz").Value = 0
                    cmd.Parameters("@rt").Value = 0L
                    cmd.Parameters("@sn").Value = ""
                    cmd.Parameters("@sid").Value = DBNull.Value
                    cmd.Parameters("@mid").Value = DBNull.Value
                    cmd.Parameters("@pr").Value = snap
                    cmd.ExecuteNonQuery() : count += 1
                Else
                    For Each item In mails
                        ' 2026/06/12 by Simon/Claude Opus 4.8: 查 _dictEmailToSenderId，無 email 時存 NULL
                        Dim emailKey = item.Mail.SenderEmail?.Trim()?.ToLower()
                        Dim sid As Object
                        Dim foundId As Integer
                        If Not String.IsNullOrEmpty(emailKey) AndAlso _dictEmailToSenderId.TryGetValue(emailKey, foundId) Then
                            sid = foundId
                        Else
                            sid = DBNull.Value
                        End If

                        cmd.Parameters("@eid").Value = HexStringToByteArray(item.Mail.EntryID)
                        cmd.Parameters("@fh").Value = FolderPathToHash64(fp)
                        cmd.Parameters("@subj").Value = If(item.Mail.Subject, "")
                        cmd.Parameters("@sz").Value = item.Mail.Size
                        cmd.Parameters("@rt").Value = LocalTimeToUnixSeconds(item.Mail.RcvTime) ' 2026/06/12 by Simon/Claude Opus 4.8: 本機時間轉 Unix 秒
                        cmd.Parameters("@sn").Value = If(item.Mail.SenderName, "")
                        cmd.Parameters("@sid").Value = sid
                        cmd.Parameters("@mid").Value = HexStringToByteArray(If(item.Mail.MsgIDhash, ""))
                        cmd.Parameters("@pr").Value = snap
                        cmd.ExecuteNonQuery() : count += 1
                    Next
                End If
            Next
        End Using
        Return count
        _dbg("結束")
    End Function
#End Region

#Region "■ 批次載入核心 (Batch Reader Core)"
    Private Function LoadFolderStatsInner() As Integer
        ' ---------------------------------------------------------------
        ' LoadFolderStatsInner — 讀回六個數字快取
        ' 使用 TryAdd：記憶體已有值時保留記憶體版本 (不覆蓋 Layer2.5 剛讀進來的較新值) 
        ' 2026/04/07 修正: 每個欄位加 IsDBNull 保護，NULL 代表「從未測量過」，
        '   不可塞入記憶體快取，否則 GetFolderCount 命中 -1 → LoadSubFolderToTreeView
        '   判斷 -1 > 0 為 False → 不顯示 TreeView "+" 加號 (bug) 。
        ' 2026/06/12 by Gemini/Simon: 把 folder_path TEXT 改成 folder_hash INTEGER 儲存，並新增一個記憶體字典 _dictHashToPath 來反查雜湊對應的路徑
        ' ---------------------------------------------------------------
        _dbg("開始")
        Dim count As Integer = 0
        Using cmd As New SqliteCommand(
            "SELECT folder_path,mail_count,mail_count_all,folder_count,folder_count_all,folder_size,folder_size_all,entry_id,store_id,is_mail,has_chinese FROM folder_stats", _dbCache)
            Using reader = cmd.ExecuteReader()
                While reader.Read()
                    Dim path = reader.GetString(0)
                    FolderPathToHash64(path)  ' 👈 註冊到字典，讓後續的明細表可以直接反查

                    ' 只有 NOT NULL 的欄位才塞入記憶體快取；NULL 代表「從未測量過」，跳過
                    If Not reader.IsDBNull(1) Then _cacheMailCount.TryAdd(path, reader.GetInt64(1))
                    If Not reader.IsDBNull(2) Then _cacheMailCountAll.TryAdd(path, reader.GetInt64(2))
                    If Not reader.IsDBNull(3) Then _cacheFolderCount.TryAdd(path, reader.GetInt64(3))
                    If Not reader.IsDBNull(4) Then _cacheFolderCountAll.TryAdd(path, reader.GetInt64(4))
                    If Not reader.IsDBNull(5) Then _cacheFolderSize.TryAdd(path, reader.GetInt64(5))
                    If Not reader.IsDBNull(6) Then _cacheFolderSizeAll.TryAdd(path, reader.GetInt64(6))

                    ' by Gemini 3.0 flash, 2026/04/18: 批量讀取時回填身分標識與標籤字典，確保 LoadCache 後狀態完整
                    ' 2026/06/10 by Gemini/Simon: 優化SQLite儲存空間把Entry_id改成BLOB二進位儲存
                    Dim eid As String = If(Not reader.IsDBNull(7), ByteArrayToHexString(reader.GetFieldValue(Of Byte())(7)), "")
                    Dim sid As String = If(Not reader.IsDBNull(8), reader.GetString(8), "")
                    Dim isMail As Integer = If(Not reader.IsDBNull(9), reader.GetInt32(9), -1)
                    Dim hasCh As Integer = If(Not reader.IsDBNull(10), reader.GetInt32(10), -1)

                    If eid <> "" Then _cacheFolderIDs.TryAdd(path, (eid, sid, isMail = 1, hasCh = 1))
                    If isMail >= 0 Then _cacheIsMailFolder.TryAdd(path, isMail = 1)

                    count += 1
                End While
            End Using
        End Using
        Return count
        _dbg("結束")

    End Function
    Private Function LoadYearCountsInner() As Integer
        ' ---------------------------------------------------------------
        ' LoadYearCountsInner — 從 year_counts 重建 _cacheYearCounts
        ' 先按 folder_path 分組收集，最後一次性寫入 (TryAdd 保留記憶體已有版本)
        ' ---------------------------------------------------------------
        _dbg(" ├ 開始")
        Dim count As Integer = 0
        Dim tempDict As New Dictionary(Of String, ConcurrentDictionary(Of Integer, Integer))()

        Using cmd As New SqliteCommand("SELECT folder_hash,year,count FROM year_counts", _dbCache)
            Using reader = cmd.ExecuteReader()
                While reader.Read()
                    Dim fp As String = "" : If Not _dictHashToPath.TryGetValue(reader.GetInt64(0), fp) Then Continue While
                    Dim yr = reader.GetInt32(1)
                    Dim cnt = reader.GetInt32(2)
                    If Not tempDict.ContainsKey(fp) Then tempDict(fp) = New ConcurrentDictionary(Of Integer, Integer)()
                    tempDict(fp)(yr) = cnt
                    count += 1
                End While
            End Using
        End Using

        For Each kvp In tempDict
            _cacheYearCounts.TryAdd(kvp.Key, kvp.Value)
        Next
        Return count
        _dbg(" ├ 結束")

    End Function
    Private Function LoadMonthCountsInner() As Integer
        ' ---------------------------------------------------------------
        ' LoadMonthCountsInner — 從 month_counts 重建 _cacheMonthCounts
        ' 按 (folder_path, year) 分組，TryAdd 保留記憶體已有版本
        '   ├ month_counts 新增函數群 (2026/04/09 by Claude)
        '   ├ 2026/04/09 修正：改用三欄 PK，從 folder_path + year 重組 cacheKey
        ' ---------------------------------------------------------------
        _dbg(" ├ 開始")
        Dim count As Integer = 0
        Dim tempDict As New Dictionary(Of String, ConcurrentDictionary(Of Integer, Integer))()

        Using cmd As New SqliteCommand("SELECT folder_hash,year,month,count FROM month_counts", _dbCache)
            Using reader = cmd.ExecuteReader()
                While reader.Read()
                    Dim fp As String = "" : If Not _dictHashToPath.TryGetValue(reader.GetInt64(0), fp) Then Continue While
                    Dim yr = reader.GetInt32(1)
                    Dim mo = reader.GetInt32(2)
                    Dim cnt = reader.GetInt32(3)
                    Dim ck = fp & "_" & yr.ToString()   ' 重組 cacheKey，與 _cacheMonthCounts key 格式一致

                    Dim value As ConcurrentDictionary(Of Integer, Integer) = Nothing
                    If Not tempDict.TryGetValue(ck, value) Then
                        value = New ConcurrentDictionary(Of Integer, Integer)()
                        tempDict(ck) = value
                    End If
                    value(mo) = cnt : count += 1
                End While
            End Using
        End Using

        For Each kvp In tempDict : _cacheMonthCounts.TryAdd(kvp.Key, kvp.Value) : Next
        _dbg(" ├ 結束", $"{count} 筆 → {tempDict.Count} 個 cache_key")
        Return count

    End Function
    Private Function LoadAttachMailListInner() As Integer
        ' ---------------------------------------------------------------
        ' LoadAttachMailListInner — 重建 _cacheAttachMailList (按 folder_path 分組)
        ' 2026/06/12 by Simon/Claude Opus 4.8: received_time 改從 INTEGER 讀取 (Unix秒→LocalDateTime)
        '   順手修正舊版 bug: mail.Size 誤讀 index 2 (subject)，應讀 index 3 (msg_size)
        ' ---------------------------------------------------------------
        _dbg("開始")
        Dim count As Integer = 0
        ' 暫存用：先按 folder_path 分組收集，最後一次性寫入 _cacheAttachMailList
        Dim tempDict As New Dictionary(Of String, (snap As Integer, mails As List(Of MailItemInfo)))()
        Using cmd As New SqliteCommand(
            "SELECT entry_id,folder_hash,subject,msg_size,received_time,sender_name,attach_count,pr_count_snap FROM attach_maillist", _dbCache)
            Using reader = cmd.ExecuteReader()
                While reader.Read()
                    Dim fp As String = ""
                    If Not _dictHashToPath.TryGetValue(reader.GetInt64(1), fp) Then Continue While
                    If Not tempDict.ContainsKey(fp) Then tempDict(fp) = (reader.GetInt64(7), New List(Of MailItemInfo)())

                    ' 2026/06/10 by Gemini/Simon: 優化SQLite儲存空間把Entry_id改成BLOB二進位儲存
                    ' 2026/06/12 by Simon/Claude Opus 4.8: received_time 改為 INTEGER (Unix秒)
                    Dim mail As New MailItemInfo With {.EntryID = ByteArrayToHexString(reader.GetFieldValue(Of Byte())(0)),
                                                       .Subject = If(reader.IsDBNull(2), "", reader.GetString(2)),
                                                       .Size = If(reader.IsDBNull(3), 0L, reader.GetInt64(3)),  ' 2026/06/12 修正 bug: 原為 index 2
                                                       .RcvTime = UnixSecondsToLocalTime(If(reader.IsDBNull(4), 0L, reader.GetInt64(4))),
                                                       .SenderName = If(reader.IsDBNull(5), "", reader.GetString(5)),
                                                       .AttachCount = reader.GetInt32(6),
                                                       .FolderPath = fp} ' 確保讀取快取時填入路徑
                    tempDict(fp).mails.Add(mail) : count += 1
                End While
            End Using
        End Using

        For Each kvp In tempDict
            _cacheAttachMailList.TryAdd(kvp.Key, New FolderCacheTab3 With {.AttachMailList = kvp.Value.mails, .ItemCountSnap = kvp.Value.snap})
        Next
        Return count
        _dbg("結束")

    End Function
    Private Function LoadAttachFilenamesInner() As Integer
        ' ---------------------------------------------------------------
        ' LoadAttachFilenamesInner — 重建 _cacheAttachFilename (JSON 反序列化)
        ' ---------------------------------------------------------------
        _dbg("開始")
        Dim count As Integer = 0
        If _dbMail Is Nothing Then Return 0   ' 2026/06/21 by Simon/Claude Opus 4.8: attach_filenames 來源改 _dbMail(OLAcacheMail.db)
        Using cmd As New SqliteCommand("SELECT entry_id,filenames FROM attach_filenames", _dbMail)
            Using reader = cmd.ExecuteReader()
                ' 2026/06/10 by Gemini/Simon: 優化SQLite儲存空間把Entry_id改成BLOB二進位儲存
                While reader.Read()
                    Dim eidStr = ByteArrayToHexString(reader.GetFieldValue(Of Byte())(0))
                    Dim fnJson = If(reader.IsDBNull(1), "[]", reader.GetString(1))
                    Try
                        Dim list = JsonSerializer.Deserialize(Of List(Of String))(fnJson)
                        _cacheAttachFilename.TryAdd(eidStr, list)
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
    Private Function LoadSendersInner() As Integer
        ' ---------------------------------------------------------------
        ' LoadSendersInner — 載入 senders 表，重建 _dictEmailToSenderId / _dictSenderIdToEmail
        ' 2026/06/12 by Simon/Claude Opus 4.8: 配合 sender_email 正規化架構新增
        '   - 寫入側 (_dictEmailToSenderId): SaveMailInfoInner 查詢 sender_id
        '   - 讀取側 (_dictSenderIdToEmail): Load/DbGet 函式還原 SenderEmail
        ' ---------------------------------------------------------------
        _dictEmailToSenderId.Clear()
        _dictSenderIdToEmail.Clear()
        If _dbCache Is Nothing Then Return 0

        Try
            Using cmd As New SqliteCommand("SELECT sender_id, sender_email FROM senders", _dbCache)
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
            _dbg(" ├ LoadSendersInner 錯誤", ex.Message)
        End Try

        ' count 是已在 While 內計數的整數，最後加：
        Return _dictSenderIdToEmail.Count

    End Function
    Private Function LoadMailInfoInner() As Integer
        ' ---------------------------------------------------------------
        ' LoadMailInfoInner — 重建 _cacheMailInfo (Tab4/5 專用)
        ' 2026/04/22 by Gemini 3.1 Pro: 補齊載入邏輯，解決重啟後重複掃描問題
        ' 2026/06/10 by Gemini/Simon: 優化SQLite儲存空間把Entry_id改成BLOB二進位儲存
        ' 2026/06/11 by Gemini/Simon: 把 message_id 轉成 xxHash64，並同時改成 BLOB 儲存節省空間
        ' 2026/06/12 by Simon/Claude Opus 4.8: topic 改由 GetCleanSubject(subject) 動態計算，不再從 DB 讀欄位
        '   sender_email 改由 _dictSenderIdToEmail 從 sender_id 還原；received_time 改 INTEGER
        ' ---------------------------------------------------------------
        _dbg("開始")
        ' 由於此表資料量可能較大，我們按 folder_path 分組收集到記憶體中
        Dim count As Integer = 0
        Dim tempDict As New Dictionary(Of String, (Mails As List(Of (Mail As MailItemInfo, Topic As String)), Snap As Integer))()

        ' 2026/06/12: 新欄位順序 — entry_id(0),folder_hash(1),subject(2),msg_size(3),
        '             received_time INTEGER(4),sender_name(5),sender_id(6),msgid_hash(7),pr_count_snap(8)
        Using cmd As New SqliteCommand(
            "SELECT entry_id,folder_hash,subject,msg_size,received_time,sender_name,sender_id,msgid_hash,pr_count_snap FROM mailinfo_list", _dbCache)
            Using reader = cmd.ExecuteReader()
                While reader.Read()
                    Dim eidStr = ByteArrayToHexString(reader.GetFieldValue(Of Byte())(0))
                    Dim fp As String = "" : If Not _dictHashToPath.TryGetValue(reader.GetInt64(1), fp) Then Continue While
                    Dim subj = If(reader.IsDBNull(2), "", reader.GetString(2))
                    Dim sz = reader.GetInt64(3)
                    Dim rt As Long = If(reader.IsDBNull(4), 0L, reader.GetInt64(4))
                    Dim sn = If(reader.IsDBNull(5), "", reader.GetString(5))
                    Dim semail = _dictSenderIdToEmail.GetValueOrDefault(If(reader.IsDBNull(6), 0, reader.GetInt32(6)), "") ' 2026/06/12 by Simon/Claude Opus 4.8: sender_id → email 反查
                    Dim mid = If(reader.IsDBNull(7), "", ByteArrayToHexString(reader.GetFieldValue(Of Byte())(7)))
                    Dim snap = If(reader.IsDBNull(8), -1L, reader.GetInt64(8))

                    ' 2026/05/06 by Claude: cache key 改為純路徑
                    Dim cacheKey = fp
                    If Not tempDict.ContainsKey(cacheKey) Then tempDict(cacheKey) = (New List(Of (Mail As MailItemInfo, Topic As String))(), snap)

                    ' by Claude Sonnet 4.6, 2026/05/06: 跳過 sentinel row，但保留 snap
                    If eidStr.StartsWith("EMPTY_BASIC_") Then Continue While

                    Dim mail As New MailItemInfo With {.EntryID = eidStr, .Subject = subj, .Size = sz, .SenderName = sn,
                                                       .RcvTime = UnixSecondsToLocalTime(rt), .FolderPath = fp, .MsgIDhash = mid, .SenderEmail = semail}
                    tempDict(cacheKey).Mails.Add((mail, GetCleanSubject(subj)))
                    count += 1
                End While
            End Using
        End Using

        For Each kvp In tempDict
            _cacheMailInfo.TryAdd(kvp.Key, kvp.Value)
        Next
        Return count
        _dbg("結束")
    End Function
#End Region

#Region "■ DbGetXXX 即時查詢 (Lazy SELECT Helpers)"
    ' Phase 2 — Layer2.5 lazy SELECT 用的 DB read helper 群
    ' ==============================================================
    ' 設計原則 (2026-04-07):
    '   1. 只做「讀」，不做「寫」。寫入仍由 SaveCachesToDB (SaveCache 按鈕) 批次處理。
    '   2. 回傳 Nothing 表示 DB 中無此筆資料，呼叫端應繼續往 Layer3 走。
    '   3. 這些函數從 UI 執行緒呼叫，SQLite keyed lookup < 1ms，不需要 Async。
    '   4. FolderStatsDbRow / AttachMailListDbResult 定義在本檔，Partial Class 跨檔可見。
    ' ==============================================================
    Friend Function DbGetFolderStats(fPath As String) As FolderStatsDbRow
        ' ---------------------------------------------------------------
        ' DbGetFolderStats — 讀取 folder_stats 單一 folder_path 的一行
        ' 回傳 Nothing 表示 DB 中無此路徑
        ' ---------------------------------------------------------------
        If _iLikeNoisy Then _dbg(" ├ 開始")
        If _dbCache Is Nothing Then Return Nothing

        Try
            Using cmd As New SqliteCommand(
                "SELECT mail_count,mail_count_all,folder_count,folder_count_all," &
                "       folder_size,folder_size_all,pr_count_snap,entry_id,store_id,is_mail,has_chinese" &
                "  FROM folder_stats WHERE folder_path=@fp", _dbCache)
                cmd.Parameters.AddWithValue("@fp", fPath)
                Using reader = cmd.ExecuteReader()
                    If Not reader.Read() Then Return Nothing

                    ' 2026/06/10 by Gemini/Simon: 優化SQLite儲存空間把Entry_id改成BLOB二進位儲存
                    Return New FolderStatsDbRow With {.mc = If(reader.IsDBNull(0), -1L, reader.GetInt64(0)),
                                                      .mca = If(reader.IsDBNull(1), -1L, reader.GetInt64(1)),
                                                      .fc = If(reader.IsDBNull(2), -1L, reader.GetInt64(2)),
                                                      .fca = If(reader.IsDBNull(3), -1L, reader.GetInt64(3)),
                                                      .fs = If(reader.IsDBNull(4), -1L, reader.GetInt64(4)),
                                                      .fsa = If(reader.IsDBNull(5), -1L, reader.GetInt64(5)),
                                                      .snap = If(reader.IsDBNull(6), -1L, reader.GetInt64(6)),
                                                      .eid = If(reader.IsDBNull(7), "", ByteArrayToHexString(reader.GetFieldValue(Of Byte())(7))),
                                                      .sid = If(reader.IsDBNull(8), "", reader.GetString(8)),
                                                      .isMail = If(reader.IsDBNull(9), -1, reader.GetInt32(9)),
                                                      .hasCh = If(reader.IsDBNull(10), -1, reader.GetInt32(10))}
                End Using
            End Using
        Catch ex As System.Exception
            _dbg(" ├ 錯誤", $"{fPath}: {ex.Message}")
        Finally : If _iLikeNoisy Then _dbg(" ├ 結束")
        End Try
        Return Nothing

    End Function
    Friend Function DbGetYearCountsForFolder(fPath As String) As ConcurrentDictionary(Of Integer, Integer)
        ' ---------------------------------------------------------------
        ' DbGetYearCountsForFolder — 讀取 year_counts WHERE folder_path=? 的所有行
        ' 供 ComputeYearCountsAsync 在記憶體 miss 時先查 DB，避免 COM 呼叫
        ' 回傳 Nothing 表示 DB 中無此資料夾的年份記錄
        ' ---------------------------------------------------------------
        If _iLikeNoisy Then _dbg(" ├ 開始")
        If _dbCache Is Nothing Then Return Nothing

        Try
            Dim result As New ConcurrentDictionary(Of Integer, Integer)()
            Using cmd As New SqliteCommand("SELECT year,count FROM year_counts WHERE folder_hash=@fh", _dbCache)
                ' 2026/06/12 by Gemini/Simon: 把 folder_path TEXT 改成 folder_hash INTEGER 儲存，並新增一個記憶體字典 _dictHashToPath 來反查雜湊對應的路徑
                cmd.Parameters.AddWithValue("@fh", FolderPathToHash64(fPath))
                Using reader = cmd.ExecuteReader()
                    While reader.Read() : result(reader.GetInt32(0)) = reader.GetInt32(1) : End While
                End Using
            End Using
            Return If(result.Count > 0, result, Nothing)
        Catch ex As System.Exception
            _dbg(" ├ 錯誤", $"{fPath}: {ex.Message}")
        Finally : If _iLikeNoisy Then _dbg(" ├ 結束")
        End Try
        Return Nothing

    End Function
    Friend Function DbGetMonthCountsForFolder(fPath As String, year As Integer) As ConcurrentDictionary(Of Integer, Integer)
        ' ---------------------------------------------------------------
        ' DbGetMonthCountsForFolder — 讀取 month_counts WHERE folder_path=? AND year=?
        ' 供 GetMonthCountsForYearL3 在記憶體 miss 時先查 DB，避免 COM 呼叫
        ' 回傳 Nothing 表示 DB 中無此 (folder_path, year) 組合
        '   ├ month_counts 新增函數群 (2026/04/09 by Claude)
        '   ├ 2026/04/09 修正：改用三欄 PK，接收 (fPath, year) 兩個參數
        ' ---------------------------------------------------------------
        If _iLikeNoisy Then _dbg(" ├ 開始")
        If _dbCache Is Nothing Then Return Nothing
        Try
            Dim result As New ConcurrentDictionary(Of Integer, Integer)()
            Using cmd As New SqliteCommand("SELECT month,count FROM month_counts WHERE folder_hash=@fh AND year=@yr", _dbCache)
                ' 2026/06/12 by Gemini/Simon: 把 folder_path TEXT 改成 folder_hash INTEGER 儲存，並新增一個記憶體字典 _dictHashToPath 來反查雜湊對應的路徑
                cmd.Parameters.AddWithValue("@fh", FolderPathToHash64(fPath))
                cmd.Parameters.AddWithValue("@yr", year)
                Using reader = cmd.ExecuteReader()
                    While reader.Read() : result(reader.GetInt32(0)) = reader.GetInt32(1) : End While
                End Using
            End Using
            Return If(Not result.IsEmpty, result, Nothing)

        Catch ex As System.Exception
            _dbg(" ├ 錯誤", $"{fPath} {year}: {ex.Message}")
        Finally : If _iLikeNoisy Then _dbg(" ├ 結束")
        End Try
        Return Nothing

    End Function
    Friend Function DbGetAttachMailList(fPath As String) As AttachMailListDbResult
        ' ---------------------------------------------------------------
        ' DbGetAttachMailList — 讀取 attach_maillist WHERE folder_path=? 的所有行
        ' 回傳 Nothing 表示 DB 中無此資料夾的郵件記錄
        ' ---------------------------------------------------------------
        If _iLikeNoisy Then _dbg(" ├ 開始")
        If _dbCache Is Nothing Then Return Nothing

        Try
            Dim result As New AttachMailListDbResult()
            Dim hasRecord As Boolean = False ' by Claude Sonnet 4.6, 2026/05/06: 補齊變數宣告
            Using cmd As New SqliteCommand("SELECT entry_id,subject,msg_size,received_time,sender_name,attach_count,pr_count_snap" &
                                           "  FROM attach_maillist WHERE folder_hash=@fh", _dbCache)

                ' 2026/06/12 by Gemini/Simon: 把 folder_path TEXT 改成 folder_hash INTEGER 儲存，並新增一個記憶體字典 _dictHashToPath 來反查雜湊對應的路徑
                cmd.Parameters.AddWithValue("@fh", FolderPathToHash64(fPath))
                hasRecord = False ' 移除 Dim，直接使用外層變數

                ' 2026/06/10 by Gemini/Simon: 優化SQLite儲存空間把Entry_id改成BLOB二進位儲存
                Using reader = cmd.ExecuteReader()
                    While reader.Read() ' pr_count_snap 整個 folder 共用同一值，每行都一樣，讀最後一次即可
                        hasRecord = True
                        result.Snap = If(reader.IsDBNull(6), -1L, reader.GetInt64(6))
                        Dim eid = ByteArrayToHexString(reader.GetFieldValue(Of Byte())(0))
                        If eid.StartsWith("EMPTY_ATTACH_") Then Continue While

                        Dim mail As New MailItemInfo With {.EntryID = eid,
                                                           .Subject = If(reader.IsDBNull(1), "", reader.GetString(1)),
                                                           .Size = If(reader.IsDBNull(2), 0L, reader.GetInt64(2)),
                                                           .RcvTime = UnixSecondsToLocalTime(If(reader.IsDBNull(3), 0L, reader.GetInt64(3))),
                                                           .SenderName = If(reader.IsDBNull(4), "", reader.GetString(4)),
                                                           .AttachCount = If(reader.IsDBNull(5), 0, reader.GetInt32(5)),
                                                           .FolderPath = fPath}
                        result.Mails.Add(mail)
                    End While
                End Using
            End Using
            Return If(hasRecord, result, Nothing) ' by Gemini 3 Flash, 2026/05/06: 只要有紀錄(含 Dummy)就算命中

        Catch ex As System.Exception
            _dbg(" ├ 錯誤", $"{fPath}: {ex.Message}")
        Finally : If _iLikeNoisy Then _dbg(" ├ 結束")
        End Try
        Return Nothing

    End Function
    Friend Function DbGetAttachFilenames(entryId As String) As List(Of String)
        ' ---------------------------------------------------------------
        ' DbGetAttachFilenames — 讀取 attach_filenames WHERE entry_id=? 的一行
        ' 回傳 Nothing 表示 DB 中無此 EntryID
        ' ---------------------------------------------------------------
        If _iLikeNoisy Then _dbg(" ├ 開始")
        If _dbMail Is Nothing Then Return Nothing   ' 2026/06/21 by Simon/Claude Opus 4.8: attach_filenames 已搬至 _dbMail(OLAcacheMail.db)

        Try
            Using cmd As New SqliteCommand("SELECT filenames FROM attach_filenames WHERE entry_id=@eid", _dbMail)
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
    Friend Function DbGetMailInfo(fPath As String) As (Mails As List(Of (Mail As MailItemInfo, Topic As String)), Snap As Integer)?
        ' ---------------------------------------------------------------
        ' DbGetMailInfo — 讀取 mailinfo_list WHERE folder_path=? 的所有行
        ' 回傳 Nothing 表示 DB 中無此資料夾的郵件記錄
        ' 2026/06/10 by Gemini/Simon: 優化SQLite儲存空間把Entry_id改成BLOB二進位儲存
        ' 2026/06/11 by Gemini/Simon: 把 message_id 轉成 xxHash64，並同時改成 BLOB 儲存節省空間
        ' 2026/06/12 by Simon/Claude Opus 4.8: topic 改動態計算；sender_id→email；received_time 改 INTEGER
        ' ---------------------------------------------------------------
        If _iLikeNoisy Then _dbg(" ├ 開始")
        If _dbCache Is Nothing Then Return Nothing
        Try
            Dim result As New List(Of (Mail As MailItemInfo, Topic As String))(1024)
            Dim snap As Integer = -1
            Dim hasRecord As Boolean = False

            ' 2026/06/12: 新欄位順序 — entry_id(0),subject(1),msg_size(2),received_time INTEGER(3),
            '             sender_name(4),sender_id(5),msgid_hash(6),pr_count_snap(7)
            Using cmd As New SqliteCommand("SELECT entry_id,subject,msg_size,received_time,sender_name,sender_id,msgid_hash,pr_count_snap" &
                                           "  FROM mailinfo_list WHERE folder_hash=@fh", _dbCache)
                ' 2026/06/12 by Gemini/Simon: folder_hash 查詢
                cmd.Parameters.AddWithValue("@fh", FolderPathToHash64(fPath))

                ' 2026/06/12 by Simon/Claude Opus 4.8: received_time 改為 INTEGER (Unix秒)
                ' 2026/06/12 by Simon/Claude Opus 4.8: sender_id → email 反查
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        hasRecord = True
                        snap = If(reader.IsDBNull(7), -1L, reader.GetInt64(7))
                        Dim eid = ByteArrayToHexString(reader.GetFieldValue(Of Byte())(0))
                        If eid.StartsWith("EMPTY_BASIC_") Then Continue While

                        Dim subj = If(reader.IsDBNull(1), "", reader.GetString(1))
                        Dim mail As New MailItemInfo With {.EntryID = eid,
                                                           .Subject = subj,
                                                           .Size = If(reader.IsDBNull(2), 0L, reader.GetInt64(2)),
                                                           .RcvTime = UnixSecondsToLocalTime(If(reader.IsDBNull(3), 0L, reader.GetInt64(3))),
                                                           .SenderName = If(reader.IsDBNull(4), "", reader.GetString(4)),
                                                           .SenderEmail = _dictSenderIdToEmail.GetValueOrDefault(If(reader.IsDBNull(5), 0, reader.GetInt32(5)), ""),
                                                           .FolderPath = fPath,
                                                           .MsgIDhash = If(reader.IsDBNull(6), "", ByteArrayToHexString(reader.GetFieldValue(Of Byte())(6)))}
                        result.Add((mail, GetCleanSubject(subj)))
                    End While
                End Using
            End Using

            If hasRecord Then ' by Gemini 3 Flash, 2026/05/06: 只要有紀錄(含 Dummy)就算命中
                _dbg(" ├ 命中 SSD", $"{ExtractFolderName(fPath)} | 取得 {result.Count} 筆 | Snap={snap}")
                Return (result, snap)
            End If
        Catch ex As System.Exception
            _dbg(" ├ 錯誤", $"{fPath}: {ex.Message}")
        Finally : If _iLikeNoisy Then _dbg(" ├ 結束")
        End Try
        Return Nothing
    End Function
    Friend Function DbGetMailInfoBatch(folderPaths As List(Of String)) As Dictionary(Of String, (Mails As List(Of (Mail As MailItemInfo, Topic As String)), Snap As Long))
        ' ---------------------------------------------------------------
        ' DbGetMailInfoBatch — 一次 SQL IN 查詢，批次讀回多個資料夾的 mailinfo_list
        ' 取代原來 300 個別查詢的瓶頸；由 PreLoadMailCacheAsync 呼叫。
        ' 回傳 key=folder_path, value=(Mails, Snap)；不做 snap 驗證，由呼叫端決定。
        ' 2026/05/11 by Simon/Claude: 優化B
        ' 2026/06/12 by Simon/Claude Opus 4.8: topic 改由 GetCleanSubject(subject) 動態計算
        ' ---------------------------------------------------------------
        Dim result As New Dictionary(Of String, (Mails As List(Of (Mail As MailItemInfo, Topic As String)), Snap As Long))(folderPaths.Count)
        If _dbCache Is Nothing OrElse folderPaths.Count = 0 Then Return result
        If _iLikeNoisy Then _dbg(" ├ 開始", $"批次查詢 {folderPaths.Count} 個路徑")

        Try
            ' 建立 Hash 清單與對照表, 2026/06/12
            Dim hashes As New List(Of Long)(folderPaths.Count)
            For Each p In folderPaths
                hashes.Add(FolderPathToHash64(p))
            Next

            ' SQLite 預設 variable 上限 999，300 個路徑絕對安全
            Dim paramNames = Enumerable.Range(0, folderPaths.Count).Select(Function(i) "@fh" & i.ToString()).ToList()

            ' 2026/06/12 by Simon/Claude Opus 4.8: 更新 SELECT 欄位：移除 topic/sender_email/updated_at，加 sender_id；received_time 改 INTEGER
            ' 新欄位順序 — entry_id(0),folder_hash(1),subject(2),msg_size(3),received_time INTEGER(4),
            '             sender_name(5),sender_id(6),msgid_hash(7),pr_count_snap(8)
            Dim sql As String = "SELECT entry_id,folder_hash,subject,msg_size,received_time,sender_name,sender_id,msgid_hash,pr_count_snap" &
                                "  FROM mailinfo_list WHERE folder_hash IN (" & String.Join(",", paramNames) & ")"

            Using cmd As New SqliteCommand(sql, _dbCache)
                For i As Integer = 0 To folderPaths.Count - 1
                    cmd.Parameters.AddWithValue(paramNames(i), hashes(i))
                Next

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
    Friend Function DbGetSubFolderIDList(rootPath As String, isIncludeAll As Boolean) As List(Of FolderStatsDbRow)
        ' ---------------------------------------------------------------
        ' DbGetSubFolderIDList — [優化 BFS] 利用 LIKE 一次抓出整棵子樹的所有資料夾身分證
        ' ---------------------------------------------------------------
        If _iLikeNoisy Then _dbg(" ├ 開始")
        If _dbCache Is Nothing Then Return Nothing
        Try
            Dim result As New List(Of FolderStatsDbRow)(512)
            ' 過濾條件: 路徑以 rootPath 開頭，且 entry_id 不為空。若沒勾全選，則只抓 is_mail=1 的。
            Dim filter = If(isIncludeAll, "", " AND is_mail=1")
            ' 2026/06/13 by Simon/Claude Opus 4.8: 補 folder_count → 填 row.fc，供 IsSubtreeComplete 的「集合內子夾數 == fc」完整性檢查
            Dim sql = $"SELECT folder_path,entry_id,store_id,is_mail,has_chinese,folder_count FROM folder_stats " &
                      $"WHERE folder_path LIKE @fp || '%' AND entry_id IS NOT NULL" & filter

            Using cmd As New SqliteCommand(sql, _dbCache)
                cmd.Parameters.AddWithValue("@fp", rootPath)
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        ' 2026/06/10 by Gemini/Simon: 優化SQLite儲存空間把Entry_id改成BLOB二進位儲存
                        result.Add(New FolderStatsDbRow With {.path = reader.GetString(0),
                                                              .eid = If(reader.IsDBNull(1), "", ByteArrayToHexString(reader.GetFieldValue(Of Byte())(1))),
                                                              .sid = If(reader.IsDBNull(2), "", reader.GetString(2)),
                                                              .isMail = reader.GetInt32(3),
                                                              .hasCh = reader.GetInt32(4),
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
    Friend Function DbGetOrderedSubFolderIDs(parentPath As String, isIncludeAll As Boolean) As List(Of FolderStatsDbRow)
        ' ---------------------------------------------------------------
        ' DbGetOrderedSubFolderIDs — [優化 TreeView] 抓出直屬子目錄的身分證，並由 SQL 完成「英文優先」排序
        ' ---------------------------------------------------------------
        If _iLikeNoisy Then _dbg(" ├ 開始")
        If _dbCache Is Nothing Then Return Nothing

        Try
            Dim result As New List(Of FolderStatsDbRow)(512)
            ' SQL 邏輯: 
            '   1. 找出 folder_path 以 parentPath + "\" 開頭。
            '   2. 且不包含更深層的 "\" (代表是直屬子項) 。注意: 此邏輯在路徑分隔符不一致時需調整。
            '   3. 按照 has_chinese ASC (0=英, 1=中, 故英優先) 排序。
            Dim filter = If(isIncludeAll, "", " AND is_mail=1")

            ' 精確匹配直屬子目錄：利用 LENGTH + REPLACE 來算出層級
            ' 或是利用路徑字串特性：新的路徑長度應該是在 parent 之後且沒有多餘的層級
            ' 簡化做法：目前專案路徑是用 \ 分隔。
            Dim sql = "SELECT folder_path,entry_id,store_id,is_mail,has_chinese FROM folder_stats " &
                      " WHERE folder_path LIKE @fp || '\%' AND entry_id IS NOT NULL " & filter &
                      "   AND folder_path NOT LIKE @fp || '\%\%' " & ' 排除第二層以後的
                      " ORDER BY has_chinese ASC, folder_path ASC"

            Using cmd As New SqliteCommand(sql, _dbCache)
                cmd.Parameters.AddWithValue("@fp", parentPath)
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        ' 2026/06/10 by Gemini/Simon: 優化SQLite儲存空間把Entry_id改成BLOB二進位儲存
                        result.Add(New FolderStatsDbRow With {.path = reader.GetString(0),
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
    Friend Function DbGetAllFolderStats() As List(Of FolderStatsDbRow)
        ''' <summary>
        ''' 一次性唯讀撈出 folder_stats 全量名單
        ''' </summary>
        Dim list As New List(Of FolderStatsDbRow)()
        If _dbCache Is Nothing Then Return list

        Try
            ' 2026/06/12 by Simon/Claude Opus 4.8: 更新 SELECT 欄位順序，並移除 has_chinese（UI 只在 TreeView 用到，且不常變動，不適合放在全量查詢裡）
            Dim sql As String = "SELECT folder_path, entry_id, store_id, pr_count_snap, 
                                        mail_count, folder_count, folder_size, mail_count_all, folder_count_all, folder_size_all, is_mail FROM folder_stats"
            ' 建立並填入原有的結構
            Using cmd As New SqliteCommand(sql, _dbCache)
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        ' 2026/06/10 by Gemini/Simon: 優化SQLite儲存空間把Entry_id改成BLOB二進位儲存
                        Dim row As New FolderStatsDbRow() With {.path = reader.GetString(0),
                                                                .eid = If(reader.IsDBNull(1), "", ByteArrayToHexString(reader.GetFieldValue(Of Byte())(1))),
                                                                .sid = If(reader.IsDBNull(2), "", reader.GetString(2)),
                                                                .snap = If(reader.IsDBNull(3), -1, reader.GetInt64(3)),
                                                                .mc = If(reader.IsDBNull(4), -1L, reader.GetInt64(4)),
                                                                .fc = If(reader.IsDBNull(5), -1L, reader.GetInt64(5)),
                                                                .fs = If(reader.IsDBNull(6), -1L, reader.GetInt64(6)),
                                                                .mca = If(reader.IsDBNull(7), -1L, reader.GetInt64(7)),
                                                                .fca = If(reader.IsDBNull(8), -1L, reader.GetInt64(8)),
                                                                .fsa = If(reader.IsDBNull(9), -1L, reader.GetInt64(9)),
                                                                .isMail = If(reader.IsDBNull(10), -1, reader.GetInt32(10))}
                        list.Add(row)
                    End While
                End Using
            End Using
        Catch ex As System.Exception
            _dbg("DbGetAllFolderStats 發生錯誤", ex.Message)
        End Try
        Return list
    End Function

    Private Sub DbSaveMonthCountsSingle(fPath As String, year As Integer, monthCounts As ConcurrentDictionary(Of Integer, Integer))
        ' ---------------------------------------------------------------
        ' DbSaveMonthCountsSingle — 增量寫入單一 (folder_path, year) 的月份分布
        ' 在 GetMonthCountsForYearL3 完成 L3 COM 計算後立刻呼叫，不等待 SaveCache 按鈕。
        ' 使用獨立 Transaction 包住最多 12 筆，確保原子性。
        '   ├ 2026/04/09 新增 by Claude：解決月份快取只在記憶體、SaveCache 才寫 DB 的問題
        '   ├ 根本原因：若該 session 沒點過月份視圖就不 SaveCache，下次仍打 COM
        '   ├ 修正後：每次 L3 計算完月份後立刻持久化，下次 DB lazy 直接命中
        ' ---------------------------------------------------------------
        If _iLikeNoisy Then _dbg(" ├ 開始")
        If _dbCache Is Nothing OrElse monthCounts Is Nothing OrElse monthCounts.IsEmpty Then Return
        Try
            Using txn = _dbCache.BeginTransaction()
                Using cmd As New SqliteCommand(
                    "INSERT OR REPLACE INTO month_counts (folder_hash,year,month,count) VALUES (@fh,@yr,@mo,@cnt)", _dbCache, txn)
                    ' 2026/06/12 by Gemini/Simon: 把 folder_path TEXT 改成 folder_hash INTEGER 儲存，並新增一個記憶體字典 _dictHashToPath 來反查雜湊對應的路徑
                    ' 2026/06/12 by Simon/Claude Opus 4.8: 移除 updated_at（只寫不讀）
                    cmd.Parameters.Add("@fh", SqliteType.Integer).Value = FolderPathToHash64(fPath)
                    cmd.Parameters.Add("@yr", SqliteType.Integer).Value = year
                    cmd.Parameters.Add("@mo", SqliteType.Integer)
                    cmd.Parameters.Add("@cnt", SqliteType.Integer)
                    For Each mo In monthCounts
                        cmd.Parameters("@mo").Value = mo.Key
                        cmd.Parameters("@cnt").Value = mo.Value
                        cmd.ExecuteNonQuery()
                    Next
                End Using
                txn.Commit()
            End Using
            _dbg(" ├ 更新", $"{fPath} {year} → {monthCounts.Count} 個月寫入 DB")

        Catch ex As System.Exception
            _dbg(" ├ 錯誤", $"{fPath} {year}: {ex.Message}")
        Finally : If _iLikeNoisy Then _dbg(" ├ 結束")
        End Try

    End Sub
    Private Sub DbDeleteMailInfoByPath(fPath As String)
        ' ---------------------------------------------------------------
        ' DbDeleteMailInfoByPath — 刪除郵件後立即清除指定 fPath 的 mailinfo_list 記錄
        ' 只針對被刪除郵件所在的資料夾，不影響其他 fPath
        ' 配合 InvalidateMailCache 一起使用，確保 DB lazy load 不會回傳舊資料
        ' 2026/05/11 by Claude Sonnet 4.6
        ' ---------------------------------------------------------------
        If _dbCache Is Nothing OrElse String.IsNullOrEmpty(fPath) Then Return

        Try
            ' 2026/06/12 by Gemini/Simon: 把 folder_path TEXT 改成 folder_hash INTEGER 儲存，並新增一個記憶體字典 _dictHashToPath 來反查雜湊對應的路徑
            Using cmd As New SqliteCommand("DELETE FROM mailinfo_list WHERE folder_hash=@fh", _dbCache)
                cmd.Parameters.AddWithValue("@fh", FolderPathToHash64(fPath))
                Dim rows = cmd.ExecuteNonQuery()
                _dbg("DbDeleteMailInfoByPath", $"{ExtractFolderName(fPath)}: 清除 {rows} 筆")
            End Using
        Catch ex As Exception
            _dbg("DbDeleteMailInfoByPath 錯誤", $"{ExtractFolderName(fPath)}: {ex.Message}")
        End Try
    End Sub
    Private Sub DbPurgeFolderMailRows(fPath As String, Optional includeAttachFilenames As Boolean = False)
        ' ---------------------------------------------------------------
        ' DbPurgeFolderMailRows — 刪除單一資料夾在逐封郵件表的全部列 (mailinfo_list/attach_maillist/month_counts，選擇性含 attach_filenames)。
        '   用於「資料夾還在但內含郵件有增刪」時清掉死列(失效 entryID)，維持「同一資料夾的 basic/attach 列共用單一 snap」不變量，根除讀取端混 snap 幽靈郵件。
        '   與 CleanupOrphanPath 的差異：那個是整夾消失才連 folder_stats 全表一起刪；本函式只清逐封郵件列，不動 folder_stats。
        ' 2026/06/20 by Simon/Claude: 取代原 RenewAttachMailList 三路比對
        ' ---------------------------------------------------------------
        If _dbCache Is Nothing Then Return
        Dim fh = FolderPathToHash64(fPath)
        Try
            Using txn As SqliteTransaction = _dbCache.BeginTransaction()
                For Each tbl In {"mailinfo_list", "attach_maillist", "month_counts"}
                    Using c As New SqliteCommand($"DELETE FROM {tbl} WHERE folder_hash=@fh", _dbCache, txn)
                        c.Parameters.AddWithValue("@fh", fh) : c.ExecuteNonQuery()
                    End Using
                Next
                txn.Commit()
            End Using
        Catch ex As System.Exception
            _dbg("DbPurgeFolderMailRows 錯誤", $"{fPath}: {ex.Message}")
        End Try
        ' 2026/06/21 by Simon/Claude: attach_filenames 已搬至 OLAcacheMail.db(_dbMail)，跨檔不能掛 _dbCache txn，改獨立交易刪除
        If includeAttachFilenames Then SimDbDeleteAttachFilenamesByFolder(fPath)
    End Sub

    Private Function SimDbDeleteAttachFilenamesByFolder(fPath As String) As Integer
        ' ---------------------------------------------------------------
        ' SimDbDeleteAttachFilenamesByFolder — 刪除單一資料夾在 attach_filenames(OLAcacheMail.db/_dbMail) 的全部列。
        '   獨立 _dbMail 交易，供 DbPurgeFolderMailRows(RenewCache 狀況 A) 與 CleanupOrphanPath(整夾消失) 兩處呼叫。
        '   回傳刪除列數(供 CleanupOrphanPath 統計顯示)。
        ' 2026/06/21 by Simon/Claude: Part 2 拆檔耦合點 — 原本掛在 _dbCache txn 的 attach_filenames DELETE 抽到此 helper
        ' ---------------------------------------------------------------
        If _dbMail Is Nothing Then Return 0
        Dim fh = FolderPathToHash64(fPath) : Dim n As Integer = 0
        Try
            Using txn = _dbMail.BeginTransaction()
                Using c As New SqliteCommand("DELETE FROM attach_filenames WHERE folder_hash=@fh", _dbMail, txn)
                    c.Parameters.AddWithValue("@fh", fh) : n = c.ExecuteNonQuery()
                End Using
                txn.Commit()
            End Using
        Catch ex As System.Exception
            _dbg("SimDbDeleteAttachFilenamesByFolder 錯誤", $"{fPath}: {ex.Message}")
        End Try
        Return n
    End Function
    Private Function SimDbDeleteMailRowsByEntryIds(entryIds As ICollection(Of String), Optional includeAttachFilenames As Boolean = True) As Integer
        ' ---------------------------------------------------------------
        ' SimDbDeleteMailRowsByEntryIds — ② Surgical：依「已刪 entryID 集合」精準清除兩張「逐封讀取極貴」的快取。
        '   記憶體：一律 TryRemove _cacheAttachFilename + _cacheSimHash(兩者 key 皆 EntryID 字串)。
        '   DB(_dbMail)：mail_simhash 一律刪(無 folder_hash，只能靠 entryID)；
        '                attach_filenames 視 includeAttachFilenames —— 狀況 A(夾內增刪)傳 True 逐封刪；
        '                CleanupOrphanPath(整夾消失)傳 False(已由 SimDbDeleteAttachFilenamesByFolder 按 folder_hash 高效刪過，免重複)。
        '   只清失效的那幾封，存活郵件的昂貴快取保留(免重讀內文/附件)。回傳 mail_simhash 刪除列數(供 log)。
        ' 2026/06/22 by Simon/Claude: ② Surgical 策略 — 嚴格清除失效 entryID，杜絕昂貴快取死列永久累積
        ' ---------------------------------------------------------------
        If entryIds Is Nothing OrElse entryIds.Count = 0 Then Return 0
        For Each eid In entryIds : _cacheAttachFilename.TryRemove(eid, Nothing) : _cacheSimHash.TryRemove(eid, Nothing) : Next
        If _dbMail Is Nothing Then Return 0
        Dim nSh As Integer = 0
        Try
            Using txn = _dbMail.BeginTransaction()
                Using cSh As New SqliteCommand("DELETE FROM mail_simhash WHERE entry_id=@eid", _dbMail, txn),
                      cAf As New SqliteCommand("DELETE FROM attach_filenames WHERE entry_id=@eid", _dbMail, txn)
                    cSh.Parameters.Add("@eid", SqliteType.Blob) : cAf.Parameters.Add("@eid", SqliteType.Blob)
                    For Each eid In entryIds
                        Dim blob = HexStringToByteArray(eid)
                        cSh.Parameters("@eid").Value = blob : nSh += cSh.ExecuteNonQuery()
                        If includeAttachFilenames Then cAf.Parameters("@eid").Value = blob : cAf.ExecuteNonQuery()
                    Next
                End Using
                txn.Commit()
            End Using
        Catch ex As System.Exception
            _dbg("SimDbDeleteMailRowsByEntryIds 錯誤", ex.Message)
        End Try
        Return nSh
    End Function
    Private Function DbGetFolderEntryIds(fPath As String) As List(Of String)
        ' ---------------------------------------------------------------
        ' DbGetFolderEntryIds — 撈出單一資料夾在 mailinfo_list 的全部 entry_id(轉 hex 字串)。
        '   mailinfo_list 是該夾「逐封郵件」的權威清單(attach_filenames/mail_simhash 的 entryID 皆其子集)，
        '   供 ② Surgical 差集找已刪郵件。注意：呼叫端若會 DELETE mailinfo_list，務必在 DELETE 之前呼叫本函式。
        ' 2026/06/22 by Simon/Claude: ② Surgical 輔助
        ' ---------------------------------------------------------------
        Dim result As New List(Of String)()
        If _dbCache Is Nothing Then Return result
        Dim fh = FolderPathToHash64(fPath)
        Try
            Using cmd As New SqliteCommand("SELECT entry_id FROM mailinfo_list WHERE folder_hash=@fh", _dbCache)
                cmd.Parameters.AddWithValue("@fh", fh)
                Using r = cmd.ExecuteReader()
                    While r.Read()
                        If Not r.IsDBNull(0) Then result.Add(ByteArrayToHexString(r.GetFieldValue(Of Byte())(0)))
                    End While
                End Using
            End Using
        Catch ex As System.Exception
            _dbg("DbGetFolderEntryIds 錯誤", $"{fPath}: {ex.Message}")
        End Try
        Return result
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
    Private Async Function DbShowDbFileStat() As Task
        ''' <summary>
        ''' 點擊 ListView6 "DB 檔案大小" 時呼叫。2026/06/21 起改為「依 db 分區塊」：
        '''   OLAcache.db 與 OLAcacheMail.db 各跑一次兩段式統計(per-file 數學自洽)，最後輸出兩檔合計。
        '''   階段一 (UI thread): PRAGMA 頁面資訊 + 各表筆數；階段二 (Task.Run): CAST AS BLOB 算淨重 + 比例估算實體。
        ''' 因 e_sqlite3.dll 預設未啟用 SQLITE_ENABLE_DBSTAT_VTAB，改用「淨重比例分配」估算(相對排名準，絕對值 ±20%)。
        ''' 2026/06/13 by Simon/Claude Opus 4.8 / 2026/06/21 by Simon/Claude: 拆 OLAcacheMail.db 後改雙檔分區塊
        ''' </summary>
        If _dbCache Is Nothing Then Return

        ' Task.Run 內呼叫 _dbg 的 stack trace 會抓到編譯器生成的 lambda 名稱，故預先封裝 forwarder
        ' 直接走 DebugForm.AddMessage3 並傳入 forcedCaller，與 _dbg() 的 Release-build 行為等效
        Dim _dbgFwd As Action(Of String, String) = Sub(a, b)
                                                       If _isDebugMode Then DebugForm.AddMessage3(a, b, "DbShowDbFileStat")
                                                   End Sub
        _dbgFwd("開始", "📊DB 檔案大小")

        Try
            ' 2026/06/21 by Simon/Claude: 兩個 db 各一段(淨重佔比/估算實體的分母不可跨檔)，依序輸出後再合計
            Dim r1 = Await ShowOneDbFileStat(_dbCache, _dbCachePath, "OLAcache.db", _dbgFwd)
            Dim file2 As Single = 0F, net2 As Single = 0F
            If _dbMail IsNot Nothing Then
                Dim r2 = Await ShowOneDbFileStat(_dbMail, _dbMailPath, "OLAcacheMail.db", _dbgFwd)
                file2 = r2.fileMB : net2 = r2.netMB
            End If

            ' ----- 兩檔合計 -----
            _dbgFwd("", "═══════════ 兩檔合計 ═══════════")
            _dbgFwd(" │", $" 檔案實體合計: {(r1.fileMB + file2):F2} MB    /    純資料淨重合計: {(r1.netMB + net2):F2} MB")
            _dbgFwd("結束", "DB 檔案大小")

        Catch ex As System.Exception
            _dbgFwd(" ├ 錯誤", $"DbShowDbFileStat: {ex.Message}")
        End Try
    End Function
    Private Async Function DbShowTableStat(tableName As String) As Task
        ''' <summary>
        ''' 深度分析快取表：動態計算精準 Bytes 淨容量、並還原真實的實體 SSD 大小誤差 (壓縮邏輯)
        ''' </summary>

        ' 2026/06/21 by Simon/Claude: attach_filenames/mail_simhash 住 OLAcacheMail.db(_dbMail)，依表名路由連線；其餘走 _dbCache
        Dim conn As SqliteConnection = If(tableName = "attach_filenames" OrElse tableName = "mail_simhash", _dbMail, _dbCache)
        If conn Is Nothing OrElse String.IsNullOrEmpty(tableName) Then Return

        _dbg("開始", $"[📊{tableName}]")
        Try
            ' 1. 動態取得欄位 MetaData
            Dim cols As New List(Of (Cid As Integer, Name As String, Type As String, Pk As String, Nn As String))
            Using cmd As New SqliteCommand($"PRAGMA table_info([{tableName}])", conn)
                Using rd = cmd.ExecuteReader()
                    While rd.Read()
                        cols.Add((Convert.ToInt32(rd("cid")), rd("name").ToString(), rd("type").ToString(), If(Convert.ToInt32(rd("pk")) > 0, "★", ""), If(Convert.ToInt32(rd("notnull")) > 0, "Y", "N")))
                    End While
                End Using
            End Using
            If cols.Count = 0 Then _dbg(" ├ 錯誤", $"找不到表格 [{tableName}]") : Return

            ' 2. 修正版 SQL：強制 CAST AS BLOB 算真實 Bytes，破解中文字元數陷阱！
            Dim sbSql As New System.Text.StringBuilder("SELECT COUNT(*)")
            For Each c In cols : sbSql.Append($", SUM(length(CAST([{c.Name}] AS BLOB)))") : Next
            sbSql.Append($" FROM [{tableName}]")

            Dim rowCount As Long = 0
            Dim totalNetMB As Single = 0
            Dim colSizes As New Dictionary(Of String, Single)
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

            ' 3. 嘗試讀取 dbstat 獲取真實的實體 SSD 佔用 (包含碎片、RowHeader 與 B-Tree Index 結構)
            Dim hasDbStat As Boolean = True
            Dim indexMB As Single = 0
            Dim physicalMB As Single = 0
            Try
                Dim statSql = $"SELECT name, SUM(pgsize)/1024/1024 FROM dbstat WHERE name='{tableName}' OR name IN (SELECT name FROM sqlite_master WHERE type='index' AND tbl_name='{tableName}') GROUP BY name;"
                Using cmd As New SqliteCommand(statSql, conn)
                    Using rd = cmd.ExecuteReader()
                        While rd.Read()
                            If rd.GetString(0) = tableName Then physicalMB = rd.GetDouble(1) Else indexMB += rd.GetDouble(1)
                        End While
                    End Using
                End Using
            Catch ex As System.Exception : hasDbStat = False : End Try ' 若編譯不支援 dbstat 則忽略

            ' 4. 輸出統計與殘酷的比對
            _dbg(" ├", $"總資料筆數: {rowCount} 筆")
            _dbg(" ├", $" {"欄位名稱".PadRight(16)}{"型態".PadRight(12)}欄位資料淨重")
            For Each c In cols : _dbg(" │", $" {$"[{c.Name}]".PadRight(17)}  {c.Type.PadRight(8)}: {colSizes(c.Name).ToString("F3")} MB") : Next
            _dbg(" │", $" 所有欄位純資料淨重 : {totalNetMB.ToString("F3")} MB (程式寫入的真正大小)")

            If hasDbStat Then
                _dbg(" │", $" 表格主體實際佔用 : {physicalMB.ToString("F3")} MB (包含 B-Tree 碎片與 Row Header)")
                _dbg(" │", $" 關聯索引佔用    : {indexMB.ToString("F3")} MB (Index 樹狀結構)")
                _dbg(" │", $" 總計空間佔用    : {(physicalMB + indexMB).ToString("F3")} MB 👈 這才是真凶！")
            Else
                _dbg(" │ 提醒", "(目前未啟用 dbstat 模組，無法精準測量索引與碎片開銷，通常佔用為淨重的 1.5~3倍)")
            End If
            _dbg("結束", $"[{tableName}]")

        Catch ex As System.Exception : _dbg(" ├ 錯誤", $"分析 {tableName} 失敗: {ex.Message}") : End Try
    End Function
    Private Async Function ShowOneDbFileStat(conn As SqliteConnection, dbPath As String, dbLabel As String, dbgFwd As Action(Of String, String)) As Task(Of (fileMB As Single, netMB As Single))
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
        Dim _displayOrder = New String() {"folder_stats", "senders", "mailinfo_list", "year_counts", "month_counts", "attach_maillist", "attach_filenames", "mail_simhash"}
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
#End Region

#Region "■ 轉換用輔助函式"
    Private Function HexStringToByteArray(idStr As String) As Byte()
        ''' <summary>
        ''' 安全將 EntryID 字串轉為 BLOB (Byte Array)。支援常規 Hex 與 EMPTY_ 哨兵字串。
        ''' 2026/06/10 by Gemini/Simon: 優化SQLite儲存空間把Entry_id改成BLOB二進位儲存
        ''' </summary>
        If String.IsNullOrEmpty(idStr) Then Return Array.Empty(Of Byte)()

        ' 檢查是否為系統內建的哨兵字串 (例如 EMPTY_BASIC_ 或 EMPTY_ATTACH_)
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
#End Region

End Class
