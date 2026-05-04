Imports System.Collections.Concurrent
Imports System.Text.Json
Imports Microsoft.Data.Sqlite
Imports Microsoft.Office.Interop
'Imports System.Runtime.Intrinsics.X86

' ==============================================================
' Form1_SQLite2.vb  —  SQLite 持久化快取層
' ==============================================================
' 目的: 把記憶體 ConcurrentDictionary 快取持久化到 SSD，下次啟動可快速重建
' 架構:
'   往上往下串接在Layer1/Layer2/Layer2.5和Layer3之間 (其他層呼叫完全不知道 SQLite 的存在)
'   Form1_SQLite2.vb  (本檔)
'   - InitDatabase()                            ' Form1_Load 呼叫，建 connection + CREATE TABLE IF NOT EXISTS
'   - CloseDatabase()                           ' FormClosing 呼叫
'   - LoadCachesFromDB()               ' LoadCache 按鈕手動讀出：Bulk Load，輸出詳細 _dbg 分項
'   - SaveCachesToDB()                 ' SaveCache 按鈕手動存入：① CleanupOrphanFolderPath → ② 批次寫入四張表
'   - CleanupOrphanFolderPath(livePaths)        ' 清除 DB 中已不存在的 folder_path (原 PurgeStaleFolders)，SaveCache 時順帶呼叫
'   - RenewCacheToDB(includeSize As Boolean)   ' RenewCache 按鈕：Phase1~6 完整更新 (2026-04-09 新增) 
'   - RenewAttachMailList(folder, fPath:=fPath)   ' 三路比對更新 attach_maillist (2026-04-09 新增) 
'
'   - DbGetFolderStats(folderPath)              ' folder_stats 單行查詢
'   - DbGetMailBasic(folderPath)                ' mail_basic WHERE folder_path=? 全部行
'   - DbGetAttachFilenames(entryId)             ' mail_attachments 單行查詢
'   - DbGetYearCountsForFolder(folderPath)      ' year_counts WHERE folder_path=? 全部行
'   - DbGetMonthCountsForFolder(cacheKey)       ' 2026-04-09 新增，cacheKey = FolderPath_year
'   - GetDBSummary() → (fc, mb, at, yc, mc, basic, kb) ' DB 統計摘要 (六張表行數 + 檔案 KB) 
' ---------------------------------------------------------------
'
'   五張表結構 (2026-04-09 新增 month_counts) 合一個 cache.db (LocalAppData)
'   六張表結構 (2026-04-22 新增 basic_maillist) 合一個 cache.db (LocalAppData)
'       folder_stats        (folder_path PK, mail_count, mail_count_all, folder_count, folder_count_all,
'                            folder_size, folder_size_all, content_count_snapshot, updated_at)
'       year_counts         (folder_path+year PK, count, updated_at)        
'       month_counts        (folder_path+year+month PK, count, updated_at)  
'       attach_maillist     (entry_id PK, folder_path, subject, msg_size, received_time, sender_name,
'                            attach_count, item_count_snap, updated_at)  ← 專供 Tab3 尋找附件使用
'       attach_filenames    (entry_id PK, folder_path, filenames TEXT JSON, msg_size, updated_at)
'       basic_maillist      (entry_id PK, folder_path, subject, msg_size, received_time, sender_name,
'                            topic, item_count_snap, updated_at)         ← 專供 Tab4/Tab5 系列與重複郵件使用
'                           
' 設計決策 (2026-04-06):
'   1. 跨表 Transaction 保證原子性，一個 Connection 管理最簡單
'   2. 手動控制 (SaveCache / LoadCache 按鈕)，Debug 階段方便目視確認正確性
'      正式版再切換成 Layer2.5 lazy SELECT + 增量寫入
'   3. content_count_snapshot 存 _cacheMailCount[path] 的值 (即 PR_CONTENT_COUNT 的讀取結果)
'      Load 後可快速判斷快取是否仍有效，完全不需要呼叫任何 COM
'   4. MailItemInfo 欄位以文字儲存；List(Of String) 附件名稱序列化為 JSON array
'   5. _cacheFolderTree / _cacheSubTreeList 含 COM 物件，永遠不寫入 SQLite
'   6. LoadFolderStatsInner 使用 TryAdd：若記憶體已有值 (Layer2.5 已讀過)，保留記憶體版本
'      若想強制以 DB 為準 (完整重置)，改用直接賦值 _cacheMailCount(path) = ...
'   7. (2026-04-22) 拆分 attach_maillist 與 basic_maillist：保持 Tab3 與 Tab4/5 邏輯與資料邊界獨立。
' ==============================================================
' ---------------------------------------------------------------

Partial Class Form1

#Region "■ 私有成員"
    Private _db As SqliteConnection = Nothing
    ' by Gemini, 2026/04/10: 將資料庫路徑移至 LocalAppData，避免與 Dropbox 同步衝突導致檔案鎖定與 Explorer 卡頓
    'Private ReadOnly _dbPath As String = IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OutlookAssistant\Cache", "OLAcache.db")
    Private _dbPath As String   ' 移除 ReadOnly 與靜態初始化

    ' DB Row 結構 (供 Form1_Outlook.vb 的 Layer2.5 函數使用)
    Friend Class FolderStatsDbRow
        ' folder_stats 一行的讀出結果；-1 代表該欄位在 DB 中為 NULL 或尚未寫入
        Public mc As Integer = -1       ' mail_count
        Public mca As Integer = -1      ' mail_count_all
        Public fc As Integer = -1       ' folder_count
        Public fca As Integer = -1      ' folder_count_all
        Public fs As Long = -1          ' folder_size
        Public fsa As Long = -1         ' folder_size_all
        Public snap As Integer = -1     ' content_count_snapshot (= PR_CONTENT_COUNT at save time)
        Public path As String = ""      ' folder_path        ' by Gemini 3.0 flash, 2026/04/16: 新增路徑標識，供 GetSubtreeToList Tuple 重建使用

        ' by Gemini, 2026/04/10: 新增身分標識與排序標籤，供 TreeView/BFS 持久化優化使用
        Public eid As String = ""       ' entry_id
        Public sid As String = ""       ' store_id
        Public isMail As Integer = -1   ' is_mail (0/1)
        Public hasCh As Integer = -1    ' has_chinese (0/1)
    End Class
    Friend Class AttachMailListDbResult
        ' attach_maillist WHERE folder_path=? 的讀出結果
        Public Snap As Integer = -1
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
        _dbPath = IO.Path.Combine(cacheDir, "OLAcache.db")

        Try
            ' by Gemini, 2026/04/10: 確保資料庫目錄存在
            Dim dbDir = IO.Path.GetDirectoryName(_dbPath)
            If Not IO.Directory.Exists(dbDir) Then IO.Directory.CreateDirectory(dbDir)

            _db = New SqliteConnection($"Data Source={_dbPath};Mode=ReadWriteCreate;Cache=Shared")
            _db.Open()
            _dbg("", $"已開啟: {_dbPath}")

            Using cmd As New SqliteCommand(BuildSQLiteTableString(), _db)
                cmd.ExecuteNonQuery()
            End Using

            Try
                Using cmd As New SqliteCommand(
                    "ALTER TABLE folder_stats ADD COLUMN entry_id TEXT;" &
                    "ALTER TABLE folder_stats ADD COLUMN store_id TEXT;" &
                    "ALTER TABLE folder_stats ADD COLUMN is_mail INTEGER DEFAULT -1;" &
                    "ALTER TABLE folder_stats ADD COLUMN has_chinese INTEGER DEFAULT -1;", _db)
                    cmd.ExecuteNonQuery()
                End Using
            Catch ex As System.Exception
                ' 資料行已存在時會拋出例外，安全忽略 (by Gemini, 2026/04/10)
            End Try

            Try ' 2026/05/06 by Claude: basic_maillist 新增 Tab5 去重欄位
                Using cmd As New SqliteCommand(
                    "ALTER TABLE basic_maillist ADD COLUMN message_id TEXT;" &
                    "ALTER TABLE basic_maillist ADD COLUMN sender_email TEXT;", _db)
                    cmd.ExecuteNonQuery()
                End Using
            Catch ex As System.Exception
                ' 欄位已存在，安全略過
            End Try

            ' by Claude Sonnet 4.6, 2026/05/06: Root Cause A 一次性資料清理 migration
            ' 舊版 Bug：SaveBasicMailInfoInner 將 _cacheBasicMailInfo 的 key (fPath|True/False)
            ' 直接當成 folder_path 寫入 DB，導致資料庫存有 "inbox|True" 這類帶後綴的路徑。
            ' 這些資料不會被 CleanupOrphanPath 自動清除，須一次性主動刪除。
            ' LIKE 查詢效能可接受，且每次啟動若無污染資料，DELETE 影響列數 = 0，開銷極小。
            Try
                Dim deletedRows As Integer = 0
                Using cmd As New SqliteCommand(
                    "DELETE FROM basic_maillist WHERE folder_path LIKE '%|True' OR folder_path LIKE '%|False'", _db)
                    deletedRows = cmd.ExecuteNonQuery()
                End Using
                If deletedRows > 0 Then
                    _dbg("", $"一次性清理：已刪除 {deletedRows} 筆 basic_maillist 污染資料 (folder_path 含 |True/|False 後綴)")
                End If
            Catch ex As System.Exception
                _dbg("", $"一次性清理 migration 失敗 (非致命): {ex.Message}")
            End Try

            _dbg("", "資料表確認完成")

            timerSaveCache.Interval = 1 * 60 * 1000 ' 每60sec自動保存一次快取資料到磁碟
            timerSaveCache.Start()                  ' 啟動定時快取保存

        Catch ex As System.Exception
            _dbg("       ├ 錯誤", ex.Message) ' by Gemini, 2026/04/11: Level 3
            _db = Nothing   ' 出錯就設 Nothing，後續所有 SQLite 操作因此自動跳過
        Finally : _dbg("結束") ' by Gemini, 2026/04/11: 修正對應開始層級 Level 0
        End Try

    End Sub
    Private Sub CloseDatabase()
        ' ---------------------------------------------------------------
        ' CloseDatabase — FormClosing 時呼叫，安全關閉 SQLite 連線
        ' ---------------------------------------------------------------
        _dbg("開始")

        If _db Is Nothing Then Return

        Try
            _db.Close() : _db.Dispose() : _db = Nothing

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

        Return "
                CREATE TABLE IF NOT EXISTS folder_stats ( 
                    folder_path             TEXT    PRIMARY KEY,
                    mail_count              INTEGER,
                    mail_count_all          INTEGER,
                    folder_count            INTEGER,
                    folder_count_all        INTEGER,
                    folder_size             INTEGER,
                    folder_size_all         INTEGER,
                    content_count_snapshot  INTEGER,
                    entry_id                TEXT,
                    store_id                TEXT,
                    is_mail                 INTEGER,
                    has_chinese             INTEGER,
                    updated_at              TEXT
                                                        );
                CREATE TABLE IF NOT EXISTS attach_maillist (
                    entry_id        TEXT    PRIMARY KEY,
                    folder_path     TEXT    NOT NULL,
                    subject         TEXT,
                    msg_size        INTEGER,
                    received_time   TEXT,
                    sender_name     TEXT,
                    attach_count    INTEGER,
                    item_count_snap INTEGER,
                    updated_at      TEXT
                                                        );
                CREATE TABLE IF NOT EXISTS basic_maillist (
                    entry_id        TEXT    PRIMARY KEY,
                    folder_path     TEXT    NOT NULL,
                    subject         TEXT,
                    msg_size        INTEGER,
                    received_time   TEXT,
                    sender_name     TEXT,
                    topic           TEXT,
                    message_id      TEXT,
                    sender_email    TEXT,
                    item_count_snap INTEGER,
                    updated_at      TEXT
                                                        );
                CREATE INDEX IF NOT EXISTS idx_mb_folder ON attach_maillist(folder_path);
                CREATE INDEX IF NOT EXISTS idx_basic_folder ON basic_maillist(folder_path);
                CREATE TABLE IF NOT EXISTS attach_filenames (
                    entry_id        TEXT    PRIMARY KEY,
                    folder_path     TEXT    NOT NULL,
                    filenames       TEXT,
                    msg_size        INTEGER,
                    updated_at      TEXT
                                                        );
                CREATE INDEX IF NOT EXISTS idx_ma_folder ON attach_filenames(folder_path);
                CREATE TABLE IF NOT EXISTS year_counts (
                    folder_path     TEXT    NOT NULL,
                    year            INTEGER NOT NULL,
                    count           INTEGER NOT NULL,
                    updated_at      TEXT,
                    PRIMARY KEY (folder_path, year)
                                                        );
                CREATE INDEX IF NOT EXISTS idx_yc_folder ON year_counts(folder_path);
                CREATE TABLE IF NOT EXISTS month_counts (
                    folder_path TEXT    NOT NULL,
                    year        INTEGER NOT NULL,
                    month       INTEGER NOT NULL,
                    count       INTEGER NOT NULL,
                    updated_at  TEXT,
                    PRIMARY KEY (folder_path, year, month)
                                                        );"
        ' 注意：month_counts 的舊版 schema 遷移 (cache_key → 三欄 PK) 在 InitDatabase() 中一次性處理，
        ' 不在此處 DROP TABLE，避免每次啟動都清空已存資料。

    End Function
    Private Function GetDBSummary() As (fc As Integer, mb As Integer, at As Integer, yc As Integer, mc As Integer, basic As Integer, kb As Long, lastTs As String)
        ' ---------------------------------------------------------------
        ' GetDBSummary — 取得 DB 統計摘要，供按鈕顯示
        ' 回傳 (folder_stats, attach_maillist, attach_filenames, year_counts, month_counts, basic_maillist, KB, lastTs)
        ' 2026/04/09 新增 mc = month_counts 行數
        ' 2026/04/10 新增 lastTs = 最後 updated_at 時間
        ' 2026/04/22 by Gemini 3 Flash: 新增 basic = basic_maillist 行數
        ' ---------------------------------------------------------------
        If _db Is Nothing Then Return (0, 0, 0, 0, 0, 0, 0L, "N/A")

        Try
            Dim fc, mb, at, yc, mcount, basicCount As Integer : Dim lastTs As String = "N/A"
            Using cmd As New SqliteCommand("SELECT COUNT(*) FROM folder_stats", _db) : fc = Convert.ToInt32(cmd.ExecuteScalar()) : End Using
            Using cmd As New SqliteCommand("SELECT COUNT(*) FROM attach_maillist", _db) : mb = Convert.ToInt32(cmd.ExecuteScalar()) : End Using
            Using cmd As New SqliteCommand("SELECT COUNT(*) FROM basic_maillist", _db) : basicCount = Convert.ToInt32(cmd.ExecuteScalar()) : End Using
            Using cmd As New SqliteCommand("SELECT COUNT(*) FROM attach_filenames", _db) : at = Convert.ToInt32(cmd.ExecuteScalar()) : End Using
            Using cmd As New SqliteCommand("SELECT COUNT(*) FROM year_counts", _db) : yc = Convert.ToInt32(cmd.ExecuteScalar()) : End Using
            Using cmd As New SqliteCommand("SELECT COUNT(*) FROM month_counts", _db) : mcount = Convert.ToInt32(cmd.ExecuteScalar()) : End Using

            ' 抓取最後一次成功的儲存時間字串 (取最大的 updated_at)
            Using cmd As New SqliteCommand("SELECT MAX(updated_at) FROM folder_stats", _db)
                Dim val = cmd.ExecuteScalar()
                If val IsNot Nothing AndAlso Not IsDBNull(val) Then lastTs = val.ToString()
            End Using

            Dim fi As New IO.FileInfo(_dbPath)
            Return (fc, mb, at, yc, mcount, basicCount, If(fi.Exists, fi.Length \ 1024L, 0L), lastTs)

        Catch ex As System.Exception
            _dbg("       ├ 錯誤", ex.Message) ' by Gemini, 2026/04/11: Level 3
            Return (0, 0, 0, 0, 0, 0, 0L, "Err")
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
            If IO.File.Exists(_dbPath) Then
                Dim timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss")
                Dim zipName = $"OLAcache_{timestamp}.zip"
                Dim zipPath = IO.Path.Combine(IO.Path.GetDirectoryName(_dbPath), zipName)

                _dbg("", $"正要壓縮備份至: {zipName}")

                ' 使用全名 (Fully Qualified Name) 避免 Imports 失敗, 並改用 Stream.CopyTo 避開擴展方法找不到的錯誤
                Using zipFileStream As New System.IO.FileStream(zipPath, System.IO.FileMode.Create)
                    Using archive As New System.IO.Compression.ZipArchive(zipFileStream, System.IO.Compression.ZipArchiveMode.Create)
                        Dim entry = archive.CreateEntry("OLAcache.db")
                        Using entryStream = entry.Open()
                            ' 加上 FileShare.ReadWrite 容許其他可能卡住的唯讀鎖，防止 IOException (by Gemini, 2026/04/10)
                            Using fs As New System.IO.FileStream(_dbPath, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite)
                                fs.CopyTo(entryStream)
                            End Using
                        End Using
                    End Using
                End Using

                ' 壓縮完後刪除原始 db 檔
                IO.File.Delete(_dbPath)
            End If

            ' 3. 重新建立
            InitDatabase()
            _dbg(" ├ 結束", "SSD 快取已重設，舊檔案已 Zip 備份") ' by Gemini, 2026/04/11: 修正對應開始層級 Level 1
        Catch ex As System.Exception
            _dbg("       ├ 錯誤", $"無法重置 SSD 資料庫: {ex.Message}")
            Throw
        End Try
    End Function
    Private Function SanitizeProfileName(name As String) As String
        ' Profile 名稱安全過濾
        ' CurrentProfileName 可能包含空格、單引號（Simon's Mail）、甚至斜線等非法路徑字元
        Dim invalid As Char() = IO.Path.GetInvalidFileNameChars()
        Return New String(name.Select(Function(c) If(Array.IndexOf(invalid, c) >= 0, "_"c, c)).ToArray())
    End Function
#End Region

#Region "■ 快取主控流程 (High-Level Cache Controllers)"
    Private Async Function SaveCachesToDB() As Task
        ' ---------------------------------------------------------------
        ' SaveCachesToDB — 把記憶體快取全部存入 SQLite
        ' 對應 Setting 頁 SaveCache 按鈕
        ' 流程: ① CleanupOrphanFolderPath (先清孤兒) → ② 批次寫入三張表 → ③ 統計顯示
        ' ---------------------------------------------------------------
        _dbg(" ├ 開始") ' by Gemini, 2026/04/10: 調整縮排層級為 Level 1
        If _db Is Nothing Then _dbg("", "DB 未初始化") : Return

        Dim sw As New Diagnostics.Stopwatch : sw.Start()
        Dim savedFolders, savedAttachMailList, savedAttachFilenames, savedBasic As Integer
        Try
            ProgressBar1.Text = "正在存入快取..." : Cursor = Cursors.WaitCursor

            ' ① 先清孤兒：收集目前記憶體快取中所有仍存在的 folder_path，清除 DB 中已不存在的行
            ' 用記憶體快取的 key 聯集代表「目前已知 live 的資料夾」 (比重新 BFS 掃 COM 快得多) 
            Dim livePaths As New HashSet(Of String)()
            For Each k In _cacheMailCount.Keys : livePaths.Add(k) : Next
            For Each k In _cacheFolderCount.Keys : livePaths.Add(k) : Next
            For Each k In _cacheAttachMailList.Keys : livePaths.Add(k) : Next
            If livePaths.Count > 0 Then Await CleanupOrphanPath(livePaths)

            ' ② SQLite I/O 在背景執行緒，不阻塞 UI
            Dim r = Await Task.Run(Function()
                                       Using txn As SqliteTransaction = _db.BeginTransaction()
                                           Try
                                               Dim f = SaveFolderStatsInner(txn)
                                               Dim b = SaveAttachMailListInner(txn)
                                               Dim a = SaveAttachFilenamesInner(txn)
                                               Dim y = SaveYearCountsInner(txn)
                                               Dim m = SaveMonthCountsInner(txn)   ' 2026/04/09 新增
                                               Dim basic = SaveBasicMailInfoInner(txn)
                                               txn.Commit()
                                               Return (f, b, a, y, m, basic)
                                           Catch ex As System.Exception
                                               txn.Rollback() : Throw
                                           End Try
                                       End Using
                                   End Function)

            savedFolders = r.f : savedAttachMailList = r.b : savedAttachFilenames = r.a
            Dim savedYears As Integer = r.y, savedMonths As Integer = r.m : savedBasic = r.basic
            sw.Stop()

            ' ③ 統計：各快取字典目前的 entry 數
            Dim statLine1 = $"① [記憶體] MailCount: {_cacheMailCount.Count} / MailCountAll: {_cacheMailCountAll.Count} / FolderCount: {_cacheFolderCount.Count} / FolderCountAll: {_cacheFolderCountAll.Count}"
            Dim statLine2 = $"② [記憶體] FolderSize: {_cacheFolderSize.Count} / FolderSizeAll: {_cacheFolderSizeAll.Count} / AttachPreScan: {_cacheAttachMailList.Count} / AttachFilename: {_cacheAttachFilename.Count}"
            Dim statLine3 = $"③ [寫入DB] folder_stats: {savedFolders} 筆 / basic_maillist: {savedBasic} 筆 / attach_maillist: {savedAttachMailList} 筆 / attach_filenames: {savedAttachFilenames} 筆 / year_counts: {savedYears} 筆 / month_counts: {savedMonths} 筆 / 耗時: {sw.Elapsed.TotalSeconds:0.000} 秒"
            Dim st = GetDBSummary()
            Dim statLine4 = $"④ [DB現況] folder_stats: {st.fc} 筆 / attach_maillist: {st.mb} 筆 / attach_filenames: {st.at} 筆 / year_counts: {st.yc} 筆 / month_counts: {st.mc} 筆 / 檔案: {st.kb} KB"

            ProgressBar1.Text = $"SaveCache 完成 — {statLine3}"
            ProgressBar2.Text = statLine4
            _dbg(" ├ ", statLine1)
            _dbg(" ├ ", statLine2)
            _dbg(" ├ ", statLine3)
            _dbg(" ├ ", statLine4)

        Catch ex As System.Exception
            ProgressBar1.Text = "SaveCache 失敗"
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
        If _db Is Nothing Then _dbg("", "DB 未初始化") : Return

        Dim sw As New Diagnostics.Stopwatch : sw.Start()
        Try
            ProgressBar1.Text = "正在載入快取..." : Cursor = Cursors.WaitCursor

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
                                       Dim m = LoadMonthCountsInner()               ' 2026/04/09 新增
                                       Dim basic = LoadBasicMailInfoInner()  ' 2026/04/22 by Gemini 3.1 Pro 新增
                                       Return (f, b, a, y, m, basic)
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
            Dim statLine_basic = $"⑦ [basic_maillist] 讀入 {r.basic} 筆 → BasicPreScan {_cacheBasicMailInfo.Count} 個資料夾" ' 2026/04/22 by Gemini 3.1 Pro
            Dim st = GetDBSummary()
            Dim statLine5 = $"⑧ [DB現況] folder_stats: {st.fc} 筆 / basic_maillist: {st.basic} 筆 / attach_maillist: {st.mb} 筆 / attach_filenames: {st.at} 筆 / year_counts: {st.yc} 筆 / month_counts: {st.mc} 筆 / 檔案: {st.kb} KB / 耗時: {sw.Elapsed.TotalSeconds:0.000} 秒" ' 2026/04/22 by Gemini 3.1 Pro: 加入 basic 統計

            ProgressBar1.Text = $"LoadCache 完成 — DB: {st.fc}/{st.basic}/{st.mb}/{st.at}/{st.yc}/{st.mc} 筆，{st.kb} KB，耗時 {sw.Elapsed.TotalSeconds:0.000} 秒" ' 2026/04/22 by Gemini 3.1 Pro
            ProgressBar2.Text = $"記憶體增量 — mailCount+{_cacheMailCount.Count - beforeMC} / attachFilename+{_cacheAttachFilename.Count - beforeAF} / basicMailInfo:{_cacheBasicMailInfo.Count} 資料夾" ' 2026/04/22 by Gemini 3.1 Pro
            _dbg(" ├ ", statLine1)
            _dbg(" ├ ", statLine2)
            _dbg(" ├ ", statLine3)
            _dbg(" ├ ", statLine4)
            _dbg(" ├ ", statLine_yc)
            _dbg(" ├ ", statLine_mc)
            _dbg(" ├ ", statLine_basic) ' 2026/04/22 by Gemini 3.1 Pro
            _dbg(" ├ ", statLine5)

        Catch ex As System.Exception
            ProgressBar1.Text = "LoadCache 失敗"
            _dbg("錯誤", ex.Message)
        Finally
            Cursor = Cursors.Default
            _dbg("結束")
        End Try

    End Function
    Private Async Function RenewCacheToDB(includeSize As Boolean) As Task
        ' ---------------------------------------------------------------
        ' RenewCacheToDB — 完整更新 DB 快取 (對應 Setting 頁 RenewCache 按鈕) 
        '
        ' 與 SaveCachesToDB 的差異：
        '   SaveCache  = 把目前記憶體快取照單全收寫入 DB (不更新過期的值) 
        '   RenewCache = 先用 COM 比對 snapshot → 只對有變動的資料夾重新計算 → 寫入 DB
        '
        ' 流程：
        '   Phase 1. BFS 掃出所有 live folders (COM，~1ms/資料夾) 
        '   Phase 2. 每個 folder 讀 GetLiveFolderSnapL3 vs DB snapshot → 找 dirty folders
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
        '   GetYearCountsForFolderL3 / GetFolderSizeL3 補入 cToken:=cToken
        ' ---------------------------------------------------------------

        Dim cToken As Threading.CancellationToken = OkayNowYouHaveToken()  ' ✅ 取得新 Token，同時取消上一次未完成的操作
        _dbg("開始", $"includeSize={includeSize}")
        If _db Is Nothing Then _dbg("", "DB 未初始化") : Return

        Dim sw As New Diagnostics.Stopwatch : sw.Start()
        Try
            ' ── Phase 1: BFS 掃出所有 live folders ──
            ' 2026/04/16: 改用 Dictionary(Of String, Outlook.Folder) liveDict
            '   key = FolderPath (一次 COM 呼叫)，value = Folder 物件
            '   後續 Phase2/3/4 直接用 kvp.Key 作 fPath，不再打 folder.FolderPath
            ProgressBar1.Text = "RenewCache Phase1: 掃描資料夾清單..." : Cursor = Cursors.WaitCursor
            Await Task.Yield

            Dim liveDict As New Dictionary(Of String, Outlook.Folder)()
            For Each store As Outlook.Store In _pstStoreList
                Dim root As Outlook.Folder = TryCast(store.GetRootFolder(), Outlook.Folder)
                If root Is Nothing Then Continue For

                ' 2026/04/24 by Gemini 3.0 flash: 使用 SafeGetPath 確保 root 取得安全
                Dim rootPath As String = SafeGetPath(root)
                If String.IsNullOrEmpty(rootPath) Then Continue For

                ' 2026/04/16 by Gemini: GetSubtreeToList 現在直接回傳 Tuple (Folder, FolderPath)
                ' 直接將計算好的路徑存入 liveDict，完成 0 COM Call 的清單建立
                For Each item In Await GetSubtreeToList(root, includeSubF:=True, cToken:=cToken)
                    If Not liveDict.ContainsKey(item.fPath) Then liveDict.Add(item.fPath, item.folder)
                Next
            Next
            Dim livePaths As New HashSet(Of String)(liveDict.Keys)  ' 供 Phase6 CleanupOrphan 使用
            _dbg("Phase1 完成", $"{liveDict.Count} 個 live folder")

            ' ── Phase 2: 比對 snapshot → 找出 dirty folders ──
            ' 2026/04/16: 迭代 liveDict，kvp.Key 直接當 fPath，省去 folder.FolderPath COM 呼叫
            '   節流改用 SmartThrottle(swThrottle2, cToken, ThrottleFreq.Low)，取代 Mod 100 + Task.Delay(1)
            ProgressBar1.Text = $"RenewCache Phase2: 比對 snapshot (共 {liveDict.Count} 個) ..."
            Dim dirtyDict As New Dictionary(Of String, Outlook.Folder)()
            ' by Claude Sonnet 4.6, 2026/04/25: 區分兩種「dirty」語意
            '   isNewFolder = True  → DB 從未記錄（清空後首次，或真正新資料夾）
            '                         Phase 3 只算 mc/fc/year_counts，跳過 attach_maillist 重掃
            '                         attach_maillist 交由使用者搜尋附件時 lazy 觸發
            '   isNewFolder = False → snapshot 不符（真正有信件增減）
            '                         Phase 3 完整重算包含 attach_maillist（三路比對）
            ' 這樣清空快取後執行 RenewCache，不會因為所有資料夾都「看起來像新的」而偷跑全量 GetTable 掃描，產生 2 萬筆非預期的 attach_maillist 內容。
            Dim dirtyNewFolderSet As New HashSet(Of String)()  ' 記錄 isNewFolder=True 的路徑
            Dim processed As Integer = 0
            Dim swThrottle2 As New Stopwatch : swThrottle2.Start()
            For Each kvp In liveDict
                cToken.ThrowIfCancellationRequested()  ' 2026/04/16: 取代 _cancelRequested + GoTo Cancelled
                Dim fPath As String = kvp.Key : Dim folder As Outlook.Folder = kvp.Value
                Dim liveSnap As Integer = GetLiveFolderSnapL3(folder, fPath:=fPath)   ' ~0.5ms，PropertyAccessor 單次呼叫 by Gemini 3.0 flash, 2026/04/16
                Dim row = DbGetFolderStats(fPath)

                ' dirty 條件：DB 無此路徑 (新資料夾) OR snapshot 不一致 (有信件增減)
                Dim isNewFolder As Boolean = (row Is Nothing)
                If isNewFolder OrElse row.snap <> liveSnap Then
                    dirtyDict.Add(fPath, folder)
                    If isNewFolder Then
                        dirtyNewFolderSet.Add(fPath)  ' 全新資料夾，Phase 3 跳過 attach_maillist

                        ' by Gemini 3.0 flash, 2026/04/24: 新資料夾確保 ID 被快取，Phase 6 寫入時需要 entry_id
                        _cacheFolderIDs.TryAdd(fPath, (folder.EntryID, folder.StoreID, IsMailFolder(folder, fPath), TextHasChineseChar(ExtractFolderName(fPath))))

                        ' 使父資料夾的樹狀快取失效，確保刷新 UI 後能顯示新成員
                        Dim parentPath As String = GetParentPath(fPath)
                        If Not String.IsNullOrEmpty(parentPath) Then
                            ' 清除父路徑的所有顯示模式快取 (|True 與 |False)
                            _cacheFolderTree.TryRemove(parentPath & "|True", Nothing)
                            _cacheFolderTree.TryRemove(parentPath & "|False", Nothing)
                        End If
                    End If
                End If

                processed += 1
                ' 2026/04/16 by Gemini 3.0 flash: 改用 ThrottleFreq.Low + SmartThrottle 與 onThrottled 委派
                Await SmartThrottle(swThrottle2, cToken:=cToken, ThrottleFreq.Low,
                                          Sub() ProgressBar1.Text = $"RenewCache Phase2: {processed}/{liveDict.Count}，dirty={dirtyDict.Count} (新={dirtyNewFolderSet.Count})...")
            Next
            _dbg("Phase2 完成", $"dirty={dirtyDict.Count}/{liveDict.Count} (其中全新資料夾={dirtyNewFolderSet.Count})")

            ' ── Phase 3: 對每個 dirty folder 重新計算 ──
            ' 2026/04/16: 迭代 dirtyDict，省去 folder.FolderPath COM 呼叫
            '   GetYearCountsForFolderL3 / GetFolderSizeL3 補入 cToken:=cToken
            '   節流改用 SmartThrottle(swThrottle3, cToken, ThrottleFreq.Low)，取代 Mod 10 + Task.Delay(1)
            ProgressBar1.Text = $"RenewCache Phase3: 更新 {dirtyDict.Count} 個 dirty 資料夾..." : Await Task.Delay(1, cToken)
            processed = 0
            Dim swThrottle3 As New Stopwatch : swThrottle3.Start()
            For Each kvp In dirtyDict
                cToken.ThrowIfCancellationRequested()  ' 2026/04/16: 取代 _cancelRequested + GoTo Cancelled
                Dim fPath As String = kvp.Key : Dim folder As Outlook.Folder = kvp.Value

                ' mc / fc — 快，~1ms，直接覆蓋記憶體快取
                _cacheMailCount(fPath) = GetMailCountL3(folder, fPath:=fPath)
                _cacheFolderCount(fPath) = GetFolderCountL3(folder, fPath:=fPath)

                ' year_counts — 清記憶體強制 L3 重算，結果回寫快取
                _cacheYearCounts.TryRemove(fPath, Nothing)
                _cacheYearCounts(fPath) = Await GetYearCountsForFolderL3(folder, fPath:=fPath, cToken:=cToken)  ' 2026/04/16: 補 cToken

                ' month_counts — 只清記憶體 (Phase5 再清 DB)，展開年份時 lazy 重算
                For Each mk In _cacheMonthCounts.Keys.Where(Function(k) k.StartsWith(fPath & "_")).ToList()
                    _cacheMonthCounts.TryRemove(mk, Nothing)
                Next

                ' attach_maillist — 三路比對，更新記憶體快取 (不碰 attach_filenames)
                ' by Claude Sonnet 4.6, 2026/04/25: 只對「真正 dirty」（snapshot 不符）的資料夾才重掃附件
                '   全新資料夾（DB 從未記錄）跳過，避免清空快取後 RenewCache 偷跑全量 GetTable 掃描
                '   全新資料夾的 attach_maillist 在使用者執行 Tab3 附件搜尋時透過 lazy load 建立
                If Not dirtyNewFolderSet.Contains(fPath) Then
                    Await RenewAttachMailList(folder, fPath:=fPath)
                End If

                ' folder_size — 選擇性 (GetTable 遍歷 PR_MESSAGE_SIZE，大資料夾需 10~30s)
                If includeSize Then _cacheFolderSize(fPath) = Await GetFolderSizeL3(folder, fPath:=fPath, cToken:=cToken)  ' 2026/04/16: 補 cToken

                ' 聚合快取清除 — 讓 parent 在下次點選時重新 BFS 加總
                ' by Claude Sonnet 4.6, 2026/04/25: 同時清除 |True 和 |False 兩個模式的鍵值
                '   因應未來 _showAllFolders 分支鍵值架構，確保兩個模式的過期聚合都被清掉
                _cacheMailCountAll.TryRemove(fPath & "|True", Nothing)
                _cacheMailCountAll.TryRemove(fPath & "|False", Nothing)
                _cacheMailCountAll.TryRemove(fPath, Nothing)    ' 兼容舊鍵值（無分支時寫入的）
                _cacheFolderCountAll.TryRemove(fPath & "|True", Nothing)
                _cacheFolderCountAll.TryRemove(fPath & "|False", Nothing)
                _cacheFolderCountAll.TryRemove(fPath, Nothing)  ' 同上
                _cacheFolderSizeAll.TryRemove(fPath, Nothing)

                processed += 1
                ' 2026/04/16 by Gemini 3.0 flash: 改用 ThrottleFreq.Low + SmartThrottle 與 onThrottled 委派
                Await SmartThrottle(swThrottle3, cToken:=cToken, ThrottleFreq.Low,
                                          Sub() ProgressBar1.Text = $"RenewCache Phase3: {processed}/{dirtyDict.Count} 個處理中...")
            Next
            _dbg("Phase3 完成", $"{processed} 個 dirty folder 重新計算完畢")

            ' ── Phase 4: 清除 dirty folders 的 ancestor 聚合快取 ──
            ' 任何 dirty leaf 都讓所有 ancestor 的 mca/fca/fsa 失效
            ' 2026/04/16: 改迭代 liveDict.Keys，直接用 key 作 fPath，省去 fs.FolderPath COM 呼叫
            ' by Gemini 3.0 flash, 2026/04/24: 優化為「精確打擊」模式，改用 GetAncestors 直接清除，效能從 O(N*D) 降至 O(D*L)
            If dirtyDict.Count > 0 Then
                For Each dp In dirtyDict.Keys
                    For Each ancestor In GetAncestors(dp)
                        ' by Claude Sonnet 4.6, 2026/04/25: 同時清除 |True / |False 兩個模式鍵值及舊式鍵值
                        _cacheMailCountAll.TryRemove(ancestor & "|True", Nothing)
                        _cacheMailCountAll.TryRemove(ancestor & "|False", Nothing)
                        _cacheMailCountAll.TryRemove(ancestor, Nothing)
                        _cacheFolderCountAll.TryRemove(ancestor & "|True", Nothing)
                        _cacheFolderCountAll.TryRemove(ancestor & "|False", Nothing)
                        _cacheFolderCountAll.TryRemove(ancestor, Nothing)
                        _cacheFolderSizeAll.TryRemove(ancestor, Nothing)
                    Next
                Next
                _dbg("Phase4 完成", $"已針對 {dirtyDict.Count} 個異動路徑精確清除祖先快取 (含 |True/|False 模式鍵值)")
            End If

            ' ── Phase 5: 批次 DELETE dirty folders 的 month_counts DB rows ──
            ' 注意: CleanupOrphan 只刪「不再存在的路徑」，不刪「仍存在但 dirty」的路徑
            '       所以 dirty folder 的舊 month rows 必須在這裡主動清除
            Dim dirtyPaths As New HashSet(Of String)(dirtyDict.Keys)  ' 供 Phase 5 使用
            If dirtyPaths.Count > 0 AndAlso _db IsNot Nothing Then
                Await Task.Run(Sub()
                                   Using txn = _db.BeginTransaction()
                                       Try
                                           Using cmd As New SqliteCommand("DELETE FROM month_counts WHERE folder_path=@p", _db, txn)
                                               cmd.Parameters.Add("@p", SqliteType.Text)
                                               For Each dp In dirtyPaths
                                                   cmd.Parameters("@p").Value = dp : cmd.ExecuteNonQuery()
                                               Next
                                           End Using
                                           txn.Commit()
                                       Catch : txn.Rollback() : Throw
                                       End Try
                                   End Using
                               End Sub)
                _dbg("Phase5 完成", $"已清 {dirtyPaths.Count} 個 dirty folder 的 month_counts DB rows")
            End If

            ' ── Phase 6: 孤兒清除 + 批次寫入 ──
            ProgressBar1.Text = "RenewCache Phase6: 清孤兒 + 寫入 DB..." : Await Task.Delay(1, cToken)
            Await CleanupOrphanPath(livePaths)
            Await SaveCachesToDB()    ' 內部會顯示 SaveCache 的進度訊息

            sw.Stop()
            Dim st = GetDBSummary()
            ProgressBar1.Text = $"RenewCache 完成 ✔ dirty:{dirtyDict.Count}/{liveDict.Count} 個 / 耗時:{sw.Elapsed.TotalSeconds:0.0}s — DB:{st.fc}/{st.mb}/{st.at}/{st.yc}/{st.mc} 筆"
            _dbg("完成", $"dirty={dirtyDict.Count}, total={liveDict.Count}, elapsed={sw.Elapsed.TotalSeconds:0.0}s")

        Catch ex As OperationCanceledException
            ' 2026/04/16: cToken 取消時 (ESC)，取代原本的 _cancelRequested + GoTo Cancelled 模式
            ProgressBar1.Text = "RenewCache 已中斷 (ESC)"
            _dbg("中斷", "使用者按 ESC")
        Catch ex As System.Exception
            ProgressBar1.Text = $"RenewCache 失敗: {ex.Message}"
            _dbg("錯誤", ex.Message)
        Finally
            Cursor = Cursors.Default
            _dbg("結束")
        End Try

    End Function
    Private Async Function RenewAttachMailList(folder As Outlook.Folder, fPath As String) As Task
        ' ---------------------------------------------------------------
        ' RenewAttachMailList — 三路比對更新單一資料夾的 attach_maillist 快取
        '
        ' 三路比對邏輯：
        '   新郵件   (live 有、DB 沒有) → 進入新的 mailList，SaveCache 時 INSERT
        '   已刪郵件 (DB 有、live 沒有) → 從 _cacheAttachFilename 清除 (DB row 留 CleanupOrphan 處理) 
        '   未變郵件 (live ∩ DB)        → 原有 filenames 快取保留，不重掃附件
        '
        ' attach_filenames 永不重掃 (設計邊界，最耗時步驟留給 Tab3 搜尋時 lazy 觸發) 
        ' 2026/04/09 by Claude
        ' ---------------------------------------------------------------
        Dim fName As String = ExtractFolderName(fPath) ' by Gemini 3.1 Pro 2026/04/16: 透過 fPath 取名，省下 COM 呼叫
        If _iLikeNoisy Then _dbg("開始", fName)

        ' 取 live 含附件郵件清單 (L3 GetTable 重掃，~2-5ms/資料夾) 
        Dim newMailList = Await GetAttachMailListL3(folder, Nothing)
        Dim liveEntrySet As New HashSet(Of String)(newMailList.Select(Function(m) m.EntryID))

        ' 取 DB 中已有的 EntryIDs (供三路比對找已刪郵件) 
        Dim dbResult = DbGetAttachMailList(fPath)
        If dbResult IsNot Nothing Then
            ' 已刪郵件：從 _cacheAttachFilename 清除 (下次 Tab3 掃描時不會再用到) 
            For Each mail In dbResult.Mails
                If Not liveEntrySet.Contains(mail.EntryID) Then _cacheAttachFilename.TryRemove(mail.EntryID, Nothing)
            Next
        End If

        ' 更新記憶體快取 (新清單直接覆蓋，ItemCountSnap 用 live snap) 
        _cacheAttachMailList(fPath) = New FolderCacheTab3 With {.AttachMailList = newMailList,
                                                               .ItemCountSnap = GetLiveFolderSnapL3(folder, fPath:=fPath)}

        If _iLikeNoisy Then _dbg("結束", $"{fName}: live={newMailList.Count}, db_was={If(dbResult?.Mails.Count, 0)}")

    End Function
    Private Async Function CleanupOrphanPath(liveFolderPaths As HashSet(Of String)) As Task
        ' ---------------------------------------------------------------
        ' CleanupOrphanPath — 刪除已不存在於 Outlook 的資料夾孤兒行 (改為非同步 by Gemini 3.1 Pro, 2026/05/05)
        ' liveFolderPaths = 目前仍有效的資料夾路徑集合
        '   呼叫來源 A: SaveCachesToDB 開頭 (用記憶體快取 key 聯集) 
        '   呼叫來源 B: RenewCache_Click (用 GetSubtreeToList BFS 掃 COM 取得完整清單) 
        ' ---------------------------------------------------------------
        _dbg("    ├ 開始", $"live 資料夾數: {liveFolderPaths.Count}") ' by Gemini, 2026/04/10: 調整縮排層級為 Level 2
        If _db Is Nothing Then Return

        Await Task.Run(Sub()
                           Try
                               ' 讀出 DB 中所有 folder_path
                               Dim dbPaths As New List(Of String)(2048)
                               Using cmd As New SqliteCommand("SELECT folder_path FROM folder_stats", _db)
                                   Using reader = cmd.ExecuteReader()
                                       While reader.Read() : dbPaths.Add(reader.GetString(0)) : End While
                                   End Using
                               End Using
                               _dbg("", $"DB 中有 {dbPaths.Count} 個資料夾路徑")

                               Dim stalePaths = dbPaths.Where(Function(p) Not liveFolderPaths.Contains(p)).ToList()
                               If stalePaths.Count = 0 Then _dbg("", "未發現孤兒快取，略過") : Return

                               ' 每個孤兒路徑輸出到 _dbg 供目視確認
                               For Each stale In stalePaths
                                   If _iLikeNoisy Then _dbg(" 孤兒", stale)
                               Next

                               Dim dF, dB, dA, dM, dBasic As Integer
                               Using txn As SqliteTransaction = _db.BeginTransaction()
                                   For Each stale In stalePaths
                                       Using c1 As New SqliteCommand("DELETE FROM folder_stats WHERE folder_path=@p", _db, txn)
                                           c1.Parameters.AddWithValue("@p", stale) : dF += c1.ExecuteNonQuery()
                                       End Using
                                       Using c2 As New SqliteCommand("DELETE FROM attach_maillist WHERE folder_path=@p", _db, txn)
                                           c2.Parameters.AddWithValue("@p", stale) : dB += c2.ExecuteNonQuery()
                                       End Using
                                       Using c3 As New SqliteCommand("DELETE FROM attach_filenames WHERE folder_path=@p", _db, txn)
                                           c3.Parameters.AddWithValue("@p", stale) : dA += c3.ExecuteNonQuery()
                                       End Using
                                       ' 2026/04/09: month_counts 以 folder_path 欄位清除 (cache_key 含 year 後綴，不可直接比對) 
                                       Using c4 As New SqliteCommand("DELETE FROM month_counts WHERE folder_path=@p", _db, txn)
                                           c4.Parameters.AddWithValue("@p", stale) : dM += c4.ExecuteNonQuery()
                                       End Using
                                       Using c5 As New SqliteCommand("DELETE FROM basic_maillist WHERE folder_path=@p", _db, txn)
                                           c5.Parameters.AddWithValue("@p", stale) : dBasic += c5.ExecuteNonQuery()
                                       End Using
                                   Next
                                   txn.Commit()
                               End Using
                               _dbg("結束", $"孤兒路徑:{stalePaths.Count} 個 / folder_stats:{dF} 行 / basic_maillist:{dBasic} 行 / attach_maillist:{dB} 行 / attach_filenames:{dA} 行 / month_counts:{dM} 行")

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
        ' content_count_snapshot = _cacheMailCount[path]，即 PR_CONTENT_COUNT 讀取結果
        ' ---------------------------------------------------------------
        _dbg("    ├ 開始") ' by Gemini, 2026/04/10: 調整縮排層級為 Level 2
        Dim sql = "INSERT OR REPLACE INTO folder_stats" &
                  " (folder_path,mail_count,mail_count_all,folder_count,folder_count_all," &
                  "  folder_size,folder_size_all,content_count_snapshot,entry_id,store_id,is_mail,has_chinese,updated_at)" &
                  " VALUES (@path,@mc,@mca,@fc,@fca,@fs,@fsa,@snap,@eid,@sid,@ism,@hasch,@ts)"

        ' 蒐集六個 dict 中所有出現過的 folder_path 聯集
        Dim allPaths As New HashSet(Of String)()
        For Each k In _cacheMailCount.Keys : allPaths.Add(k) : Next
        For Each k In _cacheMailCountAll.Keys : allPaths.Add(k) : Next
        For Each k In _cacheFolderCount.Keys : allPaths.Add(k) : Next
        For Each k In _cacheFolderCountAll.Keys : allPaths.Add(k) : Next
        For Each k In _cacheFolderSize.Keys : allPaths.Add(k) : Next
        For Each k In _cacheFolderSizeAll.Keys : allPaths.Add(k) : Next
        ' by Gemini 3.0 flash, 2026/04/18: 額外聯集身分標識字典的 Key，確保僅掃描過但未統計的資料夾也能存入 SSD
        For Each k In _cacheFolderIDs.Keys : allPaths.Add(k) : Next

        Dim count As Integer = 0
        Dim ts = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
        Using cmd As New SqliteCommand(sql, _db, txn)
            cmd.Parameters.Add("@path", SqliteType.Text)
            cmd.Parameters.Add("@mc", SqliteType.Integer)
            cmd.Parameters.Add("@mca", SqliteType.Integer)
            cmd.Parameters.Add("@fc", SqliteType.Integer)
            cmd.Parameters.Add("@fca", SqliteType.Integer)
            cmd.Parameters.Add("@fs", SqliteType.Integer)
            cmd.Parameters.Add("@fsa", SqliteType.Integer)
            cmd.Parameters.Add("@snap", SqliteType.Integer)
            cmd.Parameters.Add("@eid", SqliteType.Text)
            cmd.Parameters.Add("@sid", SqliteType.Text)
            cmd.Parameters.Add("@ism", SqliteType.Integer)
            cmd.Parameters.Add("@hasch", SqliteType.Integer)
            cmd.Parameters.Add("@ts", SqliteType.Text)

            For Each path In allPaths
                ' 2026/04/07 修正 v2: 初始值設 -1 仍然不夠，因為 -1 是整數值會被寫入 DB，
                ' LoadFolderStatsInner 讀回 -1 後直接塞入記憶體快取，
                ' GetFolderCount 命中記憶體回傳 -1 → LoadSubFolderToTreeView 判斷 -1 > 0 為 False → 不顯示 "+"。
                ' 正確做法：沒有測量過的欄位一律寫 DBNull.Value (SQL NULL)，
                ' 這樣 LoadFolderStatsInner 的 IsDBNull 保護才能正確跳過，不污染記憶體快取。
                Dim mc, mca, fc, fca As Integer : Dim fs, fsa As Long
                Dim hasMc = _cacheMailCount.TryGetValue(path, mc)
                Dim hasMca = _cacheMailCountAll.TryGetValue(path, mca)
                Dim hasFc = _cacheFolderCount.TryGetValue(path, fc)
                Dim hasFca = _cacheFolderCountAll.TryGetValue(path, fca)
                Dim hasFs = _cacheFolderSize.TryGetValue(path, fs)
                Dim hasFsa = _cacheFolderSizeAll.TryGetValue(path, fsa)
                cmd.Parameters("@path").Value = path
                cmd.Parameters("@mc").Value = If(hasMc, CObj(mc), DBNull.Value)
                cmd.Parameters("@mca").Value = If(hasMca, CObj(mca), DBNull.Value)
                cmd.Parameters("@fc").Value = If(hasFc, CObj(fc), DBNull.Value)
                cmd.Parameters("@fca").Value = If(hasFca, CObj(fca), DBNull.Value)
                cmd.Parameters("@fs").Value = If(hasFs, CObj(fs), DBNull.Value)
                cmd.Parameters("@fsa").Value = If(hasFsa, CObj(fsa), DBNull.Value)
                ' content_count_snapshot: 只有 mc 有測量過才有意義
                cmd.Parameters("@snap").Value = If(hasMc, CObj(mc), DBNull.Value)

                ' by Gemini, 2026/04/10: 寫入身分標識與標籤 (從 _cacheFolderIDs 提取)
                Dim idInfo As (eid As String, sid As String, isMail As Boolean, hasCh As Boolean) = Nothing
                If _cacheFolderIDs.TryGetValue(path, idInfo) Then
                    cmd.Parameters("@eid").Value = idInfo.eid
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
        Dim sql = "INSERT OR REPLACE INTO year_counts (folder_path,year,count,updated_at) VALUES (@fp,@yr,@cnt,@ts)"
        Dim count As Integer = 0
        Dim ts = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

        Using cmd As New SqliteCommand(sql, _db, txn)
            cmd.Parameters.Add("@fp", SqliteType.Text)
            cmd.Parameters.Add("@yr", SqliteType.Integer)
            cmd.Parameters.Add("@cnt", SqliteType.Integer)
            cmd.Parameters.Add("@ts", SqliteType.Text)

            For Each kvp In _cacheYearCounts
                Dim fp = kvp.Key
                For Each yr In kvp.Value
                    cmd.Parameters("@fp").Value = fp
                    cmd.Parameters("@yr").Value = yr.Key
                    cmd.Parameters("@cnt").Value = yr.Value
                    cmd.Parameters("@ts").Value = ts
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
        Dim sql = "INSERT OR REPLACE INTO month_counts (folder_path,year,month,count,updated_at) VALUES (@fp,@yr,@mo,@cnt,@ts)"
        Dim count As Integer = 0
        Dim ts = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

        Using cmd As New SqliteCommand(sql, _db, txn)
            cmd.Parameters.Add("@fp", SqliteType.Text)
            cmd.Parameters.Add("@yr", SqliteType.Integer)
            cmd.Parameters.Add("@mo", SqliteType.Integer)
            cmd.Parameters.Add("@cnt", SqliteType.Integer)
            cmd.Parameters.Add("@ts", SqliteType.Text)

            For Each kvp In _cacheMonthCounts
                ' cache_key 格式: "FolderPath_year"，最後一個 "_" 分隔出 year
                Dim cacheKey = kvp.Key
                Dim lastUnderscore = cacheKey.LastIndexOf("_"c)
                If lastUnderscore < 0 Then Continue For
                Dim folderPath = cacheKey.Substring(0, lastUnderscore)
                Dim yearVal As Integer
                If Not Integer.TryParse(cacheKey.Substring(lastUnderscore + 1), yearVal) Then Continue For

                For Each mo In kvp.Value
                    cmd.Parameters("@fp").Value = folderPath
                    cmd.Parameters("@yr").Value = yearVal
                    cmd.Parameters("@mo").Value = mo.Key
                    cmd.Parameters("@cnt").Value = mo.Value
                    cmd.Parameters("@ts").Value = ts
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
        ' ---------------------------------------------------------------
        _dbg("開始")
        Dim sql = "INSERT OR REPLACE INTO attach_maillist" &
                  " (entry_id,folder_path,subject,msg_size,received_time,sender_name,attach_count,item_count_snap,updated_at)" &
                  " VALUES (@eid,@fp,@subj,@sz,@rt,@sn,@ac,@snap,@ts)"

        Dim count As Integer = 0
        Dim ts = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
        Using cmd As New SqliteCommand(sql, _db, txn)
            cmd.Parameters.Add("@eid", SqliteType.Text)
            cmd.Parameters.Add("@fp", SqliteType.Text)
            cmd.Parameters.Add("@subj", SqliteType.Text)
            cmd.Parameters.Add("@sz", SqliteType.Integer)
            cmd.Parameters.Add("@rt", SqliteType.Text)
            cmd.Parameters.Add("@sn", SqliteType.Text)
            cmd.Parameters.Add("@ac", SqliteType.Integer)
            cmd.Parameters.Add("@snap", SqliteType.Integer)
            cmd.Parameters.Add("@ts", SqliteType.Text)

            ' _cacheAttachMailList: Dictionary(Of String, FolderCacheTab3)
            ' key = folder_path, value.AttachMailList = List(Of MailItemInfo)
            For Each kvp In _cacheAttachMailList
                Dim fp = kvp.Key : Dim snap = kvp.Value.ItemCountSnap
                Dim mails = kvp.Value.AttachMailList

                If mails.Count = 0 Then
                    ' by Gemini 3 Flash, 2026/05/06: 實作「空結果持久化」，記住此資料夾已掃描且為 0 筆
                    cmd.Parameters("@eid").Value = "EMPTY_ATTACH_" & fp
                    cmd.Parameters("@fp").Value = fp
                    cmd.Parameters("@subj").Value = ""
                    cmd.Parameters("@sz").Value = 0
                    cmd.Parameters("@rt").Value = ""
                    cmd.Parameters("@sn").Value = ""
                    cmd.Parameters("@ac").Value = 0
                    cmd.Parameters("@snap").Value = snap
                    cmd.Parameters("@ts").Value = ts
                    cmd.ExecuteNonQuery() : count += 1
                Else
                    For Each mail In mails
                        cmd.Parameters("@eid").Value = mail.EntryID
                        cmd.Parameters("@fp").Value = fp
                        cmd.Parameters("@subj").Value = If(mail.Subject, "")
                        cmd.Parameters("@sz").Value = mail.Size
                        cmd.Parameters("@rt").Value = mail.ReceivedTime.ToString("yyyy-MM-dd HH:mm:ss")
                        cmd.Parameters("@sn").Value = If(mail.SenderName, "")
                        cmd.Parameters("@ac").Value = mail.AttachCount
                        cmd.Parameters("@snap").Value = snap
                        cmd.Parameters("@ts").Value = ts
                        cmd.ExecuteNonQuery() : count += 1
                    Next
                End If
            Next
        End Using
        Return count
        _dbg("結束")

    End Function
    Private Function SaveAttachFilenamesInner(txn As SqliteTransaction) As Integer
        ' ---------------------------------------------------------------
        ' SaveAttachFilenamesInner — Transaction 內批次寫入 attach_filenames (Tab3 Phase2) 
        ' folder_path 透過反查 _cacheAttachMailList 取得 (_cacheAttachFilename key 是 EntryID) 
        ' 2026/04/09 修正: 移除 msg_size 欄位 (Phase2 永遠是 NULL，保留在 INSERT 造成
        '   SqliteType.Integer + DBNull.Value 不相容，丟 "Value must be set" InvalidOperationException)
        '   SQLite 未列出的欄位自動填 NULL，不需要明確傳入。
        ' ---------------------------------------------------------------
        _dbg("開始")
        Dim sql = "INSERT OR REPLACE INTO attach_filenames" &
                  " (entry_id,folder_path,filenames,updated_at)" &
                  " VALUES (@eid,@fp,@fn,@ts)"
        Dim count As Integer = 0
        Dim ts = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

        ' 反查 EntryID → folder_path (從 Phase1 快取中建立對應表) 
        Dim entryToFolder As New Dictionary(Of String, String)()
        For Each kvp In _cacheAttachMailList
            For Each mail In kvp.Value.AttachMailList
                If Not entryToFolder.ContainsKey(mail.EntryID) Then entryToFolder(mail.EntryID) = kvp.Key
            Next
        Next

        Using cmd As New SqliteCommand(sql, _db, txn)
            cmd.Parameters.Add("@eid", SqliteType.Text)
            cmd.Parameters.Add("@fp", SqliteType.Text)
            cmd.Parameters.Add("@fn", SqliteType.Text)
            cmd.Parameters.Add("@ts", SqliteType.Text)

            For Each kvp In _cacheAttachFilename
                Dim fp = "" : entryToFolder.TryGetValue(kvp.Key, fp)
                cmd.Parameters("@eid").Value = kvp.Key
                cmd.Parameters("@fp").Value = fp
                cmd.Parameters("@fn").Value = JsonSerializer.Serialize(kvp.Value)
                cmd.Parameters("@ts").Value = ts
                cmd.ExecuteNonQuery() : count += 1
            Next
        End Using
        Return count
        _dbg("結束")

    End Function
    Private Function SaveBasicMailInfoInner(txn As SqliteTransaction) As Integer
        ' ---------------------------------------------------------------
        ' SaveBasicMailInfoInner — Transaction 內批次寫入 basic_maillist (Tab4/Tab5) 
        ' ---------------------------------------------------------------
        _dbg("開始")
        ' 2026/05/06 by Claude: 新增 message_id, sender_email 欄位; cache key 已改為純路徑，不再需 .Split("|"c)(0)
        Dim sql = "INSERT OR REPLACE INTO basic_maillist" &
                  " (entry_id,folder_path,subject,msg_size,received_time,sender_name,topic,message_id,sender_email,item_count_snap,updated_at)" &
                  " VALUES (@eid,@fp,@subj,@sz,@rt,@sn,@tpc,@mid,@semail,@snap,@ts)"

        Dim count As Integer = 0
        Dim ts = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
        Using cmd As New SqliteCommand(sql, _db, txn)
            cmd.Parameters.Add("@eid", SqliteType.Text)
            cmd.Parameters.Add("@fp", SqliteType.Text)
            cmd.Parameters.Add("@subj", SqliteType.Text)
            cmd.Parameters.Add("@sz", SqliteType.Integer)
            cmd.Parameters.Add("@rt", SqliteType.Text)
            cmd.Parameters.Add("@sn", SqliteType.Text)
            cmd.Parameters.Add("@tpc", SqliteType.Text)
            cmd.Parameters.Add("@mid", SqliteType.Text)
            cmd.Parameters.Add("@semail", SqliteType.Text)
            cmd.Parameters.Add("@snap", SqliteType.Integer)
            cmd.Parameters.Add("@ts", SqliteType.Text)

            For Each kvp In _cacheBasicMailInfo
                ' by Claude Sonnet 4.6, 2026/05/06: Root Cause A 修正
                ' kvp.Key 格式為 "fPath|True" 或 "fPath|False"，直接寫入 DB 會污染 folder_path 欄位
                ' 導致 DbGetBasicMailInfo(fPath) 查詢時，WHERE folder_path=@p 永遠找不到資料
                ' 解法：剝離後綴，只取 "|" 前的純路徑
                ' Dim fp = kvp.Key.Split("|"c)(0) : Dim snap = kvp.Value.Snap

                Dim fp As String = kvp.Key   ' 2026/05/06 by Claude: key 已是純路徑，不再需 .Split
                Dim snap = kvp.Value.Snap
                Dim mails = kvp.Value.Mails

                If mails.Count = 0 Then
                    ' by Gemini 3 Flash, 2026/05/06: 實作「空結果持久化」，記住此資料夾已掃描且為 0 筆
                    ' 使用具備快取辨識性的 ID
                    cmd.Parameters("@eid").Value = "EMPTY_BASIC_" & fp
                    cmd.Parameters("@fp").Value = fp
                    cmd.Parameters("@subj").Value = ""
                    cmd.Parameters("@sz").Value = 0
                    cmd.Parameters("@rt").Value = ""
                    cmd.Parameters("@sn").Value = ""
                    cmd.Parameters("@tpc").Value = ""
                    cmd.Parameters("@mid").Value = ""
                    cmd.Parameters("@semail").Value = ""
                    cmd.Parameters("@snap").Value = snap
                    cmd.Parameters("@ts").Value = ts
                    cmd.ExecuteNonQuery() : count += 1
                Else
                    For Each item In mails
                        cmd.Parameters("@eid").Value = item.Mail.EntryID
                        cmd.Parameters("@fp").Value = fp
                        cmd.Parameters("@subj").Value = If(item.Mail.Subject, "")
                        cmd.Parameters("@sz").Value = item.Mail.Size
                        cmd.Parameters("@rt").Value = item.Mail.ReceivedTime.ToString("yyyy-MM-dd HH:mm:ss")
                        cmd.Parameters("@sn").Value = If(item.Mail.SenderName, "")
                        cmd.Parameters("@tpc").Value = If(item.Topic, "")
                        cmd.Parameters("@mid").Value = If(item.Mail.MessageID, "")
                        cmd.Parameters("@semail").Value = If(item.Mail.SenderEmail, "")
                        cmd.Parameters("@snap").Value = snap
                        cmd.Parameters("@ts").Value = ts
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
        ' ---------------------------------------------------------------
        _dbg("開始")
        Dim count As Integer = 0
        Using cmd As New SqliteCommand(
            "SELECT folder_path,mail_count,mail_count_all,folder_count,folder_count_all,folder_size,folder_size_all," &
            "entry_id,store_id,is_mail,has_chinese FROM folder_stats", _db)
            Using reader = cmd.ExecuteReader()
                While reader.Read()
                    Dim path = reader.GetString(0)
                    ' 只有 NOT NULL 的欄位才塞入記憶體快取；NULL 代表「從未測量過」，跳過
                    If Not reader.IsDBNull(1) Then _cacheMailCount.TryAdd(path, reader.GetInt32(1))
                    If Not reader.IsDBNull(2) Then _cacheMailCountAll.TryAdd(path, reader.GetInt32(2))
                    If Not reader.IsDBNull(3) Then _cacheFolderCount.TryAdd(path, reader.GetInt32(3))
                    If Not reader.IsDBNull(4) Then _cacheFolderCountAll.TryAdd(path, reader.GetInt32(4))
                    If Not reader.IsDBNull(5) Then _cacheFolderSize.TryAdd(path, reader.GetInt64(5))
                    If Not reader.IsDBNull(6) Then _cacheFolderSizeAll.TryAdd(path, reader.GetInt64(6))

                    ' by Gemini 3.0 flash, 2026/04/18: 批量讀取時回填身分標識與標籤字典，確保 LoadCache 後狀態完整
                    Dim eid As String = If(Not reader.IsDBNull(7), reader.GetString(7), "")
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

        Using cmd As New SqliteCommand("SELECT folder_path,year,count FROM year_counts", _db)
            Using reader = cmd.ExecuteReader()
                While reader.Read()
                    Dim fp = reader.GetString(0)
                    Dim yr = reader.GetInt32(1)
                    Dim cnt = reader.GetInt32(2)
                    If Not tempDict.ContainsKey(fp) Then
                        tempDict(fp) = New ConcurrentDictionary(Of Integer, Integer)()
                    End If
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

        Using cmd As New SqliteCommand("SELECT folder_path,year,month,count FROM month_counts", _db)
            Using reader = cmd.ExecuteReader()
                While reader.Read()
                    Dim fp = reader.GetString(0)
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
        ' ---------------------------------------------------------------
        _dbg("開始")
        Dim count As Integer = 0
        ' 暫存用：先按 folder_path 分組收集，最後一次性寫入 _cacheAttachMailList
        Dim tempDict As New Dictionary(Of String, (snap As Integer, mails As List(Of MailItemInfo)))()

        Using cmd As New SqliteCommand("SELECT entry_id,folder_path,subject,msg_size,received_time,sender_name,attach_count,item_count_snap FROM attach_maillist", _db)
            Using reader = cmd.ExecuteReader()
                While reader.Read()
                    Dim fp = reader.GetString(1)
                    If Not tempDict.ContainsKey(fp) Then
                        tempDict(fp) = (reader.GetInt32(7), New List(Of MailItemInfo)())
                    End If

                    Dim mail As New MailItemInfo()
                    mail.EntryID = reader.GetString(0)
                    mail.Subject = If(reader.IsDBNull(2), "", reader.GetString(2))
                    mail.Size = reader.GetInt64(2)

                    Dim dtStr = If(reader.IsDBNull(4), "", reader.GetString(4))
                    DateTime.TryParse(dtStr, mail.ReceivedTime)
                    mail.SenderName = If(reader.IsDBNull(5), "", reader.GetString(5))
                    mail.AttachCount = reader.GetInt32(6)
                    mail.FolderPath = fp ' 確保讀取快取時填入路徑
                    tempDict(fp).mails.Add(mail) : count += 1
                End While
            End Using
        End Using

        For Each kvp In tempDict
            ' 優化後：
            _cacheAttachMailList.TryAdd(kvp.Key,
                                       New FolderCacheTab3 With {.AttachMailList = kvp.Value.mails,
                                                                 .ItemCountSnap = kvp.Value.snap})
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
        Using cmd As New SqliteCommand("SELECT entry_id,filenames FROM attach_filenames", _db)
            Using reader = cmd.ExecuteReader()
                While reader.Read()
                    Dim eid = reader.GetString(0)
                    Dim fnJson = If(reader.IsDBNull(1), "[]", reader.GetString(1))
                    Try
                        Dim list = JsonSerializer.Deserialize(Of List(Of String))(fnJson)
                        _cacheAttachFilename.TryAdd(eid, list) : count += 1
                    Catch ex As System.Exception
                        _dbg("錯誤: 解析失敗", $"{eid}: {ex.Message}")
                    End Try
                End While
            End Using
        End Using
        Return count
        _dbg("結束")

    End Function
    Private Function LoadBasicMailInfoInner() As Integer
        ' ---------------------------------------------------------------
        ' LoadBasicMailInfoInner — 重建 _cacheBasicMailInfo (Tab4/5 專用)
        ' 2026/04/22 by Gemini 3.1 Pro: 補齊載入邏輯，解決重啟後重複掃描問題
        ' ---------------------------------------------------------------
        _dbg("開始")
        ' 由於此表資料量可能較大，我們按 folder_path 分組收集到記憶體中
        Dim count As Integer = 0
        Dim tempDict As New Dictionary(Of String, (Mails As List(Of (Mail As MailItemInfo, Topic As String)), Snap As Integer))()

        ' 2026/05/06 by Claude: 新增 message_id(7), sender_email(8); item_count_snap 移至索引 9
        Using cmd As New SqliteCommand(
            "SELECT entry_id,folder_path,subject,msg_size,received_time,sender_name,topic," &
            "message_id,sender_email,item_count_snap FROM basic_maillist", _db)
            Using reader = cmd.ExecuteReader()
                While reader.Read()
                    Dim eid = reader.GetString(0)
                    Dim fp = reader.GetString(1)
                    Dim subj = If(reader.IsDBNull(2), "", reader.GetString(2))
                    Dim sz = reader.GetInt64(3)
                    Dim rtStr = If(reader.IsDBNull(4), "", reader.GetString(4))
                    Dim sn = If(reader.IsDBNull(5), "", reader.GetString(5))
                    Dim tpc = If(reader.IsDBNull(6), "", reader.GetString(6))
                    Dim mid = If(reader.IsDBNull(7), "", reader.GetString(7))
                    Dim semail = If(reader.IsDBNull(8), "", reader.GetString(8))
                    Dim snap = If(reader.IsDBNull(9), -1, reader.GetInt32(9))

                    ' by Claude Sonnet 4.6, 2026/05/06: Root Cause A 修正 — fp 現在是純路徑
                    ' 記憶體 key 格式須與 GetBasicMailInfo 使用的 cacheKey 對齊
                    ' GetBasicMailInfo 使用 fPath & "|" & needTopic，Tab4 固定傳 needTopic=True
                    ' DB 存的是完整資料（含 topic），統一重建為 |"|True" 版本
                    ' Dim cacheKey = fp & "|True"
                    Dim cacheKey = fp   ' 2026/05/06 by Claude: cache key 改為純路徑（與 GetBasicMailInfo L2.5 一致）
                    If Not tempDict.ContainsKey(cacheKey) Then tempDict(cacheKey) = (New List(Of (Mail As MailItemInfo, Topic As String))(), snap)

                    ' by Claude Sonnet 4.6, 2026/05/06: Root Cause B 修正 — 跳過 sentinel row
                    ' EMPTY_BASIC_ 是「已掃描但無信」的標記行，不可建立假 MailItemInfo
                    ' 但其 snap 仍需讀取（已在上方 snap = reader.GetInt32(7) 讀取），tempDict entry 已建立
                    If eid.StartsWith("EMPTY_BASIC_") Then count += 1 : Continue While

                    Dim mail As New MailItemInfo With {
                        .EntryID = eid, .Subject = subj, .Size = sz, .SenderName = sn,
                        .FolderPath = fp, .MessageID = mid, .SenderEmail = semail}
                    DateTime.TryParse(rtStr, mail.ReceivedTime)
                    tempDict(cacheKey).Mails.Add((mail, tpc))
                    count += 1
                End While
            End Using
        End Using

        For Each kvp In tempDict
            _cacheBasicMailInfo.TryAdd(kvp.Key, kvp.Value)
        Next
        Return count
        _dbg("結束")
    End Function
#End Region

#Region "■ Layer 2.5 即時查詢 (Lazy SELECT Helpers)"
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
        If _db Is Nothing Then Return Nothing

        Try
            Using cmd As New SqliteCommand(
                "SELECT mail_count,mail_count_all,folder_count,folder_count_all," &
                "       folder_size,folder_size_all,content_count_snapshot," &
                "       entry_id,store_id,is_mail,has_chinese" &
                " FROM folder_stats WHERE folder_path=@p", _db)

                cmd.Parameters.AddWithValue("@p", fPath)
                Using reader = cmd.ExecuteReader()
                    If Not reader.Read() Then Return Nothing
                    Return New FolderStatsDbRow With {.mc = If(reader.IsDBNull(0), -1, reader.GetInt32(0)),
                                                      .mca = If(reader.IsDBNull(1), -1, reader.GetInt32(1)),
                                                      .fc = If(reader.IsDBNull(2), -1, reader.GetInt32(2)),
                                                      .fca = If(reader.IsDBNull(3), -1, reader.GetInt32(3)),
                                                      .fs = If(reader.IsDBNull(4), -1L, reader.GetInt64(4)),
                                                      .fsa = If(reader.IsDBNull(5), -1L, reader.GetInt64(5)),
                                                      .snap = If(reader.IsDBNull(6), -1, reader.GetInt32(6)),
                                                      .eid = If(reader.IsDBNull(7), "", reader.GetString(7)),
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
    Friend Function DbGetAttachMailList(fPath As String) As AttachMailListDbResult
        ' ---------------------------------------------------------------
        ' DbGetAttachMailList — 讀取 attach_maillist WHERE folder_path=? 的所有行
        ' 回傳 Nothing 表示 DB 中無此資料夾的郵件記錄
        ' ---------------------------------------------------------------
        If _iLikeNoisy Then _dbg(" ├ 開始")
        If _db Is Nothing Then Return Nothing

        Try
            Dim result As New AttachMailListDbResult()
            Dim hasRecord As Boolean = False ' by Claude Sonnet 4.6, 2026/05/06: 補齊變數宣告
            Using cmd As New SqliteCommand(
                "SELECT entry_id,subject,msg_size,received_time,sender_name,attach_count,item_count_snap" &
                " FROM attach_maillist WHERE folder_path=@p", _db)

                cmd.Parameters.AddWithValue("@p", fPath)
                hasRecord = False ' 移除 Dim，直接使用外層變數
                Using reader = cmd.ExecuteReader()
                    While reader.Read() ' item_count_snap 整個 folder 共用同一值，每行都一樣，讀最後一次即可
                        hasRecord = True
                        result.Snap = If(reader.IsDBNull(6), -1, reader.GetInt32(6))
                        Dim eid = reader.GetString(0)
                        If eid.StartsWith("EMPTY_ATTACH_") Then Continue While

                        Dim mail As New MailItemInfo()
                        mail.EntryID = eid
                        mail.Subject = If(reader.IsDBNull(1), "", reader.GetString(1))
                        mail.Size = reader.GetInt64(2)
                        Dim dtStr = If(reader.IsDBNull(3), "", reader.GetString(3))
                        DateTime.TryParse(dtStr, mail.ReceivedTime)
                        mail.SenderName = If(reader.IsDBNull(4), "", reader.GetString(4))
                        mail.AttachCount = reader.GetInt32(5)
                        mail.FolderPath = fPath ' 確保填入路徑
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
    Friend Function DbGetBasicMailInfo(fPath As String) As (Mails As List(Of (Mail As MailItemInfo, Topic As String)), Snap As Integer)?
        ' ---------------------------------------------------------------
        ' DbGetBasicMailInfo — 讀取 basic_maillist WHERE folder_path=? 的所有行
        ' 回傳 Nothing 表示 DB 中無此資料夾的郵件記錄
        ' ---------------------------------------------------------------
        If _iLikeNoisy Then _dbg(" ├ 開始")
        If _db Is Nothing Then Return Nothing
        Try
            Dim result As New List(Of (Mail As MailItemInfo, Topic As String))(1024)
            Dim snap As Integer = -1
            Dim hasRecord As Boolean = False ' by Claude Sonnet 4.6, 2026/05/06: 補齊變數宣告

            ' 2026/05/06 by Claude: 新增 message_id(6), sender_email(7); item_count_snap 移至索引 8
            Using cmd As New SqliteCommand(
                "SELECT entry_id,subject,msg_size,received_time,sender_name,topic," &
                "message_id,sender_email,item_count_snap" &
                " FROM basic_maillist WHERE folder_path=@p", _db)
                cmd.Parameters.AddWithValue("@p", fPath)
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        hasRecord = True
                        snap = If(reader.IsDBNull(8), -1, reader.GetInt32(8))
                        Dim eid = reader.GetString(0)
                        If eid.StartsWith("EMPTY_BASIC_") Then Continue While

                        Dim mail As New MailItemInfo()
                        mail.EntryID = eid
                        mail.Subject = If(reader.IsDBNull(1), "", reader.GetString(1))
                        mail.Size = reader.GetInt64(2)
                        Dim dtStr = If(reader.IsDBNull(3), "", reader.GetString(3))
                        DateTime.TryParse(dtStr, mail.ReceivedTime)
                        mail.SenderName = If(reader.IsDBNull(4), "", reader.GetString(4))
                        mail.FolderPath = fPath
                        mail.MessageID = If(reader.IsDBNull(6), "", reader.GetString(6))
                        mail.SenderEmail = If(reader.IsDBNull(7), "", reader.GetString(7))
                        Dim topic = If(reader.IsDBNull(5), "", reader.GetString(5))
                        result.Add((mail, topic))
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
    Friend Function DbGetAttachFilenames(entryId As String) As List(Of String)
        ' ---------------------------------------------------------------
        ' DbGetAttachFilenames — 讀取 attach_filenames WHERE entry_id=? 的一行
        ' 回傳 Nothing 表示 DB 中無此 EntryID
        ' ---------------------------------------------------------------
        If _iLikeNoisy Then _dbg(" ├ 開始")
        If _db Is Nothing Then Return Nothing

        Try
            Using cmd As New SqliteCommand(
                "SELECT filenames FROM attach_filenames WHERE entry_id=@eid", _db)

                cmd.Parameters.AddWithValue("@eid", entryId)
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
    Friend Function DbGetYearCountsForFolder(fPath As String) As ConcurrentDictionary(Of Integer, Integer)
        ' ---------------------------------------------------------------
        ' DbGetYearCountsForFolder — 讀取 year_counts WHERE folder_path=? 的所有行
        ' 供 ComputeYearCountsAsync 在記憶體 miss 時先查 DB，避免 COM 呼叫
        ' 回傳 Nothing 表示 DB 中無此資料夾的年份記錄
        ' ---------------------------------------------------------------
        If _iLikeNoisy Then _dbg(" ├ 開始")
        If _db Is Nothing Then Return Nothing

        Try
            Dim result As New ConcurrentDictionary(Of Integer, Integer)()
            Using cmd As New SqliteCommand("SELECT year,count FROM year_counts WHERE folder_path=@p", _db)
                cmd.Parameters.AddWithValue("@p", fPath)
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        result(reader.GetInt32(0)) = reader.GetInt32(1)
                    End While
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
        '   ├ 2026/04/09 修正：改用三欄 PK，接收 (folderPath, year) 兩個參數
        ' ---------------------------------------------------------------
        If _iLikeNoisy Then _dbg(" ├ 開始")
        If _db Is Nothing Then Return Nothing
        Try
            Dim result As New ConcurrentDictionary(Of Integer, Integer)()
            Using cmd As New SqliteCommand("SELECT month,count FROM month_counts WHERE folder_path=@fp AND year=@yr", _db)
                cmd.Parameters.AddWithValue("@fp", fPath)
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
    Friend Function DbGetSubFolderIDList(rootPath As String, isIncludeAll As Boolean) As List(Of FolderStatsDbRow)
        ' ---------------------------------------------------------------
        ' DbGetSubFolderIDList — [優化 BFS] 利用 LIKE 一次抓出整棵子樹的所有資料夾身分證
        ' ---------------------------------------------------------------
        If _iLikeNoisy Then _dbg(" ├ 開始")
        If _db Is Nothing Then Return Nothing
        Try
            Dim result As New List(Of FolderStatsDbRow)(512)
            ' 過濾條件: 路徑以 rootPath 開頭，且 entry_id 不為空。若沒勾全選，則只抓 is_mail=1 的。
            Dim filter = If(isIncludeAll, "", " AND is_mail=1")
            Dim sql = $"SELECT folder_path,entry_id,store_id,is_mail,has_chinese FROM folder_stats " &
                      $"WHERE folder_path LIKE @p || '%' AND entry_id IS NOT NULL" & filter

            Using cmd As New SqliteCommand(sql, _db)
                cmd.Parameters.AddWithValue("@p", rootPath)
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        result.Add(New FolderStatsDbRow With {.path = reader.GetString(0),
                                                              .eid = reader.GetString(1),
                                                              .sid = reader.GetString(2),
                                                              .isMail = reader.GetInt32(3),
                                                              .hasCh = reader.GetInt32(4)})
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
        If _db Is Nothing Then Return Nothing
        Try
            Dim result As New List(Of FolderStatsDbRow)(512)
            ' SQL 邏輯: 
            ' 1. 找出 folder_path 以 parentPath + "\" 開頭。
            ' 2. 且不包含更深層的 "\" (代表是直屬子項) 。注意: 此邏輯在路徑分隔符不一致時需調整。
            ' 3. 按照 has_chinese ASC (0=英, 1=中, 故英優先) 排序。
            Dim filter = If(isIncludeAll, "", " AND is_mail=1")

            ' 精確匹配直屬子目錄：利用 LENGTH + REPLACE 來算出層級
            ' 或是利用路徑字串特性：新的路徑長度應該是在 parent 之後且沒有多餘的層級
            ' 簡化做法：目前專案路徑是用 \ 分隔。
            Dim sql = "SELECT folder_path,entry_id,store_id,is_mail,has_chinese FROM folder_stats " &
                      "WHERE folder_path LIKE @p || '\%' AND entry_id IS NOT NULL " & filter &
                      " AND folder_path NOT LIKE @p || '\%\%' " & ' 排除第二層以後的
                      " ORDER BY has_chinese ASC, folder_path ASC"

            Using cmd As New SqliteCommand(sql, _db)
                cmd.Parameters.AddWithValue("@p", parentPath)
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        result.Add(New FolderStatsDbRow With {.eid = reader.GetString(1),
                                                              .sid = reader.GetString(2),
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
    Friend Sub DbSaveMonthCountsSingle(fPath As String, year As Integer, monthCounts As ConcurrentDictionary(Of Integer, Integer))
        ' ---------------------------------------------------------------
        ' DbSaveMonthCountsSingle — 增量寫入單一 (folder_path, year) 的月份分布
        ' 在 GetMonthCountsForYearL3 完成 L3 COM 計算後立刻呼叫，不等待 SaveCache 按鈕。
        ' 使用獨立 Transaction 包住最多 12 筆，確保原子性。
        '   ├ 2026/04/09 新增 by Claude：解決月份快取只在記憶體、SaveCache 才寫 DB 的問題
        '   ├ 根本原因：若該 session 沒點過月份視圖就不 SaveCache，下次仍打 COM
        '   ├ 修正後：每次 L3 計算完月份後立刻持久化，下次 DB lazy 直接命中
        ' ---------------------------------------------------------------
        If _iLikeNoisy Then _dbg(" ├ 開始")
        If _db Is Nothing OrElse monthCounts Is Nothing OrElse monthCounts.IsEmpty Then Return
        Try
            Dim ts = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            Using txn = _db.BeginTransaction()
                Using cmd As New SqliteCommand(
                    "INSERT OR REPLACE INTO month_counts (folder_path,year,month,count,updated_at) VALUES (@fp,@yr,@mo,@cnt,@ts)", _db, txn)
                    cmd.Parameters.Add("@fp", SqliteType.Text).Value = fPath
                    cmd.Parameters.Add("@yr", SqliteType.Integer).Value = year
                    cmd.Parameters.Add("@mo", SqliteType.Integer)
                    cmd.Parameters.Add("@cnt", SqliteType.Integer)
                    cmd.Parameters.Add("@ts", SqliteType.Text).Value = ts
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
#End Region


End Class
