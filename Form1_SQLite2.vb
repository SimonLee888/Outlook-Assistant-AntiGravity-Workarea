Imports Microsoft.Data.Sqlite
Imports System.Collections.Concurrent
Imports System.Text.Json

' ==============================================================
' Form1_SQLite2.vb  —  SQLite 持久化快取層
' ==============================================================
' 目的: 把記憶體 ConcurrentDictionary 快取持久化到 SSD，下次啟動可快速重建
'
' 架構:
'   L1 / L2 / L2.5  (完全不知道 SQLite 的存在)
'   Form1_SQLite2.vb  (本檔)
'     InitDatabase()              建 connection + CREATE TABLE IF NOT EXISTS
'     SaveCachesToSQLiteAsync()   手動存入 (Setting 頁 SaveCache 按鈕)
'     LoadCachesFromSQLiteAsync() 手動讀出 (Setting 頁 LoadCache 按鈕)
'     PurgeStaleFolders()         清除孤兒 row，SaveCache 時順帶呼叫
'     GetDbStats()                DB 統計供 debug 顯示
'     CloseDatabase()             FormClosing 時呼叫
'
'   三張表合一個 cache.db (Application.StartupPath):
'     folder_stats      — 資料夾層級六個數字快取 + content_count_snapshot
'     mail_basic        — Tab3 Phase1 候選郵件基本資訊 (MailItemInfo)
'     mail_attachments  — Tab3 Phase2 附件檔名清單 (JSON array)
'
' 設計決策 (2026-04-06):
'   1. 三張表合一個 cache.db，跨表 Transaction 保證原子性，一個 Connection 管理最簡單
'   2. 手動控制 (SaveCache / LoadCache 按鈕)，Debug 階段方便目視確認正確性
'      正式版再切換成 L2.5 lazy SELECT + 增量寫入
'   3. content_count_snapshot 存 _cacheMailCount[path] 的值 (即 PR_CONTENT_COUNT 的讀取結果)
'      Load 後可快速判斷快取是否仍有效，完全不需要呼叫任何 COM
'   4. MailItemInfo 欄位以文字儲存；List(Of String) 附件名稱序列化為 JSON array
'   5. _cacheFolderTree / _cacheSubFolderList 含 COM 物件，永遠不寫入 SQLite
'   6. LoadFolderStatsInner 使用 TryAdd：若記憶體已有值 (L2.5 已讀過)，保留記憶體版本
'      若想強制以 DB 為準 (完整重置)，改用直接賦值 _cacheMailCount(path) = ...
' ==============================================================

Partial Class Form1

    ' ---------------------------------------------------------------
    ' 私有成員
    ' ---------------------------------------------------------------
    Private _db As SqliteConnection = Nothing
    Private ReadOnly _dbPath As String =
        IO.Path.Combine(Application.StartupPath, "cache.db")

    ' ---------------------------------------------------------------
    ' InitDatabase — 建立或開啟 cache.db，確保三張表與索引存在
    ' 在 UI 執行緒呼叫 (Form1_Load)，SQLite DDL 量小，不需要 Async
    ' ---------------------------------------------------------------
    Friend Sub InitDatabase()
        Try
            _db = New SqliteConnection($"Data Source={_dbPath};Mode=ReadWriteCreate;Cache=Shared")
            _db.Open()
            Dbg("InitDatabase", $"已開啟: {_dbPath}")
            Using cmd As New SqliteCommand(GetCreateTablesSql(), _db)
                cmd.ExecuteNonQuery()
            End Using
            Dbg("InitDatabase", "資料表確認完成")
        Catch ex As System.Exception
            Dbg("InitDatabase ERROR", ex.Message)
            _db = Nothing   ' 出錯就設 Nothing，後續所有 SQLite 操作因此自動跳過
        End Try
    End Sub

    Private Function GetCreateTablesSql() As String
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
    updated_at              TEXT
);
CREATE TABLE IF NOT EXISTS mail_basic (
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
CREATE INDEX IF NOT EXISTS idx_mb_folder ON mail_basic(folder_path);
CREATE TABLE IF NOT EXISTS mail_attachments (
    entry_id        TEXT    PRIMARY KEY,
    folder_path     TEXT    NOT NULL,
    filenames       TEXT,
    msg_size        INTEGER,
    updated_at      TEXT
);
CREATE INDEX IF NOT EXISTS idx_ma_folder ON mail_attachments(folder_path);"
    End Function

    ' ---------------------------------------------------------------
    ' SaveCachesToSQLiteAsync — 把記憶體快取全部存入 SQLite
    ' 對應 Setting 頁 SaveCache 按鈕
    ' ---------------------------------------------------------------
    Friend Async Function SaveCachesToSQLiteAsync() As Task
        If _db Is Nothing Then Dbg("SaveCache SKIP", "DB 未初始化") : Return

        Dim sw As New Diagnostics.Stopwatch : sw.Start()
        Dim savedFolders, savedMailBasic, savedAttach As Integer
        Try
            ProgressBar1.Text = "正在存入快取..." : Cursor = Cursors.WaitCursor

            ' SQLite I/O 在背景執行緒，不阻塞 UI
            Dim r = Await Task.Run(Function()
                                       Using txn As SqliteTransaction = _db.BeginTransaction()
                                           Try
                                               Dim f = SaveFolderStatsInner(txn)
                                               Dim b = SaveMailBasicInner(txn)
                                               Dim a = SaveMailAttachmentsInner(txn)
                                               txn.Commit()
                                               Return (f, b, a)
                                           Catch ex As System.Exception
                                               txn.Rollback() : Throw
                                           End Try
                                       End Using
                                   End Function)
            savedFolders = r.Item1 : savedMailBasic = r.Item2 : savedAttach = r.Item3
            sw.Stop()
            Dim msg = $"folder_stats: {savedFolders} 筆 / mail_basic: {savedMailBasic} 筆 / mail_attachments: {savedAttach} 筆 / 耗時: {sw.Elapsed.TotalSeconds:0.000} 秒"
            ProgressBar1.Text = $"SaveCache 完成 — {msg}"
            Dbg("SaveCache 完成", msg)
        Catch ex As System.Exception
            ProgressBar1.Text = "SaveCache 失敗"
            Dbg("SaveCache ERROR", ex.Message)
        Finally
            Cursor = Cursors.Default
        End Try
    End Function

    ' ---------------------------------------------------------------
    ' SaveFolderStatsInner — Transaction 內批次寫入 folder_stats
    ' 注意: 在 Task.Run 背景執行緒呼叫，不可碰 UI 控制項
    ' content_count_snapshot = _cacheMailCount[path]，即 PR_CONTENT_COUNT 讀取結果
    ' ---------------------------------------------------------------
    Private Function SaveFolderStatsInner(txn As SqliteTransaction) As Integer
        Dim sql = "INSERT OR REPLACE INTO folder_stats" &
                  " (folder_path,mail_count,mail_count_all,folder_count,folder_count_all," &
                  "  folder_size,folder_size_all,content_count_snapshot,updated_at)" &
                  " VALUES (@path,@mc,@mca,@fc,@fca,@fs,@fsa,@snap,@ts)"

        ' 蒐集六個 dict 中所有出現過的 folder_path 聯集
        Dim allPaths As New HashSet(Of String)()
        For Each k In _cacheMailCount.Keys : allPaths.Add(k) : Next
        For Each k In _cacheMailCountAll.Keys : allPaths.Add(k) : Next
        For Each k In _cacheFolderCount.Keys : allPaths.Add(k) : Next
        For Each k In _cacheFolderCountAll.Keys : allPaths.Add(k) : Next
        For Each k In _cacheFolderSize.Keys : allPaths.Add(k) : Next
        For Each k In _cacheFolderSizeAll.Keys : allPaths.Add(k) : Next

        Dim count As Integer = 0
        Dim ts = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
        Using cmd As New SqliteCommand(sql, _db, txn)
            cmd.Parameters.Add("@path", SqliteType.Text)
            cmd.Parameters.Add("@mc",   SqliteType.Integer)
            cmd.Parameters.Add("@mca",  SqliteType.Integer)
            cmd.Parameters.Add("@fc",   SqliteType.Integer)
            cmd.Parameters.Add("@fca",  SqliteType.Integer)
            cmd.Parameters.Add("@fs",   SqliteType.Integer)
            cmd.Parameters.Add("@fsa",  SqliteType.Integer)
            cmd.Parameters.Add("@snap", SqliteType.Integer)
            cmd.Parameters.Add("@ts",   SqliteType.Text)

            For Each path In allPaths
                Dim mc, mca, fc, fca As Integer : Dim fs, fsa As Long
                _cacheMailCount.TryGetValue(path, mc)
                _cacheMailCountAll.TryGetValue(path, mca)
                _cacheFolderCount.TryGetValue(path, fc)
                _cacheFolderCountAll.TryGetValue(path, fca)
                _cacheFolderSize.TryGetValue(path, fs)
                _cacheFolderSizeAll.TryGetValue(path, fsa)
                cmd.Parameters("@path").Value = path
                cmd.Parameters("@mc").Value   = mc
                cmd.Parameters("@mca").Value  = mca
                cmd.Parameters("@fc").Value   = fc
                cmd.Parameters("@fca").Value  = fca
                cmd.Parameters("@fs").Value   = fs
                cmd.Parameters("@fsa").Value  = fsa
                cmd.Parameters("@snap").Value = mc  ' content_count_snapshot = _cacheMailCount
                cmd.Parameters("@ts").Value   = ts
                cmd.ExecuteNonQuery() : count += 1
            Next
        End Using
        Return count
    End Function

    ' ---------------------------------------------------------------
    ' SaveMailBasicInner — Transaction 內批次寫入 mail_basic（Tab3 Phase1）
    ' ---------------------------------------------------------------
    Private Function SaveMailBasicInner(txn As SqliteTransaction) As Integer
        Dim sql = "INSERT OR REPLACE INTO mail_basic" &
                  " (entry_id,folder_path,subject,msg_size,received_time,sender_name,attach_count,item_count_snap,updated_at)" &
                  " VALUES (@eid,@fp,@subj,@sz,@rt,@sn,@ac,@snap,@ts)"
        Dim count As Integer = 0
        Dim ts = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
        Using cmd As New SqliteCommand(sql, _db, txn)
            cmd.Parameters.Add("@eid",  SqliteType.Text)
            cmd.Parameters.Add("@fp",   SqliteType.Text)
            cmd.Parameters.Add("@subj", SqliteType.Text)
            cmd.Parameters.Add("@sz",   SqliteType.Integer)
            cmd.Parameters.Add("@rt",   SqliteType.Text)
            cmd.Parameters.Add("@sn",   SqliteType.Text)
            cmd.Parameters.Add("@ac",   SqliteType.Integer)
            cmd.Parameters.Add("@snap", SqliteType.Integer)
            cmd.Parameters.Add("@ts",   SqliteType.Text)

            ' _cachePhase1tab3: Dictionary(Of String, FolderCacheTab3)
            ' key = folder_path, value.mailWithAttachment = List(Of MailItemInfo)
            For Each kvp In _cachePhase1tab3
                Dim fp = kvp.Key : Dim snap = kvp.Value.ItemCountWhenCached
                For Each mail In kvp.Value.mailWithAttachment
                    cmd.Parameters("@eid").Value  = mail.EntryID
                    cmd.Parameters("@fp").Value   = fp
                    cmd.Parameters("@subj").Value = If(mail.Subject, "")
                    cmd.Parameters("@sz").Value   = mail.Size
                    cmd.Parameters("@rt").Value   = mail.ReceivedTime.ToString("yyyy-MM-dd HH:mm:ss")
                    cmd.Parameters("@sn").Value   = If(mail.SenderName, "")
                    cmd.Parameters("@ac").Value   = mail.AttachCount
                    cmd.Parameters("@snap").Value = snap
                    cmd.Parameters("@ts").Value   = ts
                    cmd.ExecuteNonQuery() : count += 1
                Next
            Next
        End Using
        Return count
    End Function

    ' ---------------------------------------------------------------
    ' SaveMailAttachmentsInner — Transaction 內批次寫入 mail_attachments（Tab3 Phase2）
    ' folder_path 透過反查 _cachePhase1tab3 取得（_cacheAttachFilename key 是 EntryID）
    ' ---------------------------------------------------------------
    Private Function SaveMailAttachmentsInner(txn As SqliteTransaction) As Integer
        Dim sql = "INSERT OR REPLACE INTO mail_attachments" &
                  " (entry_id,folder_path,filenames,msg_size,updated_at)" &
                  " VALUES (@eid,@fp,@fn,@sz,@ts)"
        Dim count As Integer = 0
        Dim ts = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

        ' 反查 EntryID → folder_path（從 Phase1 快取中建立對應表）
        Dim entryToFolder As New Dictionary(Of String, String)()
        For Each kvp In _cachePhase1tab3
            For Each mail In kvp.Value.mailWithAttachment
                If Not entryToFolder.ContainsKey(mail.EntryID) Then entryToFolder(mail.EntryID) = kvp.Key
            Next
        Next

        Using cmd As New SqliteCommand(sql, _db, txn)
            cmd.Parameters.Add("@eid", SqliteType.Text)
            cmd.Parameters.Add("@fp",  SqliteType.Text)
            cmd.Parameters.Add("@fn",  SqliteType.Text)
            cmd.Parameters.Add("@sz",  SqliteType.Integer)
            cmd.Parameters.Add("@ts",  SqliteType.Text)

            For Each kvp In _cacheAttachFilename
                Dim fp = "" : entryToFolder.TryGetValue(kvp.Key, fp)
                cmd.Parameters("@eid").Value = kvp.Key
                cmd.Parameters("@fp").Value  = fp
                cmd.Parameters("@fn").Value  = JsonSerializer.Serialize(kvp.Value)
                cmd.Parameters("@sz").Value  = DBNull.Value  ' Phase2 沒有直接存 msg_size
                cmd.Parameters("@ts").Value  = ts
                cmd.ExecuteNonQuery() : count += 1
            Next
        End Using
        Return count
    End Function

    ' ---------------------------------------------------------------
    ' LoadCachesFromSQLiteAsync — 從 SQLite 讀回所有快取（Bulk Load）
    ' 對應 Setting 頁 LoadCache 按鈕，Debug 階段使用
    ' ---------------------------------------------------------------
    Friend Async Function LoadCachesFromSQLiteAsync() As Task
        If _db Is Nothing Then Dbg("LoadCache SKIP", "DB 未初始化") : Return

        Dim sw As New Diagnostics.Stopwatch : sw.Start()
        Try
            ProgressBar1.Text = "正在載入快取..." : Cursor = Cursors.WaitCursor

            Dim r = Await Task.Run(Function()
                                       Dim f = LoadFolderStatsInner()
                                       Dim b = LoadMailBasicInner()
                                       Dim a = LoadMailAttachmentsInner()
                                       Return (f, b, a)
                                   End Function)
            sw.Stop()
            Dim msg = $"folder_stats: {r.Item1} 筆 / mail_basic: {r.Item2} 筆 / mail_attachments: {r.Item3} 筆 / 耗時: {sw.Elapsed.TotalSeconds:0.000} 秒"
            ProgressBar1.Text = $"LoadCache 完成 — {msg}"
            Dbg("LoadCache 完成", msg)
        Catch ex As System.Exception
            ProgressBar1.Text = "LoadCache 失敗"
            Dbg("LoadCache ERROR", ex.Message)
        Finally
            Cursor = Cursors.Default
        End Try
    End Function

    ' ---------------------------------------------------------------
    ' LoadFolderStatsInner — 讀回六個數字快取
    ' 使用 TryAdd：記憶體已有值時保留記憶體版本（不覆蓋 L2.5 剛讀進來的較新值）
    ' ---------------------------------------------------------------
    Private Function LoadFolderStatsInner() As Integer
        Dim count As Integer = 0
        Using cmd As New SqliteCommand(
            "SELECT folder_path,mail_count,mail_count_all,folder_count,folder_count_all,folder_size,folder_size_all FROM folder_stats",
            _db)
            Using reader = cmd.ExecuteReader()
                While reader.Read()
                    Dim path = reader.GetString(0)
                    _cacheMailCount.TryAdd(path,    reader.GetInt32(1))
                    _cacheMailCountAll.TryAdd(path, reader.GetInt32(2))
                    _cacheFolderCount.TryAdd(path,  reader.GetInt32(3))
                    _cacheFolderCountAll.TryAdd(path, reader.GetInt32(4))
                    _cacheFolderSize.TryAdd(path,   reader.GetInt64(5))
                    _cacheFolderSizeAll.TryAdd(path, reader.GetInt64(6))
                    count += 1
                End While
            End Using
        End Using
        Return count
    End Function

    ' ---------------------------------------------------------------
    ' LoadMailBasicInner — 重建 _cachePhase1tab3（按 folder_path 分組）
    ' ---------------------------------------------------------------
    Private Function LoadMailBasicInner() As Integer
        Dim count As Integer = 0
        ' 暫存用：先按 folder_path 分組收集，最後一次性寫入 _cachePhase1tab3
        Dim tempDict As New Dictionary(Of String, (snap As Integer, mails As List(Of MailItemInfo)))()

        Using cmd As New SqliteCommand(
            "SELECT entry_id,folder_path,subject,msg_size,received_time,sender_name,attach_count,item_count_snap FROM mail_basic",
            _db)
            Using reader = cmd.ExecuteReader()
                While reader.Read()
                    Dim fp = reader.GetString(1)
                    If Not tempDict.ContainsKey(fp) Then
                        tempDict(fp) = (reader.GetInt32(7), New List(Of MailItemInfo)())
                    End If
                    Dim mail As New MailItemInfo()
                    mail.EntryID    = reader.GetString(0)
                    mail.Subject    = If(reader.IsDBNull(2), "", reader.GetString(2))
                    mail.Size       = reader.GetInt64(3)
                    Dim dtStr = If(reader.IsDBNull(4), "", reader.GetString(4))
                    DateTime.TryParse(dtStr, mail.ReceivedTime)
                    mail.SenderName = If(reader.IsDBNull(5), "", reader.GetString(5))
                    mail.AttachCount = reader.GetInt32(6)
                    tempDict(fp).mails.Add(mail) : count += 1
                End While
            End Using
        End Using

        For Each kvp In tempDict
            If Not _cachePhase1tab3.ContainsKey(kvp.Key) Then
                _cachePhase1tab3(kvp.Key) = New FolderCacheTab3 With {
                    .mailWithAttachment = kvp.Value.mails,
                    .ItemCountWhenCached = kvp.Value.snap}
            End If
        Next
        Return count
    End Function

    ' ---------------------------------------------------------------
    ' LoadMailAttachmentsInner — 重建 _cacheAttachFilename（JSON 反序列化）
    ' ---------------------------------------------------------------
    Private Function LoadMailAttachmentsInner() As Integer
        Dim count As Integer = 0
        Using cmd As New SqliteCommand("SELECT entry_id,filenames FROM mail_attachments", _db)
            Using reader = cmd.ExecuteReader()
                While reader.Read()
                    Dim eid = reader.GetString(0)
                    Dim fnJson = If(reader.IsDBNull(1), "[]", reader.GetString(1))
                    Try
                        Dim list = JsonSerializer.Deserialize(Of List(Of String))(fnJson)
                        _cacheAttachFilename.TryAdd(eid, list) : count += 1
                    Catch ex As System.Exception
                        Dbg("LoadAttach 解析失敗", $"{eid}: {ex.Message}")
                    End Try
                End While
            End Using
        End Using
        Return count
    End Function

    ' ---------------------------------------------------------------
    ' PurgeStaleFolders — 刪除已不存在於 Outlook 的資料夾孤兒行
    ' liveFolderPaths = 目前 Outlook 中所有資料夾路徑的 HashSet
    ' 建議在 SaveCachesToSQLiteAsync 之前呼叫，先清後存
    ' ---------------------------------------------------------------
    Friend Sub PurgeStaleFolders(liveFolderPaths As HashSet(Of String))
        If _db Is Nothing Then Return
        Try
            ' 讀出 DB 中所有 folder_path
            Dim dbPaths As New List(Of String)()
            Using cmd As New SqliteCommand("SELECT folder_path FROM folder_stats", _db)
                Using reader = cmd.ExecuteReader()
                    While reader.Read() : dbPaths.Add(reader.GetString(0)) : End While
                End Using
            End Using

            Dim stalePaths = dbPaths.Where(Function(p) Not liveFolderPaths.Contains(p)).ToList()
            If stalePaths.Count = 0 Then Dbg("PurgeStaleFolders", "無孤兒行，略過") : Return

            Dim dF, dB, dA As Integer
            Using txn As SqliteTransaction = _db.BeginTransaction()
                For Each stale In stalePaths
                    Using c1 As New SqliteCommand("DELETE FROM folder_stats WHERE folder_path=@p", _db, txn)
                        c1.Parameters.AddWithValue("@p", stale) : dF += c1.ExecuteNonQuery()
                    End Using
                    Using c2 As New SqliteCommand("DELETE FROM mail_basic WHERE folder_path=@p", _db, txn)
                        c2.Parameters.AddWithValue("@p", stale) : dB += c2.ExecuteNonQuery()
                    End Using
                    Using c3 As New SqliteCommand("DELETE FROM mail_attachments WHERE folder_path=@p", _db, txn)
                        c3.Parameters.AddWithValue("@p", stale) : dA += c3.ExecuteNonQuery()
                    End Using
                Next
                txn.Commit()
            End Using
            Dbg("PurgeStaleFolders 完成", $"刪除 folder_stats:{dF} / mail_basic:{dB} / mail_attachments:{dA}")
        Catch ex As System.Exception
            Dbg("PurgeStaleFolders ERROR", ex.Message)
        End Try
    End Sub

    ' ---------------------------------------------------------------
    ' GetDbStats — 取得 DB 統計摘要，供按鈕顯示
    ' 回傳 (folder_stats 行數, mail_basic 行數, mail_attachments 行數, 檔案大小 KB)
    ' ---------------------------------------------------------------
    Friend Function GetDbStats() As (fc As Integer, mb As Integer, at As Integer, kb As Long)
        If _db Is Nothing Then Return (0, 0, 0, 0)
        Try
            Dim fc, mb, at As Integer
            Using cmd As New SqliteCommand("SELECT COUNT(*) FROM folder_stats",    _db) : fc = Convert.ToInt32(cmd.ExecuteScalar()) : End Using
            Using cmd As New SqliteCommand("SELECT COUNT(*) FROM mail_basic",      _db) : mb = Convert.ToInt32(cmd.ExecuteScalar()) : End Using
            Using cmd As New SqliteCommand("SELECT COUNT(*) FROM mail_attachments", _db) : at = Convert.ToInt32(cmd.ExecuteScalar()) : End Using
            Dim fi As New IO.FileInfo(_dbPath)
            Return (fc, mb, at, If(fi.Exists, fi.Length \ 1024L, 0L))
        Catch ex As System.Exception
            Dbg("GetDbStats ERROR", ex.Message) : Return (0, 0, 0, 0)
        End Try
    End Function

    ' ---------------------------------------------------------------
    ' CloseDatabase — FormClosing 時呼叫，安全關閉 SQLite 連線
    ' ---------------------------------------------------------------
    Friend Sub CloseDatabase()
        If _db Is Nothing Then Return
        Try
            _db.Close() : _db.Dispose() : _db = Nothing
            Dbg("CloseDatabase", "SQLite 連線已關閉")
        Catch ex As System.Exception
            Dbg("CloseDatabase ERROR", ex.Message)
        End Try
    End Sub

End Class
