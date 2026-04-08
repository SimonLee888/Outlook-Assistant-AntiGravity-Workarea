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
'     InitDatabase()              → 建 connection + CREATE TABLE IF NOT EXISTS
'     SaveCachesToSQLiteAsync()   → 手動存入 (Setting 頁 SaveCache 按鈕)
'     LoadCachesFromSQLiteAsync() → 手動讀出 (Setting 頁 LoadCache 按鈕)
'     CleanupOrphanFolderPath()         → 清除孤兒 row，SaveCache 時順帶呼叫
'     GetDatabaseSummary()                → DB 統計供 debug 顯示
'     CloseDatabase()             → FormClosing 時呼叫
'
'   三張表合一個 cache.db (Application.StartupPath):
'     folder_stats      — 資料夾層級六個數字快取 + content_count_snapshot
'     mail_withattachs        — Tab3 Phase1 候選郵件基本資訊 (MailItemInfo)
'     attach_filenames  — Tab3 Phase2 附件檔名清單 (JSON array)
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

    ' 私有成員
    Private _db As SqliteConnection = Nothing
    Private ReadOnly _dbPath As String = IO.Path.Combine(Application.StartupPath, "OLAcache.db")

    ' ---------------------------------------------------------------
    ' DB Row 結構 (供 Form1_ComL3.vb 的 L2.5 函數使用)
    ' ---------------------------------------------------------------
    Friend Class FolderStatsDbRow
        ' folder_stats 一行的讀出結果；-1 代表該欄位在 DB 中為 NULL 或尚未寫入
        Public mc As Integer = -1       ' mail_count
        Public mca As Integer = -1      ' mail_count_all
        Public fc As Integer = -1       ' folder_count
        Public fca As Integer = -1      ' folder_count_all
        Public fs As Long = -1          ' folder_size
        Public fsa As Long = -1         ' folder_size_all
        Public snap As Integer = -1     ' content_count_snapshot (= PR_CONTENT_COUNT at save time)
    End Class

    Friend Class MailBasicDbResult
        ' mail_withattachs WHERE folder_path=? 的讀出結果
        Public Snap As Integer = -1
        Public Mails As New List(Of MailItemInfo)()
    End Class

    Friend Sub InitDatabase()
        ' ---------------------------------------------------------------
        ' InitDatabase — 建立或開啟 cache.db，確保三張表與索引存在
        ' 在 UI 執行緒呼叫 (Form1_Load)，SQLite DDL 量小，不需要 Async
        ' ---------------------------------------------------------------
        Dbg("開始")
        Try
            _db = New SqliteConnection($"Data Source={_dbPath};Mode=ReadWriteCreate;Cache=Shared")
            _db.Open()
            Dbg("", $"已開啟: {_dbPath}")

            Using cmd As New SqliteCommand(GetCreateTablesSql(), _db)
                cmd.ExecuteNonQuery()
            End Using
            Dbg("", "資料表確認完成")

        Catch ex As System.Exception
            Dbg("錯誤", ex.Message)
            _db = Nothing   ' 出錯就設 Nothing，後續所有 SQLite 操作因此自動跳過
        Finally : Dbg("結束")
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
                CREATE TABLE IF NOT EXISTS mail_withattachs (
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
                CREATE INDEX IF NOT EXISTS idx_mb_folder ON mail_withattachs(folder_path);
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
                CREATE INDEX IF NOT EXISTS idx_yc_folder ON year_counts(folder_path);"

    End Function
    Friend Sub CloseDatabase()
        ' ---------------------------------------------------------------
        ' CloseDatabase — FormClosing 時呼叫，安全關閉 SQLite 連線
        ' ---------------------------------------------------------------
        Dbg("開始")

        If _db Is Nothing Then Return

        Try
            _db.Close() : _db.Dispose() : _db = Nothing
            Dbg("", "SQLite 連線已關閉")

        Catch ex As System.Exception
            Dbg("", ex.Message)
        Finally : Dbg("結束")
        End Try

    End Sub

    Friend Async Function SaveCachesToSQLiteAsync() As Task
        ' ---------------------------------------------------------------
        ' SaveCachesToSQLiteAsync — 把記憶體快取全部存入 SQLite
        ' 對應 Setting 頁 SaveCache 按鈕
        ' 流程: ① CleanupOrphanFolderPath (先清孤兒) → ② 批次寫入三張表 → ③ 統計顯示
        ' ---------------------------------------------------------------
        Dbg("開始")
        If _db Is Nothing Then Dbg("", "DB 未初始化") : Return

        Dim sw As New Diagnostics.Stopwatch : sw.Start()
        Dim savedFolders, savedMailBasic, savedAttach As Integer
        Try
            ProgressBar1.Text = "正在存入快取..." : Cursor = Cursors.WaitCursor

            ' ① 先清孤兒：收集目前記憶體快取中所有仍存在的 folder_path，清除 DB 中已不存在的行
            ' 用記憶體快取的 key 聯集代表「目前已知 live 的資料夾」（比重新 BFS 掃 COM 快得多）
            Dim livePaths As New HashSet(Of String)()
            For Each k In _cacheMailCount.Keys : livePaths.Add(k) : Next
            For Each k In _cacheFolderCount.Keys : livePaths.Add(k) : Next
            For Each k In _cacheAttachPreScan.Keys : livePaths.Add(k) : Next
            If livePaths.Count > 0 Then CleanupOrphanFolderPath(livePaths)

            ' ② SQLite I/O 在背景執行緒，不阻塞 UI
            Dim r = Await Task.Run(Function()
                                       Using txn As SqliteTransaction = _db.BeginTransaction()
                                           Try
                                               Dim f = SaveFolderStatsInner(txn)
                                               Dim b = SaveMailBasicInner(txn)
                                               Dim a = SaveAttachFilenamesInner(txn)
                                               Dim y = SaveYearCountsInner(txn)
                                               txn.Commit()
                                               Return (f, b, a, y)
                                           Catch ex As System.Exception
                                               txn.Rollback() : Throw
                                           End Try
                                       End Using
                                   End Function)

            savedFolders = r.Item1 : savedMailBasic = r.Item2 : savedAttach = r.Item3
            Dim savedYears As Integer = r.Item4
            sw.Stop()

            ' ③ 統計：各快取字典目前的 entry 數
            Dim statLine1 = $"① [記憶體] MailCount: {_cacheMailCount.Count} / MailCountAll: {_cacheMailCountAll.Count} / FolderCount: {_cacheFolderCount.Count} / FolderCountAll: {_cacheFolderCountAll.Count}"
            Dim statLine2 = $"② [記憶體] FolderSize: {_cacheFolderSize.Count} / FolderSizeAll: {_cacheFolderSizeAll.Count} / AttachPreScan: {_cacheAttachPreScan.Count} / AttachFilename: {_cacheAttachFilename.Count}"
            Dim statLine3 = $"③ [寫入DB] folder_stats: {savedFolders} 筆 / mail_withattachs: {savedMailBasic} 筆 / attach_filenames: {savedAttach} 筆 / year_counts: {savedYears} 筆 / 耗時: {sw.Elapsed.TotalSeconds:0.000} 秒"
            Dim st = GetDatabaseSummary()
            Dim statLine4 = $"④ [DB現況] folder_stats: {st.fc} 筆 / mail_withattachs: {st.mb} 筆 / attach_filenames: {st.at} 筆 / year_counts: {st.yc} 筆 / 檔案: {st.kb} KB"

            ProgressBar1.Text = $"SaveCache 完成 — {statLine3}"
            ProgressBar2.Text = statLine4
            Dbg(" - [SaveCache]", statLine1)
            Dbg(" - [SaveCache]", statLine2)
            Dbg(" - [SaveCache]", statLine3)
            Dbg(" - [SaveCache]", statLine4)

        Catch ex As System.Exception
            ProgressBar1.Text = "SaveCache 失敗"
            Dbg("錯誤", ex.Message)
        Finally
            Cursor = Cursors.Default
            Dbg("結束")
        End Try

    End Function
    Friend Async Function LoadCachesFromSQLiteAsync() As Task
        ' ---------------------------------------------------------------
        ' LoadCachesFromSQLiteAsync — 從 SQLite 讀回所有快取（Bulk Load）
        ' 對應 Setting 頁 LoadCache 按鈕，Debug 階段使用
        ' 完成後輸出詳細 Dbg 分項：每個快取字典各自載入了幾筆
        ' ---------------------------------------------------------------
        Dbg("開始")
        If _db Is Nothing Then Dbg("", "DB 未初始化") : Return

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
            Dim beforePS = _cacheAttachPreScan.Count
            Dim beforeAF = _cacheAttachFilename.Count

            Dim r = Await Task.Run(Function()
                                       Dim f = LoadFolderStatsInner()
                                       Dim b = LoadMailBasicInner()
                                       Dim a = LoadAttachFilenamesInner()
                                       Dim y = LoadYearCountsInner()
                                       Return (f, b, a, y)
                                   End Function)
            sw.Stop()

            Dim beforeYC As Integer = 0  ' year_counts 沒有 before snapshot，直接顯示載入後總數
            ' 詳細 Dbg：各快取字典 Load 後的增量
            Dim statLine1 = $"① [folder_stats] 讀入 {r.Item1} 筆 — " &
                            $"MailCount +{_cacheMailCount.Count - beforeMC} / " &
                            $"MailCountAll +{_cacheMailCountAll.Count - beforeMCA} / " &
                            $"FolderCount +{_cacheFolderCount.Count - beforeFC} / " &
                            $"FolderCountAll +{_cacheFolderCountAll.Count - beforeFCA}"
            Dim statLine2 = $"② [folder_stats cont.] " &
                            $"FolderSize +{_cacheFolderSize.Count - beforeFS} / " &
                            $"FolderSizeAll +{_cacheFolderSizeAll.Count - beforeFSA}"
            Dim statLine3 = $"③ [mail_withattachs] 讀入 {r.Item2} 筆 → AttachPreScan +{_cacheAttachPreScan.Count - beforePS} 個資料夾"
            Dim statLine4 = $"④ [attach_filenames] 讀入 {r.Item3} 筆 → AttachFilename +{_cacheAttachFilename.Count - beforeAF} 筆"
            Dim statLine_yc = $"⑤ [year_counts] 讀入 {r.Item4} 筆 → _yearCountsCache {_yearCountsCache.Count} 個資料夾"
            Dim st = GetDatabaseSummary()
            Dim statLine5 = $"⑥ [DB現況] folder_stats: {st.fc} 筆 / mail_withattachs: {st.mb} 筆 / attach_filenames: {st.at} 筆 / year_counts: {st.yc} 筆 / 檔案: {st.kb} KB / 耗時: {sw.Elapsed.TotalSeconds:0.000} 秒"

            ProgressBar1.Text = $"LoadCache 完成 — DB: {st.fc}/{st.mb}/{st.at}/{st.yc} 筆，{st.kb} KB，耗時 {sw.Elapsed.TotalSeconds:0.000} 秒"
            ProgressBar2.Text = $"記憶體增量 — mailCount+{_cacheMailCount.Count - beforeMC} / attachFilename+{_cacheAttachFilename.Count - beforeAF} / yearCounts:{_yearCountsCache.Count} 資料夾"
            Dbg(" - [LoadCache]", statLine1)
            Dbg(" - [LoadCache]", statLine2)
            Dbg(" - [LoadCache]", statLine3)
            Dbg(" - [LoadCache]", statLine4)
            Dbg(" - [LoadCache]", statLine_yc)
            Dbg(" - [LoadCache]", statLine5)

        Catch ex As System.Exception
            ProgressBar1.Text = "LoadCache 失敗"
            Dbg("錯誤", ex.Message)
        Finally
            Cursor = Cursors.Default
            Dbg("結束")
        End Try

    End Function
    Friend Sub CleanupOrphanFolderPath(liveFolderPaths As HashSet(Of String))
        ' ---------------------------------------------------------------
        ' CleanupOrphanFolderPath — 刪除已不存在於 Outlook 的資料夾孤兒行
        ' liveFolderPaths = 目前仍有效的資料夾路徑集合
        '   呼叫來源 A: SaveCachesToSQLiteAsync 開頭（用記憶體快取 key 聯集）
        '   呼叫來源 B: RenewCache_Click（用 GetSubFolderList BFS 掃 COM 取得完整清單）
        ' ---------------------------------------------------------------
        Dbg("開始", $"live 資料夾數: {liveFolderPaths.Count}")
        If _db Is Nothing Then Return
        Try
            ' 讀出 DB 中所有 folder_path
            Dim dbPaths As New List(Of String)()
            Using cmd As New SqliteCommand("SELECT folder_path FROM folder_stats", _db)
                Using reader = cmd.ExecuteReader()
                    While reader.Read() : dbPaths.Add(reader.GetString(0)) : End While
                End Using
            End Using
            Dbg("", $"DB 中有 {dbPaths.Count} 個資料夾路徑")

            Dim stalePaths = dbPaths.Where(Function(p) Not liveFolderPaths.Contains(p)).ToList()
            If stalePaths.Count = 0 Then Dbg("", "未發現孤兒快取，略過") : Return

            ' 每個孤兒路徑輸出到 Dbg 供目視確認
            For Each stale In stalePaths
                Dbg(" 孤兒", stale)
            Next

            Dim dF, dB, dA As Integer
            Using txn As SqliteTransaction = _db.BeginTransaction()
                For Each stale In stalePaths
                    Using c1 As New SqliteCommand("DELETE FROM folder_stats WHERE folder_path=@p", _db, txn)
                        c1.Parameters.AddWithValue("@p", stale) : dF += c1.ExecuteNonQuery()
                    End Using
                    Using c2 As New SqliteCommand("DELETE FROM mail_withattachs WHERE folder_path=@p", _db, txn)
                        c2.Parameters.AddWithValue("@p", stale) : dB += c2.ExecuteNonQuery()
                    End Using
                    Using c3 As New SqliteCommand("DELETE FROM attach_filenames WHERE folder_path=@p", _db, txn)
                        c3.Parameters.AddWithValue("@p", stale) : dA += c3.ExecuteNonQuery()
                    End Using
                Next
                txn.Commit()
            End Using
            Dbg("結束", $"孤兒路徑:{stalePaths.Count} 個 / 刪除 folder_stats:{dF} 行 / mail_withattachs:{dB} 行 / attach_filenames:{dA} 行")

        Catch ex As System.Exception
            Dbg("錯誤", ex.Message)
        End Try

    End Sub
    Friend Function GetDatabaseSummary() As (fc As Integer, mb As Integer, at As Integer, yc As Integer, kb As Long)
        ' ---------------------------------------------------------------
        ' GetDatabaseSummary — 取得 DB 統計摘要，供按鈕顯示
        ' 回傳 (folder_stats 行數, mail_withattachs 行數, attach_filenames 行數, year_counts 行數, 檔案大小 KB)
        ' ---------------------------------------------------------------
        Dbg(" - 開始")
        If _db Is Nothing Then Return (0, 0, 0, 0, 0)

        Try
            Dim fc, mb, at, yc As Integer
            Using cmd As New SqliteCommand("SELECT COUNT(*) FROM folder_stats", _db) : fc = Convert.ToInt32(cmd.ExecuteScalar()) : End Using
            Using cmd As New SqliteCommand("SELECT COUNT(*) FROM mail_withattachs", _db) : mb = Convert.ToInt32(cmd.ExecuteScalar()) : End Using
            Using cmd As New SqliteCommand("SELECT COUNT(*) FROM attach_filenames", _db) : at = Convert.ToInt32(cmd.ExecuteScalar()) : End Using
            Using cmd As New SqliteCommand("SELECT COUNT(*) FROM year_counts", _db) : yc = Convert.ToInt32(cmd.ExecuteScalar()) : End Using
            Dim fi As New IO.FileInfo(_dbPath)
            Return (fc, mb, at, yc, If(fi.Exists, fi.Length \ 1024L, 0L))

        Catch ex As System.Exception
            Dbg(" - 錯誤", ex.Message) : Return (0, 0, 0, 0, 0)
        Finally : Dbg(" - 結束")
        End Try

    End Function

    Private Function SaveFolderStatsInner(txn As SqliteTransaction) As Integer
        ' ---------------------------------------------------------------
        ' SaveFolderStatsInner — Transaction 內批次寫入 folder_stats
        ' 注意: 在 Task.Run 背景執行緒呼叫，不可碰 UI 控制項
        ' content_count_snapshot = _cacheMailCount[path]，即 PR_CONTENT_COUNT 讀取結果
        ' ---------------------------------------------------------------
        Dbg("開始")
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
            cmd.Parameters.Add("@mc", SqliteType.Integer)
            cmd.Parameters.Add("@mca", SqliteType.Integer)
            cmd.Parameters.Add("@fc", SqliteType.Integer)
            cmd.Parameters.Add("@fca", SqliteType.Integer)
            cmd.Parameters.Add("@fs", SqliteType.Integer)
            cmd.Parameters.Add("@fsa", SqliteType.Integer)
            cmd.Parameters.Add("@snap", SqliteType.Integer)
            cmd.Parameters.Add("@ts", SqliteType.Text)

            For Each path In allPaths
                ' 2026/04/07 修正 v2: 初始值設 -1 仍然不夠，因為 -1 是整數值會被寫入 DB，
                ' LoadFolderStatsInner 讀回 -1 後直接塞入記憶體快取，
                ' GetCachedFolderCount 命中記憶體回傳 -1 → LoadSubFolderToTreeView 判斷 -1 > 0 為 False → 不顯示 "+"。
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
                cmd.Parameters("@ts").Value = ts
                cmd.ExecuteNonQuery() : count += 1
            Next
        End Using

        Return count
        Dbg("結束")

    End Function
    Private Function SaveYearCountsInner(txn As SqliteTransaction) As Integer
        ' ---------------------------------------------------------------
        ' SaveYearCountsInner — Transaction 內批次寫入 year_counts（Tab2 年份分布）
        ' _yearCountsCache: ConcurrentDictionary(Of String, ConcurrentDictionary(Of Integer, Integer))
        '   key = folder_path, value = { year → count }
        ' 每筆寫入 (folder_path, year, count)，PRIMARY KEY = (folder_path, year)
        ' ---------------------------------------------------------------
        Dbg(" - 開始")
        Dim sql = "INSERT OR REPLACE INTO year_counts (folder_path,year,count,updated_at) VALUES (@fp,@yr,@cnt,@ts)"
        Dim count As Integer = 0
        Dim ts = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

        Using cmd As New SqliteCommand(sql, _db, txn)
            cmd.Parameters.Add("@fp", SqliteType.Text)
            cmd.Parameters.Add("@yr", SqliteType.Integer)
            cmd.Parameters.Add("@cnt", SqliteType.Integer)
            cmd.Parameters.Add("@ts", SqliteType.Text)

            For Each kvp In _yearCountsCache
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
        Dbg(" - 結束")

    End Function
    Private Function SaveMailBasicInner(txn As SqliteTransaction) As Integer
        ' ---------------------------------------------------------------
        ' SaveMailBasicInner — Transaction 內批次寫入 mail_withattachs（Tab3 Phase1）
        ' ---------------------------------------------------------------
        Dbg("開始")
        Dim sql = "INSERT OR REPLACE INTO mail_withattachs" &
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

            ' _cacheAttachPreScan: Dictionary(Of String, FolderCacheTab3)
            ' key = folder_path, value.mailWithAttachment = List(Of MailItemInfo)
            For Each kvp In _cacheAttachPreScan
                Dim fp = kvp.Key : Dim snap = kvp.Value.ItemCountWhenCached
                For Each mail In kvp.Value.mailWithAttachment
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
            Next
        End Using
        Return count
        Dbg("結束")

    End Function
    Private Function SaveAttachFilenamesInner(txn As SqliteTransaction) As Integer
        ' ---------------------------------------------------------------
        ' SaveAttachFilenamesInner — Transaction 內批次寫入 attach_filenames（Tab3 Phase2）
        ' folder_path 透過反查 _cacheAttachPreScan 取得（_cacheAttachFilename key 是 EntryID）
        ' ---------------------------------------------------------------
        Dbg("開始")
        Dim sql = "INSERT OR REPLACE INTO attach_filenames" &
                  " (entry_id,folder_path,filenames,msg_size,updated_at)" &
                  " VALUES (@eid,@fp,@fn,@sz,@ts)"
        Dim count As Integer = 0
        Dim ts = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

        ' 反查 EntryID → folder_path（從 Phase1 快取中建立對應表）
        Dim entryToFolder As New Dictionary(Of String, String)()
        For Each kvp In _cacheAttachPreScan
            For Each mail In kvp.Value.mailWithAttachment
                If Not entryToFolder.ContainsKey(mail.EntryID) Then entryToFolder(mail.EntryID) = kvp.Key
            Next
        Next

        Using cmd As New SqliteCommand(sql, _db, txn)
            cmd.Parameters.Add("@eid", SqliteType.Text)
            cmd.Parameters.Add("@fp", SqliteType.Text)
            cmd.Parameters.Add("@fn", SqliteType.Text)
            cmd.Parameters.Add("@sz", SqliteType.Integer)
            cmd.Parameters.Add("@ts", SqliteType.Text)

            For Each kvp In _cacheAttachFilename
                Dim fp = "" : entryToFolder.TryGetValue(kvp.Key, fp)
                cmd.Parameters("@eid").Value = kvp.Key
                cmd.Parameters("@fp").Value = fp
                cmd.Parameters("@fn").Value = JsonSerializer.Serialize(kvp.Value)
                cmd.Parameters("@sz").Value = DBNull.Value  ' Phase2 沒有直接存 msg_size
                cmd.Parameters("@ts").Value = ts
                cmd.ExecuteNonQuery() : count += 1
            Next
        End Using
        Return count
        Dbg("結束")

    End Function

    Private Function LoadFolderStatsInner() As Integer
        ' ---------------------------------------------------------------
        ' LoadFolderStatsInner — 讀回六個數字快取
        ' 使用 TryAdd：記憶體已有值時保留記憶體版本（不覆蓋 L2.5 剛讀進來的較新值）
        ' 2026/04/07 修正: 每個欄位加 IsDBNull 保護，NULL 代表「從未測量過」，
        '   不可塞入記憶體快取，否則 GetCachedFolderCount 命中 -1 → LoadSubFolderToTreeView
        '   判斷 -1 > 0 為 False → 不顯示 TreeView "+" 加號（bug）。
        ' ---------------------------------------------------------------
        Dbg("開始")
        Dim count As Integer = 0
        Using cmd As New SqliteCommand(
            "SELECT folder_path,mail_count,mail_count_all,folder_count,folder_count_all,folder_size,folder_size_all FROM folder_stats", _db)
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
                    count += 1
                End While
            End Using
        End Using
        Return count
        Dbg("結束")

    End Function
    Private Function LoadYearCountsInner() As Integer
        ' ---------------------------------------------------------------
        ' LoadYearCountsInner — 從 year_counts 重建 _yearCountsCache
        ' 先按 folder_path 分組收集，最後一次性寫入（TryAdd 保留記憶體已有版本）
        ' ---------------------------------------------------------------
        Dbg(" - 開始")
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
            _yearCountsCache.TryAdd(kvp.Key, kvp.Value)
        Next
        Return count
        Dbg(" - 結束")

    End Function
    Private Function LoadMailBasicInner() As Integer
        ' ---------------------------------------------------------------
        ' LoadMailBasicInner — 重建 _cacheAttachPreScan（按 folder_path 分組）
        ' ---------------------------------------------------------------
        Dbg("開始")
        Dim count As Integer = 0
        ' 暫存用：先按 folder_path 分組收集，最後一次性寫入 _cacheAttachPreScan
        Dim tempDict As New Dictionary(Of String, (snap As Integer, mails As List(Of MailItemInfo)))()

        Using cmd As New SqliteCommand("SELECT entry_id,folder_path,subject,msg_size,received_time,sender_name,attach_count,item_count_snap FROM mail_withattachs", _db)
            Using reader = cmd.ExecuteReader()
                While reader.Read()
                    Dim fp = reader.GetString(1)
                    If Not tempDict.ContainsKey(fp) Then
                        tempDict(fp) = (reader.GetInt32(7), New List(Of MailItemInfo)())
                    End If

                    Dim mail As New MailItemInfo()
                    mail.EntryID = reader.GetString(0)
                    mail.Subject = If(reader.IsDBNull(2), "", reader.GetString(2))
                    mail.Size = reader.GetInt64(3)

                    Dim dtStr = If(reader.IsDBNull(4), "", reader.GetString(4))
                    DateTime.TryParse(dtStr, mail.ReceivedTime)
                    mail.SenderName = If(reader.IsDBNull(5), "", reader.GetString(5))
                    mail.AttachCount = reader.GetInt32(6)
                    tempDict(fp).mails.Add(mail) : count += 1
                End While
            End Using
        End Using

        For Each kvp In tempDict
            ' 優化後：
            _cacheAttachPreScan.TryAdd(kvp.Key,
                                       New FolderCacheTab3 With {.mailWithAttachment = kvp.Value.mails,
                                                                 .ItemCountWhenCached = kvp.Value.snap})
        Next
        Return count
        Dbg("結束")

    End Function
    Private Function LoadAttachFilenamesInner() As Integer
        ' ---------------------------------------------------------------
        ' LoadAttachFilenamesInner — 重建 _cacheAttachFilename（JSON 反序列化）
        ' ---------------------------------------------------------------
        Dbg("開始")
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
                        Dbg("錯誤: 解析失敗", $"{eid}: {ex.Message}")
                    End Try
                End While
            End Using
        End Using
        Return count
        Dbg("結束")

    End Function

    ' ==============================================================
    ' Phase 2 — L2.5 lazy SELECT 用的 DB read helper 群
    ' ==============================================================
    ' 設計原則 (2026-04-07):
    '   1. 只做「讀」，不做「寫」。寫入仍由 SaveCachesToSQLiteAsync (SaveCache 按鈕) 批次處理。
    '   2. 回傳 Nothing 表示 DB 中無此筆資料，呼叫端應繼續往 L3 走。
    '   3. 這些函數從 UI 執行緒呼叫，SQLite keyed lookup < 1ms，不需要 Async。
    '   4. FolderStatsDbRow / MailBasicDbResult 定義在本檔，Partial Class 跨檔可見。
    ' ==============================================================
    Friend Function DbGetFolderStats(folderPath As String) As FolderStatsDbRow
        ' ---------------------------------------------------------------
        ' DbGetFolderStats — 讀取 folder_stats 單一 folder_path 的一行
        ' 回傳 Nothing 表示 DB 中無此路徑
        ' ---------------------------------------------------------------
        If _iLikeNoisy Then Dbg(" - 開始")
        If _db Is Nothing Then Return Nothing

        Try
            Using cmd As New SqliteCommand(
                "SELECT mail_count,mail_count_all,folder_count,folder_count_all," &
                "       folder_size,folder_size_all,content_count_snapshot" &
                " FROM folder_stats WHERE folder_path=@p", _db)

                cmd.Parameters.AddWithValue("@p", folderPath)
                Using reader = cmd.ExecuteReader()
                    If Not reader.Read() Then Return Nothing
                    Return New FolderStatsDbRow With {.mc = If(reader.IsDBNull(0), -1, reader.GetInt32(0)),
                                                      .mca = If(reader.IsDBNull(1), -1, reader.GetInt32(1)),
                                                      .fc = If(reader.IsDBNull(2), -1, reader.GetInt32(2)),
                                                      .fca = If(reader.IsDBNull(3), -1, reader.GetInt32(3)),
                                                      .fs = If(reader.IsDBNull(4), -1L, reader.GetInt64(4)),
                                                      .fsa = If(reader.IsDBNull(5), -1L, reader.GetInt64(5)),
                                                      .snap = If(reader.IsDBNull(6), -1, reader.GetInt32(6))}
                End Using
            End Using
        Catch ex As System.Exception
            Dbg(" - 錯誤", $"{folderPath}: {ex.Message}")
        Finally : If _iLikeNoisy Then Dbg(" - 結束")
        End Try
        Return Nothing

    End Function
    Friend Function DbGetMailBasic(folderPath As String) As MailBasicDbResult
        ' ---------------------------------------------------------------
        ' DbGetMailBasic — 讀取 mail_withattachs WHERE folder_path=? 的所有行
        ' 回傳 Nothing 表示 DB 中無此資料夾的郵件記錄
        ' ---------------------------------------------------------------
        Dbg("開始")
        If _db Is Nothing Then Return Nothing

        Try
            Dim result As New MailBasicDbResult()
            Using cmd As New SqliteCommand(
                "SELECT entry_id,subject,msg_size,received_time,sender_name,attach_count,item_count_snap" &
                " FROM mail_withattachs WHERE folder_path=@p", _db)

                cmd.Parameters.AddWithValue("@p", folderPath)
                Using reader = cmd.ExecuteReader()
                    While reader.Read() ' item_count_snap 整個 folder 共用同一值，每行都一樣，讀最後一次即可
                        result.Snap = If(reader.IsDBNull(6), -1, reader.GetInt32(6))
                        Dim mail As New MailItemInfo()
                        mail.EntryID = reader.GetString(0)
                        mail.Subject = If(reader.IsDBNull(1), "", reader.GetString(1))
                        mail.Size = reader.GetInt64(2)
                        Dim dtStr = If(reader.IsDBNull(3), "", reader.GetString(3))
                        DateTime.TryParse(dtStr, mail.ReceivedTime)
                        mail.SenderName = If(reader.IsDBNull(4), "", reader.GetString(4))
                        mail.AttachCount = reader.GetInt32(5)
                        result.Mails.Add(mail)
                    End While
                End Using
            End Using
            Return If(result.Mails.Count > 0, result, Nothing)

        Catch ex As System.Exception
            Dbg("錯誤", $"{folderPath}: {ex.Message}")
        Finally : Dbg("結束")
        End Try
        Return Nothing

    End Function
    Friend Function DbGetAttachFilenames(entryId As String) As List(Of String)
        ' ---------------------------------------------------------------
        ' DbGetAttachFilenames — 讀取 attach_filenames WHERE entry_id=? 的一行
        ' 回傳 Nothing 表示 DB 中無此 EntryID
        ' ---------------------------------------------------------------
        Dbg("開始")
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
            Dbg("錯誤", $"{entryId}: {ex.Message}")
        Finally : Dbg("結束")
        End Try
        Return Nothing

    End Function
    Friend Function DbGetYearCountsForFolder(folderPath As String) As ConcurrentDictionary(Of Integer, Integer)
        ' ---------------------------------------------------------------
        ' DbGetYearCountsForFolder — 讀取 year_counts WHERE folder_path=? 的所有行
        ' 供 ComputeYearCountsAsync 在記憶體 miss 時先查 DB，避免 COM 呼叫
        ' 回傳 Nothing 表示 DB 中無此資料夾的年份記錄
        ' ---------------------------------------------------------------
        If _db Is Nothing Then Return Nothing
        Try
            Dim result As New ConcurrentDictionary(Of Integer, Integer)()
            Using cmd As New SqliteCommand("SELECT year,count FROM year_counts WHERE folder_path=@p", _db)
                cmd.Parameters.AddWithValue("@p", folderPath)
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        result(reader.GetInt32(0)) = reader.GetInt32(1)
                    End While
                End Using
            End Using
            Return If(result.Count > 0, result, Nothing)
        Catch ex As System.Exception
            Dbg("DbGetYearCountsForFolder ERROR", $"{folderPath}: {ex.Message}")
        End Try
        Return Nothing

    End Function


End Class
