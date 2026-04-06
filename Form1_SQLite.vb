Imports Microsoft.Data.Sqlite
Imports System.IO

' ==============================================================
' === SQLite 快取管理模組 (by AntiGravity, 2026/04/06) ===
' 目的: 將內存中的 ConcurrentDictionary 持久化到磁碟，避免重開後重新計數
' 技術: 使用 Microsoft.Data.Sqlite 進行高效、跨平台的資料庫操作
' 更新: 2026/04/06 - 強化 Debug 訊息，分項顯示各快取處理筆數
' ==============================================================
<System.ComponentModel.DesignerCategory("")>
Partial Class Form1

    Private _dbPath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "cache.db")

    ''' <summary>
    ''' 初始化資料庫，若不存在則建立 cache.db 並建立資料表
    ''' </summary>
    Private Sub InitDatabase()
        Dim connectionString As String = "Data Source=" + _dbPath

        Try
            Using connection As New SqliteConnection(connectionString)
                connection.Open()

                Dim createTableSql As String = "
                    CREATE TABLE IF NOT EXISTS FolderCache (
                        FolderPath TEXT PRIMARY KEY,
                        MailCount INTEGER,
                        MailCountAll INTEGER,
                        FolderCount INTEGER,
                        FolderCountAll INTEGER,
                        FolderSize INTEGER,
                        FolderSizeAll INTEGER,
                        IsMailFolder INTEGER,
                        UpdateTime TEXT
                    );"

                Using command As New SqliteCommand(createTableSql, connection)
                    command.ExecuteNonQuery()
                End Using
            End Using
            Dbg("SQLite 初始化成功", _dbPath)
        Catch ex As Exception
            Dbg("SQLite 初始化失敗", ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' 將目前的快取 Dictionary 合併存入 SQLite
    ''' </summary>
    Private Async Function SaveCachesToSQLiteAsync() As Task
        Dbg("開始存入 SQLite")
        Dim sw As New Stopwatch() : sw.Start()

        Dim allKeys = _cacheMailCount.Keys.Union(_cacheMailCountAll.Keys).
                      Union(_cacheFolderCount.Keys).Union(_cacheFolderCountAll.Keys).
                      Union(_cacheFolderSize.Keys).Union(_cacheFolderSizeAll.Keys).
                      Union(_cacheIsMailFolder.Keys).Distinct().ToList()

        If allKeys.Count = 0 Then
            Dbg("存入取消", "快取中沒有任何資料") : Return
        End If

        ' 統計各快取存入筆數
        Dim sMc, sMcA, sFc, sFcA, sSz, sSzA, sIsM As Integer

        Try
            Using connection As New SqliteConnection("Data Source=" + _dbPath)
                Await connection.OpenAsync()

                Using transaction = connection.BeginTransaction()
                    Dim upsertSql As String = "
                        INSERT OR REPLACE INTO FolderCache 
                        (FolderPath, MailCount, MailCountAll, FolderCount, FolderCountAll, FolderSize, FolderSizeAll, IsMailFolder, UpdateTime)
                        VALUES (@path, @mc, @mcAll, @fc, @fcAll, @fSize, @fSizeAll, @isMail, @time);"

                    Using command As New SqliteCommand(upsertSql, connection, transaction)
                        command.Parameters.Add("@path", SqliteType.Text)
                        command.Parameters.Add("@mc", SqliteType.Integer)
                        command.Parameters.Add("@mcAll", SqliteType.Integer)
                        command.Parameters.Add("@fc", SqliteType.Integer)
                        command.Parameters.Add("@fcAll", SqliteType.Integer)
                        command.Parameters.Add("@fSize", SqliteType.Integer)
                        command.Parameters.Add("@fSizeAll", SqliteType.Integer)
                        command.Parameters.Add("@isMail", SqliteType.Integer)
                        command.Parameters.Add("@time", SqliteType.Text)

                        Dim nowStr As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

                        For Each pathKey In allKeys
                            Dim mc, mcA, fc, fcA As Integer
                            Dim fSz, fSzA As Long
                            Dim isM As Boolean

                            Dim hMc = _cacheMailCount.TryGetValue(pathKey, mc) : If hMc Then sMc += 1
                            Dim hMcA = _cacheMailCountAll.TryGetValue(pathKey, mcA) : If hMcA Then sMcA += 1
                            Dim hFc = _cacheFolderCount.TryGetValue(pathKey, fc) : If hFc Then sFc += 1
                            Dim hFcA = _cacheFolderCountAll.TryGetValue(pathKey, fcA) : If hFcA Then sFcA += 1
                            Dim hSz = _cacheFolderSize.TryGetValue(pathKey, fSz) : If hSz Then sSz += 1
                            Dim hSzA = _cacheFolderSizeAll.TryGetValue(pathKey, fSzA) : If hSzA Then sSzA += 1
                            Dim hIsM = _cacheIsMailFolder.TryGetValue(pathKey, isM) : If hIsM Then sIsM += 1

                            command.Parameters("@path").Value = pathKey
                            command.Parameters("@mc").Value = If(Not hMc, -1, mc)
                            command.Parameters("@mcAll").Value = If(Not hMcA, -1, mcA)
                            command.Parameters("@fc").Value = If(Not hFc, -1, fc)
                            command.Parameters("@fcAll").Value = If(Not hFcA, -1, fcA)
                            command.Parameters("@fSize").Value = If(Not hSz, -1, fSz)
                            command.Parameters("@fSizeAll").Value = If(Not hSzA, -1, fSzA)
                            command.Parameters("@isMail").Value = If(Not hIsM, -1, If(isM, 1, 0))
                            command.Parameters("@time").Value = nowStr

                            Await command.ExecuteNonQueryAsync()
                        Next
                    End Using
                    transaction.Commit()
                End Using
            End Using
            sw.Stop()

            Dim msg = "成功存入 " + allKeys.Count.ToString() + " 筆資料夾資訊到 SQLite (耗時 " + sw.ElapsedMilliseconds.ToString() + "ms)
" +
                      "  -> 郵件數(自身/子樹): " + sMc.ToString() + " / " + sMcA.ToString() + " 筆
" +
                      "  -> 目錄數(自身/子樹): " + sFc.ToString() + " / " + sFcA.ToString() + " 筆
" +
                      "  -> 大小(自身/子樹): " + sSz.ToString() + " / " + sSzA.ToString() + " 筆
" +
                      "  -> 類型標記: " + sIsM.ToString() + " 筆"
            Dbg("結束存入 SQLite", msg)
            ProgressBar2.Text = "快取已存入資料庫 (" + allKeys.Count.ToString() + " 筆)，耗時 " + sw.ElapsedMilliseconds.ToString() + "ms"
        Catch ex As Exception
            Dbg("存入 SQLite 發生異常", ex.Message)
            ProgressBar2.Text = "快取存入失敗！請查看 Debug Log"
        End Try
    End Function

    ''' <summary>
    ''' 從 SQLite 讀取快取並填回 Dictionary
    ''' </summary>
    Private Async Function LoadCachesFromSQLiteAsync() As Task
        Dbg("開始從 SQLite 讀取")
        Dim sw As New Stopwatch() : sw.Start()

        If Not File.Exists(_dbPath) Then
            Dbg("讀取取消", "找不到資料庫檔案") : Return
        End If

        ' 統計各快取讀取筆數
        Dim totalRows, lMc, lMcA, lFc, lFcA, lSz, lSzA, lIsM As Integer

        Try
            Using connection As New SqliteConnection("Data Source=" + _dbPath)
                Await connection.OpenAsync()
                Await EnsureColumnsExist(connection)

                Dim selectSql As String = "SELECT * FROM FolderCache;"
                Using command As New SqliteCommand(selectSql, connection)
                    Using reader = Await command.ExecuteReaderAsync()
                        While Await reader.ReadAsync()
                            totalRows += 1
                            Dim pathKey = reader.GetString(0)

                            Dim mc = reader.GetInt32(1) : If mc <> -1 Then _cacheMailCount(pathKey) = mc : lMc += 1
                            Dim mcA = reader.GetInt32(2) : If mcA <> -1 Then _cacheMailCountAll(pathKey) = mcA : lMcA += 1
                            Dim fc = reader.GetInt32(3) : If fc <> -1 Then _cacheFolderCount(pathKey) = fc : lFc += 1
                            Dim fcA = reader.GetInt32(4) : If fcA <> -1 Then _cacheFolderCountAll(pathKey) = fcA : lFcA += 1
                            Dim fSz = reader.GetInt64(5) : If fSz <> -1 Then _cacheFolderSize(pathKey) = fSz : lSz += 1
                            Dim fSzA = reader.GetInt64(6) : If fSzA <> -1 Then _cacheFolderSizeAll(pathKey) = fSzA : lSzA += 1
                            Dim isM = reader.GetInt32(7) : If isM <> -1 Then _cacheIsMailFolder(pathKey) = (isM = 1) : lIsM += 1
                        End While
                    End Using
                End Using
            End Using
            sw.Stop()

            Dim msg = "成功從 SQLite 恢復 " + totalRows.ToString() + " 筆資料夾快取 (耗時 " + sw.ElapsedMilliseconds.ToString() + "ms)
" +
                      "  -> 郵件數(自身/子樹): " + lMc.ToString() + " / " + lMcA.ToString() + " 筆
" +
                      "  -> 目錄數(自身/子樹): " + lFc.ToString() + " / " + lFcA.ToString() + " 筆
" +
                      "  -> 大小(自身/子樹): " + lSz.ToString() + " / " + lSzA.ToString() + " 筆
" +
                      "  -> 類型標記: " + lIsM.ToString() + " 筆"
            Dbg("結束從 SQLite 讀取", msg)
            ProgressBar2.Text = "已從資料庫恢復 " + totalRows.ToString() + " 筆快取，耗時 " + sw.ElapsedMilliseconds.ToString() + "ms"
        Catch ex As Exception
            Dbg("讀取 SQLite 發生異常", ex.Message)
            ProgressBar2.Text = "快取讀取失敗！請查看 Debug Log"
        End Try
    End Function

    ''' <summary>
    ''' 輔助函數：確保資料表結構包含最新欄位 (升級舊版資料庫)
    ''' </summary>
    Private Async Function EnsureColumnsExist(conn As SqliteConnection) As Task
        Dim columnsToAdd As New List(Of String) From {"FolderSize", "FolderSizeAll", "IsMailFolder"}
        For Each col In columnsToAdd
            Dim checkSql = "PRAGMA table_info(FolderCache);"
            Dim exists = False
            Using cmd = New SqliteCommand(checkSql, conn)
                Using r = Await cmd.ExecuteReaderAsync()
                    While Await r.ReadAsync()
                        If r.GetString(1).Equals(col, StringComparison.OrdinalIgnoreCase) Then
                            exists = True : Exit While
                        End If
                    End While
                End Using
            End Using
            If Not exists Then
                Dbg("升級資料庫", "新增欄位: " + col)
                Using alterCmd = New SqliteCommand("ALTER TABLE FolderCache ADD COLUMN " + col + " INTEGER DEFAULT -1;", conn)
                    Await alterCmd.ExecuteNonQueryAsync()
                End Using
            End If
        Next
    End Function

End Class