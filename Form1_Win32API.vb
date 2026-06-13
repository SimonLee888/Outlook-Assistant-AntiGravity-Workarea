Imports System.Collections.Concurrent
Imports System.Globalization
Imports Microsoft.Office.Interop

Partial Class Form1

#Region "■ 01 全域宣告"
#Region "  ├ Win32 API 宣告"
    ' ── 函數宣告 ────────────────────────────────────────────────────────────────
    ' 統一使用 DllImport (取代舊式 Declare Function)
    ' 2026-03-23 整理: 移除重複的 SendMessage Declare 版本，補齊 FindWindow / FindWindowEx
    <Runtime.InteropServices.DllImport("user32.dll", CharSet:=Runtime.InteropServices.CharSet.Auto)>
    Private Shared Function FindWindow(
        ByVal lpClassName As String,
        ByVal lpWindowName As String) As IntPtr
    End Function
    <Runtime.InteropServices.DllImport("user32.dll", CharSet:=Runtime.InteropServices.CharSet.Auto)>
    Private Shared Function FindWindowEx(
        ByVal hWndParent As IntPtr,
        ByVal hWndChildAfter As IntPtr,
        ByVal lpszClass As String,
        ByVal lpszWindow As String) As IntPtr
    End Function
    <Runtime.InteropServices.DllImport("user32.dll", CharSet:=Runtime.InteropServices.CharSet.Auto)>
    Private Shared Function SendMessage(
        ByVal hWnd As IntPtr,
        ByVal msg As Integer,
        ByVal wParam As IntPtr,
        ByVal lParam As IntPtr) As IntPtr
    End Function

    <Runtime.InteropServices.DllImport("user32.dll")>
    Private Shared Function ShowWindow(
        ByVal hWnd As IntPtr,
        ByVal nCmdShow As Integer) As Boolean
    End Function
    <Runtime.InteropServices.DllImport("user32.dll")>
    Private Shared Function LockWindowUpdate(
        ByVal hWnd As IntPtr) As Boolean
    End Function
    <Runtime.InteropServices.DllImport("user32.dll")>
    Private Shared Function RedrawWindow(
        ByVal hWnd As IntPtr,
        ByVal lprcUpdate As IntPtr,
        ByVal hrgnUpdate As IntPtr,
        ByVal flags As UInteger) As Boolean
    End Function

    <Runtime.InteropServices.DllImport("user32.dll")>
    Private Shared Function PostMessage(
        ByVal hWnd As IntPtr,
        ByVal msg As Integer,
        ByVal wParam As IntPtr,
        ByVal lParam As IntPtr) As Boolean
    End Function
    <Runtime.InteropServices.DllImport("user32.dll")>
    Private Shared Function SetWindowPos(
        ByVal hWnd As IntPtr,
        ByVal hWndInsertAfter As IntPtr,
        ByVal x As Integer,
        ByVal y As Integer,
        ByVal cx As Integer,
        ByVal cy As Integer,
        ByVal uFlags As Integer) As Boolean
    End Function

    ' === 用來強制移除 SplitContainer 焦點框 ===
    <Runtime.InteropServices.DllImport("user32.dll")>
    Private Shared Function GetWindowLong(
        hWnd As IntPtr,
        nIndex As Integer) As Integer
    End Function
    <Runtime.InteropServices.DllImport("user32.dll")>
    Private Shared Function SetWindowLong(
        hWnd As IntPtr,
        nIndex As Integer,
        dwNewLong As Integer) As Integer
    End Function

    ' === 2026/4/19, 用來改變windows計時器精度 ===
    <Runtime.InteropServices.DllImport("winmm.dll", EntryPoint:="timeBeginPeriod", SetLastError:=True)>
    Private Shared Function TimeBeginPeriod(ByVal uPeriod As Integer) As Integer
    End Function
    <Runtime.InteropServices.DllImport("winmm.dll", EntryPoint:="timeEndPeriod", SetLastError:=True)>
    Private Shared Function TimeEndPeriod(ByVal uPeriod As Integer) As Integer
    End Function

    ' ── 常數 ───
    Private Const WM_LBUTTONDOWN As Integer = &H201
    Private Const WM_LBUTTONUP As Integer = &H202
    Private Const SW_HIDE As Integer = 0

    ' TreeView 雙緩衝
    Private Const TV_FIRST As Integer = &H1100
    Private Const TVM_SETEXTENDEDSTYLE As Integer = TV_FIRST + 44
    Private Const TVS_EX_DOUBLEBUFFER As Integer = &H4

    ' ListView 雙緩衝
    Private Const LVM_SETEXTENDEDLISTVIEWSTYLE As Integer = &H1036
    Private Const LVS_EX_DOUBLEBUFFER As Integer = &H10000
    Private Const SWP_NOZORDER As Integer = &H4                 ' debugForm resize用
    Private Const SWP_NOACTIVATE As Integer = &H10              ' debugForm resize用
    Private Const SWP_NOREDRAW As Integer = &H8                 ' debugForm resize 時閃爍, 改手動redraw
    Private Const RDW_ALLCHILDREN As Integer = &H80             ' debugForm resize 時閃爍, 改手動redraw
    Private Const RDW_INVALIDATE As Integer = &H1               ' debugForm resize 時閃爍, 改手動redraw
    Private Const RDW_UPDATENOW As Integer = &H100              ' debugForm resize 時閃爍, 改手動redraw
    Private Const RDW_ERASE As Integer = &H4                    ' 2026/03/28 by Gemini: 補上缺失定義

    ' ↓ 新增 (2026-03-20) ListView1 進入資料夾用
    Private Const WM_SETREDRAW As Integer = &HB                 ' 2026/3/26 by Gemini
    Private Const WM_SIZE As Integer = &H5                      ' 視窗尺寸變更訊息, 2026/5/7 by Claude
    Private Const SIZE_MAXIMIZED As Integer = 2                 ' WM_SIZE wParam: 最大化
    Private Const SIZE_RESTORED As Integer = 0                  ' WM_SIZE wParam: 還原
#End Region
#End Region

#Region "■ 99 舊版備用 (勿刪)"

    ' 定義排序方式的列舉
    'Private _lv3SortOrder As SortOrder = SortOrder.Ascending    ' 設置初始排序方式為升序
    'Private _lv3LastSortColumn As Integer = -1                  ' 儲存上一次點選的列索引
    Friend Class ListViewItemComparer ' 用於比較 ListView 項目並依Column 進行排序
        Implements IComparer
        Private ReadOnly columnIndex As Integer
        Private ReadOnly order As SortOrder
        Public Sub New(columnIndex As Integer, order As SortOrder)
            Me.columnIndex = columnIndex
            Me.order = order
        End Sub
        Public Function Compare(x As Object, y As Object) As Integer Implements IComparer.Compare
            Dim itemX As ListViewItem = DirectCast(x, ListViewItem)
            Dim itemY As ListViewItem = DirectCast(y, ListViewItem)
            Dim compareResult As Integer
            Select Case columnIndex
                Case 1  ' 郵件大小: 從 Tag 讀 Long，O(1)，不解析字串
                    Dim sizeX As Long = GetSizeFromTag(itemX)
                    Dim sizeY As Long = GetSizeFromTag(itemY)
                    compareResult = sizeX.CompareTo(sizeY)
                Case 2  ' 日期
                    Dim dateX As DateTime, dateY As DateTime
                    If DateTime.TryParse(itemX.SubItems(2).Text, dateX) AndAlso
                       DateTime.TryParse(itemY.SubItems(2).Text, dateY) Then
                        compareResult = dateX.CompareTo(dateY)
                    Else
                        compareResult = 0
                    End If
                Case 4  ' 附件個數直接 TryParse (數量小，解析快)
                    Dim countX As Integer = GetAttachCountFromTag(itemX)
                    Dim countY As Integer = GetAttachCountFromTag(itemY)
                    compareResult = countX.CompareTo(countY)
                Case Else  ' 文字欄位 (Subject、SenderName、EntryID)
                    compareResult = String.Compare(itemX.SubItems(columnIndex).Text,
                                                   itemY.SubItems(columnIndex).Text,
                                                   StringComparison.CurrentCultureIgnoreCase)
            End Select
            Return If(order = SortOrder.Ascending, compareResult, -compareResult)

        End Function
        Private Shared Function GetSizeFromTag(item As ListViewItem) As Long
            ' Tag 存的是 Long (Phase1) 或 Long() (Phase2)
            If TypeOf item.Tag Is Long Then Return CLng(item.Tag)
            If TypeOf item.Tag Is Long() Then Return DirectCast(item.Tag, Long())(0)
            Dim v As Long   ' Fallback: 萬一 Tag 沒設，解析字串
            Long.TryParse(item.SubItems(1).Text, NumberStyles.AllowThousands, Nothing, v)
            Return v

        End Function
        Private Shared Function GetAttachCountFromTag(item As ListViewItem) As Integer
            If TypeOf item.Tag Is Long() Then Return CInt(DirectCast(item.Tag, Long())(1))
            Dim v As Integer ' ">0" 或普通數字字串
            If Integer.TryParse(item.SubItems(4).Text, v) Then Return v
            Return 0  ' ">0" 的情況視為 1

        End Function
    End Class

    Private Async Function GetTotalFolderCountAsync(folder As Outlook.Folder) As Task(Of Integer)
        _dbg("開始", folder.Name)
        Dim value As Integer
        Dim fPath As String = folder.FolderPath
        If _cacheFolderCountAll.TryGetValue(fPath, value) Then Return value     ' 檢查快取中是否已存在值, 若有則直接返回
        Dim totalSubCount As Integer = GetFolderCountL3(folder, fPath:=fPath)   ' 初始值為點選資料夾的子資料夾數量
        ' 5/21測試記錄: 這裡使用ConcurrentBag跟使用results.sum應該要比較快, 但不知為何實測結果都比GetTotalFolderCount_Old()還慢了5%, 這個函數先保留不清除
        ' 5/21最後決定: 二個函數快慢互有變化, 但GetTotalFolderCountAsync()的穩定性較好, 比New()的標準差來得小, 所以決定使用這個
        ' 使用 Parallel.ForEach 進行多線程遞迴計算subfolder數量
        Dim countingBag As New ConcurrentBag(Of Task(Of Integer))()             ' 使用 ConcurrentBag 來安全地收集每個子資料夾的數量
        Parallel.ForEach(folder.Folders.Cast(Of Outlook.Folder)(),
                         Sub(subFolder As Outlook.Folder)
                             'countingBag.Add(GetTotalFolderCountAsync(subFolder))
                             'countingBag.Add(GetFolderCountAllL3(subFolder))
                         End Sub)
        Dim results = Await Task.WhenAll(countingBag)   ' 等待所有平行出去收集的數量都確定回來了
        totalSubCount += results.Sum()                  ' 再將回傳的各個子資料夾的數量加總
        _cacheFolderCountAll.TryAdd(fPath, totalSubCount)
        ' ✅ 2026-03-16 移除多餘的 Try/Catch: ConcurrentDictionary.TryAdd 本身不拋例外 (只回傳 True/False)
        ' 原本是從 .Add() 時代留下的防護，改 TryAdd 後應一併移除
        Return totalSubCount

    End Function
    Private Async Function GetFolderSizeLegacy(folder As Outlook.Folder) As Task(Of Long)
        ' ==============================================================
        ' === GetFolderSizeLegacy — 修正版 (移除 Task.Run 包 COM) ===
        ' ==============================================================
        '
        ' 原版問題: Task.Run(Function() folder.Items.Cast(Of Object)().Sum(Function(st) st.Size))
        '          在 thread pool 執行緒上操作 Outlook COM 物件，違反 STA 規定, 在特定情況 (COM interop 敏感時機) 會造成 crash 或傳回錯誤結果
        '
        ' 修正做法: GetTable + PR_MESSAGE_SIZE 在 UI 執行緒循序讀取
        '           GetTable 回傳 MAPI binary table (低層讀取)
        '          一次只讀一個 Row，每個 Row 用後立即 ReleaseComObject，避免 RCW 累積
        '          每 100 筆 Yield 一次讓 UI 保持回應
        '          速度接近原版 LINQ (實測差距在誤差範圍內) ，但 STA 安全
        '
        ' 此函數仍為 Lazy (不主動觸發) :
        '   由 ListView1_ColumnClick 或右鍵選單「Show This Folder Size」觸發
        '   結果存入 folderSizeCache，BuildLv1Item 下次組裝時自動顯示
        ' ==============================================================
        _dbg("開始", folder.Name)
        Dim value As Long   ' 快取命中直接回傳
        If _cacheFolderSize.TryGetValue(folder, value) Then Return value
        '' 已知有問題的資料夾走舊路徑 (不明 COM 例外物件，GetTable 也可能出問題)
        'Dim exceptList As String() = {"Inbox_2000~2018", "Facebook"}
        'If exceptList.Contains(folder.Name) Then Return GetFolderSizeOld(folder)
        Dim table As Outlook.Table = Nothing
        Try
            ' GetTable + PR_MESSAGE_SIZE (0x0E080003) :
            ' PR_MESSAGE_SIZE_EXTENDED (0x0E080014, PT_I8) — PST 本地端的內建彙總屬性
            Const PR_MESSAGE_SIZE As String = "http://schemas.microsoft.com/mapi/proptag/0x0E080003"
            Const PR_SIZE_EXTENDED As String = "http://schemas.microsoft.com/mapi/proptag/0x0E080014"
            ' 只讀 Size 欄，不載入其他 MAPI 屬性，減少記憶體與 COM 開銷
            table = folder.GetTable()
            table.Columns.RemoveAll()
            table.Columns.Add(PR_SIZE_EXTENDED)
            Dim totalSize As Long = 0
            Dim rowCount As Integer = 0
            Do While Not table.EndOfTable
                Dim row As Outlook.Row = table.GetNextRow()
                totalSize += SafeGet(Of Long)(row, PR_SIZE_EXTENDED, 0L)
                TryMarshalRelease(row)
                rowCount += 1
                If rowCount Mod 100 = 0 Then Await Task.Yield()  ' 每 100 筆統計就讓 UI 回應一次
            Loop
            _cacheFolderSize.TryAdd(folder, totalSize)
            Return totalSize
        Catch ex As OverflowException
            _dbg("Error: GetFolderSizeLegacy overflow", folder.Name)
            Return -1
        Catch ex As System.Exception
            _dbg("Error: GetFolderSizeLegacy", folder.Name & " - " & ex.Message)
            Return -1
        Finally
            TryMarshalRelease(table)
        End Try

    End Function
    Private Function GetFolderSizeOld(folder As Outlook.Folder) As Long
        _dbg("開始", folder.Name)
        Dim totalSize As Long = 0
        Dim folderItems As Outlook.Items = Nothing
        Try
            folderItems = folder.Items          ' ✅ 先取出 Items 物件，才能在 Finally 釋放
            For Each item As Object In folderItems
                Try
                    Dim mailItem As Outlook.MailItem = DirectCast(item, Outlook.MailItem)
                    If mailItem IsNot Nothing Then
                        totalSize += mailItem.Size
                        'tasks.Add(Task.Run(Async Function() ' 使用非同步 IO 操作來取得郵件大小
                        '                       'Await mailItem.PropertyAccessor.GetPropertyAsync("http://schemas.microsoft.com/mapi/proptag/0x0E080014")
                        '                       Interlocked.Add(sizeAdder, mailItem.Size)
                        '                   End Function))
                    End If
                Catch
                End Try
            Next
            'Await Task.WhenAll(tasks) ' 等待所有非同步操作完成
        Finally
            TryMarshalRelease(folderItems)  ' ✅ Items 集合釋放
        End Try
        Return totalSize

    End Function
    Private Function GetMailCountRecursiveLegacy(folder As Outlook.Folder) As Integer
        _dbg("開始", folder.Name)
        Dim value As Integer
        If _cacheMailCountAll.TryGetValue(folder, value) Then Return value ' 檢查快取中是否已存在值, 若有則直接返回
        ' 改成先用 Parallel.ForEach 遍歷子文件夾並且並行處理
        Dim totalMailCount As Integer = 0
        Dim countingBag As New ConcurrentBag(Of Integer)()
        Try
            ' 5/21記錄: 模仿GetFolderSizeLegacy那一句超快速的LINQ, 但測試結果沒有現在這個快, 所以決定保留這個
            ' 2026/3/20, 重寫了底層GetMailCountAll() 但是不知為何效能還是比不過現在下面這個遞迴版本?? (todo: 暫時先保留)
            ' 原因: 原版遞迴只走一遍 COM 資料夾樹，新版走了兩遍COM:
            ' 第一遍: GetSubtreeToList()    → BFS 遍歷，存取每個 folder.Folders
            ' 第二遍: For Each allFolders   → GetMailCountL3() 再讀每個資料夾一次
            ' 2026/3/22, 導入Redemption, 應該可以刪掉這裡了? 還是讓Redemption 變成on-demand, 需要才啟動?
            'Parallel.ForEach(folder.Folders.Cast(Of Outlook.Folder),' 取得子資料夾的郵件數量並添加到 ConcurrentBag 中
            '                 Sub(subFolder As Outlook.Folder)
            '                     countingBag.Add(GetMailCountRecursive(subFolder))
            '                 End Sub)
            'totalMailCount = countingBag.Sum() ' 累加所有子資料夾的郵件數量
            ''' 最後再獲取選取文件夾自身的郵件數量 (改用MAPI table 的PR_CONTENT_COUNT屬性來getmailcount)
            ''Const PR_CONTENT_COUNT As String = "http://schemas.microsoft.com/mapi/proptag/0x36020003"
            ''totalMailCount += folder.PropertyAccessor.GetProperty(PR_CONTENT_COUNT)
            totalMailCount += GetMailCountL3(folder)  ' 單一目錄的mail count改成重寫的統一底層函數, 2026/3/20
            _cacheMailCountAll.TryAdd(folder, totalMailCount) ' 第一次計算後就存入快取
        Catch
        End Try
        Return totalMailCount

    End Function

    Private Function CalculateSimilarity(strA As String, strB As String) As Double
        ' by Gemini, 2026/04/04: Issue 1 移除 _dbg (Tab5 高頻呼叫，N封×2個函數=2N行輸出) 
        ' 計算編輯距離
        Dim editDistance As Integer = LevenshteinDistance(strA, strB)

        ' 將編輯距離歸一化為範圍在 0 到 1 之間的值
        Dim maxLength As Integer = Math.Max(strA.Length, strB.Length)
        Dim similarity As Double = 1 - CDbl(editDistance) / maxLength
        Return similarity

    End Function
    Private Function LevenshteinDistance(strA As String, strB As String) As Integer
        ' by Gemini, 2026/04/04: Issue 1 移除 _dbg (Tab5 高頻呼叫，同上) 
        ' 計算 Levenshtein 編輯距離的輔助函數
        Dim lenA As Integer = strA.Length
        Dim lenB As Integer = strB.Length
        Dim distance(lenA, lenB) As Integer
        For i As Integer = 0 To lenA : distance(i, 0) = i : Next
        For j As Integer = 0 To lenB : distance(0, j) = j : Next
        For j As Integer = 1 To lenB
            For i As Integer = 1 To lenA
                '' 改前 (5行)
                'If strA(i - 1) = strB(j - 1) Then
                '    distance(i, j) = distance(i - 1, j - 1)
                'Else
                '    distance(i, j) = Math.Min(Math.Min(distance(i - 1, j) + 1,
                '                                       distance(i, j - 1) + 1), distance(i - 1, j - 1) + 1)
                'End If

                ' 改後 (1行)
                distance(i, j) = If(strA(i - 1) = strB(j - 1),
                distance(i - 1, j - 1), Math.Min(Math.Min(distance(i - 1, j) + 1,
                                                          distance(i, j - 1) + 1), distance(i - 1, j - 1) + 1))
            Next
        Next
        Return distance(lenA, lenB)
    End Function
    Private Async Sub CacheSnifferAsync(cToken As System.Threading.CancellationToken)
        ' === CacheSniffer — 背景快取預讀系統 (B4) ===
        ' ===============================================================================
        ' 職責: 程式啟動後在背景靜默預讀 Tab1 / Tab2 / Tab3 ，快取後讓使用者點選時直接從記憶體讀取，不再等待 COM 查詢。
        '
        ' 設計原則:
        '   - 廣度優先 (BFS) : 淺層資料夾優先預讀，使用者最常點選的位置最先就緒
        '   - 固定 1 秒間隔: 每完成一個資料夾的三項快取，固定等 1 秒再繼續，讓 Outlook 有充足空閒時間回應使用者互動
        '   - COM 全在 UI 執行緒 (STA) : 所有 Await 都不切執行緒，不需要 Task.Run
        '   - CancellationToken: FormClosing 時呼叫 _cacheSnifferCts.Cancel()，確保程式關閉後不留殘餘 COM 呼叫
        '   - 快取命中就跳過: 若使用者已先點選觸發過快取，CacheSniffer 直接略過不重做
        '   - 停用方式: 把 Form1_Load 末尾的 CacheSnifferAsync(...) 那行加上 ' 即可，其餘程式碼完全不受影響
        '
        ' 預讀順序 (每個資料夾) :
        '   1. Tab1: mailCountCache + folderCountCache (GetMailCountAllL3 / GetTotalFolderCountAsync)
        '   2. Tab2: yearCountsCache (GetYearCountsForFolderAsync)
        '   3. Tab3: _cacheAttachMailList (CheckTab3CacheOrRescan)
        '
        ' 2026-03-16 B4 新增，由 PrewarmAllCachesAsync 重構整合，改名為 CacheSniffer
        '       只要偵測到正在進行 AfterSelect 或是正在跑複雜統計，就自動閉嘴等閒下來再繼續
        ' ===============================================================================

        'If _pstStoreList Is Nothing OrElse _pstStoreList.Count = 0 Then Return
        Await Task.Delay(10000, cToken)      ' 等待 10 秒: 確保 Form1_Load 完全結束、UI 呈現完畢，再開始佔用 Outlook COM

        'Try
        '    _dbg("開始", "預讀快取")
        '    ' ── BFS 初始化: 把所有 PST 的第一層子資料夾加進佇列 ─────────
        '    ' 不從 root 本身開始，因為 root ("個人資料夾") 通常不含郵件，
        '    ' 直接從第一層子資料夾 (收件匣、寄件匣…) 開始
        '    Dim queue As New Queue(Of Outlook.Folder)
        '    For Each store As Outlook.Store In _pstStoreList
        '        If cToken.IsCancellationRequested Then Return
        '        For Each subFolder As Outlook.Folder In GetSortedSubFolders(store.GetRootFolder())
        '            queue.Enqueue(subFolder)
        '        Next
        '    Next
        '    ' ── BFS 主迴圈 ───────────────────────────────────────────────
        '    ' 每次取出一個資料夾，依序預讀 Tab1 / Tab2 / Tab3 的快取，
        '    ' 完成後把它的直屬子資料夾再放入佇列 (廣度優先，淺層先完成)
        '    Dim processed As Integer = 0
        '    While queue.Count > 0
        '        If cToken.IsCancellationRequested Then Return
        '        Dim folder As Outlook.Folder = queue.Dequeue()
        '        processed += 1
        '        ' ── Tab1: mailCountCache + folderCountCache ───────────────
        '        ' GetMailCountAllL3 和 GetTotalFolderCountAsync 內部各自寫入自己的快取
        '        ' 已命中的快取直接跳過，不重複呼叫 COM
        '        Try
        '            Await GetMailCountAllAsync(folder)
        '            Await GetFolderCountAllAsync(folder)
        '        Catch ex As System.Exception
        '            _dbg("CacheSniffer Tab1 Error: ", folder.Name & " - " & ex.Message)
        '        End Try
        '        If cToken.IsCancellationRequested Then Return
        '        ' ── Tab2: yearCountsCache ─────────────────────────────────
        '        ' GetYearCountsForFolderAsync 內部有快取命中判斷，已快取直接回傳
        '        Try
        '            Dim key As String = folder.FolderPath
        '            If Not _cacheYearCounts.ContainsKey(key) Then Await GetYearCountsForFolderL3(folder)
        '        Catch ex As System.Exception
        '            _dbg("CacheSniffer Tab2 Error: ", folder.Name & " - " & ex.Message)
        '        End Try
        '        If cToken.IsCancellationRequested Then Return
        '        ' ── Tab3: _cacheAttachMailList ────────────────────────────────
        '        ' CheckTab3CacheOrRescan 內部有 Items.Count 失效判斷
        '        Try
        '            Await CheckTab3CacheOrRescan(folder, Nothing)
        '        Catch ex As System.Exception
        '            _dbg("CacheSniffer Tab3 Error: ", folder.Name & " - " & ex.Message)
        '        End Try
        '        If cToken.IsCancellationRequested Then Return
        '        ' ── 固定 1 秒間隔: 讓 Outlook 保持回應能力 ───────────────
        '        _dbg($"CacheSniffer: [{processed}] {folder.Name} 完成，等 1 秒")
        '        Await Task.Delay(1000, cToken)
        '        Await Task.Yield()
        '        ' ── 把直屬子資料夾加入佇列 (廣度優先) ────────────────────
        '        ' GetSortedSubFolders 有 folderTreeCache，不重打 COM
        '        Try
        '            For Each subFolder As Outlook.Folder In GetSortedSubFolders(folder)
        '                queue.Enqueue(subFolder)
        '            Next
        '        Catch ex As System.Exception
        '            _dbg("錯誤", folder.Name & " - " & ex.Message)
        '        End Try
        '    End While
        '    _dbg("結束", $"預讀完成 | 總計: {processed} 個資料夾")
        'Catch ex As System.Threading.Tasks.TaskCanceledException
        '    _dbg("CacheSniffer: 已取消 (FormClosing) ")
        'Catch ex As System.Exception
        '    _dbg("錯誤", ex.Message)
        'Finally
        '    _dbg("結束")
        'End Try

    End Sub

    Private Async Function ForceTvRefresh_old(tv As SimTree) As Task
        ' ── SimTree F5 強制刷新 ──────────────────────────────────────────────
        ' 職責: 不讀任何快取，重新從 Outlook COM 讀取整棵資料夾樹並更新 _cacheFolderTree
        '       ① 記錄目前展開路徑 + 選取路徑
        '       ② 清 _cacheFolderTree (確保 LoadSubFolderToTreeView 重讀 COM)
        '       ③ Nodes.Clear + LoadStoreToTreeView (重建 root 層)
        '       ④ 逐層 node.Expand() 重建已展開路徑 (觸發 LoadSubFolderToTreeView)
        '       ⑤ 還原選取，透過 FireAfterSelect 觸發正常 AfterSelect 流程更新 ListView
        '
        ' 2026/05/13 by Claude Sonnet 4.6
        ' 2026/05/17 by Simon/Claude: ⑤ 改回 FireAfterSelect，解決 ListView 未更新的問題
        '   原本直接呼叫 ComputeTab1FolderStats + RenderLv1 的方式繞過了 SimTree 標準流程，
        '   導致 AfterSelect 沒有被觸發，ListView 顯示內容不對應選取的資料夾。
        ' 2026/05/25 by Simon/Claude: 再度重構使用呼叫simTree內部方法
        ' ─────────────────────────────────────────────────────────────────────
        _dbg("開始", tv.Name)
        If _pstStoreList Is Nothing OrElse _pstStoreList.Count = 0 Then Return

        ' ① 記錄展開路徑與選取路徑
        Dim expandedPaths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        CollectExpandedPaths(tv.Nodes, expandedPaths)

        Dim selectedPaths As New List(Of String)(32)
        For Each node As TreeNode In tv.SelectedNodes
            Dim f As Outlook.Folder = TryCast(node.Tag, Outlook.Folder)
            If f IsNot Nothing Then selectedPaths.Add(SafeGetPath(f))
        Next

        ProgressBar1.Text = $"F5: 重整 {tv.Name}..." : ProgressBar2.Text = ""
        _isUserBusy = True : Cursor = Cursors.WaitCursor

        Try
            ' ② 清 _cacheFolderTree（確保 GetSortedSubFolders 重讀 COM，不用舊快取）
            _cacheFolderTree.Clear()

            ' ③ 重建 root 層
            tv.ClearSelectedNodes()
            tv.Nodes.Clear()
            LoadStoreToTreeView(_pstStoreList, tv)

            ' ④ 逐條路徑重展開（node.Expand() 觸發 BeforeExpand → LoadSubFolderToTreeView → 新鮮 COM）
            For Each path In expandedPaths.OrderBy(Function(p) p.Length)   ' 由淺到深確保父節點先展開
                ReExpandNodeByPath(tv, path)
                Dim unused = TimeBeginPeriod(1)
                Await Task.Delay(1)
                Dim unused1 = TimeEndPeriod(1)
            Next

            ' ⑤ 還原選取
            Dim firstNode As TreeNode = Nothing
            For Each path In selectedPaths
                ' by Gemini 3.5 Flash, 2026/05/21: 改用 tv.GetNodeIn 高效尋路引擎，取代舊有的暴力遞迴 FindNodeByPath
                Dim found As TreeNode = tv.GetNode(path, searchOnlyExpanded:=True)
                If found IsNot Nothing Then
                    tv.AddSelectedNode(found)
                    If firstNode Is Nothing Then firstNode = found
                End If
            Next

            If firstNode IsNot Nothing Then
                firstNode.EnsureVisible()
                ' 2026/05/17 by Simon/Claude:
                tv.FireAfterSelect(firstNode)
                ' 改回 FireAfterSelect，讓 SimTree1_AfterSelect 自行處理統計與 RenderLv1，這才是 SimTree 的標準觸發流程。
                ' 原本直接呼叫 ComputeTab1FolderStats 的方式導致 ListView1 未被正確更新。
            Else
                GotoDefaultInbox(tv)   ' 找不到舊選取時退回預設 Inbox
            End If

            ProgressBar1.Text = $"F5: {tv.Name} 重整完成" : ProgressBar2.Text = ""

        Catch ex As System.Exception
            _dbg("錯誤", ex.Message) : ProgressBar1.Text = $"F5 {tv.Name} 失敗: " & ex.Message
        Finally
            Cursor = Cursors.Default : _isUserBusy = False : _dbg("結束", tv.Name)
        End Try
    End Function
    Private Sub CollectExpandedPaths(nodes As TreeNodeCollection, paths As HashSet(Of String))
        ''' <summary>遞迴收集已展開節點的 FolderPath，供 F5 刷新前記錄狀態用</summary>
        For Each n As TreeNode In nodes
            Dim f As Outlook.Folder = TryCast(n.Tag, Outlook.Folder)
            If f Is Nothing Then Continue For   ' 跳過 ":::" 佔位節點
            If n.IsExpanded Then
                paths.Add(SafeGetPath(f))
                CollectExpandedPaths(n.Nodes, paths)
            End If
        Next
    End Sub
    Private Sub ReExpandNodeByPath(tv As SimTree, fullPath As String)
        ' by Gemini 3.5 Flash, 2026/05/21: 重構以使用底層高效的尋路與展開機制，取代舊的手動逐層循環暴力比對，以防佔用執行緒
        Dim found As TreeNode = Nothing
        If tv.TryGetNode(fullPath, found, searchOnlyExpanded:=False, expandAlongTheWay:=True) Then
            If found IsNot Nothing AndAlso Not found.IsExpanded AndAlso found.Nodes.Count > 0 Then found.Expand()
        End If
    End Sub
    Private Async Function RenewCacheToDB_old(includeSize As Boolean) As Task
        '    ' ---------------------------------------------------------------
        '    ' RenewCacheToDB — 完整更新 DB 快取 (對應 Setting 頁 RenewCache 按鈕) 
        '    '
        '    ' 與 SaveCachesToDB 的差異：
        '    '   SaveCache  = 把目前記憶體快取照單全收寫入 DB (不更新過期的值) 
        '    '   RenewCache = 先用 COM 比對 snapshot → 只對有變動的資料夾重新計算 → 寫入 DB
        '    '
        '    ' 流程：
        '    '   Phase 1. BFS 掃出所有 live folders (COM，~1ms/資料夾) 
        '    '   Phase 2. 每個 folder 讀 GetLiveFolderSnapL3 vs DB snapshot → 找 dirty folders
        '    '   Phase 3. 對每個 dirty folder 重新計算：
        '    '              mc/fc (快，~1ms) 
        '    '              year_counts (GetTable + GetArray，~10-50ms/資料夾) 
        '    '              month_counts (清記憶體， Phase5 清 DB， 展開時 lazy 重算) 
        '    '              attach_maillist (GetTable 三路比對，~5ms/資料夾) 
        '    '              folder_size (選擇性，依 includeSize，GetTable 遍歷，~10-30s/大資料夾) 
        '    '              清除 mca/fca/fsa 聚合快取 (讓下次點選重算) 
        '    '              清除此 folder 的 month_counts 記憶體快取 (不重算，展開年份時 lazy) 
        '    '   Phase 4. 清除所有 dirty folders 的 ancestor 聚合快取
        '    '   Phase 5. 批次 DELETE dirty folders 的 month_counts DB rows (不是孤兒，不靠 Cleanup) 
        '    '   Phase 6. CleanupOrphanFolderPath → SaveCachesToDB
        '    '
        '    ' 不更新項目 (設計邊界) ：
        '    '   attach_filenames — 最耗時，留給使用者搜尋附件時 lazy 觸發
        '    '   month_counts     — 清記憶體 + 清 DB，展開年份時 lazy 重算
        '    ' 2026/04/09 by Claude
        '    ' ---------------------------------------------------------------
        '    ' 2026/04/16 by Simon/Claude: 加入 cToken (OkayNowYouHaveToken)，取代 _cancelRequested + GoTo Cancelled 模式
        '    '   Phase1 改用 Dictionary(Of String, Outlook.Folder) liveDict，每個資料夾只讀一次 FolderPath COM 屬性，
        '    '   Phase2/3/4 迭代 dict 的 Key/Value，完全省去重複的 folder.FolderPath COM 呼叫（~500 資料夾省 ~250ms）
        '    '   Phase2/3 節流改用 SmartThrottle(sw, cToken, ThrottleFreq.Low)，取代 Mod N + Task.Delay(1)
        '    '   GetYearCountsForFolderL3 / GetFolderSizeL3 補入 cToken:=cToken
        '    ' ---------------------------------------------------------------

        '    Dim cToken As Threading.CancellationToken = OkayNowYouHaveToken()  ' ✅ 取得新 Token，同時取消上一次未完成的操作
        '    _dbg("開始", $"includeSize={includeSize}")
        '    If _db Is Nothing Then _dbg("", "DB 未初始化") : Return

        '    Dim sw As New Diagnostics.Stopwatch : sw.Start()
        '    Try
        '        ' ── Phase 1: BFS 掃出所有 live folders ──
        '        ' 2026/04/16: 改用 Dictionary(Of String, Outlook.Folder) liveDict
        '        '   key = FolderPath (一次 COM 呼叫)，value = Folder 物件
        '        '   後續 Phase2/3/4 直接用 kvp.Key 作 fPath，不再打 folder.FolderPath
        '        ProgressBar1.Text = "RenewCache Phase1: 掃描資料夾清單..." : Cursor = Cursors.WaitCursor
        '        Await Task.Yield

        '        Dim liveDict As New Dictionary(Of String, Outlook.Folder)()
        '        For Each store As Outlook.Store In _pstStoreList
        '            Dim root As Outlook.Folder = TryCast(store.GetRootFolder(), Outlook.Folder)
        '            If root Is Nothing Then Continue For

        '            ' 2026/04/24 by Gemini 3.0 flash: 使用 SafeGetPath 確保 root 取得安全
        '            Dim rootPath As String = SafeGetPath(root)
        '            If String.IsNullOrEmpty(rootPath) Then Continue For

        '            ' 2026/04/16 by Gemini: GetSubtreeToList 現在直接回傳 Tuple (Folder, FolderPath)
        '            ' 直接將計算好的路徑存入 liveDict，完成 0 COM Call 的清單建立
        '            For Each item In Await GetSubtreeToList(root, includeSubF:=True, cToken:=cToken)
        '                If Not liveDict.ContainsKey(item.fPath) Then liveDict.Add(item.fPath, item.folder)
        '            Next
        '        Next
        '        Dim livePaths As New HashSet(Of String)(liveDict.Keys)  ' 供 Phase6 CleanupOrphan 使用
        '        _dbg("Phase1 完成", $"{liveDict.Count} 個 live folder")

        '        ' ── Phase 2: 比對 snapshot → 找出 dirty folders ──
        '        ' 2026/04/16: 迭代 liveDict，kvp.Key 直接當 fPath，省去 folder.FolderPath COM 呼叫
        '        '   節流改用 SmartThrottle(swThrottle2, cToken, ThrottleFreq.Low)，取代 Mod 100 + Task.Delay(1)
        '        ProgressBar1.Text = $"RenewCache Phase2: 比對 snapshot (共 {liveDict.Count} 個) ..."
        '        Dim dirtyDict As New Dictionary(Of String, Outlook.Folder)()
        '        ' by Claude Sonnet 4.6, 2026/04/25: 區分兩種「dirty」語意
        '        '   isNewFolder = True  → DB 從未記錄（清空後首次，或真正新資料夾）
        '        '                         Phase 3 只算 mc/fc/year_counts，跳過 attach_maillist 重掃
        '        '                         attach_maillist 交由使用者搜尋附件時 lazy 觸發
        '        '   isNewFolder = False → snapshot 不符（真正有信件增減）
        '        '                         Phase 3 完整重算包含 attach_maillist（三路比對）
        '        ' 這樣清空快取後執行 RenewCache，不會因為所有資料夾都「看起來像新的」而偷跑全量 GetTable 掃描，產生 2 萬筆非預期的 attach_maillist 內容。
        '        Dim dirtyNewFolderSet As New HashSet(Of String)()  ' 記錄 isNewFolder=True 的路徑
        '        Dim processed As Integer = 0
        '        Dim swThrottle2 As New Stopwatch : swThrottle2.Start()
        '        For Each kvp In liveDict
        '            cToken.ThrowIfCancellationRequested()  ' 2026/04/16: 取代 _cancelRequested + GoTo Cancelled
        '            Dim fPath As String = kvp.Key : Dim folder As Outlook.Folder = kvp.Value
        '            Dim liveSnap As Integer = GetLiveFolderSnapL3(folder, fPath:=fPath)   ' ~0.5ms，PropertyAccessor 單次呼叫 by Gemini 3.0 flash, 2026/04/16
        '            Dim row = DbGetFolderStats(fPath)

        '            ' dirty 條件：DB 無此路徑 (新資料夾) OR snapshot 不一致 (有信件增減)
        '            Dim isNewFolder As Boolean = (row Is Nothing)
        '            If isNewFolder OrElse row.snap <> liveSnap Then
        '                dirtyDict.Add(fPath, folder)
        '                If isNewFolder Then
        '                    dirtyNewFolderSet.Add(fPath)  ' 全新資料夾，Phase 3 跳過 attach_maillist

        '                    ' by Gemini 3.0 flash, 2026/04/24: 新資料夾確保 ID 被快取，Phase 6 寫入時需要 entry_id
        '                    _cacheFolderIDs.TryAdd(fPath, (folder.EntryID, folder.StoreID, IsMailFolder(folder, fPath), TextHasChineseChar(ExtractFolderName(fPath))))

        '                    ' 使父資料夾的樹狀快取失效，確保刷新 UI 後能顯示新成員
        '                    Dim parentPath As String = GetParentPath(fPath)
        '                    If Not String.IsNullOrEmpty(parentPath) Then
        '                        ' 清除父路徑的所有顯示模式快取 (|True 與 |False)
        '                        _cacheFolderTree.TryRemove(parentPath & "|True", Nothing)
        '                        _cacheFolderTree.TryRemove(parentPath & "|False", Nothing)
        '                    End If
        '                End If
        '            End If

        '            processed += 1
        '            ' 2026/04/16 by Gemini 3.0 flash: 改用 ThrottleFreq.Low + SmartThrottle 與 onThrottled 委派
        '            Await SmartThrottle(swThrottle2, cToken:=cToken, ThrottleFreq.Low,
        '                                      Sub() ProgressBar1.Text = $"RenewCache Phase2: {processed}/{liveDict.Count}，dirty={dirtyDict.Count} (新={dirtyNewFolderSet.Count})...")
        '        Next
        '        _dbg("Phase2 完成", $"dirty={dirtyDict.Count}/{liveDict.Count} (其中全新資料夾={dirtyNewFolderSet.Count})")

        '        ' ── Phase 3: 對每個 dirty folder 重新計算 ──
        '        ' 2026/04/16: 迭代 dirtyDict，省去 folder.FolderPath COM 呼叫
        '        '   GetYearCountsForFolderL3 / GetFolderSizeL3 補入 cToken:=cToken
        '        '   節流改用 SmartThrottle(swThrottle3, cToken, ThrottleFreq.Low)，取代 Mod 10 + Task.Delay(1)
        '        ProgressBar1.Text = $"RenewCache Phase3: 更新 {dirtyDict.Count} 個 dirty 資料夾..." : Await Task.Delay(1, cToken)
        '        processed = 0
        '        Dim swThrottle3 As New Stopwatch : swThrottle3.Start()
        '        For Each kvp In dirtyDict
        '            cToken.ThrowIfCancellationRequested()  ' 2026/04/16: 取代 _cancelRequested + GoTo Cancelled
        '            Dim fPath As String = kvp.Key : Dim folder As Outlook.Folder = kvp.Value

        '            ' mc / fc — 快，~1ms，直接覆蓋記憶體快取
        '            _cacheMailCount(fPath) = GetMailCountL3(folder, fPath:=fPath)
        '            _cacheFolderCount(fPath) = GetFolderCountL3(folder, fPath:=fPath)

        '            ' year_counts — 清記憶體強制 L3 重算，結果回寫快取
        '            _cacheYearCounts.TryRemove(fPath, Nothing)
        '            _cacheYearCounts(fPath) = Await GetYearCountsForFolderL3(folder, fPath:=fPath, cToken:=cToken)  ' 2026/04/16: 補 cToken

        '            ' month_counts — 只清記憶體 (Phase5 再清 DB)，展開年份時 lazy 重算
        '            For Each mk In _cacheMonthCounts.Keys.Where(Function(k) k.StartsWith(fPath & "_")).ToList()
        '                _cacheMonthCounts.TryRemove(mk, Nothing)
        '            Next

        '            ' attach_maillist — 三路比對，更新記憶體快取 (不碰 attach_filenames)
        '            ' by Claude Sonnet 4.6, 2026/04/25: 只對「真正 dirty」（snapshot 不符）的資料夾才重掃附件
        '            '   全新資料夾（DB 從未記錄）跳過，避免清空快取後 RenewCache 偷跑全量 GetTable 掃描
        '            '   全新資料夾的 attach_maillist 在使用者執行 Tab3 附件搜尋時透過 lazy load 建立
        '            If Not dirtyNewFolderSet.Contains(fPath) Then
        '                Await RenewAttachMailList(folder, fPath:=fPath)
        '            End If

        '            ' folder_size — 選擇性 (GetTable 遍歷 PR_MESSAGE_SIZE，大資料夾需 10~30s)
        '            If includeSize Then _cacheFolderSize(fPath) = Await GetFolderSizeL3(folder, fPath:=fPath, cToken:=cToken)  ' 2026/04/16: 補 cToken

        '            ' 聚合快取清除 — 讓 parent 在下次點選時重新 BFS 加總
        '            ' by Claude Sonnet 4.6, 2026/04/25: 同時清除 |True 和 |False 兩個模式的鍵值
        '            '   因應未來 _showAllFolders 分支鍵值架構，確保兩個模式的過期聚合都被清掉
        '            _cacheMailCountAll.TryRemove(fPath & "|True", Nothing)
        '            _cacheMailCountAll.TryRemove(fPath & "|False", Nothing)
        '            _cacheMailCountAll.TryRemove(fPath, Nothing)    ' 兼容舊鍵值（無分支時寫入的）
        '            _cacheFolderCountAll.TryRemove(fPath & "|True", Nothing)
        '            _cacheFolderCountAll.TryRemove(fPath & "|False", Nothing)
        '            _cacheFolderCountAll.TryRemove(fPath, Nothing)  ' 同上
        '            _cacheFolderSizeAll.TryRemove(fPath, Nothing)

        '            processed += 1
        '            ' 2026/04/16 by Gemini 3.0 flash: 改用 ThrottleFreq.Low + SmartThrottle 與 onThrottled 委派
        '            Await SmartThrottle(swThrottle3, cToken:=cToken, ThrottleFreq.Low,
        '                                      Sub() ProgressBar1.Text = $"RenewCache Phase3: {processed}/{dirtyDict.Count} 個處理中...")
        '        Next
        '        _dbg("Phase3 完成", $"{processed} 個 dirty folder 重新計算完畢")

        '        ' ── Phase 4: 清除 dirty folders 的 ancestor 聚合快取 ──
        '        ' 任何 dirty leaf 都讓所有 ancestor 的 mca/fca/fsa 失效
        '        ' 2026/04/16: 改迭代 liveDict.Keys，直接用 key 作 fPath，省去 fs.FolderPath COM 呼叫
        '        ' by Gemini 3.0 flash, 2026/04/24: 優化為「精確打擊」模式，改用 GetAncestors 直接清除，效能從 O(N*D) 降至 O(D*L)
        '        If dirtyDict.Count > 0 Then
        '            For Each dp In dirtyDict.Keys
        '                For Each ancestor In GetAncestors(dp)
        '                    ' by Claude Sonnet 4.6, 2026/04/25: 同時清除 |True / |False 兩個模式鍵值及舊式鍵值
        '                    _cacheMailCountAll.TryRemove(ancestor & "|True", Nothing)
        '                    _cacheMailCountAll.TryRemove(ancestor & "|False", Nothing)
        '                    _cacheMailCountAll.TryRemove(ancestor, Nothing)
        '                    _cacheFolderCountAll.TryRemove(ancestor & "|True", Nothing)
        '                    _cacheFolderCountAll.TryRemove(ancestor & "|False", Nothing)
        '                    _cacheFolderCountAll.TryRemove(ancestor, Nothing)
        '                    _cacheFolderSizeAll.TryRemove(ancestor, Nothing)
        '                Next
        '            Next
        '            _dbg("Phase4 完成", $"已針對 {dirtyDict.Count} 個異動路徑精確清除祖先快取 (含 |True/|False 模式鍵值)")
        '        End If

        '        ' ── Phase 5: 批次 DELETE dirty folders 的 month_counts DB rows ──
        '        ' 注意: CleanupOrphan 只刪「不再存在的路徑」，不刪「仍存在但 dirty」的路徑
        '        '       所以 dirty folder 的舊 month rows 必須在這裡主動清除
        '        Dim dirtyPaths As New HashSet(Of String)(dirtyDict.Keys)  ' 供 Phase 5 使用
        '        If dirtyPaths.Count > 0 AndAlso _db IsNot Nothing Then
        '            Await Task.Run(Sub()
        '                               Using txn = _db.BeginTransaction()
        '                                   Try
        '                                       Using cmd As New SqliteCommand("DELETE FROM month_counts WHERE folder_path=@p", _db, txn)
        '                                           cmd.Parameters.Add("@p", SqliteType.Text)
        '                                           For Each dp In dirtyPaths
        '                                               cmd.Parameters("@p").Value = dp : cmd.ExecuteNonQuery()
        '                                           Next
        '                                       End Using
        '                                       txn.Commit()
        '                                   Catch : txn.Rollback() : Throw
        '                                   End Try
        '                               End Using
        '                           End Sub)
        '            _dbg("Phase5 完成", $"已清 {dirtyPaths.Count} 個 dirty folder 的 month_counts DB rows")
        '        End If

        '        ' ── Phase 6: 孤兒清除 + 批次寫入 ──
        '        ProgressBar1.Text = "RenewCache Phase6: 清孤兒 + 寫入 DB..." : Await Task.Delay(1, cToken)
        '        Await CleanupOrphanPath(livePaths)
        '        Await SaveCachesToDB()    ' 內部會顯示 SaveCache 的進度訊息

        '        sw.Stop()
        '        Dim st = GetDBSummary()
        '        ProgressBar1.Text = $"RenewCache 完成 ✔ dirty:{dirtyDict.Count}/{liveDict.Count} 個 / 耗時:{sw.Elapsed.TotalSeconds:0.0}s — DB:{st.fc}/{st.mb}/{st.at}/{st.yc}/{st.mc} 筆"
        '        _dbg("完成", $"dirty={dirtyDict.Count}, total={liveDict.Count}, elapsed={sw.Elapsed.TotalSeconds:0.0}s")

        '    Catch ex As OperationCanceledException
        '        ' 2026/04/16: cToken 取消時 (ESC)，取代原本的 _cancelRequested + GoTo Cancelled 模式
        '        ProgressBar1.Text = "RenewCache 由使用者中斷"
        '        _dbg("中斷", "使用者按 ESC")
        '    Catch ex As System.Exception
        '        ProgressBar1.Text = $"RenewCache 失敗: {ex.Message}"
        '        _dbg("錯誤", ex.Message)
        '    Finally
        '        Cursor = Cursors.Default
        '        _dbg("結束")
        '    End Try
    End Function


#End Region


End Class
