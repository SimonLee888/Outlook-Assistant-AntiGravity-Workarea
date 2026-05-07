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
    Private Const GWL_STYLE As Integer = -16
    Private Const WS_TABSTOP As Integer = &H10000
    Private Const SW_HIDE As Integer = 0
    Private Const WM_COMMAND As Integer = &H111
    Private Const WM_LBUTTONDOWN As Integer = &H201
    Private Const WM_LBUTTONUP As Integer = &H202
    Private Const BM_CLICK As Integer = &HF5

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
    Private Const RDW_FRAME As Integer = &H400                  ' 2026/03/28 by Gemini: 補上缺失定義

    ' ↓ 新增 (2026-03-20) ListView1 進入資料夾用
    Private Const TVM_SELECTITEM As Integer = &H110B            ' = &H1100 + 11
    Private Const TVGN_CARET As Integer = &H9                   ' SendMessage 選取 Treeview 游標節點
    Private Const LVM_SETITEMCOUNT As Integer = &H1000 + 47     ' = &H102F '
    Private Const WM_SETREDRAW As Integer = &HB                 ' 2026/3/26 by Gemini
    Private Const WM_SIZE As Integer = &H5                      ' 視窗尺寸變更訊息, 2026/5/7 by Claude
    Private Const SIZE_MAXIMIZED As Integer = 2                 ' WM_SIZE wParam: 最大化
    Private Const SIZE_RESTORED As Integer = 0                  ' WM_SIZE wParam: 還原
#End Region
#End Region

#Region "■ 99 舊版備用 (勿刪)"

    ' 定義排序方式的列舉
    'Private lv3SortOrder As SortOrder = SortOrder.Ascending     ' 設置初始排序方式為升序
    'Private lv3LastSortColumn As Integer = -1                     ' 儲存上一次點選的列索引
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
        If _cacheFolderCountAll.TryGetValue(fPath, value) Then Return value ' 檢查快取中是否已存在值, 若有則直接返回
        Dim totalSubCount As Integer = GetFolderCountL3(folder, fPath:=fPath)           ' 初始值為點選資料夾的子資料夾數量
        ' 5/21測試記錄: 這裡使用ConcurrentBag跟使用results.sum應該要比較快, 但不知為何實測結果都比GetTotalFolderCount_Old()還慢了5%, 這個函數先保留不清除
        ' 5/21最後決定: 二個函數快慢互有變化, 但GetTotalFolderCountAsync()的穩定性較好, 比New()的標準差來得小, 所以決定使用這個
        ' 使用 Parallel.ForEach 進行多線程遞迴計算subfolder數量
        Dim countingBag As New ConcurrentBag(Of Task(Of Integer))()     ' 使用 ConcurrentBag 來安全地收集每個子資料夾的數量
        Parallel.ForEach(folder.Folders.Cast(Of Outlook.Folder)(),
                         Sub(subFolder As Outlook.Folder)
                             'countingBag.Add(GetTotalFolderCountAsync(subFolder))
                             countingBag.Add(GetFolderCountAllL3(subFolder))
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
#End Region


End Class
