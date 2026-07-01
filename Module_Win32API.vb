Imports System.Collections.Concurrent
Imports System.Threading
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Outlook

Partial Class Form1

#Region "■ 01 全域宣告"
#Region "  ├ Win32 API 宣告"
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
    Private Const SWP_NOZORDER As Integer = &H4             ' debugForm resize用
    Private Const SWP_NOACTIVATE As Integer = &H10          ' debugForm resize用
    Private Const SWP_NOSIZE As Integer = &H1               ' 2026/06/19 by Simon/Claude: 只搬位置不改尺寸，拖曳跟隨用
    Private Const SWP_NOREDRAW As Integer = &H8             ' debugForm resize 時閃爍, 改手動redraw
    Private Const RDW_ALLCHILDREN As Integer = &H80         ' debugForm resize 時閃爍, 改手動redraw
    Private Const RDW_INVALIDATE As Integer = &H1           ' debugForm resize 時閃爍, 改手動redraw
    Private Const RDW_UPDATENOW As Integer = &H100          ' debugForm resize 時閃爍, 改手動redraw
    Private Const RDW_ERASE As Integer = &H4                ' 2026/03/28 by Gemini: 補上缺失定義

    ' ↓ 新增 (2026-03-20) ListView1 進入資料夾用
    Private Const WM_SETREDRAW As Integer = &HB             ' 2026/3/26 by Gemini
    Private Const WM_SIZE As Integer = &H5                  ' 視窗尺寸變更訊息, 2026/5/7 by Claude
    Private Const WM_ENTERSIZEMOVE As Integer = &H231       ' 2026/06/19 by Simon/Claude: 進入拖曳 size/move modal loop
    Private Const WM_EXITSIZEMOVE As Integer = &H232        ' 2026/06/19 by Simon/Claude: 離開拖曳 size/move modal loop
    Private Const SIZE_MAXIMIZED As Integer = 2             ' WM_SIZE wParam: 最大化
    Private Const SIZE_RESTORED As Integer = 0              ' WM_SIZE wParam: 還原
#End Region
#End Region

#Region "■ 99 舊版備用 (勿刪)"
    Private Async Function GetTotalFolderCountAsync(folder As Outlook.Folder) As Task(Of Integer)
        _dbg("開始", folder.Name)
        Dim value As Integer
        Dim fPath As String = folder.FolderPath
        If _cacheFolderCountAll.TryGetValue(fPath, value) Then Return value     ' 檢查快取中是否已存在值, 若有則直接返回
        Dim totalSubCount As Integer = GetFolderCountOOM(folder, fPath:=fPath)   ' 初始值為點選資料夾的子資料夾數量
        ' 5/21測試記錄: 這裡使用ConcurrentBag跟使用results.sum應該要比較快, 但不知為何實測結果都比GetTotalFolderCount_Old()還慢了5%, 這個函數先保留不清除
        ' 5/21最後決定: 二個函數快慢互有變化, 但GetTotalFolderCountAsync()的穩定性較好, 比New()的標準差來得小, 所以決定使用這個
        ' 使用 Parallel.ForEach 進行多線程遞迴計算subfolder數量
        Dim countingBag As New ConcurrentBag(Of Task(Of Integer))()             ' 使用 ConcurrentBag 來安全地收集每個子資料夾的數量
        Parallel.ForEach(folder.Folders.Cast(Of Outlook.Folder)(),
                         Sub(subFolder As Outlook.Folder)
                             'countingBag.Add(GetTotalFolderCountAsync(subFolder))
                             'countingBag.Add(GetFolderCountAllOOM(subFolder))
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
            ' 第一遍: GetSubtree()    → BFS 遍歷，存取每個 folder.Folders
            ' 第二遍: For Each allFolders   → GetMailCountOOM() 再讀每個資料夾一次
            ' 2026/3/22, 導入Redemption, 應該可以刪掉這裡了? 還是讓Redemption 變成on-demand, 需要才啟動?
            'Parallel.ForEach(folder.Folders.Cast(Of Outlook.Folder),' 取得子資料夾的郵件數量並添加到 ConcurrentBag 中
            '                 Sub(subFolder As Outlook.Folder)
            '                     countingBag.Add(GetMailCountRecursive(subFolder))
            '                 End Sub)
            'totalMailCount = countingBag.Sum() ' 累加所有子資料夾的郵件數量
            ''' 最後再獲取選取文件夾自身的郵件數量 (改用MAPI table 的PR_CONTENT_COUNT屬性來getmailcount)
            ''Const PR_CONTENT_COUNT As String = "http://schemas.microsoft.com/mapi/proptag/0x36020003"
            ''totalMailCount += folder.PropertyAccessor.GetProperty(PR_CONTENT_COUNT)
            totalMailCount += GetMailCountOOM(folder)  ' 單一目錄的mail count改成重寫的統一底層函數, 2026/3/20
            _cacheMailCountAll.TryAdd(folder, totalMailCount) ' 第一次計算後就存入快取
        Catch
        End Try
        Return totalMailCount

    End Function
    Private Function GetMailSizeL3(item As Object) As Long
        ' --------------------------------------------------------------
        ' GetMailSizeL3: 讀取單封郵件的大小 (bytes)，供 GetFolderSizeOOM fallback 路徑呼叫
        '
        ' Fallback 鏈:
        '   ⓪ Redemption : RDOMail.Size
        '                  free-threaded 安全，可在 Task.Run 內呼叫
        '                  繞過 Outlook Security Guard，不會彈出安全性警告
        '                  _rdo 未就緒時自動跳過此層
        '   ① MAPI : PR_MESSAGE_SIZE_EXTENDED (0x0E080014, PT_I8, 64-bit Long)
        '            避免 PR_MESSAGE_SIZE (PT_LONG, 32-bit) 在超大郵件時溢位
        '   ② MAPI : PR_MESSAGE_SIZE (0x0E080003, PT_LONG, 32-bit Integer)
        '            Fallback 到 32-bit 版本，CInt → CLng 安全轉型
        '   ③ OOM  : mail.Size
        '            最後手段，OOM 的 Size 屬性單位是 bytes，回傳 Integer，
        '            大郵件 (>2GB) 理論上會溢位，但實務上 Outlook 的 PST 限制在 50GB 總量，
        '            單封郵件超過 2GB 極不可能，此層可視為安全
        '
        ' 注意: 此函數接受 Object 型別參數，是因為 GetFolderSizeOOM 的 fallback 路徑
        '       用 Items.GetFirst/GetNext 取回的是 Object，省去呼叫端的 TryCast 成本
        '       若是 MailItem 就正常讀取，若是其他型別 (Contact、Appointment 等) 就回 0
        '
        ' 取代: GetFolderSizeOld 內的 mailItem.Size 直接呼叫 行3385 的同名 stub (完整替換)
        ' ---------------------------------------------------------------
        ' 2026/4/28 by simon: 目前此函數已成死碼, 沒有任何呼叫端與進入點
        ' 原始設計意圖:
        ' 	呼叫端 → GetFolderSizeAll → GetFolderSize → GetMailSizeL3 (L3)
        ' 後來使用GetTable.GetArray() 直接整個目錄的table一起讀出來在記憶體內運算
        ' 	(自己計算並直接寫入 _cacheFolderSize，完全繞過此L3層級函數)
        ' ---------------------------------------------------------------

        ' 非 MailItem 的項目 (Calendar、Contact 等) 直接略過，回 0
        If _iLikeNoisy Then _dbg("    ├ 開始")
        Dim mail As Outlook.MailItem = TryCast(item, Outlook.MailItem)
        If mail Is Nothing Then Return 0

        ' ⓪ Redemption: RDOMail.Size
        '   GetMessageFromID 的 StoreID 從 mail.Parent 取得，多一次 COM call 但避免跨 PST 找錯 item
        '   2026-03-22 新增
        If _rdo IsNot Nothing Then
            Dim rdoMail As Redemption.RDOMail = Nothing
            Try
                Dim parentFolder As Outlook.Folder = TryCast(mail.Parent, Outlook.Folder)
                Dim storeId As String = If(parentFolder?.StoreID, "")
                rdoMail = TryCast(_rdo.GetMessageFromID(mail.EntryID, storeId), Redemption.RDOMail)
                If rdoMail IsNot Nothing Then
                    Dim sz As Long = CLng(rdoMail.Size)
                    If _iLikeNoisy Then _dbg("    ├ ⓪ RDO 成功", $"size={sz}") ' 高頻率項目平時不輸出 Log
                    Return sz
                End If
            Catch ex As System.Exception
                If _iLikeNoisy Then _dbg("    ├ ⓪ RDO 失敗，走MAPI fallback", ex.Message) ' by Gemini, 2026/04/10: 調整縮排層級為 Level 2 (內部失敗路徑)
            Finally
                TryMarshalRelease(rdoMail)
            End Try
        End If

        ' ① MAPI: PR_MESSAGE_SIZE_EXTENDED (0x0E080014, PT_I8) — 64-bit，無溢位風險
        Try
            Const PR_SIZE_EXTENDED As String = "http://schemas.microsoft.com/mapi/proptag/0x0E080014"
            ' by Gemini, 2026/03/29: 移除 TypeOf 判斷，CLng() 可自動處理 Long/Integer 轉型，若屬性不存在或回傳 Nothing/DBNull，CLng 會拋例外進入 Catch
            Return CLng(mail.PropertyAccessor.GetProperty(PR_SIZE_EXTENDED))
        Catch ex As System.Exception
            If _iLikeNoisy Then _dbg("    ├ ① PR_MESSAGE_SIZE_EXTENDED失敗", ex.Message) ' by Gemini, 2026/04/10
        End Try

        ' ② MAPI: PR_MESSAGE_SIZE (0x0E080003, PT_LONG) — 32-bit，超大郵件理論上溢位
        Try
            Const PR_SIZE As String = "http://schemas.microsoft.com/mapi/proptag/0x0E080003"
            Return CLng(mail.PropertyAccessor.GetProperty(PR_SIZE))             ' by Gemini, 2026/03/29: 同上，移除 TypeOf 判斷
        Catch ex As System.Exception
            If _iLikeNoisy Then _dbg("    ├ ② PR_MESSAGE_SIZE失敗", ex.Message) ' by Gemini, 2026/04/10
        End Try

        ' ③ OOM: mail.Size (Integer，超大郵件理論上不準，但實務上 PST 內不會發生)
        Try
            Return CLng(mail.Size)
        Catch ex As System.Exception
            If _iLikeNoisy Then _dbg("    ├ ③ OOM mail.Size也失敗", ex.Message) ' by Gemini, 2026/04/10
        End Try
        Return -1

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
        '   1. Tab1: mailCountCache + folderCountCache (GetMailCountAllOOM / GetTotalFolderCountAsync)
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
        '        ' GetMailCountAllOOM 和 GetTotalFolderCountAsync 內部各自寫入自己的快取
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

        PgrsBar1.Text = $"F5: 重整 {tv.Name}..." : PgrsBar2.Text = ""
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

            PgrsBar1.Text = $"F5: {tv.Name} 重整完成" : PgrsBar2.Text = ""

        Catch ex As System.Exception
            _dbg("錯誤", ex.Message) : PgrsBar1.Text = $"F5 {tv.Name} 失敗: " & ex.Message
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

    Private Async Function RdoPreloadAttach_1(sourceList As List(Of MailItemInfo), progress As IProgress(Of ProgressReport), cToken As CancellationToken) As Task
        ' =================================================================
        ' by Gemini, 2026/04/05: Layer2.5 快取代理層 - 批次預熱附件檔名快取
        '   利用 Redemption (RDO) Free-Threaded 安全的特性，
        '   在進入 Layer2 迴圈前平行提早把附件檔名讀進 _cacheAttachFilename。
        '   完全不更改原有的迴圈運作邏輯，以預讀取的型態塞資料進快取來大幅壓縮等待時間。
        ' =================================================================
        If _rdo Is Nothing OrElse sourceList.Count = 0 Then Return

        _dbg("開始", $"RDO預載Parallel.ForEach {sourceList.Count} 筆")
        Dim swTotal As Stopwatch = Stopwatch.StartNew()     ' by Claude Sonnet 4.6, 2026/06/07
        Dim swThrottle As Stopwatch = Stopwatch.StartNew()  ' by Claude Sonnet 4.6, 2026/06/07
        Dim processed As Integer = 0
        Dim total As Integer = sourceList.Count

        ' 設定並發數：嘗試設為 CPU 核心數的 4 倍，壓榨 SSD 的 Queue Depth
        Dim maxConcurrency As Integer = Environment.ProcessorCount * 4

        Await Task.Run(Sub()
                           ' ✅ 2026/04/11 cToken 重構: CancellationToken 傳入 ParallelOptions，取消時 Parallel.ForEach 會拋 OperationCanceledException
                           Dim parallelOptions As New ParallelOptions With {.MaxDegreeOfParallelism = maxConcurrency, .CancellationToken = cToken}
                           Try
                               Parallel.ForEach(sourceList, parallelOptions,
                                                Sub(mail)
                                                    If Not _cacheAttachFilename.ContainsKey(mail.EntryID) Then
                                                        Dim rdoMsg As Redemption.RDOMail = Nothing
                                                        Try
                                                            rdoMsg = TryCast(_rdo.GetMessageFromID(mail.EntryID), Redemption.RDOMail)
                                                            If rdoMsg IsNot Nothing Then
                                                                Dim list As New List(Of String)(512)
                                                                For i As Integer = 1 To rdoMsg.Attachments.Count    ' COM 的 index 從 1 開始而不是0
                                                                    list.Add(rdoMsg.Attachments.Item(i).FileName)
                                                                Next
                                                                _cacheAttachFilename.TryAdd(mail.EntryID, list)
                                                            End If
                                                        Catch
                                                        Finally
                                                            If rdoMsg IsNot Nothing Then TryMarshalRelease(rdoMsg)
                                                        End Try
                                                    End If

                                                    Dim curProcessed As Integer = Interlocked.Increment(processed)
                                                    If swThrottle.ElapsedMilliseconds >= ThrottleFreq.Hii OrElse curProcessed = total Then
                                                        Dim eta = CalculateSpeedAndETA(total, curProcessed, swTotal.Elapsed.TotalSeconds)
                                                        progress?.Report(New ProgressReport With {.CurrentCount = curProcessed, .TotalCount = total,
                                                                                                  .Message = $"Phase 2 (RDO預載Parallel.ForEach): {curProcessed} / {total} ({eta.Speed:F0} 封/秒{eta.EtaString})"})
                                                        swThrottle.Restart()
                                                    End If
                                                End Sub)
                           Catch ex As OperationCanceledException
                               ' cToken 取消時 Parallel.ForEach 拋出，正常中斷，不需處理
                           End Try
                       End Sub, cToken)
        _dbg(" ├ 結束", $"RDO預載Parallel.ForEach完成，處理共 {processed} 筆") ' by Gemini, 2026/04/10: 調整縮排層級為 Level 1
    End Function
    Private Async Function RdoPreloadAttach_2(sourceList As List(Of MailItemInfo), progress As IProgress(Of ProgressReport), cToken As CancellationToken) As Task
        ' ==============================================================
        ' by AntiGravity, 2026/04/07: 實驗性質
        ' - 使用 Task.WhenAll + SemaphoreSlim，試圖推高 SSD I/O 並發度
        ' ==============================================================
        If _rdo Is Nothing OrElse sourceList.Count = 0 Then Return

        _dbg(" ├ 開始", $"RDO預載Task.WhenAll {sourceList.Count} 筆") ' by Gemini, 2026/04/10: 調整縮排層級為 Level 1
        Dim swThrottle As Stopwatch = Stopwatch.StartNew()  ' by Claude Sonnet 4.6, 2026/06/07
        Dim swTotal As Stopwatch = Stopwatch.StartNew()     ' by Claude Sonnet 4.6, 2026/06/07
        Dim processed As Integer = 0
        Dim total As Integer = sourceList.Count

        ' 設定並發數：嘗試設為 CPU 核心數的 4 倍，壓榨 SSD 的 Queue Depth
        Dim maxConcurrency As Integer = Environment.ProcessorCount * 4
        Dim throttler As New SemaphoreSlim(maxConcurrency)
        Dim tasks As New List(Of Task)(32)

        For Each m As MailItemInfo In sourceList
            Dim mail = m ' 在 lambda 中避免變數捕獲問題

            tasks.Add(Task.Run(Async Function()
                                   Await throttler.WaitAsync(cToken)   ' ✅ cToken 取消時直接拋 OperationCanceledException
                                   Try
                                       If Not _cacheAttachFilename.ContainsKey(mail.EntryID) Then
                                           Dim rdoMsg As Redemption.RDOMail = Nothing
                                           Try
                                               rdoMsg = TryCast(_rdo.GetMessageFromID(mail.EntryID), Redemption.RDOMail)
                                               If rdoMsg IsNot Nothing Then
                                                   Dim list As New List(Of String)(512)
                                                   For i As Integer = 1 To rdoMsg.Attachments.Count
                                                       list.Add(rdoMsg.Attachments.Item(i).FileName)
                                                   Next
                                                   _cacheAttachFilename.TryAdd(mail.EntryID, list)
                                               End If
                                           Catch
                                           Finally
                                               If rdoMsg IsNot Nothing Then TryMarshalRelease(rdoMsg)
                                           End Try
                                       End If

                                       Dim curProcessed As Integer = Interlocked.Increment(processed)
                                       If swThrottle.ElapsedMilliseconds >= ThrottleFreq.Hii OrElse curProcessed = total Then
                                           Dim eta = CalculateSpeedAndETA(total, curProcessed, swTotal.Elapsed.TotalSeconds)
                                           progress?.Report(New ProgressReport With {.CurrentCount = curProcessed, .TotalCount = total,
                                                                                     .Message = $"Phase 2 (RDO預載Task.WhenAll): {curProcessed} / {total} ({eta.Speed:F0} 封/秒{eta.EtaString})"})
                                           swThrottle.Restart()
                                       End If
                                   Finally
                                       throttler.Release()
                                   End Try
                               End Function, cToken))
        Next

        If tasks.Count > 0 Then Await Task.WhenAll(tasks)
        _dbg(" ├ 結束", $"RDO預載=Task.WhenAll完成，處理共 {processed} 筆") ' by Gemini, 2026/04/10
    End Function
    Private Async Function RdoPreloadAttach_3(sourceList As List(Of MailItemInfo), progress As IProgress(Of ProgressReport), cToken As CancellationToken) As Task
        ' =================================================================
        ' 2026/06/19 by Simon/Claude: Layer2.5 多PST獨立Session平行預載 (實驗版_3)
        '   獨立 RDOSession (自有 MAPI session) 才是 Redemption 真 free-threaded 的前提。
        '   ★ 實測結論: (a) 同一 PST 多 session 無加速(PST provider 對同檔序列化);
        '             (b) 獨立冷 session 每封成本約為 Outlook 熱 session 的 3 倍。
        '   故設計為「每個 PST 一條獨立 session、組內循序、組間平行」,
        '   加速僅來自「跨多個 PST 同時讀取」→ 適用情境是同時選取大量 PST 的整庫掃描,
        '   少數 PST 時請改用 _1/_2 (共用熱 session 反而較快)。
        '   store 物件開啟一次重複使用(store.GetMessageFromID), 避免每封信重新定位 store 的成本。
        ' =================================================================
        If _rdo Is Nothing OrElse sourceList.Count = 0 Then Return   ' _rdo 僅作「RDO 是否可用」的偵測旗標

        _dbg("開始", $"RDO預載Cross.PST {sourceList.Count} 筆")
        Dim swTotal As Stopwatch = Stopwatch.StartNew()
        Dim swThrottle As Stopwatch = Stopwatch.StartNew()
        Dim processed As Integer = 0
        Dim resolveFail As Integer = 0      ' 暫時探針: 解 EntryID 的失敗數, 確認=0 後可移除
        Dim total As Integer = sourceList.Count

        ' 依 PST(store 顯示名) 分組: 每組一條獨立 session, 組內循序, 組間平行才是真正加速來源
        Dim groups = sourceList.GroupBy(Function(m) GetStoreNameFromPath(m.FolderPath)).ToList()
        _dbg(" ├ 分組", $"涵蓋 {groups.Count} 個 PST → 開 {groups.Count} 條平行 session")

        Dim tasks As New List(Of Task)(groups.Count)
        For Each grp In groups
            Dim storeName As String = grp.Key
            Dim items = grp.ToList()
            tasks.Add(Task.Run(Sub()
                                   Dim sess As Redemption.RDOSession = Nothing
                                   Try
                                       sess = New Redemption.RDOSession()
                                       sess.Logon(_rdo.ProfileName, "", False, True)   ' (ProfileName, Pwd, ShowDialog, NewSession): 不沿用 Outlook session
                                       ' 取得該 PST 已開啟的 RDOStore 並重複使用(避免每封信重開 store 的高昂成本)
                                       Dim store As Redemption.RDOStore = Nothing
                                       For i As Integer = 1 To sess.Stores.Count
                                           If sess.Stores.Item(i).Name = storeName Then store = sess.Stores.Item(i) : Exit For
                                       Next
                                       If store Is Nothing Then
                                           _dbg(" ├ 略過", $"獨立 session 找不到 store [{storeName}]，跳過該組 {items.Count} 筆")
                                           Interlocked.Add(resolveFail, items.Count)
                                           Return
                                       End If

                                       For Each mail As MailItemInfo In items
                                           cToken.ThrowIfCancellationRequested()
                                           If Not _cacheAttachFilename.ContainsKey(mail.EntryID) Then
                                               Dim rdoMsg As Redemption.RDOMail = Nothing
                                               Try
                                                   rdoMsg = TryCast(store.GetMessageFromID(mail.EntryID), Redemption.RDOMail)   ' ★ 用已開啟的 store, 不每封重開
                                                   If rdoMsg IsNot Nothing Then
                                                       Dim list As New List(Of String)(512)
                                                       For i As Integer = 1 To rdoMsg.Attachments.Count    ' COM 的 index 從 1 開始而不是0
                                                           list.Add(rdoMsg.Attachments.Item(i).FileName)
                                                       Next
                                                       _cacheAttachFilename.TryAdd(mail.EntryID, list)
                                                   Else
                                                       Interlocked.Increment(resolveFail)
                                                   End If
                                               Catch
                                                   Interlocked.Increment(resolveFail)
                                               Finally
                                                   If rdoMsg IsNot Nothing Then TryMarshalRelease(rdoMsg)
                                               End Try
                                           End If

                                           Dim curProcessed As Integer = Interlocked.Increment(processed)
                                           If swThrottle.ElapsedMilliseconds >= ThrottleFreq.Hii OrElse curProcessed = total Then
                                               Dim eta = CalculateSpeedAndETA(total, curProcessed, swTotal.Elapsed.TotalSeconds)
                                               progress?.Report(New ProgressReport With {.CurrentCount = curProcessed, .TotalCount = total,
                                                                                         .Message = $"Phase 2 (RDO預載Cross.PST): {curProcessed} / {total} ({eta.Speed:F0} 封/秒{eta.EtaString})"})
                                               swThrottle.Restart()
                                           End If
                                       Next
                                   Catch ex As OperationCanceledException
                                       ' cToken 取消, 正常中斷
                                   Catch ex As System.Exception
                                       _dbg(" ├ 失敗", $"PST [{storeName}] 組例外: {ex.GetBaseException().Message}")
                                   Finally
                                       If sess IsNot Nothing Then
                                           Try : sess.Logoff() : Catch : End Try
                                           TryMarshalRelease(sess)
                                       End If
                                   End Try
                               End Sub, cToken))
        Next

        If tasks.Count > 0 Then Await Task.WhenAll(tasks)
        _dbg(" ├ 結束", $"RDO預載Cross.PST完成，處理共 {processed} 筆，resolve 失敗 {resolveFail} 筆")
    End Function
    Private Function GetSubtreeToListL3_Rdo(rootFolder As Redemption.RDOFolder, includeSubF As Boolean) As List(Of Redemption.RDOFolder)
        ' --------------------------------------------------------------
        ' 2026/3/24 by Gemini: GetSubtreeToListL3_Rdo
        ' 目的: 專門提供給 RDO 平行路徑使用，回傳 List(Of Redemption.RDOFolder)
        ' 說明: 因為 Redemption 是 free-threaded，可以用 Parallel.ForEach 安全平行展開子樹
        ' 2026/6/27 退役，全部轉由整合完成的GetSubtreeListRdo單一入口
        ' --------------------------------------------------------------
        'Dim rootName As String = rootFolder.Name
        '_dbg("    ├ 開始", rootName)
        'Dim sw As Stopwatch = Stopwatch.StartNew()  ' by Claude Sonnet 4.6, 2026/06/07

        'Dim resultBag As New ConcurrentBag(Of Redemption.RDOFolder)
        'resultBag.Add(rootFolder)
        'If Not includeSubF Then
        '    sw.Stop()
        '    _dbg("    ├ 結束", $"{rootName} (RDO-Single) | {sw.ElapsedMilliseconds}ms") ' by Gemini, 2026/04/10
        '    Return resultBag.ToList()
        'End If

        '' 使用兩層佇列作層級遍歷，每層用 Parallel.ForEach 探索
        'Dim currentLayer As New ConcurrentQueue(Of Redemption.RDOFolder)
        'currentLayer.Enqueue(rootFolder)
        'Do
        '    Dim layerList = currentLayer.ToList()
        '    If layerList.Count = 0 Then Exit Do

        '    ' 清空 queue 準備裝下一層的資料夾
        '    Do While currentLayer.TryDequeue(Nothing) : Loop

        '    ' 平行處理當前層的資料夾，將它們的子資料夾加進 queue 與結果中
        '    Parallel.ForEach(layerList, Sub(current)
        '                                    Try
        '                                        For Each subFolder As Redemption.RDOFolder In current.Folders
        '                                            resultBag.Add(subFolder)
        '                                            currentLayer.Enqueue(subFolder)
        '                                        Next
        '                                    Catch ex As System.Exception
        '                                        _dbg("    ├ 錯誤", current.Name & " - " & ex.Message) ' by Gemini, 2026/04/10
        '                                    End Try
        '                                End Sub)
        'Loop

        'sw.Stop()
        '_dbg("    ├ 結束", $"{rootName} (RDO-Parallel BFS) | 資料夾總計: {resultBag.Count} | {sw.ElapsedMilliseconds}ms") ' by Gemini, 2026/04/10
        'Return resultBag.ToList()
    End Function

    ' Tv4_AfterSelect, Tv4_KeyDown, RenderLv4Group() — 舊版資料夾樹選取連動事件全數移除
    '   因 SimTree4 雙軌模式拆分，結果選取已由 Lv4Topic 專職負責。點擊 SimTree4 只作搜尋參考。
    '   此事件已無業務邏輯需求，故整段註解保留，以備日後參考。
    ' 2026/05/29 by Claude Sonnet 4.6: 
    ' ---------------------------------------------------------------------------------------------------------
    'Private Sub Tv4_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles SimTree4.AfterSelect

    '    ' ✅ 2026/04/20 by Gemini 2.0 Flash: 新增雙模式選取邏輯
    '    ' 模式 A: 資料夾模式 (目前的行為是選取後僅供搜尋參考，不執行連動)
    '    _dbg("開始 (A:資料夾模式)", e.Node.Text)
    '    If Not Lv4Topic.Visible Then Return ' by Claude Sonnet 4.6, 2026/05/29: 將 _isTv4ResultMode 替換為 Lv4Topic.Visible

    '    ' 模式 B: 主旨模式 (顯示主旨下的郵件清單)
    '    _dbg("開始 (B:主旨模式)", e.Node.Text)
    '    Dim mailList As List(Of MailItemInfo) = TryCast(e.Node.Tag, List(Of MailItemInfo))
    '    If mailList Is Nothing Then Return

    '    _lv4SortOrder = SortOrder.Descending    ' 每次點選新節點時，重置排序狀態為預設 (日期降冪)
    '    _lv4LastSortColumn = 2                  ' 收到日期所在的 index
    '    mailList.Sort(Function(a, b) b.RcvTime.CompareTo(a.RcvTime))  ' 排序: 依據時間遞減 (越新的在越前面)
    '    RenderLv4Result(mailList)
    '    _dbg("結束", $"顯示 {mailList.Count} 封系列郵件")

    'End Sub
    'Private Sub Tv4_KeyDown(sender As Object, e As KeyEventArgs) Handles SimTree4.KeyDown
    '    ' ✅ 2026/04/20 by Gemini 2.0 Flash: 處理 SimTree4 的快捷鍵與模式切換
    '    _dbg("開始", e.KeyCode.ToString())

    '    Select Case e.KeyCode
    '        Case Keys.Enter
    '            ' 在結果模式下按下 Enter 切換焦點到列表
    '            ' 2026/05/29 by Simon/Claude: 拆分SimTree4的雙重模式, 讓SimTree4回復到純粹的資料夾樹行為
    '            '   這裡的 Enter 只負責開始搜尋 (等同 Button4)，不再處理切換焦點的行為
    '            '   結果模式下的主旨選取改由 Lv4Topic 處理，
    '            ' todo: 其實只剩資料夾模式就可以合併回去原本的共用熱鍵處理函數了，已經沒有雙重模式的需求
    '            'If _isTv4ResultMode AndAlso Listview4.Items.Count > 0 Then Listview4.Focus()
    '            Button4.PerformClick()
    '            e.Handled = True

    '            'Case Keys.F5
    '            '    ' 按下 F5 等同 Button4 (重新開始掃描系列郵件)
    '            '    ' ✅ 2026/04/20: 在結果模式下按 F5 會自動引用上一資料夾重新掃描
    '            '    Button4.PerformClick()
    '            '    e.Handled = True

    '            'Case Keys.F6
    '            '    ' ✅ 2026/04/20 by Gemini 2.0 Flash: 切換左側樹排序方式 (數量/名稱)
    '            '    If _isTv4ResultMode AndAlso _tv4PrevTopicResults IsNot Nothing Then
    '            '        _tv4GroupSortByCount = Not _tv4GroupSortByCount
    '            '        RenderLv4Group(_tv4PrevTopicResults)
    '            '        _dbg("F6 按下：切換排序為", If(_tv4GroupSortByCount, "數量", "主旨"))
    '            '        e.Handled = True
    '            '    End If

    '            'Case Keys.Escape
    '            '    ' 按下 ESC：從結果模式恢復為資料夾模式
    '            '    If _isTv4ResultMode Then
    '            '        _dbg("ESC 按下：恢復資料夾模式 (PopNodeSnapshot)")
    '            '        _isTv4ResultMode = False
    '            '        Listview4.Items.Clear()

    '            '        ' ✅ 2026/05/23 by Simon/Claude: 改用 SimTree 內建快照還原，取代舊版手動重插節點
    '            '        '   RestoreTreeNodeSnap 內部處理：BeginUpdate/EndUpdate、節點插回、選取還原、EnsureVisible
    '            '        '   若插槽不存在（Fallback：重新載入資料夾樹）
    '            '        If Not SimTree4.RestoreTreeNodeSnap("folder-view") Then
    '            '            LoadStoreToTreeView(_pstStoreList, SimTree4)
    '            '            GotoDefaultInbox(SimTree4)
    '            '        End If

    '            '        ProgressBar1.Text = "已恢復資料夾樹模式。" : ProgressBar2.Text = ""
    '            '        SimTree4.Focus()
    '            '        e.Handled = True : e.SuppressKeyPress = True
    '            '    End If
    '    End Select

    'End Sub
    'Private Sub RenderLv4Group(topicDict As Dictionary(Of String, List(Of MailItemInfo)))
    '    ''' <summary>
    '    ''' ✅ 2026/04/20 by Gemini 2.0 Flash: 根據目前的排序模式渲染 Tab4 的主旨群組樹
    '    ''' </summary>

    '    _dbg("開始")
    '    If topicDict Is Nothing Then Return

    '    SimTree4.BeginUpdate()
    '    SimTree4.Nodes.Clear()
    '    ' _isTv4ResultMode = True ' by Claude Sonnet 4.6, 2026/05/29: 已廢棄，改用 Lv4Topic.Visible 代替

    '    _dbg("渲染系列清單", $"模式: {If(_tv4GroupSortByCount, "按數量", "按主旨")}")
    '    ' 根據旗標決定排序方式 (by Gemini 3 Flash, 2026/05/11: 改為 AddRange 模式以提升效能)
    '    Dim nodesArray = If(Not _tv4GroupSortByCount,
    '        topicDict.Where(Function(kvp) kvp.Value.Count > 1).
    '                  OrderBy(Function(kvp) kvp.Key).
    '                  Select(Function(kvp) New TreeNode($"{kvp.Key} ({kvp.Value.Count})") With {.Tag = kvp.Value}).ToArray(),
    '        topicDict.Where(Function(kvp) kvp.Value.Count > 1).
    '                  OrderByDescending(Function(kvp) kvp.Value.Count).
    '                  ThenBy(Function(kvp) kvp.Key).
    '                  Select(Function(kvp) New TreeNode($"{kvp.Key} ({kvp.Value.Count})") With {.Tag = kvp.Value}).ToArray())

    '    If nodesArray.Length > 0 Then SimTree4.Nodes.AddRange(nodesArray)
    '    SimTree4.EndUpdate()

    '    ' ✅ by Gemini 3.0 flash, 2026/04/21: 搜尋完成後，自動選取第一個結果並 Focus
    '    ' 💡 補充: 為了確保右側 Listview4 同步更新，手動呼叫事件處理器 (by Gemini 3.0 flash, 2026/04/21)
    '    If SimTree4.Nodes.Count > 0 Then
    '        Dim firstNode = SimTree4.Nodes(0)
    '        SimTree4.SelectedNode = firstNode
    '        SimTree4.Focus()
    '        Tv4_AfterSelect(SimTree4, New TreeViewEventArgs(firstNode))
    '    End If
    '    ProgressBar1.Text = $"找到 {SimTree4.Nodes.Count} 個系列 (排序: {If(_tv4GroupSortByCount, "數量", "主旨")})"
    '    _dbg("結束")

    'End Sub

    Private Async Function RenewAttachMailList(folder As Outlook.Folder, fPath As String) As Task
        ' ---------------------------------------------------------------
        ' RenewAttachMailList — 三路比對更新單一資料夾的 attach_maillist 快取
        ' 三路比對邏輯：
        '   新郵件   (live 有、DB 沒有) → 進入新的 mailList，SaveCache 時 INSERT
        '   已刪郵件 (DB 有、live 沒有) → 從 _cacheAttachFilename 清除 (DB row 留 CleanupOrphan 處理) 
        '   未變郵件 (live ∩ DB)        → 原有 filenames 快取保留，不重掃附件
        ' attach_filenames 永不重掃 (設計邊界，最耗時步驟留給 Tab3 搜尋時 lazy 觸發) 
        ' 2026/04/09 by Claude
        ' 2026/06/20 by Simon/Claude: 死碼可刪除, 已被 DbPurgeFolderMailRows 取代
        ' ---------------------------------------------------------------
    End Function

    ' 2026/06/22 by Simon/Claude Opus 4.8: IRM 保護信隔離夾名稱 (方案 Y: 每顆 PST 各建一個同名夾, 同 store 內搬)
    Private Const QUARANTINE_NAME As String = "_IRM_Protected"
    Private Async Function ScanAndMoveRpmsgRdo() As Task
        '' ============================================================================
        '' 2026/06/22 by Simon/Claude Opus 4.8: 【一次性工具】scan-and-move — 把 message.rpmsg 保護信隔離
        ''   作法: 依 SimTree3 選定節點掃整棵子樹, 命中(任一附件 .rpmsg)就用 RDO 把該信 Move 到
        ''         「同一顆 PST 的 _IRM_Protected 夾」(方案 Y, 同 store 內搬, 避開跨 store 不確定性)。
        ''   為何 scan-and-move 而非餵 EntryID: 搬移後 EntryID 會變, 來回 rebind 脆; 掃描當下手上就有 live
        ''         RDOMail, 就地搬最穩, 且搬前再驗一次 .rpmsg 防呆。全程走 RDO 不會觸發授權 modal。
        ''   ⚠ 破壞性: 信會離開原夾。搬完那些來源夾 + 隔離夾的 SQLite 快照會 stale, 需自行對受影響夾跑 RenewCache。
        ''   ⚠ 完整性: 請先把「所有可能含 rpmsg 的 PST」都選進 SimTree3 再執行, 才能一次搬乾淨。
        '' 2026/6/27 by Simon/Claude Opus 4.8: 原有呼叫GetSubtreeToListL3_RDO()退役, 改成新的GetSubtreeRdoByBatch()
        ''   For Each r In roots
        ''       Try (root)
        ''           rdoRoot → 走子樹 nodes → 釋 rdoRoot → 組 scanEids
        ''           For Each fe In scanEids
        ''               Try (folder)
        ''                   rdoF = Store.GetFolderFromID(fe.eid) → fName → 跳隔離夾
        ''                   [L1464–1520 原 items 掃描/搬移,不動]
        ''               Finally → 釋 rdoF
        ''               End Try
        ''           Next (fe)
        ''       Catch → log
        ''       End Try
        ''   Next (roots)
        '' ============================================================================

        '' ── 0. 確保 RDO 已載入 (改用 _rdo2 獨立 session) ──
        '' 2026/06/27 by Simon/Claude Opus 4.8: _rdo → _rdo2。兩者在 InitRdoSessionWithoutEULA 同生、ReleaseRdoSession 同滅,判 _rdo2 即可。
        'If _rdo2 Is Nothing Then Await InitRdoSessionWithoutEULA()
        'If _rdo2 Is Nothing Then _dbg("RDO隔離", "Redemption (_rdo2) 初始化失敗, 中止") : Return

        '' ── 1. UI 執行緒抽出選定節點 (EntryID, StoreID, 名稱) ──
        'Dim selectedNodes As List(Of TreeNode) = SimTree3.SelectedNodes
        'If selectedNodes Is Nothing OrElse selectedNodes.Count = 0 Then _dbg("RDO隔離", "SimTree3 未選取任何 PST/資料夾") : Return

        '' 2026/06/27 by Simon/Claude: UI 緒先解 _rdo2 store(GetRdoStore 內部用一般 Dictionary,須在 UI 緒呼叫;store 物件本身 free-threaded,可帶進 Task.Run)。
        ''   改抓 path(供走訪當 rootPath + 解 store);sid 不再需要(store-scoped 單參數解夾)。
        'Dim roots As New List(Of (store As Redemption.RDOStore, eid As String, path As String, name As String))(selectedNodes.Count)
        'For Each node As TreeNode In selectedNodes
        '    Dim f As Folder = TryCast(node.Tag, Folder)
        '    If f Is Nothing Then Continue For
        '    Dim p As String = SafeGetPath(f)
        '    Dim st As Redemption.RDOStore = GetRdoStore(p)   ' 記憶化快取 store,不釋放
        '    If st Is Nothing Then _dbg("RDO隔離", $"GetRdoStore 失敗,跳過: {f.Name}") : Continue For
        '    roots.Add((st, f.EntryID, p, f.Name))
        'Next
        'If roots.Count = 0 Then _dbg("RDO隔離", "選取節點皆非有效資料夾") : Return

        '' ── 破壞性動作, 先確認 ──
        'Dim dr As DialogResult = MessageBox.Show(
        '    $"即將掃描 {roots.Count} 個根節點, 把所有 message.rpmsg 保護信搬到各自 PST 的「{QUARANTINE_NAME}」夾。" & vbCrLf & vbCrLf &
        '    "此動作會改變封存結構且不易復原, 確定執行?", "確認隔離搬移", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
        'If dr <> DialogResult.Yes Then _dbg("RDO隔離", "使用者取消") : Return
        '_dbg("RDO隔離 開始", $"掃描 {roots.Count} 個根節點 ...")

        '' ── 2. 背景 scan-and-move (RDO free-threaded) ──
        'Dim moves As New List(Of String)
        'Dim movedCount As Integer = 0, failCount As Integer = 0
        'Dim scanned As Long = 0
        'Dim quarantineCache As New Dictionary(Of String, Redemption.RDOFolder)   ' key: store EntryID → 該 store 的隔離夾

        'Await Task.Run(
        '    Sub()
        '        For Each r In roots
        '            Try
        '                Dim rdoRoot As Redemption.RDOFolder = TryCast(r.store.GetFolderFromID(r.eid), Redemption.RDOFolder)
        '                If rdoRoot Is Nothing Then _dbg("RDO隔離 根節點失敗", $"{r.name} | root 解析失敗") : Continue For

        '                ' 批次走子樹拿 (eid,name,path);失敗退枚舉。rdoRoot 用完即釋。
        '                Dim nodes As List(Of (eid As String, name As String, path As String)) = GetSubtreeRdoBatch(r.store, rdoRoot, r.path)
        '                If nodes Is Nothing Then nodes = GetSubtreeRdoEnum(rdoRoot, r.path)
        '                Dim oRoot As Object = rdoRoot : TryMarshalRelease(oRoot)
        '                If nodes Is Nothing Then _dbg("RDO隔離 根節點失敗", $"{r.name} | 子樹走訪失敗") : Continue For

        '                ' root 自己 + 所有子孫,逐一在 _rdo2 store-scoped 重解夾後掃描
        '                Dim scanEids As New List(Of (eid As String, name As String))(nodes.Count + 1)
        '                scanEids.Add((r.eid, r.name))
        '                For Each nd In nodes : scanEids.Add((nd.eid, nd.name)) : Next

        '                For Each fe In scanEids
        '                    Dim rdoF As Redemption.RDOFolder = Nothing
        '                    Try
        '                        rdoF = TryCast(r.store.GetFolderFromID(fe.eid), Redemption.RDOFolder)
        '                        If rdoF Is Nothing Then Continue For
        '                        Dim fName As String = "" : Try : fName = rdoF.Name : Catch : End Try
        '                        If String.Equals(fName, QUARANTINE_NAME, StringComparison.OrdinalIgnoreCase) Then Continue For   ' 不掃隔離夾自己
        '                        Dim items = Nothing
        '                        Try
        '                            items = rdoF.Items
        '                            Dim cnt As Integer = items.Count
        '                            ' 由後往前 (Move 會把命中信移出本夾, 降序迭代不會影響尚未處理的索引)
        '                            For i As Integer = cnt To 1 Step -1
        '                                Dim m As Redemption.RDOMail = TryCast(items.Item(i), Redemption.RDOMail)
        '                                If m Is Nothing Then Continue For
        '                                Try
        '                                    scanned += 1
        '                                    If scanned Mod 5000 = 0 Then _dbg("RDO隔離 進行中", $"已掃 {scanned}, 已搬 {movedCount} ...")

        '                                    ' 偵測: 任一附件 .rpmsg 即命中 (搬前再驗, 防呆)
        '                                    Dim matched As String = Nothing
        '                                    For k As Integer = 1 To m.Attachments.Count
        '                                        Dim att As Redemption.RDOAttachment = m.Attachments.Item(k)
        '                                        Try
        '                                            Dim afn As String = att.FileName
        '                                            If afn IsNot Nothing AndAlso afn.EndsWith(".rpmsg", StringComparison.OrdinalIgnoreCase) Then matched = afn : Exit For
        '                                        Finally : TryMarshalRelease(att)
        '                                        End Try
        '                                    Next
        '                                    If matched Is Nothing Then Continue For   ' Finally 會釋放 m

        '                                    ' 命中: 先取所屬 store, get-or-create 該 store 的隔離夾
        '                                    Dim st As Redemption.RDOStore = m.Store
        '                                    Dim stKey As String = st.EntryID
        '                                    Dim qf As Redemption.RDOFolder = Nothing
        '                                    If Not quarantineCache.TryGetValue(stKey, qf) Then
        '                                        qf = GetOrCreateQuarantineRdo(st)
        '                                        quarantineCache(stKey) = qf
        '                                    End If

        '                                    ' 搬移前先擷取資訊 (Move 後 m 會失效、EntryID 會變)
        '                                    Dim rcv As String = "" : Try : rcv = m.ReceivedTime.ToString("yyyy/MM/dd HH:mm") : Catch : End Try
        '                                    Dim subj As String = "" : Try : subj = m.Subject : Catch : End Try
        '                                    Dim sndr As String = "" : Try : sndr = m.SenderName : Catch : End Try
        '                                    Dim eidOld As String = "" : Try : eidOld = m.EntryID : Catch : End Try
        '                                    Dim stName As String = "" : Try : stName = st.Name : Catch : End Try

        '                                    m.Move(qf)   ' ← 搬到隔離夾
        '                                    movedCount += 1
        '                                    _dbg($"搬移 #{movedCount}", $"{rcv} | {sndr} | {subj}")
        '                                    moves.Add(String.Join(vbTab, {$"#{movedCount}", rcv, "寄件:" & sndr, "主旨:" & subj, "原夾:" & fName, "PST:" & stName, "舊EntryID:" & eidOld}))
        '                                    TryMarshalRelease(st)
        '                                Catch ex As System.Exception
        '                                    failCount += 1
        '                                    _dbg("RDO隔離 搬移失敗", ex.Message)
        '                                Finally
        '                                    TryMarshalRelease(m)
        '                                End Try
        '                            Next
        '                        Catch ex As System.Exception
        '                            _dbg("RDO隔離 略過夾", $"{fName} | {ex.Message}")
        '                        Finally
        '                            TryMarshalRelease(items)
        '                        End Try
        '                    Finally
        '                        Dim oOF As Object = rdoF : TryMarshalRelease(oOF)   ' 每夾 store-scoped 開出,逐一釋
        '                    End Try
        '                Next   ' For Each fe In scanEids
        '            Catch ex As System.Exception
        '                _dbg("RDO隔離 根節點失敗", $"{r.name} | {ex.Message}")
        '            End Try
        '        Next   ' For Each r In roots
        '    End Sub)

        'For Each kv In quarantineCache : TryMarshalRelease(kv.Value) : Next

        '' ── 3. 寫搬移紀錄檔 (與 OLAcache.db 同目錄) ──
        'Dim logPath As String = ""
        'Try
        '    Dim baseDir As String = If(String.IsNullOrEmpty(_dbCachePath), My.Application.Info.DirectoryPath, System.IO.Path.GetDirectoryName(_dbCachePath))
        '    logPath = System.IO.Path.Combine(baseDir, $"RpmsgMoved_{DateTime.Now:yyyyMMdd_HHmmss}.log")
        '    Dim header As New List(Of String) From {
        '        $"# RDO 保護信隔離搬移   {DateTime.Now:yyyy/MM/dd HH:mm:ss}",
        '        $"# 已掃 {scanned} 封, 搬移 {movedCount} 封, 失敗 {failCount} 封 → 各 PST 的 {QUARANTINE_NAME} 夾",
        '        ""}
        '    System.IO.File.WriteAllLines(logPath, header.Concat(moves), System.Text.Encoding.UTF8)
        'Catch ex As System.Exception
        '    _dbg("RDO隔離 寫檔失敗", ex.Message)
        'End Try

        '_dbg("RDO隔離 完成", $"掃 {scanned} | 搬 {movedCount} | 失敗 {failCount} | log: {logPath}")
    End Function
    Private Function GetOrCreateQuarantineRdo(st As Redemption.RDOStore) As Redemption.RDOFolder
        ' 2026/06/22 by Simon/Claude Opus 4.8: 取得(或建立)指定 store 頂層的 _IRM_Protected 隔離夾
        Dim root As Redemption.RDOFolder = st.IPMRootFolder   ' store 的可見頂層夾 (PST 適用)
        Try
            Dim subs = root.Folders
            For i As Integer = 1 To subs.Count
                Dim f As Redemption.RDOFolder = subs.Item(i)
                If String.Equals(f.Name, QUARANTINE_NAME, StringComparison.OrdinalIgnoreCase) Then Return f   ' 已存在
            Next
            Return subs.Add(QUARANTINE_NAME)   ' ★ 唯一沒在文件逐字確認的 API (鏡像 OOM Folders.Add); 不編譯就是這行
        Finally
            TryMarshalRelease(root)
        End Try
    End Function


    ' 2026/06/19 驗證獨立 session _rdo2
    Private Async Function SpikeRdoIndependentSession() As Task
        ' 2026/06/19 by Simon/Claude: 拋棄式 spike — 驗證 RDO 獨立 session 三件事
        '   (1) Outlook 已掛載 PST 時，獨立 RDOSession 能否 Logon (PST 共享鎖)
        '   (2) 該獨立 session 能否讀到 RdoTest 內信件的附件檔名
        '   (3) 獨立 session 給的 EntryID，能否用 OOM _olNS.GetItemFromID 還原
        ' 測完即可整段刪除。請暫時掛到一個測試按鈕呼叫。
        ' ============================================================
        Dim log As New List(Of String)
        Dim firstEntryID As String = ""

        ' ── 先取 OOM 端 Gmail_2022 的 StoreID，供步驟3b比對用 ──
        Dim oomStoreId As String = ""
        Try
            For Each st As Outlook.Store In _olNS.Stores
                If st.DisplayName = "Gmail_2022" Then oomStoreId = st.StoreID : Exit For
            Next
        Catch ex As System.Exception
            log.Add("取 OOM StoreID 失敗: " & ex.Message)
        End Try

        ' ── 步驟1+2：背景執行緒用「獨立 session」讀取 ──
        Await Task.Run(Sub()
                           Dim sess As Redemption.RDOSession = Nothing
                           Try
                               sess = New Redemption.RDOSession()
                               ' ⚠ 確認點A：Logon 參數請依你的 Redemption 版本確認
                               '   目標 = 不沿用 Outlook session，建立獨立 MAPI session、用預設 profile、不彈窗
                               sess.Logon("", "", False, True)   ' (ProfileName, Pwd, ShowDialog, NewSession)
                               log.Add("步驟1 OK：獨立 session Logon 成功 (PST 共享鎖未擋住)")

                               ' ── 導覽到 \\Gmail_2022\收件匣\RdoTest ──
                               Dim store As Redemption.RDOStore = Nothing
                               For i As Integer = 1 To sess.Stores.Count
                                   If sess.Stores.Item(i).Name = "Gmail_2022" Then store = sess.Stores.Item(i) : Exit For
                               Next
                               If store Is Nothing Then log.Add("步驟2 失敗：獨立 session 找不到 Gmail_2022 store") : Return

                               ' ⚠ 確認點B：收件匣/RdoTest 確為 IPMRootFolder 下的層級
                               Dim inbox = store.IPMRootFolder.Folders.Item("收件匣")
                               Dim testFolder = inbox.Folders.Item("RdoTest")
                               log.Add($"步驟2 導覽 OK：RdoTest 共 {testFolder.Items.Count} 項")

                               Dim n As Integer = 0
                               For i As Integer = 1 To testFolder.Items.Count
                                   Dim msg = TryCast(testFolder.Items.Item(i), Redemption.RDOMail)
                                   If msg Is Nothing Then Continue For
                                   Dim names As New List(Of String)
                                   For a As Integer = 1 To msg.Attachments.Count
                                       names.Add(msg.Attachments.Item(a).FileName)
                                   Next
                                   If firstEntryID = "" Then firstEntryID = msg.EntryID
                                   n += 1
                                   log.Add($"  信{n}: 附件{names.Count}個 [{String.Join(", ", names)}]")
                               Next
                               log.Add($"步驟2 OK：成功讀出 {n} 封信的附件檔名")
                           Catch ex As System.Exception
                               log.Add("步驟1/2 例外: " & ex.Message)
                           Finally
                               Try : If sess IsNot Nothing Then sess.Logoff()
                               Catch : End Try
                               If sess IsNot Nothing Then TryMarshalRelease(sess)
                           End Try
                       End Sub)

        ' ── 步驟3：回 UI 執行緒，用 OOM 還原「獨立 session 給的」EntryID ──
        If firstEntryID = "" Then
            log.Add("步驟3 跳過：沒有取得任何 EntryID")
        Else
            ' 3a：單參數
            Try
                Dim m1 = TryCast(_olNS.GetItemFromID(firstEntryID), Outlook.MailItem)
                log.Add(If(m1 IsNot Nothing, "步驟3a OK：單參數還原成功 → " & m1.Subject,
                                         "步驟3a 失敗：單參數回傳 Nothing"))
            Catch ex As System.Exception
                log.Add("步驟3a 例外: " & ex.Message)
            End Try
            ' 3b：帶 OOM StoreID
            If oomStoreId <> "" Then
                Try
                    Dim m2 = TryCast(_olNS.GetItemFromID(firstEntryID, oomStoreId), Outlook.MailItem)
                    log.Add(If(m2 IsNot Nothing, "步驟3b OK：帶StoreID還原成功 → " & m2.Subject,
                                             "步驟3b 失敗：帶StoreID回傳 Nothing"))
                Catch ex As System.Exception
                    log.Add("步驟3b 例外: " & ex.Message)
                End Try
            End If
        End If

        MessageBox.Show(String.Join(vbCrLf, log), "RDO Spike 結果")
    End Function   ' 2026/6/19~20 獨立 session 給的 EntryID，能否用 OOM _olNS.GetItemFromID 還原
    Private Async Function SpikeParallelReadBenchmark() As Task
        ' 2026/06/22 by Simon/Claude Opus 4.8: 拋棄式 spike P3 — 量測「同 profile 多獨立 session、各讀不同 PST」
        '   的真實平行加速。回答整輪調查唯一未解的問題: K 條 session 跨 PST 並行讀取, wall-clock 是否
        '   勝過序列, 還是被 MSPST provider / 實體磁碟 I/O 序列化。
        '   ★ 兩種 workload 分別計時(附件檔名 vs 內文), 因 Tab3/Tab5 負載特性可能不同。
        '   ★ 公平性: 每個 (workload,K) 各讀「獨立的冷 block」(不同信), 避免暖快取讓後跑的 config 假性變快。
        '   ★ 用與 production 同一支 API: sess.GetMessageFromID(EntryID) + rdoMsg.Attachments/.Body
        Const N As Integer = 2000      ' 每 PST 每個 block 的冷讀信數(想要更穩可調 2000, 時間約翻倍)
        Const M As Integer = 4         ' 取幾個「夠大」的 PST 當標的(K=4 時每 worker 各 1 個)
        Const BLOCKS As Integer = 6    ' 2 workload × 3 K-config; 每 PST 需 >= BLOCKS*N 封冷信

        If _rdo Is Nothing Then Await InitRdoSessionWithoutEULA()
        If _rdo Is Nothing Then _dbg("P3", "Redemption 初始化失敗, 中止") : Return

        Dim profileName As String = ""
        Try : profileName = CStr(CallByName(_rdo, "ProfileName", CallType.Get)) : Catch : End Try
        If profileName = "" Then _dbg("P3", "取不到 _rdo.ProfileName, 中止") : Return
        _dbg("P3", $"===== 平行讀取量測 開始 (profile=[{profileName}], N={N}, M={M}) =====")

        ' ── 1. 收集階段: 臨時一條 session 走訪, 挑 M 個有 >= BLOCKS*N 封的 PST, 各收 BLOCKS*N 個 EntryID ──
        '    (EntryID 是字串、跨 session 通用, 收一次給所有 worker 重用; RDOStore 物件不可跨 session 持有)
        Dim need As Integer = BLOCKS * N
        Dim pstEntryIds As New List(Of (pst As String, pstPath As String, ids As List(Of String)))()
        Dim swCollect As New Stopwatch() : swCollect.Start()
        Await Task.Run(Sub()
                           Dim sess As Redemption.RDOSession = Nothing
                           Try
                               sess = New Redemption.RDOSession()
                               sess.Logon(profileName, "", False, True)
                               For si As Integer = 1 To sess.Stores.Count
                                   If pstEntryIds.Count >= M Then Exit For
                                   Dim st = sess.Stores.Item(si)
                                   Dim nm As String = "" : Try : nm = st.Name : Catch : End Try
                                   Dim pp As String = "" : Try : pp = CStr(CallByName(st, "PstPath", CallType.Get)) : Catch : End Try   ' (c)store-scoped 需 PstPath 去 FindStoreByPath 開 store
                                   Dim ids As New List(Of String)()
                                   Try
                                       Dim stk As New Stack(Of Redemption.RDOFolder)()
                                       stk.Push(st.IPMRootFolder)
                                       Do While stk.Count > 0 AndAlso ids.Count < need
                                           Dim fld = stk.Pop()
                                           Dim cnt As Integer = fld.Items.Count
                                           For ii As Integer = 1 To cnt
                                               If ids.Count >= need Then Exit For
                                               Try
                                                   Dim mm = TryCast(fld.Items.Item(ii), Redemption.RDOMail)
                                                   If mm IsNot Nothing Then ids.Add(mm.EntryID)
                                               Catch : End Try
                                           Next
                                           For fi As Integer = 1 To fld.Folders.Count
                                               stk.Push(fld.Folders.Item(fi))
                                           Next
                                       Loop
                                   Catch : End Try
                                   If ids.Count >= need AndAlso pp <> "" Then
                                       pstEntryIds.Add((nm, pp, ids))
                                       _dbg(" │收集", $"採用 PST [{nm}] (收到 {ids.Count} EntryID, PstPath=[{pp}])")
                                   End If
                               Next
                           Catch ex As System.Exception
                               _dbg(" │收集", "例外: " & ex.GetBaseException().Message)
                           Finally
                               If sess IsNot Nothing Then
                                   Try : sess.Logoff() : Catch : End Try
                                   TryMarshalRelease(sess)
                               End If
                           End Try
                       End Sub)
        swCollect.Stop()
        If pstEntryIds.Count < M Then
            _dbg(" │✗", $"只湊到 {pstEntryIds.Count} 個夠大的 PST(需 {M}, 每個需 >= {need} 封)。請降低 N 或確認在 Work profile。中止。")
            Return
        End If
        _dbg(" │收集", $"完成: {pstEntryIds.Count} 個 PST, 各 {need} EntryID, 耗時 {swCollect.Elapsed.TotalSeconds:F1}s (不計入吞吐量)")

        ' ── 2. 對 2 種 workload × K=1/2/4 量測 ──
        Dim workloads = {"附件檔名", "內文"}
        Dim kConfigs = {1, 2, 4}
        Dim summary As New List(Of String)()

        For w As Integer = 0 To workloads.Length - 1
            Dim isBody As Boolean = (w = 1)
            For kc As Integer = 0 To kConfigs.Length - 1
                Dim K As Integer = kConfigs(kc)
                Dim blockIdx As Integer = w * 3 + kc            ' 0..5, 每 config 取不同冷 block
                Dim lo As Integer = blockIdx * N

                ' 把 M 個 PST round-robin 分給 K 個 worker
                Dim groups As New List(Of List(Of (pst As String, pstPath As String, ids As List(Of String))))()
                For g As Integer = 0 To K - 1 : groups.Add(New List(Of (pst As String, pstPath As String, ids As List(Of String)))()) : Next
                For pi As Integer = 0 To pstEntryIds.Count - 1 : groups(pi Mod K).Add(pstEntryIds(pi)) : Next

                Dim bag As New System.Collections.Concurrent.ConcurrentBag(Of (logonMs As Double, rs As Double, re As Double, mails As Integer, fails As Integer, payload As Long, storeMs As Double, withAttach As Integer))()
                Dim swWall As New Stopwatch() : swWall.Start()

                Dim tasks As New List(Of Task)()
                For g As Integer = 0 To K - 1
                    Dim myGroup = groups(g)
                    tasks.Add(Task.Run(Sub()
                                           Dim sess As Redemption.RDOSession = Nothing
                                           Dim mails As Integer = 0, fails As Integer = 0
                                           Dim bodyChars As Long = 0
                                           Dim payload As Long = 0          ' 附件: 總附件數; 內文: 總字元數 — 揪空轉用
                                           Dim withAttach As Integer = 0    ' 有附件(Count>0)的信數 — 確認取樣是否多為無附件信
                                           Dim storeMs As Double = 0        ' 本 worker 累計開 store(FindStoreByPath)耗時
                                           Dim swLogon As New Stopwatch() : swLogon.Start()
                                           Try
                                               sess = New Redemption.RDOSession()
                                               sess.Logon(profileName, "", False, True)
                                           Catch ex As System.Exception
                                               _dbg(" │✗", $"K={K} worker logon 失敗: {ex.GetBaseException().Message}") : Return
                                           End Try
                                           swLogon.Stop()
                                           Dim rs As Double = swWall.Elapsed.TotalSeconds
                                           Try
                                               For Each pe In myGroup
                                                   ' (c)store-scoped: 每個 PST 在本 worker session 內開一次 store, 之後該 PST 所有信都用 store.GetMessageFromID
                                                   ' (P4 已驗: 跨 session 單參數會 MAPI_E_UNKNOWN_ENTRYID, store-scoped 則 10/10)
                                                   Dim swStore As New Stopwatch() : swStore.Start()
                                                   Dim stStore As Redemption.RDOStore = FindStoreByPath(sess, pe.pstPath)
                                                   swStore.Stop() : storeMs += swStore.Elapsed.TotalMilliseconds   ' A: 開 store 耗時
                                                   If stStore Is Nothing Then fails += N : Continue For    ' 此 PST 在本 session 找不到 → 整塊計失敗
                                                   For idx As Integer = lo To lo + N - 1
                                                       Dim eid As String = pe.ids(idx)
                                                       Try
                                                           Dim rm = TryCast(stStore.GetMessageFromID(eid), Redemption.RDOMail)
                                                           If rm Is Nothing Then fails += 1 : Continue For
                                                           If isBody Then
                                                               Dim b As String = rm.Body
                                                               If b IsNot Nothing Then bodyChars += b.Length : payload += b.Length   ' 強制讀取內文 + 計字元(揪空轉)
                                                           Else
                                                               Dim ac As Integer = rm.Attachments.Count
                                                               If ac > 0 Then withAttach += 1                ' A: 這封真的有附件
                                                               For a As Integer = 1 To ac
                                                                   Dim fn As String = rm.Attachments.Item(a).FileName
                                                                   payload += 1                              ' A: 真讀到的附件檔名數
                                                               Next
                                                           End If
                                                           mails += 1
                                                       Catch
                                                           fails += 1
                                                       End Try
                                                   Next
                                               Next
                                           Catch ex As System.Exception
                                               _dbg(" │✗", $"K={K} worker 讀取例外: {ex.GetBaseException().Message}")
                                           Finally
                                               Dim re As Double = swWall.Elapsed.TotalSeconds
                                               bag.Add((swLogon.Elapsed.TotalMilliseconds, rs, re, mails, fails, payload, storeMs, withAttach))
                                               If sess IsNot Nothing Then
                                                   Try : sess.Logoff() : Catch : End Try
                                                   TryMarshalRelease(sess)
                                               End If
                                           End Try
                                       End Sub))
                Next
                Await Task.WhenAll(tasks)
                swWall.Stop()

                ' 聚合(手動迴圈, 不依賴 LINQ import)
                Dim arr = bag.ToArray()
                Dim totMails As Integer = 0, totFails As Integer = 0
                Dim sumLogon As Double = 0
                Dim readStart As Double = Double.MaxValue, readEnd As Double = 0
                For Each x In arr
                    totMails += x.mails : totFails += x.fails : sumLogon += x.logonMs
                    If x.rs < readStart Then readStart = x.rs
                    If x.re > readEnd Then readEnd = x.re
                Next
                If arr.Length = 0 Then readStart = 0 : readEnd = 0
                Dim avgLogon As Double = If(arr.Length > 0, sumLogon / arr.Length, 0)
                Dim wallRead As Double = Math.Max(0.001, readEnd - readStart)
                Dim thru As Double = totMails / wallRead
                ' A: 額外彙整 — 開store耗時、實際讀取量(揪空轉)、worker 重疊度
                Dim totPayload As Long = 0, totStoreMs As Double = 0, totWithAttach As Integer = 0
                Dim sumReadSpan As Double = 0   ' 各 worker 純讀取時間(rs..re)總和; 與 wallRead 比即重疊度
                For Each x In arr
                    totPayload += x.payload : totStoreMs += x.storeMs : totWithAttach += x.withAttach
                    sumReadSpan += (x.re - x.rs)
                Next
                Dim overlap As Double = sumReadSpan / wallRead   ' ≈K 表完全重疊平行; ≈1 表幾乎沒重疊
                Dim payloadDesc As String = If(w = 1, $"內文{totPayload}字元", $"附件{totPayload}個(有附件信{totWithAttach}/{totMails})")
                Dim line As String = $"[{workloads(w)}] K={K}: 讀 {totMails} 封, 讀取wall={wallRead:F1}s, 吞吐={thru:F0} 封/s, {payloadDesc}, 開store均={totStoreMs / Math.Max(1, arr.Length):F0}ms, 重疊={overlap:F2}x, logon均={avgLogon:F0}ms, resolve失敗={totFails}"
                _dbg(" │量測", line)
                summary.Add(line)
            Next
        Next

        _dbg("P3", "===== 量測結束, 摘要(看 K=2/4 吞吐相對 K=1 有沒有上去) =====")
        For Each s In summary : _dbg(" │摘要", s) : Next
        _dbg("P3", "===== 請把本段全部貼回 =====")
    End Function   ' 2026/06/22 P3量測「同 profile 多獨立 session、各讀不同 PST」的真實平行加速
    Private Async Function SpikeResolveFormCompare() As Task
        ' 2026/06/22 by Simon/Claude Opus 4.8: 拋棄式 spike B — 釘死「P3 附件 K=1 達 5589 封/s, 但 production_1/_2 只有 200 多」這 25 倍矛盾。空轉假設已被推翻(本批信 ~55% 有附件), 剩三混淆變數:
        '     (a)resolve 形式  (b)session 種類  (c)取樣信 vs sourceList 不同 ←本 spike 用「同批信讀三遍」消掉 c. 單執行緒(純比 per-call 成本, 不平行), 同一批信讀三種形式:
        '     (1)共用_rdo 單參數      = 現行 production
        '     (2)共用_rdo store-scoped → (1)vs(2)= resolve 形式效應(同一 session)
        '     (3)獨立session store-scoped = P3 → (2)vs(3)= session 種類效應
        '   依賴: FindStoreByPath(寫 P4 時放的 class-level 函數)。前提: Outlook 切 Work profile。測完即可整段刪除。
        Const N As Integer = 2000      ' 取樣信數(單執行緒, 同一批讀三遍; 夠大讓 封/s 穩定)
        If _rdo Is Nothing Then Await InitRdoSessionWithoutEULA()
        If _rdo Is Nothing Then _dbg("B", "Redemption 初始化失敗, 中止") : Return
        Dim profileName As String = ""
        Try : profileName = CStr(CallByName(_rdo, "ProfileName", CallType.Get)) : Catch : End Try
        _dbg("B", $"===== resolve 形式對照 (profile=[{profileName}], N={N}, 單執行緒) =====")

        Await Task.Run(Sub()
                           ' ── 1. 用共用 _rdo 走訪頭部湊 N 封, 記每封所屬 pstPath(供 store-scoped 分組) ──
                           Dim sample As New List(Of (eid As String, pstPath As String))()
                           Try
                               For si As Integer = 1 To _rdo.Stores.Count
                                   If sample.Count >= N Then Exit For
                                   Dim st = _rdo.Stores.Item(si)
                                   Dim pp As String = "" : Try : pp = CStr(CallByName(st, "PstPath", CallType.Get)) : Catch : End Try
                                   If pp = "" Then Continue For
                                   Try
                                       Dim stk As New Stack(Of Redemption.RDOFolder)() : stk.Push(st.IPMRootFolder)
                                       Do While stk.Count > 0 AndAlso sample.Count < N
                                           Dim f = stk.Pop()
                                           For ii As Integer = 1 To f.Items.Count
                                               If sample.Count >= N Then Exit For
                                               Dim mm = TryCast(f.Items.Item(ii), Redemption.RDOMail)
                                               If mm IsNot Nothing Then sample.Add((mm.EntryID, pp))
                                           Next
                                           For fi As Integer = 1 To f.Folders.Count : stk.Push(f.Folders.Item(fi)) : Next
                                       Loop
                                   Catch : End Try
                               Next
                           Catch ex As System.Exception
                               _dbg(" │收集✗", ex.GetBaseException().Message)
                           End Try
                           If sample.Count = 0 Then _dbg(" │✗", "沒取到信, 中止") : Return

                           ' 按 pstPath 分組(供 (2)(3) store-scoped 重用 store; 手動建, 不依賴 LINQ import)
                           Dim groups As New Dictionary(Of String, List(Of String))()
                           For Each s In sample
                               Dim lst As List(Of String) = Nothing
                               If Not groups.TryGetValue(s.pstPath, lst) Then lst = New List(Of String)() : groups(s.pstPath) = lst
                               lst.Add(s.eid)
                           Next
                           _dbg(" │收集", $"取樣 {sample.Count} 封(跨 {groups.Count} 個 PST)")

                           ' 小工具: resolve 後讀附件檔名數(回 -1 表 resolve 失敗)
                           Dim readAttach = Function(rm As Redemption.RDOMail) As Integer
                                                If rm Is Nothing Then Return -1
                                                Dim c As Integer = rm.Attachments.Count
                                                For a As Integer = 1 To c : Dim fn As String = rm.Attachments.Item(a).FileName : Next
                                                Return c
                                            End Function

                           ' ── (1) 共用 _rdo 單參數 (現行 production) ──
                           Dim sw1 As New Stopwatch() : sw1.Start()
                           Dim att1 As Long = 0, fail1 As Integer = 0
                           For Each s In sample
                               Try
                                   Dim c = readAttach(TryCast(_rdo.GetMessageFromID(s.eid), Redemption.RDOMail))
                                   If c < 0 Then fail1 += 1 Else att1 += c
                               Catch : fail1 += 1
                               End Try
                           Next
                           sw1.Stop()
                           _dbg(" │(1)", $"共用_rdo 單參數: {sample.Count / Math.Max(0.001, sw1.Elapsed.TotalSeconds):F0} 封/s ({sw1.Elapsed.TotalSeconds:F1}s, 附件{att1}, 失敗{fail1})")

                           ' ── (2) 共用 _rdo, store-scoped (只換 resolve 形式, 同一 session) ──
                           Dim sw2 As New Stopwatch() : sw2.Start()
                           Dim att2 As Long = 0, fail2 As Integer = 0
                           For Each kv In groups
                               Dim store = FindStoreByPath(_rdo, kv.Key)
                               If store Is Nothing Then fail2 += kv.Value.Count : Continue For
                               For Each eid In kv.Value
                                   Try
                                       Dim c = readAttach(TryCast(store.GetMessageFromID(eid), Redemption.RDOMail))
                                       If c < 0 Then fail2 += 1 Else att2 += c
                                   Catch : fail2 += 1
                                   End Try
                               Next
                           Next
                           sw2.Stop()
                           _dbg(" │(2)", $"共用_rdo store-scoped: {sample.Count / Math.Max(0.001, sw2.Elapsed.TotalSeconds):F0} 封/s ({sw2.Elapsed.TotalSeconds:F1}s, 附件{att2}, 失敗{fail2})")

                           ' ── (3) 獨立 session, store-scoped (= P3 形式) ──
                           Dim sess As Redemption.RDOSession = Nothing
                           Try
                               sess = New Redemption.RDOSession()
                               sess.Logon(profileName, "", False, True)
                               Dim sw3 As New Stopwatch() : sw3.Start()
                               Dim att3 As Long = 0, fail3 As Integer = 0
                               For Each kv In groups
                                   Dim store = FindStoreByPath(sess, kv.Key)
                                   If store Is Nothing Then fail3 += kv.Value.Count : Continue For
                                   For Each eid In kv.Value
                                       Try
                                           Dim c = readAttach(TryCast(store.GetMessageFromID(eid), Redemption.RDOMail))
                                           If c < 0 Then fail3 += 1 Else att3 += c
                                       Catch : fail3 += 1
                                       End Try
                                   Next
                               Next
                               sw3.Stop()
                               _dbg(" │(3)", $"獨立session store-scoped: {sample.Count / Math.Max(0.001, sw3.Elapsed.TotalSeconds):F0} 封/s ({sw3.Elapsed.TotalSeconds:F1}s, 附件{att3}, 失敗{fail3})")
                           Catch ex As System.Exception
                               _dbg(" │(3)✗", ex.GetBaseException().Message)
                           Finally
                               If sess IsNot Nothing Then
                                   Try : sess.Logoff() : Catch : End Try
                                   TryMarshalRelease(sess)
                               End If
                           End Try
                       End Sub)
        _dbg("B", "===== 對照結束, 請貼回(三個附件數應一致才公平) =====")
    End Function      ' 2026/6/23, 修改P3, 開始比較獨立session 形式對效能的影響倍數, 與平行度效能吞吐量測試
    Private Async Function SpikeBodyResolveCompare() As Task
        ' 2026/06/22 by Simon/Claude Opus 4.8: 拋棄式 spike B-內文版 — 驗證「內文讀取換獨立 session 是否也有 ~10×」。
        '   注意: 內文 production 路徑(GetMailBodyOOM 第2190行)走 OOM, 不是 _rdo, 故基準與附件版不同, 測三條:
        '     (1) OOM _olNS.GetItemFromID + .Body  = 內文現行 production 基準(你說的 70~80 封/s 來源)
        '     (2) 共用 _rdo store-scoped + .Body    → (2)vs(3) 對照「共用 vs 獨立 session」這條槓桿在內文是否成立
        '     (3) 獨立 session store-scoped + .Body = 目標形式
        '   防 IRM: 取樣時用 RDO 預掃 MessageClass, 跳過 IPM.Note.* 受保護(rpmsg)信, 避免 OOM .Body 卡死授權 modal。
        '   ★全程 UI/STA 緒同步跑: OOM COM 不可進 Task.Run; N=1000 單執行緒, UI 短暫凍結可接受。
        '   依賴: FindStoreByPath(P4 放的)。前提: Outlook 切 Work profile。測完即整段刪除。
        Const N As Integer = 1000
        If _rdo Is Nothing Then Await InitRdoSessionWithoutEULA()
        If _rdo Is Nothing Then _dbg("B內文", "Redemption 初始化失敗, 中止") : Return
        Dim profileName As String = ""
        Try : profileName = CStr(CallByName(_rdo, "ProfileName", CallType.Get)) : Catch : End Try
        _dbg("B內文", $"===== 內文 resolve 形式對照 (profile=[{profileName}], N={N}, 單執行緒/UI緒) =====")

        Await Task.Run(Sub()
                           ' ── 1. 用共用 _rdo 走訪頭部湊 N 封, 防 IRM: 跳過受保護信(MessageClass 含 .rpmsg 或非 IPM.Note 之保護類) ──
                           Dim sample As New List(Of (eid As String, pstPath As String))()
                           Dim skipIrm As Integer = 0
                           Try
                               For si As Integer = 1 To _rdo.Stores.Count
                                   If sample.Count >= N Then Exit For
                                   Dim st = _rdo.Stores.Item(si)
                                   Dim pp As String = "" : Try : pp = CStr(CallByName(st, "PstPath", CallType.Get)) : Catch : End Try
                                   If pp = "" Then Continue For
                                   Try
                                       Dim stk As New Stack(Of Redemption.RDOFolder)() : stk.Push(st.IPMRootFolder)
                                       Do While stk.Count > 0 AndAlso sample.Count < N
                                           Dim f = stk.Pop()
                                           For ii As Integer = 1 To f.Items.Count
                                               If sample.Count >= N Then Exit For
                                               Dim mm = TryCast(f.Items.Item(ii), Redemption.RDOMail)
                                               If mm Is Nothing Then Continue For
                                               Dim mc As String = "" : Try : mc = CStr(mm.MessageClass) : Catch : End Try
                                               ' IRM/RMS 保護信外層 MessageClass 多為 IPM.Note.SMIME 或含 rpmsg; 保守只收純 IPM.Note
                                               If mc.StartsWith("IPM.Note", StringComparison.OrdinalIgnoreCase) AndAlso
                               mc.IndexOf("rpmsg", StringComparison.OrdinalIgnoreCase) < 0 AndAlso
                               mc.IndexOf("SMIME", StringComparison.OrdinalIgnoreCase) < 0 Then
                                                   sample.Add((mm.EntryID, pp))
                                               Else
                                                   skipIrm += 1
                                               End If
                                           Next
                                           For fi As Integer = 1 To f.Folders.Count : stk.Push(f.Folders.Item(fi)) : Next
                                       Loop
                                   Catch : End Try
                               Next
                           Catch ex As System.Exception
                               _dbg(" │收集✗", ex.GetBaseException().Message)
                           End Try
                           If sample.Count = 0 Then _dbg(" │✗", "沒取到信, 中止") : Return

                           ' 按 pstPath 分組(供 (2)(3) store-scoped 重用 store)
                           Dim groups As New Dictionary(Of String, List(Of String))()
                           For Each s In sample
                               Dim lst As List(Of String) = Nothing
                               If Not groups.TryGetValue(s.pstPath, lst) Then lst = New List(Of String)() : groups(s.pstPath) = lst
                               lst.Add(s.eid)
                           Next
                           _dbg(" │收集", $"取樣 {sample.Count} 封(跨 {groups.Count} 個 PST), 跳過疑似IRM {skipIrm} 封")

                           ' ── (2) 共用 _rdo store-scoped + .Body ──
                           Dim sw2 As New Stopwatch() : sw2.Start()
                           Dim chars2 As Long = 0, fail2 As Integer = 0
                           For Each kv In groups
                               Dim store = FindStoreByPath(_rdo, kv.Key)
                               If store Is Nothing Then fail2 += kv.Value.Count : Continue For
                               For Each eid In kv.Value
                                   Try
                                       Dim rm = TryCast(store.GetMessageFromID(eid), Redemption.RDOMail)
                                       If rm Is Nothing Then fail2 += 1 : Continue For
                                       Dim b As String = rm.Body : If b IsNot Nothing Then chars2 += b.Length
                                   Catch : fail2 += 1
                                   End Try
                               Next
                           Next
                           sw2.Stop()
                           _dbg(" │(2)", $"共用_rdo .Body: {sample.Count / Math.Max(0.001, sw2.Elapsed.TotalSeconds):F0} 封/s ({sw2.Elapsed.TotalSeconds:F1}s, 字元{chars2}, 失敗{fail2})")

                           ' ── (3) 獨立 session store-scoped + .Body (RDO 在背景緒 OK, 但本支求一致仍在 UI 緒同步跑) ──
                           Dim sess As Redemption.RDOSession = Nothing
                           Try
                               sess = New Redemption.RDOSession()
                               sess.Logon(profileName, "", False, True)
                               Dim sw3 As New Stopwatch() : sw3.Start()
                               Dim chars3 As Long = 0, fail3 As Integer = 0
                               For Each kv In groups
                                   Dim store = FindStoreByPath(sess, kv.Key)
                                   If store Is Nothing Then fail3 += kv.Value.Count : Continue For
                                   For Each eid In kv.Value
                                       Try
                                           Dim rm = TryCast(store.GetMessageFromID(eid), Redemption.RDOMail)
                                           If rm Is Nothing Then fail3 += 1 : Continue For
                                           Dim b As String = rm.Body : If b IsNot Nothing Then chars3 += b.Length
                                       Catch : fail3 += 1
                                       End Try
                                   Next
                               Next
                               sw3.Stop()
                               _dbg(" │(3)", $"獨立session .Body: {sample.Count / Math.Max(0.001, sw3.Elapsed.TotalSeconds):F0} 封/s ({sw3.Elapsed.TotalSeconds:F1}s, 字元{chars3}, 失敗{fail3})")
                           Catch ex As System.Exception
                               _dbg(" │(3)✗", ex.GetBaseException().Message)
                           Finally
                               If sess IsNot Nothing Then
                                   Try : sess.Logoff() : Catch : End Try
                                   TryMarshalRelease(sess)
                               End If
                           End Try
                           _dbg("B內文", "===== 對照結束, 請貼回(三個字元數應相近才公平) =====")

                       End Sub)
    End Function      ' 驗證「內文讀取換獨立 session 效能與平行度效能吞吐量測試」
    Private Sub DumpResolve(tag As String, sess As Redemption.RDOSession, store As Redemption.RDOStore, eids As List(Of String), storeEid As String)
        Dim okA As Integer = 0, okB As Integer = 0, okC As Integer = 0
        Dim eA As String = "", eB As String = "", eC As String = ""
        For Each eid As String In eids
            Try
                If TryCast(sess.GetMessageFromID(eid), Redemption.RDOMail) IsNot Nothing Then okA += 1
            Catch ex As System.Exception
                If eA = "" Then eA = ex.GetBaseException().Message
            End Try
            Try
                If TryCast(sess.GetMessageFromID(eid, storeEid), Redemption.RDOMail) IsNot Nothing Then okB += 1
            Catch ex As System.Exception
                If eB = "" Then eB = ex.GetBaseException().Message
            End Try
            Try
                If store IsNot Nothing AndAlso TryCast(store.GetMessageFromID(eid), Redemption.RDOMail) IsNot Nothing Then okC += 1
            Catch ex As System.Exception
                If eC = "" Then eC = ex.GetBaseException().Message
            End Try
        Next
        _dbg($" │{tag}", $"(a)單參數={okA}/{eids.Count} [{eA}]　(b)雙參數={okB}/{eids.Count} [{eB}]　(c)store-scoped={okC}/{eids.Count} store={store IsNot Nothing} [{eC}]")
    End Sub ' P4 輔助: 對同一批 EntryID 試三種 resolve 形式, 各記成功數與首個例外
    Private Function FindStoreByPath(sess As Redemption.RDOSession, path As String) As Redemption.RDOStore
        If path = "" Then Return Nothing
        For i As Integer = 1 To sess.Stores.Count
            Dim pp As String = ""
            Try : pp = CStr(CallByName(sess.Stores.Item(i), "PstPath", CallType.Get)) : Catch : End Try
            If String.Equals(pp, path, StringComparison.OrdinalIgnoreCase) Then Return sess.Stores.Item(i)
        Next
        Return Nothing
    End Function ' P4 輔助: 用 PstPath 在指定 session 找 RDOStore

    ' 2026/06/23 驗證獨立 session _rdo2
    Private Async Function SpikeResolveFormOnRdo2() As Task
        ' =================================================================
        ' 2026/06/23 by Simon/Claude: 探針 — 驗證獨立 session _rdo2 的 resolve 形式
        '   目的: 用 OOM 取得的 (EntryID, OOM StoreID, FolderPath) 在 _rdo2 上分別試三種
        '         resolve, 決定 production 該走「雙參數」還是「store-scoped」。
        '   判讀: 看哪種形式 resolve 成功率高、且 Subject 對得上 (= 真解到, 非空 handle)。
        '   ※ 純診斷, 不動 production; 用完即可整段刪除。
        ' =================================================================
        If _rdo2 Is Nothing Then Await InitRdoSessionWithoutEULA()
        If _rdo2 Is Nothing Then _dbg("探針中止", "_rdo2 初始化失敗") : Return

        ' ── 1. 印 _rdo2 身分 (確認登對 profile、看得到哪些 store) ──
        Dim storeNames As New List(Of String)
        Try
            For i As Integer = 1 To _rdo2.Stores.Count : storeNames.Add(_rdo2.Stores.Item(i).Name) : Next
        Catch ex As System.Exception
            _dbg("探針", $"列舉 _rdo2.Stores 失敗: {ex.Message}")
        End Try
        _dbg("探針 _rdo2", $"ProfileName=[{_rdo2.ProfileName}] Stores={storeNames.Count}")
        _dbg("探針 _rdo2 stores", String.Join(" | ", storeNames))

        ' ── 2. 從 OOM 採樣: 最多 3 個 PST、每 PST 最多 4 封, 合計上限 ~12 ──
        Dim samples As New List(Of (eid As String, sid As String, fpath As String, subj As String))
        Dim storeTaken As Integer = 0
        For si As Integer = 1 To _olNS.Stores.Count
            If storeTaken >= 3 Then Exit For
            Dim st As Outlook.Store = Nothing
            Try
                st = _olNS.Stores.Item(si)
                If String.IsNullOrEmpty(st.FilePath) Then Continue For   ' 跳過無檔 store (iCloud 等)
                Dim grabbed As Integer = HarvestFromStore(st, st.StoreID, samples, 4)
                If grabbed > 0 Then storeTaken += 1
            Catch ex As System.Exception
                _dbg("探針採樣", $"store#{si} 失敗: {ex.Message}")
            Finally
                TryMarshalRelease(st)
            End Try
        Next
        _dbg("探針採樣", $"共取得 {samples.Count} 封樣本 (跨 {storeTaken} 個 PST)")
        If samples.Count = 0 Then _dbg("探針中止", "採樣 0 封") : Return

        ' ── 3. 三種形式逐封測試 ──
        Dim ok1, ok2, ok3, match1, match2, match3 As Integer
        Dim err1 As String = "", err2 As String = "", err3 As String = ""
        For Each s In samples
            ' (1) 單參數 (預期跨 session 失敗, 當 baseline)
            Dim m1 As Redemption.RDOMail = Nothing
            Try
                m1 = TryCast(_rdo2.GetMessageFromID(s.eid), Redemption.RDOMail)
                If m1 IsNot Nothing Then ok1 += 1 : If m1.Subject = s.subj Then match1 += 1
            Catch ex As System.Exception
                If err1 = "" Then err1 = ex.Message
            Finally
                TryMarshalRelease(m1)
            End Try
            ' (2) 雙參數 + OOM StoreID
            Dim m2 As Redemption.RDOMail = Nothing
            Try
                m2 = TryCast(_rdo2.GetMessageFromID(s.eid, s.sid), Redemption.RDOMail)
                If m2 IsNot Nothing Then ok2 += 1 : If m2.Subject = s.subj Then match2 += 1
            Catch ex As System.Exception
                If err2 = "" Then err2 = ex.Message
            Finally
                TryMarshalRelease(m2)
            End Try
            ' (3) store-scoped (依 FolderPath 取 store 名, 在 _rdo2.Stores 找 RDOStore)
            Dim m3 As Redemption.RDOMail = Nothing
            Dim rstore As Redemption.RDOStore = Nothing
            Try
                Dim wantName As String = GetStoreNameFromPath(s.fpath)
                For i As Integer = 1 To _rdo2.Stores.Count
                    Dim cand As Redemption.RDOStore = _rdo2.Stores.Item(i)
                    If cand.Name = wantName Then rstore = cand : Exit For
                    TryMarshalRelease(cand)
                Next
                If rstore IsNot Nothing Then
                    m3 = TryCast(rstore.GetMessageFromID(s.eid), Redemption.RDOMail)
                    If m3 IsNot Nothing Then ok3 += 1 : If m3.Subject = s.subj Then match3 += 1
                Else
                    If err3 = "" Then err3 = $"_rdo2.Stores 找不到 [{wantName}]"
                End If
            Catch ex As System.Exception
                If err3 = "" Then err3 = ex.Message
            Finally
                TryMarshalRelease(m3)
                TryMarshalRelease(rstore)
            End Try
        Next

        ' ── 4. 總結 ──
        Dim n As Integer = samples.Count
        _dbg("探針結果 (1)單參數", $"resolve {ok1}/{n}, subject吻合 {match1}/{n}{If(err1 = "", "", " | err: " & err1)}")
        _dbg("探針結果 (2)雙參數+OOM StoreID", $"resolve {ok2}/{n}, subject吻合 {match2}/{n}{If(err2 = "", "", " | err: " & err2)}")
        _dbg("探針結果 (3)store-scoped", $"resolve {ok3}/{n}, subject吻合 {match3}/{n}{If(err3 = "", "", " | err: " & err3)}")
    End Function       ' 驗證獨立 session _rdo2 的 resolve 形式
    Private Async Function SpikeResolveFolderOnRdo2() As Task
        ' =================================================================
        ' 2026/06/23 by Simon/Claude Opus 4.8: 探針 — 驗證 _rdo2 的 FOLDER resolve 形式
        '   目的: 用 OOM 取得的 (folder EntryID, OOM StoreID, FolderPath) 在 _rdo2 上試三種 resolve,
        '         決定 GetMailCountRdo/GetFolderCountRdo 該走「store-scoped 單參數」還是「雙參數」。
        '   判讀: 看哪種 resolve 成功率高、且 .Name 對得上 (= 真解到 folder, 非空 handle)。
        '   ※ 純診斷, 不動 production; 用完即可整段刪除。(對照 SpikeResolveFormOnRdo2 的 message 版)
        ' =================================================================
        If _rdo2 Is Nothing Then Await InitRdoSessionWithoutEULA()
        If _rdo2 Is Nothing Then _dbg("Folder探針中止", "_rdo2 初始化失敗") : Return

        ' ── 1. 從 OOM 採樣 folder: 最多 3 個 PST、每 PST 最多 4 個夾, 合計上限 ~12 ──
        Dim samples As New List(Of (eid As String, sid As String, fpath As String, name As String))
        Dim storeTaken As Integer = 0
        For si As Integer = 1 To _olNS.Stores.Count
            If storeTaken >= 3 Then Exit For
            Dim st As Outlook.Store = Nothing
            Try
                st = _olNS.Stores.Item(si)
                If String.IsNullOrEmpty(st.FilePath) Then Continue For
                Dim grabbed As Integer = HarvestFoldersFromStore(st, st.StoreID, samples, 4)
                If grabbed > 0 Then storeTaken += 1
            Catch ex As System.Exception
                _dbg("Folder探針採樣", $"store#{si} 失敗: {ex.Message}")
            Finally
                TryMarshalRelease(st)
            End Try
        Next
        _dbg("Folder探針採樣", $"共取得 {samples.Count} 個夾 (跨 {storeTaken} 個 PST)")
        If samples.Count = 0 Then _dbg("Folder探針中止", "採樣 0 個夾") : Return

        ' ── 2. 三種形式逐夾測試 ──
        Dim ok1, ok2, ok3, match1, match2, match3 As Integer
        Dim err1 As String = "", err2 As String = "", err3 As String = ""
        For Each s In samples
            ' (1) 單參數 session 級 (預期跨 session 失敗, baseline)
            Dim f1 As Redemption.RDOFolder = Nothing
            Try
                f1 = TryCast(_rdo2.GetFolderFromID(s.eid), Redemption.RDOFolder)
                If f1 IsNot Nothing Then ok1 += 1 : If f1.Name = s.name Then match1 += 1
            Catch ex As System.Exception
                If err1 = "" Then err1 = ex.Message
            Finally
                Dim o As Object = f1 : TryMarshalRelease(o)
            End Try
            ' (2) 雙參數 + OOM StoreID
            Dim f2 As Redemption.RDOFolder = Nothing
            Try
                f2 = TryCast(_rdo2.GetFolderFromID(s.eid, s.sid), Redemption.RDOFolder)
                If f2 IsNot Nothing Then ok2 += 1 : If f2.Name = s.name Then match2 += 1
            Catch ex As System.Exception
                If err2 = "" Then err2 = ex.Message
            Finally
                Dim o As Object = f2 : TryMarshalRelease(o)
            End Try
            ' (3) store-scoped (依 FolderPath 取 store, store.GetFolderFromID(eid)) — production 目標路徑
            Dim f3 As Redemption.RDOFolder = Nothing
            Dim rstore As Redemption.RDOStore = GetRdoStore(s.fpath)
            Try
                If rstore IsNot Nothing Then
                    f3 = TryCast(rstore.GetFolderFromID(s.eid), Redemption.RDOFolder)
                    If f3 IsNot Nothing Then ok3 += 1 : If f3.Name = s.name Then match3 += 1
                Else
                    If err3 = "" Then err3 = $"GetRdo2Store 找不到 store for [{s.fpath}]"
                End If
            Catch ex As System.Exception
                If err3 = "" Then err3 = ex.Message
            Finally
                Dim o As Object = f3 : TryMarshalRelease(o)   ' rstore 為 byName 參考,不在此釋放
            End Try
        Next

        ' ── 3. 總結 ──
        Dim n As Integer = samples.Count
        _dbg("Folder探針 (1)單參數", $"resolve {ok1}/{n}, name吻合 {match1}/{n}{If(err1 = "", "", " | err: " & err1)}")
        _dbg("Folder探針 (2)雙參數+OOM StoreID", $"resolve {ok2}/{n}, name吻合 {match2}/{n}{If(err2 = "", "", " | err: " & err2)}")
        _dbg("Folder探針 (3)store-scoped", $"resolve {ok3}/{n}, name吻合 {match3}/{n}{If(err3 = "", "", " | err: " & err3)}")
    End Function     ' 驗證 _rdo2 的 FOLDER resolve 形式
    Private Function HarvestFromStore(st As Outlook.Store, sid As String, samples As List(Of (eid As String, sid As String, fpath As String, subj As String)), maxN As Integer) As Integer
        ' 探針輔助: 從單一 OOM store BFS 抓最多 maxN 封 (只讀 EntryID/Subject, 不碰 .Body/.Attachments 故不撞 IRM)
        Dim taken As Integer = 0
        Dim root As Outlook.Folder = Nothing
        Dim queue As New Queue(Of Outlook.Folder)()
        Try
            root = TryCast(st.GetRootFolder(), Outlook.Folder)
            If root Is Nothing Then Return 0
            queue.Enqueue(root) : root = Nothing      ' 交給 queue 統一釋放
            Dim visited As Integer = 0
            While queue.Count > 0 AndAlso taken < maxN AndAlso visited < 60
                Dim f As Outlook.Folder = queue.Dequeue()
                visited += 1
                Try
                    Dim items As Outlook.Items = f.Items
                    Dim cnt As Integer = items.Count
                    Dim fpath As String = f.FolderPath
                    Dim i As Integer = 1
                    While i <= cnt AndAlso taken < maxN
                        Dim it As Object = items.Item(i)
                        Try
                            Dim eid As String = CStr(CallByName(it, "EntryID", CallType.Get))
                            Dim subj As String = CStr(CallByName(it, "Subject", CallType.Get))
                            If Not String.IsNullOrEmpty(eid) Then samples.Add((eid, sid, fpath, subj)) : taken += 1
                        Catch
                            ' 非郵件項目或讀取失敗, 略過
                        Finally
                            TryMarshalRelease(it)
                        End Try
                        i += 1
                    End While
                    For sfi As Integer = 1 To f.Folders.Count : queue.Enqueue(f.Folders.Item(sfi)) : Next
                    TryMarshalRelease(items)
                Catch
                    ' 該夾讀取失敗, 略過
                Finally
                    TryMarshalRelease(f)
                End Try
            End While
        Catch ex As System.Exception
            _dbg("探針採樣", $"HarvestFromStore 失敗: {ex.Message}")
        Finally
            TryMarshalRelease(root)
            While queue.Count > 0 : TryMarshalRelease(queue.Dequeue()) : End While   ' 排空殘留子夾
        End Try
        Return taken
    End Function
    Private Function HarvestFoldersFromStore(st As Outlook.Store, sid As String, samples As List(Of (eid As String, sid As String, fpath As String, Name As String)), maxN As Integer) As Integer
        ' 探針輔助: 從單一 OOM store BFS 抓最多 maxN 個子夾 (只讀 EntryID/Name, 不碰 Items 故極輕量)
        Dim taken As Integer = 0
        Dim root As Outlook.Folder = Nothing
        Dim queue As New Queue(Of Outlook.Folder)()
        Try
            root = TryCast(st.GetRootFolder(), Outlook.Folder)
            If root Is Nothing Then Return 0
            queue.Enqueue(root) : root = Nothing
            Dim visited As Integer = 0
            While queue.Count > 0 AndAlso taken < maxN AndAlso visited < 60
                Dim f As Outlook.Folder = queue.Dequeue()
                visited += 1
                Try
                    Try
                        If Not String.IsNullOrEmpty(f.EntryID) Then samples.Add((f.EntryID, sid, f.FolderPath, f.Name)) : taken += 1
                    Catch
                    End Try
                    Dim subs As Outlook.Folders = f.Folders
                    Try
                        For Each sf As Outlook.Folder In subs
                            If queue.Count < 60 Then queue.Enqueue(sf) Else TryMarshalRelease(sf)
                        Next
                    Finally
                        TryMarshalRelease(subs)
                    End Try
                Finally
                    TryMarshalRelease(f)
                End Try
            End While
        Catch ex As System.Exception
            _dbg("HarvestFolders", $"{ex.Message}")
        End Try
        Return taken
    End Function

    Private Async Function SpikeFolderVisibilityCompare() As Task
        ' 探針一: SpikeFolderVisibilityCompare — RDO vs OOM 全枚舉夾清單差集 + 隱藏判據 dump
        ' 2026/06/23 by Simon/Claude Opus 4.8: 補 _rdoFastPath 的 visibility 技術債。
        '   目的: 找出 RDO 枚舉多撈、但 OOM 看不到的夾(實測曾 27 vs 22)，並 dump 其判據
        '         (Kind / PR_CONTAINER_CLASS / PR_ATTR_HIDDEN)，決定 isRDO 旗標的判斷規則。
        '   非破壞性: 只枚舉讀取，不寫任何快取、不改任何夾。測完可整段刪除。
        '   前提: 跑前把 Outlook 切到要測的 profile (Work 27 PST)。RDO 用獨立 _rdo2 不污染 _rdo。
        ' ════════════════════════════════════════════════════════════════════════
        Const PR_CONTAINER_CLASS As String = "http://schemas.microsoft.com/mapi/proptag/0x3613001E"
        Const PR_ATTR_HIDDEN As String = "http://schemas.microsoft.com/mapi/proptag/0x10F4000B"

        If _rdo2 Is Nothing Then Await InitRdoSessionWithoutEULA()  ' ← 若你的 _rdo2 初始化函數名不同，改這行
        If _rdo2 Is Nothing Then _dbg("VisCmp", "_rdo2 初始化失敗, 中止") : Return
        If _olNS Is Nothing Then _dbg("VisCmp", "_olNS 為空, 中止") : Return

        _dbg("VisCmp", "═════ RDO vs OOM 全枚舉差集 開始 ═════")

        Await Task.Run(
            Sub()
                ' ── 1. OOM 端: 逐 store BFS 枚舉 .Folders，收 FolderPath 集合 ──
                Dim oomPaths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                Dim oomByStore As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
                Try
                    For Each st As Outlook.Store In _olNS.Stores
                        Dim stName As String = "" : Try : stName = st.DisplayName : Catch : End Try
                        Dim root As Outlook.Folder = Nothing
                        Try : root = TryCast(st.GetRootFolder(), Outlook.Folder) : Catch : End Try
                        If root Is Nothing Then Continue For
                        Dim before As Integer = oomPaths.Count
                        Dim stk As New Stack(Of Outlook.Folder)() : stk.Push(root)
                        Do While stk.Count > 0
                            Dim f = stk.Pop()
                            Dim p As String = "" : Try : p = f.FolderPath : Catch : End Try
                            If p <> "" Then oomPaths.Add(p)
                            Try
                                For i As Integer = 1 To f.Folders.Count
                                    stk.Push(TryCast(f.Folders.Item(i), Outlook.Folder))
                                Next
                            Catch : End Try
                        Loop
                        oomByStore(stName) = oomPaths.Count - before
                    Next
                Catch ex As System.Exception
                    _dbg("VisCmp", "OOM 枚舉例外: " & ex.GetBaseException().Message)
                End Try
                _dbg("VisCmp", $"OOM 可見夾總數 = {oomPaths.Count}")
                For Each kv In oomByStore : _dbg(" │OOM", $"[{kv.Key}] {kv.Value} 夾") : Next

                ' ── 2. RDO 端(_rdo2): 逐 store BFS 枚舉 .Folders，收 FolderPath 集合 ──
                '    同時記下每夾的判據, 供差集 dump
                ' ── 2. RDO 端(_rdo2): 逐 store BFS 枚舉 .Folders，收 FolderPath 集合 ──
                '    2026/06/23 by Simon/Claude: 改用 IPMRootFolder(IPM 樹根)當起點。
                '      假設: search folder/系統夾在 IPM 樹外, 用 IPMRootFolder 枚舉天生不會撈到,
                '      集合應 = OOM 可見的 822。若差集歸零即證實「從源頭用 IPMRootFolder」可行。
                Dim rdoPaths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                Dim rdoInfo As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) ' path → 判據字串
                Try
                    For si As Integer = 1 To _rdo2.Stores.Count
                        Dim st = _rdo2.Stores.Item(si)
                        Dim stName As String = "" : Try : stName = st.Name : Catch : End Try
                        Dim root As Redemption.RDOFolder = Nothing
                        Try : root = st.IPMRootFolder : Catch : End Try    ' ← 改: RootFolder → IPMRootFolder
                        If root Is Nothing Then Continue For
                        Dim stk As New Stack(Of Redemption.RDOFolder)() : stk.Push(root)
                        Do While stk.Count > 0
                            Dim f = stk.Pop()
                            Dim p As String = "" : Try : p = f.FolderPath : Catch : End Try
                            If p <> "" Then
                                rdoPaths.Add(p)
                                ' dump 判據(讀法修正): Kind 直接取列舉轉 Integer, 不套 CallByName+CStr
                                Dim kind As String = "?"
                                Try : kind = CInt(f.Kind).ToString() : Catch : kind = "?" : End Try
                                Dim cclass As String = "" : Try : cclass = CStr(f.Fields(PR_CONTAINER_CLASS)) : Catch : cclass = "" : End Try
                                Dim hidden As String = "?" : Try : hidden = CStr(f.Fields(PR_ATTR_HIDDEN)) : Catch : hidden = "?" : End Try
                                rdoInfo(p) = $"Kind={kind}, Class=[{cclass}], Hidden={hidden}"
                            End If
                            Try
                                For i As Integer = 1 To f.Folders.Count
                                    stk.Push(f.Folders.Item(i))
                                Next
                            Catch : End Try
                        Loop
                    Next
                Catch ex As System.Exception
                    _dbg("VisCmp", "RDO 枚舉例外: " & ex.GetBaseException().Message)
                End Try
                _dbg("VisCmp", $"RDO(_rdo2, IPMRootFolder) 枚舉夾總數 = {rdoPaths.Count}")

                ' ── 3. 差集 ──
                Dim rdoOnly = rdoPaths.Where(Function(p) Not oomPaths.Contains(p)).OrderBy(Function(p) p).ToList()
                Dim oomOnly = oomPaths.Where(Function(p) Not rdoPaths.Contains(p)).OrderBy(Function(p) p).ToList()

                _dbg("VisCmp", $"═════ RDO-only(RDO有 OOM無) 共 {rdoOnly.Count} 個 ═════")
                For Each p In rdoOnly
                    Dim info As String = "" : rdoInfo.TryGetValue(p, info)
                    _dbg(" │RDO-only", $"{p}  ←  {info}")
                Next
                _dbg("VisCmp", $"═════ OOM-only(OOM有 RDO無) 共 {oomOnly.Count} 個 ═════")
                For Each p In oomOnly
                    _dbg(" │OOM-only", p)
                Next
                _dbg("VisCmp", "═════ 結束, 請貼回 RDO-only 清單與判據 ═════")
            End Sub)
    End Function ' RDO vs OOM 全枚舉夾清單差集 + 隱藏判據 dump
    Private Async Function SpikeFolderTableBenchmark() As Task
        ' 探針二: SpikeFolderTableBenchmark — 單夾 GetTable 的 OOM vs RDO parity + 分段計時
        '         + 平行 K=1/2/4 × {共用_rdo2 / 各自獨立session} 對照
        ' 2026/06/23 by Simon/Claude Opus 4.8: 改自 SpikeParallelReadBenchmark(P3)。
        '   回答三問: (A)單夾 GetTable 分段耗時瓶頸在哪 (B)RDO MAPITable 與 OOM GetTable 列數 parity
        '            (C)平行值不值得 + worker 共用一條 _rdo2 是否可行/掉速 vs 各自獨立 session(實測不預防)。
        '   非破壞性: 只讀不寫。測完可整段刪除。前提: Work profile, 已勾 CheckRDO 使 _rdo2 在。
        ' ════════════════════════════════════════════════════════════════════════
        Const PR_MESSAGE_SIZE As String = "http://schemas.microsoft.com/mapi/proptag/0x0E080003"
        Const PR_SENDER_EMAIL As String = "http://schemas.microsoft.com/mapi/proptag/0x0C1F001E"
        Const PR_INTERNET_MESSAGE_ID As String = "http://schemas.microsoft.com/mapi/proptag/0x1035001E"
        Dim cols = {"EntryID", "Subject", PR_MESSAGE_SIZE, "ReceivedTime", "SenderName", PR_INTERNET_MESSAGE_ID, PR_SENDER_EMAIL}

        If _rdo2 Is Nothing Then _dbg("TblBM", "_rdo2 為空(請先勾 CheckRDO), 中止") : Return
        If _olNS Is Nothing Then _dbg("TblBM", "_olNS 為空, 中止") : Return

        Dim profileName As String = ""
        Try : profileName = CStr(CallByName(_rdo2, "ProfileName", CallType.Get)) : Catch : End Try
        _dbg("TblBM", $"═════ 開始 (profile=[{profileName}]) ═════")

        ' ── 收集標的: OOM 走訪挑 >= MINROWS 封的夾, 收 (FolderPath, OOM Folder 物件) ──
        Const MINROWS As Integer = 500
        Const MAXFOLDERS As Integer = 8
        Dim targets As New List(Of (path As String, oomFolder As Outlook.Folder))()
        Try
            For Each st As Outlook.Store In _olNS.Stores
                If targets.Count >= MAXFOLDERS Then Exit For
                Dim root As Outlook.Folder = TryCast(st.GetRootFolder(), Outlook.Folder)
                If root Is Nothing Then Continue For
                Dim stk As New Stack(Of Outlook.Folder)() : stk.Push(root)
                Do While stk.Count > 0 AndAlso targets.Count < MAXFOLDERS
                    Dim f = stk.Pop()
                    Dim cnt As Integer = 0 : Try : cnt = f.Items.Count : Catch : End Try
                    If cnt >= MINROWS Then targets.Add((f.FolderPath, f))
                    Try
                        For i As Integer = 1 To f.Folders.Count : stk.Push(TryCast(f.Folders.Item(i), Outlook.Folder)) : Next
                    Catch : End Try
                Loop
            Next
        Catch : End Try
        If targets.Count = 0 Then _dbg("TblBM", $"找不到 >= {MINROWS} 封的夾, 中止") : Return
        _dbg("TblBM", $"標的夾 {targets.Count} 個 (每個 >= {MINROWS} 封)")

        ' ════ A: OOM GetTable 分段計時 ════
        _dbg("TblBM", "───── A: OOM GetTable 分段 ─────")
        For Each tg In targets
            Dim swPath As New Stopwatch(), swTable As New Stopwatch(), swArray As New Stopwatch()
            Dim rows As Integer = 0
            Try
                swPath.Start() : Dim p As String = tg.oomFolder.FolderPath : swPath.Stop()
                swTable.Start()
                Dim tbl As Outlook.Table = tg.oomFolder.GetTable("", Outlook.OlTableContents.olUserItems)
                tbl.Columns.RemoveAll()
                For Each c In cols : tbl.Columns.Add(c) : Next
                swTable.Stop()
                swArray.Start()
                Do While Not tbl.EndOfTable
                    Dim arr = tbl.GetArray(500)
                    If arr Is Nothing Then Exit Do
                    rows += arr.GetUpperBound(0) + 1
                Loop
                swArray.Stop()
            Catch ex As System.Exception
                _dbg(" │OOM", $"例外 {tg.path}: {ex.Message}") : Continue For
            End Try
            _dbg(" │OOM", $"[{ExtractFolderName(tg.path)}] rows={rows} | Path={swPath.ElapsedMilliseconds}ms Table={swTable.ElapsedMilliseconds}ms Array={swArray.ElapsedMilliseconds}ms")
        Next

        ' ════ B: RDO 列舉 Items(設 Columns 走 table 不開信) 分段計時 + parity ════
        _dbg("TblBM", "───── B: RDO 列舉 Items 分段 ─────")
        For Each tg In targets
            Dim swResolve As New Stopwatch(), swCols As New Stopwatch(), swRead As New Stopwatch()
            Dim rows As Integer = 0
            Try
                swResolve.Start()
                Dim rf As Redemption.RDOFolder = FolderPath2RdoFolder(_rdo2, tg.path)
                swResolve.Stop()
                If rf Is Nothing Then _dbg(" │RDO", $"解析失敗 {tg.path}") : Continue For

                Dim items As Redemption.RDOItems = rf.Items
                swCols.Start()
                ' 設 MAPITable.Columns: 設好後列舉 items 只讀這些欄、不開信 (官方 RDOItems 範例)
                Try
                    Dim mt As Object = items.MAPITable
                    mt.Columns.Clear()
                    For Each c In cols : mt.Columns.Add(c) : Next
                Catch exCol As System.Exception
                    _dbg(" │RDO", $"設 Columns 失敗 {tg.path}: {exCol.GetBaseException().Message}")
                End Try
                swCols.Stop()

                swRead.Start()
                For Each m As Redemption.RDOMail In items
                    Dim s As String = "" : Try : s = m.Subject : Catch : End Try   ' 觸發實際讀取(走 table)
                    rows += 1
                Next
                swRead.Stop()
            Catch ex As System.Exception
                _dbg(" │RDO", $"例外 {tg.path}: {ex.GetBaseException().Message}") : Continue For
            End Try
            _dbg(" │RDO", $"[{ExtractFolderName(tg.path)}] rows={rows} | Resolve={swResolve.ElapsedMilliseconds}ms Cols={swCols.ElapsedMilliseconds}ms Read={swRead.ElapsedMilliseconds}ms")
        Next


        ' ════ C: 平行 K=1/2/4 × {共用 _rdo2 / 各自獨立 session} ════
        _dbg("TblBM", "───── C: 平行對照 (workload=逐夾 列舉 Items 走 table) ─────")
        Dim allPaths = targets.Select(Function(t) t.path).ToList()
        For Each useShared In {True, False}
            Dim modeName As String = If(useShared, "共用_rdo2", "各自獨立session")
            For Each K In {1, 2, 4}
                Dim groups As New List(Of List(Of String))()
                For g = 0 To K - 1 : groups.Add(New List(Of String)) : Next
                For i = 0 To allPaths.Count - 1 : groups(i Mod K).Add(allPaths(i)) : Next

                Dim swWall As New Stopwatch() : swWall.Start()
                Dim tasks As New List(Of Task)()
                For g = 0 To K - 1
                    Dim myPaths = groups(g)
                    tasks.Add(Task.Run(
                        Sub()
                            Dim sess As Redemption.RDOSession = Nothing
                            Try
                                If useShared Then
                                    sess = _rdo2
                                Else
                                    sess = New Redemption.RDOSession()
                                    sess.Logon(profileName, "", False, True)
                                End If
                                For Each pth In myPaths
                                    Try
                                        Dim rf As Redemption.RDOFolder = FolderPath2RdoFolder(sess, pth)
                                        If rf Is Nothing Then Continue For
                                        Dim items As Redemption.RDOItems = rf.Items
                                        Try
                                            Dim mt As Object = items.MAPITable
                                            mt.Columns.Clear()
                                            For Each c In cols : mt.Columns.Add(c) : Next
                                        Catch : End Try
                                        For Each m As Redemption.RDOMail In items
                                            Dim s As String = "" : Try : s = m.Subject : Catch : End Try
                                        Next
                                    Catch : End Try
                                Next
                            Catch ex As System.Exception
                                _dbg(" │" & modeName, $"K={K} worker 例外: {ex.GetBaseException().Message}")
                            Finally
                                If Not useShared AndAlso sess IsNot Nothing Then
                                    Try : sess.Logoff() : Catch : End Try
                                    TryMarshalRelease(sess)
                                End If
                            End Try
                        End Sub))
                Next
                Await Task.WhenAll(tasks)
                swWall.Stop()
                _dbg(" │C", $"{modeName} K={K}: wall={swWall.ElapsedMilliseconds}ms ({allPaths.Count}夾)")
            Next
        Next
        _dbg("TblBM", "═════ 結束, 請貼回 ═════")

    End Function    ' 單夾 GetTable 的 OOM vs RDO parity + 分段計時
    Private Function FolderPath2RdoFolder(sess As Redemption.RDOSession, folderPath As String) As Redemption.RDOFolder
        ' ── 探針二專用小 helper: 在指定 session 上用 FolderPath 解出 RDOFolder ──
        ' 2026/06/23 by Simon/Claude: 拋棄式, 隨探針二刪除。
        '   策略: 先用 GetRdoStore 取 store(僅對 _rdo2 有效); 若傳入的是別條獨立 session,
        '   則退化為走訪該 session 的 Stores 找路徑開頭吻合者, 再 BFS 比對 FolderPath。
        Try
            ' 找 store: 路徑形如 \\store顯示名\夾\子夾, 取第一段比對 store.Name
            Dim trimmed As String = folderPath.TrimStart("\"c)
            Dim firstSeg As String = trimmed.Split("\"c)(0)
            Dim targetStore As Redemption.RDOStore = Nothing
            For si As Integer = 1 To sess.Stores.Count
                Dim st = sess.Stores.Item(si)
                Dim nm As String = "" : Try : nm = st.Name : Catch : End Try
                If String.Equals(nm, firstSeg, StringComparison.OrdinalIgnoreCase) Then targetStore = st : Exit For
            Next
            If targetStore Is Nothing Then Return Nothing
            ' 從 IPMRootFolder BFS 找 FolderPath 吻合
            Dim root As Redemption.RDOFolder = targetStore.IPMRootFolder
            Dim stk As New Stack(Of Redemption.RDOFolder)() : stk.Push(root)
            Do While stk.Count > 0
                Dim f = stk.Pop()
                Dim p As String = "" : Try : p = f.FolderPath : Catch : End Try
                If String.Equals(p, folderPath, StringComparison.OrdinalIgnoreCase) Then Return f
                Try
                    For i As Integer = 1 To f.Folders.Count : stk.Push(f.Folders.Item(i)) : Next
                Catch : End Try
            Loop
        Catch : End Try
        Return Nothing
    End Function ' 探針二專用小 helper: 在指定 session 上用 FolderPath 解出 RDOFolder

    ' 2026/06/24 by Simon/Claude Opus 4.8: 拋棄式探針 — 子樹階層走訪 OOM vs RDO批次 對拍
    '   本輪唯一目的: 先確認 API 讀法寫對 + 取得「暖快取」基準值(供 GetSubtreeRdo 完工後比對是否有額外開銷)。
    '   標的: SimTree3.SelectedNodes 當 root(可多選逐一各跑;Simon 自行換不同深淺節點重跑)。
    '   對手(全單執行緒,全產出「子孫 path 集合」對拍):
    '     A  OOM        : current.Folders 逐夾 BFS(= GetSubtreeOOM 去副作用版,基準)
    '     B  RDO-Enum   : rdoFolder.Folders For Each 逐夾(診斷: 隔離 RDO 層 vs OOM 層)
    '     C  RDO-Batch  : Folders.MAPITable.GetRows 整層批次,只對 PR_SUBFOLDERS=true 遞迴(候選)
    '     C+ RDO-Batch+CC: C 多撈 PR_CONTENT_COUNT(獨立計時,驗免費搭車且不污染 A/B/C)
    '   正確性對拍用 path 集合(最穩);EntryID 經 SpikeEidToHex 統一轉 hex 供遞迴。
    ' ============================================================================
    Private Sub SpikeSubtreeWalkCompare()
        Dim log As New List(Of String)
        If _rdo2 Is Nothing Then MessageBox.Show("_rdo2 未初始化,請先勾選 CheckRDO。") : Return
        Dim roots As List(Of TreeNode) = SimTree3.SelectedNodes
        If roots Is Nothing OrElse roots.Count = 0 Then MessageBox.Show("請先在 Tab3 的樹選定至少一個節點當 root。") : Return

        For Each node As TreeNode In roots
            Dim rootF As Folder = TryCast(node.Tag, Folder)
            If rootF Is Nothing Then Continue For
            Dim rootPath As String = SafeGetPath(rootF)
            Dim rootEid As String = "" : Try : rootEid = rootF.EntryID : Catch : End Try
            log.Add("══════ ROOT: " & ExtractFolderName(rootPath) & " ══════")
            log.Add("path = " & rootPath)

            Dim store As Redemption.RDOStore = GetRdoStore(rootPath)
            If store Is Nothing Then log.Add("✗ GetRdo2Store 失敗 → 跳過此 root 的 RDO 對手")

            ' ── 暖機一次(OOM)丟棄,讓後續對手吃同樣暖快取 ──
            Try : SpikeWalk_Oom(rootF, rootPath) : Catch : End Try

            Dim ra = SpikeWalk_Oom(rootF, rootPath)
            log.Add($"A  OOM        : {ra.paths.Count} 夾 | {ra.ms} ms")

            Dim rdoRoot As Redemption.RDOFolder = Nothing
            If store IsNot Nothing AndAlso rootEid <> "" Then
                Try : rdoRoot = TryCast(store.GetFolderFromID(rootEid), Redemption.RDOFolder)
                Catch ex As System.Exception : log.Add("✗ RDO root GetFolderFromID: " & ex.Message) : End Try
            End If

            If rdoRoot IsNot Nothing Then
                Try
                    Dim rb = SpikeWalk_RdoEnum(rdoRoot, rootPath)
                    log.Add($"B  RDO-Enum   : {rb.paths.Count} 夾 | {rb.ms} ms | vs A: {SpikeDiff(ra.paths, rb.paths)}")
                Catch ex As System.Exception : log.Add("✗ B 例外: " & ex.GetBaseException().Message) : End Try

                Try
                    Dim k As String = "?" : Dim cv As Integer = 0, ce As Integer = 0
                    Dim rc = SpikeWalk_RdoBatch(store, rdoRoot, rootPath, False, cv, ce, k)
                    log.Add($"C  RDO-Batch  : {rc.paths.Count} 夾 | {rc.ms} ms | vs A: {SpikeDiff(ra.paths, rc.paths)} | EntryID型別={k}")
                Catch ex As System.Exception : log.Add("✗ C 例外: " & ex.GetBaseException().Message) : End Try

                Try
                    Dim k As String = "?" : Dim cv As Integer = 0, ce As Integer = 0
                    Dim rcc = SpikeWalk_RdoBatch(store, rdoRoot, rootPath, True, cv, ce, k)
                    log.Add($"C+ RDO-Batch+CC: {rcc.paths.Count} 夾 | {rcc.ms} ms | PR_CONTENT_COUNT 有效={cv} 缺/錯={ce}")
                Catch ex As System.Exception : log.Add("✗ C+ 例外: " & ex.GetBaseException().Message) : End Try
            End If

            Dim o As Object = rdoRoot : TryMarshalRelease(o)
        Next

        For Each ln In log : _dbg("SubtreeSpike", ln) : Next
        MessageBox.Show(String.Join(vbCrLf, log), "子樹走訪對拍結果")
    End Sub
    Private Function SpikeWalk_Oom(rootF As Folder, rootPath As String) As (paths As HashSet(Of String), ms As Long)
        Dim paths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim sw = Stopwatch.StartNew()
        Dim q As New Queue(Of (f As Folder, p As String))() : q.Enqueue((rootF, rootPath))
        While q.Count > 0
            Dim cur = q.Dequeue()
            Dim subs As Folders = Nothing
            Try
                subs = cur.f.Folders
                For Each sf As Folder In subs
                    Dim nm As String = "" : Try : nm = sf.Name : Catch : Continue For : End Try
                    Dim cp As String = cur.p & "\" & nm : paths.Add(cp) : q.Enqueue((sf, cp))
                Next
            Catch : End Try
            If subs IsNot Nothing Then TryMarshalRelease(subs)
        End While
        sw.Stop() : Return (paths, sw.ElapsedMilliseconds)
    End Function
    Private Function SpikeWalk_RdoEnum(rdoRoot As Redemption.RDOFolder, rootPath As String) As (paths As HashSet(Of String), ms As Long)
        Dim paths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim toRel As New List(Of Object)()
        Dim sw = Stopwatch.StartNew()
        Dim q As New Queue(Of (f As Redemption.RDOFolder, p As String))() : q.Enqueue((rdoRoot, rootPath))
        Try
            While q.Count > 0
                Dim cur = q.Dequeue()
                Dim subs = cur.f.Folders
                Try
                    For Each sf As Redemption.RDOFolder In subs
                        Dim nm As String = "" : Try : nm = sf.Name : Catch : Continue For : End Try
                        Dim cp As String = cur.p & "\" & nm : paths.Add(cp) : q.Enqueue((sf, cp)) : toRel.Add(sf)
                    Next
                Catch : End Try
                TryMarshalRelease(subs)
            End While
        Finally
            For Each o In toRel : Dim oo As Object = o : TryMarshalRelease(oo) : Next
        End Try
        sw.Stop() : Return (paths, sw.ElapsedMilliseconds)
    End Function
    Private Function SpikeWalk_RdoBatch(store As Redemption.RDOStore, rdoRoot As Redemption.RDOFolder, rootPath As String,
                                        withCC As Boolean, ByRef ccValid As Integer, ByRef ccErr As Integer, ByRef eidKind As String) As (paths As HashSet(Of String), ms As Long)
        Const DASL_SUB As String = "http://schemas.microsoft.com/mapi/proptag/0x360A000B"  ' PR_SUBFOLDERS
        Const DASL_CC As String = "http://schemas.microsoft.com/mapi/proptag/0x36020003"   ' PR_CONTENT_COUNT
        Dim cols As String = If(withCC, $"Name, EntryID, {DASL_SUB}, {DASL_CC}", $"Name, EntryID, {DASL_SUB}")
        Dim paths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim toRel As New List(Of Object)()
        Dim sw = Stopwatch.StartNew()
        Dim q As New Queue(Of (f As Redemption.RDOFolder, p As String))() : q.Enqueue((rdoRoot, rootPath))
        Try
            While q.Count > 0
                Dim cur = q.Dequeue()
                Try
                    Dim foldersCol = cur.f.Folders  ' 推斷型別,避免猜 RDOFolders 名稱
                    Dim tbl = foldersCol.MAPITable  ' 推斷型別,避免猜 MAPITable 名稱
                    Dim rc As Integer = CInt(tbl.RowCount)
                    If rc > 0 Then
                        tbl.Columns = cols : tbl.GoToFirst()
                        Dim rowsArr As Array = DirectCast(tbl.GetRows(rc), Array)
                        For i As Integer = rowsArr.GetLowerBound(0) To rowsArr.GetUpperBound(0)
                            Dim row As Array = DirectCast(rowsArr.GetValue(i), Array)
                            Dim lb As Integer = row.GetLowerBound(0)
                            Dim vName = row.GetValue(lb) : Dim vEid = row.GetValue(lb + 1) : Dim vSub = row.GetValue(lb + 2)
                            If eidKind = "?" AndAlso vEid IsNot Nothing Then eidKind = vEid.GetType().Name
                            Dim nm As String = If(TypeOf vName Is String, CStr(vName), "")
                            Dim cp As String = cur.p & "\" & nm : paths.Add(cp)
                            If withCC Then
                                Dim vCc = row.GetValue(lb + 3)
                                If TypeOf vCc Is Integer Then ccValid += 1 Else ccErr += 1
                            End If
                            Dim hasSub As Boolean = If(TypeOf vSub Is Boolean, CBool(vSub), True)  ' 未知→保守遞迴
                            If hasSub Then
                                Dim eidHex As String = SpikeEidToHex(vEid)
                                If eidHex <> "" Then
                                    Dim child As Redemption.RDOFolder = TryCast(store.GetFolderFromID(eidHex), Redemption.RDOFolder)
                                    If child IsNot Nothing Then q.Enqueue((child, cp)) : toRel.Add(child)
                                End If
                            End If
                        Next
                    End If
                    TryMarshalRelease(tbl) : TryMarshalRelease(foldersCol)
                Catch ex As System.Exception
                    Throw New System.Exception($"RdoBatch@{ExtractFolderName(cur.p)}: {ex.GetBaseException().Message}")  ' 探針: 明確報錯不靜默
                End Try
            End While
        Finally
            For Each o In toRel : Dim oo As Object = o : TryMarshalRelease(oo) : Next
        End Try
        sw.Stop() : Return (paths, sw.ElapsedMilliseconds)
    End Function
    Private Function SpikeEidToHex(v As Object) As String
        If v Is Nothing Then Return ""
        If TypeOf v Is String Then Return CStr(v)
        If TypeOf v Is Byte() Then Return BitConverter.ToString(DirectCast(v, Byte())).Replace("-", "")
        If TypeOf v Is Array Then
            Dim a As Array = DirectCast(v, Array)
            Dim sb As New System.Text.StringBuilder(a.Length * 2)
            For k As Integer = a.GetLowerBound(0) To a.GetUpperBound(0) : sb.Append(Convert.ToByte(a.GetValue(k)).ToString("X2")) : Next
            Return sb.ToString()
        End If
        Return ""
    End Function
    Private Function SpikeDiff(a As HashSet(Of String), b As HashSet(Of String)) As String
        Dim ao = a.Where(Function(x) Not b.Contains(x)).Count()
        Dim bo = b.Where(Function(x) Not a.Contains(x)).Count()
        If ao = 0 AndAlso bo = 0 Then Return "一致✓"
        Return $"A獨有{ao}/此法獨有{bo}✗"
    End Function

    ' 2026/6/27 開始測試foldersize用的GetRows()和ExecSQL()
    Private Sub SpikeFolderSizeReadCompare()
        ' 2026/06/27 by Simon/Claude Opus 4.8 (v2): size 讀取法對拍 — 修 v1 三問題
        '   (1)選 PST root 本層 0 封,假性「一致」無意義 → cnt=0 直接標記跳過。
        '   (2)B 讀 PR_MESSAGE_SIZE_EXTENDED(PT_I8)經 GetRows 每封都回相同垃圾常數 ≈ -2^31 → 改讀 PR_MESSAGE_SIZE(PT_LONG,0x0E080003)。
        '   (3)ExecSQL SUM 實測 AV(REDEMP~2.DLL)→ 確認不可用,移除。
        '   對手: A OOM-GetArray(PT_I8,基準) vs B RDO-GetRows(PT_LONG)。parity ✗ 時自動 dump 前3列原始型別/值。測完即刪。
        ' ============================================================================
        Const PR_SIZE_I8 As String = "http://schemas.microsoft.com/mapi/proptag/0x0E080014"   ' PR_MESSAGE_SIZE_EXTENDED (PT_I8) — A 基準
        Const PR_SIZE_LONG As String = "http://schemas.microsoft.com/mapi/proptag/0x0E080003" ' PR_MESSAGE_SIZE (PT_LONG) — B 候選
        Dim log As New List(Of String)
        If _rdo2 Is Nothing Then MessageBox.Show("_rdo2 未初始化,請先勾選 CheckRDO。") : Return
        Dim roots As List(Of TreeNode) = SimTree3.SelectedNodes
        If roots Is Nothing OrElse roots.Count = 0 Then MessageBox.Show("請先在 Tab3 的樹選定資料夾。") : Return

        For Each node As TreeNode In roots
            Dim f As Folder = TryCast(node.Tag, Folder)
            If f Is Nothing Then Continue For
            Dim fPath As String = SafeGetPath(f)
            Dim eid As String = "" : Try : eid = f.EntryID : Catch : End Try
            log.Add("══════ 夾: " & ExtractFolderName(fPath) & " ══════")

            ' ── A: OOM 基準 (PT_I8,暖機一次再計時) ──
            Dim aSum As Long = 0, aCnt As Long = 0
            Try
                SpikeSizeOom(f, PR_SIZE_I8)
                Dim swA = Stopwatch.StartNew()
                Dim ra = SpikeSizeOom(f, PR_SIZE_I8) : swA.Stop()
                aSum = ra.sum : aCnt = ra.cnt
                log.Add($"A  OOM-GetArray : {aCnt} 封 | size={aSum} | {swA.ElapsedMilliseconds} ms")
            Catch ex As System.Exception
                log.Add("✗ A 例外: " & ex.GetBaseException().Message)
            End Try

            If aCnt = 0 Then log.Add("— 本層 0 封, 跳過 RDO 對拍(無意義)") : Continue For

            ' ── 解析 _rdo2 上的 rdoFolder ──
            Dim store As Redemption.RDOStore = GetRdoStore(fPath)
            Dim rdoFolder As Redemption.RDOFolder = Nothing
            If store IsNot Nothing AndAlso eid <> "" Then
                Try : rdoFolder = TryCast(store.GetFolderFromID(eid), Redemption.RDOFolder)
                Catch ex As System.Exception : log.Add("✗ RDO 解夾失敗: " & ex.Message) : End Try
            Else
                log.Add("✗ GetRdoStore=Nothing 或 eid 空 → 跳過 RDO 對手")
            End If

            If rdoFolder IsNot Nothing Then
                ' ── B: RDO GetRows (PT_LONG,暖機一次再計時) ──
                Try
                    SpikeSizeRdoGetRows(rdoFolder, PR_SIZE_LONG)
                    Dim swB = Stopwatch.StartNew()
                    Dim rb = SpikeSizeRdoGetRows(rdoFolder, PR_SIZE_LONG) : swB.Stop()
                    Dim ok As Boolean = (rb.sum = aSum AndAlso rb.cnt = aCnt)
                    log.Add($"B  RDO-GetRows  : {rb.cnt} 封 | size={rb.sum} | {swB.ElapsedMilliseconds} ms | vs A: {If(ok, "一致✓", $"✗(封{rb.cnt}/size{rb.sum})")}")
                    If Not ok Then
                        SpikeSizeRdoDumpCol(rdoFolder, "I8 0x0E080014", PR_SIZE_I8, log)
                        SpikeSizeRdoDumpCol(rdoFolder, "LONG 0x0E080003", PR_SIZE_LONG, log)
                    End If
                Catch ex As System.Exception
                    log.Add("✗ B 例外: " & ex.GetBaseException().Message)
                End Try

                Dim o As Object = rdoFolder : TryMarshalRelease(o)
            End If
        Next

        For Each ln In log : _dbg("SizeSpike", ln) : Next
        MessageBox.Show(String.Join(vbCrLf, log), "單夾大小讀取對拍結果 (v2)")
    End Sub
    Private Sub SpikeSizeRdoDumpCol(rdoFolder As Redemption.RDOFolder, label As String, col As String, log As List(Of String))
        ' 診斷: parity 失敗時 dump 前3列在指定欄的原始 TypeName + 值,看 GetRows 到底回什麼
        Dim items = rdoFolder.Items
        Dim tbl = items.MAPITable
        Try
            tbl.Columns = col : tbl.GoToFirst()
            Dim chunk As Array = TryCast(tbl.GetRows(3), Array)
            If chunk IsNot Nothing Then
                For i As Integer = chunk.GetLowerBound(0) To chunk.GetUpperBound(0)
                    Dim row As Array = TryCast(chunk.GetValue(i), Array)
                    Dim v As Object = If(row IsNot Nothing, row.GetValue(row.GetLowerBound(0)), Nothing)
                    log.Add($"   dump[{label}] row{i}: TypeName={TypeName(v)} | val={v}")
                Next
            End If
        Catch ex As System.Exception
            log.Add($"   dump[{label}] 失敗: {ex.GetBaseException().Message}")
        Finally
            TryMarshalRelease(tbl) : TryMarshalRelease(items)
        End Try
    End Sub
    Private Function SpikeSizeOom(f As Folder, prSize As String) As (sum As Long, cnt As Long)
        ' OOM 基準: 同 GetFolderSizeOOM ① 的 GetArray 迴圈(去 SmartThrottle/progress)
        Dim total As Long = 0, n As Long = 0
        Dim table As Outlook.Table = Nothing
        Try
            table = SafeGetTable(f, "", prSize)
            Do
                Dim data = SafeGetArray(table)
                If data Is Nothing Then Exit Do
                For r As Integer = 0 To data.GetUpperBound(0)
                    n += 1
                    Dim sz = data(r, 0)
                    If sz IsNot Nothing AndAlso Not IsDBNull(sz) Then total += CLng(sz)
                Next
            Loop
        Finally
            TryMarshalRelease(table)
        End Try
        Return (total, n)
    End Function
    Private Function SpikeSizeRdoGetRows(rdoFolder As Redemption.RDOFolder, prSize As String) As (sum As Long, cnt As Long)
        ' 候選: rdoFolder.Items.MAPITable 設單欄 PR_MESSAGE_SIZE_EXTENDED,GoToFirst 後分批 GetRows(5000) 加總
        Dim total As Long = 0, n As Long = 0
        Dim items = rdoFolder.Items
        Dim tbl = items.MAPITable
        Try
            tbl.Columns = prSize
            tbl.GoToFirst()
            Do
                Dim chunk As Array = TryCast(tbl.GetRows(5000), Array)
                If chunk Is Nothing Then Exit Do
                Dim got As Integer = 0
                For i As Integer = chunk.GetLowerBound(0) To chunk.GetUpperBound(0)
                    got += 1 : n += 1
                    Dim row As Array = TryCast(chunk.GetValue(i), Array)
                    If row Is Nothing Then Continue For
                    Dim v = row.GetValue(row.GetLowerBound(0))
                    If v IsNot Nothing AndAlso Not IsDBNull(v) Then total += CLng(v)
                Next
                If got < 5000 Then Exit Do   ' 最後一批不足 → 到底
            Loop
        Finally
            TryMarshalRelease(tbl) : TryMarshalRelease(items)
        End Try
        Return (total, n)
    End Function
    Private Sub SpikeSizeRdoExecSql(rdoFolder As Redemption.RDOFolder, log As List(Of String))
        ' 診斷: ExecSQL 能否下推 SUM。先 COUNT(*)(過去已確證)驗 ExecSQL 通,再試兩種 SUM 欄名寫法。回 ADODB.Recordset(晚繫結)。
        Dim items = rdoFolder.Items
        Dim tbl = items.MAPITable
        Try
            Try
                Dim rs As Object = tbl.ExecSQL("SELECT COUNT(*) FROM Folder")
                Dim cv As Object = If(rs IsNot Nothing AndAlso Not CBool(rs.EOF), rs.Fields(0).Value, Nothing)
                log.Add($"C1 ExecSQL COUNT(*) = {cv}  (ExecSQL 可用✓)")
            Catch ex As System.Exception
                log.Add("✗ C1 ExecSQL COUNT(*) 失敗: " & ex.GetBaseException().Message)
            End Try
            For Each col As String In {"""http://schemas.microsoft.com/mapi/proptag/0x0E080014""", "PR_MESSAGE_SIZE"}
                Try
                    Dim rs As Object = tbl.ExecSQL($"SELECT SUM({col}) FROM Folder")
                    Dim sv As Object = If(rs IsNot Nothing AndAlso Not CBool(rs.EOF), rs.Fields(0).Value, Nothing)
                    log.Add($"C2 ExecSQL SUM({col}) = {sv}  ✓")
                Catch ex As System.Exception
                    log.Add($"✗ C2 ExecSQL SUM({col}) 失敗: " & ex.GetBaseException().Message)
                End Try
            Next
        Finally
            TryMarshalRelease(tbl) : TryMarshalRelease(items)
        End Try
    End Sub

#End Region


End Class
