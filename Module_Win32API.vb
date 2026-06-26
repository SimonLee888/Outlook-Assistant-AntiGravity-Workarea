Imports System.Collections.Concurrent
Imports System.Threading
Imports Microsoft.Office.Interop

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
    Private Function GetMailSizeL3(item As Object) As Long
        ' --------------------------------------------------------------
        ' GetMailSizeL3: 讀取單封郵件的大小 (bytes)，供 GetFolderSizeL3 fallback 路徑呼叫
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
        ' 注意: 此函數接受 Object 型別參數，是因為 GetFolderSizeL3 的 fallback 路徑
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
#End Region


End Class
