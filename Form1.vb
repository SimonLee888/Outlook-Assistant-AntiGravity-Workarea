Imports System
Imports System.Threading
Imports System.Globalization
Imports System.Collections.Concurrent
Imports System.Runtime.InteropServices
Imports System.Windows.Forms.DataVisualization.Charting
Imports Microsoft.Office.Interop.Outlook
Imports Outlook = Microsoft.Office.Interop.Outlook
Imports Redemption  ' 2026/3/22 正式導入Redemption, 測試logon成功, 傳回數值成功

'Imports MailKit
'Imports MailKit.Search
'Imports System.Core.dll
'Imports System.ComponentModel.Design.ObjectSelectorEditor
'Imports System.Diagnostics.Metrics
'Imports System.DirectoryServices.ActiveDirectory
'Imports System.Linq.Parallel.dll
'Imports System.Net
'Imports System.Reflection
'Imports System.Windows.Forms.VisualStyles.VisualStyleElement
'Imports System.Windows.Controls
'Imports System.ComponentModel
'Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Header
'Imports System.Threading
'Imports Microsoft.VisualBasic.Devices
'Imports Windows.Graphics.Printing.OptionDetails
'Imports Windows.Security.Authentication.Identity.Core
'Imports Exception = Microsoft.Office.Interop.Outlook.Exception

Public Class Form1

#Region "■ 01 全域宣告"
#Region "  ├ Win32 API 宣告"

    ' ── 函數宣告 ────────────────────────────────────────────────────────────────
    ' 統一使用 DllImport（取代舊式 Declare Function）
    ' 2026-03-23 整理：移除重複的 SendMessage Declare 版本，補齊 FindWindow / FindWindowEx

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
    Private Shared Function SetWindowPos(
        ByVal hWnd As IntPtr,
        ByVal hWndInsertAfter As IntPtr,
        ByVal x As Integer,
        ByVal y As Integer,
        ByVal cx As Integer,
        ByVal cy As Integer,
        ByVal uFlags As Integer) As Boolean
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
    Private Shared Function ShowWindow(ByVal hWnd As IntPtr, ByVal nCmdShow As Integer) As Boolean
    End Function

    <Runtime.InteropServices.DllImport("user32.dll")>
    Private Shared Function LockWindowUpdate(
        ByVal hWnd As IntPtr) As Boolean
    End Function

    ' ── 常數 ────────────────────────────────────────────────────────────────────
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

    Private Const SWP_NOZORDER As Integer = &H4                     ' debugForm resize用
    Private Const SWP_NOACTIVATE As Integer = &H10                  ' debugForm resize用

    Private Const SWP_NOREDRAW As Integer = &H8                     ' debugForm resize 時閃爍, 改手動redraw
    Private Const RDW_INVALIDATE As Integer = &H1                   ' debugForm resize 時閃爍, 改手動redraw
    Private Const RDW_UPDATENOW As Integer = &H100                  ' debugForm resize 時閃爍, 改手動redraw
    Private Const RDW_ALLCHILDREN As Integer = &H80                 ' debugForm resize 時閃爍, 改手動redraw
    Private Const WM_SETREDRAW As Integer = &HB                     ' 2026/3/26 by AntiGravity

    ' ↓ 新增 (2026-03-20) ListView1 進入資料夾用
    Private Const TVM_SELECTITEM As Integer = &H110B                ' = &H1100 + 11
    Private Const TVGN_CARET As Integer = &H9                       ' SendMessage 選取 Treeview 游標節點
    Private Const LVM_SETITEMCOUNT As Integer = &H1000 + 47         ' = &H102F '

#End Region
#Region "  ├ 成員變數與快取"
    <System.Diagnostics.Conditional("DEBUG")>
    Private Sub Dbg(Optional msg As String = "", Optional detail As String = "")
        Dim realCaller As String = WhoCallsMe()         ' 這裡強制傳入真正的呼叫者，跳過 Dbg() 這一層' 下面會定義
        If _isDebugMode Then DebugForm.AddMessage3(msg, detail, realCaller)   ' ← 傳第 3 個參數
    End Sub
    Private Function WhoCallsMe() As String        ' 尋找呼叫來源的輔助函數（放在 Dbg 旁邊即可）
        Dim st As New StackTrace(2, True)   ' 從第 2 層開始（跳過 Dbg + Conditional）
        For i As Integer = 0 To st.FrameCount - 1
            Dim m = st.GetFrame(i).GetMethod()
            If m IsNot Nothing AndAlso
                m.DeclaringType IsNot Nothing AndAlso
                m.DeclaringType.Name <> "DebugForm" AndAlso
                Not m.Name.Contains("Dbg") AndAlso
                Not m.Name.Contains("MoveNext") Then
                Return $"{m.DeclaringType.Name}.{m.Name}"
            End If
        Next
        Return "Unknown Method Call:"
    End Function

    Private WithEvents _olApp As Outlook.Application = Nothing
    Private _olNS As Outlook.NameSpace = Nothing
    Private _rdo As Redemption.RDOSession = Nothing ' _rdoSession 就等同是outlook.namespace 的意思, 就是Redemption的MAPI session
    Private _pstStoreList As List(Of Outlook.Store) = Nothing
    ' ■ Redemption 共用 Session（Form 層級，只初始化一次）
    ' 2026-03-22 新增：用於測試 Redemption.dll 整合 (注意：session.MAPIOBJECT 必須在 Outlook MAPI 連線建立後才能設定（Form1_Load 尾端）
    '------------------------------------------------------------------------------------------------
    ' Outlook 物件(OOM)	    Redemption 物件 (RDO)     說明
    '------------------------------------------------------------------------------------------------
    ' Outlook.Application	Redemption 本體	        Redemption 是底層 MAPI 封裝，它不負責 UI 或視窗管理。
    ' Outlook.NameSpace	    Redemption.RDOSession	最接近。 負責管理登入、StoreID、PST 檔案庫與全域設定。
    ' Outlook.Folder	    Redemption.RDOFolder	對應資料夾層級。
    ' Outlook.MailItem	    Redemption.RDOMail	    對應單封郵件層級。
    ' Outlook.Store	        Redemption.RDOStore	    對應 PST 或 Exchange 帳戶。

    Private _isFirstInit As Boolean = True          ' 第一次啟動程式
    Private _isTab3_Stop As Boolean                 ' 搜尋附件的Tab3/Button3按下ESC中斷
    Private _isDebugMode As Boolean
    Private _cancelRequested As Boolean = False     ' todo: ESC 全域中斷旗標：Tab1/Tab2/Tab3 共用，按 ESC 立刻設 True，各操作在 Yield 點檢查
    Private _cacheSnifferCts As New System.Threading.CancellationTokenSource  ' B4 CacheSniffer 取消令牌，FormClosing 時呼叫 Cancel()
    Private sw0, sw1, sw2, sw3, sw4, sw5, sw6 As New Stopwatch

    '2026/3/10重構時停止使用全域變數來記錄遞迴過程中的資料, 改用傳遞參數以避免多線程或重入呼叫時資料被改寫的問題
    'Private _intTotalMailCount As Integer          ' 在遞迴中, 記錄點選資料夾內的所有郵件總數, 不要被遞迴呼叫改變數量
    'Private _intProcessedCount As Integer          ' 在遞迴中, 加總已處理的郵件總數, 不要被遞迴呼叫改變數量

    Private _lastHoveredTreeNode As TreeNode = Nothing
    Private _lastHoveredListItem As ListViewItem = Nothing
    Private _lastHoveredPointIndex As Integer = -1                  ' 記住上一個 hover 的點，-1 表示沒有

    Private _tab1SelectSeq As Integer = 0                           ' Tab1 快速點選防護序號
    Private _tab2FolderList As List(Of Outlook.Folder) = Nothing    ' 記住目前 Tab2 的資料夾清單，供月份展開使用
    Private _tab2IsMonthView As Boolean = False                     ' 目前 ListView2 顯示的是月份視圖還是年度視圖
    Private _tab2MonthViewYear As Integer = 0                       ' 目前月份視圖顯示的是哪一年

    Private WithEvents SimTree1 As New SimTree
    Private WithEvents SimTree2 As New SimTree
    Private WithEvents SimTree3 As New SimTree
    Private WithEvents SimTree4 As New SimTree
    Private _ctxListView1 As ContextMenuStrip                       ' ContextMenu 成員變數，只初始化一次，不在每次右鍵時重新建立
    Private _ctxTreeView2 As ContextMenuStrip
    Private _ctxSimTree2 As ContextMenuStrip

    Private Shared ReadOnly _mailCountCache As New ConcurrentDictionary(Of Outlook.Folder, Integer)
    Private Shared ReadOnly _mailSizeCache As New ConcurrentDictionary(Of Outlook.MailItem, Integer)
    Private Shared ReadOnly _folderCountCache As New ConcurrentDictionary(Of Outlook.Folder, Integer)
    Private Shared ReadOnly _folderSizeCache As New ConcurrentDictionary(Of Outlook.Folder, Long)
    Private Shared ReadOnly _folderTreeCache As New ConcurrentDictionary(Of Outlook.Folder, List(Of Outlook.Folder))
    Private Shared ReadOnly _yearCountsCache As New ConcurrentDictionary(Of String, ConcurrentDictionary(Of Integer, Integer))
    Private Shared ReadOnly _monthCountsCache As New ConcurrentDictionary(Of String, ConcurrentDictionary(Of Integer, Integer))
#End Region
#Region "  └ 其他全域結構"
    Private Class FolderBfsEntry
        ' 候選待掃瞄剪枝的資料夾結構
        Public Folder As Outlook.Folder
        Public ParentIndex As Integer       ' -1 = rootFolder；>= 0 = 父節點在 allEntries 的索引
        Public DirectMailCount As Integer   ' 本層郵件數 (不含子孫) ，由 L3 填入
        Public TotalMailCount As Integer    ' 含子孫郵件總數，L2 底部向上彙總後填入
        Public TotalSubCount As Integer     ' 含子孫資料夾總數，L2 底部向上彙總後填入
        Public IsFromCache As Boolean       ' True = TotalMailCount/TotalSubCount 從快取取得，子樹已剪枝
    End Class

    ' ── Tab3 Phase1 快取 ─────────────────────────────────────────────
    ' 設計：快取「hasattachment 全集 (無大小篩選) 」，大小條件改在 LINQ 記憶體過濾
    ' 好處：相同資料夾換不同大小條件時直接命中快取，不重跑 GetTable
    ' 失效條件：folder.Items.Count (PR_CONTENT_COUNT) 改變
    ' key：FolderPath 字串 (不用 COM 物件，穩定不受 RCW 影響) 
    ' 2026-03-16 B1 新增
    Private _tab3Phase1Cache As New Dictionary(Of String, FolderCacheTab3)
    Private Structure FolderCacheTab3
        Dim mailWithAttachment As List(Of MailItemInfo)         ' 所有 hasattachment 候選 (無大小篩選) 
        Dim ItemCountWhenCached As Integer                      ' 快取當下的 PR_CONTENT_COUNT，失效偵測用
    End Structure
    Private Structure MailItemInfo
        ' 候選郵件的純資料結構 (不帶 COM 物件，不受 GC 影響)
        Dim EntryID As String
        Dim Subject As String
        Dim Size As Long
        Dim ReceivedTime As DateTime
        Dim SenderName As String
        Dim AttachCount As Integer
    End Structure

    ' 定義排序方式的列舉
    Private currentSortOrder As SortOrder = SortOrder.Ascending     ' 設置初始排序方式為升序
    Private previousColumnIndex As Integer = -1                     ' 儲存上一次點選的列索引
    Public Class ListViewItemComparer ' 用於比較 ListView 項目並依Column 進行排序
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
#End Region
#End Region

#Region "■ 02 Form 生命週期 & 外觀初始化"
#Region "  ├ 表單行為及輔助函數"
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        _isDebugMode = True
        dbg("開始：", sender.Name)
        ' 開始計時
        Dim stopwatch As New Stopwatch() : stopwatch.Start()

        ' 檢查系統中是否已經啟動 Outlook
        Dim processes() As Process = Process.GetProcessesByName("OUTLOOK")
        If processes.Length = 0 Then    ' 如果 Outlook 尚未啟動，顯示訊息並關閉應用程式
            MessageBox.Show("請先啟動 Outlook", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information) : End
        Else
            Try
                _olApp = New Outlook.Application        ' 初始化Outlook物件模型
                _olNS = _olApp.GetNamespace("MAPI")     ' 初始化outlook.nameSpace = MAPI session
                If _olApp IsNot Nothing AndAlso _olNS IsNot Nothing Then _pstStoreList = GetSortedStores(_olNS)  ' 取得所有 PST 檔
            Catch ex As System.Exception
                Dbg("Outlook App OR NameSpace init FAIL", ex.Message)
                _olApp = Nothing : _olNS = Nothing
                MessageBox.Show("Outlook Object 連接失敗!") : End
            End Try

            Try ' ── Redemption Session 初始化, 2026-03-22 測試用：
                '_rdo = New Redemption.RDOSession()  ' _rdoSession 就等同是outlook.namespace 的意思, 就是Redemption的MAPI session
                '_rdo.MAPIOBJECT = _olNS.MAPIOBJECT  ' 直接attach 到現有 Outlook MAPI session, 就不會另開視窗, 也不會另外生出不同的ol.app或ol.ns (必須在 objNameSpace 已建立之後才呼叫)
                'Dbg("Redemption init OK", $"Version={_rdo.Version}") ' 關鍵：不建新連線，直接接管現有的 Outlook MAPI session, 這樣就不會彈出第二個 Outlook 視窗，也不需要另外登入
                Dim unused = InitRedemptionSessionWithoutDeclaration()
            Catch ex As System.Exception
                Dbg("Redemption init FAIL", ex.Message)
                _rdo = Nothing
            End Try
        End If

        ' 設定程式標題 
        Dim strApp As String = My.Application.Info.DirectoryPath & "\" & My.Application.Info.ProductName & ".EXE"
        If My.Computer.FileSystem.FileExists(strApp) Then
            Dim infoReader As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(strApp)
            Me.Text = "Outlook Assistant - by Simon Lee Studio (build " & infoReader.LastWriteTime.Year & Format(infoReader.LastWriteTime, "MMdd.HHmmss") & ")"

            '' todo: 如何設置版本號自動遞增
            'Dim myApp = My.Application.Info.Version
            'Dim strVer As String = myApp.Major & ", " & myApp尸.Minor & ", " & myApp.MajorRevision & ", " & myApp.MinorRevision & ", " & myApp.Build & ", " & My.Application.Info.AssemblyName
            'myApp.MinorRevision += 1
            'myApp.Build += 1
            'Me.Text = Me.Text & strVer
        End If

        ' ✅ DebugForm 也丟背景，不等它建立完
        '   原本 CheckDebug 的 CheckedChanged 會 Show DebugForm，
        '   改成直接 Task.Run 建立，不卡 UI
        ' todo: 讓debugform自動上色, 可多選, 正確減去時間差
        Task.Run(Sub()
                     Me.Invoke(Sub()
                                   If CheckDebug.Checked Then DebugForm.Show()
                               End Sub)
                 End Sub)

        Me.KeyPreview = True    ' ✅ 讓 Form 優先攔截 ESC，否則 ESC 會先被 TreeView/ListBox 等子控制項消耗
        Cursor = Cursors.AppStarting
        LookAndFeel()   ' 設計程式外觀
        Me.BringToFront()
        Me.Show()       ' 先將表單顯示後, 再以背景執行緒加入資料夾, 提高操作反應速度

        ' PST檔太多, 啟動速度愈來愈差, 5/17全部重寫, 依照20年前的做法動態載入:
        ' 只先載入第一層, 若第二層還有subFolders則暫加一個假的":::", 讓它能顯示"+"加號表示還有子資料夾就好
        ' todo: 讓claude.ai 重構, 真正做到 lazy loading, 只有當使用者點開 "+" 號展開節點時, 才真正去讀取該資料夾的子資料夾, 而不是一開始就全部讀取進來
        ' ✅ LoadStoreToTreeView 是目前最大耗時點
        ' todo: 如果 PST 數量多，可以只先載入第一個 PST，其餘用 BeginInvoke 延遲載入 (或用RDO載入?)
        ' todo: 第一次formload的時候, 好像RDO 一直還沒init 完? 都是走MAPI??
        LoadStoreToTreeView(_pstStoreList, TreeView1)
        ExpandTreeToDefaultInbox(TreeView1)

        ' 啟動完成, 停止計時, 顯示總共花費的時間
        stopwatch.Stop() : lblStatus2.Text = "程式啟動完成花費了 " & stopwatch.Elapsed.TotalSeconds.ToString("0.00") & " 秒。"
        Cursor = Cursors.Default

        ' todo: 使用BackgroundWorker元件, 在還沒點選前偷偷在背景計算foldercount, mailcount並存入cache
        ' todo: 啟動timer, 在背景偷偷預讀foldercounts --> mailcounts --> foldersize
        ' todo: 使用BackgroundWorker元件, 在還沒點選前偷偷在背景計算foldersize並存入cache.
        'folderQueue = New Queue(Of TreeNode)(QueueAllFolderNodes(TreeView1))
        'tmrPreCache.Interval = 2000 : tmrPreCache.Enabled = True

        ' ✅ 2026-03-16 B4 CacheSniffer：啟動完成後才 fire-and-forget，不阻塞 UI
        'CacheSnifferAsync(_cacheSnifferCts.Token)  '裡面內建了等待 10 秒的延遲，確保 Form1_Load 完全結束、UI 呈現完畢，再開始佔用 Outlook COM

        dbg("結束：", sender.Name)
        CheckDebug.Checked = False  ' 先預設開啟Debug視窗, 以便在開發測試階段觀察各個事件的觸發狀況跟參數值, 等正式發佈了再預設關閉
    End Sub
    Private Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        ' ── ESC 全域中斷 ──────────────────────────────────────────────
        ' KeyPreview=True 讓 Form 優先攔截 KeyDown，子控制項不會先吃掉 ESC
        ' Tab1: ComputeFolderStatsAsync 在 Yield 點檢查 _cancelRequested → 回空 List
        ' Tab2: ComputeYearCounts  在 For Each 頭部檢查 → Exit For 回傳已算部分
        ' Tab3: 複用既有的 isTab3_Stop 旗標，不重複設計
        If e.KeyCode = Keys.Escape Then
            _cancelRequested = True

            Button3.Enabled = True
            _isTab3_Stop = True
            Button3_Stop.Visible = False

            Cursor = Cursors.Default
            lblStatus1.Text = "已中斷。"
            e.Handled = True
        End If
    End Sub
    Private Sub Form1_Move(sender As Object, e As EventArgs) Handles Me.Move
        ' 視窗移動時同步 DebugForm — 2026/3/26 by AntiGravity
        SyncDebugFormPosition()
    End Sub
    Private Sub Form1_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        ' 視窗縮放時同步 DebugForm — 2026/3/26 by AntiGravity
        SyncDebugFormPosition()

        ' 原本的 ListView1 寬度調整邏輯
        For Each c In ListView1.Columns
            c.Width = ListView1.Width * 0.168
        Next
        ListView1.Columns(0).Width = ListView1.Width * 0.32

        If TabControl1.SelectedTab Is TabPage3 Then
            Button3_Stop.Location = Button3.Location            ' 把stop按鈕跟button3重疊但不可見, 按下button3的查詢期間才visible
            GroupBox3.Visible = SplitContainer3.Width >= 1100   ' Group3的附件個數篩選平常看不到, 拉開寬度才出現
        End If

        'SimTree2.Width = TreeView1.Width
    End Sub
    Private Sub Form1_ResizeBegin(sender As Object, e As EventArgs) Handles Me.ResizeBegin
        Dbg("開始：", sender.Width & "x" & sender.Height)
    End Sub
    Private Sub Form1_ResizeEnd(sender As Object, e As EventArgs) Handles Me.ResizeEnd
        Dbg("結束：", sender.Width & "x" & sender.Height)
    End Sub
    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        _cacheSnifferCts.Cancel()   ' ✅ 2026-03-16 B4: 通知 CacheSniffer 停止，避免程式關閉後 COM 呼叫繼續進行

        ' 釋放所有的 COM 物件占用資源
        If _pstStoreList IsNot Nothing Then
            For Each store In _pstStoreList : Marshal.FinalReleaseComObject(store) : Next
            _pstStoreList.Clear() : _pstStoreList = Nothing
        End If

        If _olApp IsNot Nothing Then Marshal.FinalReleaseComObject(_olApp)
        If _olNS IsNot Nothing Then Marshal.FinalReleaseComObject(_olNS)
        If _rdo IsNot Nothing Then Marshal.FinalReleaseComObject(_rdo)

    End Sub
    Private Sub SyncDebugFormPosition()
        ' 同步 Debug 視窗與主視窗的位置與大小，並將其右側貼齊螢幕邊緣
        ' 使用 SetWindowPos 避免多個屬性分別設定導致的閃爍
        ' 2026/3/26 by AntiGravity
        If DebugForm IsNot Nothing AndAlso (DebugForm.Visible OrElse CheckDebug.Checked) Then
            Dim newLeft As Integer = Me.Left + Me.Width - 12
            Dim newTop As Integer = Me.Top
            Dim newHeight As Integer = Me.Height

            ' 計算螢幕工作區右側邊緣，並延展 DebugForm 寬度填滿剩餘空間
            Dim screenRight = Screen.FromControl(Me).WorkingArea.Right
            Dim newWidth = screenRight - newLeft
            If newWidth < 100 Then newWidth = DebugForm.Width ' 保底寬度

            ' 2026/3/26 by AntiGravity: 改用 WM_SETREDRAW 取代 LockWindowUpdate，解決黑影閃爍問題
            SendMessage(DebugForm.Handle, WM_SETREDRAW, New IntPtr(0), IntPtr.Zero)
            Try
                SetWindowPos(DebugForm.Handle, IntPtr.Zero, newLeft, newTop, newWidth, newHeight,
                             SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_NOREDRAW)
            Finally
                SendMessage(DebugForm.Handle, WM_SETREDRAW, New IntPtr(1), IntPtr.Zero)
                RedrawWindow(DebugForm.Handle, IntPtr.Zero, IntPtr.Zero,
                             RDW_INVALIDATE Or RDW_UPDATENOW Or RDW_ALLCHILDREN)
            End Try
        End If
    End Sub
#End Region
#Region "  ├ 物件初始化"
    Private Sub LookAndFeel()
        ' === 初始化共用物件的外觀及共通行為 ===
        ' 2026-03-17 C2 拆分：TreeView / ListView 各司其職的外觀設定移到獨立函數
        '   ApplyTreeViewStyles() ← TreeView / SimTree 字型、顏色、雙緩衝
        '   ApplyListViewStyles() ← ListView 字型、基本樣式、雙緩衝、欄位定義
        '   LookAndFeel()         ← 視窗位置、TabControl、ContextMenu、Chart2、Button、雜項
        Dbg("開始：")

        ' ── 視窗位置與背景色 ──
        Me.BackColor = Color.FromArgb(242, 242, 242)
        If Screen.FromControl(Me).Bounds.Height > 2560 Then
            Me.Top = Screen.FromControl(Me).Bounds.Height * 0.45                '如果在直立式的4K螢幕上啟動, 就把表單放在下半部往上移5%
            Me.Left = (Screen.FromControl(Me).Bounds.Width - Me.Width) * 0.45   '不管在什麼解析度的螢幕上啟動, 都把表單放在螢幕中央往左移5%
        End If

        ' ── 各控制項外觀 (委派給獨立函數) ──
        Dim defaultFont As New Font("Microsoft Jhenghei", 10, System.Drawing.FontStyle.Regular)
        InitTreeViews(defaultFont)
        InitListViews(defaultFont)
        InitTab5UI()

        ' ── TabControl 字型與分頁名稱 ──
        Dim strTabName As String() = {"資料夾統計", "依日期統計", "尋找附件", "尋找系列郵件", "尋找重覆郵件"}
        TabControl1.Font = defaultFont
        For i As Integer = 0 To strTabName.Length - 1
            TabControl1.TabPages(i).Text = strTabName(i)
        Next

        ' ── Chart2 樣式 ──
        CheckSub2.BringToFront()
        Chart2.Width = Chart2.Parent.ClientSize.Width
        Chart2.Height = Chart2.Parent.ClientSize.Height - ListView2.Height - 33  '調整Chart2的高度以適應父容器的高度變化, 33是預留的間距
        Chart2.BackColor = Color.FromArgb(242, 242, 242)
        Chart2.BorderlineDashStyle = ChartDashStyle.Solid
        Chart2.BorderlineColor = Color.FromArgb(224, 224, 224)
        Chart2.ChartAreas(0).BackColor = Color.FromArgb(242, 242, 242)
        Chart2.ChartAreas(0).AxisX.MajorGrid.LineColor = Color.FromArgb(224, 224, 224)
        Chart2.ChartAreas(0).AxisY.MajorGrid.LineColor = Color.FromArgb(224, 224, 224)

        ' ── 最大化 ChartArea 和 InnerPlotPosition ──
        ' ChartArea.Position：ChartArea 在整個 Chart 控制項中的佔比 (單位：%) 
        ' ✅ 讓 ChartArea 幾乎填滿整個 Chart 控制項 (上下左右各留 1%) ' 預設約 Position(5,5,90,90)，壓縮到幾乎填滿整個 Chart 控制項
        Chart2.ChartAreas(0).Position = New ElementPosition(1, 1, 99, 99)

        ' ✅ InnerPlotPosition：ChartArea 內部長條圖實際繪製區的佔比
        '    Auto=True 時 Chart 會自動縮排給軸標籤留空，通常左側縮 10~15%
        '    改成 Auto=False 並手動指定，讓左側縮排符合實際 Y 軸標籤寬度
        With Chart2.ChartAreas(0).InnerPlotPosition
            .Auto = False
            .X = 1          ' 左側留 1% (給 Y 軸數字標籤) 
            .Y = 2          ' 上方留 2%
            .Width = 90     ' 往右延伸 90%
            '.Height = 94    ' 往下延伸 94% (底部留 6% 給 X 軸標籤) 
            .Height = 100
            Chart2.Legends(0).Enabled = False   ' 這樣 Height 可以再放大到 92~94, 如果數字標籤被截到，就微調 X 和 Height 的值
        End With

        ' ── Button3 樣式 ──
        Button3.FlatStyle = FlatStyle.System
        Button3.FlatAppearance.BorderColor = Color.FromArgb(0, 120, 212)
        Button3.FlatAppearance.MouseOverBackColor = Color.FromArgb(224, 224, 224)
        Button3.ForeColor = Color.FromArgb(0, 120, 212)

        Dbg("結束：")
    End Sub
    Private Sub InitTreeViews(defaultFont As Font)
        ' ── TreeView / SimTree：字型、顏色、縮排、雙緩衝 ──
        ' TreeView / SimTree 都在 SplitContainer 內，Me.Controls 的 For Each 看不到它們，因此直接逐一設定，不用遍歷
        For Each tv As TreeView In {TreeView1, TreeView2, TreeView3, TreeView4}
            tv.Font = defaultFont
            tv.BackColor = Color.White
            tv.ForeColor = SystemColors.InactiveCaptionText
            tv.Indent = 20
        Next
        For Each st As SimTree In {SimTree1, SimTree2}
            st.Font = defaultFont
            st.BackColor = Color.White
            st.ForeColor = SystemColors.InactiveCaptionText
            st.Indent = 20
            st.Anchor = TreeView1.Anchor
        Next

        SplitContainer2.Panel1.Controls.Add(SimTree2)
        With SimTree2
            .Top = TreeView1.Top : .Height = TreeView1.Height
            .Left = TreeView1.Left : .Width = TreeView1.Width
            '.BringToFront()
        End With

        TreeView2.Top = TreeView1.Height / 5 * 4
        TreeView2.Height = TreeView1.Height / 5 * 1
        'TreeView2.BringToFront()
        SimTree2.BringToFront()

        ' 雙緩衝：解決 MouseMove hover 換色閃爍
        SendMessage(TreeView1.Handle, TVM_SETEXTENDEDSTYLE, New IntPtr(TVS_EX_DOUBLEBUFFER), New IntPtr(TVS_EX_DOUBLEBUFFER))
        SendMessage(TreeView2.Handle, TVM_SETEXTENDEDSTYLE, New IntPtr(TVS_EX_DOUBLEBUFFER), New IntPtr(TVS_EX_DOUBLEBUFFER))
        SendMessage(TreeView3.Handle, TVM_SETEXTENDEDSTYLE, New IntPtr(TVS_EX_DOUBLEBUFFER), New IntPtr(TVS_EX_DOUBLEBUFFER))
        SendMessage(TreeView4.Handle, TVM_SETEXTENDEDSTYLE, New IntPtr(TVS_EX_DOUBLEBUFFER), New IntPtr(TVS_EX_DOUBLEBUFFER))

        SendMessage(SimTree1.Handle, TVM_SETEXTENDEDSTYLE, New IntPtr(TVS_EX_DOUBLEBUFFER), New IntPtr(TVS_EX_DOUBLEBUFFER))
        SendMessage(SimTree2.Handle, TVM_SETEXTENDEDSTYLE, New IntPtr(TVS_EX_DOUBLEBUFFER), New IntPtr(TVS_EX_DOUBLEBUFFER))

        '' ── ContextMenu (只建立一次，不在每次右鍵時重複 AddHandler) ──
        '_ctxTreeView2 = New ContextMenuStrip()
        'Dim menuItem1 As New ToolStripMenuItem("切換多選模式")
        'AddHandler menuItem1.Click, AddressOf MenuItem1_Click : _ctxTreeView2.Items.Add(menuItem1)

        '_ctxSimTree2 = New ContextMenuStrip()
        'Dim menuItem2 As New ToolStripMenuItem("切換單選模式")
        'AddHandler menuItem2.Click, AddressOf MenuItem2_Click : _ctxSimTree2.Items.Add(menuItem2)

    End Sub

    Private WithEvents rbExactMatch As New RadioButton()
    Private WithEvents rbFuzzyMatch As New RadioButton()
    Private WithEvents ListView5 As New ListView()

    Private Sub InitTab5UI()
        ' 清除原有的測試控制項 (如果有) ，並移出 ListView5 到 TabPage5 下
        TabPage5.Controls.Clear()

        rbExactMatch.Text = "完全相同 (主旨+大小+時間+寄件者)"
        rbExactMatch.Location = New Point(20, 20)
        rbExactMatch.Checked = True
        rbExactMatch.AutoSize = True

        rbFuzzyMatch.Text = "相似重複 (相似主旨+大小)"
        rbFuzzyMatch.Location = New Point(320, 20)
        rbFuzzyMatch.AutoSize = True

        Button5.Location = New Point(600, 15)
        Button5.Text = "開始掃描"
        Button5.AutoSize = True

        ListView5.Location = New Point(20, 60)
        ListView5.Size = New Size(TabPage5.ClientSize.Width - 40, TabPage5.ClientSize.Height - 80)
        ListView5.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
        ListView5.View = System.Windows.Forms.View.Details
        ListView5.FullRowSelect = True
        ListView5.GridLines = True
        SendMessage(ListView5.Handle, LVM_SETEXTENDEDLISTVIEWSTYLE, New IntPtr(LVS_EX_DOUBLEBUFFER), New IntPtr(LVS_EX_DOUBLEBUFFER))

        ListView5.Columns.Add("主旨", CInt(TabPage5.Width * 0.4))
        ListView5.Columns.Add("大小", 80).TextAlign = HorizontalAlignment.Right
        ListView5.Columns.Add("收到時間", 150)
        ListView5.Columns.Add("寄件者", 150)
        ListView5.Columns.Add("重複群體", 80)
        ListView5.Columns.Add("EntryID", 0)

        TabPage5.Controls.Add(rbExactMatch)
        TabPage5.Controls.Add(rbFuzzyMatch)
        TabPage5.Controls.Add(Button5)
        TabPage5.Controls.Add(ListView5)
    End Sub

    Private Sub InitListViews(defaultFont As Font)
        ' ── ListView：字型、基本樣式、雙緩衝、欄位定義 ── (其他 ListView 欄位在 Designer 定義) 
        For Each lv As ListView In {ListView1, ListView2, ListView3, ListView4}
            lv.Font = New Font("Microsoft Jhenghei", 10, System.Drawing.FontStyle.Regular)
            lv.View = System.Windows.Forms.View.Details
            lv.Width = lv.Parent.ClientSize.Width
            lv.FullRowSelect = True
            lv.GridLines = False
            ' ✅ 2026-03-17：原本 For Each lv in Me.Controls, 根本碰不到 ListView (因為是在 SplitContainer 內) ，
            ' 拆分後直接指名才真正生效，誤顯示格線，改回 False 恢復原有外觀
        Next

        ' 雙緩衝, 解決大量資料時的滾動和更新閃爍問題
        SendMessage(ListView1.Handle, LVM_SETEXTENDEDLISTVIEWSTYLE, New IntPtr(LVS_EX_DOUBLEBUFFER), New IntPtr(LVS_EX_DOUBLEBUFFER))
        SendMessage(ListView2.Handle, LVM_SETEXTENDEDLISTVIEWSTYLE, New IntPtr(LVS_EX_DOUBLEBUFFER), New IntPtr(LVS_EX_DOUBLEBUFFER))
        SendMessage(ListView3.Handle, LVM_SETEXTENDEDLISTVIEWSTYLE, New IntPtr(LVS_EX_DOUBLEBUFFER), New IntPtr(LVS_EX_DOUBLEBUFFER))
        SendMessage(ListView4.Handle, LVM_SETEXTENDEDLISTVIEWSTYLE, New IntPtr(LVS_EX_DOUBLEBUFFER), New IntPtr(LVS_EX_DOUBLEBUFFER))

        ListView1.Columns.Clear()
        Dim headerNames As String() = {"資料夾名稱", "郵件數量", "資料夾數量", "郵件總計", "資料夾大小"}
        For Each n In headerNames
            ListView1.Columns.Add(n, n)
            ListView1.Columns(n).Width = ListView1.Width * 0.168
        Next
        ListView1.Columns("資料夾名稱").Width = ListView1.Width * 0.32
        ListView1.Columns("資料夾名稱").TextAlign = HorizontalAlignment.Left
        For Each c In ListView1.Columns
            c.TextAlign = HorizontalAlignment.Right
        Next

        _ctxListView1 = New ContextMenuStrip()  ' ── ContextMenu (只建立一次，不在每次右鍵時重複 AddHandler) ──

        Dim enterFolderMenuItem As New ToolStripMenuItem("進入資料夾")
        AddHandler enterFolderMenuItem.Click, AddressOf Me.EnterFolderMenuItem
        _ctxListView1.Items.Add(enterFolderMenuItem)

        Dim showFolderSizeMenuItem As New ToolStripMenuItem("統計資料夾大小")
        AddHandler showFolderSizeMenuItem.Click, AddressOf ListView1_ItemMenu
        _ctxListView1.Items.Add(showFolderSizeMenuItem)

        ' ── ListView4：系列郵件欄位定義 ──
        ListView4.Columns.Clear()
        Dim lv4Names As String() = {"主旨", "大小", "收到時間", "寄件者", "EntryID"}
        For Each n In lv4Names
            ListView4.Columns.Add(n, n)
        Next
        ListView4.Columns("主旨").Width = ListView4.Width * 0.4
        ListView4.Columns("大小").Width = ListView4.Width * 0.15
        ListView4.Columns("大小").TextAlign = HorizontalAlignment.Right
        ListView4.Columns("收到時間").Width = ListView4.Width * 0.2
        ListView4.Columns("寄件者").Width = ListView4.Width * 0.2
        ListView4.Columns("EntryID").Width = 0 ' 隱藏欄位

    End Sub
    Private Async Function InitRedemptionSessionWithoutDeclaration() As Task
        ' 2026-03-23 v3：
        '   Task.Run 包裝保留（讓 UI 執行緒繼續跑 LoadStoreToTreeView，平行初始化）
        '   第一次執行競爭條件改用 Thread.Sleep(1) 在 Set() 前解決，
        '   確保 AutoDismiss 輪詢 loop 已執行第一次再放行 New RDOSession()
        Try
            If _rdo IsNot Nothing Then Return

            Dim threadStarted As New System.Threading.ManualResetEventSlim(False)
            AutoDismissRedemptionDialog(threadStarted)

            ' 等 AutoDismiss thread 確認輪詢已開始，最多等 500ms
            threadStarted.Wait(500)
            Dbg("InitRedemption", "AutoDismiss thread 已就緒，開始 New RDOSession")

            ' ✅ Task.Run：UI 執行緒不阻塞，LoadStoreToTreeView 可以同時跑
            Dim session As Redemption.RDOSession = Nothing
            Await Task.Run(Sub()
                               session = New Redemption.RDOSession()
                           End Sub)

            ' MAPIOBJECT 必須回 UI 執行緒賦值（_olNS 是 STA COM 物件）
            session.MAPIOBJECT = _olNS.MAPIOBJECT
            _rdo = session
            Dbg("Redemption init OK", $"Version={_rdo.Version}")

        Catch ex As System.Exception
            _rdo = Nothing
            Dbg("Redemption init FAIL", ex.Message)
        End Try
    End Function
    Private Async Sub CacheSnifferAsync(ct As System.Threading.CancellationToken)
        ' === CacheSniffer — 背景快取預讀系統 (B4) ===
        ' ===============================================================================
        ' 職責：程式啟動後在背景靜默預讀 Tab1 / Tab2 / Tab3 ，快取後讓使用者點選時直接從記憶體讀取，不再等待 COM 查詢。
        '
        ' 設計原則：
        '   - 廣度優先 (BFS) ：淺層資料夾優先預讀，使用者最常點選的位置最先就緒
        '   - 固定 1 秒間隔：每完成一個資料夾的三項快取，固定等 1 秒再繼續，讓 Outlook 有充足空閒時間回應使用者互動
        '   - COM 全在 UI 執行緒 (STA) ：所有 Await 都不切執行緒，不需要 Task.Run
        '   - CancellationToken：FormClosing 時呼叫 _cacheSnifferCts.Cancel()，確保程式關閉後不留殘餘 COM 呼叫
        '   - 快取命中就跳過：若使用者已先點選觸發過快取，CacheSniffer 直接略過不重做
        '   - 停用方式：把 Form1_Load 末尾的 CacheSnifferAsync(...) 那行加上 ' 即可，其餘程式碼完全不受影響
        '
        ' 預讀順序 (每個資料夾) ：
        '   1. Tab1：mailCountCache + folderCountCache (GetMailCountAll / GetTotalFolderCountAsync) 
        '   2. Tab2：yearCountsCache (GetYearCountsForFolderAsync) 
        '   3. Tab3：_tab3Phase1Cache (CheckTab3CacheOrRescan) 
        '
        ' 2026-03-16 B4 新增，由 PrewarmAllCachesAsync 重構整合，改名為 CacheSniffer
        ' ===============================================================================

        If _pstStoreList Is Nothing OrElse _pstStoreList.Count = 0 Then Return
        Await Task.Delay(10000, ct)      ' 等待 10 秒：確保 Form1_Load 完全結束、UI 呈現完畢，再開始佔用 Outlook COM
        Try
            Dbg("CacheSniffer: 開始預讀")

            ' ── BFS 初始化：把所有 PST 的第一層子資料夾加進佇列 ─────────
            ' 不從 root 本身開始，因為 root ("個人資料夾") 通常不含郵件，
            ' 直接從第一層子資料夾 (收件匣、寄件匣…) 開始
            Dim queue As New Queue(Of Outlook.Folder)
            For Each store As Outlook.Store In _pstStoreList
                If ct.IsCancellationRequested Then Return
                For Each subFolder As Outlook.Folder In GetSortedSubFolders(store.GetRootFolder())
                    queue.Enqueue(subFolder)
                Next
            Next

            ' ── BFS 主迴圈 ───────────────────────────────────────────────
            ' 每次取出一個資料夾，依序預讀 Tab1 / Tab2 / Tab3 的快取，
            ' 完成後把它的直屬子資料夾再放入佇列 (廣度優先，淺層先完成) 
            Dim processed As Integer = 0
            While queue.Count > 0
                If ct.IsCancellationRequested Then Return
                Dim folder As Outlook.Folder = queue.Dequeue()
                processed += 1

                ' ── Tab1：mailCountCache + folderCountCache ───────────────
                ' GetMailCountAll 和 GetTotalFolderCountAsync 內部各自寫入自己的快取
                ' 已命中的快取直接跳過，不重複呼叫 COM
                Try
                    If Not _mailCountCache.ContainsKey(folder) Then Await GetMailCountAll(folder)
                    If Not _folderCountCache.ContainsKey(folder) Then Await GetFolderCountAll(folder)
                Catch ex As System.Exception
                    Dbg("CacheSniffer Tab1 Error: ", folder.Name & " - " & ex.Message)
                End Try
                If ct.IsCancellationRequested Then Return

                ' ── Tab2：yearCountsCache ─────────────────────────────────
                ' GetYearCountsForFolderAsync 內部有快取命中判斷，已快取直接回傳
                Try
                    Dim key As String = folder.FolderPath
                    If Not _yearCountsCache.ContainsKey(key) Then Await GetYearCountsForFolder(folder)
                Catch ex As System.Exception
                    Dbg("CacheSniffer Tab2 Error: ", folder.Name & " - " & ex.Message)
                End Try
                If ct.IsCancellationRequested Then Return

                ' ── Tab3：_tab3Phase1Cache ────────────────────────────────
                ' CheckTab3CacheOrRescan 內部有 Items.Count 失效判斷
                Try
                    Await CheckTab3CacheOrRescan(folder)
                Catch ex As System.Exception
                    Dbg("CacheSniffer Tab3 Error: ", folder.Name & " - " & ex.Message)
                End Try
                If ct.IsCancellationRequested Then Return

                ' ── 固定 1 秒間隔：讓 Outlook 保持回應能力 ───────────────
                Dbg($"CacheSniffer: [{processed}] {folder.Name} 完成，等 1 秒")
                Await Task.Delay(1000, ct)
                Await Task.Yield()

                ' ── 把直屬子資料夾加入佇列 (廣度優先) ────────────────────
                ' GetSortedSubFolders 有 folderTreeCache，不重打 COM
                Try
                    For Each subFolder As Outlook.Folder In GetSortedSubFolders(folder)
                        queue.Enqueue(subFolder)
                    Next
                Catch ex As System.Exception
                    Dbg("CacheSniffer subfolder Error: ", folder.Name & " - " & ex.Message)
                End Try
            End While

            Dbg($"CacheSniffer: 完成，共預讀 {processed} 個資料夾")

        Catch ex As System.Threading.Tasks.TaskCanceledException
            Dbg("CacheSniffer: 已取消 (FormClosing) ")
        Catch ex As System.Exception
            Dbg("CacheSniffer Error: ", ex.Message)
        End Try
    End Sub
#End Region
#Region "  └ LazyLoad動態載入"
    Private Function GetSortedStores(space As Outlook.NameSpace) As List(Of Outlook.Store)
        ' ==========================================
        ' 取得排序後的 NameSpace 下所有Outlook.Store
        ' 包含目前config內的所有帳號和所有開啟的PST檔
        ' ==========================================
        Dbg("開始：", space.CurrentProfileName)

        ' 使用 Task.Run 將同步操作包裝在獨立的工作線程中 (違反STA安全, 不再使用)
        'Await Task.Run(Sub() stores = space.Stores.Cast(Of Outlook.Store)().ToList())

        ' 遍歷所有Outlook.Store並添加到列表中, 使用LINQ擴充方法就夠快了, 不再使用非同步或Parallel.Foreach了
        Dim stores As List(Of Outlook.Store) = space.Stores.Cast(Of Outlook.Store)().ToList()

        ' 使用 LINQ 排序Outlook.Store
        stores = stores.OrderBy(Function(s) If(TextHasChineseChar(s.DisplayName), 1, 0)).ThenBy(Function(s) s.DisplayName).ToList()
        Return stores
        ' ⚠️ 注意：不在這裡 ReleaseComObject(space)
        ' space 就是外層的 objNameSpace，釋放後其他地方 (Tab2/Tab3 等) 再用 objNameSpace 會觸發 RCW 已釋放的例外
        ' objNameSpace 的生命週期就只由 Form1_FormClosing 統一管理

    End Function
    Private Function GetSortedSubFolders(folder As Outlook.Folder) As List(Of Outlook.Folder)
        ' ==========================================
        ' 取得引數folder下的所有subFolders並排序後傳回
        ' ==========================================
        Dbg("開始：", folder.Name)

        'debug: 首次切到tab2的時候, Gmail_2022不會展開, 因為沒有填入subFolders, 所以ExpandTreeToDefaultInbox()找不到預設的收件匣 
        If _folderTreeCache.ContainsKey(folder) Then Return _folderTreeCache(folder)  ' 若快取中有資料則直接回傳不再讀取

        ' 使用 LINQ 擴充方法將 Folders 集合轉換為 List(Of Outlook.Folder)
        ' 2024/5/13記錄: 已經試過很多種優化, 好像很難再比現在下面這二行LINQ還快了??
        Dim subFolders As List(Of Outlook.Folder) = folder.Folders.Cast(Of Outlook.Folder)().ToList()

        ' 使用 LINQ 排序資料夾
        subFolders = subFolders.OrderBy(Function(subFolder) If(TextHasChineseChar(subFolder.Name), 1, 0)).
                                ThenBy(Function(subFolder) subFolder.Name).ToList()
        _folderTreeCache(folder) = subFolders    '存入快取
        Return subFolders
    End Function
    Private Function GetSubFolderList(rootFolder As Outlook.Folder, includeSubFolders As Boolean) As List(Of Outlook.Folder)
        ' --------------------------------------------------------------
        ' GetSubFolderList：取得目標資料夾清單 (BFS，含子資料夾)
        ' ① OOM BFS：目前唯一的路徑，使用 Outlook Object Model 廣度優先搜尋
        ' --------------------------------------------------------------
        Dbg("開始：GetSubFolderList", rootFolder.Name)
        Dim sw As New Stopwatch() : sw.Start()

        Dim result As New List(Of Outlook.Folder)
        result.Add(rootFolder)
        If Not includeSubFolders Then
            sw.Stop()
            Dbg("結束：GetSubFolderList (Single)", $"{rootFolder.Name} | {sw.ElapsedMilliseconds}ms")
            Return result     ' 若不包含子資料夾，直接回傳只有 rootFolder 的清單
        End If

        ' 取得目標資料夾清單 (BFS，含子資料夾)
        Dim queue As New Queue(Of Outlook.Folder)
        queue.Enqueue(rootFolder)
        While queue.Count > 0
            Dim current As Outlook.Folder = queue.Dequeue()
            Try
                For Each subFolder As Outlook.Folder In current.Folders
                    result.Add(subFolder)       ' 把子資料夾加入結果清單
                    queue.Enqueue(subFolder)    ' 把子資料夾加入佇列，繼續往下搜尋
                Next
            Catch ex As System.Exception
                Dbg("GetSubFolderList ① OOM 失敗", current.Name & " - " & ex.Message)
            End Try
        End While

        sw.Stop()
        Dbg("結束：GetSubFolderList (BFS)", $"{rootFolder.Name} | folders={result.Count} | {sw.ElapsedMilliseconds}ms")
        Return result
    End Function

    ' --------------------------------------------------------------
    ' 2026/3/24 by AntiGravity: GetSubFolderList_RDO
    ' 目的：專門提供給 RDO 平行路徑使用，回傳 List(Of Redemption.RDOFolder)
    ' 說明：因為 Redemption 是 free-threaded，可以用 Parallel.ForEach 安全平行展開子樹
    ' --------------------------------------------------------------
    Private Function GetSubFolderList_RDO(rootFolder As Redemption.RDOFolder, includeSubFolders As Boolean) As List(Of Redemption.RDOFolder)
        Dbg("開始：GetSubFolderList_RDO", rootFolder.Name)
        Dim sw As New Stopwatch() : sw.Start()

        Dim resultBag As New ConcurrentBag(Of Redemption.RDOFolder)
        resultBag.Add(rootFolder)
        If Not includeSubFolders Then
            sw.Stop()
            Dbg("結束：GetSubFolderList_RDO (Single)", $"{rootFolder.Name} | {sw.ElapsedMilliseconds}ms")
            Return resultBag.ToList()
        End If

        ' 使用兩層佇列作層級遍歷，每層用 Parallel.ForEach 探索
        Dim currentLayer As New ConcurrentQueue(Of Redemption.RDOFolder)
        currentLayer.Enqueue(rootFolder)

        Do
            Dim layerList = currentLayer.ToList()
            If layerList.Count = 0 Then Exit Do

            ' 清空 queue 準備裝下一層的資料夾
            Do While currentLayer.TryDequeue(Nothing) : Loop

            ' 平行處理當前層的資料夾，將它們的子資料夾加進 queue 與結果中
            Parallel.ForEach(layerList,
                Sub(current)
                    Try
                        For Each subFolder As Redemption.RDOFolder In current.Folders
                            resultBag.Add(subFolder)
                            currentLayer.Enqueue(subFolder)
                        Next
                    Catch ex As System.Exception
                        Dbg("GetSubFolderList_RDO Error: ", current.Name & " - " & ex.Message)
                    End Try
                End Sub)
        Loop

        sw.Stop()
        Dbg("結束：GetSubFolderList_RDO (Parallel BFS)", $"{rootFolder.Name} | folders={resultBag.Count} | {sw.ElapsedMilliseconds}ms")
        Return resultBag.ToList()
    End Function
    Private Sub LoadStoreToTreeView(storeList As List(Of Outlook.Store), treeview As TreeView)
        Dbg("開始：", treeview.Name)

        ' 2024/5/17全部重寫, 只先動態載入一層的rootFolder, 不花時間遍歷所有的subFolders
        ' 2024/5/19試過Task.Run(), Parallel.Foreach跟LINQ擴充方法了, 都沒有比較快, 別再試了, 就算virtual mode也沒有比我現在的lazy load還快
        'treeview.BeginUpdate()
        'For Each store In storeList
        '    Dim root As Outlook.Folder = store.GetRootFolder
        '    'Dim node As TreeNode = Await Task.Run(Function() Me.Invoke(Function() treeview.Nodes.Add(root.Name)))
        '    Dim node As TreeNode = treeview.Nodes.Add(root.Name)
        '    node.Tag = root
        '    If root.Folders.Count > 0 Then node.Nodes.Add(":::") '若發現底下還有subFolders也不讀取, 只先填入一個假的":::"暫代, 才能出現"+"號
        'Next
        'treeview.EndUpdate()

        ' 2024/5/20昨天才說不會更快了, 今天改用Nodes.AddRange(), 又更快了一點, 連BeginUpdate/EndUpdate都不需要了
        ' 遍歷 storeList 並創建節點, 加進List而不是直接加到Treeview.Nodes
        Dim nodeList As New List(Of TreeNode) ' 創建一個 TreeNode 的 List 來暫存所有要添加的節點
        For Each store In storeList
            Dim root As Outlook.Folder = store.GetRootFolder
            Dim node As New TreeNode(root.Name) With {.Tag = root}
            node.Nodes.Add(":::")  ' ✅ 無條件加佔位節點，省掉判斷 root.Folders.Count 這一次多餘的 COM 往返
            nodeList.Add(node)
            ' PST root folder 幾乎 100% 都有子資料夾，這個假設安全；
            ' 就算 PST 真的空了，展開時 LoadSubFolderToTreeView 清除 ":::" 後不加任何子節點，節點就會自動收起 "+" 號，行為正確
            Dbg("", root.Name)
        Next
        treeview.Nodes.AddRange(nodeList.ToArray()) ' 將所有組裝好的節點一次性添加到 treeview.Nodes
        Dbg("結束：", treeview.Name)
    End Sub
    Private Sub LoadSubFolderToTreeView(sender As Object, e As TreeViewCancelEventArgs)
        Dbg("開始：", sender.Name)
        ' 2024/5/17全部重寫, 把現在要點開的資料夾, 讀出其子資料夾並加載進treeview
        ' 5/19試過Task.Run(), Parallel.Foreach跟LINQ擴充方法了, 都沒有比較快, 別再試了, 就算virtual mode也沒有比我現在的lazy load還快
        Dim selectedNode As TreeNode = e.Node                   ' 取得點選的node
        Dim selectedFolder As Outlook.Folder = selectedNode.Tag ' 取得點選的資料夾
        Dim sortedFolders = GetSortedSubFolders(selectedFolder) ' 取得所有子資料夾並排序

        If selectedNode.Nodes.Count = 1 AndAlso selectedNode.FirstNode.Text = ":::" Then
            selectedNode.Nodes.Clear()  '清除原本暫代的假node ":::"

            ' 5/20昨天才說不會更快了, 今天改用Nodes.AddRange(), 又更快了一點, 連BeginUpdate/EndUpdate都不需要了
            ' 遍歷 storeList 並創建節點, 先加進List而不是直接加到Treeview.Nodes
            Dim nodeList As New List(Of TreeNode) ' 創建一個 TreeNode 的 List 來暫存所有要添加的節點
            For Each folder As Outlook.Folder In sortedFolders
                Dim node As New TreeNode(folder.Name) With {.Tag = folder}
                If GetFolderCount(folder) > 0 Then node.Nodes.Add(":::")
                nodeList.Add(node) ' 先加進List在記憶體中快速操作, 而不是直接加到Treeview.Nodes
                Dbg("", selectedFolder.Name & folder.Name)
            Next
            selectedNode.Nodes.AddRange(nodeList.ToArray()) ' 將所有節點一次性添加到 selectedNode.Nodes
        End If
        Dbg("結束：", sender.Name)
    End Sub
    Private Function TextHasChineseChar(name As String) As Boolean
        ' 判斷名稱是否包含中文字符
        'For Each c As Char In name : If c >= ChrW(&H4E00) AndAlso c <= ChrW(&H9FFF) Then Return True
        'Next : Return False
        Return name.Any(Function(c) c >= ChrW(&H4E00) AndAlso c <= ChrW(&H9FFF)) '使用LINQ語法, 比for each快了3~5%
    End Function
    Private Function SafeGet(Of T)(row As Outlook.Row, column As String, defaultValue As T) As T
        ''' <summary>
        ''' 安全地從 Outlook.Row 讀取欄位，自動處理 Nothing / DBNull / 例外
        ''' 2026-03-22 新增 Helper，大幅減少重複程式碼
        ''' </summary>
        ' todo: SafeGet拿來替換許多COM Exception的地方??
        Try
            Dim value = row(column)
            If value Is Nothing OrElse IsDBNull(value) Then Return defaultValue
            Return CType(value, T)
        Catch ex As System.Exception
            Dbg("SafeGet 失敗", $"{column} | {ex.Message}")
            Return defaultValue
        End Try
    End Function
#End Region
#End Region

#Region "■ 03 共用控制項行為"
#Region "  ├ 滑鼠操作事件"
    Private Sub HandleTreeViewMouseMoveShared(sender As Object, e As MouseEventArgs)
        ' ---------------------------------------------------------------
        ' 共用 TreeView / SimTree MouseMove 處理：節點 hover 色管理
        '
        ' 還原規則：
        '   SimTree 選取節點 → 還原選取色 (不能用 Color.Empty，否則藍色會閃掉) 
        '   其餘節點         → Color.Empty (原生 TreeView 預設) 
        '
        ' 套用規則：
        '   SimTree 選取節點 → 跳過 (選取色優先，不蓋 hover 色) 
        '   其餘節點         → 淡灰色 hover
        '
        ' 2026-03-17 C3 最終版：兩段結構對稱，各用一個布林封裝 SimTree 例外
        ' ---------------------------------------------------------------
        Dim treeView As TreeView = CType(sender, TreeView)
        Dim node As TreeNode = treeView.GetNodeAt(e.X, e.Y)
        If node Is _lastHoveredTreeNode Then Return

        ' ── 還原上一個 hover 節點 ──
        If _lastHoveredTreeNode IsNot Nothing Then
            Dim sim As SimTree = TryCast(treeView, SimTree)
            If sim IsNot Nothing AndAlso sim.SelectedNodes.Contains(_lastHoveredTreeNode) Then
                ' SimTree 選取節點：根據焦點還原正確的選取色 (不能 Color.Empty) 
                _lastHoveredTreeNode.BackColor = If(sim.Focused, SystemColors.Highlight, Color.FromArgb(240, 240, 240))
                _lastHoveredTreeNode.ForeColor = If(sim.Focused, SystemColors.HighlightText, SystemColors.InactiveCaptionText)
            Else
                _lastHoveredTreeNode.BackColor = Color.Empty
                _lastHoveredTreeNode.ForeColor = Color.Empty
            End If
        End If

        ' ── 套用新 hover 色 ──
        If node IsNot Nothing Then
            Dim skipHover As Boolean = TypeOf treeView Is SimTree AndAlso CType(treeView, SimTree).SelectedNodes.Contains(node)
            If Not skipHover Then
                node.BackColor = Color.FromArgb(240, 240, 240)
                node.ForeColor = SystemColors.InactiveCaptionText
            End If
        End If

        _lastHoveredTreeNode = node
    End Sub
    Private Sub TreeView1_MouseMove(sender As Object, e As MouseEventArgs) Handles TreeView1.MouseMove
        HandleTreeViewMouseMoveShared(sender, e)
    End Sub
    Private Sub TreeView2_MouseMove(sender As Object, e As MouseEventArgs) Handles TreeView2.MouseMove
        HandleTreeViewMouseMoveShared(sender, e)
    End Sub
    Private Sub TreeView3_MouseMove(sender As Object, e As MouseEventArgs) Handles TreeView3.MouseMove
        HandleTreeViewMouseMoveShared(sender, e)
    End Sub
    Private Sub TreeView4_MouseMove(sender As Object, e As MouseEventArgs) Handles TreeView4.MouseMove
        HandleTreeViewMouseMoveShared(sender, e)
    End Sub

    Private Sub SimTree1_MouseMove(sender As Object, e As MouseEventArgs) Handles SimTree1.MouseMove
        HandleTreeViewMouseMoveShared(sender, e)
    End Sub
    Private Sub SimTree2_MouseMove(sender As Object, e As MouseEventArgs) Handles SimTree2.MouseMove
        HandleTreeViewMouseMoveShared(sender, e)
    End Sub
    Private Sub SimTree3_MouseMove(sender As Object, e As MouseEventArgs) Handles SimTree3.MouseMove
        HandleTreeViewMouseMoveShared(sender, e)
    End Sub
    Private Sub SimTree4_MouseMove(sender As Object, e As MouseEventArgs) Handles SimTree4.MouseMove
        HandleTreeViewMouseMoveShared(sender, e)
    End Sub

    Private Sub HandleTreeViewMouseLeaveShared(sender As Object, e As EventArgs)
        ' 共用的 MouseLeave 處理函數
        If _lastHoveredTreeNode IsNot Nothing Then
            _lastHoveredTreeNode.BackColor = Color.Empty
            _lastHoveredTreeNode.ForeColor = Color.Empty
            _lastHoveredTreeNode = Nothing
        End If
    End Sub
    Private Sub TreeView1_MouseLeave(sender As Object, e As EventArgs) Handles TreeView1.MouseLeave
        HandleTreeViewMouseLeaveShared(sender, e)
    End Sub
    Private Sub TreeView2_MouseLeave(sender As Object, e As EventArgs) Handles TreeView2.MouseLeave
        HandleTreeViewMouseLeaveShared(sender, e)
    End Sub
    Private Sub TreeView3_MouseLeave(sender As Object, e As EventArgs) Handles TreeView3.MouseLeave
        HandleTreeViewMouseLeaveShared(sender, e)
    End Sub
    Private Sub TreeView4_MouseLeave(sender As Object, e As EventArgs) Handles TreeView4.MouseLeave
        HandleTreeViewMouseLeaveShared(sender, e)
    End Sub

    Private Sub SimTree1_MouseLeave(sender As Object, e As EventArgs) Handles SimTree1.MouseLeave
        HandleTreeViewMouseLeaveShared(sender, e)
    End Sub
    Private Sub SimTree2_MouseLeave(sender As Object, e As EventArgs) Handles SimTree2.MouseLeave
        HandleTreeViewMouseLeaveShared(sender, e)
    End Sub
    Private Sub SimTree3_MouseLeave(sender As Object, e As EventArgs) Handles SimTree3.MouseLeave
        HandleTreeViewMouseLeaveShared(sender, e)
    End Sub
    Private Sub SimTree4_MouseLeave(sender As Object, e As EventArgs) Handles SimTree4.MouseLeave
        HandleTreeViewMouseLeaveShared(sender, e)
    End Sub

    Private Sub HandleListViewMouseMoveShared(sender As Object, e As MouseEventArgs)
        ' 共用的 MouseMove 處理函數
        Dim listView As ListView = CType(sender, ListView)
        Dim item As ListViewItem = listView.GetItemAt(e.X, e.Y)
        If item Is _lastHoveredListItem Then Return

        If _lastHoveredListItem IsNot Nothing Then _lastHoveredListItem.BackColor = Color.Empty
        If item IsNot Nothing Then item.BackColor = Color.FromArgb(240, 240, 240)
        _lastHoveredListItem = item
    End Sub
    Private Sub ListView1_MouseMove(sender As Object, e As MouseEventArgs) Handles ListView1.MouseMove
        HandleListViewMouseMoveShared(sender, e)
    End Sub
    Private Sub ListView2_MouseMove(sender As Object, e As MouseEventArgs) Handles ListView2.MouseMove
        HandleListViewMouseMoveShared(sender, e)
    End Sub
    Private Sub ListView4_MouseMove(sender As Object, e As MouseEventArgs) Handles ListView4.MouseMove
        HandleListViewMouseMoveShared(sender, e)
    End Sub

    Private Sub HandleListViewMouseLeaveShared(sender As Object, e As EventArgs)
        ' 共用的 MouseLeave 處理函數
        If _lastHoveredListItem IsNot Nothing Then
            _lastHoveredListItem.BackColor = Color.Empty
            _lastHoveredListItem = Nothing
        End If
    End Sub
    Private Sub ListView1_MouseLeave(sender As Object, e As EventArgs) Handles ListView1.MouseLeave
        HandleListViewMouseLeaveShared(sender, e)
    End Sub
    Private Sub ListView2_MouseLeave(sender As Object, e As EventArgs) Handles ListView2.MouseLeave
        HandleListViewMouseLeaveShared(sender, e)
    End Sub
    Private Sub ListView3_MouseLeave(sender As Object, e As EventArgs) Handles ListView3.MouseLeave
        HandleListViewMouseLeaveShared(sender, e)
    End Sub
    Private Sub ListView4_MouseLeave(sender As Object, e As EventArgs) Handles ListView4.MouseLeave
        HandleListViewMouseLeaveShared(sender, e)
    End Sub
#End Region
#Region "  ├ 鍵盤操作事件"
    Private Sub HandleTreeViewKeyPressShared(sender As Object, e As KeyPressEventArgs)
        Dbg("開始：", sender.Name)

        ' 在這裡處理所有TreeView KeyPress 事件的程式碼
        If TypeOf sender Is TreeView Then
            ' 判斷 sender 是哪個 TreeView 控制項的實例, 使用 currentTreeView 來辨識是哪個 TreeView 控制項
            'Dim currentTreeView As TreeView = DirectCast(sender, TreeView)
            If e.KeyChar = ChrW(Keys.Enter) Then
                sender.SelectedNode.Expand()            ' 按Enter展開下一層
                If sender.name = "TreeView1" Then ListView1.Focus()

            ElseIf e.KeyChar = ChrW(Keys.Escape) Then   ' 按ESC退回上一層
                If sender.SelectedNode IsNot Nothing AndAlso sender.SelectedNode.Parent IsNot Nothing Then
                    sender.SelectedNode.Collapse() : sender.SelectedNode = sender.SelectedNode.Parent
                End If

            ElseIf e.KeyChar = ChrW(Keys.Space) Then    ' 按Space切換展開/收合
                Dim node As TreeNode = sender.SelectedNode
                If node IsNot Nothing Then              ' ✅ 避免 Space 觸發系統預設行為 (捲動等) 
                    If node.IsExpanded Then node.Collapse() Else node.Expand() : e.Handled = True
                End If
            End If
        End If

    End Sub
    Private Sub TreeView1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TreeView1.KeyPress
        HandleTreeViewKeyPressShared(sender, e)
    End Sub
    Private Sub TreeView2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TreeView2.KeyPress
        HandleTreeViewKeyPressShared(sender, e)
    End Sub
    Private Sub TreeView3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TreeView3.KeyPress
        HandleTreeViewKeyPressShared(sender, e)
    End Sub
    Private Sub TreeView4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TreeView4.KeyPress
        HandleTreeViewKeyPressShared(sender, e)
    End Sub

    Private Sub SimTree1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles SimTree1.KeyPress
        HandleTreeViewKeyPressShared(sender, e)
    End Sub
    Private Sub SimTree2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles SimTree2.KeyPress
        HandleTreeViewKeyPressShared(sender, e)
    End Sub
    Private Sub SimTree3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles SimTree3.KeyPress
        HandleTreeViewKeyPressShared(sender, e)
    End Sub
    Private Sub SimTree4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles SimTree4.KeyPress
        HandleTreeViewKeyPressShared(sender, e)
    End Sub

    Private Async Sub HandleListViewKeyPressShared(sender As Object, e As KeyPressEventArgs)
        ' ---------------------------------------------------------------
        ' 共用 ListView KeyPress 處理 (完整替換舊的 HandleListViewKeyPressShared) 
        ' 2026-03-16 C3 重構：把原本分散在三個 ListView 的 KeyPress 邏輯統一到這裡，根據 sender 來辨識是哪個 ListView，實作各自的行為
        ' 目前的實現是直接在 KeyPress 事件裡處理所有邏輯，是否要把各區塊或年度視圖和月份視圖的 Enter 鍵行為分別封裝成獨立的方法?
        '
        ' 各 ListView 行為：
        '    ListView1 : Enter = 進入子資料夾, ESC = 退回上一層  (原有邏輯不變)
        '    ListView2 : Enter = 等同雙擊 (進入月份或返回年度) , ESC = 返回年度視圖
        '    ListView3 : Enter = 打開郵件, ESC = 取消選取
        ' ---------------------------------------------------------------

        Dbg("開始：", sender.name)

        Dim lv As ListView = TryCast(sender, ListView)
        If lv Is Nothing Then Return

        ' ---------------------------------------------------------------
        ' ListView1：資料夾導覽 (保留原有邏輯，從 ListView1_KeyPress 移到這裡統一管理) 
        ' ---------------------------------------------------------------
        If lv Is ListView1 Then
            If e.KeyChar = ChrW(Keys.Enter) Then
                If lv.SelectedItems.Count = 0 Then Return
                Dim selectedItem As ListViewItem = lv.SelectedItems(0)          ' 獲取點選的資料夾並進入
                If selectedItem IsNot Nothing Then Tab1_EnterSelectedFolder(selectedItem)

            ElseIf e.KeyChar = ChrW(Keys.Escape) Then                           ' 退回上一層資料夾
                Dim itemName As String = lv.Items(0).Text                       ' 記下現在所在的listviewItem
                Dim node As TreeNode = TreeView1.SelectedNode                   ' 記下現在所在的selectedNode
                If node IsNot Nothing AndAlso node.Parent IsNot Nothing Then
                    node.Collapse() : TreeView1.SelectedNode = node.Parent      ' 選取其上層資料夾
                    Dim item As ListViewItem = FindLiSVItemByName(lv, itemName) ' 找出剛才退出前的資料夾
                    If item IsNot Nothing Then item.Selected = True : item.Focused = True : lv.Focus()
                End If

            ElseIf e.KeyChar = ChrW(1) Then                                     ' Ctrl-A 選擇 listview1 所有項目 — 2026/3/26 by AntiGravity
                lv.BeginUpdate()
                For Each item As ListViewItem In lv.Items
                    item.Selected = True
                Next
                lv.EndUpdate()
                e.Handled = True
            End If
            Return
        End If

        ' ---------------------------------------------------------------
        ' ListView2：年度 / 月份視圖導覽
        ' ---------------------------------------------------------------
        If lv Is ListView2 Then
            If e.KeyChar = ChrW(Keys.Enter) Then        ' Enter = 等同雙擊目前選定的項目
                If lv.SelectedItems.Count = 0 Then Return
                Dim selectedItem As ListViewItem = lv.SelectedItems(0)

                If _tab2IsMonthView AndAlso             ' 在月份視圖按 Enter 於返回列 → 回到年度視圖
                    selectedItem.Tag IsNot Nothing AndAlso
                    selectedItem.Tag.ToString() = "BACK" Then
                    Await ShowYearView()
                    ' ✅ 2026-03-16 Bug fix: 移除此處多餘的 item.Selected = True 造成 ListView2 出現兩個 highlighted item，且位置不正確

                ElseIf Not _tab2IsMonthView Then        ' 在年度視圖按 Enter → 進入月份視圖
                    Dim selectedYear As Integer = 0
                    If Integer.TryParse(selectedItem.Text.Trim(), selectedYear) AndAlso
                        _tab2FolderList IsNot Nothing AndAlso
                        _tab2FolderList.Count > 0 Then Await ShowMonthView(selectedYear)
                End If

            ElseIf e.KeyChar = ChrW(Keys.Escape) Then   ' ESC：不管在哪個視圖，一律返回年度視圖
                If _tab2IsMonthView Then Await ShowYearView()
                ' ✅ 2026-03-16 Bug fix: 同上，移除多餘的 item.Selected，ShowYearView 已處理
            End If
            Return
        End If

        ' ---------------------------------------------------------------
        ' ListView3：Tab3 附件搜尋結果的鍵盤操作
        ' Enter = 用 EntryID 打開郵件 (第 6 欄 SubItems(5)) 
        ' ESC   = 清除目前選取
        ' 2026-03-16 實作完成
        ' ---------------------------------------------------------------
        If lv Is ListView3 Then
            If e.KeyChar = ChrW(Keys.Enter) Then
                If lv.SelectedItems.Count = 0 Then Return
                OpenMailByEntryID(lv.SelectedItems(0).SubItems(5).Text) ' Enter = 用 EntryID 打開郵件 (第 6 欄 SubItems(5)) 
            ElseIf e.KeyChar = ChrW(Keys.Escape) Then
                If lv.SelectedItems.Count > 0 Then lv.SelectedItems(0).Selected = False
            End If
            Return
        End If
    End Sub
    Private Sub ListView1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles ListView1.KeyPress
        HandleListViewKeyPressShared(sender, e)
    End Sub
    Private Sub ListView2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles ListView2.KeyPress
        HandleListViewKeyPressShared(sender, e)
    End Sub
    Private Sub ListView3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles ListView3.KeyPress
        HandleListViewKeyPressShared(sender, e)
    End Sub
    Private Sub ListView4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles ListView4.KeyPress
        HandleListViewKeyPressShared(sender, e)
    End Sub
#End Region
#Region "  ├ 其他輔助事件"
    Private Sub TreeView1_BeforeExpand(sender As Object, e As TreeViewCancelEventArgs) Handles TreeView1.BeforeExpand
        LoadSubFolderToTreeView(sender, e)
    End Sub
    Private Sub TreeView2_BeforeExpand(sender As Object, e As TreeViewCancelEventArgs) Handles TreeView2.BeforeExpand
        LoadSubFolderToTreeView(sender, e)
    End Sub
    Private Sub TreeView3_BeforeExpand(sender As Object, e As TreeViewCancelEventArgs) Handles TreeView3.BeforeExpand
        LoadSubFolderToTreeView(sender, e)
    End Sub
    Private Sub TreeView4_BeforeExpand(sender As Object, e As TreeViewCancelEventArgs) Handles TreeView4.BeforeExpand
        LoadSubFolderToTreeView(sender, e)
    End Sub

    Private Sub SimTree1_BeforeExpand(sender As Object, e As TreeViewCancelEventArgs) Handles SimTree1.BeforeExpand
        LoadSubFolderToTreeView(sender, e)
    End Sub
    Private Sub SimTree2_BeforeExpand(sender As Object, e As TreeViewCancelEventArgs) Handles SimTree2.BeforeExpand
        Dbg("", sender.Name)
        LoadSubFolderToTreeView(sender, e)
    End Sub

    Private Sub ListView3_MouseMove(sender As Object, e As MouseEventArgs) Handles ListView3.MouseMove
        HandleListViewMouseMoveShared(sender, e)
    End Sub
    Private Sub TabControl1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl1.SelectedIndexChanged
        Dbg("開始：", sender.Name)
        Dim sw As New Stopwatch : sw.Start()
        lblStatus1.Text = "" : lblStatus2.Text = ""

        Select Case CType(sender, TabControl).SelectedTab.Text

            Case "資料夾統計"
                TreeView1.Focus()
                'SimTree1.Focus()

            Case "依日期統計"
                If SimTree2.Nodes.Count = 0 Then
                    InitChart(Chart2)
                    LoadStoreToTreeView(_pstStoreList, SimTree2)
                    ExpandTreeToDefaultInbox(SimTree2)

                End If

                'If TreeView2.Visible  And TreeView2.Nodes.Count = 0 Then
                '    LoadStoreToTreeView(PstStoreList, TreeView2)
                '    ExpandTreeToDefaultInbox(TreeView2)
                'End If
                SimTree2.Focus()

            Case "尋找附件"
                If TreeView3.Nodes.Count = 0 Then
                    LoadStoreToTreeView(_pstStoreList, TreeView3)
                    ExpandTreeToDefaultInbox(TreeView3)
                End If
                TreeView3.Focus()
                Button3_Stop.Location = Button3.Location
                'Button3.BringToFront()
                'CheckSubFolder3.BringToFront()

            Case "尋找系列郵件"
                If TreeView4.Nodes.Count = 0 Then
                    LoadStoreToTreeView(_pstStoreList, TreeView4)
                    ExpandTreeToDefaultInbox(TreeView4)
                End If
                TreeView4.Focus()

            Case "尋找重覆郵件"
                TreeView5.Visible = True
                If TreeView5.Nodes.Count = 0 Then
                    LoadStoreToTreeView(_pstStoreList, TreeView5)
                    ExpandTreeToDefaultInbox(TreeView5)
                End If
                TreeView5.Focus()

            Case "Debug"
                'CheckDebug.Checked = True

            Case Else

        End Select
        sw.Stop()
        lblStatus2.Text = "切換頁面花費了 " & sw.Elapsed.TotalSeconds.ToString("0.00") & " 秒。"

    End Sub
    Private Sub Redemption_Click(sender As Object, e As EventArgs) Handles Redemption.Click
        Dbg("開始：", sender.Name)
        Dim unused = InitRedemptionSessionWithoutDeclaration()
    End Sub
    Private Sub buttonClearCache_Click(sender As Object, e As EventArgs) Handles buttonClearCache.Click
        Dbg("開始：", sender.Name)

        ' Tab2 年份統計快取 (String key，安全直接清除) 
        _yearCountsCache.Clear()
        _monthCountsCache.Clear()

        ' Tab3 Phase1 快取 (Dictionary(Of String, FolderCacheTab3)) 
        _tab3Phase1Cache.Clear()    ' 2026-03-16 B1 新增

        ' 以下快取的 Key 是 COM 物件 (Outlook.Folder) ，.Clear() 只移除 Dictionary 的參照，不釋放 COM 物件本身，
        _mailCountCache.Clear()      ' 郵件數量快取
        _mailSizeCache.Clear()       ' 郵件大小快取
        _folderCountCache.Clear()    ' 子資料夾數量快取
        _folderSizeCache.Clear()     ' 資料夾大小快取
        _folderTreeCache.Clear()     ' 資料夾樹狀快取

        Dbg("結束：ClearAllCache() - 所有快取已清除")
        lblStatus2.Text = "所有快取已清除，下次統計將重新從 Outlook 讀取。"
    End Sub
    Private Sub AutoDismissRedemptionDialog(threadStarted As System.Threading.ManualResetEventSlim)
        ' 自動點掉 Redemption EULA dialog
        ' 使用 WinSpy++ 確認的視窗結構（2026-03-23）：
        '   視窗 class    = TEULAForm（Delphi VCL 表單），title = "Outlook Redemption"
        '   "I agree"     = TRadioButton，text = "I agree"
        '   "I do NOT..." = TRadioButton，text = "I do NOT agree"
        '   "Ok"          = TButton，    text = "Ok"
        '   "Cancel"      = TButton，    text = "Cancel"
        '
        ' v1 (2026-03-23)：PostMessage 取代 SendMessage
        '   SendMessage 在 modal dialog 阻塞 UI 執行緒時死結，
        '   PostMessage 非同步丟進佇列，由 dialog 自己的訊息泵處理
        '
        ' v2 (2026-03-23)：ShowWindow(SW_HIDE) 立刻隱藏視窗
        '   找到 TEULAForm 後立刻隱藏，使用者看不到閃爍
        '
        ' v3 (2026-03-23)：輪詢間隔 100ms → 5ms，移除固定 Thread.Sleep
        '   改成輪詢等子控制項出現，控制項一建立就立刻動作
        '   Thread.Priority = AboveNormal 確保首次啟動也能及時執行
        '
        ' v4 (2026-03-23)：加入 ManualResetEventSlim 同步點
        '   threadStarted.Set() 通知呼叫端「輪詢已開始」，
        '   呼叫端等到 Set 後才呼叫 New RDOSession()，
        '   解決 thread pool 競爭導致首次執行抓不到視窗的問題

        Dim t As New System.Threading.Thread(
            Sub()
                ' ✅ 先讓輪詢 loop 跑第一次，再通知呼叫端
                '   避免 Set() 後呼叫端立刻 New RDOSession()，
                '   但此 thread 還沒執行到 FindWindow 的競爭條件
                System.Threading.Thread.Sleep(1)
                threadStarted.Set()

                Dim hWnd As IntPtr = IntPtr.Zero
                Dim timeout As Integer = 0

                ' 輪詢找 TEULAForm，最多等 30 秒（3000 × 10ms）
                Do While hWnd = IntPtr.Zero AndAlso timeout < 3000
                    hWnd = FindWindow("TEULAForm", Nothing)
                    If hWnd = IntPtr.Zero Then
                        System.Threading.Thread.Sleep(5)
                        timeout += 1
                    End If
                Loop

                If hWnd = IntPtr.Zero Then
                    Dbg("AutoDismissRedemption", "逾時：找不到 TEULAForm") : Return
                End If

                ' ✅ 立刻隱藏，使用者不會看到 EULA dialog 閃出來
                ShowWindow(hWnd, SW_HIDE)
                Dbg("AutoDismissRedemption", $"TEULAForm 隱藏 hWnd=0x{hWnd.ToString("X")}")

                ' ── Step 1："I agree" TRadioButton ──────────────────────
                ' 輪詢等子控制項建立完成（視窗已隱藏，等待時間使用者無感）
                Dim hAgree As IntPtr = IntPtr.Zero
                Dim childTimeout As Integer = 0
                Do While hAgree = IntPtr.Zero AndAlso childTimeout < 50  ' 最多 250ms
                    hAgree = FindWindowEx(hWnd, IntPtr.Zero, "TRadioButton", "I agree")
                    If hAgree = IntPtr.Zero Then
                        System.Threading.Thread.Sleep(5) : childTimeout += 1
                    End If
                Loop

                If hAgree <> IntPtr.Zero Then
                    PostMessage(hAgree, WM_LBUTTONDOWN, New IntPtr(1), IntPtr.Zero)
                    PostMessage(hAgree, WM_LBUTTONUP, New IntPtr(1), IntPtr.Zero)
                    Dbg("AutoDismissRedemption", "'I agree' PostMessage 送出")
                Else
                    Dbg("AutoDismissRedemption", "找不到 'I agree'（已逾時）")
                End If

                ' ── Step 2："Ok" TButton ────────────────────────────────
                Dim hOk As IntPtr = IntPtr.Zero
                Dim okTimeout As Integer = 0
                Do While hOk = IntPtr.Zero AndAlso okTimeout < 50        ' 最多 250ms
                    hOk = FindWindowEx(hWnd, IntPtr.Zero, "TButton", "Ok")
                    If hOk = IntPtr.Zero Then
                        System.Threading.Thread.Sleep(5) : okTimeout += 1
                    End If
                Loop

                If hOk <> IntPtr.Zero Then
                    PostMessage(hOk, WM_LBUTTONDOWN, New IntPtr(1), IntPtr.Zero)
                    PostMessage(hOk, WM_LBUTTONUP, New IntPtr(1), IntPtr.Zero)
                    Dbg("AutoDismissRedemption", "'Ok' PostMessage 送出")
                Else
                    Dbg("AutoDismissRedemption", "找不到 'Ok'（已逾時）")
                End If
            End Sub)

        t.Priority = System.Threading.ThreadPriority.AboveNormal  ' 確保首次啟動 thread pool 忙碌時仍能及時執行
        t.IsBackground = True   ' Form 關閉時自動結束，不需要手動管理生命週期
        t.Start()

    End Sub
#End Region
#End Region

#Region "■ 04 Tab1：資料夾統計 — 重構後程式碼 v4 (最終版) ==="
    ' ==============================================================
    '
    ' ── 版本演進摘要 ──────────────────────────────────────────────
    '
    '   原始版  循序 Await GetInfoForListview × N，各自等遞迴完成後才輪下一個
    '           GetFolderSizeLegacy 用 Task.Run 包 COM (STA 違規) 
    '           s4Task.Result 潛在 deadlock
    '           cache: 0.10~0.19s
    '
    '   v1      BFS 一次展開整棵子樹，GetMailCount 循序讀 PR_CONTENT_COUNT
    '           底部向上彙總後一次寫快取，之後點選子資料夾直接命中，架構最乾淨，
    '           但有 bug：root 快取命中時不展開子資料夾 → 第二次點選 ListView 只顯示 root 自身
    '           cache: 0.01s (最快，因為完全不碰 thread pool) 
    '
    '   v2      Task.WhenAll 同時發起 N 個子資料夾的計算 (並行的並行) 修掉 s4Task.Result deadlock
    '           1st read 明顯變快；但 cache 仍有 40 次 Task.Run dispatch overhead
    '           cache: 0.04~0.09s (因 Task.Run overhead 限制) 
    '
    '   v3      BFS + Task.WhenAll 試圖合併 v1 + v2 優點
    '           但 ComputeFolderDisplayList 在 UI 執行緒循序走整棵子樹 → 更慢
    '
    '   v3fix   修正 v3 過深遍歷問題，ComputeFolderDisplayList 只收 depth=0/1
    '           效能介於 v1 和 v2 之間，但仍有 Task.Run overhead
    '           cache: 0.05~0.08s
    '
    '   v4 (本版) 
    '           v1 的 BFS 架構 + 一行 bug fix：root 永遠展開直屬子資料夾
    '           保留 v1 的所有效能優勢，同時修正第二次點選只顯示 root 的問題
    '           不引入 Task.WhenAll (實測 sequential BFS 比 parallel of parallel 快) 
    '           cache: 0.01s (應當與 v1 相同) 
    '
    ' ── 為什麼 v4 不用 Task.WhenAll？─────────────────────────────
    '
    '   v2/v3fix 的「並行的並行」看起來應該更快，但實測反而輸給 v1，原因：
    '   PST 的 PR_CONTENT_COUNT 讀取是 COM overhead 主導 (不是 I/O bottleneck) 
    '   v1 的 BFS sequential：N 個資料夾 × 1 PR_CONTENT_COUNT call = O(N)，無其他 overhead
    '   v2/v3fix 的 Task.WhenAll：20 子資料夾 × 2 Task.Run = 40 次 thread pool dispatch
    '            每次 dispatch ~1~2ms，40 次 = 40~80ms → 這就是 cache 0.05s 的來源
    '
    '   PST 是單一檔案，並行讀取可能造成 I/O 競爭，在慢速 HDD 上優勢也有限
    '   → v1 的 sequential BFS 在此場景下已是最優，不需要 Task.WhenAll  ' todo: 但我還是想要再嚐試看看, 我覺得上次測試不是這個原因
    '
    ' ── 分層架構 ──────────────────────────────────────────────────
    '
    '   L1  TreeView1_AfterSelect   UI 事件層
    '       取得選中資料夾 → 呼叫 L2 → 批次更新 ListView1
    '       規則：不做計算，不直接操作 COM，只傳達意圖與呈現結果
    '
    '   L2  ComputeFolderStatsAsync 流程協調層 (核心) 
    '       BFS 展開整棵子樹 (root 永遠展開直屬子，其餘節點依快取決定) 
    '       → 呼叫 L3 讀每個節點的直接郵件數
    '       → 底部向上彙總 (O(N)，無遞迴 stack overflow 風險) 
    '       → 一次性寫快取 (整棵子樹預讀) 
    '       → 回傳 root + 直屬子資料夾清單供 L1 顯示
    '       回呼 onProgress 讓 L1 更新進度，L2 自身不碰任何 UI 控制項
    '
    '   L3  GetMailCount            COM 資料層
    '       只讀單一資料夾的 PR_CONTENT_COUNT (本層郵件數，不含子孫) 
    '       不遞迴，不展開子資料夾，最小化 COM 呼叫量
    '
    ' ── 快取策略 ──────────────────────────────────────────────────
    '
    '   快取 key：Outlook.Folder COM 物件 (沿用現有設計，接受偶爾 RCW 不同的 cache miss) 
    '
    '   mailCountCache   → TotalMailCount (含子孫郵件總數)   L2 底部向上彙總後寫入，TryAdd 不覆蓋既有值
    '   folderCountCache → TotalSubCount (含子孫資料夾總數)  L2 底部向上彙總後寫入，TryAdd 不覆蓋既有值
    '   folderSizeCache  → 資料夾大小 (Lazy，由 ColumnClick / 右鍵觸發計算) 
    '   folderTreeCache  → 子資料夾排序清單 (GetSortedSubFolders 負責維護) 
    '
    '   快取命中剪枝規則：
    '     root (BFS 起點)   → 永遠展開直屬子資料夾 (v4 bug fix 的核心) 
    '     非 root 節點      → mailCountCache + folderCountCache 都命中 → IsFromCache=True → 不再往下展開
    '     效果：第一次點選做完整 BFS；後續點選命中快取，BFS 剪枝到只剩兩層，幾乎瞬間完成
    '
    ' ── 效能特點 ──────────────────────────────────────────────────
    '
    '   第一次點選：BFS 展開整棵子樹 (N 個資料夾 × 1 PR_CONTENT_COUNT) ，快取預讀一次到位
    '   後續點選  ：命中快取，BFS 剪枝，底部向上加總純在記憶體執行 → 0.01s
    '   快速點選  ：序號機制確保只有最後一次結果寫 ListView
    '   STA 安全  ：所有 COM 呼叫在 UI 執行緒；Task.Yield() 每 20 個資料夾讓出一次 UI
    '
    ' ── 使用說明 ──────────────────────────────────────────────────
    '
    '   【加入成員變數】 (放在 _tab2... 附近) 
    '     Private _tab1SelectSeq As Integer = 0
    '
    '   【替換以下函數】
    '     - TreeView1_AfterSelect   → 本檔 L1 取代
    '     - GetInfoForListview      → 由 L2/L3 取代，舊函數可刪除
    '     - GetFolderSizeLegacy     → 本檔修正版取代 (移除 Task.Run 包 COM) 
    '
    '   【完全不動的函數】
    '     - GetMailCountByMAPINew    保留 (GetFolderSizeLegacy exception path 仍呼叫) 
    '     - GetTotalFolderCountAsync 保留 (不再由 Tab1 主流程呼叫，但其他地方可能用到) 
    '     - GetSortedSubFolders      不改 (L2 BFS 直接呼叫) 
    '     - GetFolderByName, FindNodeByName (右鍵、雙擊功能用) 
    '     - ListView1_ColumnClick, ListView1_ItemMenu, EnterFolderMenuItem
    '     - GetFolderSizeOld_Async (問題資料夾的 fallback，新版 GetFolderSizeLegacy 仍呼叫) 
    '
    ' ==============================================================


    ' ─────────────────────────────────────────────────────────────
    ' FolderBfsEntry：BFS 過程中每個資料夾節點的容器
    ' 貫穿 L2 的所有步驟 (BFS 展開 → L3 讀取 → 底部向上彙總 → 快取寫入 → 回傳清單) 
    ' ─────────────────────────────────────────────────────────────
#Region "  ├ L1 UI事件層"
    Private Async Sub TreeView1_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TreeView1.AfterSelect
        ' ==============================================================
        ' === Layer 1 (UI 事件層) ===
        ' 職責: 回應 TreeView1 點選，呼叫 L2 計算，批次更新 ListView1
        ' 規則: 不做遞迴計算，不直接操作 COM，只傳達意圖與呈現結果
        ' ==============================================================
        Dbg("開始：", sender.Name)
        _cancelRequested = False    ' ✅ 每次新點選 reset，避免上一次的 ESC 殘留影響本次

        ' 序號機制：每次點選遞增；計算完成後若序號已變，代表有更新的點選，丟棄本次結果, 避免快速切換資料夾時舊結果覆蓋新結果
        Dim mySeq As Integer = System.Threading.Interlocked.Increment(_tab1SelectSeq)

        Dim sw As New Stopwatch : sw.Start()
        lblStatus1.Text = "" : lblStatus2.Text = "" : Cursor = Cursors.WaitCursor

        Dim selectedFolder As Outlook.Folder = TryCast(e.Node.Tag, Outlook.Folder)
        If selectedFolder Is Nothing Then Cursor = Cursors.Default : Return

        ' todo: debugForm開啟的時候, addmessage拖累程式運作速度
        ' tab1 流程為何一直沒有進到GetMailCountAll()?? 就沒使用到GetArray()加速??
        ' 如何在大量迴圈起始前先beginUpdate? 或是讓debugform.addmessage非同步運作?
        ' 為何在selectednode己經統計一次了, 但點開+號的時候還會再統計一次subfodler.count? (只有第一次點開才會, 重覆開合就不會)
        ' 在同樣二個A/B 目錄之間切換, 會一再看見RDO 成功GetMailCount, 但其實不是應該去讀快取的嗎?

        Try ' L2：BFS 展開整棵子樹，快取命中剪枝，底部向上彙總，回傳顯示清單
            ' 第一次點選：完整遍歷整棵子樹並預讀快取
            ' 後續點選  ：命中快取，BFS 立即剪枝，近乎瞬間完成
            Dim rows As List(Of FolderBfsEntry) = Await ComputeFolderStatsAsync(selectedFolder,
                    Sub(processed As Integer, total As Integer) lblStatus1.Text = "正在處理: " & processed & " / " & total & " 個資料夾...")

            If _tab1SelectSeq <> mySeq Then Return
            If _cancelRequested OrElse rows.Count = 0 Then   ' ✅ ESC 中斷或 ComputeFolderStatsAsync 回空 List → 不更新 ListView
                lblStatus2.Text = "已中斷。" : Cursor = Cursors.Default : Return
            End If

            Dim items As New List(Of ListViewItem)  ' 批次建立 ListViewItem 並一次性塞入 ListView
            For i As Integer = 0 To rows.Count - 1
                items.Add(BuildListViewItem_Tab1(rows(i), isRoot:=(i = 0)))
            Next

            ListView1.BeginUpdate()                 ' BeginUpdate/AddRange/EndUpdate 避免逐筆 Add 造成重繪閃爍
            ListView1.Items.Clear()
            ListView1.Items.AddRange(items.ToArray())
            ListView1.EndUpdate()

        Catch ex As System.Exception
            Dbg("Error: TreeView1_AfterSelect", ex.Message)
        End Try
        sw.Stop()

        If Not _isFirstInit Then lblStatus2.Text = "統計花費了 " & sw.Elapsed.TotalSeconds.ToString("0.00") & " 秒。" Else _isFirstInit = False

        lblStatus1.Text = ""
        Cursor = Cursors.Default
        TreeView1.Enabled = True : TreeView1.Focus()
        Dbg("結束：", sender.Name)
    End Sub
    Private Sub TreeView1_MouseClick(sender As Object, e As MouseEventArgs) Handles TreeView1.MouseClick
        Dbg("開始：", sender.Name)
        If e.Button = MouseButtons.Left AndAlso _isFirstInit = True Then _isFirstInit = False   ' 只為了第一次啟動時自動展開第一層資料夾, 點選之後就不再自動展開了, 以免干擾使用者操作
    End Sub
    Private Sub ListView1_MouseClick(sender As Object, e As MouseEventArgs) Handles ListView1.MouseClick
        Dbg("開始：", sender.Name)
        If e.Button = MouseButtons.Right Then _ctxListView1.Show(System.Windows.Forms.Cursor.Position)    ' ✅ 直接顯示已初始化好的選單，不重複建立和 AddHandler
        ' 2026/3/6: 原有程式碼每次都會新建一個ContextMenuStrip, 每次都新建一個都要重新AddHandler會造成memory leak
        ' 現在改成只在initial的時候建立一次, 之後每次右鍵點擊的時候直接Show()就好, 不用再重複建立
    End Sub
    Private Sub ListView1_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles ListView1.MouseDoubleClick
        Dbg("開始：", sender.Name)
        If e.Button = MouseButtons.Left AndAlso e.Clicks = 2 Then           ' Double-click就跳至該資料夾統計資料顯示
            Dim selectedItem As ListViewItem = sender.GetItemAt(e.X, e.Y)   ' 獲取點選的資料夾並進入
            If selectedItem Is Nothing Then Exit Sub Else Tab1_EnterSelectedFolder(selectedItem)
            ' 上面這句的語法超簡潔又易讀, 還有其他地方可以改成這樣的寫法嗎??   
        End If
    End Sub
    Private Sub ListView1_GotFocus(sender As Object, e As EventArgs) Handles ListView1.GotFocus
        If ListView1.SelectedItems.Count = 0 AndAlso ListView1.Items.Count > 0 Then
            ListView1.Items(0).Selected = True
        End If
        'ListView1.Invalidate()
    End Sub
    Private Async Sub ListView1_ItemMenu(sender As Object, e As EventArgs)
        Dbg("開始：", sender.Name)

        Dim stopwatch As New Stopwatch : stopwatch.Start()
        Dim selectedItems As ListView.SelectedListViewItemCollection = ListView1.SelectedItems  ' 如果有選中項目, 獲取所選中的項目
        'Dim selectedItem As ListViewItem = selectedItems(0)                                    ' 取得第一個選中項目
        'Dim folderSizeSubItem As ListViewItem.ListViewSubItem = selectedItem.SubItems(1)       ' 假設 FolderSize 是第二列的子項目

        If selectedItems.Count > 0 Then
            For Each s As ListViewItem In selectedItems
                If s.Index = 0 Then Continue For ' 若選中是本體目錄則跳過 ' todo: 為何跳過??

                ' 如果已經有FolderSize的子項目就先把它改成「計算中...」, 如果還沒有就先加一個占位用的子項目
                If s.SubItems.Count > 4 Then s.SubItems(4).Text = "計算中..." Else s.SubItems.Add("計算中...")
                'ListView1.Refresh()    ' ✅ 強制立即重繪，確保「計算中...」在計算開始前就顯示出來
                'Await Task.Yield()     ' 讓 UI 有機會更新，確保「計算中...」在計算開始前就顯示出來
                'Await Task.Delay(0)    ' 讓 UI 有機會更新，確保「計算中...」在計算開始前就顯示出來
                'Task.Yield 只是把後續程式碼排進訊息佇列，不保證重繪先完成。Task.Delay(0) 則是真正等待至少一個訊息泵循環，重繪一定在計算開始前完成。
            Next

            For Each s As ListViewItem In selectedItems
                If s.Index = 0 Then Continue For ' 若選中是本體目錄則跳過

                ' 2026/3/24 by AntiGravity: 改用 Tag 取回 Folder，避免 GetFolderByName 遞迴展開 TreeView
                Dim folder As Outlook.Folder = TryCast(s.Tag, Outlook.Folder)
                If folder Is Nothing Then Continue For
                Dim folderSize As Long = Await GetFolderSizeAll(folder)
                Dim strFolderSize As String

                If folderSize < 0 Then
                    strFolderSize = "計算失敗"
                ElseIf folderSize = 0 AndAlso GetMailCount(folder) > 0 Then
                    '如果明明有東西卻讀回快取內容大小為零, 就重新用傳統mailitem迴圈讀取, 重新存入cache
                    'todo: 這裡有可能會進來嗎??
                    GetFolderSizeOld_Async(folder)
                    strFolderSize = "**重新計算..."
                    Dbg("警告: " & folder.Name & " 的資料夾大小讀取到快取是 0KB, 已重新讀取: ", strFolderSize)
                Else
                    strFolderSize = (folderSize / 1024).ToString("###,###,###,##0KB")
                End If

                If s.SubItems.Count > 4 Then s.SubItems(4).Text = strFolderSize Else s.SubItems.Add(strFolderSize)
            Next
        End If
        lblStatus2.Text = "統計資料夾大小花費了 " & stopwatch.Elapsed.TotalSeconds.ToString("0.00") & " 秒。"
    End Sub
    Private Sub Tab1_EnterSelectedFolder(selectedItem As ListViewItem)
        ' ★ 核心修正 (2026-03-20) by Claude.ai：
        '
        ' TreeView 使用 lazy loading：子節點未展開時 .Nodes 只有 ":::" 佔位節點。
        ' 因此在搜尋目標節點前，必須先確保 SelectedNode 已展開，讓真實子節點載入。
        '
        ' 展開 SelectedNode 只觸發一次 BeforeExpand → LoadSubFolderToTreeView，這是正確且必要的。
        ' 問題出在舊版後續的 TreeView1.SelectedNode = foundNode：
        '   WinForms setter 內部呼叫 Win32 TVM_ENSUREVISIBLE，
        '   TVM_ENSUREVISIBLE 沿祖先鏈逐一 Expand()，每個都觸發 BeforeExpand → LoadSubFolderToTreeView，
        '   即使 foundNode 是直屬子節點也如此 (Win32 層不知道它已可見) 。
        '
        ' 修正方案：
        '   ① SelectedNode.Expand()         → 只展開父節點，載入真實子節點 (一次 BeforeExpand，正確) 
        '   ② 在真實子節點裡找 foundNode     → 不遞迴 (FindNodeByName 每個節點都 Expand()，已知錯誤) 
        '   ③ foundNode.Tag判斷folder.count → 確認目標資料夾有子資料夾才進入
        '   ④ SendMessage TVM_SELECTITEM    → 直接在 Win32 層選取 foundNode，
        '      繞過 WinForms setter 的 EnsureVisible 路徑，不再展開任何額外節點。
        '      Win32 TVM_SELECTITEM 仍會發出 TVN_SELCHANGED，
        '      WinForms 收到後自動觸發 TreeView1_AfterSelect，行為與原本完全一致。

        Dbg("開始：", selectedItem.SubItems(0).Text)

        If TreeView1.SelectedNode Is Nothing Then Return

        ' ① 確保父節點已展開 (若只有 ":::" 則展開一次，載入真實子節點) 
        '   若已展開則 Expand() 無作用 (WinForms 不會重複觸發 BeforeExpand) 
        TreeView1.SelectedNode.Expand()

        ' ② 在直屬子節點裡找目標 (不遞迴，不呼叫任何 Expand) 
        Dim subject As String = selectedItem.SubItems(0).Text.Replace(" - ", "")
        Dim foundNode As TreeNode = Nothing
        For Each node As TreeNode In TreeView1.SelectedNode.Nodes
            If node.Text.Replace(" - ", "") = subject Then
                foundNode = node : Exit For
            End If
        Next
        If foundNode Is Nothing Then Return

        ' ③ 確認目標資料夾有子資料夾才進入
        '    以 foundNode.Tag 取得 Outlook.Folder，呼叫 GetFolderCount 判斷
        '    GetFolderCount 內有快取 (folderTreeCache) ，重複點選不重讀 COM
        Dim targetFolder As Outlook.Folder = TryCast(foundNode.Tag, Outlook.Folder)
        If targetFolder Is Nothing OrElse GetFolderCount(targetFolder) = 0 Then
            Dbg("已攔截：目標資料夾無子資料夾，不進入", subject)
            Return
        End If

        foundNode.EnsureVisible()   ' 捲動使節點可見 (不展開祖先，因父節點已展開) 

        ' ④ 用 Win32 直接選取treeview.selectednode，繞過 WinForms SelectedNode setter 的 EnsureVisible 路徑
        SendMessage(TreeView1.Handle, TVM_SELECTITEM, New IntPtr(TVGN_CARET), foundNode.Handle)

        ListView1.Focus()
        If ListView1.Items.Count > 0 Then ListView1.Items(0).Selected = True
    End Sub
#End Region
#Region "  ├ L2 流程協調層"
    Private Async Function ComputeFolderStatsAsync(rootFolder As Outlook.Folder, onProgress As Action(Of Integer, Integer)) As Task(Of List(Of FolderBfsEntry))
        ' ==============================================================
        ' === Layer 2 (流程協調層) ===
        ' 職責: BFS 展開整棵子樹，管理快取剪枝，驅動 L3，底部向上彙總，回傳顯示清單
        '
        ' 四個步驟：
        '   Step 1  BFS 展開：收集整棵子樹的所有節點；快取命中 (非 root) 則剪枝不再往下
        '   Step 2  L3 讀取：對未快取節點逐一呼叫 GetMailCount()，取本層郵件數
        '   Step 3  底部向上彙總：利用 BFS「父索引 < 子索引」的特性，從尾端往前掃一次完成
        '   Step 4  寫快取：TryAdd 整棵子樹，後續任何子資料夾點選都可命中
        '   Step 5  組裝回傳清單：root + 直屬子資料夾 (ParentIndex=0) ，快取命中節點補讀 DirectMailCount
        ' ' todo: 這裡邏輯是否有點過度複雜? 真的需要判斷 "是否來自快取" 嗎? (讓claude再自己檢查一次)
        '
        ' onProgress 回呼：
        '   由 L1 傳入，每處理 20 個資料夾回呼一次 (同時 Task.Yield 讓出 UI 執行緒) 
        '   L2 自身完全不碰任何 UI 控制項，保持分層乾淨
        '
        ' v4 bug fix (相對於 v1) ：
        '   BFS 剪枝規則改為「root (parentIdx=-1) 永遠展開直屬子資料夾，不論快取」
        '   修正 v1 第二次點選時 root 快取命中 → 子資料夾不展開 → ListView 只顯示 root 的問題
        ' todo: 如何在tab1的邏輯簡化?
        ' ==============================================================
        Dbg("開始：ComputeFolderStatsAsync: ", rootFolder.Name)

        ' ── Step 1: BFS 展開整棵子樹 ──────────────────────────────────────────
        Dim allEntries As New List(Of FolderBfsEntry)
        Dim queue As New Queue(Of (folderObj As Outlook.Folder, parentIdx As Integer))
        queue.Enqueue((rootFolder, -1))

        Do While queue.Count > 0
            Dim curr = queue.Dequeue()
            Dim entry As New FolderBfsEntry With {.Folder = curr.folderObj,
                                                  .ParentIndex = curr.parentIdx,
                                                  .IsFromCache = False}
            Dim myIdx As Integer = allEntries.Count
            allEntries.Add(entry)

            ' 快取命中判斷：兩個快取都有才算完整命中 (任一失效都重新計算，確保一致性) 
            Dim cachedMail As Integer, cachedSub As Integer
            If _mailCountCache.TryGetValue(curr.folderObj, cachedMail) AndAlso _folderCountCache.TryGetValue(curr.folderObj, cachedSub) Then
                entry.TotalMailCount = cachedMail
                entry.TotalSubCount = cachedSub
                entry.IsFromCache = True
                ' ★ v4 bug fix：root (parentIdx=-1) 即使快取命中，也要繼續展開直屬子資料夾
                '   v1 在此直接 Continue (不展開) ，導致第二次點選 ListView 只顯示 root 自身
                '   修正：只有非 root 節點才允許剪枝
                If curr.parentIdx <> -1 Then Continue Do  ' 非 root 快取命中 → 剪枝，不再往下展開
            End If

            ' 未命中，或是 root (不論有無快取) → 展開直屬子資料夾
            ' GetSortedSubFolders 內有 folderTreeCache，重複點選不重讀 COM ' todo: 為什麼每次在按右鍵展開子樹的時候都要loadsubfolder? (貼debug message詢問)
            For Each subFolder As Outlook.Folder In GetSortedSubFolders(curr.folderObj)
                queue.Enqueue((subFolder, myIdx))
            Next
        Loop

        Dim total As Integer = allEntries.Count
        Dbg("BFS 完成: ", total & " 個節點 (含快取命中剪枝) ")

        ' ── Step 2: L3 讀取各節點的直接郵件數 ────────────────────────────────
        ' IsFromCache=True 的節點已有 TotalMailCount (從快取) ，不需再讀 COM
        ' IsFromCache=False 的節點讀 PR_CONTENT_COUNT，作為底部向上彙總的初始值
        Dim processed As Integer = 0
        For i As Integer = 0 To total - 1
            Dim entry As FolderBfsEntry = allEntries(i)
            If Not entry.IsFromCache Then
                entry.DirectMailCount = GetMailCount(entry.Folder)
                entry.TotalMailCount = entry.DirectMailCount    ' 初始值 = 本層，後面底部向上累加子孫
                entry.TotalSubCount = 0                         ' 初始為 0，後面累加子孫資料夾數
            End If

            processed += 1
            If processed Mod 20 = 0 Then    ' 每掃瞄20個郵件就讓出一次控制權
                onProgress?.Invoke(processed, total) : Await Task.Yield()
                If _cancelRequested Then Return New List(Of FolderBfsEntry)  ' ✅ ESC 中斷：放棄本次計算，L1 收到空 List 不更新 ListView
            End If
        Next

        ' ── Step 3: 底部向上彙總 ────────────────────────────────────────────
        ' BFS 特性保證：父節點索引 < 子節點索引 (因為父節點先入佇列先出佇列先被加入 allEntries) 
        ' 從尾端往前掃一次 = 底部向上 (所有子孫都累加完才輪到父節點) 
        ' ★ 邏輯等同遞迴，但 O(N) 線性掃描，無 stack overflow 風險
        '
        ' ★ 重要：parent.IsFromCache=True 代表該節點的 TotalMailCount 已是含子孫的正確快取值
        '   此時不能再疊加 child.TotalMailCount，否則會雙重計算導致第二次點選數字膨脹
        '   只有 parent.IsFromCache=False 的節點才需要從子孫累加 
        For i As Integer = allEntries.Count - 1 To 1 Step -1
            Dim child As FolderBfsEntry = allEntries(i)
            Dim parent As FolderBfsEntry = allEntries(child.ParentIndex)
            ' todo: 如何修改讓這裡的TotalMailCount 加速? 有rdo就直接讀全數, 沒有的才去累加?
            '       或, 直接一律套用GetMailCountAll(), 不管有沒有rdo 就去L3函數裡判斷?
            If Not parent.IsFromCache Then                          ' ★ 快取命中的 parent 已含子孫總計，不再疊加
                parent.TotalMailCount += child.TotalMailCount
                parent.TotalSubCount += child.TotalSubCount + 1     ' +1 = child 這個資料夾本身也計入
            End If
        Next

        ' ── Step 4: 把新計算的彙總結果寫入快取 ──────────────────────────────
        ' 一次性快取整棵子樹：後續點選任何子資料夾都能命中快取
        ' TryAdd 不覆蓋既有值，避免污染快取 (RCW 相同才能命中，不同的 RCW 只是 cache miss) 
        For Each entry As FolderBfsEntry In allEntries
            If Not entry.IsFromCache Then
                _mailCountCache.TryAdd(entry.Folder, entry.TotalMailCount)
                _folderCountCache.TryAdd(entry.Folder, entry.TotalSubCount)
            End If
        Next

        ' ── Step 5: 組裝回傳清單 (root + 直屬子資料夾) ────────────────────
        ' 直屬子資料夾 = ParentIndex = 0 (rootFolder 在 allEntries 的索引永遠為 0) 
        ' 快取命中的直屬子資料夾沒有讀 DirectMailCount (Step 2 跳過) ，在此補讀一次
        ' 只有直屬子資料夾需要補讀，子孫不顯示在 ListView 所以不補讀
        Dim result As New List(Of FolderBfsEntry)
        result.Add(allEntries(0))   ' index 0 = rootFolder 本身

        For i As Integer = 1 To allEntries.Count - 1
            Dim entry As FolderBfsEntry = allEntries(i)
            If entry.ParentIndex = 0 Then
                If entry.IsFromCache Then entry.DirectMailCount = GetMailCount(entry.Folder)
                result.Add(entry)
            End If
        Next

        ' root 自身的 DirectMailCount (Step 2 已讀，IsFromCache 通常為 False 所以不需補讀) 
        ' 但若 root 自身快取命中 (IsFromCache=True) ，DirectMailCount 仍為 0，補讀一次
        If allEntries(0).IsFromCache Then allEntries(0).DirectMailCount = GetMailCount(allEntries(0).Folder)

        onProgress?.Invoke(total, total)
        Dbg("結束：", "回傳 " & result.Count & " 列 (1 root + " & (result.Count - 1) & " 直屬子資料夾) ")
        Return result
    End Function
    Private Function BuildListViewItem_Tab1(entry As FolderBfsEntry, isRoot As Boolean) As ListViewItem
        ' ─────────────────────────────────────────────────────────────
        ' 組裝 ListView1 的單一 ListViewItem
        ' 欄位: 資料夾名稱 / 本層郵件數 / 含子孫資料夾總數 / 含子孫郵件總數 / 大小 (Lazy) 
        ' isRoot=True  → 顯示名稱不加前綴 (選中的資料夾本身) 
        ' isRoot=False → 顯示名稱加「 - 」前綴 (直屬子資料夾，視覺上縮排) 
        ' ─────────────────────────────────────────────────────────────
        Dim displayName As String = If(isRoot, entry.Folder.Name, " - " & entry.Folder.Name)

        ' 大小：Lazy，從快取讀；未計算過則留空，等 ColumnClick 或右鍵選單觸發計算
        Dim sizeStr As String = ""
        Dim sizeVal As Long
        If _folderSizeCache.TryGetValue(entry.Folder, sizeVal) AndAlso sizeVal > 0 Then sizeStr = (sizeVal \ 1024L).ToString("###,###,###,##0") & "KB"

        Dim lvi As New ListViewItem({displayName,
                                 entry.DirectMailCount.ToString("###,###,##0"),  ' 欄1: 本層郵件數 (不含子孫) 
                                 entry.TotalSubCount.ToString("###,###,##0"),    ' 欄2: 含子孫資料夾總數
                                 entry.TotalMailCount.ToString("###,###,##0"),   ' 欄3: 含子孫郵件總數
                                 sizeStr})                                       ' 欄4: 大小 (Lazy) 
        lvi.Tag = entry.Folder ' 將 Folder 物件存在 Tag，以避免後續在 TreeView 中遞迴搜尋
        Return lvi

    End Function
#End Region
#Region "  └ 輔助函數"
    Private Async Sub ExpandTreeToDefaultInbox(treeview As TreeView)
        Dbg("開始：", treeview.Name)

        If treeview.Nodes.Count = 0 Then Return

        treeview.BeginUpdate()
        Dim rootNode = treeview.Nodes(0)
        If treeview.Nodes(0).Nodes.Count > 0 Then
            rootNode.Expand()
            ' ---------------------------------------------------------------
            ' 【修正1】ExpandTreeToDefaultInbox —— 迴圈上限寫錯
            ' 把 treeview.Nodes.Count - 1 改成 treeview.Nodes(0).Nodes.Count - 1
            ' ---------------------------------------------------------------
            ' ✅ 修正: 應遍歷第一個 PST 的「子資料夾」數量，而非根節點數量
            ' 舊版: treeview.Nodes.Count - 1 = PST 個數 (通常=1) ，只會檢查第一個子資料夾
            ' 新版: treeview.Nodes(0).Nodes.Count - 1 = 第一個 PST 下的所有子資料夾數
            For i As Integer = 0 To rootNode.Nodes.Count - 1
                Try
                    Dim node As TreeNode = rootNode.Nodes(i)
                    If node.Text.Contains("Inbox") Or node.Text.Contains("收件匣") Then
                        Dbg("Found default inbox: ", node.FullPath)
                        If TypeOf treeview Is SimTree Then
                            ' 2026/3/18: 必須明確 TryCast 到 SimTree，才能正確呼叫 AddSelectedNode 更新 _selectedNodes 和高亮色
                            ' Hack: 同時, 把自訂控制項裡面的 FireAfterSelect() 從 private 改成 public, 直接手動觸發 AfterSelect 事件
                            Dim st As SimTree = DirectCast(treeview, SimTree)
                            st.AddSelectedNode(node)    ' ← SimTree 專用路徑：直接更新 _selectedNodes + 高亮
                            st.FireAfterSelect(node)    ' ← 直接手動觸發 AfterSelect 事件，讓統計邏輯跑起來
                        ElseIf TypeOf treeview Is TreeView Then
                            ' 2026/3/18: debug找了好幾天, 首次切換到tab2時, SimTree2無法正確選取到預設的收件匣
                            ' 結果原來是下面這行 treeview.SelectedNode = node 送到SimTree控制項, 沒有被觸發選中的event.
                            treeview.SelectedNode = node
                        End If
                        treeview.Focus() : treeview.Refresh() : treeview.EndUpdate() : Await Task.Yield : Exit Sub
                    End If
                Catch
                End Try
            Next
        End If

    End Sub
    Private Async Function GetTotalFolderCountAsync(folder As Outlook.Folder) As Task(Of Integer)
        Dbg("開始：", folder.Name)
        Dim value As Integer
        If _folderCountCache.TryGetValue(folder, value) Then Return value ' 檢查快取中是否已存在值, 若有則直接返回

        Dim totalSubCount As Integer = GetFolderCount(folder)           ' 初始值為點選資料夾的子資料夾數量

        ' 5/21測試記錄: 這裡使用ConcurrentBag跟使用results.sum應該要比較快, 但不知為何實測結果都比GetTotalFolderCount_Old()還慢了5%, 這個函數先保留不清除
        ' 5/21最後決定: 二個函數快慢互有變化, 但GetTotalFolderCountAsync()的穩定性較好, 比New()的標準差來得小, 所以決定使用這個
        ' 使用 Parallel.ForEach 進行多線程遞迴計算subfolder數量
        Dim countingBag As New ConcurrentBag(Of Task(Of Integer))()     ' 使用 ConcurrentBag 來安全地收集每個子資料夾的數量
        Parallel.ForEach(folder.Folders.Cast(Of Outlook.Folder)(),
                         Sub(subFolder As Outlook.Folder)
                             'countingBag.Add(GetTotalFolderCountAsync(subFolder))
                             countingBag.Add(GetFolderCountAll(subFolder))
                         End Sub)
        Dim results = Await Task.WhenAll(countingBag)   ' 等待所有平行出去收集的數量都確定回來了
        totalSubCount += results.Sum()                  ' 再將回傳的各個子資料夾的數量加總

        _folderCountCache.TryAdd(folder, totalSubCount)
        ' ✅ 2026-03-16 移除多餘的 Try/Catch：ConcurrentDictionary.TryAdd 本身不拋例外 (只回傳 True/False) 
        ' 原本是從 .Add() 時代留下的防護，改 TryAdd 後應一併移除

        Return totalSubCount
    End Function
    Private Sub EnterFolderMenuItem(sender As Object, e As EventArgs)
        Dbg("開始：", sender.Name)
        Tab1_EnterSelectedFolder(ListView1.SelectedItems(0))
    End Sub
    Private Function GetFolderSizeOld_Async(folder As Outlook.Folder) As Long
        Dbg("開始：", folder.Name)

        Dim totalSize As Long = 0
        Dim folderItems As Outlook.Items = Nothing
        Try
            folderItems = folder.Items          ' ✅ 先取出 Items 物件，才能在 Finally 釋放
            For Each item As Object In folderItems
                Try
                    ' todo: 有好幾處都分別使用mailItem.Size或MAPI屬性來讀取資料夾大小, 是否可以統一成為一個函數供各處呼叫, 並且讀完直接加入快取? (先給我看你的建議, 先不要改)
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
            If folderItems IsNot Nothing Then Marshal.ReleaseComObject(folderItems)  ' ✅ Items 集合釋放
        End Try
        Return totalSize

    End Function
    Private Function GetFolderByName(folderName As String) As Outlook.Folder
        Dbg("開始：", folderName)
        ' 從目前 TreeView1.SelectedNode 的folderName, 回傳該資料夾的outlook.folder物件 (記錄在TreeNode的Tag)
        Dim node As TreeNode = FindNodeByName(TreeView1.SelectedNode, folderName)
        Return If(node Is Nothing, Nothing, node.Tag)
        ' 上面這行也很漂亮, 找找還有沒有類似的情況可以套用
    End Function
    Private Function FindNodeByName(selectedNode As TreeNode, ByVal findName As String) As TreeNode
        Dbg("開始：", selectedNode.Name)
        selectedNode.Expand()

        ' todo: 目前的實現是每次比較都把 " - " 去掉再比對，
        ' 建議改成在一開始就把 findName 處理成 cleanName，然後在整個遞迴過程中都使用 cleanName 來比對，
        ' 這樣就不需要每次都呼叫 Replace 了，效能應該會更好... 嗎??
        ' GetFolderByName()跟FindNodeByName()有更好的寫法嗎? 更快的或是更簡潔的?
        Dim cleanName As String = findName.Replace(" - ", "")
        For Each node As TreeNode In selectedNode.Nodes
            If node.Text = cleanName Then Return node

            Dim foundNode As TreeNode = FindNodeByName(node, findName)  ' 遞迴往下搜尋直到符合才return，找到就不再繼續往下搜尋了
            If foundNode IsNot Nothing Then Return foundNode
        Next
        Return Nothing
    End Function
    Private Function FindLiSVItemByName(listview As ListView, itemName As String) As ListViewItem
        Dbg("開始：", listview.Name)
        For Each item As ListViewItem In listview.Items
            If item.Text.Replace(" - ", "") = itemName.Replace(" - ", "") Then Return item
        Next : Return Nothing
    End Function
    Private Function FindNodeOrItemByName(ByVal nodesOrItems As IEnumerable, ByVal itemName As String) As Object
        Dbg("開始：", itemName)
        For Each item As Object In nodesOrItems
            Dim text As String = If(TypeOf item Is TreeNode, DirectCast(item, TreeNode).Text, DirectCast(item, ListViewItem).Text)
            If text.Replace(" - ", "") = itemName.Replace(" - ", "") Then Return item
        Next
        Return Nothing
    End Function
#End Region
#End Region

#Region "■ 05 Tab2：依日期統計"
    ' ==============================================================
    ' 重構目標: COM/UI/流程邏輯與業務分離清晰分層，去除全域狀態，優化快取機制
    ' 1. 分層架構: 將原本混在一起的程式碼重構成三個明確的層次
    '    - Layer 1 (UI 事件層)    : 回應使用者操作，組裝參數後交給 L2 執行，最後把結果交給顯示函數
    '    - Layer 2 (流程協調層)   : BFS 遍歷 folderList，管理快取，驅動 L3 計算，合併結果，回報進度
    '    - Layer 3 (COM 資料層)   : 對 Outlook 發出 COM 呼叫，回傳單一資料夾的年份郵件分佈
    ' 2. 去除全域狀態: 原本的 _intTotalMailCount 和 _intProcessedCount 全域變數已改成局部變數，避免多次點選時的計數錯亂
    ' 3. 優化快取機制: 快取的 key 改為純字串 FolderPath，避免 COM 物件當 key 導致 RCW 殘留問題；快取只存單一資料夾的結果，由 L2 負責合併
    ' 4. 進度回報改為 callback 機制: L2 執行統計時，透過 onProgress callback 回報已處理的郵件數和總郵件數，L1 負責更新 UI 顯示，保持分層乾淨
    ' by: Claude AI (2026/3/10)
    ' ==============================================================
    '
    ' 替換說明: 
    '   以下程式碼完整取代 Tab2 相關的所有邏輯函數。
    '   請同時刪除以下舊的函數與宣告: 
    '     - Private _intTotalMailCount As Integer   (全域變數宣告，已改成局部)
    '     - Private _intProcessedCount As Integer   (全域變數宣告，已改成局部)
    '     - TreeView2_AfterSelect()                 (已重寫)
    '     - SimTree2_AfterSelect()                  (已重寫，不再 commented out)
    '     - CheckSub2_CheckedChanged()              (已重寫)
    '     - GetYearCountsAsync_CL()                 (已由 ComputeYearCounts 取代)
    '     - CountMailByYearAsync_CL2()              (已由 GetYearCountsForFolderAsync 取代)
    '     - UpdateCounterProgress()                 (已改成 callback 機制，函數可刪除)
    '     - UpdateTab2Status()                          (簽章已更改，請替換)
    '
    '   以下函數不需要改動，保留原有: 
    '     - BuildFilterDateRangeTab2()
    '     - Find1stYear()
    '     - MergeDictionaries()
    '     - ShowTab2Result()
    '     - UpdateChart2()
    '     - TreeView2_MouseClick()
    '     - SimTree2_MouseClick()
    '     - MenuItem1_Click(), MenuItem2_Click(), ToggleTreeViewSelectMode()
    '     - ListView2_MouseDoubleClick()
    '     - Chart2_MouseMove(), Chart2_MouseLeave()
    '
    ' 分層架構: 
    '   Layer 1 (UI 事件層)    : TreeView2_AfterSelect, SimTree2_AfterSelect, CheckSub2_CheckedChanged, ShowTab2Result, UpdateTab2Status
    '   Layer 2 (流程協調層)   : ComputeYearCounts
    '   Layer 3 (COM 資料層)   : GetYearCountsForFolderAsync
    ' ==============================================================
#Region "  ├ L1 UI事件層"
    Private Async Sub SimTree2_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles SimTree2.AfterSelect
        ' ---------------------------------------------------------------
        ' === Layer 1: UI 事件層 ===
        ' 職責: 回應使用者操作，組裝參數後交給 L2 執行，最後把結果交給顯示函數
        ' 規則: 不做業務計算，不直接碰 COM，只傳達意圖
        ' ---------------------------------------------------------------
        ' SimTree2_AfterSelect：多選模式 SimTree2 的節點點選事件, 完整替換舊版
        ' 與 TreeView2_AfterSelect 對齊，補上月份展開所需的狀態賦值
        ' 支援 Ctrl+Click 多選，每個選定節點各自 BFS 展開後合併統計
        ' ---------------------------------------------------------------
        Dbg("開始：", sender.Name)

        Dim stopwatch As New Stopwatch() : stopwatch.Start()    ' 開始計時，初始化畫面狀態
        lblStatus1.Text = "" : lblStatus2.Text = "" : Cursor = Cursors.WaitCursor
        _cancelRequested = False                                ' ✅ reset ESC 旗標

        Dim selectedNodes As List(Of TreeNode) = SimTree2.SelectedNodes ' 取得 SimTree2 多選清單 (SelectedNodes 是 SimTree 提供的 List(Of TreeNode))
        If selectedNodes Is Nothing OrElse selectedNodes.Count = 0 Then ' 選擇節點為空，直接結束
            Cursor = Cursors.Default
            Dbg("End (no selected nodes): SimTree2_AfterSelect()")
            Return
        End If

        Dim targetFolderList =                                  ' 把所有已選 TreeNode 的 Tag 轉換成 Outlook.Folder，過濾掉無效節點
            selectedNodes.Select(Function(n) TryCast(n.Tag, Outlook.Folder)).Where(Function(f) f IsNot Nothing).ToList()

        If targetFolderList.Count = 0 Then                      ' 如果沒有任何有效的資料夾 (List.Count=0) 就直接結束
            Cursor = Cursors.Default
            Dbg("End (all nodes invalid): SimTree2_AfterSelect()")
            Return
        End If

        Dim folderList As New List(Of Outlook.Folder)           ' 對每個選定的根資料夾執行 BFS，合併成一個完整的目標資料夾清單
        Dim addedPaths As New HashSet(Of String)                ' 用 HashSet(Of String) 以 FolderPath 去重，避免使用者選到父子資料夾時重複計算
        For Each rootFolder As Outlook.Folder In targetFolderList
            For Each f As Outlook.Folder In GetSubFolderList(rootFolder, CheckSub2.Checked)
                If addedPaths.Add(f.FolderPath) Then folderList.Add(f)
                ' 若Add() 回傳 False 代表已存在，自動去重
            Next
        Next

        _tab2FolderList = folderList                            ' ✅ 記住本次統計的資料夾清單，供 ListView2 月份展開 (ShowMonthView) 使用
        _tab2IsMonthView = False                                ' 切換選取時，重置視圖狀態為年度視圖

        'Dim totalMailCount As Integer =                                                     ' 計算所有選定根資料夾的郵件總數作為進度分母
        '    If(CheckSub2.Checked, rootFolders.Sum(Function(f) GetMailCountRecursive(f)),    ' CheckSub2.Checked = True  → 含子資料夾: 各自完整子樹的總和
        '                          rootFolders.Sum(Function(f) GetMailCount(f)))             ' CheckSub2.Checked = False → 只算選定的那一層

        '' 2026/3/20, 重寫了底層GetMailCountAll() 但是不知為何效能還是比不過現在上面的遞迴版本??
        ' 原因: 原版遞迴只走一遍 COM 資料夾樹，新版走了兩遍COM：
        ' 第一遍：GetSubFolderList()    → BFS 遍歷，存取每個 folder.Folders
        ' 第二遍：For Each allFolders   → GetMailCount() 再讀每個資料夾一次

        ' 計算所有選定根資料夾的郵件總數，作為 ComputeYearCounts 進度條的分母
        ' GetMailCountAll 是 Async，不能像上面放在 LINQ Sum lambda 裡，改用明確的 For Each + Await (光是這二點, 效能也差一大截)
        ' 2026/3/20, 再次嚐試把GetMailCountAll() 改成平行處理, 效能回復到原有的遞迴函數, 但速度並不穩定
        Dim totalMailCount As Long = 0
        For Each rf As Outlook.Folder In targetFolderList
            If CheckSub2.Checked Then
                Dim c As Long = Await GetMailCountAll(rf) : If c > 0 Then totalMailCount += c   ' -1 表示讀取失敗，略過不累加
            Else
                Dim c As Integer = GetMailCount(rf) : If c > 0 Then totalMailCount += c         ' 只算本層，L3 同步函數，不需要 Await
            End If
        Next

        Dim yearCounts As ConcurrentDictionary(Of Integer, Integer) =   ' 呼叫 L2 流程協調層執行統計 (跟單選模式走一樣的路徑，只是 folderList 不同) 
            Await ComputeYearCounts(folderList, totalMailCount,         ' 進度更新透過 callback 傳回，L2 不直接碰 lblStatus1，保持分層乾淨
                Sub(processed As Integer, total As Integer) lblStatus1.Text = $"正在統計全部 {total} 郵件裡的 {processed} 封")

        stopwatch.Stop()                                                ' ✅ 統計完成後才停錶 (與 TreeView2_AfterSelect 一致) 
        If _cancelRequested Then                                        ' ✅ ESC 中斷：還原 UI 狀態
            lblStatus1.Text = "" : lblStatus2.Text = "已中斷。"
            sender.Enabled = True : sender.Focus()
            Cursor = Cursors.Default : Return
        End If
        ShowTab2Result(yearCounts)                                      ' 顯示結果到 ListView2 和 Chart2
        UpdateTab2Status(yearCounts, stopwatch.Elapsed)                 ' 顯示執行時間與處理速度到 lblStatus2

        sender.Enabled = True : sender.Focus() : Cursor = Cursors.Default
        Dbg("結束：", sender.Name)
    End Sub
    Private Sub TreeView2_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TreeView2.AfterSelect
        ' ---------------------------------------------------------------
        ' === Layer 1: UI 事件層 ===
        ' 職責: 回應使用者操作，組裝參數後交給 L2 執行，最後把結果交給顯示函數
        ' 規則: 不做業務計算，不直接碰 COM，只傳達意圖
        ' ---------------------------------------------------------------
        Dbg("開始：", sender.Name)
        '_cancelRequested = False    ' ✅ reset ESC 旗標

        '' 開始計時，初始化畫面狀態
        'Dim stopwatch As New Stopwatch() : stopwatch.Start()
        'lblStatus1.Text = "" : lblStatus2.Text = "" : Cursor = Cursors.WaitCursor

        '' 取得使用者點選的資料夾，Tag 存放的是 Outlook.Folder COM 物件
        'Dim rootFolder As Outlook.Folder = TryCast(e.Node.Tag, Outlook.Folder)
        'If rootFolder Is Nothing Then
        '    Cursor = Cursors.Default
        '    dbg("End (rootFolder is Nothing): TreeView2_AfterSelect()")
        '    Return
        'End If

        '' BFS 展開目標資料夾清單 (複用 Tab3 的 GetSubFolderList) 
        '' CheckSub2.Checked = True  → 把整棵子樹全部展開成一個清單 → 含子資料夾: 整棵樹的總數
        '' CheckSub2.Checked = False → 清單只有rootFolder 自己一個 → 單一資料夾: 只算自己這層
        'Dim folderList As List(Of Outlook.Folder) = GetSubFolderList(rootFolder, CheckSub2.Checked)
        '_tab2FolderList = folderList     ' 記住本次統計使用的資料夾清單，供月份展開用
        '_tab2IsMonthView = False         ' 切換資料夾時，重置視圖狀態為年度視圖

        '' 取得總郵件數作為進度條的分母: GetMailCountByMAPINew 已包含所有子資料夾的郵件數，並有 mailCountCache 加速
        'Dim folderItems As Outlook.Items = rootFolder.Items
        'Dim totalMailCount As Integer = If(CheckSub2.Checked, GetMailCountByMAPINew(rootFolder), folderItems.Count)
        'Marshal.ReleaseComObject(folderItems)   ' 釋放 COM 物件，避免 RCW 殘留

        '' 呼叫 L2 流程協調層執行統計
        '' 進度更新透過 callback 傳回，L2 不直接碰 lblStatus1，保持分層乾淨
        '' L2 執行完會回傳整合好的 yearCounts 結果，L1 負責把結果交給 ShowTab2Result 顯示，並呼叫 UpdateTab2Status 顯示執行時間和速度
        '' 下面這句的語法是 VB.NET 的 Lambda Sub，代表定義一個匿名 Sub 來接收 L2 的進度回報，然後更新 lblStatus1 的文字顯示
        '' (這句好像有點複雜，可以拆成兩行來寫，先定義一個 Sub 變數來接收 Lambda，再把它傳給 L2)
        'Dim yearCounts As ConcurrentDictionary(Of Integer, Integer) =
        '    Await ComputeYearCounts(folderList, totalMailCount,
        '        Sub(total As Integer, processed As Integer) lblStatus1.Text = $"正在統計全部 {total} 郵件裡的 {processed} 封") ' ✅ UI 更新在 UI 執行緒，透過 callback 傳回

        'stopwatch.Stop()                            ' 統計完成
        'If _cancelRequested Then                    ' ✅ ESC 中斷：還原 UI 狀態，不更新 ListView/Chart
        '    lblStatus1.Text = "" : lblStatus2.Text = "已中斷。" : sender.Enabled = True : sender.Focus() : Cursor = Cursors.Default : Return
        'End If
        'ShowTab2Result(yearCounts)                  ' 顯示結果到 ListView2 和 Chart2
        'UpdateTab2Status(yearCounts, stopwatch.Elapsed) ' 顯示執行時間與處理速度到 lblStatus2

        sender.Enabled = True : sender.Focus() : Cursor = Cursors.Default
        Dbg("結束：")
    End Sub
    Private Sub ListView2_GotFocus(sender As Object, e As EventArgs) Handles ListView2.GotFocus
        If ListView2.SelectedItems.Count = 0 AndAlso ListView2.Items.Count > 0 Then
            ListView2.Items(0).Selected = True
        End If
    End Sub
    Private Sub ListView2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListView2.SelectedIndexChanged
        ' ---------------------------------------------------------------
        ' ListView2 選取變更 ↔ Chart2 對應長條同步高亮
        ' 年度視圖：選取某年 → 高亮 Chart2 中對應年份的長條
        ' 月份視圖：選取某月 → 高亮 Chart2 中對應月份的長條
        ' 與 Chart2_MouseMove 共用 _lastHoveredPointIndex，確保兩者高亮互斥、不累積
        ' 注意：Chart2_MouseLeave 會清掉 _lastHoveredPointIndex，所以滑鼠離開圖表後
        '       ListView 的選取高亮也會消失 — 這是可接受的行為 (簡化狀態管理) 
        ' 2026-03-18, by Claude.ai
        ' ---------------------------------------------------------------

        If ListView2.SelectedItems.Count = 0 Then Return
        If Chart2.Series.Count = 0 OrElse Chart2.Series(0).Points.Count = 0 Then Return  ' Chart 尚未載入資料，直接結束

        Dim selectedItem As ListViewItem = ListView2.SelectedItems(0)
        Dim selectedText As String = selectedItem.Text.Trim()

        ' ── 找出目標 DataPoint index ──
        Dim targetIndex As Integer = -1

        If Not _tab2IsMonthView Then
            ' 年度視圖：selectedText = "2023"，直接解析成整數，比對 pt.XValue
            Dim selectedYear As Integer = 0
            If Not Integer.TryParse(selectedText, selectedYear) Then Return  ' 非數字 (理論上不應發生) 
            For i = 0 To Chart2.Series(0).Points.Count - 1
                If CInt(Chart2.Series(0).Points(i).XValue) = selectedYear Then targetIndex = i : Exit For
            Next

        Else
            ' 月份視圖：selectedText = "2024 /  01月"
            ' 過濾掉特殊列：返回列 (Tag="BACK") 和標題列 (包含 "──") 
            If selectedItem.Tag IsNot Nothing AndAlso selectedItem.Tag.ToString() = "BACK" Then Return
            If selectedText.Contains("──") Then Return

            ' 從文字尾端的 "MM月" 提取月份數字 (從 "月" 往前讀取連續數字字元) 
            Dim moonIdx As Integer = selectedText.IndexOf("月")
            If moonIdx < 0 Then Return  ' 沒有 "月" 字 (不應發生，但防護) 
            Dim numStr As String = ""
            For k As Integer = moonIdx - 1 To 0 Step -1
                If Char.IsDigit(selectedText(k)) Then numStr = selectedText(k) & numStr Else Exit For
            Next
            Dim selectedMonth As Integer = 0
            If Not Integer.TryParse(numStr, selectedMonth) OrElse selectedMonth < 1 OrElse selectedMonth > 12 Then Return
            ' UpdateChart2ForMonths 依 1~12 月順序加入 DataPoints，月份N = index N-1
            targetIndex = selectedMonth - 1
        End If

        If targetIndex < 0 OrElse targetIndex >= Chart2.Series(0).Points.Count Then Return

        ' ── 還原上一個高亮，套用新的高亮 ──
        If _lastHoveredPointIndex >= 0 AndAlso _lastHoveredPointIndex < Chart2.Series(0).Points.Count Then
            Chart2.Series(0).Points(_lastHoveredPointIndex).Color = Color.Empty  ' 還原成 Series 預設色
        End If
        Chart2.Series(0).Points(targetIndex).Color = Color.Red  ' 與 MouseMove hover 同色，統一體驗
        _lastHoveredPointIndex = targetIndex
        Chart2.Refresh()

        ' 2026/3/22, rewrited by Grok, but not working...
        'If Chart2.Series.Count = 0 OrElse Chart2.Series(0).Points.Count = 0 Then Return
        'If ListView2.SelectedItems.Count = 0 Then Return

        'Dim selText As String = ListView2.SelectedItems(0).Text.Trim()

        '' 清除舊 hover 高亮
        'If _lastHoveredPointIndex >= 0 Then
        '    Chart2.Series(0).Points(_lastHoveredPointIndex).Color = Color.SteelBlue
        '    _lastHoveredPointIndex = -1
        'End If

        'For Each pt As DataPoint In Chart2.Series(0).Points
        '    Dim match As Boolean = False
        '    If _tab2IsMonthView Then
        '        match = pt.AxisLabel.Contains(selText) OrElse pt.AxisLabel = selText & "月"
        '    Else
        '        match = CInt(pt.XValue).ToString() = selText
        '    End If

        '    If match Then
        '        pt.Color = Color.OrangeRed
        '        _lastHoveredPointIndex = pt.PointIndex 
        '        Chart2.Invalidate()
        '        Exit For
        '    End If
        'Next
    End Sub
    Private Async Sub ListView2_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles ListView2.MouseDoubleClick
        ' ---------------------------------------------------------------
        ' ListView2 雙擊事件 (完整替換舊版) 
        ' 年度視圖：雙擊某一年 → 展開顯示該年的月份分佈 + 更新 Chart2
        ' 月份視圖：雙擊「← 返回」 → 回到年度視圖 + 更新 Chart2
        ' ---------------------------------------------------------------
        Dbg("開始：", sender.Name)

        Dim clickedItem As ListViewItem = ListView2.GetItemAt(e.X, e.Y)
        If clickedItem Is Nothing Then Return

        ' 月份視圖 → 雙擊「← 返回年度統計」：回到年度視圖
        If _tab2IsMonthView AndAlso clickedItem.Tag IsNot Nothing AndAlso
                                    clickedItem.Tag.ToString() = "BACK" Then
            Await ShowYearView() : Return
        End If

        ' 年度視圖 → 雙擊某一年：展開為月份視圖
        ' 2026/3/16: monthCountsCache 已在 GetMonthCountsForYear 內部實作，重複展開同一年直接命中快取
        Dim selectedYear As Integer = 0
        If Not Integer.TryParse(clickedItem.Text.Trim(), selectedYear) Then Return
        If _tab2FolderList Is Nothing OrElse _tab2FolderList.Count = 0 Then Return
        Await ShowMonthView(selectedYear)

        Dbg("結束：", $"year={selectedYear}")
    End Sub
    Private Sub Chart2_MouseMove(sender As Object, e As MouseEventArgs) Handles Chart2.MouseMove
        ' ✅ 改用 MouseMove，滑鼠移動時持續觸發，才能追蹤到每個長條
        Dim chart As Chart = CType(sender, Chart)
        If chart.Series.Count = 0 OrElse chart.Series(0).Points.Count = 0 Then Return

        Dim hit As HitTestResult = chart.HitTest(e.X, e.Y)

        If hit.ChartElementType = ChartElementType.DataPoint Then
            Dim pointIndex As Integer = hit.PointIndex
            If pointIndex = _lastHoveredPointIndex Then Return ' 如果跟上次是同一個點就不重複處理，避免閃爍

            ' ✅ 先把上一個點的顏色還原
            If _lastHoveredPointIndex >= 0 AndAlso
                _lastHoveredPointIndex < chart.Series(0).Points.Count Then
                chart.Series(0).Points(_lastHoveredPointIndex).Color = Color.Empty  ' Empty = 還原成 Series 預設色
            End If

            ' ✅ 把目前這個點變成紅色
            chart.Series(0).Points(pointIndex).Color = Color.Red
            _lastHoveredPointIndex = pointIndex

            ' ✅ 顯示數值 (用 Series 的 ToolTip 屬性，Chart 控制項內建支援)
            Dim dataPoint As DataPoint = chart.Series(0).Points(pointIndex)
            chart.Series(0).ToolTip = $"年份: {dataPoint.AxisLabel}, 數量: {dataPoint.YValues(0):###,###,##0}"
            ' todo: tooltip 的位置被滑鼠遮住, 要往上移動一點

        Else
            ' 滑鼠離開所有長條，還原上一個點
            If _lastHoveredPointIndex >= 0 AndAlso
                _lastHoveredPointIndex < chart.Series(0).Points.Count Then
                chart.Series(0).Points(_lastHoveredPointIndex).Color = Color.Empty
                _lastHoveredPointIndex = -1
            End If
            chart.Series(0).ToolTip = String.Empty
        End If
    End Sub
    Private Sub Chart2_MouseClick(sender As Object, e As MouseEventArgs) Handles Chart2.MouseClick
        ' ---------------------------------------------------------------
        ' Chart2 點擊長條 → 同步高亮 ListView2 對應的年份或月份列
        ' 反向對應：ListView2_SelectedIndexChanged 負責 ListView → Chart2
        ' 年度視圖：比對 pt.XValue (整數年份) 找 ListView2 中對應的年份列
        ' 月份視圖：pt.AxisLabel = "N月"，解析月份數字，找對應的月份列
        ' 設定 item.Selected = True 會觸發 ListView2_SelectedIndexChanged，
        ' 後者會再次把 Chart2 同一條塗紅 — 因為是同一條，行為是 idempotent 不會閃爍
        ' 2026-03-18, by Claude.ai
        ' ---------------------------------------------------------------
        Dbg("開始：", sender.Name)
        If Chart2.Series.Count = 0 OrElse Chart2.Series(0).Points.Count = 0 Then Return

        Dim hit As HitTestResult = Chart2.HitTest(e.X, e.Y)
        If hit.ChartElementType <> ChartElementType.DataPoint Then Return

        Dim pt As DataPoint = Chart2.Series(0).Points(hit.PointIndex)

        ' ── 根據目前視圖找目標 ListViewItem ──
        Dim targetItem As ListViewItem = Nothing

        If Not _tab2IsMonthView Then
            ' 年度視圖：pt.XValue = 年份 (Double，轉 Integer 比對) 
            Dim clickedYear As Integer = CInt(pt.XValue)
            For Each item As ListViewItem In ListView2.Items
                Dim yr As Integer = 0
                If Integer.TryParse(item.Text.Trim(), yr) AndAlso yr = clickedYear Then
                    targetItem = item : Exit For
                End If
            Next

        Else
            ' 月份視圖：pt.AxisLabel = "N月"，解析出月份數字
            Dim label As String = pt.AxisLabel  ' e.g. "3月"
            Dim moonIdx As Integer = label.IndexOf("月")
            If moonIdx < 0 Then Return
            Dim monthNum As Integer = 0
            If Not Integer.TryParse(label.Substring(0, moonIdx), monthNum) Then Return

            ' ListView2 月份列的文字格式："{year} /  MM月"，只要月份數字符合就算
            Dim monthStr As String = monthNum.ToString("D2") & "月"  ' e.g. "03月"
            For Each item As ListViewItem In ListView2.Items
                If item.Text.Contains(monthStr) AndAlso
               (item.Tag Is Nothing OrElse item.Tag.ToString() <> "BACK") Then
                    targetItem = item : Exit For
                End If
            Next
        End If

        If targetItem Is Nothing Then Return

        For Each item As ListViewItem In ListView2.Items    ' ✅ 先清除所有現有選取，避免多次點擊累積多個 highlighted item
            item.Selected = False                           ' 改用逐一設 Selected = False，安全可靠
        Next                                                ' 不可用 ListView.SelectedItems.Clear() (會丟 NotSupportedException) 

        ' ── 選取並捲動到目標列 (會觸發 SelectedIndexChanged 同步塗色) ──
        targetItem.Selected = True
        targetItem.Focused = True
        ListView2.Focus()
        targetItem.EnsureVisible()
    End Sub
    Private Sub Chart2_MouseLeave(sender As Object, e As EventArgs) Handles Chart2.MouseLeave
        If _lastHoveredPointIndex >= 0 AndAlso Chart2.Series.Count > 0 AndAlso
            _lastHoveredPointIndex < Chart2.Series(0).Points.Count Then
            Chart2.Series(0).Points(_lastHoveredPointIndex).Color = Color.Empty
            _lastHoveredPointIndex = -1
        End If
        Chart2.Series(0).ToolTip = String.Empty
        Chart2.Refresh()        ' ✅ 同步重繪，立刻執行
        'Me.BeginInvoke(Sub() Chart2.Invalidate())  ' ← 取代 Refresh()，等內部狀態穩定再重繪
        'Await Task.Yield()     ' ✅ 讓出 UI 執行緒，確保 MouseLeave 事件處理器能完成剩餘的還原操作
        'Await Task.Delay(0)    ' ✅ 小延遲，確保 Chart 控制項有機會處理完 MouseLeave 事件的內部狀態更新，避免因為 Chart 控制項內部狀態還沒更新而導致的顏色還原失效
    End Sub
    Private Sub CheckSub2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckSub2.CheckedChanged
        ' CheckSub2 勾選狀態改變時，以目前選定的節點重新觸發統計
        ' 判斷目前顯示的是 TreeView2 (單選) 還是 SimTree2 (多選) ，各自觸發對應的事件
        ' todo: 若SimTree多選控制項已完整實作，TreeView2的單選模式已經不再使用，考慮直接移除TreeView2相關的程式碼，專注在SimTree2的多選功能上，避免維護兩套類似功能的程式碼造成混亂
        Dbg("開始：", sender.Name)

        ' SimTree2 (多選模式) 可見時，重新觸發目前選定節點的統計
        ' 傳第一個選定節點作為 TreeViewEventArgs 的參數 (AfterSelect 內部會自己取 SelectedNodes 清單) 
        If SimTree2.Visible Then
            Dim selectedNodes As List(Of TreeNode) = SimTree2.SelectedNodes
            If selectedNodes IsNot Nothing AndAlso selectedNodes.Count > 0 Then SimTree2_AfterSelect(SimTree2, New TreeViewEventArgs(selectedNodes(0)))
        End If

        '' TreeView2 (單選模式) 可見時，重新觸發目前選定節點的統計
        '2026/3/11, 去除 TreeView2 的 AfterSelect 事件觸發，以SimTree2 為主的多選模式取代原本的單選模式，避免原有treeview2一直搶到焦點
        'If TreeView2.Visible Then
        '    Dim selectedNode As TreeNode = TreeView2.SelectedNode
        '    If selectedNode IsNot Nothing Then       TreeView2_AfterSelect(TreeView2, New TreeViewEventArgs(selectedNode))
        'End If
    End Sub
#End Region
#Region "  ├ L2 流程協調層"
    Private Async Function ComputeYearCounts(folderList As List(Of Outlook.Folder), totalMailCount As Integer, onProgress As Action(Of Integer, Integer)) As Task(Of ConcurrentDictionary(Of Integer, Integer))
        ' ---------------------------------------------------------------
        ' === Layer 2: 流程協調層 ===
        ' 職責: BFS 遍歷 folderList，管理快取，驅動 L3 計算，合併結果，回報進度
        '       逐資料夾計算年份統計並合併，是 Tab2 所有統計流程的唯一入口
        ' 規則: 不直接碰 UI 控制項 (lblStatus1 等) ，進度透過 onProgress callback 傳出, 自己不會知道上一層是單選還是多選，只知道接受傳入的 folderList 清單
        '
        ' 參數: 
        '   folderList    : 由 L1 組裝好的目標資料夾清單 (已包含 BFS 展開結果) 
        '   totalMailCount: 總郵件數，用來計算進度百分比的分母
        '   onProgress    : 進度 callback，每處理完一個資料夾呼叫一次，回傳 (已處理, 總數)
        ' ---------------------------------------------------------------
        ' todo: tab2在跑到某些資料夾時, 會發生很多COM exception,
        ' 尤其是在日誌, 工作, 行事曆, 這種非郵件目錄
        ' 但在有些老舊郵件目錄也會偶而出現, 可能裡面混入了一些其他 "不是mailitem" 的項目, 要如何篩選掉? (包括計算數目, 以及總數)
        ' tab1在計數不會, 但在計算foldersize的時候會.

        Dbg("開始：ComputeYearCounts(), 資料夾數量: ", folderList.Count)

        Dim merged As New ConcurrentDictionary(Of Integer, Integer)
        Dim processedCount As Integer = 0       ' ✅ 局部計數器，取代全域的 _intProcessedCount 和 _intTotalMailCount, 不會被其他事件汙染，快速點選時不會計數錯亂

        For Each folder As Outlook.Folder In folderList
            If _cancelRequested Then Exit For   ' ✅ ESC 中斷：Exit For 回傳已算的部分結果，L1 會偵測到 _cancelRequested 並跳過顯示
            'If folder.FolderPath.Contains("\行事曆") OrElse folder.FolderPath.Contains("\Task Done") Then Exit For  ' 這二種folder會抛出超多未知的COM Exception 怎麼辦? 扣除的話又統計錯誤

            Dim folderResult As ConcurrentDictionary(Of Integer, Integer)
            Dim cacheKey As String = folder.FolderPath
            ' 快取 key 只用 FolderPath (純字串) ，不用 COM 物件當 key
            ' 理由: COM 物件當 key 會造成 RCW 殘留無法被 GC 回收 (已知架構問題) 
            ' 只快取「單一資料夾」的結果，合併邏輯由本層負責

            If _yearCountsCache.ContainsKey(cacheKey) Then   ' ✅ 快取命中: 直接取結果，完全不再讀 COM
                Dbg("Cache Hit: ", folder.Name)
                folderResult = _yearCountsCache(cacheKey)
            Else                                            ' ❌ 快取未命中: 呼叫 L3 COM 資料層，計算這個資料夾的年份分佈
                Dbg("Cache miss: ", folder.Name)
                folderResult = Await GetYearCountsForFolder(folder)
                _yearCountsCache(cacheKey) = folderResult    ' 計算完成後存入快取，下次點選同一資料夾直接命中
                ' ✅ 用 "=" 賦值 (非 .Add()) ，有重複 key 時直接覆蓋，不拋例外
            End If

            merged = MergeDictionaries(merged, folderResult)  ' 把這個資料夾的結果合併到總計 (純 .NET 運算，不碰 COM) 
            processedCount += folderResult.Values.Sum()     ' 累加已處理郵件數，透過 callback 通知 L1 更新進度顯示
            onProgress(processedCount, totalMailCount)      ' onProgress callback
            If processedCount Mod 3 = 0 Then Await Task.Delay(0) ' ✅ 每處理完3個資料夾讓出一次控制權，保持 UI 可回應, 大量子資料夾時，UI 仍然可以回應使用者操作
        Next

        Dbg("結束：ComputeYearCounts()", $"年份數:{merged.Count}, 總郵件:{merged.Values.Sum}")
        Return merged
    End Function
    Private Async Function GetYearCountsForFolder(folder As Outlook.Folder) As Task(Of ConcurrentDictionary(Of Integer, Integer))
        ' ---------------------------------------------------------------
        ' === Layer 3: COM 資料層 ===
        ' 職責: 對 Outlook 發出 COM 呼叫，回傳單一資料夾的年份郵件分佈
        ' 規則: 不遞迴、不碰 UI、不修改任何全域狀態，
        '       只做一件事: 詢問 Outlook 某資料夾每年有幾封郵件，回傳結果
        '       不遞迴、不知道上層的進度計數、不碰 UI，完全純粹的資料查詢函數
        ' 2026/3/24 by AntiGravity: 從逐年 Restrict 改為 GetTable + GetArray 一次讀完再記憶體分組
        '   原本每年一次 Restrict + Items.Count = ~30 次 COM call
        '   現在 1 次 GetTable + ceil(N/1000) 次 GetArray，大幅減少 COM 跨程序呼叫
        ' ---------------------------------------------------------------
        Dbg("開始：GetYearCountsForFolder()", folder.Name)

        ' 2026/3/11再次重構: 優化 COM 呼叫，減少 RCW 物件積累，提升效能和穩定性
        'Dim folderItems As Outlook.Items = Nothing
        Dim yearCounts As New ConcurrentDictionary(Of Integer, Integer)
        Const BATCH_SIZE As Integer = 1000  ' 2026/3/24 by AntiGravity: 每次批量讀取的筆數
        Dim table As Outlook.Table = Nothing

        Try
            ' 2026/3/24 by AntiGravity: 改用 GetTable + GetArray 取代逐年 Restrict
            ' 只讀 ReceivedTime 一欄，最小化每 row 的傳輸量
            table = folder.GetTable()
            table.Columns.RemoveAll()
            table.Columns.Add("ReceivedTime")   ' 欄位索引 0

            Do While Not table.EndOfTable
                If _cancelRequested Then Exit Do
                Dim arr As Object = table.GetArray(BATCH_SIZE)
                If arr Is Nothing Then Exit Do
                Dim data(,) As Object = DirectCast(arr, Object(,))
                Dim rows As Integer = data.GetUpperBound(0) + 1

                For r As Integer = 0 To rows - 1
                    Try
                        Dim val As Object = data(r, 0)
                        If val Is Nothing OrElse IsDBNull(val) Then Continue For
                        Dim receivedTime As DateTime = CDate(val)
                        Dim year As Integer = receivedTime.Year
                        If year > 0 AndAlso year <= Date.Today.Year Then
                            yearCounts.AddOrUpdate(year, 1, Function(k, v) v + 1)
                        End If
                    Catch ex As System.Exception
                        ' 個別 row 讀取失敗不影響整體統計
                    End Try
                Next
                Await Task.Yield()  ' ✅ 每批次讓出一次，讓 ESC 按鍵能被處理
            Loop

        Catch ex As System.Exception
            Dbg("GetYearCountsForFolder Error: ", folder.Name & " - " & ex.Message)
        Finally
            If table IsNot Nothing Then Marshal.ReleaseComObject(table)
        End Try

        Await Task.Yield()   ' ✅ 函數結束前再讓出一次，確保畫面有機會更新
        Dbg("結束：GetYearCountsForFolder()", $"{folder.Name}, 年份數:{yearCounts.Count}")
        Return yearCounts

    End Function
    Private Async Function GetMonthCountsForYear(folder As Outlook.Folder, year As Integer) As Task(Of ConcurrentDictionary(Of Integer, Integer))
        ' ---------------------------------------------------------------
        ' GetMonthCountsForYear (完整替換舊版，加入快取和進度支援) 
        ' L3 COM 資料層：計算單一資料夾在指定年份中每個月的郵件數量
        ' 快取 key = FolderPath + "_" + year，與 yearCountsCache 的命名慣例一致
        ' 2026/3/24 by AntiGravity: 從逐月 Restrict 改為 GetTable + GetArray 一次讀完再記憶體分組
        '   原本 12 次 Restrict + 12 次 Items.Count = 24 次 COM call
        '   現在 1 次 GetTable (含日期範圍 filter) + ceil(N/1000) 次 GetArray
        ' ---------------------------------------------------------------
        Dbg("開始：GetMonthCountsForYear()", $"{folder.Name}, year={year}")

        ' ✅ 快取命中：直接回傳，不打任何 COM
        Dim cacheKey As String = folder.FolderPath & "_" & year.ToString()
        If _monthCountsCache.ContainsKey(cacheKey) Then
            Dbg("Cache Hit: GetMonthCountsForYear()", $"{folder.Name}, year={year}")
            Return _monthCountsCache(cacheKey)
        End If

        Dim monthCounts As New ConcurrentDictionary(Of Integer, Integer)
        Const BATCH_SIZE As Integer = 1000  ' 2026/3/24 by AntiGravity
        Dim table As Outlook.Table = Nothing

        Try
            ' 2026/3/24 by AntiGravity: 改用 GetTable + 日期範圍 DASL filter + GetArray
            ' 用整年的日期範圍一次篩選，不再逐月 Restrict
            Dim startDate As New Date(year, 1, 1, 0, 0, 0)
            Dim endDate As New Date(year, 12, 31, 23, 59, 59)
            Dim dateFilter As String = $"[ReceivedTime] >= '{startDate}' AND [ReceivedTime] <= '{endDate}'"

            table = folder.GetTable(dateFilter)
            table.Columns.RemoveAll()
            table.Columns.Add("ReceivedTime")   ' 欄位索引 0

            Do While Not table.EndOfTable
                Dim arr As Object = table.GetArray(BATCH_SIZE)
                If arr Is Nothing Then Exit Do
                Dim data(,) As Object = DirectCast(arr, Object(,))
                Dim rows As Integer = data.GetUpperBound(0) + 1

                For r As Integer = 0 To rows - 1
                    Try
                        Dim val As Object = data(r, 0)
                        If val Is Nothing OrElse IsDBNull(val) Then Continue For
                        Dim month As Integer = CDate(val).Month
                        monthCounts.AddOrUpdate(month, 1, Function(k, v) v + 1)
                    Catch ex As System.Exception
                        ' 個別 row 讀取失敗不影響整體統計
                    End Try
                Next
                Await Task.Yield()
            Loop

        Catch ex As System.Exception
            Dbg("GetMonthCountsForYear Error: ", folder.Name & $", year={year} - " & ex.Message)
        Finally
            If table IsNot Nothing Then Marshal.ReleaseComObject(table)
        End Try

        _monthCountsCache(cacheKey) = monthCounts    ' ✅ 第一次統計完, 一律存入快取，下次進入同一年份直接命中

        Dbg("結束：GetMonthCountsForYear()", $"{folder.Name}, 有郵件的月份數={monthCounts.Count}")
        Return monthCounts
    End Function
    Private Async Function ShowYearView() As Task
        ' ---------------------------------------------------------------
        ' 回到年度視圖 (返回按鈕、ESC 鍵都呼叫這裡) 
        ' ---------------------------------------------------------------
        Dbg("開始：ShowYearView: ")

        Dim yearToRestore As Integer = _tab2MonthViewYear  ' 先記住要回去的年份
        _tab2IsMonthView = False
        _tab2MonthViewYear = 0

        ' ★ 直接重算年度統計，若資料已在 yearCountsCache, 則ComputeYearCounts 快取全部命中瞬間完成 (< 5ms) 
        If _tab2FolderList IsNot Nothing AndAlso _tab2FolderList.Count > 0 Then
            Dim sw As New Stopwatch : sw.Start()
            ' ✅ 2026-03-17 ESC regression fix：
            '    ShowYearView 是「還原畫面」操作，不是「中斷操作」
            '    但 _cancelRequested = True 時 ComputeYearCounts 會 Exit For 回傳空 Dict
            '    → ShowTab2Result 收到空資料就清空 ListView2，畫面變空白
            '    解法：暫時保存並清除旗標，讓這次靜默重算順利完成；完成後還原旗標
            Dim savedCancel As Boolean = _cancelRequested
            _cancelRequested = False
            Dim yearCounts As ConcurrentDictionary(Of Integer, Integer) = Await ComputeYearCounts(_tab2FolderList, 0, Sub(a, b)
                                                                                                                      End Sub)
            _cancelRequested = savedCancel  ' 還原 (若使用者是在統計期間按 ESC，不影響後續操作的旗標) 
            sw.Stop()
            ShowTab2Result(yearCounts)
            UpdateTab2Status(yearCounts, sw.Elapsed)
        End If

        ' 回到年度視圖後，嘗試選定剛才進入前的那一年, 讓使用者感覺是「回到剛才看的地方」，而不是每次都回到頂部
        If yearToRestore > 0 AndAlso ListView2.Items.Count > 0 Then
            For Each item As ListViewItem In ListView2.Items
                Dim yr As Integer
                If Integer.TryParse(item.Text.Trim(), yr) AndAlso yr = yearToRestore Then
                    item.Selected = True : item.Focused = True : item.EnsureVisible()
                    ListView2.Focus() : Exit For
                End If
            Next
        End If

        Dbg("結束：ShowYearView: ")
    End Function
    Private Async Function ShowMonthView(selectedYear As Integer) As Task
        ' ---------------------------------------------------------------
        ' 顯示月份視圖 (年度視圖進入時呼叫，Enter 鍵也呼叫這裡) 
        ' 包含：進度顯示、快取、ListView2 月份清單、Chart2 月份長條圖、UpdateTab2Status
        ' ---------------------------------------------------------------
        Dbg("開始：ShowMonthView: ", selectedYear)

        _tab2IsMonthView = True
        _tab2MonthViewYear = selectedYear
        lblStatus1.Text = "" : lblStatus2.Text = "" : Cursor = Cursors.WaitCursor

        Dim sw As New Stopwatch() : sw.Start()
        Dim monthCounts As New ConcurrentDictionary(Of Integer, Integer)
        Dim totalFolders As Integer = _tab2FolderList.Count
        Dim processedFolders As Integer = 0

        For Each folder As Outlook.Folder In _tab2FolderList
            ' ✅ 進度顯示：每完成一個資料夾更新一次
            processedFolders += 1
            lblStatus1.Text = $"正在統計 {selectedYear} 年月份分佈... 讀取({processedFolders}/{totalFolders})個資料夾。"

            Dim folderMonthCounts As ConcurrentDictionary(Of Integer, Integer) = Await GetMonthCountsForYear(folder, selectedYear)
            monthCounts = MergeDictionaries(monthCounts, folderMonthCounts)
            Await Task.Yield()
        Next
        sw.Stop()
        Cursor = Cursors.Default

        ' ---------------------------------------------------------------
        ' 顯示月份清單到 ListView2
        ' ---------------------------------------------------------------
        ListView2.BeginUpdate()
        ListView2.Items.Clear()

        ' 第一行：返回按鈕
        Dim backItem As New ListViewItem("← 返回年度統計")
        backItem.SubItems.Add("") : backItem.Tag = "BACK"
        backItem.ForeColor = Color.Gray
        backItem.Font = New Font(ListView2.Font, System.Drawing.FontStyle.Italic)
        ListView2.Items.Add(backItem)

        ' 第二行：年份標題
        Dim titleItem As New ListViewItem($"── {selectedYear} 年月份分佈 ──")
        titleItem.SubItems.Add($"共 {monthCounts.Values.Sum:###,###,##0} 封")
        titleItem.ForeColor = Color.DimGray
        titleItem.Font = New Font(ListView2.Font, System.Drawing.FontStyle.Bold)
        ListView2.Items.Add(titleItem)

        ' 逐月顯示 (只顯示有郵件的月份) 
        For month As Integer = 1 To 12
            Dim count As Integer = 0
            monthCounts.TryGetValue(month, count)
            If count > 0 Then
                Dim monthItem As New ListViewItem($"{selectedYear} /  {month:D2}月")
                monthItem.SubItems.Add(count.ToString("###,###,##0"))
                ListView2.Items.Add(monthItem)
            End If
        Next
        ListView2.EndUpdate()

        ' 更新 Chart2 為月份長條圖
        UpdateChart2ForMonths(monthCounts, selectedYear)

        ' 左側 TreeView 選定節點保持可見, todo: 等確定SimTree2完全替代TreeView2後，可以移除TreeView2相關的程式碼
        If TreeView2.Visible AndAlso TreeView2.SelectedNode IsNot Nothing Then
            TreeView2.SelectedNode.EnsureVisible()
        ElseIf SimTree2.Visible Then
            Dim nodes As List(Of TreeNode) = SimTree2.SelectedNodes
            If nodes IsNot Nothing AndAlso nodes.Count > 0 Then nodes(0).EnsureVisible()
        End If

        ' ✅ UpdateTab2Status：顯示花費時間和速度
        Dim countedItems As Integer = monthCounts.Values.Sum
        Dim speed As Double = If(sw.Elapsed.TotalSeconds > 0, countedItems / sw.Elapsed.TotalSeconds, 0)
        lblStatus1.Text = $"共 {countedItems:###,###,##0} 封 / {selectedYear} 年"
        lblStatus2.Text = $"{selectedYear} 年月份統計花費了 {sw.Elapsed.TotalSeconds:0.00} 秒。({speed:###,##0}/sec) (按 ESC 或雙擊返回列返回年度視圖) "

        Dbg("結束：ShowMonthView: ", $"year={selectedYear}, 有郵件的月份數={monthCounts.Count}")
    End Function
    Private Sub ShowTab2Result(yearCounts As ConcurrentDictionary(Of Integer, Integer))
        ' 顯示結果的子程序
        Dbg("開始：", yearCounts.Values.Sum)

        ' 把統計完yearCounts的結果, 分別傳到ListView2和Chart2顯示
        ListView2.Items.Clear()                         ' 清空之前的統計結果
        If yearCounts Is Nothing OrElse yearCounts.Count = 0 Then
            ListView2.Items.Add(New ListViewItem("找不到郵件"))
            ' ★ 空資料夾時也要清除 Chart2，否則前一個資料夾的圖表會殘留
            Chart2.Series(0).Points.Clear()
            Dim existingAvg As Series = Chart2.Series.FindByName("平均線")
            If existingAvg IsNot Nothing Then Chart2.Series.Remove(existingAvg)
            Dim existingAnnotation = Chart2.Annotations.FindByName("avgLabel")
            If existingAnnotation IsNot Nothing Then Chart2.Annotations.Remove(existingAnnotation)
        Else
            ' 5/28修改, 二個AI都說第二段性能較好, 因為排序後轉成ToList再傳入, 才不會每次遍歷都再排序一次
            ListView2.BeginUpdate()                                                     ' ✅ 批次更新，避免每次 Add 都觸發重繪
            Dim sortedYearCounts = yearCounts.OrderBy(Function(pair) pair.Key).ToList() ' 將年份按照升序排序
            For Each pair In sortedYearCounts
                ListView2.Items.Add(New ListViewItem({pair.Key, pair.Value.ToString("###,###,##0")})) ' ✅ 改寫後: 直接從 UI 執行緒更新 ListView，不需要 Invoke
            Next
            ListView2.EndUpdate()
            UpdateChart2(sortedYearCounts)

        End If
    End Sub
    Private Sub UpdateChart2(sortedYearCounts As List(Of KeyValuePair(Of Integer, Integer)))
        Dbg("開始：")

        ' 清除之前的統計結果, 包括 Series Points 和 平均線 Series 以及平均值標籤 Annotation (避免重複加入)
        Chart2.Series(0).Points.Clear()                 ' 清除之前的 Series Points
        Dim existingAvg As Series = Chart2.Series.FindByName("平均線") ' 清除舊的平均線 Series (避免重複加入)
        If existingAvg IsNot Nothing Then Chart2.Series.Remove(existingAvg)
        Dim existingAnnotation = Chart2.Annotations.FindByName("avgLabel")  ' 先清除舊的 Annotation (避免重複加入) 
        If existingAnnotation IsNot Nothing Then Chart2.Annotations.Remove(existingAnnotation)

        ' 添加數據到 Series, 在 Chart2 中顯示統計結果
        Dim series As Series = Chart2.Series(0)
        For Each pair In sortedYearCounts
            series.Points.AddXY(pair.Key, pair.Value)
        Next

        ' 依內容大小來設置 Chart2 的 X軸上下限
        With Chart2.ChartAreas(0).AxisX
            .Minimum = sortedYearCounts.Min(Function(p) p.Key) - 0.5
            .Maximum = sortedYearCounts.Max(Function(p) p.Key) + 0.5
            .Interval = 1
            .IntervalOffset = 0                 ' ✅ 還原年度視圖的長條置中偏移
            .LabelStyle.Format = "####"         ' ✅ 還原年份格式
            .LabelStyle.Interval = 1
            .LabelStyle.IntervalOffset = 0.5    ' ✅ 校正還原上面max/min的0.5偏移
            .MajorTickMark.IntervalOffset = 0   ' ✅ 還原刻度偏移
        End With

        ' 添加一條代表平均值的線, 2026/3/6 by Claude Code  
        ' ✅ 改用獨立 Series 畫平均線，才能控制線型 (StripLine 不支援虛線)
        Dim average As Double = sortedYearCounts.Average(Function(pair) pair.Value)
        Dim xMin As Double = sortedYearCounts.Min(Function(pair) pair.Key)
        Dim xMax As Double = sortedYearCounts.Max(Function(pair) pair.Key)

        ' ✅ 新增平均線 Series，用 Line 類型才能設虛線
        Dim avgSeries As New Series("平均線") With {.ChartType = SeriesChartType.Line,
                                                    .Color = Color.Red,
                                                    .BorderWidth = 2,
                                                    .BorderDashStyle = ChartDashStyle.Dash,  ' ✅ 虛線
                                                    .ChartArea = Chart2.ChartAreas(0).Name,
                                                    .IsVisibleInLegend = False}
        avgSeries.Points.AddXY(xMin - 1, average)  ' 從 X 軸最小值開始
        avgSeries.Points.AddXY(xMax + 1, average)  ' 到 X 軸最大值結束

        ' ✅ 用 TextAnnotation 顯示平均值標籤
        Dim avgLabel As New TextAnnotation With {.Name = "avgLabel",
                                                 .Text = "AVG: " & average.ToString("#,###,##0"),
                                                 .ForeColor = Color.Red,
                                                 .Font = New Font("Tahoma", 9, System.Drawing.FontStyle.Bold),
                                                 .AnchorDataPoint = avgSeries.Points(1),  ' 標籤錨定在平均線右端
                                                 .AnchorOffsetX = -1,   ' 往左微調，避免超出右邊界
                                                 .AnchorOffsetY = -3,   ' 往上微調，讓標籤在線的上方
                                                 .BackColor = Color.Transparent,
                                                 .LineColor = Color.Transparent}
        Chart2.Series.Add(avgSeries)
        Chart2.Annotations.Add(avgLabel)
        Chart2.Invalidate() ' 強制重新繪製圖表
        Dbg("結束：")
    End Sub
    Private Sub UpdateChart2ForMonths(monthCounts As ConcurrentDictionary(Of Integer, Integer), year As Integer)
        ' ---------------------------------------------------------------
        ' 月份長條圖 (只畫 1~12 月，X 軸標籤顯示「M月」，不畫平均線) 
        ' 完整替換 Chart2 的內容，與 UpdateChart2 平行存在
        ' ---------------------------------------------------------------
        Dbg("開始：", year)

        ' 清除之前的所有圖表內容 (同 UpdateChart2 的清除邏輯) 
        Chart2.Series(0).Points.Clear()
        Dim existingAvg As Series = Chart2.Series.FindByName("平均線")
        If existingAvg IsNot Nothing Then Chart2.Series.Remove(existingAvg)
        Dim existingAnnotation = Chart2.Annotations.FindByName("avgLabel")
        If existingAnnotation IsNot Nothing Then Chart2.Annotations.Remove(existingAnnotation)

        ' 把 1~12 月的資料全部加入 (沒有郵件的月份補 0，讓 X 軸保持完整 12 格) 
        Dim series As Series = Chart2.Series(0)
        For month As Integer = 1 To 12
            Dim count As Integer = 0
            monthCounts.TryGetValue(month, count)
            ' ✅ 用月份名稱當 X 軸標籤，比純數字 1~12 更易讀
            Dim pt As DataPoint = New DataPoint()
            pt.SetValueXY(month, count)
            pt.AxisLabel = $"{month}月"
            series.Points.Add(pt)
            pt.IsVisibleInLegend = True         ' ✅ 讓圖例顯示每個月的標籤
        Next

        ' X 軸固定顯示 1~12，不根據資料範圍自動縮放 
        ' X 軸重置所有從 InitChart 繼承的年度設定，改成月份專用設定
        With Chart2.ChartAreas(0).AxisX
            '.IsMarginVisible = True             ' ✅ 月份圖保留左右空白，讓長條不緊貼 Y 軸，更美觀
            .Minimum = 0.5
            .Maximum = 12.5
            .Interval = 1
            .IntervalOffset = 0                 ' ✅ 清除 InitChart 的 0.5 偏移量
            .LabelStyle.Format = ""             ' ✅ 清除 "####" 年份格式，讓 AxisLabel 屬性生效
            .LabelStyle.Interval = 1
            .LabelStyle.IntervalOffset = 0.5    ' ✅ 清除偏移
            .MajorTickMark.IntervalOffset = 0   ' ✅ 清除刻度偏移
        End With
        Chart2.Invalidate()

        Dbg("結束：", year)
    End Sub
    Private Sub UpdateTab2Status(yearCounts As ConcurrentDictionary(Of Integer, Integer), elapsed As TimeSpan)
        Dbg("開始：")

        ' 顯示執行時間與統計速度 (lblStatus2) ，yearCounts.Values.Sum 是最可靠的實際計數來源: 
        '   - 含子資料夾時:   Sum = 整棵樹的郵件數
        '   - 不含子資料夾時: Sum = 只有選定資料夾的郵件數
        '   兩種情況都正確，不需要再透過 sender.SelectedNode 取值 (舊版 HACK 的根源) 
        Dim countedItems As Integer = yearCounts.Values.Sum
        Dim speed As Double = If(elapsed.TotalSeconds > 0, countedItems / elapsed.TotalSeconds, 0)
        lblStatus2.Text = $"更新年度統計花費了 {elapsed.TotalSeconds:0.00} 秒。({speed:###,##0}/sec)"
    End Sub
#End Region
#Region "  └ 輔助函數"
    Private Sub InitChart(chart As Chart)
        Dbg("開始：", chart.Name)
        ' 清除原有的設定
        chart.ChartAreas.Clear()
        chart.Legends.Clear()
        chart.Series.Clear()

        ' 添加 Chart 的 ChartArea
        Dim area1 As New ChartArea("chartArea")
        chart.ChartAreas.Add(area1)

        ' 添加 Chart 的 Series
        Dim series As New Series With {
        .ChartType = SeriesChartType.Column,
        .Name = "郵件數量"}
        chart.Series.Add(series)
        'series("PixelPointWidth") = "50" ' 設置長條圖的寬度

        ' 設置 Chart 的外觀
        With chart
            ' 設置抗鋸齒和文本抗鋸齒品質
            .AntiAliasing = AntiAliasingStyles.All
            .TextAntiAliasingQuality = TextAntiAliasingQuality.High

            ' 設置 ChartArea 的背景色和邊框顏色
            With area1
                .BackColor = Color.FromArgb(245, 245, 245) ' 淡灰色背景
                .BorderColor = Color.DarkGray
                .ShadowColor = Color.Transparent ' 關閉陰影效果

                ' 設置背景格線顏色和寬度
                .AxisX.LineColor = Color.DimGray
                .AxisY.LineColor = Color.DimGray
                .AxisX.MajorGrid.LineColor = Color.LightGray ' 淡灰色
                .AxisY.MajorGrid.LineColor = Color.LightGray ' 淡灰色
                .AxisX.MajorGrid.LineWidth = 1 ' 寬度1
                .AxisY.MajorGrid.LineWidth = 1 ' 寬度1

                With .AxisX
                    .Minimum = 1900             ' 設置 X 軸的最小值
                    .Maximum = 2100             ' 設置 X 軸的最大值
                    .Interval = 1               ' 設置 X 軸的間隔為 1
                    .IntervalOffset = 0.5       ' 設置偏移量為0.5，將長條置中在兩個刻度之間
                    .LabelStyle.Format = "####" ' 設置 X 軸標籤的格式
                    .LabelStyle.Interval = 1    ' 設置 X 軸標籤的顯示間隔為 1
                    .LabelStyle.IntervalOffset = 1
                    .IsLabelAutoFit = True      ' 自動調整 X 軸標籤

                    ' 設置 X 軸的刻度線位置
                    .MajorTickMark.Enabled = True
                    '.MajorTickMark.Enabled = False ' 隱藏 X 軸的主要刻度線
                    .MajorTickMark.Interval = 1
                    .MajorTickMark.IntervalOffset = 0.5
                    .MajorTickMark.TickMarkStyle = TickMarkStyle.AcrossAxis
                End With
            End With

            ' 設置 Series 的顏色
            series.Color = Color.FromArgb(70, 130, 180) ' 深藍色
        End With

    End Sub
    Private Function Find1stYear(selectedFolder As Outlook.Folder) As Integer
        Dbg("開始：", selectedFolder.Name)

        ' =============================================================
        ' 尋找資料夾中最早的郵件年份，作為統計的起點
        ' 2026/3/10, by Claude, 重構 Find1stYear 函數
        ' 改進: 多層try/catch加強錯誤處理、確保 COM 物件正確釋放，避免 RCW 殘留問題
        ' =============================================================
        Dim mail As Outlook.MailItem = Nothing
        Dim allItems As Outlook.Items = Nothing
        Dim validItems As Outlook.Items = Nothing

        ' 改用一層一層的 Try-Catch 包裹過濾，確保物件讀取失敗或類型轉換失敗都能被捕捉到
        Try
            ' 資料夾裡可能混有 MeetingRequest / ContactItem / Note 等, 這些物件沒有 ReceivedTime
            ' 透過 COM late binding 存取會拋 COMException 或 AccessViolationException (.NET 4+ 的 corrupted state exception) ，bare Catch 接不住
            ' ✅ 先 Restrict 過濾掉 null/零值 ReceivedTime 的壞項目，再升冪排序取最舊年份
            allItems = selectedFolder.Items : If allItems Is Nothing OrElse allItems.Count = 0 Then Return 1974
            validItems = allItems.Restrict("[ReceivedTime] > '1974/01/01'") : If validItems.Count = 0 Then Return 1974
            validItems.Sort("[ReceivedTime]", OlSortOrder.olDescending)
            Dim firstItem As Object = validItems.GetFirst() : If firstItem Is Nothing Then Return 1974
            mail = TryCast(firstItem, Outlook.MailItem) : If mail Is Nothing Then Return 1974
            Dim year As Integer = mail.ReceivedTime.Year : Return If(year <= 0 OrElse year > Date.Today.Year, 1974, year)

        Catch ex As System.Exception
            Dbg("Find1stYear Error: ", selectedFolder.Name & " - " & ex.Message)
            Return 1974

        Finally ' ✅ Finally 確保不管正常結束或例外都一定釋放，包括 Return 提前返回的情況
            If mail IsNot Nothing Then Marshal.ReleaseComObject(mail)
            If validItems IsNot Nothing Then Marshal.ReleaseComObject(validItems)
            If allItems IsNot Nothing Then Marshal.ReleaseComObject(allItems)
        End Try

    End Function
    Private Function BuildFilterDateRangeTab2(year As Integer, Optional mon1 As Integer = 1, Optional mon2 As Integer = 12) As String
        If year < 1974 Then Return Nothing

        'Const DATE_FORMAT As String = "yyyy/MM/dd HH:mm:ss"
        'Dim startDate As Date = Date.ParseExact($"{year}/01/01 00:00:00", DATE_FORMAT, Nothing) ' 建立當年的起始日期和結束日期
        'Dim endDate As Date = Date.ParseExact($"{year}/12/31 23:59:59", DATE_FORMAT, Nothing)   ' 設置結束日期的時間為23:59:59

        'Return $"[ReceivedTime] >= '{startDate}' AND [ReceivedTime] <= '{endDate}'"

        ' 2026/3/11, by Claude, 重構BuildFilterDateRangeTab2 函數: 增加了月份參數，並且直接用 Date 物件來建立日期範圍，避免字串格式問題
        Dim startDate As New Date(year, mon1, 1, 0, 0, 0)                                   ' ✅ 用 mon1/mon2 決定起訖月份，預設 1~12 代表整年
        Dim endDate As New Date(year, mon2, Date.DaysInMonth(year, mon2), 23, 59, 59)       ' mon2 的結束日用該月最後一天，避免硬寫 31 日造成 2 月等短月份抓不準

        Return $"[ReceivedTime] >= '{startDate}' AND [ReceivedTime] <= '{endDate}'"

    End Function
    Private Function MergeDictionaries(dict1 As ConcurrentDictionary(Of Integer, Integer), dict2 As ConcurrentDictionary(Of Integer, Integer)) As ConcurrentDictionary(Of Integer, Integer)

        If dict1 Is Nothing Then dict1 = New ConcurrentDictionary(Of Integer, Integer)
        If dict2 Is Nothing Then Return dict1

        '' 逐一遍歷合併 dict2 的鍵值對到 dict1 中，如果 dict1 已經有相同的鍵，則將值相加
        'For Each kvp As KeyValuePair(Of Integer, Integer) In dict2
        '    If dict1.ContainsKey(kvp.Key) Then: dict1(kvp.Key) += kvp.Value
        '    Else:                               dict1.Add(kvp.Key, kvp.Value)
        '    End If
        'Next

        ' ✅ LINQ 改寫後, 效能更好，因為不需要每次都檢查 dict1 是否包含鍵，直接使用 GetValueOrDefault 來獲取值，如果鍵不存在則返回 0，然後加上 dict2 的值
        For Each kvp In dict2
            dict1(kvp.Key) = dict1.GetValueOrDefault(kvp.Key, 0) + kvp.Value
        Next
        Return dict1
    End Function
#End Region
#End Region

#Region "■ 06 Tab3：依附件條件搜尋"
    ' ===================================================
    ' TabPage3 搜尋附件 — 重新設計 v2 by Claude, 2026/3/7
    ' 策略: Phase1 GetTable (快速掃描中繼資料)
    '       Phase2 GetItemFromID (僅在需要附件細節時)
    ' 優點: 大幅減少對 MailItem 物件的依賴和操作，提升搜尋效率和穩定性
    ' 可以用來替代原本的 Button3_Click 事件處理器，並且在 UI 上保持相同的使用體驗
    ' ===================================================

    '## 架構說明與各步驟分析 Button3_Click (主控流程)
    '├── BuildFilterAttachmentTab3()    → 純字串建構，無 COM
    '├── GetSubFolderList()             → COM，UI 執行緒，BFS 資料夾遍歷
    '├── FilterFolderWithAttachment()   → COM，UI 執行緒，Phase 1 核心
    '├── FilterAttachmentByName()       → COM，UI 執行緒 + Yield，Phase 2
    '├── BuildListViewItem_Tab3()       → 純 .NET，無 COM
    '└── ShowTab3Result()               → UI，BeginUpdate/AddRange/EndUpdate
#Region "  ├ L1 UI事件層"
    Private Async Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Dbg("開始：", sender.Name)
        ' ── 驗證選取的資料夾 ──
        If TreeView3.SelectedNode Is Nothing OrElse
            TryCast(TreeView3.SelectedNode.Tag, Folder) Is Nothing Then
            MessageBox.Show("請先在左側選擇目標資料夾。", "提示") : Return
        End If
        Dim rootFolder = DirectCast(TreeView3.SelectedNode.Tag, Folder)

        ' ── 鎖定 UI ──
        ListView3.Items.Clear()
        lblStatus1.Text = "準備中..." : lblStatus2.Text = ""
        Button3.Enabled = False : Button3_Stop.Visible = True
        _isTab3_Stop = False : _cancelRequested = False : TextBox3.Enabled = False
        Cursor = Cursors.WaitCursor

        Dim sw As New Stopwatch : sw.Start()

        Try
            ' ── Step 1: 驗證大小設定 (矛盾就提早返回，快取查詢在 Step3 做 LINQ 過濾) ──
            If CheckSize.Checked Then
                Dim minSize = CLng(NumberMin.Value) * GetSizeMultiplier(UnitMin.SelectedItem.ToString)
                Dim maxSize = CLng(NumberMax.Value) * GetSizeMultiplier(UnitMax.SelectedItem.ToString)
                If minSize > maxSize Then
                    MessageBox.Show("大小設定錯誤: 最小值不能大於最大值。", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If
            End If

            ' ── Step 2: 收集目標資料夾清單 ──
            Dim folderList = GetSubFolderList(rootFolder, CheckSubFolder3.Checked)
            lblStatus1.Text = $"準備掃描 {folderList.Count} 個資料夾..."
            Await Task.Yield

            ' ── Step 3: Phase 1 — GetTable 快速掃描 (含快取) ──
            ' 設計：快取存「hasattachment 全集，無大小篩選」；大小篩選在此處用 LINQ 做
            ' 好處：換大小條件不重跑 GetTable，直接從快取 LINQ，速度接近瞬間
            ' 失效：folder.Items.Count 改變時重掃 (偵測到有新信進來或刪信) 
            ' 2026-03-16 B1
            Dim targetMailList As New List(Of MailItemInfo)
            For Each folder In folderList
                If _isTab3_Stop Then Return
                lblStatus1.Text = $"Phase 1 掃描: {folder.Name}  (已找 {targetMailList.Count} 封)"
                If targetMailList.Count Mod 10 = 0 Then Await Task.Delay(0)     ' 每存入10封郵件就讓出一次控制權
                Dim folderResult = Await CheckTab3CacheOrRescan(folder)
                targetMailList.AddRange(folderResult)
            Next
            If _isTab3_Stop Then Return

            ' ── Step 3b: 大小篩選 (LINQ 記憶體過濾，不重打 GetTable) ──
            If CheckSize.Checked Then
                Dim minSz = CLng(NumberMin.Value) * GetSizeMultiplier(UnitMin.SelectedItem.ToString)
                Dim maxSz = CLng(NumberMax.Value) * GetSizeMultiplier(UnitMax.SelectedItem.ToString)
                targetMailList = targetMailList.Where(Function(c) c.Size >= minSz AndAlso c.Size <= maxSz).ToList
            End If

            ' ── Step 4: 決定是否需要 Phase 2 附件細查 ──
            Dim hasKeyword = CheckAttachName.Checked AndAlso TextBox3.Text.Trim.Length > 0
            Dim finalItems As List(Of ListViewItem)
            If hasKeyword OrElse CheckAttCount.Checked Then
                finalItems = Await ScanAttachmentByName(targetMailList)
            Else
                finalItems = BuildListViewItem_Tab3(targetMailList)
            End If

            ' ── Step 5: 顯示結果 ──
            sw.Stop()
            ShowTab3Result(finalItems, sw.Elapsed.TotalSeconds, targetMailList.Count)

        Catch ex As System.Exception
            MessageBox.Show("搜尋發生錯誤: " & ex.Message, "錯誤")
            Dbg("Button3_Click Error: ", ex.Message & vbCrLf & ex.StackTrace)
        Finally
            ' ── 無論如何都解鎖 UI ──
            TextBox3.Enabled = CheckAttachName.Checked
            Button3.Enabled = True : Button3_Stop.Visible = False
            Cursor = Cursors.Default
        End Try

    End Sub
    Private Sub Button3_Stop_Click(sender As Object, e As EventArgs) Handles Button3_Stop.Click
        Dbg("開始：", sender.Name)
        _isTab3_Stop = True : Button3_Stop.Visible = False
        lblStatus1.Text = "使用者已停止搜尋。"
    End Sub
    Private Sub ListView3_GotFocus(sender As Object, e As EventArgs) Handles ListView3.GotFocus
        If ListView3.SelectedItems.Count = 0 AndAlso ListView3.Items.Count > 0 Then
            ListView3.Items(0).Selected = True
        End If
    End Sub
    Private Sub ListView3_ColumnClick(sender As Object, e As ColumnClickEventArgs) Handles ListView3.ColumnClick
        Dbg("Begin Sorting: ", sender.Name)
        Dim sw As New Stopwatch : sw.Start()

        ' 判斷是否點選的是同一個列標題, 如果是，則切換排序方式, 否則預設使用升序排序
        currentSortOrder = If(e.Column = previousColumnIndex AndAlso currentSortOrder = SortOrder.Ascending, SortOrder.Descending, SortOrder.Ascending)
        previousColumnIndex = e.Column  ' 儲存目前點選的列索引

        ListView3.BeginUpdate()
        ListView3.ListViewItemSorter = New ListViewItemComparer(e.Column, currentSortOrder)
        ListView3.EndUpdate()

        sw.Stop()
        lblStatus2.Text = $"ListView 排序 {ListView3.Items.Count} 項，耗時 {sw.Elapsed.TotalSeconds:0.00} 秒"

        Dbg("End Sorting: ", sender.Name)
    End Sub
    Private Sub ListView3_MouseClick(sender As Object, e As MouseEventArgs) Handles ListView3.MouseClick
        ' 單擊左鍵 → 複製郵件主旨到剪貼簿 (方便貼到搜尋欄或筆記) 
        ' 2026-03-16 確認：原有行為保留
        Dbg("開始：", sender.Name)
        Dim item As ListViewItem = sender.GetItemAt(e.X, e.Y)
        If item IsNot Nothing AndAlso e.Button = MouseButtons.Left Then Clipboard.SetText(item.SubItems(0).Text)
    End Sub
    Private Sub ListView3_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles ListView3.MouseDoubleClick

        Dbg("開始：", sender.Name)
        Dim lvItem As ListViewItem = sender.GetItemAt(e.X, e.Y)
        If lvItem Is Nothing OrElse e.Button <> MouseButtons.Left Then Return

        ' 雙擊左鍵 → 複製主旨 + 用 EntryID 在 Outlook 中打開郵件
        ' 2026-03-16 確認：原有行為保留，移除舊版死碼
        Clipboard.SetText(lvItem.SubItems(0).Text)  ' 先複製主旨
        OpenMailByEntryID(lvItem.SubItems(5).Text)  ' 用 EntryID 打開郵件 (第 6 欄 SubItems(5)) 

    End Sub
    Private Sub CheckAttachName_CheckedChanged(sender As Object, e As EventArgs) Handles CheckAttachName.CheckedChanged
        TextBox3.Enabled = CheckAttachName.Checked
        If CheckAttachName.Checked Then TextBox3.Focus() : TextBox3.SelectAll()
    End Sub
    Private Sub CheckAttCount_CheckedChanged(sender As Object, e As EventArgs) Handles CheckAttCount.CheckedChanged
        Dbg("開始：", sender.Name)
        CountMin.Enabled = CheckAttCount.Checked
        CountMax.Enabled = CheckAttCount.Checked
    End Sub
    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        Dbg("開始：", sender.Name)
        If e.KeyCode = Keys.Enter Then
            Button3.PerformClick()
            TextBox3.SelectAll()
            e.SuppressKeyPress = True
        End If
    End Sub
    Private Sub NumberMin_ValueChanged(sender As Object, e As EventArgs) Handles NumberMin.ValueChanged
        NumberMin.Increment = If(NumberMin.Value < 100, 1, 10)
    End Sub
    Private Sub NumberMax_ValueChanged(sender As Object, e As EventArgs) Handles NumberMax.ValueChanged
        NumberMax.Increment = If(NumberMax.Value < 100, 1, 10)
    End Sub
#End Region
#Region "  ├ L2 流程協調層"
    Private Async Function CheckTab3CacheOrRescan(targetFolder As Outlook.Folder) As Task(Of List(Of MailItemInfo))
        ' ── Tab3 Phase1 快取查詢入口 ─────────────────────────────────────
        ' 呼叫端：Button3_Click Step3
        ' 邏輯：
        '   1. 讀取 folder.Items.Count 做失效判斷
        '   2. 快取命中且 ItemCount 未變 → 直接回傳快取，零 GetTable
        '   3. 快取失效 → 呼叫 FilterFolderWithAttachment (無大小篩選) → 存入快取
        ' 2026-03-16 B1 新增
        Dbg("開始：", targetFolder.Name)

        Dim key As String = targetFolder.FolderPath
        Dim currentCount As Integer = GetMailCount(targetFolder)    ' 只做單次 COM，代價極低

        Dim entry As FolderCacheTab3
        If _tab3Phase1Cache.TryGetValue(key, entry) AndAlso entry.ItemCountWhenCached = currentCount Then
            Dbg("Tab3 Cache Hit: ", targetFolder.Name & $" ({currentCount} items)")
            Return entry.mailWithAttachment
        End If
        Dbg("Tab3 Cache Miss: ", targetFolder.Name)   ' 快取未命中或已失效：重新掃描 (使用無大小篩選的基礎 filter) 

        ' 開始逐一掃瞄所有資料夾
        Dim targetMailList As List(Of MailItemInfo) = Await ScanFolderWithAttachment(targetFolder)
        _tab3Phase1Cache(key) = New FolderCacheTab3 With {.mailWithAttachment = targetMailList, .ItemCountWhenCached = currentCount}

        Return targetMailList
    End Function
    Private Async Function ScanFolderWithAttachment(folder As Outlook.Folder) As Task(Of List(Of MailItemInfo))

        ' Phase 1: GetTable + GetArray 批次掃描單一資料夾
        ' 2026/3/24 by AntiGravity: 從 GetNextRow 逐行讀取改為 GetArray(1000) 批次讀取
        ' 說明: GetArray 一次把最多 N 筆 row 以 Object(,) 二維陣列傳回，大幅減少 COM 跨程序呼叫次數
        '       原本每封信一次 COM call (GetNextRow)，現在每 1000 封只需一次 COM call (GetArray)
        ' GetTable 同時套用 DASL 篩選，MAPI 層就已過濾，不用逐一判斷
        Const PR_MESSAGE_SIZE As String = "http://schemas.microsoft.com/mapi/proptag/0x0E080003"   ' MAPI 屬性的 DASL URI
        Const BATCH_SIZE As Integer = 1000  ' 2026/3/24 by AntiGravity: 每次批量讀取的筆數
        Dim table As Outlook.Table = Nothing        ' ✅ 宣告在 Try 外，初始化在 Try 內，才能在 Finally 正確釋放
        Dbg("開始：", folder.Name)

        ' 比對用的 hasattachment 基礎 DASL (不含大小條件) 
        Dim strFilterHasAttachment As String = "@SQL=" & Chr(34) & "urn:schemas:httpmail:hasattachment" & Chr(34) & " = True"
        Dim result As New List(Of MailItemInfo)
        Try                                     ' ✅ GetTable() 移進 Try，拋例外時才能被 Catch 接住
            table = folder.GetTable(strFilterHasAttachment)
            table.Columns.RemoveAll()           ' 清除預設欄位，只保留需要的，最小化資料傳輸量
            table.Columns.Add("EntryID")        ' 欄位索引 0，稍後 GetItemFromID 用
            table.Columns.Add("Subject")        ' 欄位索引 1
            table.Columns.Add(PR_MESSAGE_SIZE)  ' 欄位索引 2
            table.Columns.Add("ReceivedTime")   ' 欄位索引 3
            table.Columns.Add("SenderName")     ' 欄位索引 4

            ' 2026/3/24 by AntiGravity: 改用 GetArray 批次讀取，減少 COM 跨程序呼叫
            ' todo: 這裡用GetArray() 改寫後, 速度也快太多了吧???!!!
            Dim rowCount As Integer = 0
            Do While Not table.EndOfTable
                If _isTab3_Stop Then Exit Do
                Dim arr As Object = table.GetArray(BATCH_SIZE)  ' 一次讀取最多 BATCH_SIZE 筆，回傳 Object(,) 二維陣列
                If arr Is Nothing Then Exit Do
                Dim data(,) As Object = DirectCast(arr, Object(,))
                Dim rows As Integer = data.GetUpperBound(0) + 1  ' 實際讀回的筆數 (可能 < BATCH_SIZE)

                For r As Integer = 0 To rows - 1
                    Try
                        Dim entryID As String = If(data(r, 0) Is Nothing OrElse IsDBNull(data(r, 0)), "", data(r, 0).ToString())
                        If entryID = "" Then Continue For   ' ✅ EntryID 是空的就跳過，這筆資料沒有用

                        Dim info As New MailItemInfo With {
                            .EntryID = entryID,
                            .Subject = If(data(r, 1) Is Nothing OrElse IsDBNull(data(r, 1)), "", data(r, 1).ToString()),
                            .Size = If(data(r, 2) Is Nothing OrElse IsDBNull(data(r, 2)), 0L, CLng(data(r, 2))),
                            .ReceivedTime = If(data(r, 3) Is Nothing OrElse IsDBNull(data(r, 3)), DateTime.MinValue, CDate(data(r, 3))),
                            .SenderName = If(data(r, 4) Is Nothing OrElse IsDBNull(data(r, 4)), "", data(r, 4).ToString())}
                        result.Add(info)
                    Catch ex As System.Exception
                        Dbg("GetArray ROW error: " & rowCount, ex.Message)
                    End Try
                    rowCount += 1
                Next

                ' 2026/3/24 by AntiGravity: 每批次讀完讓出一次 UI 執行緒，保持 ESC 回應
                Await Task.Delay(0)
                If _isTab3_Stop Then Exit Do
            Loop

        Catch ex As System.Exception
            Dbg("FilterFolderWithAttachment Error: ", folder.Name & " — " & ex.Message)
        Finally
            If table IsNot Nothing Then Marshal.ReleaseComObject(table)
        End Try
        Return result
    End Function
    Private Async Function ScanAttachmentByName(targetMailList As List(Of MailItemInfo)) As Task(Of List(Of ListViewItem))

        ' Phase 2: 逐一載入 MailItem，檢查附件名稱/數量
        ' 說明: 只在有 keyword 或 count filter 時才執行
        '       COM STA 安全: 所有 GetItemFromID 都在 UI 執行緒
        '       Await Task.Yield() 每 10 封讓 UI 更新一次
        ' todo: phase 2 加入快取 (同資料夾比對不同檔名時直接從快取)
        ' todo: phase 2 很多COM exception
        ' todo: phase 2 無法ESC 中斷
        Dbg("開始：")

        Dim mustCountAttach As Boolean = CheckAttCount.Checked
        Dim minCount As Integer = If(mustCountAttach, CInt(CountMin.Value), 0)
        Dim maxCount As Integer = If(mustCountAttach, CInt(CountMax.Value), Integer.MaxValue)


        Dim resultItems As New List(Of ListViewItem)
        Dim total As Integer = targetMailList.Count
        Dim processed As Integer = 0

        ' ← 新增這兩行，讓進度從 0 開始顯示，格式與後續一致
        lblStatus1.Text = $"Phase 2: 0 / {total}，已符合 0 封"
        If processed Mod 10 = 0 Then Await Task.Delay(1)     ' 每掃瞄20個郵件就讓出一次控制權

        Dim keyword As String = If(CheckAttachName.Checked, TextBox3.Text.Trim.ToLower(), "")
        For Each mail As MailItemInfo In targetMailList
            If _isTab3_Stop Then Exit For

            Dim tempMail As MailItem = Nothing
            Dim attachments As Outlook.Attachments = Nothing
            Try
                ' GetItemFromID 必須在 UI 執行緒 (COM STA)
                tempMail = TryCast(_olNS.GetItemFromID(mail.EntryID), MailItem)
                If tempMail Is Nothing Then Continue For

                ' ── 數量篩選 ──
                attachments = tempMail.Attachments
                Dim attachCount As Integer = attachments.Count
                If mustCountAttach AndAlso (attachCount < minCount OrElse attachCount > maxCount) Then Continue For

                ' ── 檔名關鍵字篩選 ──
                If keyword.Length > 0 Then
                    Dim found As Boolean = False
                    For Each att As Outlook.Attachment In attachments
                        If att.FileName.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0 Then
                            found = True : Exit For
                        End If
                    Next
                    If Not found Then Continue For
                End If

                ' 通過所有篩選的項目就加進清單準備傳回
                ' todo: 再次嚐試使用MAPI屬性: PR_ATTACH_SIZE, PR_ATTACH_NUM, PR_ATTACH_FILENAME / PR_ATTACH_LONG_FILENAME, PR_NORMALIZED_SUBJECT
                resultItems.Add(New ListViewItem({mail.Subject,
                                                  mail.Size.ToString("###,###,##0"),
                                                  mail.ReceivedTime.ToShortDateString(),
                                                  mail.SenderName,
                                                  attachCount.ToString(),
                                                  mail.EntryID}))
            Catch ex As System.Exception
                Dbg("FilterByAttachDetail Error: ", ex.Message)
            Finally
                If attachments IsNot Nothing Then Marshal.ReleaseComObject(attachments)
                If tempMail IsNot Nothing Then Marshal.ReleaseComObject(tempMail)
            End Try

            processed += 1
            lblStatus1.Text = $"Phase 2: {processed} / {total}，已符合 {resultItems.Count} 封"

            If processed Mod 10 = 0 Then            ' ✅ 每 20 封讓出一次，用 Delay(0) 確保 message pump 有機會處理 ESC/STOP
                Await Task.Delay(1)                 '    Task.Yield() 只是把 continuation 貼回 queue 尾端，queue 滿時 ESC 排不進來
                If _isTab3_Stop Then Exit For    '    Task.Delay(0) 用 Timer 實作，這 1ms 內 message pump 優先處理鍵盤/滑鼠事件
            End If
        Next

        Return resultItems
    End Function
    Private Function BuildListViewItem_Tab3(targetMailList As List(Of MailItemInfo)) As List(Of ListViewItem)

        ' 從 Phase 1 候選資料建立 ListViewItem (不需附件細節時)
        ' AttachmentCount 欄顯示 ">0" (已知有附件但未精確計數)
        Dbg("開始：")
        Dim items As New List(Of ListViewItem)(targetMailList.Count)
        For Each c As MailItemInfo In targetMailList
            items.Add(New ListViewItem({c.Subject,
                                        c.Size.ToString("###,###,##0"),
                                        c.ReceivedTime.ToShortDateString(),
                                        c.SenderName,
                                        ">0",           ' 有附件但未計數，避免載入全部 MailItem
                                        c.EntryID}))
        Next
        Return items
    End Function
    Private Sub ShowTab3Result(items As List(Of ListViewItem), elapsedSeconds As Double, totalProcessed As Integer)
        Dbg("開始：", items.Count)

        ListView3.Items.Clear()
        Dim lvCount As Integer = items.Count
        ' 先告訴 ListView 總共會有幾筆，讓它一次配置好記憶體，不要每次 Add 都 realloc
        If lvCount > 50 Then SendMessage(ListView3.Handle, LVM_SETITEMCOUNT, New IntPtr(lvCount), IntPtr.Zero)
        If lvCount > 10 Then ListView3.BeginUpdate()
        If lvCount > 0 Then ListView3.Items.AddRange(items.ToArray()) Else ListView3.Items.Add("找不到符合條件的郵件")
        ListView3.EndUpdate()
        lblStatus1.Text = $"共找到 {lvCount} 封郵件"

        ' totalProcessed 避免除以零
        Dim speedText As String = ""
        If elapsedSeconds > 0 AndAlso totalProcessed > 0 Then speedText = $" ({CInt(totalProcessed / elapsedSeconds):###,##0}/sec)"
        lblStatus2.Text = $"耗時 {elapsedSeconds:0.00} 秒{speedText}"

        Dbg("ShowTab3Result: ", $"{items.Count} 封，{elapsedSeconds:0.00}s")
    End Sub
#End Region
#Region "  └ 輔助函數"
    Private Function BuildFilterAttachmentTab3() As String
        ' 2026-03-16 B1：大小篩選移到 Button3_Click Step3b 的 LINQ，
        '               此函數保留但現在只回傳 hasattachment 基礎 filter (與 strFilterHasAttachment 一致) 
        '               保留原有大小條件建構邏輯以備日後參考，但 Button3_Click 已不呼叫此函數
        Dim q As String = Chr(34)
        Return "@SQL=" & q & "urn:schemas:httpmail:hasattachment" & q & " = True"
    End Function
    Private Function GetSizeMultiplier(sizeUnit As String, Optional base1024 As Boolean = False) As Integer
        Dbg("開始：")
        ' 獲取大小單位的倍數
        Dim multi As Long = If(base1024, 1024, 1000)
        Select Case sizeUnit.ToLower()
            Case "kb" : Return multi
            Case "mb" : Return multi ^ 2
            Case "gb" : Return multi ^ 3
            Case Else : Return 1
        End Select
    End Function
    Private Sub OpenMailByEntryID(strEntryID As String)
        Dbg("OpenMailByEntryID: " & strEntryID)
        ' 依照傳入的Mailitem's EntryID, 呼叫NameSpace打開郵件再釋放object
        If strEntryID Is Nothing Then Return

        'Dim mail As MailItem = Nothing
        'Try
        '    mail = CType(objNameSpace.GetItemFromID(strEntryID), MailItem)
        '    mail.Display()
        'Catch ex As System.Exception
        '    MessageBox.Show("無法開啟郵件: " & ex.Message)
        'Finally
        '    If mail IsNot Nothing Then Marshal.ReleaseComObject(mail)   ' ✅ MailItem 釋放
        'End Try

        ' 2026/3/20, by Claude.ai, 建立獨立執行緒fire-and-forget
        ' 讓作業系統跟outlook.exe 自己去做它們的事, 我們不用等它開啟完畢, 可以直接回到自己的程式介面
        Dim ns As Outlook.NameSpace = Nothing
        Dim mail As MailItem = Nothing
        Dim th As New Thread(Sub()
                                 Try
                                     ns = _olApp.GetNamespace("MAPI")
                                     mail = CType(ns.GetItemFromID(strEntryID), MailItem)
                                     mail.Display()
                                 Catch ex As System.Exception
                                     ' 開視窗失敗，靜默忽略（或 BeginInvoke 到 UI 執行緒顯示 MessageBox）
                                     'MessageBox.Show("無法開啟郵件: " & ex.Message)
                                 Finally
                                     If mail IsNot Nothing Then Marshal.ReleaseComObject(mail)
                                     If ns IsNot Nothing Then Marshal.ReleaseComObject(ns)       ' ✅ 補上
                                 End Try
                             End Sub)
        th.SetApartmentState(ApartmentState.STA)    ' ✅ 新執行緒設 STA，COM 呼叫合法
        th.IsBackground = True                      ' ✅ 主程式關閉時不等這條執行緒
        th.Start()                                  ' ✅ fire-and-forget，直接 return，主程式UI 立刻恢復回應

    End Sub
#End Region
#End Region

#Region "■ 07 Tab4：系列郵件"
    Private Async Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dbg("開始：", sender.Name)

        Dim rootFolder As Outlook.Folder = TryCast(TreeView1.SelectedNode?.Tag, Outlook.Folder)
        If rootFolder Is Nothing Then
            MessageBox.Show("請先在左側 Tab1 選擇要掃描的資料夾", "提示")
            Return
        End If

        Button4.Enabled = False
        Cursor = Cursors.WaitCursor
        TreeView4.Nodes.Clear()
        ListView4.Items.Clear()
        lblStatus1.Text = "開始掃描系列郵件..."

        Dim sw As New Stopwatch() : sw.Start()
        Dim topicDict As New Dictionary(Of String, List(Of MailItemInfo))(StringComparer.OrdinalIgnoreCase)

        Try
            ' 取得所有子資料夾
            Dim targetFolderList As List(Of Outlook.Folder) = GetSubFolderList(rootFolder, includeSubFolders:=True)
            Dim processed As Integer = 0

            For Each folder In targetFolderList
                Dim table As Outlook.Table = Nothing
                Try
                    table = folder.GetTable()
                    Const PR_MESSAGE_SIZE As String = "http://schemas.microsoft.com/mapi/proptag/0x0E080003"
                    Const PR_CONVERSATION_TOPIC As String = "http://schemas.microsoft.com/mapi/proptag/0x0070001E"

                    table.Columns.RemoveAll()
                    table.Columns.Add("EntryID")
                    table.Columns.Add("Subject")
                    table.Columns.Add(PR_MESSAGE_SIZE)
                    table.Columns.Add("ReceivedTime")
                    table.Columns.Add("SenderName")
                    table.Columns.Add(PR_CONVERSATION_TOPIC)

                    Do While Not table.EndOfTable
                        Dim arr As Object = table.GetArray(1000)
                        If arr Is Nothing Then Exit Do
                        Dim data(,) As Object = DirectCast(arr, Object(,))

                        For r As Integer = 0 To data.GetUpperBound(0)
                            Dim topic As String = If(data(r, 5) Is Nothing OrElse IsDBNull(data(r, 5)), "", data(r, 5).ToString())
                            If topic = "" Then Continue For ' 沒有 Conversation Topic 的信件略過

                            Dim entryID As String = If(data(r, 0) Is Nothing OrElse IsDBNull(data(r, 0)), "", data(r, 0).ToString())
                            If entryID = "" Then Continue For

                            Dim info As New MailItemInfo With {
                                .EntryID = entryID,
                                .Subject = If(data(r, 1) Is Nothing OrElse IsDBNull(data(r, 1)), "", data(r, 1).ToString()),
                                .Size = If(data(r, 2) Is Nothing OrElse IsDBNull(data(r, 2)), 0L, CLng(data(r, 2))),
                                .ReceivedTime = If(data(r, 3) Is Nothing OrElse IsDBNull(data(r, 3)), DateTime.MinValue, CDate(data(r, 3))),
                                .SenderName = If(data(r, 4) Is Nothing OrElse IsDBNull(data(r, 4)), "", data(r, 4).ToString())
                            }

                            If Not topicDict.ContainsKey(topic) Then
                                topicDict(topic) = New List(Of MailItemInfo)()
                            End If
                            topicDict(topic).Add(info)
                        Next
                    Loop
                Catch ex As System.Exception
                    Dbg("Button4 GetTable Error: " & folder.Name, ex.Message)
                Finally
                    If table IsNot Nothing Then Marshal.ReleaseComObject(table)
                End Try

                processed += 1
                lblStatus1.Text = $"掃描中: {processed} / {targetFolderList.Count}"
                If processed Mod 5 = 0 Then Await Task.Yield()
            Next

            ' 將結果加入 TreeView4 (只加數量 > 1 的)
            TreeView4.BeginUpdate()
            For Each kvp In topicDict
                If kvp.Value.Count > 1 Then
                    Dim node As New TreeNode($"{kvp.Key} ({kvp.Value.Count})")
                    node.Tag = kvp.Value ' 存入 List(Of MailItemInfo)
                    TreeView4.Nodes.Add(node)
                End If
            Next
            TreeView4.EndUpdate()

            sw.Stop()
            lblStatus1.Text = $"掃描完成，找到 {TreeView4.Nodes.Count} 個系列郵件"
            lblStatus2.Text = $"耗時 {sw.Elapsed.TotalSeconds:0.00} 秒"

        Catch ex As System.Exception
            MessageBox.Show("掃描系列郵件時發生錯誤: " & ex.Message, "錯誤")
            Dbg("Button4_Click Error: ", ex.Message)
        Finally
            Button4.Enabled = True
            Cursor = Cursors.Default
        End Try
    End Sub
    Private Sub TreeView4_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TreeView4.AfterSelect
        Dbg("開始：", e.Node.Text)
        Dim mailList As List(Of MailItemInfo) = TryCast(e.Node.Tag, List(Of MailItemInfo))
        If mailList Is Nothing Then Return

        ' 排序：依據時間遞減 (越新的在越前面)
        mailList.Sort(Function(a, b) b.ReceivedTime.CompareTo(a.ReceivedTime))

        ListView4.BeginUpdate()
        ListView4.Items.Clear()
        For Each mailItem In mailList
            Dim lvi As New ListViewItem({
                mailItem.Subject,
                (mailItem.Size \ 1024L).ToString("###,###,###,##0") & "KB",
                mailItem.ReceivedTime.ToString("yyyy/MM/dd HH:mm:ss"),
                mailItem.SenderName,
                mailItem.EntryID
            })
            ListView4.Items.Add(lvi)
        Next
        ListView4.EndUpdate()
    End Sub
    Private Sub ListView4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListView4.SelectedIndexChanged
        Dbg("開始：", sender.Name)
        ' todo: double-click就直接把subject name 送到outlook 搜尋欄位

    End Sub
    Private Sub ListView4_MouseClick(sender As Object, e As MouseEventArgs) Handles ListView4.MouseClick
        ' ── ListView4 滑鼠事件 (Tab4 系列郵件搜尋結果) ──
        ' 2026-03-16 新增：單擊/雙擊都複製主旨到剪貼簿
        Dim item As ListViewItem = sender.GetItemAt(e.X, e.Y)
        If item IsNot Nothing AndAlso e.Button = MouseButtons.Left Then Clipboard.SetText(item.SubItems(0).Text)
    End Sub
    Private Sub ListView4_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles ListView4.MouseDoubleClick
        Dim item As ListViewItem = sender.GetItemAt(e.X, e.Y)
        If item IsNot Nothing AndAlso e.Button = MouseButtons.Left Then Clipboard.SetText(item.SubItems(0).Text)
    End Sub
#End Region

#Region "■ 08 Tab5：重複郵件"
    Private Async Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dbg("開始：", sender.Name)

        If _pstStoreList Is Nothing OrElse _pstStoreList.Count = 0 Then
            MessageBox.Show("PST 檔案庫尚未載入完成，請稍後再試", "提示")
            Return
        End If

        Button5.Enabled = False
        Cursor = Cursors.WaitCursor
        ListView5.BeginUpdate()
        ListView5.Items.Clear()
        ListView5.EndUpdate()
        lblStatus1.Text = "開始全信箱掃描重複郵件..."

        Dim sw As New Stopwatch() : sw.Start()

        ' Key = 雜湊值, Value = 該雜湊對應的郵件清單
        Dim exactDict As New Dictionary(Of String, List(Of MailItemInfo))(StringComparer.OrdinalIgnoreCase)
        Dim isExact As Boolean = rbExactMatch.Checked

        Try
            ' 遍歷所有 Store
            Dim totalProcessed As Integer = 0
            For Each store In _pstStoreList
                If _cancelRequested Then Exit For

                Try
                    Dim rootFolder As Outlook.Folder = store.GetRootFolder()
                    Dim targetFolderList As List(Of Outlook.Folder) = GetSubFolderList(rootFolder, includeSubFolders:=True)

                    For Each folder In targetFolderList
                        If _cancelRequested Then Exit For

                        Dim table As Outlook.Table = Nothing
                        Try
                            table = folder.GetTable()
                            Const PR_MESSAGE_SIZE As String = "http://schemas.microsoft.com/mapi/proptag/0x0E080003"

                            table.Columns.RemoveAll()
                            table.Columns.Add("EntryID")
                            table.Columns.Add("Subject")
                            table.Columns.Add(PR_MESSAGE_SIZE)
                            table.Columns.Add("ReceivedTime")
                            table.Columns.Add("SenderName")

                            Do While Not table.EndOfTable
                                Dim arr As Object = table.GetArray(1000)
                                If arr Is Nothing Then Exit Do
                                Dim data(,) As Object = DirectCast(arr, Object(,))

                                For r As Integer = 0 To data.GetUpperBound(0)
                                    Dim entryID As String = If(data(r, 0) Is Nothing OrElse IsDBNull(data(r, 0)), "", data(r, 0).ToString())
                                    If entryID = "" Then Continue For

                                    Dim subject As String = If(data(r, 1) Is Nothing OrElse IsDBNull(data(r, 1)), "", data(r, 1).ToString())
                                    Dim size As Long = If(data(r, 2) Is Nothing OrElse IsDBNull(data(r, 2)), 0L, CLng(data(r, 2)))
                                    Dim recvTime As DateTime = If(data(r, 3) Is Nothing OrElse IsDBNull(data(r, 3)), DateTime.MinValue, CDate(data(r, 3)))
                                    Dim senderName As String = If(data(r, 4) Is Nothing OrElse IsDBNull(data(r, 4)), "", data(r, 4).ToString())

                                    Dim info As New MailItemInfo With {
                                        .EntryID = entryID,
                                        .Subject = subject,
                                        .Size = size,
                                        .ReceivedTime = recvTime,
                                        .SenderName = senderName
                                    }

                                    Dim hashKey As String
                                    If isExact Then
                                        hashKey = $"{subject}|{size}|{recvTime:yyyyMMddHHmmss}|{senderName}"
                                    Else
                                        Dim cleanSubj As String = subject.ToUpper().Replace("RE:", "").Replace("FW:", "").Replace("回覆:", "").Replace("轉寄:", "").Replace(" ", "").Trim()
                                        If cleanSubj.Length > 20 Then cleanSubj = cleanSubj.Substring(0, 20)
                                        hashKey = $"{cleanSubj}|{size}"
                                    End If

                                    If Not exactDict.ContainsKey(hashKey) Then
                                        exactDict(hashKey) = New List(Of MailItemInfo)()
                                    End If
                                    exactDict(hashKey).Add(info)
                                Next
                                Await Task.Yield()
                            Loop
                        Catch ex As System.Exception
                            Dbg("Button5 GetTable Error: " & folder.Name, ex.Message)
                        Finally
                            If table IsNot Nothing Then Marshal.ReleaseComObject(table)
                        End Try

                        totalProcessed += 1
                        lblStatus1.Text = $"掃描中 ({store.DisplayName}): 已處理 {totalProcessed} 個資料夾"
                        If totalProcessed Mod 10 = 0 Then Await Task.Yield()
                    Next
                Catch ex As System.Exception
                    Dbg("Button5 Store Error: ", ex.Message)
                End Try
            Next

            ' 尋找符合條件的群組
            ListView5.BeginUpdate()
            Dim groupID As Integer = 1
            Dim totalDuplicateMails As Integer = 0

            For Each kvp In exactDict
                If kvp.Value.Count > 1 Then
                    Dim isValidGroup As Boolean = True

                    ' 若是 Fuzzy 模式，還需確認 Levenshtein 距離不超過門檻 (至少大於 0.8 相似度)
                    If Not isExact Then
                        Dim firstSubject As String = kvp.Value(0).Subject.ToUpper()
                        For i As Integer = 1 To kvp.Value.Count - 1
                            Dim sim As Double = CalculateSimilarity(firstSubject, kvp.Value(i).Subject.ToUpper())
                            If sim < 0.8 Then
                                isValidGroup = False
                                Exit For
                            End If
                        Next
                    End If

                    If isValidGroup Then
                        Dim groupColor As Color = If(groupID Mod 2 = 0, Color.FromArgb(240, 248, 255), Color.White)
                        For Each mailItem In kvp.Value
                            Dim lvi As New ListViewItem({
                                mailItem.Subject,
                                (mailItem.Size \ 1024L).ToString("###,###,###,##0") & "KB",
                                mailItem.ReceivedTime.ToString("yyyy/MM/dd HH:mm:ss"),
                                mailItem.SenderName,
                                "群組 " & groupID.ToString(),
                                mailItem.EntryID
                            })
                            lvi.BackColor = groupColor
                            ListView5.Items.Add(lvi)
                            totalDuplicateMails += 1
                        Next
                        groupID += 1
                    End If
                End If
            Next
            ListView5.EndUpdate()

            sw.Stop()
            lblStatus1.Text = $"掃描完成，找到 {groupID - 1} 組 ({totalDuplicateMails} 封) 重複郵件"
            lblStatus2.Text = $"耗時 {sw.Elapsed.TotalSeconds:0.00} 秒"

        Catch ex As System.Exception
            MessageBox.Show("掃描重複郵件時發生錯誤: " & ex.Message, "錯誤")
            Dbg("Button5_Click Error: ", ex.Message)
        Finally
            Button5.Enabled = True
            Cursor = Cursors.Default
        End Try
    End Sub
    Private Function CalculateSimilarity(strA As String, strB As String) As Double
        Dbg("開始：")
        ' 計算編輯距離
        Dim editDistance As Integer = LevenshteinDistance(strA, strB)

        ' 將編輯距離歸一化為範圍在 0 到 1 之間的值
        Dim maxLength As Integer = Math.Max(strA.Length, strB.Length)
        Dim similarity As Double = 1 - CDbl(editDistance) / maxLength

        Return similarity
    End Function
    Private Function LevenshteinDistance(strA As String, strB As String) As Integer
        Dbg("開始：")

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
#End Region

#Region "■ 09 Tab6：Debug & 設定"
    Private Sub CheckDebug_CheckedChanged(sender As Object, e As EventArgs) Handles CheckDebug.CheckedChanged
        _isDebugMode = CheckDebug.Checked
        Dbg("開始：", sender.Name)
        Me.Left = If(CheckDebug.Checked, Me.Left - 240, Me.Left + 240)

        ' 2026/3/26 by AntiGravity: 先同步位置與大小再顯示，確保第一次 Load 時就能抓到正確的視窗寬度
        If CheckDebug.Checked Then SyncDebugFormPosition()
        DebugForm.Visible = CheckDebug.Checked
    End Sub


    Private Sub OST_Click(sender As Object, e As EventArgs)

        'Dim outlookApp As Outlook.Application = Nothing
        'Dim ns As Outlook.NameSpace = Nothing
        'Dim inbox As Outlook.Folder = Nothing

        'Try
        '    ReadEmailsFromOST("D:\Users\Simon\Documents\Outlook 檔案\Work\Inbox_2011_GLI.ost")

        'Finally
        '    If inbox IsNot Nothing Then Marshal.ReleaseComObject(inbox)
        '    If ns IsNot Nothing Then Marshal.ReleaseComObject(ns)
        '    If outlookApp IsNot Nothing Then Marshal.ReleaseComObject(outlookApp)
        'End Try

    End Sub
    Private Sub ReadEmailsFromOST(path As String)
        Dbg("開始：")
        ' 創建 Outlook 應用程序對象
        Dim outlookApp As New Outlook.Application()

        ' 獲取 Outlook 命名空間
        Dim ns As Outlook.NameSpace = outlookApp.GetNamespace("MAPI")

        ' 添加本地 OST 文件
        ns.AddStore(path)

        ' 獲取默認的收件箱文件夾
        Dim inbox As Outlook.Folder = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox)

        ' 讀取郵件
        ReadFolderEmails(inbox)

        ' 釋放 COM 對象
        Marshal.ReleaseComObject(inbox)
        Marshal.ReleaseComObject(ns)
        Marshal.ReleaseComObject(outlookApp)
    End Sub
    Private Sub ReadFolderEmails(folder As Outlook.Folder)
        ' 迭代郵件項
        For Each item As Object In folder.Items
            If TypeOf item Is Outlook.MailItem Then
                Dim mail As Outlook.MailItem = CType(item, Outlook.MailItem)
                ' 在這裡處理郵件，比如顯示主題和內容等
                MessageBox.Show($"Subject: {mail.Subject}, Received: {mail.ReceivedTime}")
                Marshal.ReleaseComObject(mail)
            End If
        Next
    End Sub
    Private Sub Debug_Click(sender As Object, e As EventArgs) Handles btnDebug.Click
        Dbg("開始：Debug Button: ")







    End Sub
#End Region

#Region "■ 10 底層 COM 函數群（新設計，現役主力）"
    ' === 從頭重新設計底層計數函數 ===
    ' 目的：提供一個純粹的 COM 資料層函數，專注於讀取資料，不做任何流程控制或快取邏輯
    '       取代目前散落在各處的 GetMailCountByMAPINew、GetFolderSizeLegacy 等函數，統一為一個簡單的 GetXxxL3 函數
    ' 架構：L3 純資料層，L2 流程協調層，L1 UI 事件層
    '       L3 只負責讀取資料夾的本層郵件數 (GetDirectMailCountL3) ，不遞迴、不展開子資料夾，最小化 COM 呼叫量
    '       上層流程 (如 ComputeFolderStatsAsync) 負責決定何時呼叫、如何使用結果、快取管理等
    ' ==============================================================
    ' === L3 底層 COM 資料層函數群 ===
    ' 設計原則：
    '   1. 每個函數只負責一件事：讀取單一資料夾或單封郵件的一種屬性
    '   2. 不做快取、不做遞迴、不做 BFS 展開——這些全部交給 L2 流程協調層
    '   3. Fallback 鏈：RDO → MAPI GetArray → OOM最後手段
    '                   parallel.foreach → BFS → Recursive，每層不論成功失敗都丟 Debug message
    '   4. 失敗統一回傳 -1 (不回 0) ，讓 L2 能區分「真的是 0」或「讀取失敗」
    '   5. 所有 COM 物件在 Finally 中釋放，確保 RCW 不殘留
    ' ==============================================================
    Private Function GetMailCount(folder As Outlook.Folder) As Integer
        ' --------------------------------------------------------------
        ' GetMailCount：只讀單一資料夾的本層郵件數 (不含子孫)
        ' Fallback 鏈：
        '   ⓪ Redemption : RDOFolder.Items.Count (可在非 STA 執行緒呼叫)
        '   ① MAPI : PR_CONTENT_COUNT (0x36020003) (最快快取屬性)
        '   ② OOM  : folder.Items.Count (會建立 Items 集合)
        '   ③ fail : Return -1
        ' --------------------------------------------------------------
        Dbg("開始：GetMailCount", folder.Name)
        Dim sw As New Stopwatch() : sw.Start()

        ' ⓪ Redemption：RDOFolder.Items.Count
        If _rdo IsNot Nothing Then
            Dim rdoFolder As Redemption.RDOFolder = Nothing
            Try
                rdoFolder = _rdo.GetFolderFromID(folder.EntryID, folder.StoreID)
                Dim count As Integer = rdoFolder.Items.Count
                sw.Stop()
                Dbg("GetMailCount ⓪ RDO 成功", $"{folder.Name} | count={count} | {sw.ElapsedMilliseconds}ms")
                Return count
            Catch ex As System.Exception
                Dbg("GetMailCount ⓪ RDO 失敗", $"{folder.Name} | {ex.Message}")
            Finally
                If rdoFolder IsNot Nothing Then Marshal.ReleaseComObject(rdoFolder)
            End Try
        End If

        ' ① MAPI：PR_CONTENT_COUNT (0x36020003)
        Try
            Const PR_CONTENT_COUNT As String = "http://schemas.microsoft.com/mapi/proptag/0x36020003"
            Dim count As Integer = CInt(folder.PropertyAccessor.GetProperty(PR_CONTENT_COUNT))
            sw.Stop()
            Dbg("GetMailCount ① MAPI 成功", $"{folder.Name} | count={count} | {sw.ElapsedMilliseconds}ms")
            Return count
        Catch ex As System.Exception
            Dbg("GetMailCount ① MAPI 失敗", $"{folder.Name} | {ex.Message}")
        End Try

        ' ② OOM：folder.Items.Count
        Try
            Dim items As Outlook.Items = Nothing
            Try
                items = folder.Items
                Dim count As Integer = items.Count
                sw.Stop()
                Dbg("GetMailCount ② OOM 成功", $"{folder.Name} | count={count} | {sw.ElapsedMilliseconds}ms")
                Return count
            Finally
                If items IsNot Nothing Then Marshal.ReleaseComObject(items)
            End Try
        Catch ex As System.Exception
            Dbg("GetMailCount ② OOM 也失敗", $"{folder.Name} | {ex.Message}")
        End Try

        sw.Stop()
        Dbg("結束：GetMailCount (FAIL)", $"{folder.Name} | -1 | {sw.ElapsedMilliseconds}ms")
        Return -1
    End Function
    Private Async Function GetMailCountAll(rootFolder As Outlook.Folder, Optional onProgress As Action(Of Integer, Integer) = Nothing) As Task(Of Long)
        ' --------------------------------------------------------------
        ' GetMailCountAll v3.0：讀取某資料夾及其整棵子樹的郵件總數
        ' todo: 改成平行處理跟GetArray() 的v4.0
        '
        ' v3.0 變更說明 (2026-03-22)：
        '   合併原 GetMailCountAll + GetMailCountAllParallel 為單一函數，
        '   統一 fallback 鏈，呼叫端不再需要選擇要用哪個版本。
        '   GetMailCountAllParallel 可標記廢棄或直接刪除。
        '
        ' 設計說明：
        '   為何呼叫 GetMailCount() 而非直接用 GetTable()：
        '     PR_CONTENT_COUNT 是 Folder 物件上的已儲存屬性，Outlook 自動維護，讀取等於讀一個整數，一次 COM call 結束。
        '     GetTable() 會把資料夾內所有郵件 row 逐一回傳，只為了計數代價太高。GetTable 適合讀郵件內容 (大小、日期)，不適合純計數。
        '
        '   回傳型別 Long 而非 Integer：
        '     單一資料夾用 Integer 夠 (PR_CONTENT_COUNT 是 PT_LONG 32-bit)，
        '     但整棵子樹加總若有多個大資料夾，理論上可能超過 Integer.MaxValue (2,147,483,647)，用 Long 安全。
        '
        ' Fallback 鏈（依速度由快到慢）：
        '   ⓪ Redemption : rdoFolder.TotalItemCount
        '            MAPI 快取的彙總屬性，一次 COM call 直接取得整棵子樹總數，完全不需 BFS 遍歷
        '            Redemption 可正確讀取 PST 上此屬性（原生 OOM 的 PR_MESSAGE_SIZE_EXTENDED 在 PST 上無效）
        '            _rdoSession 未就緒時自動跳過此層
        '            注意：走此路徑時 onProgress callback 不會被觸發（無中間進度可回報）
        '   ① Task.WhenAll 平行 BFS：
        '            BFS 展開後每個資料夾各建一個 Task.Run，全部 WhenAll 等待
        '            Task.Run 內的 GetMailCount(f) 走 Redemption ⓪ 時是 free-threaded 安全的
        '            若 GetMailCount fallback 到 MAPI PropertyAccessor，仍有 STA 違規風險，需留意
        '   ② BFS 循序累加：
        '            GetSubFolderList BFS 展開 + GetMailCount(L3) 逐一加總
        '            支援取消檢查和 onProgress 進度回報
        '            平行路徑失敗時的安全 fallback
        '   ③ 遞迴 fallback：
        '            GetSubFolderList 本身失敗時 (極少見) 的最後保險
        '            無法精確回報進度，但確保加總結果正確
        '            todo: 這裡遞迴會重複呼叫 GetSubFolderList，若 ③ 常被觸發需檢查根本原因
        '   ④ Return -1：四層都失敗，由 L2 決定如何處理
        '
        ' cancelRequested：
        '   檢查 _cancelRequested 旗標，取消時回傳 -1，由 L1 判斷是否需要清空 UI
        '   ⓪ Redemption 路徑不插入取消檢查（單次 call，幾乎瞬間完成）
        '
        ' onProgress 參數 (可選)：
        '   傳入 Action(Of Integer, Integer) callback
        '   L2 每處理一個資料夾回報 (已完成數, 總數)，讓 L1 更新狀態列
        '   不需要進度回報時傳 Nothing
        '   ⓪ 和 ① 路徑不觸發 onProgress，② 路徑才會逐一回報
        '
        ' 取代：
        '   GetMailCountByMAPINew 的整棵子樹加總用途
        '   GetMailCountAllParallel (v3.0 已合併，舊版可廢棄)
        ' --------------------------------------------------------------
        Dbg("開始：GetMailCountAll v3.0", rootFolder.Name)

        ' ⓪ Redemption：TotalItemCount 直接回傳整棵子樹郵件總數
        '   一次 COM call 結束，不需要任何 BFS 遍歷或平行處理
        '   2026-03-22 新增
        If _rdo IsNot Nothing Then
            Dim rdoFolder As Redemption.RDOFolder = Nothing
            Try
                rdoFolder = _rdo.GetFolderFromID(rootFolder.EntryID, rootFolder.StoreID)
                Dim total As Long = CLng(rdoFolder.TotalItemCount)
                Dbg("GetMailCountAll ⓪ RDO 成功取得rdoFolder.TotalItemCount: ", $"{rootFolder.Name} | TotalItemCount={total}")
                Return total
            Catch ex As System.Exception
                Dbg("GetMailCountAll ⓪ RDO 失敗，走平行BFS fallback", $"{rootFolder.Name} | {ex.Message}")
            Finally
                If rdoFolder IsNot Nothing Then Marshal.ReleaseComObject(rdoFolder)
            End Try
        End If

        ' 2026/3/24 by AntiGravity: ① 平行 BFS (RDO)
        '   使用 GetSubFolderList_RDO 取得清單，以 Parallel.ForEach 搭配 Interlocked.Add 快速加總
        '   Redemption (RDO) 是 free-threaded，在背景平行執行安全且極為高效
        If _rdo IsNot Nothing Then
            Dim rdoRoot As Redemption.RDOFolder = Nothing
            Try
                rdoRoot = _rdo.GetFolderFromID(rootFolder.EntryID, rootFolder.StoreID)
                Dim rdoFolderList As List(Of Redemption.RDOFolder) = GetSubFolderList_RDO(rdoRoot, includeSubFolders:=True)
                Dim targetFolderCount As Integer = rdoFolderList.Count

                Dim totalCount As Long = 0
                Dim processedCount As Integer = 0

                Parallel.ForEach(rdoFolderList,
                    Sub(rdoF As Redemption.RDOFolder)
                        If _cancelRequested Then Return
                        Try
                            Dim count As Integer = rdoF.Items.Count
                            Interlocked.Add(totalCount, CLng(count))
                        Catch ex As System.Exception
                            Dbg("GetMailCountAll ① 略過失敗資料夾", rdoF.Name)
                        End Try
                        Dim done As Integer = Interlocked.Increment(processedCount)
                        onProgress?.Invoke(done, targetFolderCount)
                    End Sub)

                If _cancelRequested Then
                    Dbg("GetMailCountAll ① 已取消", $"總資料夾數：{targetFolderCount}") : Return -1
                End If

                Dbg("GetMailCountAll ① 平行BFS成功 (RDO)", $"{rootFolder.Name} | total={totalCount} | folders={targetFolderCount}")
                Return totalCount

            Catch ex As System.Exception
                Dbg("GetMailCountAll ① 平行BFS失敗，走循序BFS fallback", $"{rootFolder.Name} | {ex.Message}")
            Finally
                If rdoRoot IsNot Nothing Then Marshal.ReleaseComObject(rdoRoot)
            End Try
        End If

        ' ② BFS 循序累加：GetSubFolderList 展開 + GetMailCount(L3) 逐一加總
        '   支援取消檢查和 onProgress 進度回報，比平行版保守但穩定
        Try
            Dim targetFolderList As List(Of Outlook.Folder) = GetSubFolderList(rootFolder, includeSubFolders:=True)
            Dim grandTotal As Long = 0

            For i As Integer = 0 To targetFolderList.Count - 1
                If _cancelRequested Then
                    Dbg("GetMailCountAll ② 被取消", $"已處理 {i}/{targetFolderList.Count}") : Return -1
                End If

                Dim f As Outlook.Folder = targetFolderList(i)
                Dim count As Integer = GetMailCount(f)
                ' GetMailCount 的所有 fallback 都失敗才會到這個 else，記錄但不中止整體加總
                If count >= 0 Then grandTotal += CLng(count) Else Dbg("GetMailCountAll ② 略過失敗資料夾", f.Name)

                onProgress?.Invoke(i + 1, targetFolderList.Count)  ' 知道 total，進度條可以準確顯示百分比
                If i Mod 10 = 0 Then Await Task.Yield()            ' 每 10 個資料夾讓出一次，保持 UI 回應
            Next

            Dbg("GetMailCountAll ② 循序BFS成功", $"{rootFolder.Name} | total={grandTotal}")
            Return grandTotal

        Catch ex As System.Exception
            Dbg("GetMailCountAll ② 循序BFS失敗，走遞迴fallback", $"{rootFolder.Name} | {ex.Message}")
        End Try

        ' ③ 遞迴 fallback：GetSubFolderList 本身失敗時的最後保險
        '   無法精確回報進度，但確保加總結果正確
        '   注意：遞迴呼叫會重新進入本函數，⓪ Redemption 已失敗所以 _rdoSession 仍 Nothing 或故障
        '         ① ② 也已失敗，只會走到 ③ 再次遞迴——理論上 ③ 不會無限展開，因為每層只遞迴直屬子資料夾
        '         todo: 若 ③ 常被觸發，需回頭檢查 GetSubFolderList 失敗的根本原因
        Try
            Dim totalCount As Long = 0
            Dim count As Integer = GetMailCount(rootFolder)     ' 本層 mailcount
            If count >= 0 Then totalCount += count
            Await Task.Yield()

            For Each f As Outlook.Folder In rootFolder.Folders
                Dim subCount As Long = Await GetMailCountAll(f) ' 遞迴，每個直屬子資料夾各自展開
                If subCount >= 0 Then totalCount += subCount
            Next

            Dbg("GetMailCountAll ③ 遞迴fallback成功", $"{rootFolder.Name} | total={totalCount}")
            Return totalCount

        Catch ex As System.Exception
            Dbg("GetMailCountAll ③ 遞迴fallback也失敗", $"{rootFolder.Name} | {ex.Message}")
            Return -1   ' ④ 四層都失敗，回傳 -1 讓 L2 知道這是「讀取失敗」而非「真的是 0 封」
        End Try

    End Function

    Private Function GetFolderCount(folder As Outlook.Folder) As Integer
        ' --------------------------------------------------------------
        ' GetFolderCount：讀取單一資料夾的本層直屬子資料夾數
        '
        ' Fallback 鏈：
        '   ⓪ Redemption : RDOFolder.Folders.Count
        '            可從非 STA 執行緒呼叫，繞過 Outlook Security Guard
        '            _rdoSession 未就緒時自動跳過此層
        '   ① MAPI : PR_FOLDER_CHILD_COUNT (0x66380003, PT_LONG) 一次 PropertyAccessor call，在大多數情況下準確
        '            注意：PST 上此屬性在剛移動資料夾後可能短暫不同步，但 Outlook 關閉再開就會修正，日常使用可接受
        '            2026/3/20 實測：PR_FOLDER_CHILD_COUNT 沒有一次成功過，已暫時 comment 出
        '   ② OOM  : folder.Folders.Count
        '            Folders 集合比 Items 輕量，載入速度可接受，且永遠準確
        '   ③ fail : Return -1
        '
        ' 關於「先讀 PR_SUBFOLDERS (0x360A000B) 再讀個數」的設計討論：
        '   PR_SUBFOLDERS 是 PT_BOOLEAN，只告訴你有沒有子資料夾 (不告訴你幾個) 
        '   先讀它再讀 PR_FOLDER_CHILD_COUNT 等於多一次 COM call，只有「大多數資料夾都沒有子資料夾」時才划算，
        '   實際 PST 不符合此條件，因此直接讀 PR_FOLDER_CHILD_COUNT，不做 PR_SUBFOLDERS 前置判斷
        '
        ' 取代：散落各處的 folder.Folders.Count 直接呼叫 (建議逐一替換) 
        ' --------------------------------------------------------------
        'dbg("開始：GetFolderCount(): ", folder.Name)

        ' ⓪ Redemption：RDOFolder.Folders.Count
        '   與 OOM folder.Folders.Count 等價，但可在任意執行緒呼叫
        '   2026-03-22 新增
        If _rdo IsNot Nothing Then
            Dim rdoFolder As Redemption.RDOFolder = Nothing
            Try
                rdoFolder = _rdo.GetFolderFromID(folder.EntryID, folder.StoreID)
                Dim count As Integer = rdoFolder.Folders.Count
                Dbg("GetFolderCount ⓪ RDO 成功", $"{folder.Name} | count={count}")
                Return count
            Catch ex As System.Exception
                Dbg("GetFolderCount ⓪ RDO 失敗，走OOM fallback", $"{folder.Name} | {ex.Message}")
            Finally
                If rdoFolder IsNot Nothing Then Marshal.ReleaseComObject(rdoFolder)
            End Try
        End If

        ' ① MAPI：PR_FOLDER_CHILD_COUNT (0x66380003)
        ' 2026/3/20, 奇怪PR_FOLDER_CHILD_COUNT 沒有一次成功過??? 乾脆先拿掉這個try, 省得一直fallback也是浪費開銷
        'Try
        '    Const PR_FOLDER_CHILD_COUNT As String = "http://schemas.microsoft.com/mapi/proptag/0x66380003"
        '    Dim intFolderCount As Object = folder.PropertyAccessor.GetProperty(PR_FOLDER_CHILD_COUNT)
        '    If TypeOf intFolderCount Is Integer Then
        '        dbg("GetFolderCount ① MAPI成功: ", $"{folder.Name}")
        '        Return intFolderCount
        '    End If
        'Catch ex As System.Exception
        '    dbg("GetFolderCount ① MAPI失敗，走OOM fallback", $"{folder.Name} | {ex.Message}")
        'End Try

        ' ② OOM：folder.Folders.Count (準確，Folders 集合比 Items 輕量) 
        Try
            Return folder.Folders.Count
        Catch ex As System.Exception
            Dbg("GetFolderCount ② OOM也失敗", $"{folder.Name} | {ex.Message}")
        End Try

        ' ③ 若前兩層都失敗，回傳 -1 讓 L2 知道這是「讀取失敗」而非「真的是 0 封」
        Return -1

    End Function
    Private Async Function GetFolderCountAll(rootFolder As Outlook.Folder) As Task(Of Integer)
        ' --------------------------------------------------------------
        ' GetFolderCountAll：讀取某資料夾整棵子樹的資料夾總數 (不含 rootFolder 自身) 
        '
        ' 2026/3/24 by AntiGravity: Fallback 鏈 (由快到慢)：
        '   ⓪ Redemption + Parallel.ForEach (最快)：RDO 是 free-threaded，平行展開子樹
        '   ① Redemption + BFS 循序累加：RDO 循序，平行失敗時的安全路徑
        '   ② OOM + BFS 循序：無 Redemption 時，走 OOM COM 循序處理
        '   ③ Return -1：全部失敗
        '
        ' 取代：GetTotalFolderCountAsync (快取邏輯移至 L2 呼叫端) 
        '
        ' [Redemption說明] 2026-03-22
        '   此函數計算的是整棵子樹的遞迴總數，Redemption 沒有單一 API 可直接取得遞迴資料夾總數
        '   （rdoFolder.Folders.Count 只回傳直屬子資料夾數，與 OOM 相同）。
        '   因此此函數本身不需要直接加 Redemption 呼叫。
        '   ① BFS 路徑：GetSubFolderList 內部走 OOM folder.Folders 展開，展開後直接 .Count，不需 L3 讀取。
        '   ② 遞迴 fallback：內部的 rootFolder.Folders.Count 和 ForEach 走 OOM，
        '      若日後改為呼叫 GetFolderCount(L3)，即可自動走 Redemption ⓪ 路徑。
        ' --------------------------------------------------------------
        Dbg("開始：GetFolderCountAll(): ", rootFolder.Name)

        ' 2026/3/24 by AntiGravity: ⓪ Redemption + 平行處理 (最快路徑)
        '   使用 GetSubFolderList_RDO 取得清單，以 Parallel.ForEach 搭配 Interlocked.Add(rdoF.Folders.Count) 快速加總
        If _rdo IsNot Nothing Then
            Dim rdoRoot As Redemption.RDOFolder = Nothing
            Try
                rdoRoot = _rdo.GetFolderFromID(rootFolder.EntryID, rootFolder.StoreID)
                Dim rdoFolderList As List(Of Redemption.RDOFolder) = GetSubFolderList_RDO(rdoRoot, includeSubFolders:=True)

                Dim totalCount As Integer = 0
                Parallel.ForEach(rdoFolderList,
                    Sub(rdoF As Redemption.RDOFolder)
                        If _cancelRequested Then Return
                        Try
                            Dim count As Integer = rdoF.Folders.Count
                            Interlocked.Add(totalCount, count)
                        Catch ex As System.Exception
                            Dbg("GetFolderCountAll ⓪ RDO 略過失敗資料夾", rdoF.Name)
                        End Try
                    End Sub)

                If _cancelRequested Then
                    Dbg("GetFolderCountAll ⓪ 已取消", "") : Return -1
                End If

                Dbg("GetFolderCountAll ⓪ RDO平行成功", $"{rootFolder.Name} | total={totalCount}")
                Return totalCount

            Catch ex As System.Exception
                Dbg("GetFolderCountAll ⓪ RDO平行失敗，走OOM循序fallback", $"{rootFolder.Name} | {ex.Message}")
            Finally
                If rdoRoot IsNot Nothing Then Marshal.ReleaseComObject(rdoRoot)
            End Try
        End If

        ' 2026/3/24 by AntiGravity: ② OOM + BFS 循序 (無 Redemption 時的最後手段)
        '   必須循序處理 OOM COM 物件以避免 STA 違規
        Try
            Dim allFolders As List(Of Outlook.Folder) = GetSubFolderList(rootFolder, includeSubFolders:=True)
            Await Task.Yield()
            Dbg("GetFolderCountAll ② OOM BFS成功", $"{rootFolder.Name} | total={allFolders.Count - 1}")
            Return allFolders.Count - 1
        Catch ex As System.Exception
            Dbg("GetFolderCountAll ② OOM BFS失敗", $"{rootFolder.Name} | {ex.Message}")
        End Try

        ' ③ 全部失敗
        Return -1

    End Function

    Private Async Function GetFolderSize(folder As Outlook.Folder) As Task(Of Long)
        ' --------------------------------------------------------------
        ' GetFolderSize：讀取單一資料夾本層大小 (bytes)
        ' 2026/3/24 by AntiGravity: Fallback 鏈重構
        '   ⓪ Redemption : rdoFolder.Fields(PR_MESSAGE_SIZE_EXTENDED) (部分 Exchange 支援，極快)
        '   ① OOM  : folder.GetTable(PR_MESSAGE_SIZE_EXTENDED) + GetArray(1000) (最快安全招式)
        '   ② OOM  : folder.GetTable(PR_MESSAGE_SIZE_EXTENDED) + GetNextRow() (備案)
        '   ③ fail : Return -1
        ' --------------------------------------------------------------
        Dbg("開始：GetFolderSize", folder.Name)
        Dim sw As New Stopwatch() : sw.Start()

        ' ⓪ Redemption 層 (嘗試讀取資料夾本身的總量屬性)
        ' RDO 沒有 GetTable().GetArray()，故若屬性讀不到直接 fallback
        If _rdo IsNot Nothing Then
            Dim rdoFolder As Redemption.RDOFolder = Nothing
            Try
                rdoFolder = _rdo.GetFolderFromID(folder.EntryID, folder.StoreID)

                ' PR_MESSAGE_SIZE_EXTENDED (0x0E080014) 
                Const PR_SIZE_EX As Integer = &HE080014
                Dim val As Object = rdoFolder.Fields(PR_SIZE_EX)
                If val IsNot Nothing AndAlso Not IsDBNull(val) Then
                    Dim totalSize As Long = CLng(val)
                    sw.Stop()
                    Dbg("GetFolderSize ⓪ RDO Fields 成功", $"{folder.Name} | size={totalSize} | {sw.ElapsedMilliseconds}ms")
                    Return totalSize
                End If
            Catch ex As System.Exception
                Dbg("GetFolderSize ⓪ RDO 失敗，走 OOM GetArray fallback", $"{folder.Name} | {ex.Message}")
            Finally
                If rdoFolder IsNot Nothing Then Marshal.ReleaseComObject(rdoFolder)
            End Try
        End If

        Const PR_SIZE_EX_STR As String = "http://schemas.microsoft.com/mapi/proptag/0x0E080014"

        ' ① OOM GetTable + GetArray(1000) (目前最穩、最快的批次讀取)
        Dim table As Outlook.Table = Nothing
        Try
            table = folder.GetTable()
            table.Columns.RemoveAll()
            table.Columns.Add(PR_SIZE_EX_STR)

            Dim totalSize As Long = 0
            Do While Not table.EndOfTable
                Dim arr As Object = table.GetArray(1000)
                If arr Is Nothing Then Exit Do
                Dim data(,) As Object = DirectCast(arr, Object(,))
                For r As Integer = 0 To data.GetUpperBound(0)
                    Dim sz = data(r, 0)
                    If sz IsNot Nothing AndAlso Not IsDBNull(sz) Then totalSize += CLng(sz)
                Next
                Await Task.Yield() ' 讓出 UI 避免卡死
            Loop
            sw.Stop()
            Dbg("GetFolderSize ① OOM GetTable.GetArray 成功", $"{folder.Name} | size={totalSize} | {sw.ElapsedMilliseconds}ms")
            Return totalSize
        Catch ex As System.Exception
            Dbg("GetFolderSize ① OOM GetArray 失敗，走 GetNextRow fallback", $"{folder.Name} | {ex.Message}")
        Finally
            If table IsNot Nothing Then Marshal.ReleaseComObject(table)
        End Try

        ' ② OOM GetTable + GetNextRow() (不依賴二維陣列的最後保險)
        Dim table2 As Outlook.Table = Nothing
        Try
            table2 = folder.GetTable()
            table2.Columns.RemoveAll()
            table2.Columns.Add(PR_SIZE_EX_STR)

            Dim totalSize As Long = 0
            Dim loopCount As Integer = 0
            Do While Not table2.EndOfTable
                Dim row As Outlook.Row = table2.GetNextRow()
                If row IsNot Nothing Then
                    Dim sz = row(PR_SIZE_EX_STR)
                    If sz IsNot Nothing AndAlso Not IsDBNull(sz) Then totalSize += CLng(sz)
                    Marshal.ReleaseComObject(row)
                End If
                loopCount += 1
                If loopCount Mod 500 = 0 Then Await Task.Yield()
            Loop
            sw.Stop()
            Dbg("GetFolderSize ② OOM GetNextRow 成功", $"{folder.Name} | size={totalSize} | {sw.ElapsedMilliseconds}ms")
            Return totalSize
        Catch ex As System.Exception
            Dbg("GetFolderSize ② OOM GetNextRow 失敗", $"{folder.Name} | {ex.Message}")
        Finally
            If table2 IsNot Nothing Then Marshal.ReleaseComObject(table2)
        End Try

        sw.Stop()
        Dbg("結束：GetFolderSize (FAIL)", $"{folder.Name} | -1 | {sw.ElapsedMilliseconds}ms")
        Return -1
    End Function
    Private Async Function GetFolderSizeAll(rootFolder As Outlook.Folder) As Task(Of Long)
        ' --------------------------------------------------------------
        ' GetFolderSizeAll v1.0：讀取某資料夾及整棵子樹的大小總計 (bytes)
        ' 
        ' 2026/3/24 by AntiGravity: 落實新的 Fallback 鏈設計，並修正平行處理的 STA 問題
        '   ⓪ Redemption 平行路徑 (最快): 
        '      利用 GetSubFolderList_RDO 一次把該子樹下所有 RDOFolder 拿出來，
        '      放到 Parallel.ForEach 中，各別讀取 MAPI 屬性 PR_MESSAGE_SIZE_EXTENDED。
        '      (RDOFolder 不支援 GetTable().GetArray()，故依賴屬性直讀)
        '
        '   ① OOM 循序路徑 (最安全): 
        '      當 RDO 平行路徑失敗（或是未匯入 Redemption），退回使用 OOM。
        '      OOM 絕對不可以在 Task.Run / WhenAll 等背景執行緒內呼叫 COM，否則會觸發 STA 錯誤。
        '      故改為嚴格的 For 迴圈，逐一 Await GetFolderSize()。
        '      而內部的 GetFolderSize 會走到它專屬的 GetTable().GetArray(1000) OOM 極速路徑。
        '
        '   ② 兩層都失敗：回傳 -1，交給上一層流程處理。
        ' --------------------------------------------------------------
        Dbg("開始：GetFolderSizeAll v1.0", rootFolder.Name)

        ' 2026/3/24 by AntiGravity: ⓪ Redemption 平行累加 PR_MESSAGE_SIZE_EXTENDED
        If _rdo IsNot Nothing Then
            Dim rdoRoot As Redemption.RDOFolder = Nothing
            Try
                rdoRoot = _rdo.GetFolderFromID(rootFolder.EntryID, rootFolder.StoreID)
                Dim rdoFolderList As List(Of Redemption.RDOFolder) = GetSubFolderList_RDO(rdoRoot, includeSubFolders:=True)

                Dim grandTotal As Long = 0
                Const PR_SIZE_EX As Integer = &HE080014

                ' 利用 Parallel.ForEach 與 Interlocked.Add 達到極致的多核並發加總
                Dim validCount As Integer = 0
                Parallel.ForEach(rdoFolderList,
                    Sub(rdoF As Redemption.RDOFolder)
                        If _cancelRequested Then Return
                        Try
                            Dim val As Object = rdoF.Fields(PR_SIZE_EX)
                            If val IsNot Nothing AndAlso Not IsDBNull(val) Then
                                Interlocked.Add(grandTotal, CLng(val))
                                Interlocked.Increment(validCount)
                            End If
                        Catch ex As System.Exception
                            Dbg("GetFolderSizeAll ⓪ RDO 略過讀取失敗的資料夾", rdoF.Name)
                        End Try
                    End Sub)

                If _cancelRequested Then
                    Dbg("GetFolderSizeAll ⓪ 已取消", $"總資料夾數：{rdoFolderList.Count}") : Return -1
                End If

                If validCount = 0 AndAlso rdoFolderList.Count > 0 Then
                    Dbg("GetFolderSizeAll ⓪ RDO 讀取失敗（無支援的屬性）", "退回 OOM")
                    Throw New System.Exception("RDO PR_SIZE_EX returned empty for all folders")
                End If

                Dbg("GetFolderSizeAll ⓪ RDO平行成功", $"{rootFolder.Name} | totalSize={grandTotal} | folders={rdoFolderList.Count}")
                Return grandTotal

            Catch ex As System.Exception
                Dbg("GetFolderSizeAll ⓪ RDO平行失敗，走 OOM 循序 fallback", $"{rootFolder.Name} | {ex.Message}")
            Finally
                If rdoRoot IsNot Nothing Then Marshal.ReleaseComObject(rdoRoot)
            End Try
        End If

        ' 2026/3/24 by AntiGravity: ① OOM 循序 BFS 累加 (避免 STA 錯誤的保險路徑)
        ' 因為 OOM 的 GetTable() 必須在 UI Thread，我們必須循序 Await 每一層
        Try
            Dim targetFolderList As List(Of Outlook.Folder) = GetSubFolderList(rootFolder, includeSubFolders:=True)
            Dim grandTotal As Long = 0

            For i As Integer = 0 To targetFolderList.Count - 1
                If _cancelRequested Then
                    Dbg("GetFolderSizeAll ① 被取消", $"已處理 {i}/{targetFolderList.Count}") : Return -1
                End If

                Dim f As Outlook.Folder = targetFolderList(i)
                Dim sz As Long = Await GetFolderSize(f) ' GetFolderSize 內部會使用最快的 GetTable().GetArray(1000)

                If sz >= 0 Then
                    grandTotal += sz
                Else
                    Dbg("GetFolderSizeAll ① 略過了大小計算失敗的資料夾", f.Name)
                End If

                ' 避免卡死 UI
                If i Mod 5 = 0 Then Await Task.Yield()
            Next

            Dbg("GetFolderSizeAll ① 循序BFS成功", $"{rootFolder.Name} | totalSize={grandTotal}")
            Return grandTotal

        Catch ex As System.Exception
            Dbg("GetFolderSizeAll ① 循序BFS失敗，放棄計算", $"{rootFolder.Name} | {ex.Message}")
        End Try

        ' ② 兩層都失敗，回傳 -1 讓呼叫端知道失敗了
        Return -1
    End Function

    Private Function GetMailSize(item As Object) As Long
        ' --------------------------------------------------------------
        ' GetMailSize：讀取單封郵件的大小 (bytes)，供 GetFolderSize fallback 路徑呼叫
        '
        ' Fallback 鏈：
        '   ⓪ Redemption : RDOMail.Size
        '            free-threaded 安全，可在 Task.Run 內呼叫
        '            繞過 Outlook Security Guard，不會彈出安全性警告
        '            _rdo 未就緒時自動跳過此層
        '   ① MAPI : PR_MESSAGE_SIZE_EXTENDED (0x0E080014, PT_I8, 64-bit Long)
        '             避免 PR_MESSAGE_SIZE (PT_LONG, 32-bit) 在超大郵件時溢位
        '   ② MAPI : PR_MESSAGE_SIZE (0x0E080003, PT_LONG, 32-bit Integer)
        '             Fallback 到 32-bit 版本，CInt → CLng 安全轉型
        '   ③ OOM  : mail.Size
        '             最後手段，OOM 的 Size 屬性單位是 bytes，回傳 Integer，
        '             大郵件 (>2GB) 理論上會溢位，但實務上 Outlook 的 PST 限制在 50GB 總量，
        '             單封郵件超過 2GB 極不可能，此層可視為安全
        '
        ' 注意：此函數接受 Object 型別參數，是因為 GetFolderSize 的 fallback 路徑
        '       用 Items.GetFirst/GetNext 取回的是 Object，省去呼叫端的 TryCast 成本
        '       若是 MailItem 就正常讀取，若是其他型別 (Contact、Appointment 等) 就回 0
        '
        ' 取代：GetFolderSizeOld_Async 內的 mailItem.Size 直接呼叫
        '       行 3385 的同名 stub (完整替換)
        ' --------------------------------------------------------------

        ' 非 MailItem 的項目 (Calendar、Contact 等) 直接略過，回 0
        Dim mail As Outlook.MailItem = TryCast(item, Outlook.MailItem)
        If mail Is Nothing Then Return 0

        ' ⓪ Redemption：RDOMail.Size
        '   GetMessageFromID 的 StoreID 從 mail.Parent 取得，多一次 COM call 但避免跨 PST 找錯 item
        '   2026-03-22 新增
        If _rdo IsNot Nothing Then
            Dim rdoMail As Redemption.RDOMail = Nothing
            Try
                Dim parentFolder As Outlook.Folder = TryCast(mail.Parent, Outlook.Folder)
                Dim storeId As String = If(parentFolder IsNot Nothing, parentFolder.StoreID, "")
                rdoMail = TryCast(_rdo.GetMessageFromID(mail.EntryID, storeId), Redemption.RDOMail)
                If rdoMail IsNot Nothing Then
                    Dim sz As Long = CLng(rdoMail.Size)
                    Dbg("GetMailSize ⓪ RDO 成功", $"size={sz}")
                    Return sz
                End If
            Catch ex As System.Exception
                Dbg("GetMailSize ⓪ RDO 失敗，走MAPI fallback", ex.Message)
            Finally
                If rdoMail IsNot Nothing Then Marshal.ReleaseComObject(rdoMail)
            End Try
        End If

        ' ① MAPI：PR_MESSAGE_SIZE_EXTENDED (0x0E080014, PT_I8) — 64-bit，無溢位風險
        Try
            Const PR_SIZE_EXTENDED As String = "http://schemas.microsoft.com/mapi/proptag/0x0E080014"
            Dim val As Object = mail.PropertyAccessor.GetProperty(PR_SIZE_EXTENDED)
            If TypeOf val Is Long Then Return CLng(val)
            If TypeOf val Is Integer Then Return CLng(CInt(val))    ' 某些環境回傳 Integer，安全轉型
            ' todo: try/catch裡面包住的 TypeOf 都可以直接拿掉
        Catch ex As System.Exception
            Dbg("GetMailSize ① PR_MESSAGE_SIZE_EXTENDED失敗", ex.Message)
        End Try

        ' ② MAPI：PR_MESSAGE_SIZE (0x0E080003, PT_LONG) — 32-bit，超大郵件理論上溢位
        Try
            Const PR_SIZE As String = "http://schemas.microsoft.com/mapi/proptag/0x0E080003"
            Dim val As Object = mail.PropertyAccessor.GetProperty(PR_SIZE)
            If TypeOf val Is Integer Then Return CLng(CInt(val))
            ' todo: try/catch裡面包住的 TypeOf 都可以直接拿掉
        Catch ex As System.Exception
            Dbg("GetMailSize ② PR_MESSAGE_SIZE失敗", ex.Message)
        End Try

        ' ③ OOM：mail.Size (Integer，超大郵件理論上不準，但實務上 PST 內不會發生)
        Try
            Return CLng(mail.Size)
        Catch ex As System.Exception
            Dbg("GetMailSize ③ OOM mail.Size也失敗", ex.Message)
        End Try

        Return -1
    End Function
#End Region

#Region "■ 99 舊版備用 (勿刪)"

    Private Function GetMailCountRecursive(folder As Outlook.Folder) As Integer
        Dbg("開始：", folder.Name)

        Dim value As Integer
        If _mailCountCache.TryGetValue(folder, value) Then Return value ' 檢查快取中是否已存在值, 若有則直接返回

        ' 改成先用 Parallel.ForEach 遍歷子文件夾並且並行處理
        Dim totalMailCount As Integer = 0
        Dim countingBag As New ConcurrentBag(Of Integer)()
        Try
            ' 5/21記錄: 模仿GetFolderSizeLegacy那一句超快速的LINQ, 但測試結果沒有現在這個快, 所以決定保留這個
            ' 2026/3/20, 重寫了底層GetMailCountAll() 但是不知為何效能還是比不過現在下面這個遞迴版本?? (todo: 暫時先保留)
            ' 原因: 原版遞迴只走一遍 COM 資料夾樹，新版走了兩遍COM：
            ' 第一遍：GetSubFolderList()    → BFS 遍歷，存取每個 folder.Folders
            ' 第二遍：For Each allFolders   → GetMailCount() 再讀每個資料夾一次
            ' 2026/3/22, 導入Redemption, 應該可以刪掉這裡了? 還是讓Redemption 變成on-demand, 需要才啟動?
            Parallel.ForEach(folder.Folders.Cast(Of Outlook.Folder),' 取得子資料夾的郵件數量並添加到 ConcurrentBag 中
                             Sub(subFolder As Outlook.Folder)
                                 countingBag.Add(GetMailCountRecursive(subFolder))
                             End Sub)
            totalMailCount = countingBag.Sum() ' 累加所有子資料夾的郵件數量

            ''' 最後再獲取選取文件夾自身的郵件數量 (改用MAPI table 的PR_CONTENT_COUNT屬性來getmailcount)
            ''Const PR_CONTENT_COUNT As String = "http://schemas.microsoft.com/mapi/proptag/0x36020003"
            ''totalMailCount += folder.PropertyAccessor.GetProperty(PR_CONTENT_COUNT)
            totalMailCount += GetMailCount(folder)  ' 單一目錄的mail count改成重寫的統一底層函數, 2026/3/20

            _mailCountCache.TryAdd(folder, totalMailCount) ' 第一次計算後就存入快取
        Catch
        End Try

        Return totalMailCount
    End Function
    Private Async Function GetMailCountAll_1(rootFolder As Outlook.Folder, Optional onProgress As Action(Of Integer, Integer) = Nothing) As Task(Of Long)
        ' --------------------------------------------------------------
        ' GetMailCountAll_1：讀取某資料夾及其整棵子樹的郵件總數
        ' 先RDO, 再BFS累加, 再遞迴
        '
        ' 設計說明：不自己遞迴，改用 GetSubFolderList() BFS 展開完整資料夾清單後逐一加總
        '           ① 可以在任意點插入取消檢查 (_cancelRequested) ，遞迴做不到
        '           ② 知道總資料夾數，可以回報準確進度
        '           ③ 多個統計 (mail count + folder size) 可共用同一份清單，不用各跑一次 BFS ' todo: 有沒有地方可以用上這個好處??
        '           ④ 沒有 stack overflow 風險 (BFS 用 Queue，不用 call stack)
        '
        '           為何呼叫 GetMailCount() 而非直接用 GetTable()：
        '             PR_CONTENT_COUNT 是 Folder 物件上的已儲存屬性，Outlook 自動維護，讀取等於讀一個整數，一次 COM call 結束。
        '             GetTable() 會把資料夾內所有郵件 row 逐一回傳，只為了計數代價太高。GetTable 適合讀郵件內容 (大小、日期) ，不適合純計數。
        '
        '           回傳型別 Long 而非 Integer：
        '             單一資料夾用 Integer 夠 (PR_CONTENT_COUNT 是 PT_LONG 32-bit) ，
        '             但整棵子樹加總若有多個大資料夾，理論上可能超過 Integer.MaxValue (2,147,483,647) ，用 Long 安全。
        '
        ' Fallback 鏈：
        '   ⓪ Redemption : rdoFolder.TotalItemCount
        '            直接回傳整棵子樹的郵件總數，MAPI 層面的快取彙總值，一次 COM call 結束，完全不需要 BFS 遍歷
        '            Redemption 可正確讀取 PST 上此屬性（原生 OOM 無法取得）
        '            _rdoSession 未就緒時自動跳過此層
        '   ① GetSubFolderList + GetMailCount(L3) 逐一加總, BFS 展開後逐一呼叫，清單與計算邏輯分離，支援取消和進度回報
        '   ② 遞迴 fallback: GetSubFolderList 本身失敗時 (極少見) 的保險方案, 遞迴版本無法回報精確進度，但確保結果正確
        '   ③ 兩層都失敗就回傳 Return -1 並記錄 DebugForm，不讓單一資料夾的讀取失敗影響整體加總。
        '
        ' cancelRequested 參數：' todo: 如何使用??
        '   傳入 _cancelRequested 旗標的 ByRef，讓呼叫端可以中途 ESC 取消
        '   取消時回傳 -1，由 L1 判斷是否需要清空 UI
        '
        ' onProgress 參數 (可選) ：' todo: 如何使用??
        '   傳入 Action(Of Integer, Integer) callback，
        '   L2 每處理一個資料夾回報 (已完成數, 總數)，讓 L1 更新狀態列
        '   不需要進度回報時傳 Nothing
        '   注意：⓪ Redemption 路徑一次取得結果，不會觸發 onProgress callback（無中間進度可回報）
        '
        ' 取代：GetMailCountByMAPINew 的整棵子樹加總用途
        '       (GetMailCountByMAPINew 內的 Parallel.ForEach 遞迴整段, 效能超快, 但不是好的做法)
        '
        ' 2026-03-22 新增 ⓪ Redemption TotalItemCount，_rdoSession 就緒時完全跳過 BFS
        ' --------------------------------------------------------------
        Dbg("開始：GetMailCountAll(): ", rootFolder.Name)

        ' ⓪ Redemption：TotalItemCount 直接回傳整棵子樹郵件總數
        '   MAPI 快取的彙總屬性，一次 call 結束，不需要 BFS 遍歷，也不需要平行處理
        '   原生 OOM 的 PR_MESSAGE_SIZE_EXTENDED 在 PST 上找不到，Redemption 可正確讀取
        '   2026-03-22 新增
        If _rdo IsNot Nothing Then
            Dim rdoFolder As Redemption.RDOFolder = Nothing
            Try
                rdoFolder = _rdo.GetFolderFromID(rootFolder.EntryID, rootFolder.StoreID)
                Dim total As Long = rdoFolder.TotalItemCount
                Dbg("GetMailCountAll ⓪ RDO 成功", $"{rootFolder.Name} | TotalItemCount={total}")
                Return total
            Catch ex As System.Exception
                Dbg("GetMailCountAll ⓪ RDO 失敗，走BFS fallback", $"{rootFolder.Name} | {ex.Message}")
            Finally
                If rdoFolder IsNot Nothing Then Marshal.ReleaseComObject(rdoFolder)
            End Try
        End If

        ' ① 標準路徑：GetSubFolderList BFS 展開 + GetMailCount(L3) 逐一加總
        Try
            ' BFS 展開整棵子樹的資料夾清單 (復用現有函數，不重寫)
            Dim targetFolderList As List(Of Outlook.Folder) = GetSubFolderList(rootFolder, includeSubFolders:=True)

            Dim grandTotal As Long = 0
            For i As Integer = 0 To targetFolderList.Count - 1
                If _cancelRequested Then    ' ✅ 取消檢查：任意點都可以乾淨中止，不像遞迴版難以插入
                    Dbg("GetMailCountAll 被取消", $"已處理 {i}/{targetFolderList.Count}") : Return -1
                End If

                Dim f As Outlook.Folder = targetFolderList(i)
                Dim count As Integer = GetMailCount(f)
                ' GetMailCount 的所有 fallback 都失敗才會到這個else，記錄但不中止整體加總
                If count >= 0 Then grandTotal += CLng(count) Else Dbg("GetMailCountAll 略過失敗資料夾", f.Name)

                onProgress?.Invoke(i + 1, targetFolderList.Count) ' 進度回報 (optional callback，呼叫端不需要時傳 Nothing 即可) 因為知道 total，進度條可以準確顯示百分比 'todo: 這個進度回報如何使用?
                If i Mod 10 = 0 Then Await Task.Yield()     ' 每掃瞄10個資料夾處理完就讓出一次，保持 UI 回應 (GetMailCount 本身是同步的，所以這裡的 Yield 是唯一的讓出點)
            Next
            Return grandTotal

        Catch ex As System.Exception
            Dbg("GetMailCountAll ① BFS路徑失敗，走遞迴fallback", $"{rootFolder.Name} | {ex.Message}")
        End Try

        ' ② 遞迴 fallback：GetSubFolderList 本身失敗時使用 (無法精確回報進度，但至少確保加總結果正確)
        '   注意：遞迴層數受 PST 資料夾巢狀深度限制，實務上 PST 不會太深
        Try
            Dim totalCount As Long = 0
            Dim count As Integer = GetMailCount(rootFolder) ' 本層mailcount
            If count >= 0 Then totalCount += count : Await Task.Yield()

            For Each f As Outlook.Folder In rootFolder.Folders
                Dim subCount As Long = Await GetMailCountAll_1(f) ' todo: 這裡遞迴的話, 會一直重複呼叫上面的GetSubFolderList(), 會跑到死....
                If subCount >= 0 Then totalCount += subCount
            Next
            Return totalCount

        Catch ex As System.Exception ' ③ 全部失敗就傳回 -1 讓上層流程去處理
            Dbg("GetMailCountAll ② 遞迴fallback也失敗", $"{rootFolder.Name} | {ex.Message}")
            Return -1   ' ③ 若前兩層都失敗，回傳 -1 讓 L2 知道這是「讀取失敗」而非「真的是 0 封」
        End Try

    End Function
    Private Async Function GetMailCountAll_2(rootFolder As Outlook.Folder, Optional onProgress As Action(Of Integer, Integer) = Nothing) As Task(Of Long)
        ' --------------------------------------------------------------
        ' 就是 GetMailCountAllParallel  v2.0：讀取某資料夾及其整棵子樹的郵件總數
        ' 先RDO, 再平行, 再BFS累加
        '
        ' 平行策略：
        '   BFS 展開後，對每個資料夾各建一個 Task.Run，全部 Task.WhenAll 等待。不需再依PST StoreID 分組，結構最簡潔。
        '   PR_CONTENT_COUNT 是 Folder 上的已快取屬性，bottleneck 是 cross-process COM overhead，Outlook.exe 端能否真正並發處理需實測確認。
        '
        ' [2026-03-22 重要說明] Redemption 就緒後此函數實質上已被 GetMailCountAll ⓪ 取代
        '   原本設計平行處理是為了加速 BFS 逐一累加的瓶頸，
        '   但 Redemption 的 TotalItemCount 一次 call 就取得整棵子樹總數，
        '   平行處理的必要性消失。此函數保留作為：
        '   (a) _rdoSession 未就緒時的備用高速路徑（走 Task.WhenAll 平行版）
        '   (b) 將來跨 PST 加總時的協調層（多個 PST 的 GetMailCountAll 可以 Task.WhenAll）
        '   若確認 Redemption 穩定，日後可考慮廢棄此函數，呼叫端直接改用 GetMailCountAll。
        '
        ' [Redemption說明] 2026-03-22
        '   ⓪ Redemption TotalItemCount 一次取得，走此路徑時整個平行展開邏輯完全跳過
        '   ① Task.WhenAll 平行路徑：_rdoSession 未就緒時的 fallback
        '      Task.Run 內的 GetMailCount(f) 若走 Redemption ⓪，是 free-threaded 安全的
        '      若 fallback 到 MAPI PropertyAccessor，仍有 STA 違規風險，需留意
        ' --------------------------------------------------------------
        Dbg("開始：GetMailCountAllParallel(): ", rootFolder.Name)

        ' ⓪ Redemption：TotalItemCount 直接回傳整棵子樹郵件總數
        '   就緒時完全跳過下方所有平行 BFS 邏輯，等同於 GetMailCountAll ⓪ 的行為
        '   2026-03-22 新增
        If _rdo IsNot Nothing Then
            Dim rdoFolder As Redemption.RDOFolder = Nothing
            Try
                rdoFolder = _rdo.GetFolderFromID(rootFolder.EntryID, rootFolder.StoreID)
                Dim total As Long = rdoFolder.TotalItemCount
                Dbg("GetMailCountAllParallel ⓪ RDO 成功", $"{rootFolder.Name} | TotalItemCount={total}")
                Return total
            Catch ex As System.Exception
                Dbg("GetMailCountAllParallel ⓪ RDO 失敗，走平行BFS fallback", $"{rootFolder.Name} | {ex.Message}")
            Finally
                If rdoFolder IsNot Nothing Then Marshal.ReleaseComObject(rdoFolder)
            End Try
        End If

        ' ① 標準路徑：BFS 展開 → 每個資料夾一個 Task → Task.WhenAll
        Try
            Dim targetFolderList As List(Of Outlook.Folder) = GetSubFolderList(rootFolder, includeSubFolders:=True)
            Dim targetFolderCount As Integer = targetFolderList.Count
            Dim processedCount As Integer = 0   ' Interlocked 保證多 Task 同時更新的執行緒安全

            Dim folderTasks = targetFolderList.Select(
                Function(f) Task.Run(Function() As Integer
                                         If _cancelRequested Then Return 0

                                         Dim count As Integer = GetMailCount(f)
                                         If count < 0 Then
                                             Dbg("GetMailCountAllParallel 略過失敗資料夾", f.Name)
                                             count = 0
                                         End If

                                         Dim done As Integer = Interlocked.Increment(processedCount)
                                         onProgress?.Invoke(done, targetFolderCount) : Return count
                                     End Function)).ToList()

            Dim results As Integer() = Await Task.WhenAll(folderTasks)

            If _cancelRequested Then
                Dbg("GetMailCountAllParallel 已取消", $"總資料夾數：{targetFolderCount}") : Return -1
            End If
            Return results.Sum(Function(c) CLng(c))

        Catch ex As System.Exception
            Dbg("GetMailCountAllParallel ① 平行路徑失敗，走循序fallback", $"{rootFolder.Name} | {ex.Message}")
        End Try

        ' ② 循序 fallback：平行路徑失敗時使用，退回單純的逐一加總
        '   不用遞迴 (避免重複呼叫 GetSubFolderList) ，直接重跑 BFS 循序版
        Try
            Dim targetFolderList As List(Of Outlook.Folder) = GetSubFolderList(rootFolder, includeSubFolders:=True)
            Dim grandTotal As Long = 0

            For i As Integer = 0 To targetFolderList.Count - 1
                If _cancelRequested Then Return -1
                Dim count As Integer = GetMailCount(targetFolderList(i))
                If count >= 0 Then grandTotal += CLng(count)
                If i Mod 10 = 0 Then Await Task.Yield()     ' 每掃瞄10個資料夾處理完就讓出一次，保持 UI 回應
            Next
            Return grandTotal

        Catch ex As System.Exception
            Dbg("GetMailCountAllParallel ② 循序fallback也失敗", $"{rootFolder.Name} | {ex.Message}")
            Return -1       ' ③ 若前兩層都失敗，回傳 -1 讓 L2 知道這是「讀取失敗」而非「真的是 0 封」
        End Try
    End Function
    Private Async Function GetFolderSizeLegacy(folder As Outlook.Folder) As Task(Of Long)
        ' ==============================================================
        ' === GetFolderSizeLegacy — 修正版 (移除 Task.Run 包 COM) ===
        ' ==============================================================
        '
        ' 原版問題：Task.Run(Function() folder.Items.Cast(Of Object)().Sum(Function(s) s.Size))
        '          在 thread pool 執行緒上操作 Outlook COM 物件，違反 STA 規定, 在特定情況 (COM interop 敏感時機) 會造成 crash 或傳回錯誤結果
        '
        ' 修正做法：GetTable + PR_MESSAGE_SIZE 在 UI 執行緒循序讀取
        '           GetTable 回傳 MAPI binary table (低層讀取) 
        '          一次只讀一個 Row，每個 Row 用後立即 ReleaseComObject，避免 RCW 累積
        '          每 100 筆 Yield 一次讓 UI 保持回應
        '          速度接近原版 LINQ (實測差距在誤差範圍內) ，但 STA 安全
        '
        ' 此函數仍為 Lazy (不主動觸發) ：
        '   由 ListView1_ColumnClick 或右鍵選單「Show This Folder Size」觸發
        '   結果存入 folderSizeCache，BuildListViewItem_Tab1 下次組裝時自動顯示
        ' ==============================================================
        Dbg("開始：GetFolderSizeLegacy: ", folder.Name)

        Dim value As Long   ' 快取命中直接回傳
        If _folderSizeCache.TryGetValue(folder, value) Then Return value

        '' 已知有問題的資料夾走舊路徑 (不明 COM 例外物件，GetTable 也可能出問題) 
        'Dim exceptList As String() = {"Inbox_2000~2018", "Facebook"}
        'If exceptList.Contains(folder.Name) Then Return GetFolderSizeOld_Async(folder)

        Dim table As Outlook.Table = Nothing
        Try
            ' GetTable + PR_MESSAGE_SIZE (0x0E080003) ：
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
                Try
                    Dim sz As Object = row(PR_SIZE_EXTENDED)
                    totalSize += CLng(CInt(sz))
                Catch ex As System.Exception : Dbg("GetFolderSizeLegacy 例外: " & folder.Name, ex.Message)
                Finally : Marshal.ReleaseComObject(row)   ' 每個 Row 用完立即釋放，避免 RCW 累積
                End Try
                rowCount += 1
                If rowCount Mod 100 = 0 Then Await Task.Yield()  ' 每 100 筆統計就讓 UI 回應一次
            Loop

            _folderSizeCache.TryAdd(folder, totalSize)
            Return totalSize

        Catch ex As OverflowException
            Dbg("Error: GetFolderSizeLegacy overflow", folder.Name)
            Return -1
        Catch ex As System.Exception
            Dbg("Error: GetFolderSizeLegacy", folder.Name & " - " & ex.Message)
            Return -1
        Finally
            If table IsNot Nothing Then Marshal.ReleaseComObject(table)
        End Try
    End Function

#End Region

End Class


