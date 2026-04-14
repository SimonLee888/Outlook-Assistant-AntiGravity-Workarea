Imports System.Numerics
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Windows.Forms.DataVisualization.Charting
Imports Microsoft.Office.Interop.Outlook
Imports Windows.Graphics.Printing.OptionDetails
Imports Outlook = Microsoft.Office.Interop.Outlook

'Imports Redemption      ' 2026/3/22 正式導入Redemption, 測試logon成功, 傳回數值成功
'Imports MailKit        ' MailKit is a cross-platform mail client library built on top of MimeKit.
'Imports MailKit.Search
'Imports System
'Imports System.Core.dll
'Imports System.ComponentModel
'Imports System.ComponentModel.Design.ObjectSelectorEditor
'Imports System.Collections.Concurrent
'Imports System.Diagnostics.Metrics
'Imports System.DirectoryServices.ActiveDirectory
'Imports System.Linq.Parallel.dll
'Imports System.Net
'Imports System.Threading
'Imports System.Windows.Controls
'Imports System.Windows.Forms.VisualStyles.VisualStyleElement
'Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Header
'Imports Windows.Graphics.Printing.OptionDetails
'Imports Microsoft.VisualBasic.Devices

<System.ComponentModel.DesignerCategory("Form")>
Partial Class Form1

#Region "■ 01 全域宣告"
    <System.Diagnostics.Conditional("DEBUG")>
    Private Sub _dbg(Optional msg As String = "", Optional detail As String = "")
        ' 2026/03/31 by Gemini: 改用 DebugForm 統一提供的 GetCallerName，此版本支援解析 Async 非同步方法名稱
        Dim realCaller As String = DebugForm.GetCallerName()
        If _isDebugMode Then DebugForm.AddMessage3(msg, detail, realCaller)

    End Sub

    'Private _isFirstInit As Boolean = True          ' 第一次啟動程式
    ' by Gemini, 2026/04/01: 延遲載入 UI 的狀態旗標
    ' Index   0: 取代原 _isFirstInit，標記 Form 與 Tab1 是否處於「首次啟動/首次選定」階段 (True=首次啟動中)
    ' Index 1~5: 對應 Tab1~Tab5 的 UI 是否已完成掛載 (True=已完成)
    Private _isTabInitialized(5) As Boolean         ' 記錄每個 Tab 的 UI 是否已經初始化完成, (0)是FormLoad的第一次啟動, (1)~(5)分別對應 Tab1~Tab5
    Private _isUserBusy As Boolean = False          ' ✅ 2026/04/01 by Gemini: 使用者操作忙碌旗標，用於暫緩背景預載程序
    Private _isDebugMode As Boolean                 ' 是否為 Debug 模式，根據 VS 的編譯組態自動設定，是否顯示 DebugForm 以及是否啟用內部調試訊息
    Private _iLikeNoisy As Boolean = False          ' 是否啟用過濾debug message 噪音的功能，預設為 False 不顯示高頻率的迴圈訊息，想要詳細訊息轟炸就切成 True
    Private _isClosing As Boolean = False           ' added by Gemini, 2026/04/08: 關閉流程旗標，確保 FormClosing 中的非同步儲存完成後再釋放資源並允許關閉
    Private _cancelRequested As Boolean = False     ' ESC 全域中斷旗標: Tab1/Tab2/Tab3 共用，按 ESC 立刻設 True，各操作在 Yield 點檢查
    Private _cts As CancellationTokenSource         ' ✅ 2026/04/10: 導入現代化非同步中斷信號源作ESC中斷取代布林旗標
    '2026/3/10重構時停止使用全域變數來記錄遞迴過程中的資料, 改用傳遞參數以避免多線程或重入呼叫時資料被改寫的問題
    'Private _intTotalMailCount As Integer          ' 在遞迴中, 記錄點選資料夾內的所有郵件總數, 不要被遞迴呼叫改變數量
    'Private _intProcessedCount As Integer          ' 在遞迴中, 加總已處理的郵件總數, 不要被遞迴呼叫改變數量
    ' Private _isTab3_Stop As Boolean                 ' 2026/04/05 by Gemini: 已併入全域 _cancelRequested，不再單獨使用專屬旗標以簡化邏輯內容流程處理機制
    ' Private _cacheSnifferCts As New System.Threading.CancellationTokenSource  ' B4 CacheSniffer 取消令牌，FormClosing 時呼叫 Cancel()

    Private pnlOptions_tab3 As Panel
    Private _ctxListView1 As ContextMenuStrip
    Private rbExactMatch As New RadioButton()   ' tab5 用到的radio button
    Private rbFuzzyMatch As New RadioButton()   ' tab5 用到的radio button
    Private WithEvents ListView5 As New ListView()
    ' 可複選Treeview 自訂控制項 及 ContextMenu 成員變數，只初始化一次，不在每次右鍵時重新建立
    ' delete: 2026/04/01 by simon, 直接從設計工具建立 SimTree 控制項到 Form1
    'Private WithEvents SimTree1 As New SimTree
    'Private WithEvents SimTree2 As New SimTree
    'Private WithEvents SimTree3 As New SimTree
    'Private WithEvents SimTree4 As New SimTree

    ' [新增ProgressBar歷史紀錄 2026/4/2, by Gemini]

    Private Const MAX_HISTORY_COUNT As Integer = 100
    Private WithEvents HistoryListBox As ListBox
    Private _historyHoverIndex As Integer = -1
    Private _historyPopup As ToolStripDropDown
    Private _statusHistory As New List(Of StatusHistoryItem)()
    Private Structure StatusHistoryItem
        Dim Time As DateTime
        Dim Message As String
        Dim Source As String
    End Structure
#End Region

#Region "■ 02 Form 生命週期 & 外觀初始化"
#Region "  ├ 表單行為及輔助函數"
    Private Async Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
#If DEBUG Then
        _isDebugMode = True ' by Gemini, 2026/04/01: 自動依據 VS 的編譯組態判斷是否為 Debug 模式
#Else
        _isDebugMode = False 
#End If
        _dbg("開始") ' debugForm 開始計時
        Dim stopwatch As New Stopwatch() : stopwatch.Start() : Cursor = Cursors.AppStarting

        InitLookAndFeel()       ' 設計程式外觀
        InitProgressBarEvents() ' 2026/04/02 by Gemini: 集中掛載 ProgressBar 互動事件 (取代 Handles 宣告)
        Me.KeyPreview = True    ' ✅ 讓 Form 優先攔截 ESC，否則 ESC 會先被 TreeView/ListBox 等子控制項消耗
        Me.BringToFront()       ' 先將表單顯示後, 再以背景執行緒加入資料夾, 提高操作反應速度
        Me.Show()

        stopwatch.Stop() : Cursor = Cursors.Default ' 啟動完成, 停止計時, 顯示總共花費的時間
        ProgressBar1.Text = "啟動花費 " & stopwatch.Elapsed.TotalSeconds.ToString("0.00") & " 秒。"
        ProgressBar2.Text = ""
        _dbg("結束")

    End Sub
    Private Async Sub Form1_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        _dbg("開始")

        ' by Gemini, 2026/04/01: 利用背景躲藏時間，預先載入其他 Tab 的 UI 與目錄樹，實現「切換瞬間無感」的流暢體驗
        ' 讓第一頁先穩穩地顯示出來，不要與使用者剛啟動後的第一波對 TreeView1 的操作搶資源
        Await Task.Yield()

        InitOutlookNamespace()
        'InitRdoSession()
        InitDatabase()          ' by Gemini, 2026/04/06: 初始化 SQLite 快取資料庫

        ' 2024/5/17, PST檔太多, 啟動速度愈來愈差, 全部重寫. 依照20年前的做法動態載入:
        ' 啟動時只載入第一層表皮, 若下層有subFolders=True 則暫塞一個假的":::" 讓它能顯示"+"加號表示還有子資料夾就好
        ' 只有當使用者點開 "+" 號展開節點時, 才真正去讀該項目的子資料夾, 不要一開始就花時間全讀
        LoadStoreToTreeView(_pstStoreList, TreeView1)
        ExpandTreeToDefaultInbox(TreeView1)
        ' 2026/04/13 by Simon/Claude: SimTree1 與 TreeView1 同步載入 PST 目錄樹
        LoadStoreToTreeView(_pstStoreList, SimTree1)
        ExpandTreeToDefaultInbox(SimTree1)
        Await Task.Yield()

        If _isDebugMode Then    ' by Gemini, 2026/04/01: 如果是 debug mode，就顯示 debugForm跟 debug button
            CheckDebug.Visible = True
            ' ✅ 2026/03/30 by Gemini: 改用 BeginInvoke 延遲啟動，避免 Load 期間同步觸發事件造成 UI 卡頓或 Handle 競爭, 移除原本導致Exception 的Task.Run 呼叫
            Me.BeginInvoke(Sub() CheckDebug.Checked = True) ' Memo: 這裡設成True 就會預設開啟 DebugForm，False 不開啟，設計階段方便debug
        End If

        ' by Gemini, 2026/04/05: 將表單移動與縮放事件改為 AddHandler，保持類別簡潔
        AddHandler Me.Resize, Sub() SyncDebugFormPosition()
        AddHandler Me.Move, Sub() SyncDebugFormPosition()

        ' ── 初始化全域狀態變更事件 (by Gemini, 2026/04/10) ──
        AddHandler CheckSubFolder2.CheckedChanged, Sub() _includeSubTab2 = CheckSubFolder2.Checked
        AddHandler CheckSubFolder3.CheckedChanged, Sub() _includeSubTab3 = CheckSubFolder3.Checked
        AddHandler checkShowAllFolders.CheckedChanged, Sub() _showAllFolders = checkShowAllFolders.Checked
        AddHandler OKiLikeNoisy.CheckedChanged, Sub() _iLikeNoisy = OKiLikeNoisy.Checked        ' ✅ 2026/04/12 by simon: 改為動態掛載過濾噪音開關事件

        ' Tab1 順利載入後，才開始載入 Tab2~Tab5 的 UI 與資料，避免一開始就全部載入造成卡頓
        ' 使用 TryToRelaxFor 確保使用者正在操作時會暫緩預載
        ' 依序初始化後面的標籤頁，拉出間隔避免卡住使用者剛進入畫面的第一波操作
        Dim delaySame As Integer = 100  ' 每個 Tab 之間的預載延遲，單位毫秒 (ms)，可以根據需要調整
        'Dim delayDepends() As Integer = {500, 1000, 2000, 3000, 4000, 5000}

        ' by Gemini, 2026/04/03: 增加載入各 Tab 之間的視覺區隔
        Await TryToRelaxFor(delaySame)
        If Not _isTabInitialized(2) Then
            InitTab2UI() : _isTabInitialized(2) = True
            LoadStoreToTreeView(_pstStoreList, SimTree2) : ExpandTreeToDefaultInbox(SimTree2)
        End If

        Await TryToRelaxFor(delaySame)
        If Not _isTabInitialized(3) Then
            InitTab3UI() : _isTabInitialized(3) = True
            LoadStoreToTreeView(_pstStoreList, SimTree3) : ExpandTreeToDefaultInbox(SimTree3)
        End If

        Await TryToRelaxFor(delaySame)
        If Not _isTabInitialized(4) Then
            InitTab4UI() : _isTabInitialized(4) = True
            LoadStoreToTreeView(_pstStoreList, TreeView4) : ExpandTreeToDefaultInbox(TreeView4)
        End If

        Await TryToRelaxFor(delaySame)
        If Not _isTabInitialized(5) Then
            InitTab5UI() : _isTabInitialized(5) = True
            LoadStoreToTreeView(_pstStoreList, TreeView5) : ExpandTreeToDefaultInbox(TreeView5)
        End If

        _dbg("結束", "全部 Tab 背景載入完畢") ' by Gemini, 2026/04/11: UI 頂層 Level 0

        ' added by Gemini 3.1 Pro, 2026/04/12: 在所有的背景預載跟 Tab 初始化完成後，把一開始被蓋掉的啟動時間訊息重新顯示到 ProgressBar 上
        Dim firstMsgItem = _statusHistory.FirstOrDefault(Function(x) x.Source = "PB1")
        If Not String.IsNullOrEmpty(firstMsgItem.Message) Then ProgressBar1.Text = firstMsgItem.Message

    End Sub
    Private Async Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        ' 第一次進到這裡時, _isClosing = False, 跳過FinalRelease(), 執行下面的SaveCache()
        ' 確保下方的SaveCacheAsync()在Form關閉前完成, 然後再重新觸發FormClosing,
        ' designed by Gemini, 2026/4/8
        ' 如果已經進入正式關閉流程，則釋放資源並允許關閉
        If _isClosing Then
            '_cacheSnifferCts.Cancel()   ' ✅ 2026-03-16 B4: 通知 CacheSniffer 停止，避免程式關閉後 COM 呼叫繼續進行
            CloseDatabase()

            ' 釋放所有的 COM 物件占用資源
            If _pstStoreList IsNot Nothing Then
                For Each store In _pstStoreList : Marshal.FinalReleaseComObject(store) : Next
                _pstStoreList.Clear() : _pstStoreList = Nothing
            End If

            If _olApp IsNot Nothing Then Marshal.FinalReleaseComObject(_olApp)
            If _olNS IsNot Nothing Then Marshal.FinalReleaseComObject(_olNS)
            If _rdo IsNot Nothing Then Marshal.FinalReleaseComObject(_rdo)
            _dbg("結束")
            Return
        End If

        ' 第一次攔截關閉事件：取消預設關閉，暫停 Timer，執行非同步儲存後再手動關閉
        e.Cancel = True
        _isClosing = True
        timerSaveCache.Enabled = False

        ProgressBar1.Text = "正在儲存快取，準備關閉程式..."
        Await SaveCachesToSQLiteAsync()

        ' by Gemini, 2026/04/10: 關閉前顯式呼叫記憶體清理，確保資源釋放
        ClearMemoryCachesInternal()

        Me.Close() ' 觸發第二次進入此函式，執行上方釋放資源的區塊

    End Sub
    Private Sub Form1_ResizeEnd(sender As Object, e As EventArgs) Handles Me.ResizeEnd
        _dbg("結束", sender.Width & "x" & sender.Height)
        ' 視窗縮放時同步 DebugForm — 2026/3/26 by Gemini
        ' 原本的 ListView1 寬度調整邏輯已移至 ListViewResize 中，由 ListView 自行處理 Resize 事件
        ' Tab3 GroupBox3 顯示邏輯已改由 _pnlOptionsTab3.Resize 獨立處理，不再依賴 Form1_Resize
    End Sub
    Private Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        ' ── ESC 全域中斷 ──────────────────────────────────────────────
        ' KeyPreview=True 讓 Form 優先攔截 KeyDown，子控制項不會先吃掉 ESC
        If e.KeyCode = Keys.Escape Then
            ' 只有正在運算中才觸發中斷（透過 WaitCursor 判定）
            If Cursor = Cursors.WaitCursor OrElse ProgressBar1.Text.StartsWith("正在") Then
                _cancelRequested = True
                ' ✅ 發送標準取消信號
                If _cts IsNot Nothing AndAlso Not _cts.IsCancellationRequested Then _cts.Cancel()
                Cursor = Cursors.Default : ProgressBar1.Text = "已中斷。"
                e.Handled = True
                e.SuppressKeyPress = True ' ✅ 防止事件繼續傳遞觸發 ListView2_KeyPress 等回上一頁邏輯
            End If
            ' 非運算中: 完全不設 _cancelRequested，不呼叫 _cts.Cancel()，直接放行。
            ' 讓 ListView/TreeView 等子控制項的原生 KeyDown/KeyPress 自己處理 ESC (例如 ListView2 返回年份視圖)。
            ' 2026/04/11 by Claude: 修正舊版非運算中按 ESC 仍呼叫 _cts.Cancel() 的 bug，會汙染 token，導致下一個操作取得的 cToken 已經是 cancelled 狀態。
        End If

    End Sub
#End Region
#Region "  ├ 物件及外觀初始化"
    Private Sub InitLookAndFeel()
        ' === 初始化共用物件的外觀及共通行為 ===
        _dbg("開始")
        ' 2026-03-17 拆分: TreeView / ListView 各司其職的外觀設定移到獨立函數
        '   InitLookAndFeel()   ← 視窗位置、TabControl、ContextMenu、Chart2、Button、雜項
        '   InitTreeview()  ← TreeView / SimTree 字型、顏色、雙緩衝
        '   InitListview()  ← ListView 字型、基本樣式、雙緩衝、欄位定義

        ' 設定程式標題
        Dim strApp As String = My.Application.Info.DirectoryPath & "\" & My.Application.Info.ProductName & ".EXE"
        If My.Computer.FileSystem.FileExists(strApp) Then
            Dim infoReader As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(strApp)
            Dim modeStr As String = If(_isDebugMode, "(Debug)", "(Release)")
            Me.Text = $"Outlook Assistant - by Simon Lee Studio (build {infoReader.LastWriteTime:yyyy/MM/dd HH:mm:ss}) {modeStr}"
            '' todo: 如何設置版本號自動遞增 'myApp.MinorRevision += 1
        End If

        ' ── 視窗位置與背景色 ──
        If Screen.FromControl(Me).Bounds.Height > 2560 Then
            Me.Top = Screen.FromControl(Me).Bounds.Height * 0.45                '如果在直立式的4K螢幕上啟動, 就把表單放在下半部往上移5%
            Me.Left = (Screen.FromControl(Me).Bounds.Width - Me.Width) * 0.45   '不管在什麼解析度的螢幕上啟動, 都把表單放在螢幕中央往左移5%
        End If
        Me.BackColor = ThemeColors.Gray95

        ' ── TabControl 字型與分頁名稱 ──
        Dim strTabName As String() = {"資料夾統計", "依日期統計", "尋找附件", "尋找系列郵件", "尋找重覆郵件", "Setting"}
        For i As Integer = 0 To strTabName.Length - 1
            TabControl1.TabPages(i).Text = strTabName(i)
        Next
        TabControl1.Font = New Font(_fontDefault, _fontBold)
        TabControl1.Padding = New Point(12, 8)
        txtDatabaseStats.Font = New Font(_fontDefault, _fontRegular)

        ' ── 容器化佈局與動態控制項掛載 ──
        ' by Gemini, 2026/04/01: 只初始化 Tab1，其餘 Tab 在切換時才載入 (Lazy Load)
        InitTab1UI()

        ' 2026/3/27 by Gemini: 修復 StatusStrip1 被 TabControl1 遮擋的問題
        StatusStrip1.SendToBack()
        TabControl1.BringToFront()

        DebugGroup.Visible = _isDebugMode   ' 只在 Debug 模式才顯示 DebugGroup 及相關控制項
        _dbg("結束")

    End Sub
    Private Sub InitTreeView(tv As TreeView)
        ' ---------------------------------------------------------------------------------------------------------
        ' ── 共用 Treeview 外觀設定 (by Gemini, 2026/04/01: 重構成接受單一參數，避免重複造輪子) ──
        ' ---------------------------------------------------------------------------------------------------------
        tv.Font = New Font(_fontDefault, _fontRegular)
        tv.BackColor = Color.White
        tv.ForeColor = SystemColors.InactiveCaptionText
        tv.Dock = DockStyle.Fill
        tv.Indent = 25          ' 樹狀目錄縮排距離

        ' 雙重緩衝區優化
        SendMessage(tv.Handle, TVM_SETEXTENDEDSTYLE, New IntPtr(TVS_EX_DOUBLEBUFFER), New IntPtr(TVS_EX_DOUBLEBUFFER))

        AddHandler tv.BeforeExpand, AddressOf LoadSubFolderToTreeView
        AddHandler tv.KeyPress, AddressOf HandleTreeViewKeyPress
        AddHandler tv.MouseMove, AddressOf HandleTreeViewMouseHover
        AddHandler tv.MouseLeave, AddressOf HandleTreeViewMouseHover

    End Sub
    Private Sub InitListView(lv As ListView)
        ' ---------------------------------------------------------------------------------------------------------
        ' ── 共用 Listview 外觀設定 (by Gemini, 2026/04/01: 重構成接受單一參數，避免重複造輪子) ──
        ' ---------------------------------------------------------------------------------------------------------
        lv.Font = New Font(_fontDefault, _fontRegular)
        lv.GridLines = False
        lv.View = System.Windows.Forms.View.Details
        lv.FullRowSelect = True
        lv.BringToFront()
        lv.Anchor = AnchorStyles.None
        lv.Dock = DockStyle.Fill
        lv.Cursor = Cursors.Default

        ' 雙重緩衝區優化
        SendMessage(lv.Handle, LVM_SETEXTENDEDLISTVIEWSTYLE, New IntPtr(LVS_EX_DOUBLEBUFFER), New IntPtr(LVS_EX_DOUBLEBUFFER))

        ' by Gemini, 2026/04/10: 使用反射開啟隱藏的 DoubleBuffered 屬性，解決虛擬模式下的重繪閃爍問題
        GetType(ListView).GetProperty("DoubleBuffered", System.Reflection.BindingFlags.Instance Or System.Reflection.BindingFlags.NonPublic).SetValue(lv, True, Nothing)

        AddHandler lv.MouseMove, AddressOf HandleListViewMouseHover
        AddHandler lv.MouseLeave, AddressOf HandleListViewMouseHover
        AddHandler lv.GotFocus, AddressOf ListViewGotFocus
        AddHandler lv.KeyPress, AddressOf HandleListViewKeyPress
        AddHandler lv.Resize, AddressOf ListViewResize  ' 2026/04/01 by Gemini: 加入共用自動縮放事件

    End Sub
    Private Sub InitSplitContainer(scnr As SplitContainer)
        scnr.Panel1MinSize = 0

        AddHandler scnr.MouseMove, Sub(s, ev) DirectCast(s, SplitContainer).Cursor = Cursors.SizeWE
        AddHandler scnr.MouseLeave, Sub(s, ev) DirectCast(s, SplitContainer).Cursor = Cursors.Default
        AddHandler scnr.MouseDown, AddressOf SplitContainerMouseDown

    End Sub
    Private Sub InitProgressBarEvents()
        ' ── ProgressBar 歷史紀錄 (by Gemini, 2026/04/02) ──
        ''' <summary>
        ''' 集中初始化 ProgressBar1 與 ProgressBar2 的互動事件 (TextChanged, Click, Hover)
        ''' 2026/04/02 by Gemini
        ''' </summary>
        _dbg("開始")

        ' 1. 文字變更紀錄
        AddHandler ProgressBar1.TextChanged, Sub() AppendStatusHistory(ProgressBar1.Text, "PB1")
        AddHandler ProgressBar2.TextChanged, Sub() AppendStatusHistory(ProgressBar2.Text, "PB2")

        ' 2. 點擊彈出歷史選單
        AddHandler ProgressBar1.Click, Sub() ShowHistoryPopup("PB1", ProgressBar1)
        AddHandler ProgressBar2.Click, Sub() ShowHistoryPopup("PB2", ProgressBar2)

        ' 3. 滑鼠進入/離開視覺效果 (使用共用處理邏輯)
        Dim hoverIn = Sub(s, ev)
                          Dim lbl = DirectCast(s, ToolStripStatusLabel)
                          lbl.BackColor = ThemeColors.MercuryGray
                      End Sub
        Dim hoverOut = Sub(s, ev)
                           Dim lbl = DirectCast(s, ToolStripStatusLabel)
                           lbl.BackColor = Color.Transparent
                       End Sub

        ' 註：ProgressBar1/2 實際上是 ToolStripStatusLabel
        AddHandler ProgressBar1.MouseEnter, hoverIn
        AddHandler ProgressBar2.MouseEnter, hoverIn
        AddHandler ProgressBar1.MouseLeave, hoverOut
        AddHandler ProgressBar2.MouseLeave, hoverOut
        _dbg("結束")

    End Sub

    Private Sub InitTab1UI()
        _dbg("開始")

        ' 2026/04/13 by Simon/Claude: SimTree1 由設計工具建立，位置與層級已在 Designer.vb 設定
        ' 不需要 Controls.Add / BringToFront，直接 InitTreeView 初始化樣式即可
        InitTreeView(SimTree1)
        InitListView(ListView1)
        InitSplitContainer(SplitContainer1)

        ' 2026/04/13 v2: 移除「所屬父資料夾」欄，回歸 5 欄（該欄內容永遠等於群組標題行，元余）
        ' 欄位順序: 資料夾名稱 / 郵件數量 / 資料夾數量 / 郵件總計 / 資料夾大小
        ListView1.Columns.Clear()
        Dim headerNames As String() = {"資料夾名稱", "郵件數量", "資料夾數量", "郵件總計", "資料夾大小"}
        For Each n In headerNames
            ListView1.Columns.Add(n, n)
            ListView1.Columns(n).Width = ListView1.Width * 0.16
        Next
        ListView1.Columns("資料夾名稱").Width = ListView1.Width * 0.32
        ListView1.Columns("資料夾名稱").TextAlign = HorizontalAlignment.Left
        ListView1.Columns("資料夾大小").Width = ListView1.Width * 0.188
        For Each c In ListView1.Columns
            c.TextAlign = HorizontalAlignment.Right
        Next
        ListView1.Columns("資料夾名稱").TextAlign = HorizontalAlignment.Left

        ' 2026/04/13 by Simon/Claude: OwnerDraw=True 讓群組標題行 / 合計列的 BackColor
        ' 在 hover / select 狀態下不被 OS 覆蓋（DrawSubItem handler 在 Form1_MainTabs.vb）
        ListView1.OwnerDraw = True

        _ctxListView1 = New ContextMenuStrip()
        _ctxListView1.Items.Add("進入資料夾 (&E)", Nothing, Sub(sender, e) EnterSelectedFolder(ListView1.SelectedItems(0)))
        _ctxListView1.Items.Add("統計資料夾大小 (&C)", Nothing, AddressOf ComputeFolderSize)
        _isTabInitialized(1) = True
        _dbg("結束")

    End Sub
    Private Sub InitTab2UI()
        ' ── Tab2 (日期統計) 佈局重構 (2026/3/27 by Gemini) ──
        _dbg("開始")

        InitTreeView(SimTree2)
        InitListView(ListView2)
        InitSplitContainer(SplitContainer2)
        InitChart2()

        ' 使用者建議 CheckSubFolder2 應在列表下方、SimTree2 右方
        ' 我們在 Panel2 內建立一個中間層面板
        Dim pnlCheckbox_tab2 As Panel
        pnlCheckbox_tab2 = New Panel With {.Dock = DockStyle.Top,
                                           .Height = 35,
                                           .BackColor = ThemeColors.Gray95}
        SplitContainer2.Panel2.Controls.Add(pnlCheckbox_tab2)

        pnlCheckbox_tab2.Controls.Add(CheckSubFolder2)
        CheckSubFolder2.Location = New Point(10, 8)

        ' 確保所有組件都在右側容器 (Panel2) 內
        ListView2.Parent = SplitContainer2.Panel2
        Chart2.Parent = SplitContainer2.Panel2
        ListView2.Height = 450
        ListView2.Dock = DockStyle.Top
        Chart2.Dock = DockStyle.Fill

        ' 💡 關鍵 Dock 順序 (by Gemini, 2026/03/28 修正 Z-order 以符合預期佈局)：
        ' 在 WinForms 中，Z-order 在最底層 (SendToBack) 的控制項會最先進行 Dock (搶佔邊緣)。
        Chart2.BringToFront()           ' 3. 圖表移至最前方 (最後才進行 Dock=Fill，填滿剩餘空間)
        pnlCheckbox_tab2.SendToBack()   ' 2. 面板移至最後方
        ListView2.SendToBack()          ' 1. 列表最後被移至最後方 (最優先 Dock=Top，確保在最上面)

        Chart2.BorderlineDashStyle = ChartDashStyle.Solid
        Chart2.BorderlineColor = ThemeColors.AltoGray
        Chart2.ChartAreas(0).BackColor = ThemeColors.bgColor
        Chart2.ChartAreas(0).AxisX.MajorGrid.LineColor = ThemeColors.gridLine
        Chart2.ChartAreas(0).AxisY.MajorGrid.LineColor = ThemeColors.gridLine
        Chart2.ChartAreas(0).Position = New ElementPosition(1, 1, 99, 99)   ' ── 最大化 ChartArea 和 InnerPlotPosition ──
        ' ChartArea.Position: ChartArea 在整個 Chart 控制項中的佔比 (單位: %)

        ' ✅ 讓 ChartArea 幾乎填滿整個 Chart 控制項 (上下左右各留 1%) ' 預設約 Position(5,5,90,90)，壓縮到幾乎填滿整個 Chart 控制項
        ' ✅ InnerPlotPosition: ChartArea 內部長條圖實際繪製區的佔比
        '    Auto=True 時 Chart 會自動縮排給軸標籤留空，通常左側縮 10~15%
        '    改成 Auto=False 並手動指定，讓左側縮排符合實際 Y 軸標籤寬度
        With Chart2.ChartAreas(0).InnerPlotPosition
            .Auto = False
            .X = 8          ' 左側留 8% (給 Y 軸數字標籤)
            .Y = 2          ' 上方留 2%
            .Width = 90     ' 往右延伸 90%
            .Height = 90    ' 往下延伸 90% (底部留 10% 給 X 軸標籤)
            ' InitChart2 會清除 Legends()，如果為空就不應該去存取，避免引發 ArgumentOutOfRangeException
            If Chart2.Legends.Count > 0 Then Chart2.Legends(0).Enabled = False
        End With
        _dbg("結束")

    End Sub
    Private Sub InitTab3UI()
        ' ── Tab3 (尋找附件) 佈局優化 ──
        _dbg("開始")
        InitTreeView(SimTree3)
        InitListView(ListView3)
        InitSplitContainer(SplitContainer3)
        ' 建立頂部面板，將所有原本散落在 Panel2 的搜尋控制項集中
        'Dim pnlOptions_tab3 As Panel
        pnlOptions_tab3 = New Panel With {.Dock = DockStyle.Top,
                                          .Height = 115,
                                          .BackColor = ThemeColors.Gray95,
                                          .Font = New Font(_fontDefault, _fontRegular)}

        ' 將原本 Panel2 中的控制項移入新增的 pnlOptions_tab3
        ' 這些控制項原本的 Location 已經適合在 Top Panel 中運作
        pnlOptions_tab3.SendToBack() ' 2026/3/27 by Gemini: 正確設定 Dock 計算順序
        pnlOptions_tab3.Controls.Add(GroupBox1)
        pnlOptions_tab3.Controls.Add(GroupBox2)
        pnlOptions_tab3.Controls.Add(GroupBox3)
        pnlOptions_tab3.Controls.Add(Button3)
        pnlOptions_tab3.Controls.Add(CheckSubFolder3)
        SplitContainer3.Panel2.Controls.Add(pnlOptions_tab3)

        ' 2026/04/05 by Gemini: 優化顯示邏輯「純淨版」
        ' 改用面板自身的 Resize 事件與 Lambda 運算，不需類別變數。
        ' 這樣無論是調整視窗還是隱藏側邊欄，GroupBox3 都會依據「右側實際可用空間 (820px)」決定顯現與否。
        AddHandler pnlOptions_tab3.Resize, Sub() GroupBox3.Visible = pnlOptions_tab3.Width >= 820
        GroupBox3.Visible = pnlOptions_tab3.Width >= 820

        ' ── Button3 樣式 ──
        Button3.FlatStyle = FlatStyle.System
        Button3.FlatAppearance.BorderColor = ThemeColors.Brand_Blue
        Button3.FlatAppearance.MouseOverBackColor = ThemeColors.AltoGray
        Button3.ForeColor = ThemeColors.Brand_Blue
        Button3.BringToFront()
        CheckSubFolder3.BringToFront()

        ' ── 2026/03/28 by Gemini: 對齊邏輯優化 ──
        CheckSubFolder3.CheckAlign = ContentAlignment.MiddleLeft
        CheckSubFolder3.TextAlign = ContentAlignment.MiddleLeft
        CheckSubFolder3.AutoSize = True                                                 ' 1. 開啟 AutoSize 解決勾選框與文字「離得太遠」的問題 (寬度會自動縮短到剛好)
        CheckSubFolder3.Anchor = AnchorStyles.Top Or AnchorStyles.Left                  ' 2. 清除 Anchor 避免設計時的自動定位干擾手動計算，之後再重設為右側關聯
        'CheckSubFolder3.Left = (Button3.Left + Button3.Width) - CheckSubFolder3.Width   ' 3. 重新計算右側對齊 (會在 AutoSize 完後的正確 Width 基礎上計算)
        CheckSubFolder3.Anchor = AnchorStyles.Top Or AnchorStyles.Right                 ' 4. 最後設定 Anchor，讓它在之後的視窗縮放中保持與 Button3 的右側對齊

        ' ------------------------------------------
        ' ── 2026/03/28 by Gemini: 集中掛載 Tab3 專屬互動邏輯 (Lambda 重構) ──
        AddHandler TextBox3.KeyDown, Sub(s, ev)
                                         If ev.KeyCode = Keys.Enter Then
                                             Button3.PerformClick() : TextBox3.SelectAll()
                                             ev.SuppressKeyPress = True
                                         End If
                                     End Sub
        AddHandler CheckAttachName.CheckedChanged, Sub()
                                                       TextBox3.Enabled = CheckAttachName.Checked
                                                       If CheckAttachName.Checked Then
                                                           TextBox3.Focus() : TextBox3.SelectAll()
                                                       End If
                                                   End Sub
        AddHandler CheckAttCount.CheckedChanged, Sub()
                                                     CountMin.Enabled = CheckAttCount.Checked
                                                     CountMax.Enabled = CheckAttCount.Checked
                                                     AutoResizeListViewColumns(ListView3) ' 加入自動縮放，依勾選狀態動態隱藏/顯示欄位
                                                 End Sub

        ' 2026/04/05 by Gemini: 優化數值微調邏輯，根據單位 (KB/MB/GB) 與當前數值動態調整增幅
        ' 並加入長按加速 (Accelerations) 提升大範圍調整效率
        For Each num In {NumberMin, NumberMax}
            num.Accelerations.Clear()
            num.Accelerations.Add(New NumericUpDownAcceleration(2, 5))  ' 2 秒後加速 5 倍
            num.Accelerations.Add(New NumericUpDownAcceleration(5, 50)) ' 5 秒後極速
        Next
        AddHandler NumberMin.ValueChanged, Sub() UpdateNumericIncrement(NumberMin, UnitMin)
        AddHandler NumberMax.ValueChanged, Sub() UpdateNumericIncrement(NumberMax, UnitMax)

        ' by Gemini, 2026/04/08: 為數字輸入框增加 Enter 鍵觸發搜尋功能
        AddHandler NumberMin.KeyDown, Sub(s, ev) If ev.KeyCode = Keys.Enter Then Button3.PerformClick() : ev.SuppressKeyPress = True
        AddHandler NumberMax.KeyDown, Sub(s, ev) If ev.KeyCode = Keys.Enter Then Button3.PerformClick() : ev.SuppressKeyPress = True

        AddHandler UnitMin.SelectedIndexChanged, Sub() UpdateNumericIncrement(NumberMin, UnitMin)
        AddHandler UnitMax.SelectedIndexChanged, Sub() UpdateNumericIncrement(NumberMax, UnitMax)

        ' 初始化時先執行一次以同步正確增額
        UpdateNumericIncrement(NumberMin, UnitMin)
        UpdateNumericIncrement(NumberMax, UnitMax)

        _dbg("結束")

    End Sub
    Private Sub InitTab4UI()
        ' ── Tab4 (系列郵件) 佈局優化 ──
        _dbg("開始")

        InitTreeView(TreeView4)
        InitListView(ListView4)
        InitSplitContainer(SplitContainer4)

        Dim pnlOptions_tab4 As Panel
        pnlOptions_tab4 = New Panel With {.Dock = DockStyle.Top,
                                          .Height = 45,
                                          .BackColor = ThemeColors.Gray95}

        ' 將按鈕移入面板
        Button4.Location = New Point(10, 10) ' 稍微修正位置使其在面板內美觀
        pnlOptions_tab4.Controls.Add(Button4)
        SplitContainer4.Panel2.Controls.Add(pnlOptions_tab4)

        pnlOptions_tab4.SendToBack()        ' 2026/3/27 by Gemini: 正確設定 Dock 計算順序

        ' ── ListView4: 系列郵件欄位定義 ──
        With ListView4
            .Columns.Clear()
            Dim lv4Names As String() = {"主旨", "大小", "收到時間", "寄件者", "EntryID"}
            For Each n In lv4Names
                .Columns.Add(n, n)
            Next
            .Columns("主旨").Width = .Width * 0.4
            .Columns("大小").Width = .Width * 0.15
            .Columns("大小").TextAlign = HorizontalAlignment.Right
            .Columns("收到時間").Width = .Width * 0.2
            .Columns("寄件者").Width = .Width * 0.2
            .Columns("EntryID").Width = 0 ' 隱藏欄位
        End With
        _dbg("結束")

    End Sub
    Private Sub InitTab5UI()
        ' 清除原有的測試控制項 (如果有) ，並移出 ListView5 到 TabPage5 下
        ' ── Tab5 選項面板 (為了支持穩定佈局，將頂部按鈕放入獨立 Panel) ──
        _dbg("開始")

        TreeView5.Visible = True
        InitTreeView(TreeView5)
        InitListView(ListView5)
        InitSplitContainer(SplitContainer5)

        Dim pnlOptions5 As Panel
        pnlOptions5 = New Panel With {.Dock = DockStyle.Top,
                                      .Height = 55,
                                      .BackColor = TabPage5.BackColor}

        rbExactMatch.Text = "完全相同 (主旨+大小+時間+寄件者)"
        rbExactMatch.Location = New Point(20, 18)
        rbExactMatch.Checked = True
        rbExactMatch.AutoSize = True
        rbFuzzyMatch.Text = "相似重複 (相似主旨+大小)"
        rbFuzzyMatch.Location = New Point(340, 18)
        rbFuzzyMatch.AutoSize = True
        Button5.Location = New Point(600, 13)
        Button5.Text = "開始掃描"
        Button5.AutoSize = True

        ' ── 組裝 UI ──
        pnlOptions5.Controls.Add(rbExactMatch)
        pnlOptions5.Controls.Add(rbFuzzyMatch)
        pnlOptions5.Controls.Add(Button5)
        TabPage5.Controls.Add(ListView5)    ' 先加 ListView，讓 Dock = Fill 佔滿底部
        TabPage5.Controls.Add(pnlOptions5)  ' 後加 Panel，Dock = Top 會佔據上方

        ' 2026/3/27 by Gemini: 正確設定 Dock 計算順序
        pnlOptions5.SendToBack()
        ListView5.BringToFront()
        _dbg("結束")

    End Sub
    Private Sub InitChart2()
        _dbg("開始", Chart2.Name)

        With Chart2
            ' 清除原有的設定
            .Series.Clear()
            .Legends.Clear()
            .ChartAreas.Clear()

            ' 設置抗鋸齒和文本抗鋸齒品質
            .AntiAliasing = AntiAliasingStyles.All
            .TextAntiAliasingQuality = TextAntiAliasingQuality.High

            ' 添加 Chart 的 Series
            Dim mailCount As New Series With {.Name = "郵件數量",
                                              .ChartType = SeriesChartType.Column,
                                              .Color = ThemeColors.barNormal}
            ' 添加 Chart 的 ChartArea
            Dim mailChart As New ChartArea With {.Name = "長條圖",
                                                 .BackColor = ThemeColors.Gray95,
                                                 .BorderColor = Color.DarkGray}
            With mailChart
                ' 設置背景格線顏色和寬度
                .AxisX.LineColor = Color.DimGray
                .AxisY.LineColor = Color.DimGray
                .AxisX.MajorGrid.LineColor = ThemeColors.gridLine ' 淡灰色
                .AxisY.MajorGrid.LineColor = ThemeColors.gridLine ' 淡灰色
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

            .Series.Add(mailCount)
            .ChartAreas.Add(mailChart)
        End With
        _dbg("結束")

    End Sub
#End Region
#Region "  └ 輔助函數"
    ''' <summary>
    ''' 標準化取消訊號重置流程：中斷舊任務、釋放資源並產生新的 Token
    ''' 取消時以下三選一:
    ''' 彈射座椅 --> 開門下車 --> 開門帶行李走
    '''     Await Task.Yield() : cToken.ThrowIfCancellationRequested() (這個實測幾乎無效!!)
    '''     Await Task.Delay(1, cToken) ' 這裡一定要保留至少 .delay(1) 才能讓 ESC 中斷生效 (simon, 2026/04/05)
    '''     If cToken.IsCancellationRequested Then Return New List(Of FolderBfsEntry) 
    ''' </summary>
    Private Function OkayNowYouHaveToken() As CancellationToken
        If _cts IsNot Nothing Then
            Try : _cts.Cancel() : _cts.Dispose() : Catch : End Try
        End If
        _cancelRequested = False  ' by Claude Opus, 2026/04/11: 重置舊旗標，否則 Layer3 的 GetMailCountAll / GetFolderCountAll 等仍在檢查它，ESC 後就永遠回傳 0 (全換token之後就可以不檢查這個旗標了，但為了保險起見還是重置它)
        _cts = New CancellationTokenSource()
        Return _cts.Token
    End Function
    Private Async Function TryToRelaxFor(baseDelayMs As Integer) As Task
        ''' <summary>
        ''' 智慧等待輔助函式
        ''' 先睡眠預定時間，若使用者正在忙碌(例如正在 AfterSelect 統計中)，則每 1000ms 檢查一次直到閒置。
        ''' 2026/04/01 by Gemini
        ''' </summary>

        Await Task.Delay(baseDelayMs)   ' 1. 先執行基礎延遲
        While _isUserBusy               ' 2. 醒來後檢查旗標，若忙碌則循環等待
            _dbg("    ├ 使用者忙碌中，背景預載暫緩 1000ms...") ' by Gemini, 2026/04/11: 內部細節 Level 2
            Await Task.Delay(1000)
        End While

    End Function
    Private Sub SyncDebugFormPosition()
        ''' <summary>
        ''' 同步 Debug 視窗與主視窗的位置與大小，並將其右側貼齊螢幕邊緣
        ''' 使用 SetWindowPos 避免多個屬性分別設定導致的閃爍
        ''' 2026/3/26 by Gemini
        ''' </summary>

        If DebugForm IsNot Nothing AndAlso (DebugForm.Visible OrElse CheckDebug.Checked) Then
            Dim newLeft As Integer = Me.Left + Me.Width - 12
            Dim newTop As Integer = Me.Top
            Dim newHeight As Integer = Me.Height

            ' 計算螢幕工作區右側邊緣，並延展 DebugForm 寬度填滿剩餘空間
            Dim screenRight = Screen.FromControl(Me).WorkingArea.Right
            Dim newWidth = screenRight - newLeft
            If newWidth < 100 Then newWidth = DebugForm.Width ' 保底寬度

            ' 2026/3/28 by Gemini: 簡化重繪策略 — 不干預 Windows 的原生重繪機制，
            ' 讓 SetWindowPos 自然觸發 WM_PAINT，確保佈局即時生效 (供 DebugForm_Load 的 Delay 計時用)
            SetWindowPos(DebugForm.Handle, IntPtr.Zero, newLeft, newTop, newWidth, newHeight, SWP_NOZORDER Or SWP_NOACTIVATE)
        End If

    End Sub
    Private Function SafeGet(Of T)(row As Outlook.Row, column As String, defaultValue As T) As T
        ''' <summary>
        ''' 安全地從 Outlook.Row 讀取欄位，自動處理 Nothing / DBNull / 例外
        ''' 2026/04/01 by Gemini
        ''' </summary>
        Try
            Dim value = row(column)
            If value Is Nothing OrElse IsDBNull(value) Then Return defaultValue
            Return CType(value, T)
        Catch ex As System.Exception
            _dbg("       ├ SafeGet(Row) 失敗", $"{column} | {ex.Message}") ' by Gemini, 2026/04/11: 底層詳細 Level 3
            Return defaultValue
        End Try

    End Function
    Private Function SafeGet(Of T)(data(,) As Object, row As Integer, col As Integer, defaultValue As T) As T
        ''' <summary>
        ''' SafeGet 的二維陣列（GetArray）Overload 版
        ''' 2026/04/01 by Gemini
        ''' </summary>
        Try
            Dim value = data(row, col)
            If value Is Nothing OrElse IsDBNull(value) Then Return defaultValue
            ' 使用 Convert.ChangeType 確保數值型態（如 Long/Int/DateTime）能正確轉換
            Return CType(Convert.ChangeType(value, GetType(T)), T)
        Catch
            Return defaultValue
        End Try

    End Function
#End Region
#End Region

#Region "■ 03 共用控制項行為"
#Region "  ├ 共用 UI控制項"
    Private Sub TabControl1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl1.SelectedIndexChanged
        _isUserBusy = True
        _dbg("開始")

        Try ' by Gemini, 2026/04/01: 根據選定的分頁動態載入 UI 與資料 (Lazy Load UI)
            ProgressBar1.Text = "" : ProgressBar2.Text = ""
            Dim selectedTab As TabPage = CType(sender, TabControl).SelectedTab
            Dim tabIndex As Integer = TabControl1.SelectedIndex + 1 ' 產生 1, 2, 3, 4, 5 (Tab1~5)

            ' 如果切換到 Debug 分頁 (TabIndex >= 6)，不需要執行底下的陣列檢查，不需初始化 UI 直接離開
            If tabIndex > 5 Then
                If selectedTab.Text = "Setting" Then
                    ' 保留給 Debug 分頁的特別處理 (如果有)
                    ' 保留給 Debug 分頁的特別處理 (如果有)
                    ' 保留給 Debug 分頁的特別處理 (如果有)
                    ' 保留給 Debug 分頁的特別處理 (如果有)
                End If
                Return
            End If

            ' ── 步驟 1: 即時建構目前分頁專屬 UI ──
            If Not _isTabInitialized(tabIndex) Then
                selectedTab.SuspendLayout()
                Select Case tabIndex
                    Case 1 : InitTab1UI()
                    Case 2 : InitTab2UI()
                    Case 3 : InitTab3UI()
                    Case 4 : InitTab4UI()
                    Case 5 : InitTab5UI()
                End Select
                selectedTab.ResumeLayout()
                _isTabInitialized(tabIndex) = True
            End If

            ' ── 步驟 2: 依照不同的頁面載入不同的treeview，並展開到預設的收件匣位置 ──
            Dim currentTree As TreeView = GetActiveTreeView()
            If currentTree IsNot Nothing Then
                If currentTree.Nodes.Count = 0 Then
                    LoadStoreToTreeView(_pstStoreList, currentTree)
                    ExpandTreeToDefaultInbox(currentTree)
                End If
                currentTree.Focus()
            End If
            _dbg("結束")
        Finally
            _isUserBusy = False
        End Try

    End Sub
    Private Sub CheckShowAllFolders_CheckedChanged(sender As Object, e As EventArgs) Handles checkShowAllFolders.CheckedChanged
        ' by Gemini, 2026/03/30: 當切換顯示所有資料夾時，清空快取並標記所有 TreeView 為無效 (Nodes.Clear)
        ' 分頁在切換時，由 SelectedIndexChanged 自動按新過濾條件重新載入, 不需要在這裡重複載入，避免不必要的 COM 呼叫和 UI 重繪
        _cacheFolderTree.Clear()
        _dbg("已切換顯示所有資料夾 (_cacheFolderTree 快取已清空)")

        For Each tv In GetAllTreeViews(Me)
            tv.Nodes.Clear()
        Next
        ProgressBar2.Text = "全域資料夾過濾已變更，各頁面將於切換時自動重新整理。"

    End Sub
    Private Sub CheckRDO_CheckedChanged(sender As Object, e As EventArgs) Handles CheckRDO.CheckedChanged
        ' 用一個checkbox 動態決定是否載入Redemption
        If CheckRDO.Checked Then
            ' 已知限制: 卸載後就無法再重新載入第二次, 不會成功
            If _rdo Is Nothing Then Dim unused = InitRedemptionSessionWithoutDeclaration()
        Else
            TryMarshalRelease(_rdo)
        End If

    End Sub
    Private Sub ClearMemoryCachesInternal()
        ' ---------------------------------------------------------------
        ' ClearMemoryCachesInternal — [記憶體層] 統一清理所有 ConcurrentDictionary
        ' ---------------------------------------------------------------
        _cacheMailCount.Clear()
        _cacheMailCountAll.Clear()
        _cacheFolderCount.Clear()
        _cacheFolderCountAll.Clear()
        _cacheFolderSize.Clear()
        _cacheFolderSizeAll.Clear()

        _cacheYearCounts.Clear()
        _cacheMonthCounts.Clear()
        _cacheAttachMailList.Clear()
        _cacheAttachFilename.Clear()

        _cacheFolderTree.Clear()
        _cacheSubFolderList.Clear()
        _cacheIsMailFolder.Clear()
        _cacheFolderIDs.Clear() ' 2026/04/10 新增 ID 快取清理

    End Sub
    Private Sub ListViewGotFocus(sender As Object, e As EventArgs)
        ' 2026/03/28 by Gemini: 集中處理 ListView 獲得焦點時自動選取第一項的邏輯
        Dim lv = DirectCast(sender, ListView)

        ' by Gemini, 2026/04/10: 虛擬模式下存取 SelectedItems 會拋出 InvalidOperationException
        ' 必須改用 SelectedIndices.Count 判斷
        If lv.VirtualMode Then
            If lv.SelectedIndices.Count = 0 AndAlso lv.VirtualListSize > 0 Then
                lv.SelectedIndices.Add(0)
            End If
        Else
            If lv.SelectedItems.Count = 0 AndAlso lv.Items.Count > 0 Then
                lv.Items(0).Selected = True
            End If
        End If
    End Sub
    Private Sub ListViewResize(sender As Object, e As EventArgs)
        ''' <summary>
        ''' 處理所有 ListView 的 Resize 共用事件 (2026/04/01 by Gemini)
        ''' </summary>
        Dim lv As ListView = TryCast(sender, ListView)
        If lv IsNot Nothing Then AutoResizeListViewColumns(lv)

    End Sub

    Private Sub HistoryListBox_MouseMove(sender As Object, e As MouseEventArgs) Handles HistoryListBox.MouseMove
        Dim newHoverIndex = HistoryListBox.IndexFromPoint(e.Location)
        If _historyHoverIndex <> newHoverIndex Then
            ' 只重繪前一個和新碰觸的項目，而不是整個 ListBox，大幅改善 Hover 和捲動的效能卡頓
            Dim oldIndex = _historyHoverIndex
            _historyHoverIndex = newHoverIndex

            If oldIndex <> -1 AndAlso
                oldIndex < HistoryListBox.Items.Count Then HistoryListBox.Invalidate(HistoryListBox.GetItemRectangle(oldIndex))
            If _historyHoverIndex <> -1 AndAlso
                _historyHoverIndex < HistoryListBox.Items.Count Then HistoryListBox.Invalidate(HistoryListBox.GetItemRectangle(_historyHoverIndex))
        End If

    End Sub
    Private Sub HistoryListBox_MouseLeave(sender As Object, e As EventArgs) Handles HistoryListBox.MouseLeave
        If _historyHoverIndex <> -1 Then
            Dim oldIndex = _historyHoverIndex
            _historyHoverIndex = -1
            If oldIndex < HistoryListBox.Items.Count Then HistoryListBox.Invalidate(HistoryListBox.GetItemRectangle(oldIndex))
        End If

        ' 當滑鼠真正離開 Popup 範圍時再自動關閉 Popup
        Dim pt = System.Windows.Forms.Cursor.Position
        If _historyPopup IsNot Nothing AndAlso _historyPopup.Visible Then
            If Not _historyPopup.Bounds.Contains(pt) Then _historyPopup.Close()
        End If

    End Sub
    Private Sub HistoryListBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles HistoryListBox.SelectedIndexChanged
        If HistoryListBox.SelectedIndex >= 0 Then
            Dim selectedText = HistoryListBox.SelectedItem.ToString()
            Try
                Clipboard.SetText(selectedText)
            Catch
            End Try
        End If

    End Sub
    Private Sub HistoryListBox_DrawItem(sender As Object, e As DrawItemEventArgs) Handles HistoryListBox.DrawItem
        If e.Index < 0 Then Return
        Dim isHovered = (e.Index = _historyHoverIndex)
        Dim isSelected = ((e.State And DrawItemState.Selected) = DrawItemState.Selected)

        ' 使用與 ListView 相似的 Hover 與 Select 背景色
        Dim backColor = If(isSelected, ThemeColors.AltoGray, If(isHovered, ThemeColors.MercuryGray, Color.White))
        Dim foreColor = Color.Black

        Using brush = New SolidBrush(backColor)
            e.Graphics.FillRectangle(brush, e.Bounds)
        End Using

        '' 開啟 GDI+ 的平滑抗鋸齒渲染，解決字體邊緣粗糙的問題
        'e.Graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit
        'Dim textRect As New RectangleF(e.Bounds.X + 4, e.Bounds.Y + 3, e.Bounds.Width - 8, e.Bounds.Height - 4)
        'Using brush = New SolidBrush(foreColor)
        '    e.Graphics.DrawString(HistoryListBox.Items(e.Index).ToString(), e.Font, brush, textRect)
        'End Using

        ' 恢復使用系統原生的 TextRenderer 確保呈現與普通 ListBox 相同的柔和抗鋸齒
        Dim textRect As New Rectangle(e.Bounds.X + 4, e.Bounds.Y, e.Bounds.Width - 4, e.Bounds.Height)
        Dim flags = TextFormatFlags.VerticalCenter Or TextFormatFlags.Left Or TextFormatFlags.EndEllipsis Or TextFormatFlags.PreserveGraphicsClipping
        TextRenderer.DrawText(e.Graphics, HistoryListBox.Items(e.Index).ToString(), e.Font, textRect, foreColor, flags)

    End Sub

    Private Async Sub SaveCache_Click(sender As Object, e As EventArgs) Handles SaveCache.Click
        Await SaveCachesToSQLiteAsync()
    End Sub
    Private Async Sub LoadCache_Click(sender As Object, e As EventArgs) Handles LoadCache.Click
        Await LoadCachesFromSQLiteAsync()
        Dim st = GetDatabaseSummary()
        ProgressBar2.Text = $"DB 統計 — folder_stats:{st.fc} 筆 / attach_maillist:{st.mb} 筆 / attach_filenames:{st.at} 筆 / year_counts:{st.yc} 筆 / month_counts:{st.mc} 筆 / {st.kb} KB"

    End Sub
    Private Async Sub RenewCache_Click(sender As Object, e As EventArgs) Handles RenewCache.Click
        ' 2026/04/09 重構: 原本只做孤兒清除，現在改呼叫完整的 RenewCacheAsync
        '   RenewCacheAsync 內含: Phase1 BFS → Phase2 snapshot 比對 → Phase3 dirty 重算
        '                         Phase4 ancestor 聚合清除 → Phase5 month_counts DB 清除
        '                         Phase6 CleanupOrphan + SaveCachesToSQLiteAsync
        '   RenewIncludeSize 勾選時才重算 folder_size（GetTable 遍歷，大資料夾較慢）
        Await RenewCacheAsync(RenewIncludeSize.Checked)
    End Sub
    Private Async Sub ClearCache_Click(sender As Object, e As EventArgs) Handles ClearCache.Click
        ' ---------------------------------------------------------------
        ' ClearCache_Click — [透明化控制] 分流清理記憶體或 SSD 快取
        ' by Gemini, 2026/04/10: 實作三路自選對話框
        ' ---------------------------------------------------------------
        _dbg("開始")
        Dim st = GetDatabaseSummary()
        Dim lastTimeStr As String = st.lastTs

        ' 1. 準備訊息文字
        Dim msg As String = $"【快取清理選項】" & vbCrLf & vbCrLf &
                            $"--- 目前 SSD 快取現況 ---" & vbCrLf &
                            $"最後儲存時間：{lastTimeStr}" & vbCrLf &
                            $"資料夾統計：{st.fc} 筆" & vbCrLf &
                            $"附件郵件：{st.mb} 筆" & vbCrLf &
                            $"檔案大小：{st.kb} KB" & vbCrLf & vbCrLf &
                            $"請選擇你要清理的範圍："

        ' 2. 使用動態 Form 實作三按鈕對話框 (為了精確符合使用者需求)
        Using f As New Form()
            f.Text = "清理快取" : f.Size = New Size(450, 280)
            f.StartPosition = FormStartPosition.CenterParent : f.FormBorderStyle = FormBorderStyle.FixedDialog
            f.MaximizeBox = False : f.MinimizeBox = False : f.BackColor = Color.White
            f.Font = New Font("Microsoft JhengHei UI", 10)

            Dim lbl As New Label() With {.Text = msg, .Location = New Point(20, 20), .Size = New Size(400, 150)}
            f.Controls.Add(lbl)

            Dim btnMem As New Button() With {.Text = "僅記憶體", .DialogResult = DialogResult.Yes, .Location = New Point(20, 180), .Size = New Size(120, 40), .BackColor = Color.LightBlue}
            Dim btnSSD As New Button() With {.Text = "僅 SSD (重建)", .DialogResult = DialogResult.No, .Location = New Point(155, 180), .Size = New Size(120, 40), .BackColor = Color.MistyRose}
            Dim btnBoth As New Button() With {.Text = "兩者皆清", .DialogResult = DialogResult.Retry, .Location = New Point(290, 180), .Size = New Size(120, 40), .BackColor = Color.Orange}

            f.Controls.AddRange({btnMem, btnSSD, btnBoth})
            f.AcceptButton = btnMem

            Dim result = f.ShowDialog()

            ' 3. 根據選擇執行處置
            Select Case result
                Case DialogResult.Yes ' 僅記憶體
                    ClearMemoryCachesInternal()
                    ProgressBar2.Text = "已完成：僅清除記憶體快取 (SSD 保留)"
                    _dbg("清理", "僅記憶體")

                Case DialogResult.No ' 僅 SSD
                    If MessageBox.Show("【安全提示】這將把目前的 SSD 快取檔更名備份 ( .zip) 並重新建立空白資料表。" & vbCrLf & "這可以解決 Schema 不相容問題且具備救援機制，確定嗎？", "重置 SSD 快取", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = DialogResult.OK Then
                        Await ResetSSDatabaseAsync()
                        ProgressBar2.Text = "已完成：SSD 資料庫已備份並重新初始化"
                        _dbg("清理", "僅 SSD (已備份)")
                    End If

                Case DialogResult.Retry ' 兩者皆清
                    If MessageBox.Show("確定要清除記憶體並備份重置 SSD 快取嗎？", "最後確認", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) = DialogResult.OK Then
                        ClearMemoryCachesInternal()
                        Await ResetSSDatabaseAsync()
                        ProgressBar2.Text = "已完成：記憶體與 SSD 快取已全數歸零 (舊 SSD 檔已備份)"
                        _dbg("清理", "FULL CLEAN (已備份)")
                    End If
            End Select
        End Using

        _dbg("結束")

    End Sub
#End Region
#Region "  ├ 滑鼠 & 鍵盤操作事件"
    Private Sub SplitContainerMouseDown(sender As Object, e As MouseEventArgs)
        ''' <summary>
        ''' 強制讓 SplitContainer 完全無法被點選、不顯示虛線焦點框
        ''' 使用 Win32 API 直接修改視窗樣式 (最強力做法)
        ''' </summary>
        ''' <param name="sc">要禁用的 SplitContainer 控制項</param>
        ''' <remarks>
        ''' 解決 SplitContainer 預設會顯示焦點虛線框的問題，僅保留 MouseMove 改變游標的功能。
        ''' 不影響內部控制項的操作。
        ''' </remarks>
        ''' 共用的側邊欄切換事件 (2026/03/28 by Gemini 改良：偵測雙擊分隔線縮放)

        _dbg("開始")
        ' 只針對滑鼠左鍵，且連按二下 (Double Click) 觸發
        If e.Button = MouseButtons.Left AndAlso e.Clicks = 2 Then
            Dim sc = TryCast(sender, SplitContainer)
            If sc Is Nothing Then Return
            ' 臨界值 20px，如果大於此寬度則進行縮合
            If sc.SplitterDistance > 20 Then
                sc.Tag = sc.SplitterDistance        ' 💡 記憶當前寬度在 Tag 屬性，以便下次恢復
                sc.SplitterDistance = 10            ' 縮合至 10px 觸控區
                _dbg("縮合側邊欄", $"{sc.Name} → 10px (原 {sc.Tag}px)") ' by Gemini, 2026/04/04: Issue 4 格式標準化
            Else
                ' 💡 恢復寬度，若無紀錄則預設為 250px
                Dim prevDist As Integer = If(TypeOf sc.Tag Is Integer, DirectCast(sc.Tag, Integer), 250)
                If prevDist < 50 Then prevDist = 250    ' 防止恢復值過小
                sc.SplitterDistance = prevDist
                _dbg("恢復側邊欄", $"{sc.Name} → {prevDist}px") ' by Gemini, 2026/04/04: Issue 4 格式標準化
            End If
        End If
        _dbg("結束")

    End Sub
    Private Sub HandleTreeViewMouseHover(sender As Object, e As EventArgs)
        ' ---------------------------------------------------------------
        ' 共用 TreeView / SimTree MouseHover 處理 (MouseMove + MouseLeave)
        ' by Gemini, 2026/04/03 整合優化，提升 SimTree 離開控制項時的視覺穩定性
        '
        ' [2026-03-17 原始規劃保留]: 兩段結構對稱，各用一個布林封裝 SimTree 例外
        ' 還原規則:
        '   SimTree 選取節點 → 還原選取色 (不能用 Color.Empty，否則藍色會閃掉)
        '   其餘節點         → Color.Empty (原生 TreeView 預設)
        '
        ' 套用規則:
        '   SimTree 選取節點 → 跳過 (選取色優先，不蓋 hover 色)
        '   其餘節點         → 淡灰色 hover
        ' ---------------------------------------------------------------
        '_dbg("開始")
        ' todo: 這裡一直在被持續觸發，滑鼠沒在移動也一直發生, 如何設條件限制不要一直進來??
        Dim tv As TreeView = CType(sender, TreeView)
        Dim mouseE = TryCast(e, MouseEventArgs)
        Dim node As TreeNode = If(mouseE IsNot Nothing, tv.GetNodeAt(mouseE.Location), Nothing)

        If node Is _lastHoveredTreeNode Then Return

        ' ── 還原上一個 hover 節點 (對稱結構第一部分) ──
        If _lastHoveredTreeNode IsNot Nothing Then
            Dim sim As SimTree = TryCast(tv, SimTree)

            If sim IsNot Nothing AndAlso sim.SelectedNodes.Contains(_lastHoveredTreeNode) Then
                ' SimTree 選取節點: 根據焦點還原正確的選取色 (不能 Color.Empty)
                _lastHoveredTreeNode.BackColor = If(sim.Focused, SystemColors.Highlight, ThemeColors.MercuryGray)
                _lastHoveredTreeNode.ForeColor = If(sim.Focused, SystemColors.HighlightText, SystemColors.InactiveCaptionText)
            Else
                _lastHoveredTreeNode.BackColor = Color.Empty
                _lastHoveredTreeNode.ForeColor = Color.Empty
            End If
        End If

        ' ── 套用新 hover 色 (對稱結構第二部分) ──
        If node IsNot Nothing Then
            Dim skipHover As Boolean = TypeOf tv Is SimTree AndAlso CType(tv, SimTree).SelectedNodes.Contains(node)
            If Not skipHover Then
                node.BackColor = ThemeColors.MercuryGray
                node.ForeColor = SystemColors.InactiveCaptionText
            End If
        End If
        _lastHoveredTreeNode = node

    End Sub
    Private Sub HandleListViewMouseHover(sender As Object, e As EventArgs)
        ' by Gemini, 2026/04/03: 整合 MouseMove 與 MouseLeave 為單一維護點
        '_dbg("開始") ' todo: 這裡一直在被持續觸發，滑鼠沒在移動也一直發生, 如何設條件限制不要一直進來??

        Dim listView As ListView = TryCast(sender, ListView)
        If listView Is Nothing Then Return

        ' by Gemini, 2026/04/10: 虛擬模式下頻繁修改 BackColor 會觸發大量繪製導致閃爍
        ' 虛擬模式應由系統處理自己的 Hover 狀態，或僅針對非虛擬模式執行此邏輯
        If listView.VirtualMode Then Return

        ' 1. 判斷目前的目標項目 (如果是 MouseLeave 則為 Nothing)
        Dim currentItem As ListViewItem = Nothing
        Dim mouseE = TryCast(e, MouseEventArgs)
        If mouseE IsNot Nothing Then currentItem = listView.GetItemAt(mouseE.X, mouseE.Y)

        ' 2. 檢查目標是否改變 (優化效能，若相同則不重繪)
        If currentItem Is _lastHoveredListItem Then Return

        ' 3. 處理狀態轉變: 清除舊背景色並套用新色
        ' 2026/04/14 by Simon/Claude: Tag=Nothing 的行（群組標題 / 合計列）有固定 BackColor，
        '   離開時要還原原色而非 Color.Empty；進入時也不套 hover 灰，保持原色不被蓋掉。
        If _lastHoveredListItem IsNot Nothing Then
            If _lastHoveredListItem.Tag Is Nothing Then
                _lastHoveredListItem.BackColor = GetHeaderRowBackColor(_lastHoveredListItem)  ' 還原標題/合計列原色
            Else
                _lastHoveredListItem.BackColor = Color.Empty
            End If
        End If
        If currentItem IsNot Nothing AndAlso currentItem.Tag IsNot Nothing Then
            currentItem.BackColor = ThemeColors.MercuryGray   ' 只對一般列套 hover 色
        End If

        _lastHoveredListItem = currentItem

    End Sub
    Private Function GetHeaderRowBackColor(item As ListViewItem) As Color
        ' 2026/04/14 by Simon/Claude: 根據 Tag=Nothing 列的文字前綴還原正確的 BackColor
        '   ▸ 開頭 = 群組標題行 → SystemColors.GradientInactiveCaption（淡藍）
        '   ▶ 開頭 = 合計列     → Color.FromArgb(220, 235, 252)（稍深藍）
        '   其他    = fallback   → Color.Empty
        If item.Text.StartsWith("▸") Then Return SystemColors.GradientInactiveCaption
        If item.Text.StartsWith("▶") Then Return Color.FromArgb(220, 235, 252)
        Return Color.Empty
    End Function
    Private Sub HandleTreeViewKeyPress(sender As Object, e As KeyPressEventArgs)
        ' 在這裡處理所有TreeView KeyPress 事件的程式碼
        _dbg("開始")

        If TypeOf sender Is TreeView Then
            If e.KeyChar = ChrW(Keys.Enter) Then
                sender.SelectedNode.Expand()            ' 按Enter展開下一層
                Select Case sender.Name
                    Case "TreeView1" : ListView1.Focus()
                    Case "SimTree2" : ListView2.Focus()
                End Select

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

        _dbg("結束")
    End Sub
    Private Async Sub HandleListViewKeyPress(sender As Object, e As KeyPressEventArgs)
        ' ---------------------------------------------------------------
        ' 共用 ListView KeyPress 處理 (完整替換舊的 HandleListViewKeyPress)
        ' 2026-03-16 C3 重構: 把原本分散在三個 ListView 的 KeyPress 邏輯統一到這裡，根據 s 來辨識是哪個 ListView，實作各自的行為
        ' 目前的實現是直接在 KeyPress 事件裡處理所有邏輯，是否要把各區塊或年度視圖和月份視圖的 Enter 鍵行為分別封裝成獨立的方法?
        '
        ' 各 ListView 行為:
        '    ListView1 : Enter = 進入子資料夾, ESC = 退回上一層  (原有邏輯不變)
        '    ListView2 : Enter = 等同雙擊 (進入月份或返回年度) , ESC = 返回年度視圖
        '    ListView3 : Enter = 打開郵件, ESC = 取消選取
        ' ---------------------------------------------------------------
        _dbg("開始", e.KeyChar)
        Dim cToken As CancellationToken = OkayNowYouHaveToken()  ' ✅ 取得新 Token，同時取消上一次未完成的操作

        Dim lv As ListView = TryCast(sender, ListView)
        If lv Is Nothing Then Return

        ' ---------------------------------------------------------------
        ' ListView1: 資料夾導覽 (保留原有邏輯，從 ListView1_KeyPress 移到這裡統一管理)
        ' ---------------------------------------------------------------
        If lv Is ListView1 Then
            If e.KeyChar = ChrW(Keys.Enter) Then
                If lv.SelectedItems.Count = 0 Then Return

                ' by Gemini 3 Flash, 2026/04/13: 選取多個項目時，改用 MessageBox 顯示數量加總 (參考 DebugForm.CalculateSelectedTimeSpan)
                If lv.SelectedItems.Count > 1 Then
                    Dim sumDirect As Long = 0 : Dim sumTotal As Long = 0
                    For Each item As ListViewItem In lv.SelectedItems
                        ' 2026/04/13 v2: 移除「所屬父資料夾」欄後，索引回歸原位
                        ' SubItems(1): 郵件數量(直屬)；SubItems(3): 郵件總計(含子孫)
                        ' 群組標題行 / 合計列 Tag=Nothing，但其數字欄也可加總（代表該組彙總值）
                        Dim strDirect As String = item.SubItems(1).Text.Replace(",", "").Trim()
                        Dim strTotal As String = item.SubItems(3).Text.Replace(",", "").Trim()
                        Dim valDirect As Long = 0 : Dim valTotal As Long = 0
                        Long.TryParse(strDirect, valDirect) : Long.TryParse(strTotal, valTotal)
                        sumDirect += valDirect : sumTotal += valTotal
                    Next
                    MessageBox.Show($"已選取 {lv.SelectedItems.Count} 個資料夾統計結果：" & vbCrLf & vbCrLf &
                                    $"【本層郵件】加總：{sumDirect:N0} 封" & vbCrLf &
                                    $"【包含子樹】加總：{sumTotal:N0} 封", "複選數量加總", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If

                Dim selectedItem As ListViewItem = lv.SelectedItems(0)          ' 獲取點選的資料夾並進入 (單選時維持原邏輯)
                If selectedItem IsNot Nothing Then EnterSelectedFolder(selectedItem)

            ElseIf e.KeyChar = ChrW(Keys.Escape) Then                           ' 退回上一層資料夾
                Dim itemName As String = lv.Items(0).Text                       ' 記下現在所在的listviewItem
                ' 2026/04/13 by Simon/Claude: Tab1 改用 SimTree1
                Dim node As TreeNode = SimTree1.SelectedNode
                If node IsNot Nothing AndAlso node.Parent IsNot Nothing Then
                    node.Collapse() : SimTree1.SelectedNode = node.Parent      ' 選取其上層資料夾
                    Dim item As ListViewItem = FindLiSVItemByName(lv, itemName) ' 找出剛才退出前的資料夾
                    If item IsNot Nothing Then item.Selected = True : item.Focused = True : lv.Focus()
                End If

            ElseIf e.KeyChar = ChrW(1) Then ' Ctrl-A 全選 listview1 所有項目 — 2026/3/26 by Gemini
                lv.BeginUpdate()
                For Each item As ListViewItem In lv.Items
                    item.Selected = True
                Next
                lv.EndUpdate()
                e.Handled = True
            End If
            _dbg("結束")
            Return
        End If

        ' ---------------------------------------------------------------
        ' ListView2: 年度 / 月份視圖導覽
        ' ---------------------------------------------------------------
        If lv Is ListView2 Then
            If e.KeyChar = ChrW(Keys.Enter) Then                ' Enter = 等同雙擊目前選定的項目
                If lv.SelectedItems.Count = 0 Then Return

                ' by Gemini 3 Flash, 2026/04/13: 選取多個項目時，改用 MessageBox 顯示數量加總 (參考 DebugForm.CalculateSelectedTimeSpan)
                If lv.SelectedItems.Count > 1 Then
                    Dim sumYearMonth As Long = 0
                    For Each item As ListViewItem In lv.SelectedItems
                        ' 跳過特殊控制列或標題列 (例如 "BACK" 或含有 "──")
                        If item.Tag?.ToString() = "BACK" OrElse item.Text.Contains("──") Then Continue For
                        ' SubItems(1): 郵件個數
                        Dim strCount As String = item.SubItems(1).Text.Replace(",", "").Trim()
                        Dim valCount As Long = 0
                        Long.TryParse(strCount, valCount)
                        sumYearMonth += valCount
                    Next
                    MessageBox.Show($"已選取 {lv.SelectedItems.Count} 個統計項目：" & vbCrLf & vbCrLf &
                                    $"郵件總計：{sumYearMonth:N0} 封", "複選數量加總", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If

                Dim selectedItem As ListViewItem = lv.SelectedItems(0)
                If _tab2IsMonthView AndAlso                     ' 在月份視圖按 Enter 於返回列 → 回到年度視圖
                    selectedItem.Tag IsNot Nothing AndAlso
                    selectedItem.Tag.ToString() = "BACK" Then
                    Await GoToYearView()                        ' ✅ 2026/04/12 重構: ShowYearView 已改名為 GoToYearView (共用導覽，DoubleClick/KeyPress 共用)

                ElseIf Not _tab2IsMonthView Then                ' 在年度視圖按 Enter → 進入月份視圖
                    Dim selectedYear As Integer = 0
                    If Integer.TryParse(selectedItem.Text.Trim(), selectedYear) AndAlso ' ✅ 2026/04/12 重構: ShowMonthView 已改名為 GoToMonthView
                        _tab2FolderList IsNot Nothing AndAlso _tab2FolderList.Count > 0 Then Await GoToMonthView(selectedYear, cToken)
                End If

            ElseIf e.KeyChar = ChrW(Keys.Escape) Then           ' ESC: 不管在哪個視圖，一律返回年度視圖
                If _tab2IsMonthView Then Await GoToYearView()   ' ✅ 2026/04/12 重構: ShowYearView 已改名為 GoToYearView
            End If
            _dbg("結束")
            Return
        End If

        ' ---------------------------------------------------------------
        ' ListView3: Tab3 附件搜尋結果的鍵盤操作
        ' Enter = 用 EntryID 打開郵件 (第 6 欄 SubItems(5))
        ' ESC   = 清除目前選取
        ' 2026-03-16 實作完成
        ' ---------------------------------------------------------------
        If lv Is ListView3 Then
            If e.KeyChar = ChrW(Keys.Enter) Then
                If lv.SelectedItems.Count = 0 Then Return
                OpenMailByEntryID(lv.SelectedItems(0).SubItems(5).Text) ' Enter = 用 EntryID 打開郵件 (第 6 欄 SubItems(5))

            ElseIf e.KeyChar = ChrW(Keys.Escape) AndAlso Not lv.VirtualMode Then
                If lv.SelectedItems.Count > 0 Then lv.SelectedItems(0).Selected = False
            End If
            _dbg("結束")
            Return
        End If

    End Sub
    Private Function FindLiSVItemByName(listview As ListView, itemName As String) As ListViewItem
        _dbg("開始", listview.Name)
        For Each item As ListViewItem In listview.Items
            If item.Text.Replace(" - ", "") = itemName.Replace(" - ", "") Then Return item
        Next : Return Nothing

    End Function
#End Region
#Region "  └ 其他輔助事件"
    Private Async Sub ExpandTreeToDefaultInbox(tv As TreeView)
        ' by Gemini, 2026/04/06: 使用Guard Clauses重構，減少巢狀層數並確保 EndUpdate 執行安全性
        _dbg("開始", tv.Name)

        ' 1. 第一層Guard Clauses：沒節點直接走人
        If tv.Nodes.Count = 0 Then Return

        ' 2. 第二層Guard Clauses：根節點沒子節點也沒什麼好展開的
        Dim rootNode = tv.Nodes(0)
        If rootNode.Nodes.Count = 0 Then Return

        ' by Gemini, 2026/04/07: 分離 UI 展開 與 資料載入(AfterSelect)，讓樹狀圖以最快極速展開完畢，不再卡這 100ms
        Dim nodeToSelect As TreeNode = Nothing
        tv.BeginUpdate()
        Try
            rootNode.Expand()
            ' ✅ 修正: 應遍歷第一個 PST 的「子資料夾」數量，而非根節點數量
            ' 舊版: tv.Nodes.Count - 1 = PST 個數 (通常=1) ，只會檢查第一個子資料夾
            ' 新版: tv.Nodes(0).Nodes.Count - 1 = 第一個 PST 下的所有子資料夾數
            ' 遍歷第一個 PST 的「子資料夾」
            For Each node As TreeNode In rootNode.Nodes
                ' 3. 第三層Guard Clauses：不是收件匣就繼續找下一個 (過濾模式)
                If Not (node.Text.Contains("Inbox") Or node.Text.Contains("收件匣")) Then Continue For
                _dbg("發現預設收件匣", node.FullPath)
                nodeToSelect = node
                Exit For
            Next
            If nodeToSelect Is Nothing Then
                _dbg("結束", $"{tv.Name}: 找不到預設收件匣，根節點共 {rootNode.Nodes.Count} 個子資料夾")
            End If
        Finally
            ' 💡 確保無論中途 Return 或發生 Exception，UI 都不會卡在 BeginUpdate
            tv.EndUpdate()
        End Try

        ' 找到節點的話，在 EndUpdate 解鎖 UI 後，由以下區塊執行「資料載入觸發」
        If nodeToSelect IsNot Nothing Then
            Await Task.Yield() ' 讓 UI 執行緒去把因為 EndUpdate 而要畫的圖立刻畫出來

            ' 4. 使用 TryCast 簡化類型判斷，減少多層 If
            Dim st = TryCast(tv, SimTree)
            If st IsNot Nothing Then
                ' 2026/3/18: 必須明確 TryCast 到 SimTree，才能正確呼叫 AddSelectedNode 更新 _selectedNodes 和高亮色
                ' 同時, 把自訂控制項裡面的 FireAfterSelect() 從 private 改成 public, 直接手動觸發 AfterSelect 事件
                st.AddSelectedNode(nodeToSelect)    ' ← SimTree 專用路徑: 直接更新 _selectedNodes + 高亮
                st.FireAfterSelect(nodeToSelect)    ' ← 直接手動觸發 AfterSelect 事件，讓統計邏輯跑起來
            ElseIf TypeOf tv Is TreeView Then
                ' 2026/3/18: debug找了好幾天, 首次切換到tab2時, SimTree2無法正確選取到預設的收件匣
                ' 結果原來是下方的 tv.SelectedNode = node 送到SimTree控制項, 沒有被觸發選中的event
                ' 一定要自己主動去手動觸發 FireAfterSelect 事件
                tv.SelectedNode = nodeToSelect
            End If
            tv.Focus()
            _dbg("結束", $"{tv.Name}: 已成功選取預設收件匣")
        End If

    End Sub
    Private Function GetActiveTreeView() As TreeView
        ''' <summary>
        ''' 根據 TabControl1 的選擇索引，判斷並傳回當前畫面上活動中的 TreeView/SimTree, by Gemini, 2026/03/30
        ''' </summary>
        ' 在需要觸發 AfterSelect 或其他操作時，能夠根據目前選中的 Tab 頁面，準確地獲取對應的 TreeView 控制項
        Select Case TabControl1.SelectedIndex
            Case 0 : Return SimTree1   ' 2026/04/13 by Simon/Claude: Tab1 改用 SimTree1
            Case 1 : Return SimTree2
            Case 2 : Return SimTree3
            Case 3 : Return TreeView4
            Case 4 : Return TreeView5
            Case Else : Return Nothing
        End Select

    End Function
    Private Function GetAllTreeViews(container As Control) As List(Of TreeView)
        ''' <summary>
        ''' 遞迴搜尋容器內所有的 TreeView (含其衍生子類如 SimTree)
        ''' </summary>
        Dim list As New List(Of TreeView)
        For Each ctrl As Control In container.Controls
            If TypeOf ctrl Is TreeView Then list.Add(CType(ctrl, TreeView)) ' 如果是 TreeView 或其衍生類 (SimTree)
            If ctrl.HasChildren Then list.AddRange(GetAllTreeViews(ctrl))   ' 如果有子容器 (如 SplitContainer, TabControl, Panel)，繼續遞迴往下層掃描
        Next
        Return list

    End Function
    Private Sub TriggerAfterSelect(tv As TreeView)
        ''' <summary>
        ''' 定向觸發活動控制項的數據刷新 (AfterSelect)
        ''' 僅作為補強機制，確保右側統計數據在特殊情況下能被手動刷新, by Gemini, 2026/03/30
        ''' </summary>

        ' 確保有選定節點才執行統計，否則統計函數會報錯
        If tv Is Nothing Then Return
        Dim targetNode As TreeNode = tv.SelectedNode
        If targetNode Is Nothing Then Return

        Dim args As New TreeViewEventArgs(targetNode)
        ' 2026/04/13 by Simon/Claude: Tab1 改用 SimTree1，所有 AfterSelect 都走 FireAfterSelect 路徑
        If TypeOf tv Is SimTree Then
            DirectCast(tv, SimTree).FireAfterSelect(targetNode)
        ElseIf TypeOf tv Is SimTree Then
            DirectCast(tv, SimTree).FireAfterSelect(targetNode)   ' SimTree2/3/4
        End If

    End Sub

    Private Sub AutoResizeListViewColumns(lv As ListView)
        ''' <summary>
        ''' 定義各個 ListView 縮放時的欄位寬度比例
        ''' 2026/04/01 by Gemini
        ''' </summary>

        If lv.Columns.Count = 0 OrElse lv.Width <= 0 Then Return

        Dim w As Integer = lv.ClientSize.Width ' 使用 ClientSize 避免捲軸吃掉寬度
        If lv Is ListView1 Then ' Tab1: 資料夾名稱 / 郵件數量 / 資料夾數量 / 郵件總計 / 大小
            ' 2026/04/13 v2: 移除「所屬父資料夾」欄，回歸 5 欄
            If lv.Columns.Count >= 5 Then
                lv.Columns(1).Width = CInt(w * 0.15)
                lv.Columns(2).Width = CInt(w * 0.15)
                lv.Columns(3).Width = CInt(w * 0.15)
                lv.Columns(4).Width = CInt(w * 0.188)
                lv.Columns(0).Width = w - (lv.Columns(1).Width + lv.Columns(2).Width + lv.Columns(3).Width + lv.Columns(4).Width) - 5
            End If

        ElseIf lv Is ListView2 Then ' Tab2: 年度 / 郵件個數 / 空白欄位
            If lv.Columns.Count >= 3 Then
                lv.Columns(0).Width = Math.Max(120, CInt(w * 0.3)) ' 第一欄(年度/月份)至少保底 120px
                lv.Columns(1).Width = Math.Max(100, CInt(w * 0.2)) ' 第二欄(郵件數量)至少保底 100px
                lv.Columns(2).Width = Math.Max(0, w - lv.Columns(0).Width - lv.Columns(1).Width - 5)  ' 第三欄吸收所有剩餘空間
                ' 2026/04/03 by Gemini: 將無用的第三欄作為彈性緩衝區。當視窗縮小時，優先壓縮第三欄位，確保前兩欄至少有基本的顯示空間而不會擠在一起。
            End If

        ElseIf lv Is ListView3 Then ' Tab3: 郵件主旨 / 郵件大小 / 收到日期 / 寄件者 / 附件個數 / EntryID
            If lv.Columns.Count >= 6 Then
                lv.Columns(1).Width = CInt(w * 0.15)    ' 郵件大小
                lv.Columns(2).Width = CInt(w * 0.2)     ' 收到日期
                lv.Columns(3).Width = CInt(w * 0.15)    ' 寄件者
                lv.Columns(5).Width = CInt(w * 0.03)    ' EntryID (隱藏?)
                lv.Columns(4).Width = If(CheckAttCount.Checked, CInt(w * 0.1), 0.03)   ' 2026/04/01 by Gemini: 根據勾選狀態 動態顯示/隱藏 附件個數欄位
                lv.Columns(0).Width = w - (lv.Columns(1).Width + lv.Columns(2).Width + lv.Columns(3).Width + lv.Columns(4).Width + lv.Columns(5).Width) - 5
            End If

        ElseIf lv Is ListView4 Then ' Tab4: 主旨 / 大小 / 收到時間 / 寄件者 / EntryID
            If lv.Columns.Count >= 5 Then
                lv.Columns(1).Width = CInt(w * 0.15)    ' 大小
                lv.Columns(2).Width = CInt(w * 0.2)     ' 收到時間
                lv.Columns(3).Width = CInt(w * 0.2)     ' 寄件者
                lv.Columns(4).Width = CInt(w * 0.03)    ' EntryID (隱藏?)
                lv.Columns(0).Width = w - (lv.Columns(1).Width + lv.Columns(2).Width + lv.Columns(3).Width + lv.Columns(4).Width) - 5
            End If
        Else
            ' 預設比例: 首欄固定40%，其餘均分
            If lv.Columns.Count > 0 Then
                Dim firstWidth As Integer = CInt(w * 0.4)
                lv.Columns(0).Width = firstWidth
                If lv.Columns.Count > 1 Then
                    Dim remainWidth As Integer = w - firstWidth - 5
                    Dim avgWidth As Integer = remainWidth \ (lv.Columns.Count - 1)
                    For i As Integer = 1 To lv.Columns.Count - 1
                        lv.Columns(i).Width = avgWidth
                    Next
                End If
            End If
        End If

    End Sub
    Private Sub UpdateNumericIncrement(num As NumericUpDown, unitCombobox As ComboBox)
        ''' <summary>
        ''' 根據當前選擇的單位與數值，動態更新 NumericUpDown 的增減幅度 (2026/04/05 by Gemini)
        ''' </summary>

        If num Is Nothing OrElse unitCombobox Is Nothing Then Return
        Dim unit As String = If(unitCombobox.SelectedItem?.ToString(), "KB")

        If unit = "MB" OrElse unit = "GB" Then  ' MB/GB 單位下，固定增量為 1
            num.Maximum = 1024
            num.Minimum = 0.1
            num.Increment = 0.1
            num.DecimalPlaces = 1
        Else                                    ' KB 單位下，根據數值範圍採用不同的階梯式增量
            num.Maximum = 9999
            num.Minimum = 1
            num.DecimalPlaces = 0
            Dim val = num.Value
            If val < 50 Then
                num.Increment = 1
            ElseIf val < 210 Then
                num.Increment = 10
            Else
                num.Increment = 100
            End If
        End If
    End Sub
    Private Sub AppendStatusHistory(msg As String, source As String)
        If String.IsNullOrWhiteSpace(msg) Then Return

        ' Smart Overwrite 原則: 因為改為最新一筆在底下，所以比對最新一筆為 (Count - 1)
        If source = "PB2" AndAlso _statusHistory.Count > 0 Then
            Dim lastItem = _statusHistory(_statusHistory.Count - 1)
            If lastItem.Source = "PB2" Then
                Dim prefixLen = Math.Min(10, Math.Min(msg.Length, lastItem.Message.Length))
                If prefixLen > 0 AndAlso msg.Substring(0, prefixLen) = lastItem.Message.Substring(0, prefixLen) Then
                    _statusHistory(_statusHistory.Count - 1) = New StatusHistoryItem With {.Time = DateTime.Now, .Message = msg, .Source = source}
                    Return
                End If
            End If
        End If

        _statusHistory.Add(New StatusHistoryItem With {.Time = DateTime.Now, .Message = msg, .Source = source})
        If _statusHistory.Count > MAX_HISTORY_COUNT Then _statusHistory.RemoveAt(0)

    End Sub
    Private Sub ShowHistoryPopup(source As String, clickedLabel As ToolStripStatusLabel)
        ' 分別過濾 PB1 與 PB2 紀錄
        Dim filteredHistory = _statusHistory.Where(Function(hisItem) hisItem.Source = source).ToList()
        If filteredHistory.Count = 0 Then Return

        If _historyPopup Is Nothing Then
            HistoryListBox = New ListBox() With {.BorderStyle = BorderStyle.None,
                                                 .Font = New Font(Me.Font.FontFamily, 9),
                                                 .IntegralHeight = False,
                                                 .DrawMode = DrawMode.OwnerDrawFixed,
                                                 .ItemHeight = 24}

            Dim host = New ToolStripControlHost(HistoryListBox)
            host.Margin = New Padding(0)
            host.Padding = New Padding(0)

            _historyPopup = New ToolStripDropDown()
            _historyPopup.Items.Add(host)
            _historyPopup.Padding = New Padding(1)
            _historyPopup.BackColor = ThemeColors.AltoGray
            _historyPopup.DropShadowEnabled = True

            ' 防止點選 ListBox 項目時 ToolStrip 自動關閉
            AddHandler _historyPopup.Closing, Sub(s, ev)
                                                  If ev.CloseReason = ToolStripDropDownCloseReason.ItemClicked Then ev.Cancel = True
                                              End Sub
        End If

        _historyHoverIndex = -1
        HistoryListBox.Items.Clear()
        For Each item In filteredHistory
            HistoryListBox.Items.Add($"[{item.Time:HH:mm:ss}] {item.Message}")
        Next

        ' 動態計算最佳寬度與高度
        Dim maxWidth As Integer = 300
        Using g = HistoryListBox.CreateGraphics()
            For Each item In HistoryListBox.Items
                Dim sz = g.MeasureString(item.ToString(), HistoryListBox.Font)
                If sz.Width > maxWidth Then maxWidth = CInt(sz.Width)
            Next
        End Using

        Dim numItemsToShow = Math.Min(15, filteredHistory.Count) ' 最多同時顯示 15 筆
        Dim targetWidth = maxWidth + 40
        Dim targetHeight = (HistoryListBox.ItemHeight * numItemsToShow) + 2

        ' by Gemini, 2026/04/02: 先解除前一次的限制，確保這次能夠正常縮小
        HistoryListBox.MinimumSize = Size.Empty

        ' 鐵血手段鎖死所有容器尺寸，解決被壓縮成 20px 的 Bug
        _historyPopup.AutoSize = False
        _historyPopup.Size = New Size(targetWidth, targetHeight)

        Dim hpHost = DirectCast(_historyPopup.Items(0), ToolStripControlHost)
        hpHost.AutoSize = False
        hpHost.Size = New Size(targetWidth, targetHeight)

        HistoryListBox.Size = New Size(targetWidth, targetHeight)
        HistoryListBox.MinimumSize = New Size(targetWidth, targetHeight)
        HistoryListBox.ClearSelected()

        ' 自動捲動到最底下 (因為最新的資料被放在清單底端)
        If HistoryListBox.Items.Count > numItemsToShow Then
            HistoryListBox.TopIndex = HistoryListBox.Items.Count - numItemsToShow
        End If

        Dim popupX As Integer = clickedLabel.Bounds.Left                    ' 將 Popup 精準顯示在 被點擊的 Label 正上方
        Dim maxRight As Integer = Screen.FromControl(Me).WorkingArea.Right  ' by Gemini, 2026/04/02: 防止 Popup 彈出時超過螢幕右緣被系統強制擠壓變形

        ' 若超過螢幕右緣，則將彈出位置往左平移，保留 5px 邊距
        Dim statusScreenPt = StatusStrip1.PointToScreen(New Point(popupX, 0))
        If statusScreenPt.X + targetWidth > maxRight Then
            Dim shift = (statusScreenPt.X + targetWidth) - maxRight + 5
            popupX -= shift
            If popupX < 0 Then popupX = 0
        End If

        Dim ptOffset = New Point(popupX, -(_historyPopup.Height + 5))
        _historyPopup.Show(StatusStrip1, ptOffset)

    End Sub

#End Region
    Public Class ThemeColors
        ' by Gemini, 2026/04/01: 統一管理專案色彩, 方便日後切換深色/淺色主題
        ''' <summary>主要視窗或Panel背景色 (#F2F2F2)</summary>
        Public Shared ReadOnly Gray95 As Color = Color.FromArgb(242, 242, 242)
        ''' <summary>滑鼠懸停(Hover)的背景色 (#E5E5E5)</summary>
        Public Shared ReadOnly MercuryGray As Color = Color.FromArgb(229, 229, 229)
        ''' <summary>輕微的格線或邊框色 (#E0E0E0)</summary>
        Public Shared ReadOnly AltoGray As Color = Color.FromArgb(224, 224, 224)
        ''' <summary>主視覺品牌藍色 (如按鈕、連結文字) (#0078D4)</summary>
        Public Shared ReadOnly Brand_Blue As Color = Color.FromArgb(0, 120, 212)
        ''' <summary>穩重的簡報藍色 (#4682B4)</summary>
        Public Shared ReadOnly Steel_Blue As Color = Color.FromArgb(70, 130, 180)
        ''' <summary>輕快的藍色 (#8DB3D3)</summary>
        Public Shared ReadOnly Polo_Blue As Color = Color.FromArgb(141, 179, 211)
        ''' <summary>深珊瑚紅 (#D83933)</summary>
        Public Shared ReadOnly CoralRed As Color = Color.FromArgb(216, 57, 51)
        ''' <summary>鐵鏽紅 (#A22C29)</summary>
        Public Shared ReadOnly RustRed As Color = Color.FromArgb(162, 44, 41)
        ''' <summary>深橘金色，用於平均線或參考線，具備極佳辨識度 (#E67E22)</summary>
        Public Shared ReadOnly DeepAmber As Color = Color.FromArgb(230, 126, 34)
        ''' <summary>在紅藍灰上都能看清的青色，用於平均線或參考線，具備極佳辨識度 (#00D4FF)</summary>
        Public Shared ReadOnly Cyan As Color = Color.FromArgb(0, 212, 255)

        ''' <summary>Chart2 背景 很淺的藍色 (#EDF4FF)</summary>
        Public Shared ReadOnly bgColor As Color = Color.FromArgb(237, 244, 255)
        ''' <summary>Chart2 格線 稍明顯的淡藍色 (#FFAA00)</summary>
        Public Shared ReadOnly gridLine As Color = Color.FromArgb(208, 223, 245)
        ''' <summary>Chart2 普通柱 天藍 (#4A8FD4)</summary>
        Public Shared ReadOnly barNormal As Color = Color.FromArgb(74, 143, 212)
        ''' <summary>Chart2 突顯柱 珊瑚紅 (#FF5533)</summary>
        Public Shared ReadOnly barHighlight As Color = Color.FromArgb(255, 85, 51)
        ''' <summary>Chart2 平均線 琥珀 (#FFAA00)</summary>
        Public Shared ReadOnly avgLineColor As Color = Color.FromArgb(255, 170, 0)
    End Class
#End Region

End Class
