Imports System.Numerics
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Windows
Imports System.Windows.Forms.DataVisualization.Charting
Imports Microsoft.Office.Interop.Outlook

Partial Class Form1

#Region "■ 00 Form 雙緩衝"
    'Protected Overrides ReadOnly Property CreateParams As CreateParams
    '    ' 2026/04/18 by Claude: 開啟 WS_EX_COMPOSITED 視窗合成模式
    '    ' 原理: Windows 把所有子控制項的繪製先合成到 offscreen buffer，再一次性 blit 到螢幕。
    '    '       與 InitTreeView / InitListView 裡各別控制項的 LVS_EX_DOUBLEBUFFER 不同層次:
    '    '         - 控制項層 (已有): 解決 ListView / TreeView 內部滾動的閃爍
    '    '         - Form 層 (本設定): 解決切換 Tab1~Tab5、Resize 視窗時整個視窗的撕裂感
    '    ' 注意: WS_EX_COMPOSITED 在某些透明子控制項或 OpenGL 繪圖上可能有副作用，
    '    '       若日後加入此類控制項出現異常，移除此 Override 即可還原。
    '    Get
    '        Dim cp As CreateParams = MyBase.CreateParams
    '        cp.ExStyle = cp.ExStyle Or &H2000000    ' WS_EX_COMPOSITED
    '        Return cp
    '    End Get
    'End Property
    Protected Overrides Sub WndProc(ByRef m As Message)
        ' 2026/05/07 by Claude: 攔截 WM_SIZE，修復最大化/還原瞬間 OwnerDraw ListView 殘影
        ' 原因：雙擊標題列觸發的 WindowState 切換是瞬間完成的，不像拖動是漸進式，
        '       OwnerDraw ListView 在視窗尺寸突變時，內部繪製管線會誤判 dirty region，
        '       導致部分項目或欄位未被正確重繪而產生殘影。
        ' 解法：在 SIZE_MAXIMIZED / SIZE_RESTORED 完成後，對所有 ListView 強制 Invalidate。
        MyBase.WndProc(m)
        ' 2026/06/19 by Simon/Claude: 追蹤是否在拖曳 size/move modal loop。拖曳邊緣會 ENTER/EXIT；
        ' 而 double-click 邊緣垂直頂天 / Shift+Win+↑↓ / 最大化還原不進此 loop，旗標維持 False。
        If m.Msg = WM_ENTERSIZEMOVE Then _inSizeMove = True
        If m.Msg = WM_EXITSIZEMOVE Then _inSizeMove = False
        If m.Msg = WM_SIZE Then
            Dim sizeType As Integer = m.WParam.ToInt32()
            If sizeType = SIZE_MAXIMIZED OrElse sizeType = SIZE_RESTORED Then
                For Each lv In GetAllLvList(Me) : lv.Invalidate() : Next
            End If
        End If
    End Sub
#End Region

#Region "■ 01 全域宣告"
    <System.Diagnostics.Conditional("DEBUG")>
    Private Sub _dbg(Optional msg As String = "", Optional detail As String = "", <System.Runtime.CompilerServices.CallerMemberName> Optional caller As String = "")
        ' Tier 1, 2026/06/15 by Simon/Claude Opus 4.8: 守衛提前 — 顯示關閉時直接 return。
        ' 原本守衛在最後一行，導致 _isDebugMode=False 時 573 處呼叫每次仍付出呼叫端解析成本後才丟掉。
        If Not _isDebugMode Then Return

        ' 2026/03/31 by Gemini: 改用 DebugForm 統一提供的 GetCallerName，此版本支援解析 Async 非同步方法名稱
        ' Tier 2, 2026/06/15 by Simon/Claude Opus 4.8: 預設改走編譯期注入的 CallerMemberName (零成本, 無 StackTrace/反射/Regex)。
        ' CallerMemberName 對 async 方法會自動還原乾淨原始名稱，但不帶 [Async] 標記。
        ' 若要恢復 [Async] 辨識：把 _useStackCaller 設 True 改走 GetCallerName；確定永遠不需要時，可直接把下面那行 If 註解掉。
        Dim realCaller As String = caller
        If _useStackCaller Then realCaller = DebugForm.GetCallerName()

        ' 2026/06/10 by Simon/Claude Opus 4.8: GetCallerName() 回傳 "Form1.MethodName"，
        ' 因所有呼叫端都在 Form1，"Form1." 前綴是冗餘資訊，直接 strip
        ' (CallerMemberName 回傳純方法名不含前綴，此行對它為 no-op；僅 GetCallerName 路徑會 strip)
        If realCaller.StartsWith("Form1.") Then realCaller = realCaller.Substring(6)

        ' by Gemini 3.5 Flash, 2026/06/19: 優先使用 ActiveInstance 以避免背景執行緒對 VB 預設實例的 Thread-Local 存取問題
        If DebugForm.ActiveInstance IsNot Nothing Then
            DebugForm.ActiveInstance.AddMessage3(msg, detail, realCaller)
        Else
            DebugForm.AddMessage3(msg, detail, realCaller)   ' fallback: ActiveInstance 未設(Load 前/已關閉), 退回原行為
        End If
    End Sub

    Private _isDebugMode As Boolean                     ' 是否為 Debug 模式，根據 VS 的編譯組態自動設定，是否顯示 DebugForm 以及是否啟用內部調試訊息
    Private _iLikeNoisy As Boolean = False              ' 是否啟用過濾debug message 噪音的功能，預設為 False 不顯示高頻率的迴圈訊息，想要詳細訊息轟炸就切成 True
    ' Tier 2, 2026/06/15 by Simon/Claude Opus 4.8: 呼叫端名稱解析方式開關。
    ' False = 用 CallerMemberName (編譯期注入, 零成本, 但 async 不帶 [Async] 標記)；
    ' True  = 用 GetCallerName (StackTrace, 較慢, 保留 [Async])
    Private _useStackCaller As Boolean = False

    'Private _isFirstInit As Boolean = True            ' 第一次啟動程式
    ' by Gemini, 2026/04/01: 延遲載入 UI 的狀態旗標
    ' Index   0: 取代原 _isFirstInit，標記 Form 與 Tab1 是否處於「首次啟動/首次選定」階段 (True=首次啟動中)
    ' Index 1~5: 對應 Tab1~Tab5 的 UI 是否已完成掛載 (True=已完成)
    Private _isTabInitialized(10) As Boolean            ' 記錄每個 Tab 的 UI 是否已經初始化完成, (0)是FormLoad的第一次啟動, (1)~(5)分別對應 Tab1~Tab5
    Private _isUserBusy As Boolean = False              ' ✅ 2026/04/01 by Gemini: 使用者操作忙碌旗標，用於暫緩背景預載程序
    Private _isClosing As Boolean = False               ' added by Gemini, 2026/04/08: 關閉流程旗標，確保 FormClosing 中的非同步儲存完成後再釋放資源並允許關閉
    ' Private _cancelRequested As Boolean = False        ' ESC 全域中斷旗標: Tab1/Tab2/Tab3 共用，按 ESC 立刻設 True，各操作在 Yield 點檢查 (2026/04/10 by simon&claude&gemini: 全域改用 CancellationTokenSource 發送取消信號，取代布林旗標)
    ' Private _cacheSnifferCts As New System.Threading.CancellationTokenSource  ' B4 CacheSniffer 取消令牌，FormClosing 時呼叫 Cancel()
    Private _cts As CancellationTokenSource             ' ✅ 2026/04/10: 導入現代化非同步中斷信號源作ESC中斷取代布林旗標

    ' ── 全域勾選狀態變數 (by Gemini, 2026/04/10: 優化效能，避免頻繁讀取 UI) ──
    '2026/3/10重構時停止使用全域變數來記錄遞迴過程中的資料, 改用傳遞參數以避免多線程或重入呼叫時資料被改寫的問題
    'Private _intTotalMailCount As Integer              ' 在遞迴中, 記錄點選資料夾內的所有郵件總數, 不要被遞迴呼叫改變數量
    'Private _intProcessedCount As Integer              ' 在遞迴中, 加總已處理的郵件總數, 不要被遞迴呼叫改變數量
    Private _showAllFolders As Boolean = False
    Private _includeSubTab2 As Boolean = False
    Private _includeSubTab3 As Boolean = False
    Private _lastTvMousePoint As Point = Point.Empty    ' by Gemini 3.1 Pro, 2026/04/26: 拆分 TreeView 與 ListView 的全域座標紀錄變數，避免互相干擾
    Private _lastLvMousePoint As Point = Point.Empty    ' by Gemini 3.1 Pro, 2026/04/26: 拆分 TreeView 與 ListView 的全域座標紀錄變數，避免互相干擾
    Private _lastHoveredPointIndex As Integer = -1      ' 記住上一個 hover 的點，-1 表示沒有
    'Private _lastHoveredTreeNode As TreeNode = Nothing ' 2026/5/14 by simon/Gemini: 將mouse hover作成內建功能
    Private _lastHoveredLvItem As ListViewItem = Nothing

    Private _inSizeMove As Boolean = False              ' 2026/06/19 by Simon/Claude: 是否處於拖曳 size/move modal loop；供 Form1_Resize 區分「拖曳縮放」與「瞬間頂天/最大化」
    Private _isForceRefreshing As Boolean = False       ' ✅ 2026/05/31 新增：F5 強制更新旗標，指示底層完全繞過 SSD 快取
    Private _isResizingLv As Boolean = False            ' ✅ 2026/05/09 by Gemini 3 Flash: 用於在欄位縮放期間暫停 OwnerDraw 繪製，消除 Reflow 殘影
    Private _lvResizePending As ListView = Nothing
    Private _lvResizeTimer As New Forms.Timer() With {.Interval = 100}

    ' [新增ProgressBar歷史紀錄 2026/4/2, by Gemini]
    Private Const MAX_HISTORY_COUNT As Integer = 100
    Private WithEvents HistoryListBox As ListBox
    Private _historyHoverIndex As Integer = -1
    Private _historyPopup As ToolStripDropDown
    Private _statusHistory As New List(Of StatusHistoryItem)(1024)

    Private Structure StatusHistoryItem
        Dim Time As DateTime
        Dim Message As String
        Dim Source As String
    End Structure
    Friend NotInheritable Class ThrottleFreq
        ' 2026/04/16 by Simon/Claude: 統一管理 SmartThrottle 的讓出間隔常數，取代散落的 100ms 魔術數字
        '   Hii (100ms) ：高頻迴圈，如 GetTable 掃郵件 (Tab2/Tab3)、Tab2 郵件總數預計算
        '   Mid (200ms) ：中頻迴圈，如 ComputeFolderSize 右鍵大小計算
        '   Low (300ms) ：低頻迴圈，如 RenewCache Phase2/3，每次操作 ~0.5ms，300ms 約每 600 個資料夾讓出一次
        Public Const Hii As Integer = 100   ' 高頻更新：適用於單純資料計算或記憶體操作
        Public Const Mid As Integer = 200   ' 中等更新：一般進度更新
        Public Const Low As Integer = 300   ' 低頻慢速：極耗時的附件掃描，不需過度更新
        Private Sub New() : End Sub         ' 防止被實例化
    End Class

    Private _fontDefault As New Font("Microsoft Jhenghei", 10.0F, _fontRegular, GraphicsUnit.Point, 0)
    Private _fontHeader As New Font("Microsoft Jhenghei", 10.0F, _fontBold, GraphicsUnit.Point, 0)
    Private _fontRegular = FontStyle.Regular
    Private _fontBold = FontStyle.Bold
    Private _fontItalic = FontStyle.Italic
#End Region

#Region "■ 02 Form 生命週期 & 外觀初始化"
#Region "  ├ 主畫面表單行為及事件"
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
#If DEBUG Then
        _isDebugMode = True ' by Gemini, 2026/04/01: 自動依據 VS 的編譯組態判斷是否為 Debug 模式
#Else
        _isDebugMode = False
#End If
        _dbg("開始") ' debugForm 開始計時
        Dim stopwatch As Stopwatch = Stopwatch.StartNew() : Cursor = Cursors.AppStarting  ' by Claude Sonnet 4.6, 2026/06/07

        InitLookAndFeel()       ' 設計程式外觀
        InitPgrsBarEvents()      ' 2026/04/02 by Gemini: 集中掛載 ProgressBar 互動事件 (取代 Handles 宣告)

        ' 2026/04/18 by Claude: Form 自身背景繪製的雙緩衝
        ' WS_EX_COMPOSITED (CreateParams, ■00) 管子控制項合成層；DoubleBuffered 管 Form 自身的 WM_PAINT。
        ' 兩者作用層次不同，互補無衝突。
        Me.DoubleBuffered = True
        Me.KeyPreview = True    ' ✅ 讓 Form 優先攔截 ESC，否則 ESC 會先被 TreeView/ListBox 等子控制項消耗
        Me.BringToFront()       ' 先將表單顯示後, 再以背景執行緒加入資料夾, 提高操作反應速度
        Me.Show()

        stopwatch.Stop() : Cursor = Cursors.Default ' 啟動完成, 停止計時, 顯示總共花費的時間
        PgrsBar1.Text = "啟動花費 " & stopwatch.Elapsed.TotalSeconds.ToString("0.00") & " 秒。"
        PgrsBar2.Text = ""
        _dbg("結束")

    End Sub
    Private Sub Form1_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        ' 2026/05/09 by Gemini 3 Flash: 整合最大化/還原偵測與節流機制
        ' 若 WindowState 改變，不直接呼叫，而是轉發給當前 ListView 的 Resize 處理器併入 100ms 節流。
        Static lastState As FormWindowState = Me.WindowState
        If Me.WindowState <> lastState Then
            Dim activeLv = GetCurrentLv()
            If activeLv IsNot Nothing Then HandleLvResize(activeLv, EventArgs.Empty)
            lastState = Me.WindowState
        End If

        ' 2026/06/19 by Simon/Claude Opus 4.8: 改以「是否在拖曳 size/move modal loop」(_inSizeMove) 判定同步策略 —
        '   拖曳中只搬位置(平滑)，完整貼齊由 ResizeEnd 收尾；
        '   非拖曳的尺寸變化(邊緣 double-click 垂直頂天 / Shift+Win+↑↓ / 最大化還原)瞬間完成、不進 loop 也無 ResizeEnd，
        '   故 _inSizeMove=False 時在此立即完整貼齊。
        If _inSizeMove Then SyncDebugFormMoveOnly() Else SyncDebugFormResize()

    End Sub
    Private Sub Form1_ResizeEnd(sender As Object, e As EventArgs) Handles Me.ResizeEnd
        ' 原本的 ListView1 寬度調整邏輯已移至 HandleLvResize 中，由 ListView 自行處理 Resize 事件
        ' Tab3 GroupBox3 顯示邏輯已改由 _pnlOptionsTab3.Resize 獨立處理，不再依賴 Form1_Resize

        If Me.Left < 0 Then Me.Left = 0 ' 2026/6/23 by simon: 防止 Form1 被拖到螢幕左邊界外

        ' 視窗縮放時同步 DebugForm — 2026/3/26 by Gemini
        SyncDebugFormResize()          ' 2026/06/19 by Simon/Claude: 拖曳(移動/縮放)結束後，一次完整貼齊右緣(含寬度/高度)
        _dbg("結束", sender.Width & "x" & sender.Height)
    End Sub
    Private Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        ' ── ESC 全域中斷 ──────────────────────────────────────────────
        ' KeyPreview=True 讓 Form 優先攔截 KeyDown，子控制項不會先吃掉 ESC
        If e.KeyCode = Keys.Escape Then
            ' 只有正在運算中才觸發中斷 (透過 WaitCursor 判定)
            If Cursor = Cursors.WaitCursor OrElse PgrsBar1.Text.StartsWith("正在") Then
                '_cancelRequested = True    ' 2026/04/10 by simon&claude&gemini: 全域改用 CancellationTokenSource 發送取消信號，取代布林旗標
                ' ✅ 發送標準取消信號
                If _cts IsNot Nothing AndAlso Not _cts.IsCancellationRequested Then _cts.Cancel()
                Cursor = Cursors.Default : PgrsBar1.Text = "由使用者中斷。"
                e.Handled = True
                e.SuppressKeyPress = True ' ✅ 防止事件繼續傳遞觸發 Lv2_KeyPress 等回上一頁邏輯
            End If

            ' 非運算中: 完全不設 _cancelRequested，不呼叫 _cts.Cancel()，直接放行。
            ' 讓 ListView/TreeView 等子控制項的原生 KeyDown/KeyPress 自己處理 ESC (例如 ListView2 返回年份視圖)。
            ' 2026/04/11 by Claude: 修正舊版非運算中按 ESC 仍呼叫 _cts.Cancel() 的 bug，會汙染 token，導致下一個操作取得的 cToken 已經是 cancelled 狀態。
        ElseIf e.KeyCode = Keys.F1 Then
            ' ── F1 側邊欄切換 ────────────────────────────────────────────────
            ' by Gemini 3 Flash, 2026/05/09: 讓使用者在任何地方按 F1 都能切換側邊欄收合/恢復
            Dim sc = GetCurrentSplitter()
            If sc IsNot Nothing Then
                SplitterToggle(sc)
                e.Handled = True
                e.SuppressKeyPress = True
            End If
        End If

    End Sub
    Private Async Sub Form1_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        _dbg("開始")

        ' by Gemini, 2026/04/01: 利用背景躲藏時間，預先載入其他 Tab 的 UI 與目錄樹，實現「切換瞬間無感」的流暢體驗
        ' 讓第一頁先穩穩地顯示出來，不要與使用者剛啟動後的第一波對 TreeView1 的操作搶資源
        InitMapiNamespace()
        InitDatabase()          ' by Gemini, 2026/04/06: 初始化 SQLite 快取資料庫
        CheckRDO.Checked = True ' added by simon 2026/6/25, to init RDO by default

        If _isDebugMode Then    ' by Gemini, 2026/04/01: 如果是 debug mode，就顯示 debugForm跟 debug button
            CheckDebug.Visible = True
            ' ✅ 2026/03/30 by Gemini: 改用 BeginInvoke 延遲啟動，避免 Load 期間同步觸發事件造成 UI 卡頓或 Handle 競爭, 移除原本導致Exception 的Task.Run 呼叫
            Me.BeginInvoke(Sub() CheckDebug.Checked = True) ' Memo: 這裡設成True 就會預設開啟 DebugForm，False 不開啟，設計階段方便debug
        End If

        ' by Gemini, 2026/04/05: 將表單移動與縮放事件改為 AddHandler，保持類別簡潔
        'AddHandler Me.Resize, Sub() SyncDebugFormResize()      ' 改善最大化時的lv column resize 效能, 移回到自己的 Resize 事件裡處理, 2026/5/9 by simon
        AddHandler Me.Move, Sub() SyncDebugFormMoveOnly()       ' 2026/06/19 by Simon/Claude: 拖曳移動只搬位置，完整貼齊延到 ResizeEnd
        Await Task.Yield()

        ' 2026/5/7 by Claude, 在HandleLvResize 加節流, 拖動過程中完全不重算欄寬，停手後才算一次
        AddHandler _lvResizeTimer.Tick, Sub()
                                            _lvResizeTimer.Stop()
                                            If _lvResizePending IsNot Nothing Then CalculateLvColumnSize(_lvResizePending)
                                        End Sub

        ' ── 初始化全域狀態變更事件 (by Gemini, 2026/04/10) ──
        AddHandler CheckSubFolder2.CheckedChanged, Sub() _includeSubTab2 = CheckSubFolder2.Checked
        AddHandler CheckSubFolder3.CheckedChanged, Sub() _includeSubTab3 = CheckSubFolder3.Checked
        'AddHandler checkShowAllFolders.CheckedChanged, Sub() _showAllFolders = checkShowAllFolders.Checked
        AddHandler OKiLikeNoisy.CheckedChanged, Sub() _iLikeNoisy = OKiLikeNoisy.Checked        ' ✅ 2026/04/12 by simon: 改為動態掛載過濾噪音開關事件
        Await Task.Yield()

        ' 2024/5/17, PST檔太多, 啟動速度愈來愈差, 全部重寫. 依照20年前的做法動態載入:
        ' 啟動時只載入第一層表皮, 若下層有subFolders=True 則暫塞一個假的":::" 讓它能顯示"+"加號表示還有子資料夾就好
        ' 只有當使用者點開 "+" 號展開節點時, 才真正去讀該項目的子資料夾, 不要一開始就花時間全讀
        ' 2026/04/13 by Simon/Claude: SimTree1 與 TreeView1 同步載入 PST 目錄樹
        LoadStoreToTreeView(_pstStoreList, SimTree1)
        GotoDefaultInbox(SimTree1)
        Await Task.Yield()

        ' Tab1 順利載入後，才開始載入 Tab2~Tab5 的 UI 與資料，避免一開始就全部載入造成卡頓
        ' 使用 TryToRelaxFor 確保使用者正在操作時會暫緩預載
        ' 依序初始化後面的標籤頁，拉出間隔避免卡住使用者剛進入畫面的第一波操作
        ' by Gemini, 2026/04/03: 增加載入各 Tab 之間的視覺區隔

        Dim delaySame As Integer = 100  ' 每個 Tab 之間的預載延遲，單位毫秒 (ms)，可以根據需要調整
        Await TryToRelaxFor(delaySame) : If Not _isTabInitialized(2) Then
            InitTab2UI() : _isTabInitialized(2) = True
            LoadStoreToTreeView(_pstStoreList, SimTree2) : GotoDefaultInbox(SimTree2)
        End If

        Await TryToRelaxFor(delaySame) : If Not _isTabInitialized(3) Then
            InitTab3UI() : _isTabInitialized(3) = True
            LoadStoreToTreeView(_pstStoreList, SimTree3) : GotoDefaultInbox(SimTree3)
        End If

        Await TryToRelaxFor(delaySame) : If Not _isTabInitialized(4) Then
            InitTab4UI() : _isTabInitialized(4) = True
            LoadStoreToTreeView(_pstStoreList, SimTree4) : GotoDefaultInbox(SimTree4)
        End If

        Await TryToRelaxFor(delaySame) : If Not _isTabInitialized(5) Then
            InitTab5UI() : _isTabInitialized(5) = True
            ' 2026/05/02 by Claude: Tab5 改用 SimTree5，對齊 Tab1~4
            LoadStoreToTreeView(_pstStoreList, SimTree5) : GotoDefaultInbox(SimTree5)
        End If
        _dbg("結束", "全部 Tab 背景載入完畢") ' by Gemini, 2026/04/11: UI 頂層 Level 0

        ' added by Gemini 3.1 Pro, 2026/04/12: 在所有的背景預載跟 Tab 初始化完成後，把一開始被蓋掉的啟動時間訊息重新顯示到 ProgressBar 上
        Dim firstMsgItem = _statusHistory.FirstOrDefault(Function(x) x.Source = "PB1")
        If Not String.IsNullOrEmpty(firstMsgItem.Message) Then PgrsBar1.Text = firstMsgItem.Message

        ' PROBE_BASICINFO_RDO  ↓↓↓ 整塊可刪 ↓↓↓ ----------------------------------------------------------
        ' 2026/07/02 by Claude: 無 GUI 自動觸發,命令列帶 /autoprobe 或 /autoprobe:StoreName|FolderName 才會啟動,正常啟動完全不受影響。
        '   ⚠ /autoprobedb 要先判斷(否則會被下面 StartsWith("/autoprobe") 誤吃)。
        Dim autoprobeDbArg = Environment.GetCommandLineArgs().FirstOrDefault(Function(a) a.StartsWith("/autoprobedb", StringComparison.OrdinalIgnoreCase))
        If autoprobeDbArg IsNot Nothing Then Dim unused5 = RunAutoProbeMailInfoDbRoundtrip(autoprobeDbArg)
        Dim autoprobeArg = Environment.GetCommandLineArgs().FirstOrDefault(Function(a) a.StartsWith("/autoprobe", StringComparison.OrdinalIgnoreCase) AndAlso Not a.StartsWith("/autoprobedb", StringComparison.OrdinalIgnoreCase))
        If autoprobeArg IsNot Nothing Then Dim unused3 = RunAutoProbeBasicInfoRdo(autoprobeArg)
        If Environment.GetCommandLineArgs().Any(Function(a) a.Equals("/liststores", StringComparison.OrdinalIgnoreCase)) Then Dim unused4 = RunListStoresDump()
        ' PROBE_BASICINFO_RDO  ↑↑↑ 整塊可刪 ↑↑↑ ----------------------------------------------------------

    End Sub
    Private Async Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        DebugForm.Visible = False
        Me.Visible = False

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

            If _olNS IsNot Nothing Then Marshal.FinalReleaseComObject(_olNS)
            If _olApp IsNot Nothing Then Marshal.FinalReleaseComObject(_olApp)
            ReleaseRdoSession() ' 2026/06/23 by Simon/Claude: 對稱釋放 _rdo2 獨立 session + store 快取
            _dbg("結束")
            Return
        End If

        ' 第一次攔截關閉事件：取消預設關閉，暫停 Timer，執行非同步儲存後再手動關閉
        e.Cancel = True
        _isClosing = True
        timerSaveCache.Enabled = False

        PgrsBar1.Text = "正在儲存快取，準備關閉程式..."
        Await SaveCachesToDB()
        ClearMemoryCachesCore() ' by Gemini, 2026/04/10: 關閉前顯式呼叫記憶體清理，確保資源釋放
        Me.Close()                  ' 觸發第二次進入此函式，執行上方釋放資源的區塊

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

        ' 設定程式標題 (如何設置版本號自動遞增 'myApp.MinorRevision += 1?? --> 在專案目錄維護一個ver.log 檔案裡面寫版本號, 每次編譯前就自動讀取, 加一再寫回)
        Dim strApp As String = My.Application.Info.DirectoryPath & "\" & My.Application.Info.ProductName & ".EXE"
        If My.Computer.FileSystem.FileExists(strApp) Then
            Dim infoReader As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(strApp)
            Dim modeStr As String = If(_isDebugMode, "(Debug)", "(Release)")
            Me.Text = $"Outlook Assistant - by Simon Lee Studio (build {infoReader.LastWriteTime:yyyy/MM/dd HH:mm:ss}) {modeStr}"
        End If

        ' ── 視窗位置與背景色 ──
        If Screen.FromControl(Me).Bounds.Height > 2560 Then
            Me.Top = Screen.FromControl(Me).Bounds.Height * 0.45                '如果在直立式的4K螢幕上啟動, 就把表單放在下半部往上移5%
            Me.Left = (Screen.FromControl(Me).Bounds.Width - Me.Width) * 0.45   '不管在什麼解析度的螢幕上啟動, 都把表單放在螢幕中央往左移5%
        End If
        Me.BackColor = ThemeColors.Gray95

        ' ── TabControl 字型與分頁名稱 ──
        Dim strTabName As String() = {"1.資料夾統計", "2.依日期統計", "3.尋找附件", "4.尋找系列郵件", "5.尋找重覆郵件", "6.Setting", "7.OST 解析"}
        For i As Integer = 0 To strTabName.Length - 1
            TabControl1.TabPages(i).Text = strTabName(i)
        Next
        TabControl1.Font = New Font(_fontDefault, _fontBold)
        TabControl1.Padding = New Point(12, 8)

        ' ── 容器化佈局與動態控制項掛載 ──
        ' by Gemini, 2026/04/01: 只初始化 Tab1，其餘 Tab 在切換時才載入 (Lazy Load)
        InitTab1UI()

        ' 2026/3/27 by Gemini: 修復 StatusStrip1 被 TabControl1 遮擋的問題
        StatusStrip1.SendToBack()
        TabControl1.BringToFront()

        DebugGroup.Visible = _isDebugMode   ' 只在 Debug 模式才顯示 DebugGroup 及相關控制項
        _dbg("結束")

    End Sub
    Private Sub InitTreeView(tv As SimTree)
        ' ---------------------------------------------------------------------------------------------------------
        ' ── 共用 Treeview 外觀設定 (by Gemini, 2026/04/01: 重構成接受單一參數，避免重複造輪子) ──
        ' ---------------------------------------------------------------------------------------------------------
        tv.Font = New Font(_fontDefault, _fontRegular)
        tv.BackColor = Color.White
        tv.ForeColor = SystemColors.InactiveCaptionText
        tv.Dock = DockStyle.Fill
        tv.Indent = 25                  ' 樹狀目錄縮排距離
        tv.EnableHoverHighlight = True  ' 啟用滑鼠懸停高亮
        tv.HoverColor = Form1.ThemeColors.MercuryGray

        ' 雙重緩衝區優化
        SendMessage(tv.Handle, TVM_SETEXTENDEDSTYLE, New IntPtr(TVS_EX_DOUBLEBUFFER), New IntPtr(TVS_EX_DOUBLEBUFFER))

        AddHandler tv.GotFocus, AddressOf HandleTvGotFocus  ' 2026/05/30, added by Simon/Claude
        AddHandler tv.BeforeExpand, AddressOf LoadSubFolderToTreeView
        AddHandler tv.MouseMove, AddressOf HandleTvMouseHover
        AddHandler tv.MouseLeave, AddressOf HandleTvMouseHover
        AddHandler tv.KeyPress, AddressOf HandleTvKeyPress

    End Sub
    Private Sub InitListView(lv As ListView)
        ' ---------------------------------------------------------------------------------------------------------
        ' ── 共用 Listview 外觀設定 (by Gemini, 2026/04/01: 重構成接受單一參數，避免重複造輪子) ──
        ' ---------------------------------------------------------------------------------------------------------
        lv.Font = New Font(_fontDefault, _fontRegular)
        lv.GridLines = False
        lv.View = Forms.View.Details
        lv.FullRowSelect = True
        lv.BringToFront()
        lv.Anchor = AnchorStyles.None
        lv.Dock = DockStyle.Fill
        lv.Cursor = Cursors.Default

        ' 雙重緩衝區優化
        SendMessage(lv.Handle, LVM_SETEXTENDEDLISTVIEWSTYLE, New IntPtr(LVS_EX_DOUBLEBUFFER), New IntPtr(LVS_EX_DOUBLEBUFFER))

        ' by Gemini, 2026/04/10: 使用反射開啟隱藏的 DoubleBuffered 屬性，解決虛擬模式下的重繪閃爍問題
        GetType(ListView).GetProperty("DoubleBuffered", System.Reflection.BindingFlags.Instance Or System.Reflection.BindingFlags.NonPublic).SetValue(lv, True, Nothing)

        ' ✅ 2026/04/21 by Gemini 3.0 flash: 統一掛載共通互動邏輯 (鍵盤開啟、雙擊開啟、路徑同步)
        ' 📢 僅針對 Tab3 (ListView3) 與 Tab4 (Listview4) 開啟，其餘 Tab 保持原有獨立邏輯
        ' 📢 2026/05/03 by Gemini 3.1 Pro: 加入 ListView5 共享共通邏輯
        If lv.Name = "ListView3" OrElse lv.Name = "Listview4" OrElse lv.Name = "ListView5" Then
            AddHandler lv.DrawColumnHeader, Sub(s, ev) ev.DrawDefault = True    ' 統一表頭繪製, 2026/4/26 by Gemini
            AddHandler lv.DrawItem, AddressOf HandleLv3Lv4Lv5_DrawItem          ' 統一項目背景繪製, 2026/05/05 by Gemini 3 Flash
            AddHandler lv.DrawSubItem, AddressOf HandleLv3Lv4Lv5_DrawSubItem    ' 統一 SubItem 繪製 (處理 Hover 變色與對齊), 2026/4/26 by Gemini

            AddHandler lv.SelectedIndexChanged, AddressOf ShowLv3Lv4Lv5PathToPgrsBar    ' 路徑更新邏輯統一由 ShowLv3Lv4Lv5PathToPgrsBar 接管

            ' AddHandler lv.KeyPress, AddressOf HandleLv3Lv4Lv5_KeyPress        ' 2026/4/22 by Gemini, 整合到KeyDown事件裡了
            AddHandler lv.KeyDown, AddressOf HandleLv3Lv4Lv5_KeyDown            ' 整合：共通快捷鍵 (ESC 回歸聚焦, Ctrl+A)
            AddHandler lv.MouseClick, AddressOf HandleLv3Lv4Lv5_MouseClick      ' 整合：單擊左鍵複製與點擊顯示路徑
            AddHandler lv.MouseDoubleClick, AddressOf HandleLv3Lv4Lv5_DoubleClick
            AddHandler lv.MouseDown, AddressOf HandleLv3Lv4Lv5_MouseDown        ' 2026/06/14 by Simon/Claude Opus 4.8: 右鍵先選取

            ' 2026/06/14 by Simon/Claude Opus 4.8: 建立 Lv3/4/5 共用刷新右鍵選單 (須在 InitListView(ListView3/4/5) 之前)
            InitLv3Lv4Lv5ContextMenu()
            lv.ContextMenuStrip = ctxMenuLv3Lv4Lv5
        End If

        AddHandler lv.GotFocus, AddressOf HandleLvGotFocus
        AddHandler lv.Resize, AddressOf HandleLvResize              ' 2026/04/01 by Gemini: 加入共用自動縮放事件
        ' AddHandler lv.KeyPress, AddressOf HandleListViewKeyPress  ' 2026/04/16 by Gemini 3.1 Pro: 已拆分至各 ListView 獨立處理
        AddHandler lv.MouseMove, AddressOf HandleLvMouseHover
        AddHandler lv.MouseLeave, AddressOf HandleLvMouseHover

    End Sub
    Private Sub InitSplitter(scnr As SplitContainer)
        scnr.Panel1MinSize = 0
        scnr.TabStop = False
        AddHandler scnr.MouseMove, Sub(s, ev) DirectCast(s, SplitContainer).Cursor = Cursors.SizeWE
        AddHandler scnr.MouseLeave, Sub(s, ev) DirectCast(s, SplitContainer).Cursor = Cursors.Default
        AddHandler scnr.MouseDown, AddressOf HandleSplitterMouseDown
    End Sub
    Private Sub InitPgrsBarEvents()
        ' ── ProgressBar 歷史紀錄 (by Gemini, 2026/04/02) ──
        ''' <summary>
        ''' 集中初始化 ProgressBar1 與 ProgressBar2 的互動事件 (TextChanged, Click, Hover)
        ''' 2026/04/02 by Gemini
        ''' </summary>
        _dbg("開始")

        ' 1. 文字變更紀錄
        AddHandler PgrsBar1.TextChanged, Sub() AppendStatusHistory(PgrsBar1.Text, "PB1")
        AddHandler PgrsBar2.TextChanged, Sub() AppendStatusHistory(PgrsBar2.Text, "PB2")

        ' 2. 點擊彈出歷史選單
        AddHandler PgrsBar1.Click, Sub() ShowHistoryPopup("PB1", PgrsBar1)
        AddHandler PgrsBar2.Click, Sub() ShowHistoryPopup("PB2", PgrsBar2)

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
        AddHandler PgrsBar1.MouseEnter, hoverIn
        AddHandler PgrsBar2.MouseEnter, hoverIn
        AddHandler PgrsBar1.MouseLeave, hoverOut
        AddHandler PgrsBar2.MouseLeave, hoverOut
        _dbg("結束")

    End Sub

    Private Sub InitTab1UI()
        _dbg("開始")

        ' 2026/04/13 by Simon/Claude: SimTree1 由設計工具建立，位置與層級已在 Designer.vb 設定
        ' 不需要 Controls.Add / BringToFront，直接 InitTreeView 初始化樣式即可
        InitTreeView(SimTree1)
        InitListView(ListView1)
        InitSplitter(SplitContainer1)

        ' 2026/04/13 v2: 移除「所屬父資料夾」欄，回歸 5 欄 (該欄內容永遠等於群組標題行，元余)
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

        ' by Gemini 3.0 Flash, 2026/04/20: 置中日期欄位 (Index 2)
        ' ✅ 2026/04/21: 因為在 Tab1 初始化的最後呼叫，需確認 ListView3 是否已初始化欄位
        If ListView3.Columns.Count > 2 Then ListView3.Columns(2).TextAlign = HorizontalAlignment.Center

        ' 2026/04/13 by Simon/Claude: OwnerDraw=True 讓群組標題行 / 合計列的 BackColor
        ' 在 hover / select 狀態下不被 OS 覆蓋 (DrawSubItem handler 在 Form1_MainTabs.vb)
        ListView1.OwnerDraw = True

        ctxMenuLv1 = New ContextMenuStrip()
        ctxMenuLv1.Items.Add("進入資料夾 (&E)", Nothing, Sub(sender, e) EnterSelectedFolder(ListView1.SelectedItems(0)))
        ctxMenuLv1.Items.Add("統計資料夾大小 (&C)", Nothing, AddressOf ComputeFolderSize)
        _isTabInitialized(1) = True
        _dbg("結束")

    End Sub
    Private Sub InitTab2UI()
        ' ── Tab2 (日期統計) 佈局重構 (2026/3/27 by Gemini) ──
        _dbg("開始")

        InitTreeView(SimTree2)
        InitListView(ListView2)
        InitSplitter(SplitContainer2)
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

        _dbg("結束")

    End Sub
    Private Sub InitTab3UI()
        ' ── Tab3 (尋找附件) 佈局優化 ──
        _dbg("開始")
        InitTreeView(SimTree3)
        InitListView(ListView3)
        InitSplitter(SplitContainer3)
        ' 建立頂部面板，將所有原本散落在 Panel2 的搜尋控制項集中
        'Dim layoutPanel3 As Panel
        layoutPanel3 = New Panel With {.Dock = DockStyle.Top,
                                          .Height = 115,
                                          .BackColor = ThemeColors.Gray95,
                                          .Font = New Font(_fontDefault, _fontRegular)}

        ' 將原本 Panel2 中的控制項移入新增的 layoutPanel3
        ' 這些控制項原本的 Location 已經適合在 Top Panel 中運作
        layoutPanel3.SendToBack() ' 2026/3/27 by Gemini: 正確設定 Dock 計算順序
        layoutPanel3.Controls.Add(GroupBox1)
        layoutPanel3.Controls.Add(GroupBox2)
        layoutPanel3.Controls.Add(GroupBox3)
        layoutPanel3.Controls.Add(Button3)
        layoutPanel3.Controls.Add(CheckSubFolder3)
        SplitContainer3.Panel2.Controls.Add(layoutPanel3)

        ' 2026/04/05 by Gemini: 優化顯示邏輯「純淨版」
        ' 改用面板自身的 Resize 事件與 Lambda 運算，不需類別變數。
        ' 這樣無論是調整視窗還是隱藏側邊欄，GroupBox3 都會依據「右側實際可用空間 (820px)」決定顯現與否。
        AddHandler layoutPanel3.Resize, Sub() GroupBox3.Visible = layoutPanel3.Width >= 820
        GroupBox3.Visible = layoutPanel3.Width >= 820

        ' ── Button3 樣式 ──
        Button3.FlatStyle = FlatStyle.System
        Button3.FlatAppearance.BorderColor = ThemeColors.Brand_Blue
        Button3.FlatAppearance.MouseOverBackColor = ThemeColors.MercuryGray
        Button3.ForeColor = ThemeColors.Brand_Blue
        Button3.BringToFront()
        CheckSubFolder3.BringToFront()

        ' ── 2026/03/28 by Gemini: 對齊邏輯優化 ──
        CheckSubFolder3.CheckAlign = ContentAlignment.MiddleLeft
        CheckSubFolder3.TextAlign = ContentAlignment.MiddleLeft
        CheckSubFolder3.AutoSize = True                                                 ' 1. 開啟 AutoSize 解決勾選框與文字「離得太遠」的問題 (寬度會自動縮短到剛好)
        CheckSubFolder3.Anchor = AnchorStyles.Top Or AnchorStyles.Left                  ' 2. 清除 Anchor 避免設計時的自動定位干擾手動計算，之後再重設為右側關聯
        'CheckSubFolder3.Left = (Button3.Left + Button3.Width) - CheckSubFolder3.Width  ' 3. 重新計算右側對齊 (會在 AutoSize 完後的正確 Width 基礎上計算)
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
                                                     CalculateLvColumnSize(ListView3) ' 加入自動縮放，依勾選狀態動態隱藏/顯示欄位
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

        ' ── ListView3: 搜尋結果欄位定義 ──
        With ListView3
            .Columns.Clear()
            .Columns.Add("主旨", "主旨", CInt(ListView3.Width * 0.45))
            .Columns.Add("郵件大小", "郵件大小", CInt(ListView3.Width * 0.13)) : .Columns(1).TextAlign = HorizontalAlignment.Right
            .Columns.Add("收到日期", "收到日期", CInt(ListView3.Width * 0.17)) : .Columns(2).TextAlign = HorizontalAlignment.Center
            .Columns.Add("寄件者", "寄件者", CInt(ListView3.Width * 0.22))
            .Columns.Add("附件個數", "附件個數", 0) : .Columns(4).TextAlign = HorizontalAlignment.Center ' [4] by Gemini 3 Flash, 2026/05/06: 初始隱藏
            .Columns.Add("EntryID", "EntryID", 0)                                                    ' [5] by Gemini 3 Flash, 2026/05/06: 初始隱藏
            .OwnerDraw = True ' by Gemini 3 Flash, 2026/04/26: 開啟 OwnerDraw 以在 VirtualMode 下實作流暢的 MouseHover 效果
        End With

        _dbg("結束")

    End Sub
    Private Sub InitTab4UI()
        _dbg("開始")

        InitTreeView(SimTree4)
        SimTree4.HideHorizontalScrollBar = True ' by Gemini 3.0 Flash, 2026/04/21: 隱藏水平捲軸並防止位移
        InitListView(Listview4)
        InitSplitter(SplitContainer4)

        ' 1. 設定 SimTree4 (左側目錄樹選取器)
        SplitContainer4.Panel1.Controls.Clear()
        With SimTree4
            .Parent = SplitContainer4.Panel1
            .Dock = DockStyle.Fill
            .Visible = True
            .BringToFront()
        End With

        ' 2026/05/29 by Simon/Claude: Phase 1 — Lv4Topic 與 SimTree4 共用 Panel1，Visible 切換顯示
        ' 尺寸對齊 SimTree4，Anchor 取代 Dock 避免兩個 Fill 控制項互相衝突
        Lv4Topic.SetBounds(SimTree4.Left, SimTree4.Top, SimTree4.Width, SimTree4.Height)
        Lv4Topic.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        Lv4Topic.Visible = False        ' 預設隱藏，資料夾模式優先
        Lv4Topic.VirtualMode = True     ' 💡 2026/05/30 by Gemini: 新增這行開啟虛擬模式

        ' 雙重緩衝區優化
        SendMessage(Lv4Topic.Handle, LVM_SETEXTENDEDLISTVIEWSTYLE, New IntPtr(LVS_EX_DOUBLEBUFFER), New IntPtr(LVS_EX_DOUBLEBUFFER))
        ' by Gemini, 2026/04/10: 使用反射開啟隱藏的 DoubleBuffered 屬性，解決虛擬模式下的重繪閃爍問題
        GetType(ListView).GetProperty("DoubleBuffered", System.Reflection.BindingFlags.Instance Or System.Reflection.BindingFlags.NonPublic).SetValue(Lv4Topic, True, Nothing)

        AddHandler Lv4Topic.Resize, AddressOf HandleLvResize
        AddHandler Lv4Topic.MouseMove, AddressOf HandleLvMouseHover
        AddHandler Lv4Topic.MouseLeave, AddressOf HandleLvMouseHover
        SplitContainer4.Panel1.Controls.Add(Lv4Topic)

        ' 2. 建立中間/右側的二階分欄 (Nested SplitContainer)
        ' ✅ 2026/04/20 by Gemini 2.0 Flash: 大手術！恢復為二欄佈局，徹底移除 TreeView4 與嵌套分欄
        Dim tp4 = TabControl1.TabPages(3)
        tp4.Controls.Clear()

        ' 1. 重設 SplitContainer4 (左: 樹, 右: 列表)
        SplitContainer4.Dock = DockStyle.Fill
        SplitContainer4.Orientation = Orientation.Vertical
        tp4.Controls.Add(SplitContainer4)

        ' 2. 左側：SimTree4 (直接放在 Panel1)
        SimTree4.Parent = SplitContainer4.Panel1
        SimTree4.Dock = DockStyle.Fill
        SimTree4.Nodes.Clear()

        ' 3. 右側：選項面板 + 列表
        ' 建立面板 (還原原始高度 80px)
        Dim layoutPanel4 = New Panel With {.Name = "layoutPanel4",
                                         .Height = 80,
                                         .Dock = DockStyle.Top,
                                         .BackColor = ThemeColors.Gray95}

        ButtonDeleteMail.Parent = layoutPanel4
        ButtonDeleteMail.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        ButtonDeleteMail.Size = New Size(102, 60)
        ButtonDeleteMail.Location = New Drawing.Point(layoutPanel4.Width - ButtonDeleteMail.Width - 10, 10)

        Button4.Parent = layoutPanel4
        Button4.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        Button4.Size = New Size(100, 60)
        Button4.Location = New Drawing.Point(ButtonDeleteMail.Left - Button4.Width - 10, 10)

        layoutPanel4.Parent = SplitContainer4.Panel2
        layoutPanel4.BringToFront()

        ' 列表掛載於 Panel2
        Listview4.Parent = SplitContainer4.Panel2
        Listview4.Dock = DockStyle.Fill
        Listview4.BringToFront()

        ' 4. 初始化寬度
        ' 設定 SplitterDistance 時必須確保控制項已在畫面上
        SplitContainer4.SplitterDistance = SimTree1.Width

        ' 再次強制清空與狀態初始化
        Listview4.Items.Clear()
        Listview4.OwnerDraw = True ' by Gemini 3.1 Pro, 2026/04/26: 開啟 OwnerDraw，徹底解決 ListView 有分組時 Hover 修改 BackColor 導致的 O(N) 嚴重卡頓

        ' ── Listview4: 系列郵件欄位定義 ──
        With Listview4
            .Columns.Clear()
            Dim lv4Names As String() = {"主旨", "郵件大小", "收到日期", "寄件者", "相似", "EntryID"}
            For Each n In lv4Names : .Columns.Add(n, n) : Next
            .Columns("主旨").Width = .Width * 0.4
            .Columns("郵件大小").Width = CInt(.Width * 0.13) : .Columns("郵件大小").TextAlign = HorizontalAlignment.Right
            .Columns("收到日期").Width = CInt(.Width * 0.17) : .Columns("收到日期").TextAlign = HorizontalAlignment.Center
            .Columns("寄件者").Width = .Width * 0.18
            .Columns("相似").Width = .Width * 0.08 : .Columns("相似").TextAlign = HorizontalAlignment.Center
            .Columns("EntryID").Width = 0   ' 隱藏，僅供 OpenMailByEntryID 使用
        End With
        _dbg("結束")

    End Sub
    Private Sub InitTab5UI()
        ' ---------------------------------------------------------------------------------------------------------
        ' ── Tab5 (尋找重複郵件) 佈局重構 (by Gemini 3 Flash, 2026/05/03) ──
        ' 加入 SimTree5 於左側，對齊 Tab1~4 操作行為，讓 Button5 搜尋範圍限定在選取的資料夾
        ' 原本直接掛在 TabPage5 的 TreeView5 + ListView5 改為用 SplitContainer5 左右分欄
        ' ---------------------------------------------------------------------------------------------------------
        _dbg("開始")

        ' 1. 初始化基礎控制項
        InitTreeView(SimTree5)
        InitListView(ListView5)
        ListView5.OwnerDraw = True ' by Gemini 3 Flash, 2026/05/05: 開啟 OwnerDraw，徹底解決 ListView5 懸停時 BackColor 被覆蓋與效能問題
        InitSplitter(SplitContainer5)
        'TreeView5.Visible = False ' 使用 SimTree5 取代，TreeView5 設為不可見

        ' 2. 準備右側選項面板 (layoutPanel5)
        Dim layoutPanel5 As New Panel With {.Name = "layoutPanel5",
                                           .Dock = DockStyle.Top,
                                           .Height = 110, ' 2026/05/05 by Gemini 3 Flash: 增加高度以容納 CheckSubFolder5
                                           .BackColor = ThemeColors.Gray95}
        ' 設定 RadioButtons 樣式與位置
        rbExactMatch.Text = "精確模式比對 (主旨+大小+時間+寄件者篩選)"
        rbExactMatch.Location = New Point(20, 15)
        rbExactMatch.Checked = True
        rbExactMatch.AutoSize = True

        rbFuzzyMatch.Text = "內文模糊比對 (SimHash + Jaccard 找相似重複)"
        rbFuzzyMatch.Location = New Point(20, 45)
        rbFuzzyMatch.AutoSize = True

        '' 設定 Label2 (搜尋模式顯示)
        'Label2.Text = "搜尋模式:"
        'Label2.Location = New Point(20, 0) ' 根據需求排好位置
        'Label2.AutoSize = True
        'Label2.Visible = True

        ' 設定 Button5 (開始掃描)
        Button5.Text = "開始掃描"
        Button5.Size = New Size(100, 60)
        Button5.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        Button5.Location = New Point(layoutPanel5.Width - Button5.Width - 10, 10)
        Button5.BringToFront()

        ' CheckSubFolder5：暫用程式建立，日後改設計工具放置。2026/05/05 by Claude
        CheckSubFolder5.Text = "含子資料夾"
        CheckSubFolder5.Checked = True
        CheckSubFolder5.AutoSize = True
        CheckSubFolder5.FlatStyle = FlatStyle.System
        CheckSubFolder5.Location = New Point(20, 75) ' 2026/05/05 by Gemini 3 Flash: 手動指定位置確保不被裁切
        AddHandler CheckSubFolder5.CheckedChanged, Sub() _includeSubTab5 = CheckSubFolder5.Checked

        ' 2026/06/17 by Simon/Claude Opus 4.8: Fuzzy 相似度檔位即時顯示 Label。
        '   拿掉 numberSimilarity 後 TrackBar1 只剩 1~5 刻度，使用者看不出「3=高 95%」，
        '   故綁 ValueChanged 顯示「檔位名 + 門檻%」(資料源 _fuzzyTierName / GetFuzzyTargetT，與掃描實際用值一致)。
        Dim lblFuzzyTier As New Label With {.Name = "lblFuzzyTier",
                                            .AutoSize = True,
                                            .Location = New Point(288, 80), ' TrackBar1(150,73 130x29) 右側、Button5 下方空區
                                            .Visible = False}
        AddHandler TrackBar1.ValueChanged, Sub()
                                               Dim v As Integer = Math.Clamp(TrackBar1.Value, 1, 5)
                                               lblFuzzyTier.Text = _fuzzyTierName(v) & " " & (GetFuzzyTargetT() * 100).ToString("0.##") & "%"
                                           End Sub
        ' 初始顯示一次(預設 TrackBar1.Value=4 → 極高 98%)，否則啟動到第一次拖動前是空的
        lblFuzzyTier.Text = _fuzzyTierName(Math.Clamp(TrackBar1.Value, 1, 5)) & " " & (GetFuzzyTargetT() * 100).ToString("0.##") & "%"

        AddHandler rbFuzzyMatch.CheckedChanged, Sub() CheckSubFolder5.Checked = Not rbFuzzyMatch.Checked
        AddHandler rbFuzzyMatch.CheckedChanged, Sub() lblFuzzyTier.Visible = rbFuzzyMatch.Checked
        AddHandler rbFuzzyMatch.CheckedChanged, Sub() TrackBar1.Visible = rbFuzzyMatch.Checked

        ' 3. 組裝右側面板
        layoutPanel5.Controls.AddRange({rbExactMatch, rbFuzzyMatch, CheckSubFolder5, TrackBar1, Button5, lblFuzzyTier})
        CheckSubFolder5.BringToFront() ' ✅ by Gemini 3 Flash, 2026/05/05: 顯式移至最前，防止被遮擋

        ' 4. 將控制項掛載到 SplitContainer5 的正確 Panel 中
        ' 左側：SimTree5 填滿 Panel1
        SplitContainer5.Panel1.Controls.Clear()
        SimTree5.Parent = SplitContainer5.Panel1
        SimTree5.Dock = DockStyle.Fill

        ' 右側 Panel2（上方按鈕列）+ ListView5（填滿剩餘空間）
        SplitContainer5.Panel2.Controls.Clear()
        layoutPanel5.Parent = SplitContainer5.Panel2
        ListView5.Parent = SplitContainer5.Panel2

        ' 設定 Dock 順序 (Z-Order)
        layoutPanel5.SendToBack()  ' 第一個 Dock=Top
        ListView5.Dock = DockStyle.Fill
        ListView5.BringToFront() ' 填滿剩餘空間

        ' 5. 顯式設定 SplitContainer 屬性與分割線
        SplitContainer5.Dock = DockStyle.Fill
        SplitContainer5.Panel2Collapsed = False
        SplitContainer5.SplitterDistance = 317 ' 對齊 Tab4

        ' 確保 SplitContainer5 在 TabPage5 中佔滿
        If Not TabPage5.Controls.Contains(SplitContainer5) Then
            TabPage5.Controls.Clear()
            TabPage5.Controls.Add(SplitContainer5)
        End If

        ' ListView5 欄位：加入「相似度」欄（Fuzzy 模式才有意義，Exact 模式顯示 100%）
        With ListView5
            .Columns.Clear()
            Dim lv5Names As String() = {"主旨", "郵件大小", "收到日期", "寄件者", "群組", "相似", "EntryID"}
            For Each n In lv5Names : .Columns.Add(n, n) : Next
            .Columns("主旨").Width = CInt(.Width * 0.34)
            .Columns("郵件大小").Width = CInt(.Width * 0.12) : .Columns("郵件大小").TextAlign = HorizontalAlignment.Right
            .Columns("收到日期").Width = CInt(.Width * 0.17) : .Columns("收到日期").TextAlign = HorizontalAlignment.Center
            .Columns("寄件者").Width = .Width * 0.17
            .Columns("群組").Width = CInt(.Width * 0.08) : .Columns("群組").TextAlign = HorizontalAlignment.Right
            .Columns("相似").Width = CInt(.Width * 0.08) : .Columns("相似").TextAlign = HorizontalAlignment.Center
            .Columns("EntryID").Width = 0 : .Columns("EntryID").TextAlign = HorizontalAlignment.Right ' 隱藏，僅供 OpenMailByEntryID 使用
        End With
        AddHandler ListView5.ColumnClick, AddressOf Lv5_ColumnClick

        _dbg("結束")

    End Sub
    Private Sub InitChart2()
        _dbg("開始", Chart2.Name)

        With Chart2
            ' 清除原有的設定
            .Series.Clear()
            .Legends.Clear()
            .ChartAreas.Clear()

            ' ── [新增/遷移] 設置 Chart 本身的外觀設定 by Gemini 3.0 Flash, 2026/04/17 ──
            .BorderlineDashStyle = ChartDashStyle.Solid
            .BorderlineColor = ThemeColors.AltoGray

            ' 設置抗鋸齒和文本抗鋸齒品質
            .AntiAliasing = AntiAliasingStyles.All
            .TextAntiAliasingQuality = TextAntiAliasingQuality.High

            ' 添加 Chart 的 Series
            Dim mailCount As New Series With {.Name = "郵件數量",
                                              .ChartType = SeriesChartType.Column, .Color = ThemeColors.barNormal}

            ' 添加 Chart 的 ChartArea
            ' by Gemini 3.0 Flash, 2026/04/17: 修正 BackColor 為 ThemeColors.bgColor 以確保與原本樣式統一
            Dim mailChart As New ChartArea With {.Name = "長條圖",
                                                 .BackColor = ThemeColors.bgColor, .BorderColor = Color.DarkGray}

            ' ── [遷移] 最大化 ChartArea 和 InnerPlotPosition by Gemini 3.0 Flash, 2026/04/17 ──
            With mailChart
                ' ChartArea.Position: ChartArea 在整個 Chart 控制項中的佔比 (單位: %)
                .Position = New ElementPosition(1, 1, 99, 99)

                ' ✅ 讓 ChartArea 幾乎填滿整個 Chart 控制項 (上下左右各留 1%) ' 預設約 Position(5,5,90,90)，壓縮到幾乎填滿整個 Chart 控制項
                ' ✅ InnerPlotPosition: ChartArea 內部長條圖實際繪製區的佔比
                With .InnerPlotPosition
                    ' Auto=True的話, Chart 會自動縮排給軸標籤留空，通常左側縮 10~15%
                    ' 改成 Auto=False 並手動指定，讓左側縮排符合實際 Y 軸標籤寬度
                    .Auto = False
                    .X = 8          ' 左側留 8% (給 Y 軸數字標籤)
                    .Y = 2          ' 上方留 2%
                    .Width = 90     ' 往右延伸 90%
                    .Height = 90    ' 往下延伸 90% (底部留 10% 給 X 軸標籤)
                End With

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

            ' ── [遷移] 圖例安全性檢查 by Gemini 3.0 Flash, 2026/04/17 ──
            ' InitChart2 會清除 Legends()，如果為空就不應該去存取，避免引發 ArgumentOutOfRangeException
            If .Legends.Count > 0 Then .Legends(0).Enabled = False
        End With
        _dbg("結束")

    End Sub
#End Region
#Region "  └ 全域輔助函數"
    Private Function OkayNowYouHaveToken() As CancellationToken
        ''' <summary>
        ''' 標準化取消訊號重置流程：中斷舊任務、釋放資源並產生新的 Token
        ''' 取消時以下三選一:
        ''' 彈射座椅 --> 開門下車 --> 開門帶行李走
        '''     Await Task.Yield() : cToken.ThrowIfCancellationRequested() (這個實測幾乎無效!!)
        '''     Await Task.Delay(1, cToken:=cToken) ' 這裡一定要保留至少 .delay(1) 才能讓 ESC 中斷生效 (simon, 2026/04/05)
        '''     If cToken.IsCancellationRequested Then Return New List(Of FolderBfsEntry)
        ''' </summary>
        If _cts IsNot Nothing Then
            Try : _cts.Cancel() : _cts.Dispose() : Catch : End Try
        End If
        ' 2026/04/10 by simon&claude&gemini: 全域改用 CancellationTokenSource 發送取消信號，取代布林旗標)
        '_cancelRequested = False  ' by Claude Opus, 2026/04/11: 重置舊旗標，否則 Layer3 的 GetMailCountAllOOM / GetFolderCountAllOOM 等仍在檢查它，ESC 後就永遠回傳 0 (全換token之後就可以不檢查這個旗標了，但為了保險起見還是重置它)
        _cts = New CancellationTokenSource()
        Return _cts.Token
    End Function
    Private Function SmartThrottle(sw As Stopwatch, cToken As CancellationToken, Optional intervalMs As Integer = ThrottleFreq.Mid, Optional onThrottled As System.Action = Nothing) As Task

        ' '' <summary>
        ''' 統一的節流讓出點，適用於所有需要在長時間迴圈中偶爾讓出 UI 執行權的情境
        ' '' </summary>
        '''
        ' ====================================================================================
        ' 2026/04/15 by Claude: 統一節流讓出點，取代各處散落的 swThrottle + Task.Delay(1) 組合
        ' 2026/04/19 by Gemini 3.0 flash: 加入 TimeBeginPeriod(1) 局部提速，縮短 Delay 偏差
        ' 2026/04/25 by Gemini 3.1 Pro: 修正 TimeEndPeriod 提早執行的問題，將等待邏輯拆分到內層 Async 函式
        '
        ' 設計說明:
        '   熱路徑 (sw < intervalMs): 直接回 Task.CompletedTask，零分配、零 await 開銷，編譯器不產生狀態機
        '   冷路徑 (sw >= intervalMs): 觸發 onThrottled (若有) → Restart sw → Task.Delay(1, cToken:=cToken)
        '     Task.Delay(1) 在 Windows 預設計時器下實際等 ~15.6ms，但每 100ms 才觸發一次，
        '     整體開銷比 < 16%，對使用者無感，且 15.6ms 已足以讓消息泵處理 ESC 的 WM_KEYDOWN。
        '     cToken 取消時 Task.Delay 拋 OperationCanceledException，讓呼叫端的 Catch OCE 接住。
        '   不使用 Application.DoEvents()：Async 函數中 DoEvents 引發再入 (reentrancy)，堆疊深度持續增長。
        '
        ' 呼叫端範例:
        '   Await SmartThrottle(swThrottle, cToken:=cToken)   ← OCE 由呼叫端 Catch 接住
        '   If cToken.IsCancellationRequested Then Return ...  ← 也可在呼叫端檢查取消狀態，視情況決定是否提前結束
        ' ====================================================================================

        If sw.ElapsedMilliseconds < intervalMs Then Return Task.CompletedTask

        Return SmartThrottleCore(sw, onThrottled, cToken:=cToken)

    End Function
    Private Async Function SmartThrottleCore(sw As Stopwatch, onThrottled As System.Action, cToken As CancellationToken) As Task

        onThrottled?.Invoke() : sw.Restart()

        ' ✅ 進入節流時暫時將系統計時器解析度拉高到 1ms，確保 Task.Delay(1) 真的只停 ~1ms 而非 15.6ms (by Gemini 3.0 flash, 2026/04/19)
        '     不使用 timeBeginPeriod(1)：Windows 10 為全域設定，拉高系統計時器解析度會增加 CPU 喚醒頻率，
        '     在桌機上額外耗電約 13~25W，筆電縮短續航 10~25%，代價不成比例。
        Try
            'Dim unused = TimeBeginPeriod(1)
            Await Task.Delay(1, cancellationToken:=cToken)
        Finally
            'Dim unused1 = TimeEndPeriod(1)
        End Try
    End Function
    Private Async Function PreciseDelay(baseDelayMs As Integer, Optional cToken As CancellationToken = Nothing) As Task
        ''' <summary>
        ''' 精準等待, 確保在 Windows 預設計時器解析度下也能盡可能接近指定的 baseDelayMs (而非誤差約 15.6ms 的 Task.Delay(baseDelayMs))
        ''' 2026/05/26 by simon
        ''' </summary>
        Try
            Dim unused = TimeBeginPeriod(1)
            Await Task.Delay(1, cancellationToken:=cToken)
        Finally
            Dim unused1 = TimeEndPeriod(1)
        End Try

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
#End Region
#End Region

#Region "■ 03 共用控制項行為"
#Region "  ├ 全域控制項事件"
    Private Sub TabControl1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl1.SelectedIndexChanged
        _isUserBusy = True
        _dbg("開始")

        Try ' by Gemini, 2026/04/01: 根據選定的分頁動態載入 UI 與資料 (Lazy Load UI)
            PgrsBar1.Text = "" : PgrsBar2.Text = ""
            Dim selectedTab As TabPage = CType(sender, TabControl).SelectedTab
            Dim tabIndex As Integer = TabControl1.SelectedIndex + 1 ' 產生 1, 2, 3, 4, 5 (Tab1~5)

            ' 如果切換到 Debug 分頁 (TabIndex >= 6)，不需要執行底下的陣列檢查，不需初始化 UI 直接離開
            If tabIndex > 5 Then
                ' 切換到 Setting 或 OST 解析頁時，更新快取統計顯示
                ' 2026/04/19 by Claude: 利用 tabIndex > 5 的早返回點順帶刷新 txtDatabaseStats
                If selectedTab.Text = "6.Setting" Then
                    ' ── 步驟 1: 第一次執行時，動態建立並配置 ListView ──
                    If ListView6.Items.Count = 0 Then
                        ListView6.BorderStyle = BorderStyle.None
                        ListView6.BackColor = Color.White
                        ListView6.Columns.Add("Item", 200)
                        ListView6.Columns.Add("Value", 120, HorizontalAlignment.Right)
                        ListView6.BringToFront()
                    End If
                    RefreshLv6DbStats() ' 現在是 Async Sub，呼叫後會立即返回不阻塞 Tab 切換

                ElseIf selectedTab.Text = "7.OST 解析" Then

                    InitTab7UI()   ' by Gemini 3.0 Flash, 2026/04/23: 設置控制項 Anchor 與事件

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
            Dim currentTree As SimTree = GetCurrentTv()
            If currentTree IsNot Nothing Then
                If currentTree.Nodes.Count = 0 Then
                    LoadStoreToTreeView(_pstStoreList, currentTree)
                    GotoDefaultInbox(currentTree)
                End If
                currentTree.Focus()
            End If

            ' ✅ by Gemini 3.1 Pro, 2026/05/27: 切換分頁時，確保當前 ListView 依最新 ClientSize 進行零開銷的初始欄寬計算
            Dim currentLv As ListView = GetCurrentLv()
            If currentLv IsNot Nothing Then CalculateLvColumnSize(currentLv)

            _dbg("結束")
        Finally
            _isUserBusy = False
        End Try

    End Sub
    Private Sub CheckShowAllFolders_CheckedChanged(sender As Object, e As EventArgs) Handles checkShowAllFolders.CheckedChanged
        ' by Gemini, 2026/03/30: 當切換顯示所有資料夾時，清空快取並標記所有 TreeView 為無效 (Nodes.Clear)
        ' by Gemini 3.0 Flash, 2026/04/17: 加入「路徑還原」機制，並整合全域旗標同步
        ' 職責：
        '   1. 同步 _showAllFolders 旗標
        '   2. 備份當前選取之路徑
        '   3. 清空所有舊節點 (並強制清理 SimTree 內部 stale 引用)
        '   4. 重新載入樹並嘗試恢復選取項目

        _showAllFolders = checkShowAllFolders.Checked
        _cacheFolderTree.Clear()
        ' by Claude Sonnet 4.6, 2026/04/25: 聚合快取（含子孫的加總）帶有模式語意，切換後必須清空，
        '   否則 BFS 會命中舊模式的 _cacheMailCountAll 直接剪枝，導致另一個模式的加總數字顯示錯誤。
        '   例如：False 模式下 TotalMailCount 不含行事曆，切到 True 模式後 BFS 直接用舊值 → 數字偏低。
        _cacheMailCountAll.Clear()
        _cacheFolderCountAll.Clear()
        _dbg("已切換顯示所有資料夾 (FolderTree/MailCountAll/FolderCountAll 快取已清空)", $"Mode: {_showAllFolders}")

        ' A. 備份 Tab1 完整展開與選取狀態 (路徑字串，Nodes.Clear 後仍有效)
        ' 2026/06/15 by Simon/Claude Opus 4.8: 改用 SaveTreeStateByPath 快照「所有」展開路徑，
        '   取代舊版僅備份單一選取節點 (oldPath/wasExpanded)，
        '   使切換 _showAllFolders 後已展開的節點保持展開、不被擅自收合。
        ' todo: 為什麼只備份 Tab1 的狀態？
        Dim st1State = SimTree1.SaveTreeStateByPath()

        ' B. 清理所有 TreeView
        For Each tv In GetAllTvList(Me)
            ' 強制重置 SimTree 內部選取清單，防止 Issue 2 發生 (重複物件重複統計)
            Dim st = TryCast(tv, SimTree)
            st?.ClearSelectedNodes()
            tv.Nodes.Clear()
        Next

        ' C. 重新載入 Tab1 的樹 (其餘 Tab 採 Lazy Load，切換時才重載)
        ' todo: 若這裡強制重載全部Lv會不會又有副作用?
        LoadStoreToTreeView(_pstStoreList, SimTree1)

        ' D. 還原所有展開路徑 + 選取，並觸發統計 (RestoreTreeState 對已消失資料夾天然容錯)
        '   若選取項目於切換後被過濾消失 (SelectedNode = Nothing)，Fallback 回收件匣
        ' by Gemini 3.5 Flash, 2026/05/21: 原採 SimTree1.SelectNode 還原單一焦點
        ' 2026/06/15 by Simon/Claude Opus 4.8: 改用 RestoreTreeState 一併還原全部展開狀態
        SimTree1.RestoreTreeState(st1State, selectAndFire:=True)
        If SimTree1.SelectedNode Is Nothing Then GotoDefaultInbox(SimTree1)

        PgrsBar2.Text = "全域資料夾過濾已變更，各頁面焦點已嘗試恢復。"

    End Sub
    Private Sub CheckRDO_CheckedChanged(sender As Object, e As EventArgs) Handles CheckRDO.CheckedChanged
        ' 用一個checkbox 動態決定是否載入Redemption
        If CheckRDO.Checked Then
            ' 2026/7/1 by simon, 所有RDO都已切換至獨立session的 _rdo2, 不再沿用 Outlook MAPI session, 讓原有的 _rdo 完全退役
            'If _rdo Is Nothing Then Dim unused = InitRdoSessionWithoutEULA()  ' 已知限制: 掛載在_olNS上的RDO, 卸載後就無法再重新載入第二次, 不會成功
            If _rdo2 Is Nothing Then Dim unused = InitRdoSessionWithoutEULA()
        Else
            ReleaseRdoSession()   ' 2026/06/23 by Simon/Claude Opus 4.8: 取消 RDO 時連 _rdo2 主讀取來源一起拆,避免 dispatcher 仍走 _rdo2
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
    Private Function GetCurrentSplitter() As SplitContainer
        ''' <summary>
        ''' 根據 TabControl1 的選擇索引，傳回當前分頁對應的 SplitContainer, by Gemini 3 Flash, 2026/05/09
        ''' </summary>
        Select Case TabControl1.SelectedIndex
            Case 0 : Return SplitContainer1
            Case 1 : Return SplitContainer2
            Case 2 : Return SplitContainer3
            Case 3 : Return SplitContainer4
            Case 4 : Return SplitContainer5
            Case Else : Return Nothing
        End Select
    End Function

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
#Region "  ├ 滑鼠 & 鍵盤操作事件"
    Private Async Sub HandleTvKeyDown(sender As Object, e As KeyEventArgs) Handles SimTree1.KeyDown, SimTree2.KeyDown, SimTree3.KeyDown, SimTree4.KeyDown, SimTree5.KeyDown
        ''' <summary>
        ''' 所有 SimTree 共用的鍵盤事件處理器，主要處理 F5 強制重整。
        ''' by Gemini 3.0 Flash, 2026/05/18: 原專屬 SimTree1_KeyDown 已移至 Form1.vb 的通用 HandleTvKeyDown() 中
        ''' </summary>
        Dim tv As SimTree = TryCast(sender, SimTree)
        If tv Is Nothing Then Return

        If e.KeyCode = Keys.F5 Then
            ' 特殊過濾：若為 SimTree4 且目前正處於話題搜尋結果模式，F5 代表重新掃描系列郵件，
            ' 此時應由 Tv4_KeyDown 專屬處理，此處直接 Return 提早退出。
            'If tv Is SimTree4 AndAlso _isTv4ResultMode Then Return
            ' 2026/05/29 by Simon/Claude: 拆分SimTree4的雙重模式, 讓SimTree4回復到純粹的資料夾樹行為

            e.Handled = True : e.SuppressKeyPress = True
            ForceTvRefresh(tv)
            If GetCurrentTv() Is SimTree1 Then Await ForceLv1Refresh()
            ' by Gemini 3.5 Flash, 2026/05/29: 限制 Await ForceLv1Refresh() 只有在當前是 Tab1 (SimTree1) 的時候才需要執行

        ElseIf e.KeyCode = Keys.Space Then  ' 按Space切換展開/收合
            ' by Claude Sonnet 4.6, 2026/05/19: 從 HandleTvKeyPress 移至此處
            ' KeyDown 比 KeyPress 更早觸發，搭配 SuppressKeyPress 可確實阻止系統捲動行為
            Dim node As TreeNode = tv.SelectedNode
            If node IsNot Nothing Then
                If node.IsExpanded Then node.Collapse() Else node.Expand()
                e.Handled = True : e.SuppressKeyPress = True  ' ✅ 阻止後續 KeyPress 觸發系統捲動
            End If

        ElseIf e.KeyCode = Keys.Enter Then
            If tv Is SimTree4 Then
                ' by Claude Sonnet 4.6, 2026/05/29: 將原先 Tv4_KeyDown 的 Enter 觸發搜尋移至此通用處理器，回復 SimTree4 單一掛載
                Button4.PerformClick()
                e.Handled = True : e.SuppressKeyPress = True
            End If
        End If
    End Sub
    Private Sub HandleTvMouseHover(sender As Object, e As EventArgs)
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
        ' 2026/5/14 by simon/Gemini: 試著將mouse hover作進SimTree自訂控制項內建功能, 但效果不太好, 跟原生長相有一滴滴不同
        ' ---------------------------------------------------------------

        '' 0. 如果滑鼠座標跟上一次完全一樣就直接離開，避免滑鼠沒有移動也一直觸發, 2026/4/19 by simon
        'Dim mouseE = TryCast(e, MouseEventArgs)
        'If mouseE IsNot Nothing Then
        '    If mouseE.Location = _lastTvMousePoint Then Return
        '    _lastTvMousePoint = mouseE.Location
        'End If

        'Dim tv As SimTree = CType(sender, SimTree)
        'Dim node As TreeNode = If(mouseE IsNot Nothing, tv.GetNodeAt(mouseE.Location), Nothing)
        'If node Is _lastHoveredTreeNode Then Return

        '' ── 還原上一個 hover 節點 (對稱結構第一部分) ──
        'If _lastHoveredTreeNode IsNot Nothing Then
        '    Dim sim As SimTree = TryCast(tv, SimTree)

        '    If sim IsNot Nothing AndAlso sim.SelectedNodes.Contains(_lastHoveredTreeNode) Then
        '        ' SimTree 選取節點: 根據焦點還原正確的選取色 (不能 Color.Empty)
        '        _lastHoveredTreeNode.BackColor = If(sim.Focused, SystemColors.Highlight, ThemeColors.AltoGray)
        '        _lastHoveredTreeNode.ForeColor = If(sim.Focused, SystemColors.HighlightText, SystemColors.InactiveCaptionText)
        '    Else
        '        _lastHoveredTreeNode.BackColor = Color.Empty
        '        _lastHoveredTreeNode.ForeColor = Color.Empty
        '    End If
        'End If

        '' ── 套用新 hover 色 (對稱結構第二部分) ──
        'If node IsNot Nothing Then
        '    Dim skipHover As Boolean = TypeOf tv Is SimTree AndAlso CType(tv, SimTree).SelectedNodes.Contains(node)
        '    If Not skipHover Then
        '        node.BackColor = ThemeColors.AltoGray
        '        node.ForeColor = SystemColors.InactiveCaptionText
        '    End If
        'End If
        '_lastHoveredTreeNode = node

    End Sub
    Private Sub HandleTvKeyPress(sender As Object, e As KeyPressEventArgs)
        ' 在這裡處理所有TreeView KeyPress 事件的程式碼
        _dbg("開始")

        If TypeOf sender Is TreeView Then
            If e.KeyChar = ChrW(Keys.Enter) Then
                sender.SelectedNode.Expand()            ' 按Enter展開下一層
                Select Case sender.Name
                    Case "SimTree1" : ListView1.Focus() : ListView1.Items(1).Selected = True
                    Case "SimTree2" : ListView2.Focus()
                End Select

            ElseIf e.KeyChar = ChrW(Keys.Escape) Then   ' 按ESC退回上一層
                ' ✅ by Gemini 3.0 flash, 2026/04/21: SimTree4 的 ESC 邏輯由 KeyDown 獨佔處理，此處予以排除避免衝突
                ' If sender.Name = "SimTree4" Then Return
                ' 2026/05/29 by Simon/Claude: 拆分SimTree4的雙重模式, 讓SimTree4回復到純粹的資料夾樹行為

                ' 2026/5/30 by simon, 取消這個ESC退回上一層的功能, 操作上不太直覺
                'If sender.SelectedNode IsNot Nothing AndAlso sender.SelectedNode.Parent IsNot Nothing Then
                '    sender.SelectedNode.Collapse() : sender.SelectedNode = sender.SelectedNode.Parent
                'End If
            End If
        End If

        _dbg("結束")

    End Sub
    Private Sub HandleLvMouseHover(sender As Object, e As EventArgs)
        ' by Gemini, 2026/04/03: 整合 MouseMove 與 MouseLeave 為單一維護點
        Dim listView As ListView = TryCast(sender, ListView)
        If listView Is Nothing Then Return

        ' 0. 如果滑鼠座標跟上一次完全一樣就直接離開，避免滑鼠沒有移動也一直觸發, 2026/4/19 by simon
        Dim mouseE = TryCast(e, MouseEventArgs)
        If mouseE IsNot Nothing Then
            If mouseE.Location = _lastLvMousePoint Then Return
            _lastLvMousePoint = mouseE.Location
        End If

        ' 1. 判斷目前的目標項目 (如果是 MouseLeave 則為 Nothing)
        Dim currentItem As ListViewItem = If(mouseE IsNot Nothing, listView.GetItemAt(mouseE.X, mouseE.Y), Nothing)

        ' 2. 檢查目標是否改變 (優化效能，若相同則不重繪)
        If currentItem Is _lastHoveredLvItem Then Return

        ' 3. 處理狀態轉變: 清除舊背景色並套用新色
        ' 2026/04/14 by Gemini 3.0 flash: 小步優化，OwnerDraw 模式下只 Invalidate 矩形，
        ' 絕對不改 BackColor 屬性，避免 WinForms ListView 在多項目時觸發 O(N) 版面重算拖慢效能。
        ' ⚠️ 注意: 這裡絕對不要加 .Refresh() 或 .BeginUpdate/EndUpdate()，讓 Windows 自然處理重繪，否則會導致嚴重的效能問題和 UI 卡頓。
        If listView.OwnerDraw Then
            If _lastHoveredLvItem IsNot Nothing Then listView.Invalidate(_lastHoveredLvItem.Bounds) ' 1. 讓「舊」項目重繪 (清除舊底色)
            If currentItem IsNot Nothing Then listView.Invalidate(currentItem.Bounds)                   ' 2. 讓「新」項目重繪 (畫上新底色)
            _lastHoveredLvItem = currentItem                                                          ' 3. 更新全域紀錄
            Return                                                                                      ' 4. 結束，因為 UI 已經標記要重繪了
        End If

        ' by Gemini, 2026/04/10: 虛擬模式下頻繁修改 BackColor 會觸發大量繪製導致閃爍
        ' 2026/04/26 by Gemini 3 Flash: 非 OwnerDraw 的虛擬模式仍需 Return 避免直接修改 BackColor 屬性報錯
        If listView.VirtualMode Then Return

        ' ----- 以下為非 OwnerDraw 模式 (例如 ListView2~5) -----
        ' 2026/04/14 by Simon/Claude: Tag=Nothing 的行 (群組標題 / 合計列) 有固定 BackColor，
        '   離開時要還原原色而非 Color.Empty；進入時也不套 hover 灰，保持原色不被蓋掉。
        ' 2026/6/5 by Simon/Claude: 再度簡化套用新 hover 色的邏輯，直接在一般列套灰色，標題/合計列保持原色不變
        If _lastHoveredLvItem IsNot Nothing Then _lastHoveredLvItem.BackColor = GetHeaderRowBackColor(_lastHoveredLvItem) ' 還原標題/合計列原色

        If currentItem IsNot Nothing Then currentItem.BackColor = ThemeColors.MercuryGray       ' 只對一般列套 hover 色
        _lastHoveredLvItem = currentItem

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
        Dim pt = Forms.Cursor.Position
        If _historyPopup IsNot Nothing AndAlso _historyPopup.Visible Then
            If Not _historyPopup.Bounds.Contains(pt) Then _historyPopup.Close()
        End If

    End Sub
    Private Sub HandleSplitterMouseDown(sender As Object, e As MouseEventArgs)
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
            If sc IsNot Nothing Then SplitterToggle(sc)
        End If
        _dbg("結束")

    End Sub
    Private Sub SplitterToggle(sc As SplitContainer)
        ''' <summary>
        ' 執行側邊欄的收合與恢復 (2026/05/09 by Gemini 3 Flash 抽離共用)
        ' 改變過程中觸發多次 DrawItem/DrawSubItem，造成殘影與卡頓。
        ' 原因: SplitterDistance 變動 → Panel2 Resize → CalculateLvColumnSize (欄寬改變)
        '        → ListView Invalidate × N 次中間狀態 → DrawSubItem × 項目數 × 欄位數
        ' 解法: WM_SETREDRAW=False 凍結所有子控制項重繪，改完後一次性 Invalidate + Update。

        '   2026/05/07 by Claude: 凍結 Panel2 的重繪，防止 OwnerDraw ListView 在 SplitterDistance
        '   2026/05/09 by Gemini 3 Flash: 在凍結前先對 Panel2 標記無效，確保舊區域在解凍後能被正確擦除背景
        ''' 2026/05/10 by Simon/Claude: 移除 WM_SETREDRAW，改用 BeginInvoke 延後重繪
        '''   根本原因：在 SplitterDistance 變更的同一 call stack 中呼叫 Invalidate/Update，
        '''   WinForms Layout 尚未完全結算新座標，Panel1 又未被凍結，導致舊像素殘留。
        '''   BeginInvoke 把重繪推入訊息佇列，保證 Layout 完全結算後才執行，徹底消除殘影。
        ''' </summary>
        If sc Is Nothing Then Return

        Try
            If sc.SplitterDistance > 20 Then    ' 臨界值 20px，如果大於此寬度則進行縮合
                sc.Tag = sc.SplitterDistance    ' 💡 記憶當前寬度在 Tag 屬性，以便下次恢復
                sc.SplitterDistance = 6         ' 縮合至 6px 觸控區
                _dbg("縮合側邊欄", $"{sc.Name} → 10px (原 {sc.Tag}px)")
            Else
                ' 💡 恢復寬度，若無紀錄則預設為 250px
                Dim prevDist As Integer = If(TypeOf sc.Tag Is Integer, DirectCast(sc.Tag, Integer), 250)
                If prevDist < 50 Then prevDist = 250    ' 防止恢復值過小
                sc.SplitterDistance = prevDist
                _dbg("恢復側邊欄", $"{sc.Name} → {prevDist}px")
            End If
        Finally
            '' 2026/05/09 by Gemini 3 Flash: 強化重繪指令，使用 Invalidate(True) 包含所有子控制項並強制背景重繪
            '' 2026/05/10 by Simon/Claude: 修復 Splitter 收合/展開後 OwnerDraw ListView 殘影
            ' BeginInvoke 確保 Layout 完全結算後才強制重繪全樹
            ' sc.Invalidate(True) 同時涵蓋 Panel1 (Tree) + Panel2 (ListView) 及其所有子控制項
            BeginInvoke(Sub()
                            sc.Invalidate(True)
                            sc.Update()         ' ← 改為 sc.Update() 確保整個 SplitContainer 同步重繪

                            ' OwnerDraw ListView 需要額外 RDW_ERASE 強制清除舊的 off-screen buffer
                            Dim lv = GetCurrentLv()
                            If lv IsNot Nothing Then RedrawWindow(lv.Handle, IntPtr.Zero, IntPtr.Zero,
                                                                  CUInt(RDW_INVALIDATE Or RDW_ERASE Or RDW_UPDATENOW Or RDW_ALLCHILDREN))
                        End Sub)
        End Try
    End Sub
    Private Sub HandleTvGotFocus(sender As Object, e As EventArgs)
        ' 當 SimTree 取得焦點時，若左側 Panel1 處於收合狀態，自動展開
        ' 適用情境：ESC 從 ListView 退回 SimTree、或任何其他讓 SimTree 得焦的操作
        ' 2026/05/30 by Simon/Claude
        Dim sc = GetCurrentSplitter()
        If sc IsNot Nothing AndAlso sc.SplitterDistance <= 20 Then SplitterToggle(sc)
    End Sub
#End Region
#Region "  ├ TreeView 導覽工具"
    ' [歷史變更與演進紀錄 by Gemini 3.5 Flash, 2026/05/21]
    ' =========================================================================
    ' 舊有的 FindNodeByPath、SelectNode、SelectNodeByPathRecursive 函數已被廢棄並刪除。
    ' 其核心職責與功能（包含 Lazy Load 的節點展開與狀態還原）已完全重構至自訂控制項 SimTree.vb
    ' 內建的 TryGetNode()、GetNodeIn() 與 SelectNode() 核心引擎中，以提升搜尋效能並達成控制項自治。
    ' -------------------------------------------------------------------------
    ' 過去的 Debug 與演進歷程備忘如下：
    ' 1. FindNodeByPath:
    '    - by Gemini 3.1 Pro, 2026/04/24: 遞迴尋找符合 targetPath 的 TreeNode，比對 Tag (Outlook.Folder) 的 FolderPath。
    '    - 2026/05/01 by Claude: 新增 searchOnlyExpanded 參數，控制是否只搜尋已展開的節點，優化性能和使用者體驗。
    '      searchOnlyExpanded：只搜已展開層（如果 =True 就只搜已展開的節點，=False則全部搜到底）。
    ' 2. SelectNodeByPathRecursive & SelectNode:
    '    - 2026/04/17 by Gemini: 若原本節點是展開的，還原時也必須主動展開，維持使用者體感的一致性。
    '    - 剪枝檢查：如果目標路徑開頭包含目前路徑，則進入子層搜尋，若未載入則呼叫 LoadSubFolderToTreeView。
    ' =========================================================================
    Private Async Sub GotoDefaultInbox(tv As SimTree)
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
            ' 舊版: tv.Nodes.Count - 1 = PST 個數 (通常=1)，只會檢查第一個子資料夾
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
    Private Async Function RefreshAllTreeViews() As Task

        ' ── UI 刷新工具 (by Gemini 3.0 flash, 2026/04/24) ──────────────────────
        ''' <summary>
        ''' 非同步刷新所有 SimTree，確保快取更新後的新資料夾結構能正確顯示。
        ''' 流程：清空節點 → 重新加載根節點 → 釋放 UI 執行緒。
        ''' </summary>
        _dbg("開始", "準備刷新所有 TreeView...")
        'PgrsBar1.Text = "正在刷新 UI 樹狀結構..."

        ' 讓出 UI 執行緒，確保進度文字能顯示
        Await Task.Yield()

        Dim trees() As SimTree = {SimTree1, SimTree2, SimTree3, SimTree4, SimTree5}
        Dim processedCount As Integer = 0

        For Each tv In trees
            If tv IsNot Nothing Then
                tv.BeginUpdate()
                Try
                    ' 1. 清空舊節點與 SimTree 內部選取快取
                    Dim st = TryCast(tv, SimTree)
                    st?.ClearSelectedNodes()
                    tv.Nodes.Clear()

                    ' 2. 重新加載 Store 根節點 (Lazy Load)
                    LoadStoreToTreeView(_pstStoreList, tv)
                    GotoDefaultInbox(tv)
                Finally
                    tv.EndUpdate()
                End Try

                processedCount += 1
                Await Task.Yield()  ' 每個 Tree 處理完後讓出 UI，防止大規模介面重繪導致短暫卡頓
            End If
        Next

        _dbg("結束", $"已刷新 {processedCount} 個 TreeView")
        'PgrsBar1.Text = "UI 刷新完成 ✔"
    End Function
    Private Sub ForceTvRefresh(tv As SimTree)
        ' ── F5 強制刷新整棵 SimTree ──────────────────────────────────────────────
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
        ' ─────────────────────────────────────────────────────────────────────
        ' 2026/05/25 by Simon/Claude: 再度重構使用呼叫simTree內部方法
        '       ① SaveTreeStateByPath（路徑字串快照）→
        '       ② 清快取 + 清節點 + 重建 root →
        '       ③ RestoreTreeState（按路徑重展開，觸發 LazyLoad 重讀 Outlook COM）→
        '       ④ Fallback
        '   若資料夾在 Outlook 已消失，重展開時天然不出現，不需另寫 diff 邏輯。
        '   2026/05/25 by Simon/Claude: 重構使用 SaveTreeStateByPath / RestoreTreeState
        ' ─────────────────────────────────────────────────────────────────────
        _dbg("開始", tv.Name)
        If _pstStoreList Is Nothing OrElse _pstStoreList.Count = 0 Then Return

        Dim sw As Stopwatch = Stopwatch.StartNew()   ' 2026/07/01 by Claude: F5 計時
        PgrsBar1.Text = $"F5: 重整 {tv.Name}..." : PgrsBar2.Text = ""
        _isUserBusy = True : Cursor = Cursors.WaitCursor

        Dim currentTvState = tv.SaveTreeStateByPath()   ' ① 快照狀態收集（展開路徑 + 選取路徑，Nodes.Clear 後仍有效）
        Try
            '_cacheFolderTree.Clear()                   ' ② 清快取，確保 GetSortedSubFolders 重讀 Outlook
            ClearMemoryCachesCore()                     ' 【修復關鍵 1】徹底清除所有快取，包含 _cacheFolderIDs，防止幽靈復活, 2026/6/1 by Simon/Gemini 3.1 Pro
            _isForceRefreshing = True                   ' 確保繞過 SSD 快取，強制打 COM (若原 codebase 沒加這行請務必補上), 2026/6/1 by Simon/Gemini 3.1 Pro
            tv.ClearSelectedNodes()
            tv.Nodes.Clear()

            LoadStoreToTreeView(_pstStoreList, tv)      ' ③ 重建 root 層
            tv.RestoreTreeState(currentTvState)         ' ④ 重展開 + 還原選取 + 觸發 AfterSelect
            If tv.SelectedNodes.Count = 0 Then GotoDefaultInbox(tv) ' ⑤ Fallback: 找不到舊選取時退回預設 Inbox

            PgrsBar1.Text = $"F5: {tv.Name} 重整完成，花費 {sw.Elapsed.TotalSeconds:0.00} 秒。" : PgrsBar2.Text = ""

        Catch ex As System.Exception
            _dbg("錯誤", ex.Message) : PgrsBar1.Text = $"F5 {tv.Name} 失敗: " & ex.Message
        Finally
            Cursor = Cursors.Default : _isUserBusy = False : _dbg("結束", tv.Name)
            _isForceRefreshing = False                  ' 關閉強制更新旗標, 2026/6/1 by Simon/Gemini 3.1 Pro
        End Try

    End Sub
    Private Function GetCurrentTv() As SimTree
        ''' <summary>
        ''' 根據 TabControl1 的選擇索引，判斷並傳回當前畫面上活動中的 TreeView/SimTree, by Gemini, 2026/03/30
        ''' </summary>
        ' 在需要觸發 AfterSelect 或其他操作時，能夠根據目前選中的 Tab 頁面，準確地獲取對應的 TreeView 控制項
        Select Case TabControl1.SelectedIndex
            Case 0 : Return SimTree1   ' 2026/04/13 by Simon/Claude: Tab1 改用 SimTree1
            Case 1 : Return SimTree2
            Case 2 : Return SimTree3
            Case 3 : Return SimTree4
            Case 4 : Return SimTree5   ' 2026/05/02 by Claude: Tab5 改用 SimTree5
            Case Else : Return Nothing
        End Select

    End Function
    Private Function GetAllTvList(container As Control) As List(Of SimTree)
        ''' <summary>
        ''' 遞迴搜尋容器內所有的 TreeView (含其衍生子類如 SimTree)
        ''' </summary>
        Dim list As New List(Of SimTree)(16) ' 預設容量 16，避免頻繁擴容
        For Each ctrl As Control In container.Controls
            If TypeOf ctrl Is SimTree Then list.Add(CType(ctrl, SimTree))
            If ctrl.HasChildren Then list.AddRange(GetAllTvList(ctrl))   ' 如果有子容器 (如 SplitContainer, TabControl, Panel)，繼續遞迴往下層掃描
        Next
        Return list
    End Function
    Private Function GetSelectedFolderPath(tv As SimTree) As String
        ''' <summary>
        ''' 取得指定 TreeView 選中節點的 Outlook 路徑 (優先從 Tag.FolderPath 讀取)
        ''' </summary>
        If tv Is Nothing Then Return ""

        Dim st = TryCast(tv, SimTree)
        Dim node As TreeNode = If(st IsNot Nothing, st.SelectedNode, tv.SelectedNode)
        If node Is Nothing Then Return ""

        Dim folder = TryCast(node.Tag, Folder)
        Return If(folder IsNot Nothing, SafeGetPath(folder), node.FullPath)
    End Function
#End Region
#Region "  ├ ListView 格式工具"
    Private Sub HandleLvGotFocus(sender As Object, e As EventArgs)
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
    Private Sub HandleLvResize(sender As Object, e As EventArgs)
        ''' <summary>
        ''' 處理所有 ListView 的 Resize 共用事件 (2026/04/01 by Gemini)
        ''' </summary>
        Dim lv As ListView = TryCast(sender, ListView)
        If lv Is Nothing Then Return

        ' 2026/05/09 by Gemini 3 Flash: 引入寬度比對防護
        ' 只有寬度實質改變時才啟動節流計時器。排除「高度改變」或「WindowState 切換中的無效訊息」。
        Static lastLvWidths As New Dictionary(Of String, Integer)
        If lastLvWidths.ContainsKey(lv.Name) AndAlso lastLvWidths(lv.Name) = lv.Width Then Return
        lastLvWidths(lv.Name) = lv.Width

        ' 在 Resize 邏輯中使用
        SendMessage(lv.Handle, WM_SETREDRAW, New IntPtr(0), IntPtr.Zero) ' 關閉重繪

        ' 2026/5/7 by Claude, 在HandleLvResize 加節流, 拖動過程中完全不重算欄寬，停手後才算一次
        _lvResizePending = lv
        _lvResizeTimer.Stop()
        _lvResizeTimer.Start()  ' 每次 Resize 重設計時，停止移動後 50ms 才真正重算

        SendMessage(lv.Handle, WM_SETREDRAW, New IntPtr(1), IntPtr.Zero) ' 開啟重繪
        lv.Invalidate() ' 強制刷新

    End Sub
    Private Sub CalculateLvColumnSize(lv As ListView)
        ''' <summary>
        ''' 定義各個 ListView 縮放時的欄位寬度比例
        ''' 2026/04/01 by Gemini
        ''' </summary>
        If lv.Columns.Count = 0 OrElse lv.Width <= 0 Then Return

        ' 2026/05/09 by Gemini 3 Flash: 終極防線 — 若寬度未變，絕不執行昂貴的 UI 重算與 _dbg 輸出
        Static lastProcessedWidths As New Dictionary(Of String, Integer)
        Dim value As Integer = Nothing
        If lastProcessedWidths.TryGetValue(lv.Name, value) AndAlso value = lv.Width Then Return

        _dbg("開始", lv.Name)
        _isResizingLv = True    ' ✅ 2026/05/09 by Gemini 3 Flash: 開啟旗標，暫停 DrawSubItem 繪製
        lv.BeginUpdate()        ' 開始更新，避免多次 Resize 造成的閃爍
        SendMessage(lv.Handle, WM_SETREDRAW, New IntPtr(0), IntPtr.Zero)    ' 在 Resize 邏輯中使用 關閉重繪
        Try
            Dim w As Integer = lv.ClientSize.Width  ' 使用 ClientSize 避免捲軸吃掉寬度
            If w <= 0 Then Exit Try

            ' 2026/05/09 by Gemini 3 Flash: 採用「計算與賦值分離」策略
            ' 先算出所有欄位的寬度清單，最後統一寫入，徹底消除 Header 逐個變寬的視覺差。
            Dim newWidths(lv.Columns.Count - 1) As Integer

            If lv Is ListView1 Then ' Tab1: 資料夾名稱 / 郵件數量 / 資料夾數量 / 郵件總計 / 大小
                ' 2026/04/13 v2: 移除「所屬父資料夾」欄，回歸 5 欄
                If lv.Columns.Count >= 5 Then
                    newWidths(1) = CInt(w * 0.15)
                    newWidths(2) = CInt(w * 0.15)
                    newWidths(3) = CInt(w * 0.15)
                    newWidths(4) = CInt(w * 0.188)
                    newWidths(0) = w - (newWidths(1) + newWidths(2) + newWidths(3) + newWidths(4)) - 5
                End If

            ElseIf lv Is ListView2 Then ' Tab2: 年度 / 郵件個數 / 空白欄位
                If lv.Columns.Count >= 2 Then
                    newWidths(0) = Math.Max(120, CInt(w * 0.3)) ' 第一欄(年度/月份)至少保底 120px
                    newWidths(1) = Math.Max(100, CInt(w * 0.2)) ' 第二欄(郵件數量)至少保底 100px
                    newWidths(2) = Math.Max(0, w - newWidths(0) - newWidths(1) - 5) ' 第三欄吸收所有剩餘空間
                End If

            ElseIf lv Is ListView3 Then ' Tab3: 郵件主旨 / 郵件大小 / 收到日期 / 寄件者 / 附件個數 / EntryID
                If lv.Columns.Count >= 6 Then
                    newWidths(1) = CInt(w * 0.15)    ' 郵件大小
                    newWidths(2) = CInt(w * 0.17)    ' 收到日期 (by Gemini 3 Flash, 2026/05/06: 統一 17%)
                    newWidths(3) = CInt(w * 0.18)    ' 寄件者
                    newWidths(5) = CInt(w * 0.01)    ' EntryID 極小保留

                    ' by Gemini 3 Flash, 2026/05/06: 實作連動邏輯 ——
                    ' 當使用者勾選「附件個數」或左側側邊欄收攏時，自動展開此欄位（寬度 60px 即可，顯示數字用）
                    Dim isLeftCollapsed As Boolean = (SplitContainer3.SplitterDistance < 50)
                    newWidths(4) = If(CheckAttCount.Checked OrElse isLeftCollapsed, 60, 0)
                    newWidths(0) = w - (newWidths(1) + newWidths(2) + newWidths(3) + newWidths(4) + newWidths(5)) - 5
                End If

            ElseIf lv Is Lv4Topic Then ' Tab4: 重複主旨 / 重複數量
                If lv.Columns.Count >= 1 Then
                    newWidths(1) = CInt(w * 0.2)        ' 重複數量
                    newWidths(0) = w - newWidths(1) - 24 ' 重複主旨
                End If
            ElseIf lv Is Listview4 Then ' Tab4: 主旨 / 大小 / 收到時間 / 寄件者 / 相似度 / EntryID
                If lv.Columns.Count >= 5 Then
                    newWidths(1) = CInt(w * 0.13)    ' 大小
                    newWidths(2) = CInt(w * 0.17)    ' 收到時間 (by Gemini 3 Flash, 2026/05/06: 統一改為 17%)
                    newWidths(3) = CInt(w * 0.18)    ' 寄件者
                    newWidths(4) = CInt(w * 0.08)    ' 相似度
                    newWidths(5) = CInt(w * 0.01)    ' EntryID 極小保留，避免 Resize 事件蓋掉 by Claude Sonnet 4.6, 2026/05/03
                    newWidths(0) = w - (newWidths(1) + newWidths(2) + newWidths(3) + newWidths(4) + newWidths(5)) - 5 ' by Claude Sonnet 4.6, 2026/05/03: 補上 EntryID 欄扣除
                End If

            ElseIf lv Is ListView5 Then ' Tab5: 主旨/大小/收到日期/寄件者/群組/相似/EntryID (7欄，比 LV4 多「群組」欄) by Claude Sonnet 4.6, 2026/05/03
                If lv.Columns.Count >= 7 Then
                    newWidths(1) = CInt(w * 0.12)    ' 郵件大小
                    newWidths(2) = CInt(w * 0.17)    ' 收到日期
                    newWidths(3) = CInt(w * 0.15)    ' 寄件者
                    newWidths(4) = CInt(w * 0.065)   ' 群組
                    newWidths(5) = CInt(w * 0.08)    ' 相似
                    newWidths(6) = CInt(w * 0.01)    ' EntryID 極小保留
                    newWidths(0) = w - (newWidths(1) + newWidths(2) + newWidths(3) + newWidths(4) + newWidths(5) + newWidths(6)) - 5
                End If
            Else
                ' 預設比例: 首欄固定40%，其餘均分
                If lv.Columns.Count > 0 Then
                    newWidths(0) = CInt(w * 0.4)
                    If lv.Columns.Count > 1 Then
                        Dim avgWidth As Integer = (w - newWidths(0) - 5) \ (lv.Columns.Count - 1)
                        For i As Integer = 1 To lv.Columns.Count - 1 : newWidths(i) = avgWidth : Next
                    End If
                End If
            End If

            ' ── 執行批量賦值 ──
            For i As Integer = 0 To lv.Columns.Count - 1
                ' 僅在數值改變時賦值，減少對 Header 的內部連動
                If lv.Columns(i).Width <> newWidths(i) Then lv.Columns(i).Width = newWidths(i)
            Next

            ' ✅ by Gemini 3.1 Pro, 2026/05/27: 將防線記錄移至賦值成功後，避免 ClientSize=0 時鎖死防線
            lastProcessedWidths(lv.Name) = lv.Width

            lv.Invalidate()         ' 強制刷新
        Finally
            SendMessage(lv.Handle, WM_SETREDRAW, New IntPtr(1), IntPtr.Zero) ' 開啟重繪
            lv.EndUpdate()
            _isResizingLv = False   ' ✅ 2026/05/09 by Gemini 3 Flash: 恢復繪製
        End Try
        _dbg("結束", lv.Name)
    End Sub
    Private Sub LviSelectAll(lv As ListView, Optional e As KeyEventArgs = Nothing)
        ' ---------------------------------------------------------------
        ' LviSelectAll — 共用的 ListView 全選輔助函數
        ' 整合自 Lv1_KeyDown / Lv2_KeyDown / HandleLv3Lv4Lv5_KeyDown 三處原本重複的 Ctrl+A 邏輯
        ' 2026/05/18 by Gemini 3.0 flash: 抽離共用的 ListView 全選邏輯 (支援一般模式與虛擬模式)
        '
        ' 一般模式 (ListView1 / ListView2 / Listview4 / ListView5 非虛擬部分):
        '   直接遍歷 Items 集合，逐一設為 Selected = True
        '   預分配容量設計來源：by Gemini 3 Flash, 2026/05/04
        '
        ' 虛擬模式 (ListView3 / 未來擴充):
        '   改用索引循環操作 SelectedIndices，不可枚舉 Items 集合否則例外
        '   修復虛擬模式全選當機問題：by Gemini 3 Flash, 2026/05/09
        '   改用 VirtualListSize 取代 Items.Count，虛擬模式下更正確：by Simon, 2026/05/18
        '
        ' 注意：e.Handled / e.SuppressKeyPress 屬於鍵盤事件責任，
        '       不在此函數內設定，由呼叫端 KeyDown 事件處理函式自行負責。
        ' ---------------------------------------------------------------
        If lv Is Nothing Then Return
        Dim itemCount As Integer = If(lv.VirtualMode, lv.VirtualListSize, lv.Items.Count)
        If itemCount = 0 Then Return

        _dbg("開始", $"全選 {lv.Name} (VirtualMode={lv.VirtualMode}, 共 {itemCount} 項)")
        lv.BeginUpdate()
        Try
            If lv.VirtualMode Then  ' 虛擬模式：直接將所有索引加進 SelectedIndices
                For i As Integer = 0 To lv.VirtualListSize - 1 : lv.SelectedIndices.Add(i) : Next
            Else                    ' 一般模式：遍歷實體項目並設為 Selected
                For Each item As ListViewItem In lv.Items : item.Selected = True : Next
            End If
        Finally
            lv.EndUpdate()
        End Try

        ' 如果有傳入 KeyEventArgs，直接在這裡將按鍵事件標記為已處理
        If e IsNot Nothing Then
            e.Handled = True
            e.SuppressKeyPress = True
        End If

        _dbg("結束", $"共選取 {lv.SelectedIndices.Count} 個項目")
    End Sub
    Private Sub LviCopyToClipboard(lv As ListView, Optional e As KeyEventArgs = Nothing)
        ' ---------------------------------------------------------------
        ' 抽離共用的 ListView 複製到剪貼簿邏輯
        ' 格式: Tab 分隔欄位，換行分隔列，直接貼入 Excel 即可對齊欄位。
        '
        ' [修改歷程]
        ' - 2026/04/27 by Claude Sonnet 4.6: 建立 Ctrl-C 複製選取列初始功能 (Sub為VB保留字，變數改用si)
        ' - 2026/04/27 by Gemini 3.1 Pro: 新增先加入標題列 (Header) 功能
        ' - 2026/05/04 by Gemini 3 Flash: 加入預分配容量 (Capacity) 優化多欄位 List 操作效能
        ' - 2026/05/18 by Gemini 3.0 Flash: 抽離為共用函數，改用動態容量並加入 Try-Catch 剪貼簿保護
        ' - 2026/05/19 by Gemini 3 Flash: 新增 ListView1 特殊判定，選取複製時連同內含 "▸"(PST名稱) 與 "▶"(合計列) 一併匯出
        ' ---------------------------------------------------------------
        If lv Is Nothing OrElse lv.SelectedItems.Count = 0 Then Return

        _dbg("開始", $"複製 {lv.Name} 資料")

        ' 1. 加入標題列 (Header) (by Gemini 3.1 Pro, 2026/04/27)
        Dim sb As New System.Text.StringBuilder()
        Dim headers As New List(Of String)(lv.Columns.Count)    ' 動態預分配容量，取代舊版的固定容量 (改良自 Gemini 3 Flash, 2026/05/04)
        For Each col As ColumnHeader In lv.Columns
            headers.Add(col.Text.Trim())
        Next
        sb.AppendLine(String.Join(vbTab, headers))

        ' 2026/05/19 by Gemini 3 Flash: 根據 ListView 控制項名稱決定要走訪的項目集合
        Dim itemsToCopy As IEnumerable(Of ListViewItem)
        If lv.Name = "ListView1" Then
            ' 【ListView1 特殊處理】因為 PST 名稱列與合計列屬於不開放選取的特殊列，無法存在於 SelectedItems 中。
            ' 故改為走訪「全體項目」，並篩選出「被選取的項目」或是「包含 ▸、▶ 裝飾字元的特殊列」，以維持完整報表結構。
            itemsToCopy = lv.Items.Cast(Of ListViewItem)().Where(Function(item) item.Selected OrElse item.Text.Contains("▸"c) OrElse item.Text.Contains("▶"c))
        Else
            ' 【其他 ListView】維持原邏輯：只複製真正被選取的項目
            itemsToCopy = lv.SelectedItems.Cast(Of ListViewItem)()
        End If

        ' 用來統計最終實質複製了多少列的計數器 (by Gemini 3 Flash, 2026/05/19)
        Dim copiedCount As Integer = 0

        ' 2. 遍歷所有被選取(或符合ListView1條件)的項目，把子欄位以 vbTab (Tab字元) 串接
        For Each item As ListViewItem In itemsToCopy
            Dim cols As New List(Of String)(item.SubItems.Count)
            ' 走訪所有子項目 (by Claude Sonnet 4.6, 2026/04/27: sub 是 VB 保留字，改用 si)
            For Each si As ListViewItem.ListViewSubItem In item.SubItems
                cols.Add(si.Text.Trim(" "c, "-"c, "▸"c, "▶"c))  ' 去除頭尾空白，以及 "-"、"▸" 等視覺裝飾字元
            Next
            sb.AppendLine(String.Join(vbTab, cols))
            copiedCount += 1
        Next

        Try
            Clipboard.SetText(sb.ToString())
            PgrsBar2.Text = $"已複製標題與 {copiedCount:N0} 列到剪貼簿。"   ' 將顯示數量改為實際複製的總列數計數 (by Gemini 3 Flash, 2026/05/19)

            ' 如果有傳入 KeyEventArgs，直接在這裡將按鍵事件標記為已處理
            If e IsNot Nothing Then
                e.Handled = True
                e.SuppressKeyPress = True
            End If

        Catch ex As System.Exception
            _dbg("剪貼簿存取失敗", ex.Message)
            MessageBox.Show("無法存取剪貼簿，可能被其他程式佔用。", "複製失敗", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try

        _dbg("結束")
    End Sub
    Private Function GetCurrentLv() As ListView
        ''' <summary>
        ''' 根據 TabControl1 的選擇索引，判斷並傳回當前畫面上活動中的 ListView, by simon, 2026/05/09
        ''' </summary>
        ' 在需要觸發 AfterSelect 或其他操作時，能夠根據目前選中的 Tab 頁面，準確地獲取對應的 ListView 控制項
        Select Case TabControl1.SelectedIndex
            Case 0 : Return ListView1   ' 2026/04/13 by Simon/Claude: Tab1 改用 ListView1
            Case 1 : Return ListView2
            Case 2 : Return ListView3
            Case 3 : Return Listview4
            Case 4 : Return ListView5
            Case Else : Return Nothing
        End Select
    End Function
    Private Function GetAllLvList(container As Control) As List(Of ListView)
        ''' <summary>
        ''' 遞迴搜尋容器內所有的 ListView
        ''' </summary>
        Dim list As New List(Of ListView)(16) ' 預設容量 16，避免頻繁擴容
        For Each ctrl As Control In container.Controls
            If TypeOf ctrl Is ListView Then list.Add(CType(ctrl, ListView)) ' 如果是 ListView
            If ctrl.HasChildren Then list.AddRange(GetAllLvList(ctrl))   ' 如果有子容器 (如 SplitContainer, TabControl, Panel)，繼續遞迴往下層掃描
        Next
        Return list
    End Function
    Private Function GetHeaderRowBackColor(item As ListViewItem) As Color
        ' 2026/04/14 by Simon/Claude: 根據 Tag=Nothing 列的文字前綴還原正確的 BackColor
        '   ▸ 開頭 = 群組標題行 → SystemColors.GradientInactiveCaption (淡藍)
        '   ▶ 開頭 = 合計列     → Color.FromArgb(220, 235, 252) (稍深藍)
        '   其他    = fallback   → Color.Empty
        If item.Text.StartsWith("▸") Then Return SystemColors.GradientInactiveCaption
        If item.Text.StartsWith("▶") Then Return Color.FromArgb(220, 235, 252)
        Return Color.Empty
    End Function
    Private Function FindLvItemByName(lv As ListView, itemName As String) As ListViewItem
        _dbg("開始", lv.Name)
        For Each item As ListViewItem In lv.Items
            If item.Text.Replace(" - ", "") = itemName.Replace(" - ", "") Then Return item
        Next : Return Nothing

    End Function
#End Region
#Region "  └ 其他輔助函數"
    Private Sub SyncDebugFormMoveOnly()
        ''' <summary>
        ''' 拖曳期間的輕量跟隨 — 只搬位置，不改尺寸 (SWP_NOSIZE)。
        ''' 原 SyncDebugFormResize 每 tick 都用 screenRight-newLeft 重算寬度 = 對 debugForm 做完整 resize
        ''' (WM_SIZE→重佈局→重繪)，同執行緒同步執行卡住 Form1 拖動迴圈，造成抖動/殘影。
        ''' 此處只 blit 搬位置 (近乎零成本)，完整貼齊右緣(含寬高)延到 Form1_ResizeEnd 一次處理。
        ''' 2026/06/19 by Simon/Claude Opus 4.8
        ''' </summary>
        If DebugForm IsNot Nothing AndAlso
            (DebugForm.Visible OrElse CheckDebug.Checked) Then SetWindowPos(DebugForm.Handle, IntPtr.Zero,
                                                                            Me.Left + Me.Width - 12, Me.Top, 0, 0, SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_NOSIZE)
    End Sub
    Private Sub SyncDebugFormResize()
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
            SetWindowPos(DebugForm.Handle, IntPtr.Zero,
                         newLeft, newTop, newWidth, newHeight, SWP_NOZORDER Or SWP_NOACTIVATE)
        End If

    End Sub
    Private Sub ClearMemoryCachesCore()
        ' ---------------------------------------------------------------
        ' ClearMemoryCachesCore — [記憶體層] 統一清理所有 ConcurrentDictionary
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
        _cacheSubTreeList.Clear()
        _cacheIsMailFolder.Clear()
        _cacheFolderIDs.Clear()     ' 2026/04/10 新增 ID 快取清理

        _cacheMailBody.Clear()      ' 2026/6/18 by simon, 之前漏掉了現在補上
        _cacheMailInfo.Clear() ' 2026/6/18 by simon, 之前漏掉了現在補上

    End Sub
#End Region

    Public Class ThemeColors
        ' by Gemini, 2026/04/01: 統一管理專案色彩, 方便日後切換深色/淺色主題
        ''' <summary>主要視窗或Panel背景色 (#F2F2F2)</summary>
        Public Shared ReadOnly Gray95 As Color = Color.FromArgb(242, 242, 242)
        ''' <summary>比DimGray更深的灰色, 主要用在字型 (#404040)</summary>
        Public Shared ReadOnly DarkerDimGray As Color = Color.FromArgb(64, 64, 64)
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
