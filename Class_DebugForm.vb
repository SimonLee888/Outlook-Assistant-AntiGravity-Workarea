Imports System.Reflection
Imports System.Text.RegularExpressions

' ==============================================================
' DebugForm.vb  —  執行期除錯視窗
' ==============================================================
' 功能:
'   即時顯示 Form1._dbg() 呼叫的訊息 (訊息文字、呼叫者、時間戳記、間隔毫秒)
'   雙擊單行 → 複製該行完整文字到剪貼簿，並重算與配對 Begin/End 的時間差
'   Ctrl+C   → 複製所有已選取的行 (Tab 分隔，每行一列)
'   點選 End: 行 → 向前搜尋相符的 Begin: 行並以黃色標示
'
' 設計說明:
'   AddMessage3 由 Form1._dbg() 呼叫，forcedCaller 由 Form1.WhoCallsMe() 填入
'   WhoCallsMe() 為 fallback，正常情況下不會被走到 (Form1 已先解析好呼叫者)
'   ListView 啟用 MultiSelect=True (於 Load 覆寫 Designer 設定) ，支援 Ctrl+C 多選複製
'
' 改動記錄:
'   2026/3/6  - AddMessage2: 合併寫入與更新邏輯，BeginUpdate 批次重繪 (by Claude.ai)
'   2026/3/22 - AddMessage3: 支援 forcedCaller 參數；WhoCallsMe 支援 skipLevels (by Grok.ai)
'   2026/3/23 - 結構整理: 加入 Region、改名 DebugForm_Load、移除無用 Imports
'  (by Claude)  補實作 Ctrl+C 多選複製；移除空白 SelectedIndexChanged；統一 WhoCallsMe 風格
'
' ==============================================================

Public Class DebugForm

#Region "■ 00 Form 雙緩衝"
    '' 2026/04/18 by Claude: 與 Form1 相同的雙緩衝設定
    '' DebugForm 開啟時主要卡頓來源是 Timer_Tick 高頻觸發 BeginUpdate/EndUpdate + EnsureVisible，
    '' 這兩項設定可改善切換焦點與 Resize 時的撕裂感，但無法根治高頻更新本身的開銷。
    'Protected Overrides ReadOnly Property CreateParams As CreateParams
    '    Get
    '        Dim cp As CreateParams = MyBase.CreateParams
    '        cp.ExStyle = cp.ExStyle Or &H2000000    ' WS_EX_COMPOSITED：子控制項合成層雙緩衝
    '        Return cp
    '    End Get
    'End Property
    Protected Overrides Sub OnLoad(e As EventArgs)
        Me.DoubleBuffered = True    ' Form 自身 WM_PAINT 雙緩衝，與 WS_EX_COMPOSITED 互補無衝突
        MyBase.OnLoad(e)
    End Sub
#End Region

#Region "■ 01 Win32 API & 常數"
    <Runtime.InteropServices.DllImport("user32.dll")>
    Private Shared Function SendMessage(hWnd As IntPtr, msg As Integer, wParam As IntPtr, lParam As IntPtr) As IntPtr
    End Function

    ' 2026/06/19 by Simon/Claude Opus 4.8: 改 OS 層 class background brush，消除撐高瞬間新區域的黑塊 (受控層 BackColor 太晚、壓不到這一幀)
    <Runtime.InteropServices.DllImport("user32.dll", EntryPoint:="SetClassLongPtrW")>
    Private Shared Function SetClassLongPtr(hWnd As IntPtr, nIndex As Integer, dwNewLong As IntPtr) As IntPtr
    End Function
    <Runtime.InteropServices.DllImport("gdi32.dll")>
    Private Shared Function CreateSolidBrush(crColor As Integer) As IntPtr
    End Function

    Private Const WM_SETFONT As Integer = &H30
    Private Const WM_SETREDRAW As Integer = &HB         ' 2026/3/26 by Gemini
    Private Const WM_SIZE As Integer = &H5              ' by Claude Opus 4.6, 2026/04/11: 攔截視窗尺寸變更
    Private Const SIZE_MAXIMIZED As Integer = 2
    Private Const SIZE_RESTORED As Integer = 0
    Private Const LVM_FIRST As Integer = &H1000
    Private Const LVM_GETHEADER As Integer = LVM_FIRST + 31
    Private Const LVM_SETEXTENDEDLISTVIEWSTYLE As Integer = LVM_FIRST + 54
    Private Const LVM_GETEXTENDEDLISTVIEWSTYLE As Integer = LVM_FIRST + 55
    Private Const LVM_SETTOOLTIPS As Integer = LVM_FIRST + 74   ' by Gemini 3 Flash, 2026/04/13: 用於切斷 ToolTip 控制項關聯
    Private Const LVS_EX_LABELTIP As Integer = &H4000           ' by Claude Opus 4.6, 2026/04/11: 移除此樣式以修復 OwnerDraw 下的文字重疊殘影
    Private Const LVS_EX_DOUBLEBUFFER As Integer = &H10000      ' by Claude, 2026/04/12: Native ListView 真正的雙緩衝 flag，解決 ScrollBar 消失時 OwnerDraw dirty region 只有右側條帶導致項目消失的 Bug
    Private Const GCLP_HBRBACKGROUND As Integer = -10
#End Region

#Region "■ 02 成員變數"
    Private WithEvents QueueTimer As New Timer() With {.Interval = 16}      ' 啟動時先預設每 16ms 清空一次message queue
    Private _msgQueue As New Concurrent.ConcurrentQueue(Of PendingDebugMsg) ' 2026/07/11 by Simon/Sonnet 5: 改存輕量 DTO，ListViewItem 建構延後到 Timer_Tick
    Private _lastRecalcWidth As Integer = 0
    Private _previousTimestampTicks As Long             ' 2026/07/11 by Simon/Sonnet 5: 原本是 Date 型別，多執行緒同時呼叫 AddMessage3 時「讀取-計算-寫回」不具原子性，交錯執行會讓 Step 間隔算錯或遺失更新。改存 Ticks 用 Interlocked.Exchange 原子交換。
    Public Shared ActiveInstance As DebugForm = Nothing ' by Gemini 3.5 Flash, 2026/06/19: 儲存作用中的 DebugForm 實例以供背景執行緒存取，解決 VB 預設實例在非 UI 執行緒的 Thread-Local 陷阱。

    Private _searchPattern As String = ""
    Private _searchRegex As Regex = Nothing             ' 2026/06/15 by Simon/Claude: 由 _searchPattern 預編譯, DrawSubItem 直接套用, 省去每格字串多載重複解析
    Private _fillBrush As New SolidBrush(Color.White)   ' 2026/06/15 by Simon/Claude: 背景填色重用同一支 brush (改 .Color 即可), 取代每格 New SolidBrush/Dispose
    Private _classBgBrush As IntPtr = IntPtr.Zero       ' 2026/06/19: OS 層白底 brush，存欄位避免被回收

    Private _lastHighlightedPair As ListViewItem        ' by Gemini, 2026/03/29: O(1) 顏色還原，取代 For Each 全域清除
    Private _suppressPairing As Boolean = False         ' 2026/07/11 by Simon/Sonnet 5: Timer_Tick 自動捲動選取時設 True, 讓 ItemSelectionChanged 略過配對高亮掃描 (原本每 100ms 觸發一次 O(N) FindSimilarPair)
    Private _historyDebug As New List(Of String)(256)   ' by AntiGravity, 2026/04/07: 搜尋歷史紀錄
    Private _historyIndex As Integer = 0                ' by AntiGravity, 2026/04/07: 目前歷史紀錄索引 (與 Count 相同時代表原始輸入區)
    Private _tempInput As String = ""                   ' by AntiGravity, 2026/04/07: 暫存回溯前的原始輸入內容
    Private _cachedKeywordsLower As New List(Of String) ' 2026/07/11 by Simon/Sonnet 5: RefreshSearchCache 預先算好的小寫關鍵字，AddMessage3 只讀欄位, 不碰 txtDebug/checkAndOr (可能由背景執行緒呼叫)
    Private _cachedAndMode As Boolean = False           ' 2026/07/11: 同上，快取 AND/OR 模式
    Private Class DebugItemTag              ' 2026/3/28 by Gemini: 定義快取結構，加速 OwnerDraw 繪製
        Public textFullRow As String        ' 預先合併好的整行小寫文字 (用於搜尋)
        Public isHit As Boolean             ' 是否命中目前搜尋關鍵字
        Public timeStamp As Date            ' 2026/3/28 by Gemini: 原始時間戳記 (供雙擊重算時間差，免去 TryParse 反解)
        Public coreKey As String            ' 2026/07/11 by Simon/Sonnet 5: FindSimilarPair 用的比對核心鍵 (RemoveBeginEnd 結果)，建立時算一次
        Public isBeginRow As Boolean        ' 2026/07/11: 是否為「開始」行，建立時算一次
        Public isEndRow As Boolean          ' 2026/07/11: 是否為「結束」行，建立時算一次
    End Class
    Private Class PendingDebugMsg
        ' 2026/07/11 by Simon/Sonnet 5: AddMessage3 端的輕量 DTO，只夾帶字串與時間戳。
        ' ListViewItem/SubItems/DebugItemTag 的建構全部延後到 Timer_Tick 批次處理時才做，讓呼叫端(可能是密集迴圈或背景執行緒)不必再付出 WinForms 物件配置的成本。
        Public msgContent As String
        Public timeNow As Date
        Public timeSpan As TimeSpan
        Public lineNo As Integer
    End Class

    ' 2026/07/11 by Simon/Sonnet 5: 兩個狀態機名稱 Regex 改成 Shared ReadOnly + Compiled，只編譯一次終身重用，
    ' 取代原本 GetCallerName 每次呼叫、每個 Async 呼叫者都用字串多載重新解析 pattern
    Private Shared ReadOnly _vbStateMachineRegex As New Regex("^VB\$StateMachine_\d+_(.*)$", RegexOptions.Compiled)
    Private Shared ReadOnly _csStateMachineRegex As New Regex("^<(.*)>d__.*$", RegexOptions.Compiled)
    ' 2026/07/11 by Simon/Sonnet 5: 拆詞 pattern 改成 Shared ReadOnly + Compiled，只編譯一次終身重用，
    ' 取代原本每次搜尋框變動 (每敲一鍵) 都用字串多載重新解析 pattern
    Private Shared ReadOnly _keywordSplitRegex As New Regex("(?:""(?<q>[^""]*)""|(?<w>\S+))", RegexOptions.Compiled)
#End Region

#Region "■ 03 表單生命週期"
    Protected Overrides Sub WndProc(ByRef m As Message)
        ' 💡 2026/04/11 by Claude Opus 4.6: 攔截 WM_SIZE 修復最大化/還原時 ListView 項目消失
        ' 原因：雙擊標題列觸發的 WindowState 切換是瞬間完成的（不像拖動是漸進式），
        ' OwnerDraw ListView 在項目不足以填滿視窗高度時，內部繪製管線會誤判不需要重繪。
        ' 解法：在最大化/還原完成後，強制觸發一次 Invalidate() 讓 ListView 重新繪製所有可視項目。
        MyBase.WndProc(m)
        If m.Msg = WM_SIZE Then
            Dim sizeType As Integer = m.WParam.ToInt32()
            If sizeType = SIZE_MAXIMIZED OrElse sizeType = SIZE_RESTORED Then lvwDebug.Invalidate()
            ' 2026/06/19 by Simon/Claude: 加 lvwDebug.Update() 強制重繪或改 Me.Refresh() 同步重繪都無效，
            ' 無法消除一次性撐高時底部新區域的黑塊空窗
        End If
    End Sub
    Private Sub ApplyListViewFixes()
        ''' <summary>
        ''' 2026/04/13 by Gemini 3 Flash: 集中處理 Win32 樣式修復 (LabelTip 移除、隔離 ToolTip、原生雙緩衝)
        ''' </summary>
        Try
            ' 1. 移除 LVS_EX_LABELTIP (防止 OwnerDraw 時出現鬼影文字標籤)
            '       💡 2026/04/11 by Claude Opus 4.6 移除 LVS_EX_LABELTIP 擴充樣式
            '       在 OwnerDraw 模式下， ListView 內建的「文字超寬時浮出全文標籤」會跟自訂繪製衝突，產生滑鼠移過時的文字重疊殘影。移除此樣式即可根治。
            '       wParam = mask(指定要修改哪些位元), lParam = 0 (關閉這些位元)
            ' by Gemini, 2026/04/13: 全部整合到 ApplyListViewFixes，確保 Handle 重建後依然生效
            SendMessage(lvwDebug.Handle, LVM_SETEXTENDEDLISTVIEWSTYLE, New IntPtr(LVS_EX_LABELTIP), IntPtr.Zero)

            ' 2. 徹底隔離 ToolTip 控制項 (切斷與系統調度器的關聯)
            SendMessage(lvwDebug.Handle, LVM_SETTOOLTIPS, IntPtr.Zero, IntPtr.Zero)

            ' 3. 啟用 LVS_EX_DOUBLEBUFFER (強化滾動與 Resize 時的渲染穩定性)
            '       💡 2026/04/12 by Claude 啟用 LVS_EX_DOUBLEBUFFER
            '       這是 Native Win32 ListView 的真正雙緩衝， 與.NET DoubleBuffered 屬性完全不同
            '       啟用後 ListView 每次 WM_PAINT 都對整個 client area 做 offscreen buffer blit，不再做 partial dirty-region paint， 從根本解決 ScrollBar 消失時 DrawSubItem 不被呼叫的問題
            ' by Gemini, 2026/04/13: 全部整合到 ApplyListViewFixes，確保 Handle 重建後依然生效
            SendMessage(lvwDebug.Handle, LVM_SETEXTENDEDLISTVIEWSTYLE, New IntPtr(LVS_EX_DOUBLEBUFFER), New IntPtr(LVS_EX_DOUBLEBUFFER))

        Catch ex As Exception
            ' 僅為 UI 修正，不應中斷程式
        End Try

    End Sub
    Private Sub OnLvwHandleCreated(sender As Object, e As EventArgs)
        ''' <summary>
        ''' 2026/04/13 by Gemini 3 Flash: 當 ListView Handle 建立時，自動重新套用樣式修復與 ToolTip 隔離
        ''' </summary>
        ApplyListViewFixes()
    End Sub
    Private Sub DebugForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ActiveInstance = Me        ' by Gemini 3.5 Flash, 2026/06/19: 在 Load 時設定 ActiveInstance 為目前實例

        '' 2026/04/01 by Gemini: 恢復 ListView 內建雙緩衝設置
        ''   此設定可避免 AddMessage3 (Timer 批次新增) 時產生的背景擦除閃爍。
        '' 2026/6/19 關閉此設定，因為已經在 ApplyListViewFixes() 中啟用 LVS_EX_DOUBLEBUFFER, 二者重疊是多餘的。
        'Dim pi = lvwDebug.GetType().GetProperty("DoubleBuffered", BindingFlags.Instance Or BindingFlags.NonPublic)
        'If pi IsNot Nothing Then pi.SetValue(lvwDebug, True, Nothing)

        _previousTimestampTicks = Now.Ticks : QueueTimer.Start()   ' 啟動時先預設每 16ms 清空一次message queue
        AddHandler lvwDebug.HandleCreated, AddressOf OnLvwHandleCreated ' 💡 2026/04/13 by Gemini 3 Flash: 註冊 HandleCreated，確保 ListView 重建時修復依然生效

        ' 2026/06/19 by Simon/Claude: 把 debugForm 的 class 背景 brush 換成白色，讓 OS 在 SetWindowPos 撐高瞬間用白色填新區域，取代預設 NULL→黑。
        _classBgBrush = CreateSolidBrush(&HFFFFFF)
        SetClassLongPtr(Me.Handle, GCLP_HBRBACKGROUND, _classBgBrush)

    End Sub
    Private Sub DebugForm_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        ' ==============================================================
        ' by Gemini, 2026/04/01: 將重型 UI 佈局校算移到 Shown 事件
        ' 目的: 讓 Form1 觸發開啟除錯視窗後能立即返回，不等待 UI 佈局渲染，優化啟動延遲感
        ' ==============================================================

        ' 1. 將搜尋列移至頂部 (比照 Tab5)，並設定固定高度
        ' 2026/3/27 by Gemini: ── 佈局一致化優化 (穩定 Dock 佈局) ──
        pnlSearch.Dock = DockStyle.Top : pnlSearch.Height = 45
        pnlSearch.BackColor = Form1.ThemeColors.Gray95

        ' 2. 設定搜尋文字框與邏輯切換開關的對齊與錨定
        ' 2026/3/28 by Gemini: 重寫佈局 — 先固定右側 CheckBox，再讓 TextBox 填滿剩餘空間
        Dim targetTop As Integer = (pnlSearch.ClientSize.Height - txtDebug.Height) \ 2

        ' 3. 設定右側: checkAndOr (AND/OR 切換) ──
        checkAndOr.AutoSize = True
        checkAndOr.FlatStyle = FlatStyle.System
        checkAndOr.TabStop = False
        checkAndOr.Location = New Point(pnlSearch.ClientSize.Width - checkAndOr.Width - 12,
        targetTop + (txtDebug.Height - checkAndOr.Height) \ 2 + 8.5) ' 垂直中心對齊 TextBox
        checkAndOr.Anchor = AnchorStyles.Top Or AnchorStyles.Right   ' 隨表單放大貼緊右上

        ' 4. 設定左側: txtDebug (搜尋輸入框) ──
        txtDebug.ImeMode = ImeMode.Alpha            ' by AntiGravity, 2026/04/07: 強制預設英文/半形英數，解決輸入法自動切換中文問題
        txtDebug.Location = New Point(8, targetTop) ' 距左側 8px
        txtDebug.Width = checkAndOr.Left - 8 - 12   ' 右邊預留 12px 間距到 CheckBox
        txtDebug.Anchor = AnchorStyles.Top Or
        AnchorStyles.Left Or AnchorStyles.Right     ' 隨表單放大自動展寬

        ' 5. 設定列表填滿剩餘空間
        lvwDebug.Anchor = AnchorStyles.None ' 2026/3/27 by Gemini: 清除 Anchor 避免與 Dock 衝突
        lvwDebug.Dock = DockStyle.Fill
        lvwDebug.OwnerDraw = True           ' 準備允許自訂重繪搜尋字串高亮
        lvwDebug.MultiSelect = True         ' 2026-03-23: 啟用多選才能讓 Ctrl+C KeyDown 複製多行, ✅ 覆寫 Designer 的 MultiSelect=False，支援 Ctrl+C 多選複製

        ' 6. 設定正確的Z-Order填充順序
        ' 2026/3/27 by Gemini: 正確的 Dock Z-Order 邏輯
        lvwDebug.BringToFront()     ' WinForms 中，Dock=Top 的面板必須在 Controls 最尾端 (SendToBack) 才會最先取得空間。
        pnlSearch.SendToBack()      ' Dock=Fill 的控制項必須在 Controls 最前端 (BringToFront) 才會填滿剩餘空間。
        ' simon: 使用視覺設計表單物件30年, 我到今天才知道它們的 Z-Order 前後順序會影響 Dock 的填充邏輯?!

        ' 7. 初始化欄位
        With lvwDebug.Columns
            .Clear()
            .Add("Debug Message", 400, HorizontalAlignment.Left)    ' 寬度會在 Load 時被 RecalcColumnWidths 調整，這裡先給個預設值
            .Add("Timestamp", 115, HorizontalAlignment.Center)
            .Add("Step (ms)", 85, HorizontalAlignment.Right)        ' by Gemini 1.5 Pro, 2026/04/11: 原 Time Span，顯示物理步進間隔
            .Add("Elapsed (ms)", 85, HorizontalAlignment.Right)     ' by Gemini 1.5 Pro, 2026/04/11: 新增，顯示函數從開始到結束的總耗時
            '.Insert(0, New ColumnHeader() With {.Text = "Debug Message", .Width = -2, .TextAlign = HorizontalAlignment.Left})   ' 2026/3/28 by Gemini: Width=-2 讓第一欄自動填滿剩餘空間，避免寫死寬度在 Load 時擠掉右側欄位
        End With
        RecalcColumnWidths(Nothing, Nothing)    ' 2026/3/30 by Gemini: 在 Load 時手動觸發強制調整一次，確保初始顯示正確 (特別是第一欄填滿剩餘空間)

        AddHandler lvwDebug.ItemSelectionChanged, AddressOf lvwDebug_ItemSelectionChanged
        AddHandler lvwDebug.ClientSizeChanged, AddressOf RecalcColumnWidths
        AddHandler txtDebug.KeyDown, AddressOf txtDebug_KeyDown ' by AntiGravity, 2026/04/07: 支持搜尋歷史回溯

        ' 2026/3/28 by Gemini: 監聽 lvwDebug 本身的 ClientSizeChanged 事件，
        ' 無論何時 ListView 可用空間改變 (Dock 佈局結算、表單 Resize、SyncDebugFormResize)，都自動重算欄寬, 不再需要猜延遲值或一次性 Timer
        ' by Gemini, 2026/03/29: 右鍵管理選單 (只建立一次，不重複 AddHandler)
        Dim ctx As New ContextMenuStrip()
        ctx.Items.Add("計算選取耗時", Nothing, AddressOf CalculateSelectedTimeSpan)
        ctx.Items.Add("刪除選取項目", Nothing, AddressOf DeleteSelectedItems)
        ctx.Items.Add("清除所有項目", Nothing, Sub(s, ev) lvwDebug.Items.Clear())
        lvwDebug.ContextMenuStrip = ctx

        ' 💡 2026/04/11 by Gemini: 原生 Header 粗體化優化 (不破壞 MouseOver 顏色變化)
        ' 直接透過 Win32 套用字型，比 OwnerDraw 更穩定且保留原生互動。
        Try
            Dim hHeader As IntPtr = SendMessage(lvwDebug.Handle, LVM_GETHEADER, IntPtr.Zero, IntPtr.Zero)
            If hHeader <> IntPtr.Zero Then
                Static boldFont As Font = New Font(lvwDebug.Font, FontStyle.Bold) ' 使用 Static 確保 Font 物件生命週期隨 Form 存在
                SendMessage(hHeader, WM_SETFONT, boldFont.ToHfont(), New IntPtr(1))
            End If
        Catch ex As Exception
            ' 僅為 UI 裝飾，失敗則跳過，不影響核心邏輯
        End Try

        ' 💡 2026/04/13 by Gemini 3 Flash: 執行 ListView 樣式修復 (原放在 Shown 的邏輯現已整合進 ApplyListViewFixes)
        ApplyListViewFixes()

        QueueTimer.Interval = 100   ' 2026/6/19 by simon: 啟動完成後把更新間隔減慢為每 100ms 清空一次message queue

    End Sub
    Private Sub DebugForm_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed

        Form1.CheckDebug.Checked = False
        ActiveInstance = Nothing        ' by Gemini 3.5 Flash, 2026/06/19: 表單關閉後清除 ActiveInstance

    End Sub
    Private Sub RecalcColumnWidths(sender As Object, e As EventArgs)

        ' 2026/04/01 by Gemini: 修正 ListView 項目在卷軸消失時跟著消失的致命 Bug
        ' 1. 加入門檻判定 (Threshold): 寬度變動極小時不觸發重設，避免拖動尺寸時的頻繁重發 (Throttle)
        If lvwDebug.Columns.Count < 3 Then Return
        If Math.Abs(lvwDebug.ClientSize.Width - _lastRecalcWidth) < 2 Then Return
        _lastRecalcWidth = lvwDebug.ClientSize.Width

        ' 讓第一欄 (Debug Message) 填滿剩餘空間，其餘欄位維持既有寬度
        Dim reservedWidth As Integer = 0
        For i As Integer = 1 To lvwDebug.Columns.Count - 1
            reservedWidth += lvwDebug.Columns(i).Width
        Next

        ' 💡 絕不可以在 ClientSizeChanged 期間呼叫 BeginUpdate/EndUpdate！
        ' 否則在「卷軸隱藏/顯示」的重算週期中會癱瘓底層訊息傳遞，導致所有項目全部消失。
        Dim newWidth As Integer = lvwDebug.ClientSize.Width - reservedWidth - 4

        '' by Claude Opus 4.6, 2026/04/11: 修復卷軸消失時所有 ListView 項目消失的致命 Bug
        '' 💡為什麼 Invalidate() 無效: 在 ClientSizeChanged resize 訊息鏈進行中，Windows 會抑制 WM_PAINT 派發。Invalidate() 只是排入 WM_PAINT 到佇列，等 resize 結束時 item bounds 快取早已壞掉。
        '' 💡為什麼 BeginInvoke + Refresh() 有效: BeginInvoke 將 delegate 排入訊息泵，保證在 resize 訊息鏈**完全結束後**才執行。Refresh() = Invalidate() + Update()，Update() 同步處理 WM_PAINT，不會被延遲。
        'If newWidth > 100 AndAlso lvwDebug.Columns(0).Width <> newWidth Then
        '    lvwDebug.Columns(0).Width = newWidth
        '    BeginInvoke(Sub() If lvwDebug IsNot Nothing AndAlso Not lvwDebug.IsDisposed AndAlso lvwDebug.Items.Count > 0 Then lvwDebug.Refresh())
        'End If

        '' 2026/04/12 by Claude: 修復 ScrollBar 消失時 ListView 項目消失的 Bug
        '' 根本原因：Columns(0).Width 賦值會觸發 ListView 內部同步 repaint，此時 DoubleBuffer backbuffer 被清空，但 GDI 系統 clip 只有右側17px條帶，導致 TextRenderer.DrawText (GDI) 被 clip 住畫不出文字。
        '' 解法：用 WM_SETREDRAW 壓住 column 改變觸發的內部 paint，改成寬度設定完後呼叫一次完整 Refresh()，此時 dirty region 是全區域。
        'If newWidth > 100 AndAlso lvwDebug.Columns(0).Width <> newWidth Then
        '    SendMessage(lvwDebug.Handle, WM_SETREDRAW, New IntPtr(0), IntPtr.Zero)
        '    lvwDebug.Columns(0).Width = newWidth
        '    SendMessage(lvwDebug.Handle, WM_SETREDRAW, New IntPtr(1), IntPtr.Zero)
        '    lvwDebug.Refresh()  ' Invalidate() + Update()，同步執行，此時 clip = 完整區域
        '    ' 移除原本的 BeginInvoke(Refresh()) — 由上面的同步 Refresh() 取代
        'End If

        ' 2026/04/12 by Claude: ScrollBar 消失後 EnsureVisible 殘留的 scroll offset 未歸零
        '   導致 item(0).Bounds.Y 為負數，所有項目偏移至底部，Refresh 在錯誤座標執行也無效
        '   檢查 item(0).Bounds.Y：不等於 0 代表 scroll offset 殘留，強制設 TopItem 歸零

        ' ── 失敗修補史精簡 (2026/06/19 by Simon/Claude Opus 4.8 整理；原始多段嘗試碼已濃縮) ──
        ' 「卷軸消失瞬間 listviewitem 全數消失」為長期未解 bug，已嘗試 10+ 次。
        ' 2026/06/15 探針定論：vscroll=True 時 Refresh()/Invalidate() 觸發 200~440 次 DrawSubItem；
        '   vscroll 一變 False，Invalidate / Refresh / RedrawItems+Update 全部 = 0 次 DrawSubItem (空白緩衝)，且卡死狀態延續到卷軸回來之後，直到 item 層級事件(hover) 或足夠的重新佈局才解除。
        ' 已否證/作廢的修法 (全屬「視窗層級」重繪，無卷軸狀態下皆 0 draw)：
        '   ① BeginInvoke+Refresh()  ② WM_SETREDRAW 包夾+Refresh()  ③ RedrawItems(LVM_REDRAWITEMS) 2026/06/15 實測仍 0 draw、item 仍消失
        '   ④ 檢查 Items(0).Bounds.Y<>0 後重設 TopItem — 探針證實 bounds 無殘留，此修法在修不存在的病因，且會在 resize 時把使用者捲動位置硬拉回頂端 → 2026/06/19 移除。
        ' 真因方向 (待驗證)：疑似多層雙緩衝 (WS_EX_COMPOSITED + native LVS_EX_DOUBLEBUFFER + managed DoubleBuffered) 在 client 寬度 ±17px(卷軸增減) 時後備緩衝失同步。
        ' 2026/06/19 by Simon/Claude Opus 4.8: 實測否證、真因確認 -->（多層雙緩衝經四格矩陣否證 + 真因是轉換週期設欄寬）

        ' 2026/06/19 by Simon/Claude Opus 4.8: 兩個月老 bug 根因確認 ——
        ' 在「垂直卷軸顯隱轉換」的同一個 ClientSizeChanged 訊息週期內【同步】設定 Columns(0).Width，會把 native ListView 推進壞掉的繪製狀態 (DrawSubItem 歸零 → 所有 item 消失)。
        ' 旁證：純改寬度拖曳不觸發卷軸顯隱，從來正常；只有改高度跨越「塞得下/塞不下」門檻、卷軸±17px 那刻會壞。
        ' 修法：把欄寬賦值「延後」到訊息鏈結束之後 (BeginInvoke) 才設，此時卷軸已穩定，不再與轉換同週期。

        If newWidth > 100 AndAlso lvwDebug.Columns(0).Width <> newWidth Then
            BeginInvoke(Sub()
                            If lvwDebug Is Nothing OrElse lvwDebug.IsDisposed Then Return
                            lvwDebug.Columns(0).Width = newWidth
                            If lvwDebug.Items.Count > 0 Then
                                lvwDebug.RedrawItems(0, lvwDebug.Items.Count - 1, False)
                                lvwDebug.Update()
                            End If
                        End Sub)
        End If

    End Sub
#End Region

#Region "■ 04 訊息寫入"
    Public Sub AddMessage3(Optional strA As String = "", Optional strB As String = "", Optional forcedCaller As String = "")
        ' ============================================================
        ' 功能:   在主程式只簡單傳入字串, debugForm就自己做好一切訊息準備
        ' 傳入值:
        '          strA: 某動作開始或結束, 或是循環中的更新狀態
        '          strB: 可傳入函數中的碼表計時, 或是函數內某物件的計數
        '   forceCaller: 若傳入的函數名數無法顯示, 可在第3個字串強制指定
        ' ============================================================
        ' ⚠️ 【安全性關鍵警語 - 跨執行緒風險】 by Gemini 3.5 Flash, 2026/06/19
        ' 這份安全性 100% 建立在「AddMessage3 是 enqueue-only」（僅將訊息包成 DTO 寫入 Queue，不直接觸發 UI 更新、不建構任何 WinForms 物件）。
        ' 哪天若有人在此方法內部（或 enqueue 之後的呼叫鏈中）直接更動 ListView、Label 等 UI 控制項，
        ' 背景路徑 (如 Task.Run) 就會立刻發生跨執行緒存取崩潰 (Cross-thread violation)。修補此方法時請務必守住這條紅線！
        ' ============================================================

        ' 2026/3/22 by Grok.ai:
        ' forcedCaller: Form1.WhoCallsMe() 預先解析好的呼叫者字串 (避免 stack trace 在 DebugForm 裡走不回去)
        Dim callingMethod As String = If(forcedCaller <> "", forcedCaller, WhoCallsMe(1))

        ' 計算時間差
        ' 2026/07/11 by Simon/Sonnet 5: Interlocked.Exchange 原子地「取回舊值+寫入新值」，
        ' 修正多執行緒交錯呼叫時的資料競爭 (讀-改-寫非原子，會算錯 Step 或遺失更新)
        Dim timeNow As Date = Now
        Dim prevTicks As Long = System.Threading.Interlocked.Exchange(_previousTimestampTicks, timeNow.Ticks)
        Dim timeSpan As New TimeSpan(timeNow.Ticks - prevTicks)

        Static lineCount As Integer
        Dim newLine As Integer = System.Threading.Interlocked.Increment(lineCount)

        ' 2026/03/31 by Gemini: 優化顯示格式，若第二個參數為空則不顯示括號
        Dim msgContent As String = $"{newLine.ToString("00")} {strA} {callingMethod}"
        If Not String.IsNullOrEmpty(strB) Then msgContent &= $" ({strB})"

        ' 2026/07/11 by Simon/Sonnet 5: AddMessage3 現在只組字串、塞進輕量 DTO 就 Enqueue。
        ' ListViewItem/SubItems/DebugItemTag(含 coreKey/textFullRow 等) 的建構全部移到 Timer_Tick 批次處理時才做 (見 BuildListViewItem)，
        ' 呼叫端 (常是密集迴圈或背景執行緒) 不再需要付出 WinForms 物件配置與字串快取準備的成本。
        _msgQueue.Enqueue(New PendingDebugMsg With {.msgContent = msgContent, .timeNow = timeNow, .timeSpan = timeSpan, .lineNo = newLine})

    End Sub
    Private Function WhoCallsMe(Optional skipLevels As Integer = 1) As String
        ' ================================================================
        ' WhoCallsMe: 若Form1 傳入時未指定呼叫者, 自行從 stack trace 解析呼叫者
        ' skipLevels: 要跳過幾層 wrapper (預設 1 = 跳過 AddMessage3 本身)
        ' 2026-03-23: 統一為 Form1.WhoCallsMe 的邏輯風格
        ' ================================================================
        Dim st As New StackTrace(skipLevels + 1, False)
        For i As Integer = 0 To st.FrameCount - 1
            Dim m = st.GetFrame(i)?.GetMethod()
            If m IsNot Nothing AndAlso
                m.DeclaringType IsNot Nothing AndAlso
                m.DeclaringType IsNot GetType(DebugForm) AndAlso
                Not m.Name.Contains("Dbg") AndAlso
                m.Name <> "MoveNext" Then : Return $"{m.DeclaringType.Name}.{m.Name}"
            End If
        Next
        Return "Unknown Method"

    End Function
    Public Shared Function GetCallerName(Optional skipLevels As Integer = 2) As String
        ''' <summary>
        ''' 2026/03/31 by Gemini: 集中化追蹤呼叫者名稱，支援 Async 非同步方法解析與編譯器生成的狀態機器名稱還原
        ''' </summary>
        ''' <param name="skipLevels">跳過的堆疊層次 (自 GetCallerName 的呼叫層算起)</param>
        ' 2026/03/31 by Gemini: 依照先前計畫重構，

        ' 使用 st.FrameCount 遍歷以確保在複雜非同步環境下仍能抓到正確層級
        Dim st As New StackTrace(skipLevels, False)
        For i As Integer = 0 To st.FrameCount - 1
            Dim frame As StackFrame = st.GetFrame(i)
            Dim m As MethodBase = frame.GetMethod()
            If m Is Nothing OrElse m.DeclaringType Is Nothing Then Continue For

            ' 排除 DebugForm 內部成員與 _dbg 相關噪音
            Dim typeName As String = m.DeclaringType.Name
            If m.DeclaringType Is GetType(DebugForm) OrElse typeName.Contains("DebugForm") Then Continue For
            If m.Name.Contains("Dbg") OrElse m.Name.Contains("WhoCallsMe") Then Continue For

            ' 💡 處理 Async 非同步狀態機器 (MoveNext)
            If m.Name = "MoveNext" AndAlso m.DeclaringType.GetInterface("IAsyncStateMachine") IsNot Nothing Then
                ' 1. VB.NET 格式: VB$StateMachine_123_MethodName
                Dim matchVB = _vbStateMachineRegex.Match(typeName)
                If matchVB.Success Then
                    Dim originalName As String = matchVB.Groups(1).Value
                    Dim parentType = m.DeclaringType.DeclaringType ' 狀態機器類別的上一層通常就是原始類別
                    Return If(parentType IsNot Nothing, $"{parentType.Name}.{originalName} [Async]", $"{originalName} [Async]")
                End If

                ' 2. C# 格式: <MethodName>d__XX (兼顧未來可能混合 C# 專案的情況)
                Dim matchCS = _csStateMachineRegex.Match(typeName)
                If matchCS.Success Then
                    Dim originalName As String = matchCS.Groups(1).Value
                    Dim parentType = m.DeclaringType.DeclaringType
                    Return If(parentType IsNot Nothing, $"{parentType.Name}.{originalName} [Async]", $"{originalName} [Async]")
                End If
            End If
            ' 一般同步方法
            Return $"{m.DeclaringType.Name}.{m.Name}"
        Next
        Return "Unknown Caller"

    End Function
    Private Sub Timer_Tick(sender As Object, e As EventArgs) Handles QueueTimer.Tick
        ' 2026-03-25 by Gemini: 改用 ConcurrentQueue 與 Timer 定期批次新增，大幅提升迴圈寫入效能
        If _msgQueue.IsEmpty Then Return

        ' 2026/07/11 by Simon/Sonnet 5: 從輕量 DTO 建構 ListViewItem 移到這裡批次處理 (見 BuildListViewItem)，
        ' 取代原本在 AddMessage3 呼叫端就建構整組 ListViewItem/SubItems/Tag 的做法
        Dim itemsToAdd As New List(Of ListViewItem)(256)
        Dim pending As PendingDebugMsg = Nothing
        While _msgQueue.TryDequeue(pending) : itemsToAdd.Add(BuildListViewItem(pending)) : End While

        If itemsToAdd.Count > 0 Then
            ' 2026/03/31 by Gemini: 自動為「結束」行預算總耗時
            ' 2026/04/11 by Gemini: 改填入 SubItems(3) (Elapsed 欄位)，並支援跨集合搜尋 (itemsToAdd)
            ' 2206/06/13 by Claude: 效能優化 - 現在要新增的項目如果包含「結束」，才進行配對搜尋，避免每次 Timer 都無差別地 O(N²) 搜尋整個 itemsToAdd 集合
            ' 2026/07/11 by Simon/Sonnet 5: 改讀 tag.isEndRow 快取，取代 lvi.Text.Contains("結束") 的重複字串掃描
            If itemsToAdd.Any(Function(lvi) DirectCast(lvi.Tag, DebugItemTag).isEndRow) Then
                For Each lvi In itemsToAdd
                    Dim tagCurrent = DirectCast(lvi.Tag, DebugItemTag)
                    If tagCurrent.isEndRow Then
                        Dim pair As ListViewItem = FindSimilarPair(lvi, itemsToAdd)
                        If pair IsNot Nothing Then
                            Dim tagPair = DirectCast(pair.Tag, DebugItemTag)
                            Dim totalMs As Double = Math.Abs((tagCurrent.timeStamp - tagPair.timeStamp).TotalMilliseconds)
                            lvi.SubItems(3).Text = totalMs.ToString("#,##0.00")
                            ' by Gemini, 2026/04/11: 填入數值後同步更新搜尋快取
                            ' 2026/07/11 by Simon/Sonnet 5: textFullRow 是延遲建構的，只有先前已經建過(代表搜尋曾經作用中)才需要在這裡同步更新；
                            ' 若從未建過就不必在此補建，維持延遲建構的原則 (RefreshSearchCache 需要時自然會補)
                            If tagCurrent.textFullRow IsNot Nothing Then tagCurrent.textFullRow = BuildFullRowText(lvi)
                        End If
                    End If
                Next
            End If

            With lvwDebug
                .BeginUpdate()
                .Items.AddRange(itemsToAdd.ToArray())
                .EndUpdate()

                ' 💡 2026/04/01 by Gemini:
                ' EnsureVisible 必須在 EndUpdate 之後呼叫，避免在暫停繪製期間滾動引發的瞬間畫面撕裂與閃爍
                If .Items.Count > 0 Then
                    Dim lastItem = .Items(.Items.Count - 1)
                    lastItem.EnsureVisible()

                    ' 2026/04/09 by Gemini: 修正游標前進但舊選取殘留的問題：
                    '   由於已開啟 MultiSelect=True，直接設 Selected=True 會變成加選，因此需先手動清除前次的選取，再設定最後一項，並賦予 Focused 確保游標真正前進
                    ' 2026/07/11 by Simon/Sonnet 5: 用 _suppressPairing 包住這段自動捲動選取，避免每次 Timer_Tick 都觸發 ItemSelectionChanged → FindSimilarPair 的 O(N) 配對掃描 (原本每 100ms 一次)
                    _suppressPairing = True
                    .SelectedItems.Clear()
                    lastItem.Selected = True
                    lastItem.Focused = True
                    _suppressPairing = False
                End If
            End With
        End If

    End Sub
    Private Function BuildListViewItem(pending As PendingDebugMsg) As ListViewItem
        ' 2026/07/11 by Simon/Sonnet 5: ListViewItem/DebugItemTag 建構邏輯從 AddMessage3 移到這裡，
        ' 在 Timer_Tick 批次處理時才建構，讓呼叫端只需組字串+enqueue輕量DTO (PendingDebugMsg)
        Dim newItem As New ListViewItem(pending.msgContent)
        newItem.SubItems.Add(pending.timeNow.ToString("HH:mm:ss.ff"))
        newItem.SubItems.Add(If(pending.lineNo > 1, pending.timeSpan.TotalMilliseconds.ToString("#,##0.00"), "-"))  ' Index 2: Step (物理間隔)
        newItem.SubItems.Add("")                                                                                    ' Index 3: Elapsed (邏輯耗時，預設空，由上方配對邏輯填入)

        ' 2026/3/28 by Gemini: 預先計算快取資訊存入tag備用
        Dim tag As New DebugItemTag()
        tag.timeStamp = pending.timeNow                       ' 2026/3/28 by Gemini: 保留原始精度供日後計算
        tag.isBeginRow = pending.msgContent.Contains("開始")  ' 2026/07/11 by Simon/Sonnet 5: FindSimilarPair 配對搜尋鍵在此算一次存入 Tag，
        tag.isEndRow = pending.msgContent.Contains("結束")    ' 取代原本每次配對搜尋都對整個 ListView 逐列重跑 RemoveBeginEnd (Substring+6xReplace+Trim)
        tag.coreKey = RemoveBeginEnd(pending.msgContent)
        newItem.Tag = tag

        ' 2026/07/11 by Simon/Sonnet 5: textFullRow 延遲建構 — 只有搜尋作用中才需要組這段字串+ToLower，
        ' 沒開搜尋 (最常見情況) 完全跳過，等 RefreshSearchCache 真的需要時才補建
        If _searchPattern.Length > 0 Then
            tag.textFullRow = BuildFullRowText(newItem)
            tag.isHit = CheckIsHitInternal(tag.textFullRow)
        End If

        Return newItem

    End Function
    Private Function BuildFullRowText(item As ListViewItem) As String
        ' 2026/07/11 by Simon/Sonnet 5: 抽出共用 helper，取代原本 AddMessage3 與 Timer_Tick 兩處重複的字串串接邏輯
        Return (item.Text & " " & item.SubItems(1).Text & " " & item.SubItems(2).Text & " " & item.SubItems(3).Text).ToLower()

    End Function
#End Region

#Region "■ 05 ListView 操作事件"
    Private Sub lvwDebug_ItemSelectionChanged(sender As Object, e As ListViewItemSelectionChangedEventArgs) Handles lvwDebug.ItemSelectionChanged

        ' by Gemini, 2026/04/01: 解決點選配對項目時的閃爍問題 (Flickering) 不直接修改 ListViewItem.BackColor 屬性，改為記錄目標並用 Invalidate() 局部重繪。
        ' 2026/07/11 by Simon/Sonnet 5: Timer_Tick 的自動捲動選取不需要配對高亮，直接跳過整段掃描
        If _suppressPairing Then Return
        If Not e.IsSelected Then Return

        ' by Gemini, 2026/03/29: 選取變更時自動標記配對的「開始/結束」行
        ' 效能優化: 使用 _lastHighlightedPair 做 O(1) 顏色還原，取代原本的 For Each 全域清除。原本 Shift 多選 100 筆時會觸發 100 次事件 × N 筆 = O(N²) 重繪，改為 O(1)
        ' 2026/04/01 by Gemini: 效能閥值管理, 防止 Shift 多選上千筆時產生 O(N²) 的效能雪崩 (延遲)
        ' 當使用者框選多筆資料時，配對高光沒有意義，直接清除高光並 Return 離開，省下幾百萬次的字串比對
        If lvwDebug.SelectedIndices.Count > 1 Then
            If _lastHighlightedPair IsNot Nothing Then
                Dim oldP As ListViewItem = _lastHighlightedPair
                _lastHighlightedPair = Nothing
                If oldP.ListView IsNot Nothing Then lvwDebug.Invalidate(oldP.Bounds)
            End If
            Return
        End If

        ' O(1) 還原上次標記的顏色 (現在改為只記錄並稍後 Invalidate，不再直接塞 Color.White)
        ' 只針對「上一筆」被高亮的項目進行處理，不走全表掃描，準備稍後進行快狀局部重繪。
        Dim oldPair As ListViewItem = _lastHighlightedPair
        _lastHighlightedPair = Nothing

        ' 雙向配對搜尋 (支援巢狀 Stack 計數)
        Dim selectedItem As ListViewItem = e.Item
        Dim newPair As ListViewItem = FindSimilarPair(selectedItem)
        If newPair IsNot Nothing Then _lastHighlightedPair = newPair ' 記住這次標記，下次進來時只需還原這一個

        ' 💡 針對舊的與新的配對項目，僅發送局部重繪指令，避免觸發全表排版與屬性變更重繪
        If oldPair IsNot Nothing AndAlso oldPair.ListView IsNot Nothing Then lvwDebug.Invalidate(oldPair.Bounds)
        If newPair IsNot Nothing AndAlso newPair.ListView IsNot Nothing Then lvwDebug.Invalidate(newPair.Bounds)

    End Sub
    Private Sub lvwDebug_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles lvwDebug.MouseDoubleClick

        ' 雙擊: 複製該行完整文字到剪貼簿，並重算與配對 Begin/End 的時間差
        If e.Button <> MouseButtons.Left OrElse e.Clicks <> 2 Then Return
        Dim selectedItem As ListViewItem = sender.GetItemAt(e.X, e.Y)
        If selectedItem Is Nothing Then Return

        ' by Gemini, 2026/04/05: ── 標記功能 ──
        ' 直接切換項目底色：橘色 ↔ 白色
        ' DrawSubItem 本來就讀 e.Item.BackColor 來畫底色，這裡直接改它最簡單
        ' (低頻雙擊操作，不會有閃爍問題)
        If selectedItem.BackColor = Color.FromArgb(255, 140, 0) Then
            selectedItem.BackColor = Color.White
            selectedItem.ForeColor = Color.Black
        Else
            selectedItem.BackColor = Color.FromArgb(255, 140, 0) ' 亮橘色 #FF8C00
            selectedItem.ForeColor = Color.White                  ' 白字確保對比度
        End If

        ' ✅ 複製該行完整文字 (Tab 分隔三欄) 到剪貼簿
        Dim fullText As String = String.Join(vbTab, selectedItem.SubItems.Cast(Of ListViewItem.ListViewSubItem)().Select(Function(s) s.Text))
        Clipboard.SetText(fullText)

        ' ✅ 重算時間差: 優先找配對的 Begin/End，其次嘗試與前一行比較 (by Gemini, 2026/03/31)
        ' 2026/3/28 by Gemini: 直接從 DebugItemTag.timeStamp 讀取原始時間，免去 TryParse 反解與精度損失
        Dim tagCurrent = TryCast(selectedItem.Tag, DebugItemTag)
        If tagCurrent Is Nothing Then Return
        Dim t_anchor As Date
        Dim anchorFound As Boolean = False

        ' 1. 優先嘗試尋找配對 (只限於對「結束」行尋找回頭的「開始」)
        ' by Gemini, 2026/03/31: 點擊「開始」應維持與前一行的間隔，只有點擊「結束」才計算程序總耗時
        If selectedItem.Text.Contains("結束") Then
            Dim pairItem As ListViewItem = FindSimilarPair(selectedItem)
            If pairItem IsNot Nothing Then
                Dim tagPair = TryCast(pairItem.Tag, DebugItemTag)
                If tagPair IsNot Nothing Then
                    t_anchor = tagPair.timeStamp
                    anchorFound = True
                End If
            End If
        End If

        ' 2. 若無配對且不是第一行，則與上一行對比 (fallback)
        If Not anchorFound AndAlso selectedItem.Index > 0 Then
            Dim tagPrev = TryCast(lvwDebug.Items(selectedItem.Index - 1).Tag, DebugItemTag)
            If tagPrev IsNot Nothing Then
                t_anchor = tagPrev.timeStamp
                anchorFound = True
            End If
        End If

        ' 3. 執行計算並更新 UI (使用絕對值確保正數顯示)
        If anchorFound Then
            Dim diffMs As Double = Math.Abs((tagCurrent.timeStamp - t_anchor).TotalMilliseconds)
            ' by Gemini, 2026/04/11: 根據配對來源決定填入哪一欄
            If selectedItem.Text.Contains("結束") Then
                selectedItem.SubItems(3).Text = diffMs.ToString("#,##0.00") ' 總耗時
            Else
                selectedItem.SubItems(2).Text = diffMs.ToString("#,##0.00") ' 物理間隔
            End If
        End If

    End Sub
    Private Sub lvwDebug_KeyDown(sender As Object, e As KeyEventArgs) Handles lvwDebug.KeyDown
        ' Ctrl+C: 複製所有已選取的行 (Tab 分隔欄位，vbNewLine 分隔行)
        ' D5 2026-03-23: 補實作多行複製，需 MultiSelect=True (於 DebugForm_Load 設定)
        If e.KeyCode = Keys.Enter Then
            CalculateSelectedTimeSpan(Nothing, Nothing)    ' Enter 鍵也觸發計算選取耗時的功能，方便快速查看

        ElseIf e.Control AndAlso e.KeyCode = Keys.C Then
            Dim selected = lvwDebug.SelectedItems.Cast(Of ListViewItem)().ToList()
            If selected.Count = 0 Then Return
            Dim lines = selected.Select(Function(item)
                                            Return String.Join(vbTab, item.SubItems.Cast(Of ListViewItem.ListViewSubItem)().Select(Function(s) s.Text))
                                        End Function)
            Clipboard.SetText(String.Join(Environment.NewLine, lines))
            e.Handled = True
        End If

    End Sub
    Private Sub lvwDebug_DrawColumnHeader(sender As Object, e As DrawListViewColumnHeaderEventArgs) Handles lvwDebug.DrawColumnHeader
        e.DrawDefault = True
    End Sub
    Private Sub lvwDebug_DrawItem(sender As Object, e As DrawListViewItemEventArgs) Handles lvwDebug.DrawItem
        ' 2026/03/31 by Gemini: 在 OwnerDraw=True 時若設為 True 會覆蓋掉 SubItem 的高亮，必須設為 False
        e.DrawDefault = False

    End Sub
    Private Sub lvwDebug_DrawSubItem(sender As Object, e As DrawListViewSubItemEventArgs) Handles lvwDebug.DrawSubItem
        ' =======================================================
        ' 完全自訂繪製，確保文字在有/無搜尋條件時位置絕對不跳動
        ' 2026/05/09 by Gemini 3 Flash: Resize 期間暫停繪製
        ' 2026/3/27 by Gemini (simon: 幹這些東西沒有用過好難)
        ' =======================================================

        e.DrawDefault = False
        ' ✅ 2026/03/31 by Gemini: 強制設定 Clip 區域。在極寬視窗 (>2000px) 且無捲軸時，
        ' GDI+ 可能因內部座標計算偏移而遺失繪圖，顯式 SetClip 可解決此問題。
        e.Graphics.SetClip(e.Bounds)

        ''' 2026/04/12 by Claude: 同時清除 GDI 系統 clip，確保 TextRenderer 不受 dirty region 限制
        ''' 否則 ScrollBar 消失時 dirty region 只有右側條帶，TextRenderer (GDI) 畫不出左側文字

        ' Step 1. Background
        Dim backColor As Color = e.Item.BackColor
        Dim foreColor As Color = e.Item.ForeColor
        ' 2026/04/01 by Gemini: 取代原本動態修改 Item.BackColor 的做法，改在渲染時直上高光
        ' (標記項目的橘色底色已直接存在 item.BackColor，自然被上面這行讀到，不需要額外判斷)
        If _lastHighlightedPair IsNot Nothing AndAlso e.Item Is _lastHighlightedPair Then
            backColor = Color.Cyan   ' 配對「開始/結束」行的高亮（此為渲染時注入，不污染 BackColor 屬性）
        End If
        If e.Item.Selected Then
            backColor = SystemColors.Highlight
            foreColor = SystemColors.HighlightText
        End If
        'Using brush As New SolidBrush(backColor) : e.Graphics.FillRectangle(brush, e.Bounds) : End Using
        ' Tier 3, 2026/06/15 by Simon/Claude: 重用 member brush，改 .Color 取代每格 New SolidBrush/Dispose
        _fillBrush.Color = backColor : e.Graphics.FillRectangle(_fillBrush, e.Bounds)

        Dim itemText As String = e.SubItem.Text
        If String.IsNullOrEmpty(itemText) Then Return

        ' Step 2. Alignment Flags
        Dim align As HorizontalAlignment = lvwDebug.Columns(e.ColumnIndex).TextAlign
        Dim flags As TextFormatFlags = TextFormatFlags.VerticalCenter Or TextFormatFlags.PreserveGraphicsClipping Or
                                       TextFormatFlags.NoPrefix Or TextFormatFlags.NoPadding
        ' 2026/3/27 by Gemini: 必須強制使用 NoPadding 才能跟 MeasureText(NoPadding) 的手動座標完全對齊
        ' Text bounds (一致的 6px 留白，模擬預設繪製但不產生跳躍)
        Dim textRect As Rectangle = e.Bounds : textRect.Inflate(-6, 0)

        Dim tag = TryCast(e.Item.Tag, DebugItemTag)
        Dim isHitCell As Boolean = False
        Dim matches As System.Text.RegularExpressions.MatchCollection = Nothing

        ' 2026/3/28 by Gemini: 讀取快取狀態與預先定義好的 Regex 模式，達成 O(1) 繪製準備
        ' Tier 3, 2026/06/15 by Simon/Claude: 改用預編譯的 _searchRegex 實例，取代每格 Regex.Matches 字串多載
        If tag IsNot Nothing AndAlso tag.isHit AndAlso _searchRegex IsNot Nothing Then
            matches = _searchRegex.Matches(itemText)
            If matches IsNot Nothing AndAlso matches.Count > 0 Then isHitCell = True
        End If

        ' Step 3. Draw Text
        If Not isHitCell Then
            Select Case align
                Case HorizontalAlignment.Center : flags = flags Or TextFormatFlags.HorizontalCenter
                Case HorizontalAlignment.Right : flags = flags Or TextFormatFlags.Right
                Case Else : flags = flags Or TextFormatFlags.Left
            End Select
            ' 無論有無命中，均使用相同對齊與邊界，確保不跳動
            TextRenderer.DrawText(e.Graphics, itemText, e.Item.Font, textRect, foreColor, flags Or TextFormatFlags.WordEllipsis)
        Else
            ' Highlighted drawing
            Dim totalSize = TextRenderer.MeasureText(e.Graphics, itemText, e.Item.Font, New Size(Integer.MaxValue, e.Bounds.Height), TextFormatFlags.NoPadding)
            Dim currentX As Integer
            Select Case align
                Case HorizontalAlignment.Right : currentX = textRect.Right - totalSize.Width
                Case HorizontalAlignment.Center : currentX = textRect.X + (textRect.Width - totalSize.Width) \ 2
                Case Else : currentX = textRect.X
            End Select
            Dim lastPos As Integer = 0
            For Each m As System.Text.RegularExpressions.Match In matches
                ' 繪製命中前的普通文字
                If m.Index > lastPos Then
                    Dim normalPart As String = itemText.Substring(lastPos, m.Index - lastPos)
                    Dim szNormal = TextRenderer.MeasureText(e.Graphics, normalPart, e.Item.Font, New Size(Integer.MaxValue, e.Bounds.Height), TextFormatFlags.NoPadding)

                    ' 💡 2026/04/13 by Gemini 3 Flash: 邊界防禦 — 若即將溢出，則強制截斷並退出
                    If currentX + szNormal.Width > textRect.Right Then
                        TextRenderer.DrawText(e.Graphics, normalPart, e.Item.Font, New Rectangle(currentX, e.Bounds.Y, textRect.Right - currentX, e.Bounds.Height), foreColor, flags Or TextFormatFlags.EndEllipsis)
                        lastPos = itemText.Length : Exit For
                    End If

                    Dim rDraw As New Rectangle(currentX, e.Bounds.Y, szNormal.Width, e.Bounds.Height)
                    TextRenderer.DrawText(e.Graphics, normalPart, e.Item.Font, rDraw, foreColor, flags)
                    currentX += szNormal.Width
                End If

                ' 繪製高亮背景與文字
                Dim matchPart As String = m.Value
                Dim szMatch = TextRenderer.MeasureText(e.Graphics, matchPart, e.Item.Font, New Size(Integer.MaxValue, e.Bounds.Height), TextFormatFlags.NoPadding)

                ' 💡 2026/04/13 by Gemini 3 Flash: 邊界防禦 (高亮塊)
                If currentX + szMatch.Width > textRect.Right Then
                    ' Tier 3, 2026/06/15 by Simon/Claude: 改用 framework 快取的 Brushes.Yellow，零配置
                    ' 2026/07/11 by Simon/Sonnet 5: 移除多餘的 Using New SolidBrush 外殼 (實際繪製早已改用 Brushes.Yellow，外殼只是浪費一次 New+Dispose)
                    e.Graphics.FillRectangle(Brushes.Yellow, New Rectangle(currentX, e.Bounds.Y + 2, textRect.Right - currentX, e.Bounds.Height - 4))
                    TextRenderer.DrawText(e.Graphics, matchPart, e.Item.Font, New Rectangle(currentX, e.Bounds.Y, textRect.Right - currentX, e.Bounds.Height), Color.Black, flags Or TextFormatFlags.EndEllipsis)
                    lastPos = itemText.Length : Exit For
                End If

                ' Tier 3, 2026/06/15 by Simon/Claude: 同上，改用 Brushes.Yellow
                ' 2026/07/11 by Simon/Sonnet 5: 移除多餘的 Using New SolidBrush 外殼
                e.Graphics.FillRectangle(Brushes.Yellow, New Rectangle(currentX, e.Bounds.Y + 2, szMatch.Width, e.Bounds.Height - 4))
                Dim rMatch As New Rectangle(currentX, e.Bounds.Y, szMatch.Width, e.Bounds.Height)
                TextRenderer.DrawText(e.Graphics, matchPart, e.Item.Font, rMatch, Color.Black, flags)
                currentX += szMatch.Width
                lastPos = m.Index + m.Length
            Next

            ' 繪製剩餘文字
            If lastPos < itemText.Length Then
                Dim remainingPart As String = itemText.Substring(lastPos)
                Dim szRemaining = TextRenderer.MeasureText(e.Graphics, remainingPart, e.Item.Font, New Size(Integer.MaxValue, e.Bounds.Height), TextFormatFlags.NoPadding)

                ' 💡 2026/04/13 by Gemini 3 Flash: 邊界防禦 (剩餘塊)
                If currentX + szRemaining.Width > textRect.Right Then
                    TextRenderer.DrawText(e.Graphics, remainingPart, e.Item.Font, New Rectangle(currentX, e.Bounds.Y, textRect.Right - currentX, e.Bounds.Height), foreColor, flags Or TextFormatFlags.EndEllipsis)
                Else
                    Dim rRem As New Rectangle(currentX, e.Bounds.Y, szRemaining.Width, e.Bounds.Height)
                    TextRenderer.DrawText(e.Graphics, remainingPart, e.Item.Font, rRem, foreColor, flags)
                End If
            End If
        End If
        ' ✅ 恢復 Clip 區域
        e.Graphics.ResetClip()

    End Sub
#End Region

#Region "■ 06 其他UI操作事件"
    Private Sub txtDebug_TextChanged(sender As Object, e As EventArgs) Handles txtDebug.TextChanged
        ' 搜尋字串變更: 重繪 ListView
        UpdateSearchCaption() ' by Gemini, 2026/03/31: 同步更新視窗標題
        RefreshSearchCache()  ' 2026/3/28 by Gemini: 批次更新快取
        lvwDebug.Invalidate() ' 2026/07/11 by Simon/Sonnet 5: 改用 Invalidate() 讓 Windows 合併重繪，取代 Refresh() 的強制同步重繪

    End Sub
    Private Sub chkSearchLogic_CheckedChanged(sender As Object, e As EventArgs) Handles checkAndOr.CheckedChanged
        checkAndOr.Text = If(checkAndOr.Checked, "AND", "OR")
        UpdateSearchCaption() ' by Gemini, 2026/03/31: 切換模式時同步更新視窗標題
        RefreshSearchCache()  ' 2026/3/28 by Gemini: 批次更新快取
        lvwDebug.Invalidate() ' 2026/07/11 by Simon/Sonnet 5: 改用 Invalidate() 讓 Windows 合併重繪，取代 Refresh() 的強制同步重繪

    End Sub
    Private Sub UpdateSearchCaption()
        ''' <summary>
        ''' 2026/03/31 by Gemini: 集中標題管理邏輯，解決文字清空未回復、及 logic 切換未同步問題
        ''' </summary>
        If String.IsNullOrWhiteSpace(txtDebug.Text) Then
            Me.Text = "執行期除錯視窗"
        Else
            Dim logic As String = If(checkAndOr.Checked, "AND", "OR")
            Me.Text = $"除錯視窗 - 搜尋比對 ({logic}): {txtDebug.Text}"
        End If

    End Sub
#End Region

#Region "■ 07 輔助函數"
    Private Sub txtDebug_KeyDown(sender As Object, e As KeyEventArgs)
        ''' <summary>
        ''' 2026/04/07 by AntiGravity: 攔截搜尋框按鍵，支援 Enter 紀錄歷史與上下鍵回溯
        ''' </summary>
        Select Case e.KeyCode
            Case Keys.Enter
                AddToHistoryDebug(CType(sender, TextBox).Text)
                e.SuppressKeyPress = True
            Case Keys.Up
                NavigateHistoryDebug(-1)
                e.Handled = True
            Case Keys.Down
                NavigateHistoryDebug(1)
                e.Handled = True
        End Select
    End Sub
    Private Sub AddToHistoryDebug(query As String)
        ''' <summary>
        ''' 2026/04/07 by AntiGravity: 將關鍵字存入歷史紀錄 (去重與限額 50 筆)
        ''' </summary>
        Dim trimmed = query.Trim()
        If String.IsNullOrEmpty(trimmed) Then Return
        If _historyDebug.Count = 0 OrElse _historyDebug.Last() <> trimmed Then
            _historyDebug.Add(trimmed)
            If _historyDebug.Count > 50 Then _historyDebug.RemoveAt(0)
        End If
        _historyIndex = _historyDebug.Count
        _tempInput = ""
    End Sub
    Private Sub NavigateHistoryDebug(direction As Integer)
        ''' <summary>
        ''' 2026/04/07 by AntiGravity: 導覽歷史紀錄，並處理暫存原始輸入的邏輯
        ''' </summary>
        If _historyDebug.Count = 0 Then Return
        If _historyIndex = _historyDebug.Count AndAlso direction < 0 Then
            _tempInput = txtDebug.Text
        End If
        Dim targetIndex As Integer = _historyIndex + direction
        If targetIndex < 0 Then
            targetIndex = 0
        ElseIf targetIndex > _historyDebug.Count Then
            targetIndex = _historyDebug.Count
        End If
        If targetIndex = _historyIndex Then Return
        _historyIndex = targetIndex
        If _historyIndex = _historyDebug.Count Then
            txtDebug.Text = _tempInput
        Else
            txtDebug.Text = _historyDebug(_historyIndex)
        End If
        txtDebug.SelectionStart = txtDebug.Text.Length
    End Sub
    Private Sub CalculateSelectedTimeSpan(sender As Object, e As EventArgs)
        ' by Gemini, 2026/03/29: 加總選取項目各自的耗時間隔 (使用 .Tag.timeStamp)
        ' 每個項目的耗時 = 該項目的 timeStamp - ListView 中前一項的 timeStamp
        If lvwDebug.SelectedItems.Count = 0 Then
            MessageBox.Show("請至少選取 1 個項目") : Return
        End If
        Dim totalMs As Double = 0
        For Each item As ListViewItem In lvwDebug.SelectedItems
            Dim tag = TryCast(item.Tag, DebugItemTag)
            If tag Is Nothing OrElse item.Index = 0 Then Continue For
            ' 取 ListView 中的前一項 (不是前一個選取項) 計算間隔
            Dim prevTag = TryCast(lvwDebug.Items(item.Index - 1).Tag, DebugItemTag)
            If prevTag IsNot Nothing Then
                totalMs += (tag.timeStamp - prevTag.timeStamp).TotalMilliseconds
            End If
        Next
        MessageBox.Show($"已選擇 {lvwDebug.SelectedItems.Count} 個項目" & vbCrLf &
                        $"耗時加總: {totalMs:N0} ms ({totalMs / 1000:N2} s)", "計算結果")

    End Sub
    Private Sub DeleteSelectedItems(sender As Object, e As EventArgs)
        ' by Gemini, 2026/03/29: 刪除選取項目，剩餘項目自動往上遞補，行號序號保留不變
        ' 2026/07/11 by Simon/Sonnet 5: 修正大量刪除時崩潰 —
        '   原本邊列舉 SelectedItems 邊 Items.Remove()會在列舉過程中改動集合本身 (SelectedItems 是即時反映 Items 的視圖)，導致列舉失效。
        '   改成先複製索引並反向由大到小刪除，刪除較高索引不會影響尚未處理的較低索引。
        Dim indices() As Integer = lvwDebug.SelectedIndices.Cast(Of Integer)().OrderByDescending(Function(i) i).ToArray()
        If indices.Length = 0 Then Return

        lvwDebug.BeginUpdate()
        For Each idx As Integer In indices : lvwDebug.Items.RemoveAt(idx) : Next
        lvwDebug.EndUpdate()

    End Sub
    Private Sub RefreshSearchCache()
        ''' 2026/3/28 by Gemini: 根據目前關鍵字更新所有項目的 isHit 狀態，並預先產生 Regex 高亮模式
        ' by Gemini, 2026/04/03: 增加邏輯區隔空白
        If lvwDebug.Items.Count = 0 Then Return

        ' 預先產生 Regex 模式，徹底移除 DrawSubItem 中的 LINQ 與字串運算
        Dim keywords = ParseSearchKeywords(txtDebug.Text.Trim())
        _searchPattern = If(keywords.Count > 0, String.Join("|", keywords.OrderByDescending(Function(k) k.Length).Select(Function(kw) System.Text.RegularExpressions.Regex.Escape(kw))), "")

        ' Tier 3, 2026/06/15 by Simon/Claude: 僅在搜尋字串變動時重建 Regex 實例 (含 IgnoreCase)，供 DrawSubItem 重複使用，避免每格用字串多載重新解析 pattern
        _searchRegex = If(String.IsNullOrEmpty(_searchPattern), Nothing, New Regex(_searchPattern, RegexOptions.IgnoreCase))

        _cachedKeywordsLower = keywords.Select(Function(k) k.ToLower()).ToList()    ' 2026/07/11 by Simon/Sonnet 5: 關鍵字轉小寫一次存欄位，供 CheckIsHitInternal 重複使用；
        _cachedAndMode = checkAndOr.Checked                                         ' 2026/07/11 by Simon/Sonnet 5: 同時把 checkAndOr.Checked 也快取成欄位，讓 AddMessage3 (可能來自背景執行緒) 不必再讀 UI 控制項

        For Each lvi As ListViewItem In lvwDebug.Items
            Dim tag = TryCast(lvi.Tag, DebugItemTag)
            If tag IsNot Nothing Then
                ' 2026/07/11 by Simon/Sonnet 5: textFullRow 現在是延遲建構的 (沒開搜尋時不會預先組好)，
                ' 第一次真的需要搜尋時才在這裡補建一次，之後就常駐快取，不用每次都重組
                If tag.textFullRow Is Nothing Then tag.textFullRow = BuildFullRowText(lvi)
                tag.isHit = CheckIsHitInternal(tag.textFullRow)
            End If
        Next

    End Sub

    Private Function ParseSearchKeywords(searchText As String) As List(Of String)
        ' 使用 Regex 拆分搜尋關鍵字，支援雙引號括起來的片語
        ' Regex 模式: (?:""(?<q>[^""]*)""|(?<w>\S+))
        Dim keywords As New List(Of String)(16)
        If String.IsNullOrWhiteSpace(searchText) Then Return keywords
        Dim matches = _keywordSplitRegex.Matches(searchText)
        For Each m As System.Text.RegularExpressions.Match In matches
            If m.Groups("q").Success Then
                keywords.Add(m.Groups("q").Value)
            ElseIf m.Groups("w").Success Then
                keywords.Add(m.Groups("w").Value)
            End If
        Next
        Return keywords

    End Function
    Private Function CheckIsHitInternal(fullText As String) As Boolean
        ' 2026/3/28 by Gemini: 內部判斷邏輯
        ' 2026/07/11 by Simon/Sonnet 5: 改讀 RefreshSearchCache 預先準備好的 _cachedKeywordsLower/_cachedAndMode，
        ' 不再現場讀 txtDebug.Text / checkAndOr.Checked (AddMessage3 可能來自背景執行緒呼叫，直接碰 UI 控制項不安全)，也不再每個關鍵字都重新 ToLower()
        If _cachedKeywordsLower.Count = 0 Then Return False

        Return If(_cachedAndMode, _cachedKeywordsLower.All(Function(kw) fullText.Contains(kw)), ' AND
                                  _cachedKeywordsLower.Any(Function(kw) fullText.Contains(kw))) ' OR

    End Function
    Private Function FindSimilarPair(selectedItem As ListViewItem, Optional additionalItems As List(Of ListViewItem) = Nothing) As ListViewItem
        ' by Gemini, 2026/03/29: 巢狀雙向配對搜尋 (Stack 計數器演算法)
        ' by Gemini, 2026/04/11: 支援 additionalItems 參數，解決同一批 Timer 寫入時新項目未掛載 ListView 的配對問題。
        '   點選「開始」→ 向下找配對的「結束」
        '   點選「結束」→ 向上找配對的「開始」
        '   遇到同名的巢狀呼叫時，使用 depth 計數器確保配對到正確的層級
        '   _dbg("Start: This is a test.")
        '   _dbg("Enter: This is a test.")
        '   _dbg("Done: This is a test.")
        '   _dbg("Begin: This is a test.")
        '   _dbg("Ended: This is a test.")
        '   _dbg("Finish: This is a test.")

        ' 2026/07/11 by Simon/Sonnet 5: coreKey/isBeginRow/isEndRow 已在 AddMessage3 建立時算好存入 Tag，
        ' 這裡直接讀快取，取代原本每次配對搜尋都對逐一列重跑 RemoveBeginEnd (Substring+6xReplace+Trim)，
        ' 這是配對搜尋 (Timer_Tick 每則「結束」訊息、選取高亮、雙擊) 的主要 GC 來源
        Dim depth As Integer = 0
        Dim tagSel = DirectCast(selectedItem.Tag, DebugItemTag)
        Dim coreName As String = tagSel.coreKey
        Dim isBegin As Boolean = tagSel.isBeginRow
        Dim isEnd As Boolean = tagSel.isEndRow
        If Not isBegin AndAlso Not isEnd Then Return Nothing

        If isBegin Then
            ' 向下搜尋配對的「結束」
            For i As Integer = selectedItem.Index + 1 To lvwDebug.Items.Count - 1
                ' todo: debug: 這裡一直存取物件不會有性能問題嗎? 改成 with lvwDebug? OR??
                Dim item As ListViewItem = lvwDebug.Items(i)
                Dim itemTag = DirectCast(item.Tag, DebugItemTag)
                If IsContentSimilar(coreName, itemTag.coreKey) Then
                    If itemTag.isBeginRow Then
                        depth += 1                      ' 同名的巢狀開始，深度 +1
                    ElseIf itemTag.isEndRow Then
                        If depth = 0 Then Return item   ' 深度歸零 = 正確配對
                        depth -= 1                      ' 消耗一層巢狀
                    End If
                End If
            Next

        ElseIf isEnd Then
            ' 向上搜尋配對的「開始」
            ' 2026/07/11 by Simon/Sonnet 5: 原本這裡有兩段幾乎一模一樣的向回掃描迴圈 (待處理批次 / 已顯示 ListView)，
            ' 抽成 ScanBackwardForBegin 共用。depth 用 ByRef 傳遞，讓 Level 2 能接續 Level 1 已累積的巢狀深度繼續找 (與原邏輯等價)。

            ' 💡 Level 1: 先在「同一批批次(待處理)」清單中往回搜尋 (優先級最高，因為距離最近)
            If additionalItems IsNot Nothing Then
                Dim selfIdx As Integer = additionalItems.IndexOf(selectedItem)  ' 從 selectedItem 在清單中的位置往前找
                If selfIdx > 0 Then
                    Dim pairInBatch As ListViewItem = ScanBackwardForBegin(additionalItems, selfIdx - 1, coreName, depth)
                    If pairInBatch IsNot Nothing Then Return pairInBatch
                End If
            End If

            ' 💡 Level 2: 若在當前批次沒找到，再往「已顯示(ListView)」中搜尋
            ' 2026/04/11 by Gemini: 修正搜尋起點
            ' 當 selectedItem 尚未加入 ListView 時 (Timer_Tick 批次處理中)，Index 會是 -1。
            ' 此時應從 ListView 的最末端 (Items.Count - 1) 開始往回找。
            Dim startIdx As Integer = If(selectedItem.Index >= 0, selectedItem.Index - 1, lvwDebug.Items.Count - 1)
            Return ScanBackwardForBegin(lvwDebug.Items, startIdx, coreName, depth)
        End If

        Return Nothing

    End Function
    Private Function ScanBackwardForBegin(items As Collections.IList, startIdx As Integer, coreName As String, ByRef depth As Integer) As ListViewItem
        ' 2026/07/11 by Simon/Sonnet 5: FindSimilarPair「向上找開始」的共用邏輯，原本在 additionalItems 與 lvwDebug.Items 兩處各重複一份。
        ' 參數吃非泛型 IList 是因為 List(Of ListViewItem) 與 ListView.ListViewItemCollection 都實作 IList，但彼此沒有共同的泛型介面。
        For i As Integer = startIdx To 0 Step -1
            Dim item As ListViewItem = DirectCast(items(i), ListViewItem)
            Dim itemTag = DirectCast(item.Tag, DebugItemTag)
            If IsContentSimilar(coreName, itemTag.coreKey) Then
                If itemTag.isEndRow Then
                    depth += 1                      ' 同名的巢狀結束，深度 +1
                ElseIf itemTag.isBeginRow Then
                    If depth = 0 Then Return item   ' 深度歸零 = 正確配對
                    depth -= 1                      ' 消耗一層巢狀
                End If
            End If
        Next
        Return Nothing
    End Function
    Private Function RemoveBeginEnd(content As String) As String
        ' 2026/03/31 by Gemini: 強化提取比對核心 (Key) 的邏輯
        ' 1. 移除行號前綴 (第一個空格前)
        Dim idx As Integer = content.IndexOf(" "c)
        Dim result As String = If(idx >= 0, content.Substring(idx + 1), content)
        ' 2. 移除 開始/結束 標籤 (兼容 有/無冒號、全形/半形)
        result = result.Replace("開始: ", "").Replace("結束: ", "") _
                       .Replace("開始：", "").Replace("結束：", "") _
                       .Replace("開始", "").Replace("結束", "").Trim()
        ' 3. 去噪: 為了讓「開始 (參數A)」能配對到「結束 (結果B)」，
        '    在提取核心時直接移除結尾的整組括號。
        '    比對時重點在於「方法名稱」與「狀態標記」是否一致。
        If result.EndsWith(CChar(")")) Then
            Dim lastOpenParen As Integer = result.LastIndexOf(CChar("("))
            If lastOpenParen >= 0 Then
                result = result.Substring(0, lastOpenParen).Trim()
            End If
        End If
        Return result

    End Function
    Private Function IsContentSimilar(content1 As String, content2 As String) As Boolean
        ' IsContentSimilar: 判斷兩段文字是否相似 (完全相符 or 包含關係)
        If content1 = content2 Then Return True
        Return content1.Contains(content2) OrElse content2.Contains(content1)

    End Function
#End Region

#Region "■ 99 舊版備用 (勿刪)"
    Public Sub AddMessage(Optional strA As String = "", Optional strB As String = "")
        ' 添加項目的方法
        Static lineCount As Integer : lineCount += 1
        'Dim callingMethod As String = GetActualCallingMethod()
        'Dim newItem As New ListViewItem($"{lineCount.ToString("000")} {strA} {callingMethod} ({strB})")
        ''newItem.Tag = Now
        'lvwDebug.Items.Add(newItem)
        'TriggerItemAddedEvent(newItem)

    End Sub
    Public Sub AddMessage2(Optional strA As String = "", Optional strB As String = "")
        ' =======================================================
        ' Claude AI 的優化建議: 將所有欄位準備好後再一次性加入 ListView，減少重繪次數，提升性能。
        ' 合併了AddMessage 和 UpdateListViewItem 的邏輯，避免了不必要的事件觸發和跨線程調用，簡化了代碼結構。
        ' TriggerItemAddedEvent 和 OnItemAddedAsync 也可以整個刪掉，邏輯已合併進來。
        ' 2026/3/6, by Claude.ai
        ' =======================================================
        Static lineCount As Integer : lineCount += 1
        'Dim callingMethod As String = GetActualCallingMethod()
        'Dim timeNow As Date = Now
        'Dim timeSpan As TimeSpan = timeNow - previousTimestamp
        'previousTimestamp = timeNow
        '' ✅ 先把所有欄位準備好，再一次性加入並重繪
        'Dim newItem As New ListViewItem($"{lineCount.ToString("000")} {strA} {callingMethod} ({strB})")
        'newItem.SubItems.Add(timeNow.ToString("HH:mm:ss.fff"))
        'newItem.SubItems.Add(timeSpan.TotalMilliseconds.ToString("#,##0.000  "))
        'newItem.Tag = timeNow
        '' ✅ BeginUpdate 包住整個 Add，只重繪一次
        'lvwDebug.BeginUpdate()
        'lvwDebug.Items.Add(newItem)
        'newItem.EnsureVisible()
        'lvwDebug.EndUpdate()

    End Sub
    Private Sub TriggerItemAddedEvent(newItem As ListViewItem)
        ' 觸發項目新增事件
        Task.Run(Async Function()
                     Await OnItemAddedAsync(newItem)
                 End Function)

    End Sub
    Private Async Function OnItemAddedAsync(newItem As ListViewItem) As Task
        ' todo: 處理項目新增的非同步方法
        Await Task.Yield()
        'Dim timeNow As Date = Now
        'Dim timeSpan As TimeSpan = timeNow - previousTimestamp
        'If Me.InvokeRequired Then
        '    Me.Invoke(New Action(Sub() UpdateListViewItem(newItem, timeNow, timeSpan)))
        'Else
        '    UpdateListViewItem(newItem, timeNow, timeSpan)
        'End If
        'previousTimestamp = timeNow

    End Function
    Private Sub UpdateListViewItem(newItem As ListViewItem, currentTimestamp As Date, timeSpan As TimeSpan)
        ' 更新 ListView 項目的方法
        lvwDebug.BeginUpdate()
        With newItem
            .SubItems.Add(currentTimestamp.ToString("HH:mm:ss.ff"))
            '.SubItems.Add(timeSpan.TotalSeconds.ToString("F6"))            ' 顯示秒, 到小數點後六位
            .SubItems.Add(timeSpan.TotalMilliseconds.ToString("#,##0.00  ")) ' 顯示ms, 到小數點後三位
            .EnsureVisible()
            .Tag = currentTimestamp
            '.SubItems(0).Tag = timeNow
            '.SubItems(1).Tag = timeSpan
        End With
        lvwDebug.EndUpdate()

    End Sub
#End Region

End Class
