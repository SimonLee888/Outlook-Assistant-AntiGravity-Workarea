Imports System.Diagnostics
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
    ' 2026/04/18 by Claude: 與 Form1 相同的雙緩衝設定
    ' DebugForm 開啟時主要卡頓來源是 Timer_Tick 高頻觸發 BeginUpdate/EndUpdate + EnsureVisible，
    ' 這兩項設定可改善切換焦點與 Resize 時的撕裂感，但無法根治高頻更新本身的開銷。
    Protected Overrides ReadOnly Property CreateParams As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            cp.ExStyle = cp.ExStyle Or &H2000000    ' WS_EX_COMPOSITED：子控制項合成層雙緩衝
            Return cp
        End Get
    End Property
    Protected Overrides Sub OnLoad(e As EventArgs)
        Me.DoubleBuffered = True    ' Form 自身 WM_PAINT 雙緩衝，與 WS_EX_COMPOSITED 互補無衝突
        MyBase.OnLoad(e)
    End Sub
#End Region

#Region "■ 01 Win32 API & 常數"
    <Runtime.InteropServices.DllImport("user32.dll")>
    Private Shared Function SendMessage(
        hWnd As IntPtr,
        msg As Integer,
        wParam As IntPtr,
        lParam As IntPtr) As IntPtr
    End Function
    <Runtime.InteropServices.DllImport("gdi32.dll")>
    Private Shared Function CreateRectRgn(x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer) As IntPtr : End Function
    <Runtime.InteropServices.DllImport("gdi32.dll")>
    Private Shared Function SelectClipRgn(hDC As IntPtr, hRgn As IntPtr) As Integer : End Function
    <Runtime.InteropServices.DllImport("gdi32.dll")>
    Private Shared Function DeleteObject(hObject As IntPtr) As Boolean : End Function

    Private Const WM_SETREDRAW As Integer = &HB  ' 2026/3/26 by Gemini
    Private Const WM_SETFONT As Integer = &H30
    Private Const WM_SIZE As Integer = &H5       ' by Claude Opus 4.6, 2026/04/11: 攔截視窗尺寸變更
    Private Const SIZE_MAXIMIZED As Integer = 2
    Private Const SIZE_RESTORED As Integer = 0
    Private Const LVM_FIRST As Integer = &H1000
    Private Const LVM_GETHEADER As Integer = LVM_FIRST + 31
    Private Const LVM_SETEXTENDEDLISTVIEWSTYLE As Integer = LVM_FIRST + 54
    Private Const LVM_GETEXTENDEDLISTVIEWSTYLE As Integer = LVM_FIRST + 55
    Private Const LVM_SETTOOLTIPS As Integer = LVM_FIRST + 74            ' by Gemini 3 Flash, 2026/04/13: 用於切斷 ToolTip 控制項關聯
    Private Const LVS_EX_LABELTIP As Integer = &H4000       ' by Claude Opus 4.6, 2026/04/11: 移除此樣式以修復 OwnerDraw 下的文字重疊殘影
    Private Const LVS_EX_DOUBLEBUFFER As Integer = &H10000  ' by Claude, 2026/04/12: Native ListView 真正的雙緩衝 flag，解決 ScrollBar 消失時 OwnerDraw dirty region 只有右側條帶導致項目消失的 Bug
#End Region

#Region "■ 02 成員變數"
    Private sw0, sw1, sw2, sw3, sw4, sw5, sw6 As New Stopwatch
    Private _previousTimestamp As Date
    Private _msgQueue As New System.Collections.Concurrent.ConcurrentQueue(Of ListViewItem)
    Private WithEvents QueueTimer As New System.Windows.Forms.Timer() With {.Interval = 100} ' 每 100ms 清空一次message queue
    Private _lastRecalcWidth As Integer = 0

    Private _searchPattern As String = ""
    Private _searchRegex As Regex = Nothing             ' Tier 3, 2026/06/15 by Simon/Claude: 由 _searchPattern 預編譯, DrawSubItem 直接套用, 省去每格字串多載重複解析
    Private _fillBrush As New SolidBrush(Color.White)   ' Tier 3, 2026/06/15 by Simon/Claude: 背景填色重用同一支 brush (改 .Color 即可), 取代每格 New SolidBrush/Dispose

    Private _lastHighlightedPair As ListViewItem        ' by Gemini, 2026/03/29: O(1) 顏色還原，取代 For Each 全域清除
    Private _historyDebug As New List(Of String)(256)   ' by AntiGravity, 2026/04/07: 搜尋歷史紀錄
    Private _historyIndex As Integer = 0                ' by AntiGravity, 2026/04/07: 目前歷史紀錄索引 (與 Count 相同時代表原始輸入區)
    Private _tempInput As String = ""                   ' by AntiGravity, 2026/04/07: 暫存回溯前的原始輸入內容
    Private Class DebugItemTag          ' 2026/3/28 by Gemini: 定義快取結構，加速 OwnerDraw 繪製
        Public textFullRow As String    ' 預先合併好的整行小寫文字 (用於搜尋)
        Public isHit As Boolean         ' 是否命中目前搜尋關鍵字
        Public timeStamp As Date        ' 2026/3/28 by Gemini: 原始時間戳記 (供雙擊重算時間差，免去 TryParse 反解)
    End Class
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
        End If
    End Sub
    Private Sub ApplyListViewFixes()
        ''' <summary>
        ''' 2026/04/13 by Gemini 3 Flash: 集中處理 Win32 樣式修復 (LabelTip 移除、隔離 ToolTip、原生雙緩衝)
        ''' </summary>
        Try
            ' 1. 移除 LVS_EX_LABELTIP (防止 OwnerDraw 時出現鬼影文字標籤)
            '       💡 2026/04/11 by Claude Opus 4.6 移除 LVS_EX_LABELTIP 擴充樣式
            '       在 OwnerDraw 模式下， ListView 內建的「文字超寬時浮出全文標籤」會跟自訂繪製衝突，
            '       產生滑鼠移過時的文字重疊殘影。移除此樣式即可根治。
            '       wParam = mask(指定要修改哪些位元), lParam = 0 (關閉這些位元)
            ' by Gemini, 2026/04/13: 全部整合到 ApplyListViewFixes，確保 Handle 重建後依然生效
            SendMessage(lvwDebug.Handle, LVM_SETEXTENDEDLISTVIEWSTYLE, New IntPtr(LVS_EX_LABELTIP), IntPtr.Zero)

            ' 2. 徹底隔離 ToolTip 控制項 (切斷與系統調度器的關聯)
            SendMessage(lvwDebug.Handle, LVM_SETTOOLTIPS, IntPtr.Zero, IntPtr.Zero)

            ' 3. 啟用 LVS_EX_DOUBLEBUFFER (強化滾動與 Resize 時的渲染穩定性)
            '       💡 2026/04/12 by Claude 啟用 LVS_EX_DOUBLEBUFFER
            '       這是 Native Win32 ListView 的真正雙緩衝， 與.NET DoubleBuffered 屬性完全不同
            '       啟用後 ListView 每次 WM_PAINT 都對整個 client area 做 offscreen buffer blit，
            '       不再做 partial dirty-region paint， 從根本解決 ScrollBar 消失時 DrawSubItem 不被呼叫的問題
            ' by Gemini, 2026/04/13: 全部整合到 ApplyListViewFixes，確保 Handle 重建後依然生效
            SendMessage(lvwDebug.Handle, LVM_SETEXTENDEDLISTVIEWSTYLE,
                        New IntPtr(LVS_EX_DOUBLEBUFFER), New IntPtr(LVS_EX_DOUBLEBUFFER))
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
        ' 2026/04/01 by Gemini: 恢復 ListView 內建雙緩衝設置
        ' 先前為了排查 2000px 高度 Bug 暫時移除，現已確認該 Bug 兇手為 ClientSizeChanged 內的 BeginUpdate。
        ' 恢復此設定可徹底避免 AddMessage3 (Timer 批次新增) 時產生的背景擦除閃爍。
        Dim pi = lvwDebug.GetType().GetProperty("DoubleBuffered", BindingFlags.Instance Or BindingFlags.NonPublic)
        If pi IsNot Nothing Then pi.SetValue(lvwDebug, True, Nothing)

        _previousTimestamp = Now
        QueueTimer.Start()          ' .Interval = 100

        ' 💡 2026/04/13 by Gemini 3 Flash: 註冊 HandleCreated，確保 ListView 重建時修復依然生效
        AddHandler lvwDebug.HandleCreated, AddressOf OnLvwHandleCreated

    End Sub
    Private Sub DebugForm_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        ' ==============================================================
        ' by Gemini, 2026/04/01: 將重型 UI 佈局校算移到 Shown 事件
        ' 目的: 讓 Form1 觸發開啟除錯視窗後能立即返回，不等待 UI 佈局渲染，優化啟動延遲感
        ' ==============================================================
        ' todo: 重構簡化formload? InitSearchPanel() : InitListView() : InitLayout()

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
            '.Insert(0, New ColumnHeader() With {.Text = "Debug Message", .Width = -2,    ' 2026/3/28 by Gemini: Width=-2 讓第一欄自動填滿剩餘空間，避免寫死寬度在 Load 時擠掉右側欄位
            '                                    .TextAlign = HorizontalAlignment.Left})
        End With

        RecalcColumnWidths(Nothing, Nothing)    ' 2026/3/30 by Gemini: 在 Load 時手動觸發強制調整一次，確保初始顯示正確 (特別是第一欄填滿剩餘空間)

        AddHandler lvwDebug.ItemSelectionChanged, AddressOf lvwDebug_ItemSelectionChanged
        AddHandler lvwDebug.ClientSizeChanged, AddressOf RecalcColumnWidths
        AddHandler txtDebug.KeyDown, AddressOf txtDebug_KeyDown ' by AntiGravity, 2026/04/07: 支持搜尋歷史回溯

        ' 2026/3/28 by Gemini: 監聽 lvwDebug 本身的 ClientSizeChanged 事件，
        ' 無論何時 ListView 可用空間改變 (Dock 佈局結算、表單 Resize、SyncDebugFormPosition)，都自動重算欄寬, 不再需要猜延遲值或一次性 Timer
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

    End Sub
    Private Sub DebugForm_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form1.CheckDebug.Checked = False
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
        '' ── 為什麼 Invalidate() 無效 ──
        '' 在 ClientSizeChanged resize 訊息鏈進行中，Windows 會抑制 WM_PAINT 派發。
        '' Invalidate() 只是排入 WM_PAINT 到佇列，等 resize 結束時 item bounds 快取早已壞掉。
        '' ── 為什麼 BeginInvoke + Refresh() 有效 ──
        '' BeginInvoke 將 delegate 排入訊息泵，保證在 resize 訊息鏈**完全結束後**才執行。
        '' Refresh() = Invalidate() + Update()，Update() 同步處理 WM_PAINT，不會被延遲。
        'If newWidth > 100 AndAlso lvwDebug.Columns(0).Width <> newWidth Then
        '    lvwDebug.Columns(0).Width = newWidth
        '    BeginInvoke(Sub() If lvwDebug IsNot Nothing AndAlso Not lvwDebug.IsDisposed AndAlso lvwDebug.Items.Count > 0 Then lvwDebug.Refresh())
        'End If

        '' 2026/04/12 by Claude: 修復 ScrollBar 消失時 ListView 項目消失的 Bug
        '' 根本原因：Columns(0).Width 賦值會觸發 ListView 內部同步 repaint，
        '' 此時 DoubleBuffer backbuffer 被清空，但 GDI 系統 clip 只有右側17px條帶，
        '' 導致 TextRenderer.DrawText (GDI) 被 clip 住畫不出文字。
        '' 解法：用 WM_SETREDRAW 壓住 column 改變觸發的內部 paint，
        '' 改成寬度設定完後呼叫一次完整 Refresh()，此時 dirty region 是全區域。
        'If newWidth > 100 AndAlso lvwDebug.Columns(0).Width <> newWidth Then
        '    SendMessage(lvwDebug.Handle, WM_SETREDRAW, New IntPtr(0), IntPtr.Zero)
        '    lvwDebug.Columns(0).Width = newWidth

        '    SendMessage(lvwDebug.Handle, WM_SETREDRAW, New IntPtr(1), IntPtr.Zero)
        '    lvwDebug.Refresh()  ' Invalidate() + Update()，同步執行，此時 clip = 完整區域
        '    ' 移除原本的 BeginInvoke(Refresh()) — 由上面的同步 Refresh() 取代
        'End If

        ' 2026/04/12 by Claude: ScrollBar 消失後 EnsureVisible 殘留的 scroll offset 未歸零
        ' 導致 item(0).Bounds.Y 為負數，所有項目偏移至底部，Refresh 在錯誤座標執行也無效
        ' 檢查 item(0).Bounds.Y：不等於 0 代表 scroll offset 殘留，強制設 TopItem 歸零
        If newWidth > 100 AndAlso lvwDebug.Columns(0).Width <> newWidth Then
            lvwDebug.Columns(0).Width = newWidth
            BeginInvoke(Sub()
                            If lvwDebug Is Nothing OrElse lvwDebug.IsDisposed OrElse lvwDebug.Items.Count = 0 Then Return
                            Try
                                If lvwDebug.Items(0).Bounds.Y <> 0 Then lvwDebug.TopItem = lvwDebug.Items(0)
                            Catch : End Try
                            lvwDebug.RedrawItems(0, lvwDebug.Items.Count - 1, False)
                            lvwDebug.Update()
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

        ' 2026/3/22 by Grok.ai:
        ' forcedCaller: Form1.WhoCallsMe() 預先解析好的呼叫者字串 (避免 stack trace 在 DebugForm 裡走不回去)
        Dim callingMethod As String = If(forcedCaller <> "", forcedCaller, WhoCallsMe(1))

        ' 計算時間差 (todo: 要不要改成與上一行的時間差? 現在是與上次 AddMessage3 的時間差)
        Dim timeNow As Date = Now
        Dim timeSpan As TimeSpan = timeNow - _previousTimestamp
        _previousTimestamp = timeNow

        Static lineCount As Integer
        Dim newLine As Integer = System.Threading.Interlocked.Increment(lineCount)

        ' 2026/03/31 by Gemini: 優化顯示格式，若第二個參數為空則不顯示括號
        Dim msgContent As String = $"{newLine.ToString("00")} {strA} {callingMethod}"
        If Not String.IsNullOrEmpty(strB) Then msgContent &= $" ({strB})"

        Dim newItem As New ListViewItem(msgContent)
        newItem.SubItems.Add(timeNow.ToString("HH:mm:ss.ff"))
        newItem.SubItems.Add(If(newLine > 1, timeSpan.TotalMilliseconds.ToString("#,##0.00"), "-")) ' Index 2: Step (物理間隔)
        newItem.SubItems.Add("")                                                                    ' Index 3: Elapsed (邏輯耗時，預設空，由 Timer_Tick 填入)

        ' 2026/3/28 by Gemini: 預先計算快取資訊存入tag備用
        Dim tag As New DebugItemTag()
        tag.textFullRow = (newItem.Text & " " & newItem.SubItems(1).Text & " " & newItem.SubItems(2).Text & " " & newItem.SubItems(3).Text).ToLower()
        tag.isHit = _searchPattern.Length > 0 AndAlso CheckIsHitInternal(tag.textFullRow)   ' 新訊息加入瞬間也要先比對是否已符合搜尋字串 ' 2026/6/13 改成：_searchPattern 空的話直接跳過
        tag.timeStamp = timeNow                                                             ' 2026/3/28 by Gemini: 保留原始精度供日後計算
        newItem.Tag = tag

        ' by Gemini, 2026/04/03: 區隔邏輯與 Enqueue
        _msgQueue.Enqueue(newItem)                      ' 2026-03-25 by Gemini: 改用 ConcurrentQueue 與 Timer 定期批次新增，大幅提升迴圈寫入效能

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
                Dim matchVB = Regex.Match(typeName, "^VB\$StateMachine_\d+_(.*)$")
                If matchVB.Success Then
                    Dim originalName As String = matchVB.Groups(1).Value
                    Dim parentType = m.DeclaringType.DeclaringType ' 狀態機器類別的上一層通常就是原始類別
                    Return If(parentType IsNot Nothing, $"{parentType.Name}.{originalName} [Async]", $"{originalName} [Async]")
                End If

                ' 2. C# 格式: <MethodName>d__XX (兼顧未來可能混合 C# 專案的情況)
                Dim matchCS = Regex.Match(typeName, "^<(.*)>d__.*$")
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

        Dim itemsToAdd As New List(Of ListViewItem)(256)
        Dim item As ListViewItem = Nothing
        While _msgQueue.TryDequeue(item) : itemsToAdd.Add(item) : End While

        If itemsToAdd.Count > 0 Then
            ' 2026/03/31 by Gemini: 自動為「結束」行預算總耗時
            ' 2026/04/11 by Gemini: 改填入 SubItems(3) (Elapsed 欄位)，並支援跨集合搜尋 (itemsToAdd)
            ' 2206/06/13 by Claude: 效能優化 - 現在要新增的項目如果包含「結束」，才進行配對搜尋，避免每次 Timer 都無差別地 O(N²) 搜尋整個 itemsToAdd 集合
            If itemsToAdd.Any(Function(lvi) lvi.Text.Contains("結束")) Then
                For Each lvi In itemsToAdd
                    If lvi.Text.Contains("結束") Then
                        Dim pair As ListViewItem = FindSimilarPair(lvi, itemsToAdd)
                        If pair IsNot Nothing Then
                            Dim tagCurrent = TryCast(lvi.Tag, DebugItemTag)
                            Dim tagPair = TryCast(pair.Tag, DebugItemTag)

                            If tagCurrent IsNot Nothing AndAlso tagPair IsNot Nothing Then
                                Dim totalMs As Double = Math.Abs((tagCurrent.timeStamp - tagPair.timeStamp).TotalMilliseconds)
                                lvi.SubItems(3).Text = totalMs.ToString("#,##0.00")
                                ' by Gemini, 2026/04/11: 填入數值後同步更新搜尋快取
                                tagCurrent.textFullRow = (lvi.Text & " " & lvi.SubItems(1).Text & " " & lvi.SubItems(2).Text & " " & lvi.SubItems(3).Text).ToLower()
                            End If
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

                        ' 2026/04/09 by Gemini: 修正游標前進但舊選取殘留的問題
                        ' 由於已開啟 MultiSelect=True，直接設 Selected=True 會變成加選
                        ' 因此需先手動清除前次的選取，再設定最後一項，並賦予 Focused 確保游標真正前進
                        .SelectedItems.Clear()
                        lastItem.Selected = True
                        lastItem.Focused = True
                    End If
                End With
            End If

    End Sub
#End Region

#Region "■ 05 ListView 操作事件"
    Private Sub lvwDebug_ItemSelectionChanged(sender As Object, e As ListViewItemSelectionChangedEventArgs) Handles lvwDebug.ItemSelectionChanged

        ' by Gemini, 2026/04/01: 解決點選配對項目時的閃爍問題 (Flickering)
        ' 不直接修改 ListViewItem.BackColor 屬性，改為記錄目標並用 Invalidate() 局部重繪。
        If Not e.IsSelected Then Return

        ' by Gemini, 2026/03/29: 選取變更時自動標記配對的「開始/結束」行
        ' 效能優化: 使用 _lastHighlightedPair 做 O(1) 顏色還原，取代原本的 For Each 全域清除
        '           原本 Shift 多選 100 筆時會觸發 100 次事件 × N 筆 = O(N²) 重繪，改為 O(1)
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
        ''Dim hDC As IntPtr = e.Graphics.GetHdc()
        ''Dim hRgn As IntPtr = CreateRectRgn(e.Bounds.Left, e.Bounds.Top, e.Bounds.Right, e.Bounds.Bottom)
        ''SelectClipRgn(hDC, hRgn)
        ''DeleteObject(hRgn)
        ''e.Graphics.ReleaseHdc(hDC)
        ''e.Graphics.SetClip(e.Bounds)    ' GDI+ clip 同步設定

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
                    Using highlightBrush As New SolidBrush(Color.Yellow)
                        ' Tier 3, 2026/06/15 by Simon/Claude: 改用 framework 快取的 Brushes.Yellow，零配置
                        e.Graphics.FillRectangle(Brushes.Yellow, New Rectangle(currentX, e.Bounds.Y + 2, textRect.Right - currentX, e.Bounds.Height - 4))
                    End Using
                    TextRenderer.DrawText(e.Graphics, matchPart, e.Item.Font, New Rectangle(currentX, e.Bounds.Y, textRect.Right - currentX, e.Bounds.Height), Color.Black, flags Or TextFormatFlags.EndEllipsis)
                    lastPos = itemText.Length : Exit For
                End If

                Using highlightBrush As New SolidBrush(Color.Yellow)
                    ' Tier 3, 2026/06/15 by Simon/Claude: 同上，改用 Brushes.Yellow
                    e.Graphics.FillRectangle(Brushes.Yellow, New Rectangle(currentX, e.Bounds.Y + 2, szMatch.Width, e.Bounds.Height - 4))
                End Using
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
        lvwDebug.Refresh()

    End Sub
    Private Sub chkSearchLogic_CheckedChanged(sender As Object, e As EventArgs) Handles checkAndOr.CheckedChanged
        checkAndOr.Text = If(checkAndOr.Checked, "AND", "OR")
        UpdateSearchCaption() ' by Gemini, 2026/03/31: 切換模式時同步更新視窗標題
        RefreshSearchCache()  ' 2026/3/28 by Gemini: 批次更新快取
        lvwDebug.Refresh()

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
        ' debug: 大量刪除時崩潰
        lvwDebug.BeginUpdate()
        For Each item As ListViewItem In lvwDebug.SelectedItems
            lvwDebug.Items.Remove(item)
        Next
        lvwDebug.EndUpdate()

    End Sub
    Private Sub RefreshSearchCache()
        ''' 2026/3/28 by Gemini: 根據目前關鍵字更新所有項目的 isHit 狀態，並預先產生 Regex 高亮模式
        ' by Gemini, 2026/04/03: 增加邏輯區隔空白
        If lvwDebug.Items.Count = 0 Then Return

        ' 預先產生 Regex 模式，徹底移除 DrawSubItem 中的 LINQ 與字串運算
        Dim keywords = ParseSearchKeywords(txtDebug.Text.Trim())
        _searchPattern = If(keywords.Count > 0,
            String.Join("|", keywords.OrderByDescending(Function(k) k.Length).Select(Function(kw) System.Text.RegularExpressions.Regex.Escape(kw))), "")

        ' Tier 3, 2026/06/15 by Simon/Claude: 僅在搜尋字串變動時重建 Regex 實例 (含 IgnoreCase)，供 DrawSubItem 重複使用，避免每格用字串多載重新解析 pattern
        _searchRegex = If(String.IsNullOrEmpty(_searchPattern), Nothing, New Regex(_searchPattern, RegexOptions.IgnoreCase))

        lvwDebug.BeginUpdate()
        For Each lvi As ListViewItem In lvwDebug.Items
            Dim tag = TryCast(lvi.Tag, DebugItemTag)
            If tag IsNot Nothing Then tag.isHit = CheckIsHitInternal(tag.textFullRow, keywords)
        Next
        lvwDebug.EndUpdate()

    End Sub

    Private Function ParseSearchKeywords(searchText As String) As List(Of String)
        ' 使用 Regex 拆分搜尋關鍵字，支援雙引號括起來的片語
        ' Regex 模式: (?:""(?<q>[^""]*)""|(?<w>\S+))
        Dim keywords As New List(Of String)(16)
        If String.IsNullOrWhiteSpace(searchText) Then Return keywords
        Dim pattern As String = "(?:""(?<q>[^""]*)""|(?<w>\S+))"
        Dim matches = System.Text.RegularExpressions.Regex.Matches(searchText, pattern)
        For Each m As System.Text.RegularExpressions.Match In matches
            If m.Groups("q").Success Then
                keywords.Add(m.Groups("q").Value)
            ElseIf m.Groups("w").Success Then
                keywords.Add(m.Groups("w").Value)
            End If
        Next
        Return keywords

    End Function
    Private Function CheckIsHitInternal(fullText As String, Optional preParsedKeywords As List(Of String) = Nothing) As Boolean
        ''' 2026/3/28 by Gemini: 內部判斷邏輯，可傳入預解析關鍵字以加速批次處理
        Dim keywords = If(preParsedKeywords, ParseSearchKeywords(txtDebug.Text.Trim()))
        If keywords.Count = 0 Then Return False
        Return If(checkAndOr.Checked,
            keywords.All(Function(kw) fullText.Contains(kw.ToLower())), ' AND
            keywords.Any(Function(kw) fullText.Contains(kw.ToLower()))) ' OR

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
        Dim txt As String = selectedItem.Text
        Dim coreName As String = RemoveBeginEnd(txt)
        Dim isBegin As Boolean = txt.Contains("開始")
        Dim isEnd As Boolean = txt.Contains("結束")
        If Not isBegin AndAlso Not isEnd Then Return Nothing
        Dim depth As Integer = 0
        If isBegin Then
            ' 向下搜尋配對的「結束」
            For i As Integer = selectedItem.Index + 1 To lvwDebug.Items.Count - 1
                ' debug: 這裡一直存取物件不會有性能問題嗎? 改成 with lvwDebug? OR??
                Dim item As ListViewItem = lvwDebug.Items(i)
                Dim itemCore As String = RemoveBeginEnd(item.Text)
                If IsContentSimilar(coreName, itemCore) Then
                    If item.Text.Contains("開始") Then
                        depth += 1                      ' 同名的巢狀開始，深度 +1
                    ElseIf item.Text.Contains("結束") Then
                        If depth = 0 Then Return item   ' 深度歸零 = 正確配對
                        depth -= 1                      ' 消耗一層巢狀
                    End If
                End If
            Next
        ElseIf isEnd Then
            ' 向上搜尋配對的「開始」

            ' 💡 Level 1: 先在「同一批批次(待處理)」清單中往回搜尋 (優先級最高，因為距離最近)
            If additionalItems IsNot Nothing Then
                ' 從 selectedItem 在清單中的位置往前找
                Dim selfIdx As Integer = additionalItems.IndexOf(selectedItem)
                If selfIdx > 0 Then
                    For i As Integer = selfIdx - 1 To 0 Step -1
                        Dim item As ListViewItem = additionalItems(i)
                        Dim itemCore As String = RemoveBeginEnd(item.Text)
                        If IsContentSimilar(coreName, itemCore) Then
                            If item.Text.Contains("結束") Then
                                depth += 1
                            ElseIf item.Text.Contains("開始") Then
                                If depth = 0 Then Return item
                                depth -= 1
                            End If
                        End If
                    Next
                End If
            End If

            ' 💡 Level 2: 若在當前批次沒找到，再往「已顯示(ListView)」中搜尋
            ' 2026/04/11 by Gemini: 修正搜尋起點
            ' 當 selectedItem 尚未加入 ListView 時 (Timer_Tick 批次處理中)，Index 會是 -1。
            ' 此時應從 ListView 的最末端 (Items.Count - 1) 開始往回找。
            Dim startIdx As Integer = If(selectedItem.Index >= 0, selectedItem.Index - 1, lvwDebug.Items.Count - 1)
            For i As Integer = startIdx To 0 Step -1
                Dim item As ListViewItem = lvwDebug.Items(i)
                Dim itemCore As String = RemoveBeginEnd(item.Text)
                If IsContentSimilar(coreName, itemCore) Then
                    If item.Text.Contains("結束") Then
                        depth += 1                      ' 同名的巢狀結束，深度 +1
                    ElseIf item.Text.Contains("開始") Then
                        If depth = 0 Then Return item   ' 深度歸零 = 正確配對
                        depth -= 1                      ' 消耗一層巢狀
                    End If
                End If
            Next
        End If

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
