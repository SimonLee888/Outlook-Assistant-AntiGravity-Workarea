Imports System.Diagnostics
Imports System.Reflection

Public Class DebugForm

    ' ==============================================================
    ' DebugForm.vb  —  執行期除錯視窗
    ' ==============================================================
    ' 功能:
    '   即時顯示 Form1.Dbg() 呼叫的訊息（訊息文字、呼叫者、時間戳記、間隔毫秒）
    '   雙擊單行 → 複製該行完整文字到剪貼簿，並重算與配對 Begin/End 的時間差
    '   Ctrl+C   → 複製所有已選取的行（Tab 分隔，每行一列）
    '   點選 End: 行 → 向前搜尋相符的 Begin: 行並以黃色標示
    '
    ' 設計說明:
    '   AddMessage3 由 Form1.Dbg() 呼叫，forcedCaller 由 Form1.WhoCallsMe() 填入
    '   WhoCallsMe() 為 fallback，正常情況下不會被走到（Form1 已先解析好呼叫者）
    '   ListView 啟用 MultiSelect=True（於 Load 覆寫 Designer 設定），支援 Ctrl+C 多選複製
    '
    ' 改動記錄:
    '   2026/3/6  - AddMessage2：合併寫入與更新邏輯，BeginUpdate 批次重繪（by Claude.ai）
    '   2026/3/22 - AddMessage3：支援 forcedCaller 參數；WhoCallsMe 支援 skipLevels（by Grok.ai）
    '   2026/3/23 - 結構整理：加入 Region、改名 DebugForm_Load、移除無用 Imports
    '  (by Claude)  補實作 Ctrl+C 多選複製；移除空白 SelectedIndexChanged；統一 WhoCallsMe 風格
    '
    ' todo:
    '      2. 如何點一個begin/end 就自動highlight 配對的另一端?
    ' ==============================================================

#Region "■ 01 Win32 API & 常數"

    <Runtime.InteropServices.DllImport("user32.dll")>
    Private Shared Function SendMessage(hWnd As IntPtr, msg As Integer, wParam As IntPtr, lParam As IntPtr) As IntPtr
    End Function

    Private Const LVM_SETEXTENDEDLISTVIEWSTYLE As Integer = &H1036
    Private Const LVS_EX_DOUBLEBUFFER As Integer = &H10000
    Private Const WM_SETREDRAW As Integer = &HB  ' 2026/3/26 by AntiGravity

#End Region

#Region "■ 02 成員變數"

    Private _previousTimestamp As Date
    Private _msgQueue As New System.Collections.Concurrent.ConcurrentQueue(Of ListViewItem)
    Private WithEvents _uiTimer As New System.Windows.Forms.Timer() With {.Interval = 100}

#End Region

#Region "■ 03 表單生命週期"
    Private Sub DebugForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' D1 2026-03-23：改名 Form2_Load → DebugForm_Load
        ' 2026/3/26 by AntiGravity: 啟用 DoubleBuffered 減少視窗閃爍
        Me.DoubleBuffered = True
        _previousTimestamp = Now
        _uiTimer.Start()

        ' ✅ 啟用 ListView 雙緩衝，減少閃爍
        SendMessage(lvwDebug.Handle, LVM_SETEXTENDEDLISTVIEWSTYLE, New IntPtr(LVS_EX_DOUBLEBUFFER), New IntPtr(LVS_EX_DOUBLEBUFFER))

        ' ✅ 覆寫 Designer 的 MultiSelect=False，支援 Ctrl+C 多選複製
        ' D5 2026-03-23：啟用多選才能讓 Ctrl+C KeyDown 複製多行
        lvwDebug.MultiSelect = True

        ' 設置 ListView 欄位
        With lvwDebug.Columns
            .Clear()
            .Add("Debug Message", 624, HorizontalAlignment.Left)
            .Add("Timestamp", 312, HorizontalAlignment.Center)
            .Add("Time Span", 99, HorizontalAlignment.Right)
        End With

        AddHandler lvwDebug.ItemSelectionChanged, AddressOf lvwDebug_ItemSelectionChanged

        ' ✅ 設置 Layout 與高亮
        'txtDebug.Dock = DockStyle.Bottom ' 已在 Designer 中放入 pnlSearch
        lvwDebug.Dock = DockStyle.Fill
        'txtDebug.BringToFront()
        lvwDebug.OwnerDraw = True

        ' ✅ 2026/3/26 by AntiGravity: 第一次顯示時強制執行一次縮放邏輯，確保欄位不會超框被遮蔽
        DebugForm_Resize(Nothing, Nothing)
    End Sub
    Private Sub DebugForm_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form1.CheckDebug.Checked = False
    End Sub
    Private Sub DebugForm_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        ' 讓第一欄 (Debug Message) 填滿剩餘空間，右側時間欄位固定
        ' 2026/3/26 by AntiGravity: 使用 WM_SETREDRAW 防止調整欄位寬度時的閃爍
        If lvwDebug.Columns.Count >= 3 Then
            SendMessage(lvwDebug.Handle, WM_SETREDRAW, New IntPtr(0), IntPtr.Zero)
            Try
                Dim otherWidths As Integer = lvwDebug.Columns(1).Width + lvwDebug.Columns(2).Width
                Dim newWidth As Integer = lvwDebug.ClientSize.Width - otherWidths - 4
                If newWidth > 100 Then lvwDebug.Columns(0).Width = newWidth
            Finally
                SendMessage(lvwDebug.Handle, WM_SETREDRAW, New IntPtr(1), IntPtr.Zero)
                lvwDebug.Invalidate()
            End Try
        End If
        If lvwDebug.Items.Count > 0 Then lvwDebug.Items(lvwDebug.Items.Count - 1).EnsureVisible()
    End Sub
#End Region

#Region "■ 04 訊息寫入"
    Public Sub AddMessage3(Optional strA As String = "", Optional strB As String = "",
                           Optional forcedCaller As String = "")
        ' AddMessage3：主要入口，由 Form1.Dbg() 呼叫
        ' forcedCaller: Form1.WhoCallsMe() 預先解析好的呼叫者字串（避免 stack trace 在 DebugForm 裡走不回去）
        ' 2026/3/22 by Grok.ai
        ' 2026-03-25 by AntiGravity: 改用 ConcurrentQueue 與 Timer 定期批次新增，大幅提升迴圈寫入效能

        Static lineCount As Integer
        Dim currentLine As Integer = System.Threading.Interlocked.Increment(lineCount)
        Dim callingMethod As String =
            If(forcedCaller <> "", forcedCaller, WhoCallsMe(1))

        Dim currentTimestamp As Date = Now
        Dim timeSpan As TimeSpan = currentTimestamp - _previousTimestamp
        _previousTimestamp = currentTimestamp

        Dim newItem As New ListViewItem($"{currentLine.ToString("00")} {strA} {callingMethod} ({strB})")
        newItem.SubItems.Add(currentTimestamp.ToString("HH:mm:ss.ff"))
        newItem.SubItems.Add(timeSpan.TotalMilliseconds.ToString("#,##0.00 "))
        newItem.Tag = currentTimestamp

        _msgQueue.Enqueue(newItem)
    End Sub
    Private Sub _uiTimer_Tick(sender As Object, e As EventArgs) Handles _uiTimer.Tick
        If _msgQueue.IsEmpty Then Return

        Dim itemsToAdd As New List(Of ListViewItem)()
        Dim item As ListViewItem = Nothing
        While _msgQueue.TryDequeue(item)
            itemsToAdd.Add(item)
        End While

        If itemsToAdd.Count > 0 Then
            lvwDebug.BeginUpdate()
            lvwDebug.Items.AddRange(itemsToAdd.ToArray())
            lvwDebug.Items(lvwDebug.Items.Count - 1).EnsureVisible()
            lvwDebug.EndUpdate()
        End If
    End Sub
    Private Function WhoCallsMe(Optional skipLevels As Integer = 1) As String
        ' WhoCallsMe：fallback 用途，當 forcedCaller 為空時從 stack trace 解析呼叫者
        ' 正常執行路徑下 Form1.Dbg() 已傳入 forcedCaller，此函數不會被走到
        ' skipLevels: 要跳過幾層 wrapper（預設 1 = 跳過 AddMessage3 本身）
        ' D7 2026-03-23：統一為 Form1.WhoCallsMe 的邏輯風格
        Dim st As New StackTrace(skipLevels + 1, True)
        For i As Integer = 0 To st.FrameCount - 1
            Dim m = st.GetFrame(i)?.GetMethod()
            If m IsNot Nothing AndAlso
                m.DeclaringType IsNot Nothing AndAlso
                m.DeclaringType IsNot GetType(DebugForm) AndAlso
                Not m.Name.Contains("Dbg") AndAlso
                m.Name <> "MoveNext" Then
                Return $"{m.DeclaringType.Name}.{m.Name}"
            End If
        Next
        Return "Unknown Method"
    End Function
#End Region

#Region "■ 05 ListView 操作事件"
    Private Sub txtDebug_TextChanged(sender As Object, e As EventArgs) Handles txtDebug.TextChanged
        ' 搜尋字串變更：重繪 ListView
        Dim logic As String = If(chkSearchLogic.Checked, "OR", "AND")
        Me.Text = $"Debug Form - Searching ({logic}): {txtDebug.Text}"
        lvwDebug.Refresh() ' 💡 使用 Refresh() 比 Invalidate() 更能解決 double buffering 沒刷新的問題
    End Sub
    Private Sub chkSearchLogic_CheckedChanged(sender As Object, e As EventArgs) Handles chkSearchLogic.CheckedChanged
        chkSearchLogic.Text = If(chkSearchLogic.Checked, "OR", "AND")
        lvwDebug.Refresh()
    End Sub

    Private Sub lvwDebug_ItemSelectionChanged(sender As Object, e As ListViewItemSelectionChangedEventArgs) Handles lvwDebug.ItemSelectionChanged
        ' 選取變更：向前搜尋配對的 Begin: 行並黃色標示
        If Not e.IsSelected Then Return

        ' 清除之前的黃色標示
        For Each item As ListViewItem In lvwDebug.Items
            item.BackColor = Color.White
        Next

        Dim selectedItem As ListViewItem = e.Item
        Dim currentContent As String = RemoveBeginEnd(selectedItem.Text)

        ' 向前搜尋含 "Begin" 的相符項目，標示黃色
        For i As Integer = selectedItem.Index - 1 To 0 Step -1
            Dim existingItem As ListViewItem = lvwDebug.Items(i)
            Dim existingContent As String = RemoveBeginEnd(existingItem.Text)
            If existingItem.Text.Contains("Begin") AndAlso IsContentSimilar(currentContent, existingContent) Then
                existingItem.BackColor = Color.Yellow
            End If
        Next
    End Sub
    Private Sub lvwDebug_DrawColumnHeader(sender As Object, e As DrawListViewColumnHeaderEventArgs) Handles lvwDebug.DrawColumnHeader
        e.DrawDefault = True
    End Sub
    Private Sub lvwDebug_DrawItem(sender As Object, e As DrawListViewItemEventArgs) Handles lvwDebug.DrawItem
        ' Details view 下不需要特別處理，交給 DrawSubItem 即可
        ' 設定為 False 避免預設繪製覆寫 Column 0 的文字
        e.DrawDefault = False
    End Sub
    Private Sub lvwDebug_DrawSubItem(sender As Object, e As DrawListViewSubItemEventArgs) Handles lvwDebug.DrawSubItem
        Dim searchText As String = txtDebug.Text.Trim()
        If String.IsNullOrEmpty(searchText) Then
            e.DrawDefault = True
            Return
        End If

        ' 1. 解析關鍵字 (支援雙引號)
        Dim keywords As List(Of String) = ParseSearchKeywords(searchText)
        If keywords.Count = 0 Then
            e.DrawDefault = True
            Return
        End If

        ' 2. 判斷整行是否符合 AND/OR 邏輯
        ' 注意：我們需要檢查整行（跨所有子項目）是否滿足條件
        Dim fullRowText As String = ""
        For Each si As ListViewItem.ListViewSubItem In e.Item.SubItems
            fullRowText &= si.Text & " "
        Next

        Dim isMatchPerRow As Boolean
        If chkSearchLogic.Checked Then ' OR 模式
            isMatchPerRow = keywords.Any(Function(kw) fullRowText.Contains(kw, StringComparison.OrdinalIgnoreCase))
        Else ' AND 模式
            isMatchPerRow = keywords.All(Function(kw) fullRowText.Contains(kw, StringComparison.OrdinalIgnoreCase))
        End If

        If Not isMatchPerRow Then
            e.DrawDefault = True
            Return
        End If

        ' 3. 在當前 SubItem 中尋找哪些關鍵字命中了，並進行繪製
        Dim itemText As String = e.SubItem.Text
        e.DrawDefault = False

        Dim backColor As Color = e.Item.BackColor
        Dim foreColor As Color = e.Item.ForeColor

        ' 處理選取狀態
        If e.Item.Selected Then
            backColor = SystemColors.Highlight
            foreColor = SystemColors.HighlightText
        End If

        ' 繪製背景
        Using backBrush As New SolidBrush(backColor)
            e.Graphics.FillRectangle(backBrush, e.Bounds)
        End Using

        ' 取得對齊方式
        Dim align As HorizontalAlignment = lvwDebug.Columns(e.ColumnIndex).TextAlign
        Dim flags As TextFormatFlags = TextFormatFlags.VerticalCenter Or TextFormatFlags.NoPrefix Or TextFormatFlags.NoPadding

        Select Case align
            Case HorizontalAlignment.Center : flags = flags Or TextFormatFlags.HorizontalCenter
            Case HorizontalAlignment.Right : flags = flags Or TextFormatFlags.Right
            Case Else : flags = flags Or TextFormatFlags.Left
        End Select

        ' 建立匹配所有關鍵字的 Regex (由最長的關鍵字排前面，避免短字吃掉長字)
        Dim sortedKeywords = keywords.OrderByDescending(Function(k) k.Length).ToList()
        Dim pattern As String = String.Join("|", sortedKeywords.Select(Function(kw) System.Text.RegularExpressions.Regex.Escape(kw)))
        Dim matches = System.Text.RegularExpressions.Regex.Matches(itemText, pattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase)

        If matches.Count = 0 Then
            ' 雖然此行匹配（可能在其它格命中），但此格沒中關鍵字，直接繪製一般文字 (並尊重對齊)
            TextRenderer.DrawText(e.Graphics, itemText, e.Item.Font, Rectangle.Inflate(e.Bounds, -2, 0), foreColor, flags Or TextFormatFlags.WordEllipsis)
            Return
        End If

        ' 繪製文字與多重高亮
        ' 先計算總文字寬度以決定起始位置 (為了支援 Right/Center 對齊)
        Dim totalSize = TextRenderer.MeasureText(e.Graphics, itemText, e.Item.Font, New Size(Integer.MaxValue, e.Bounds.Height), TextFormatFlags.NoPadding)
        Dim currentX As Integer
        Select Case align
            Case HorizontalAlignment.Right
                currentX = e.Bounds.Right - totalSize.Width - 2
            Case HorizontalAlignment.Center
                currentX = e.Bounds.X + (e.Bounds.Width - totalSize.Width) / 2
            Case Else
                currentX = e.Bounds.X + 2
        End Select

        Dim lastPos As Integer = 0
        For Each m As System.Text.RegularExpressions.Match In matches
            ' 畫命中前的文字
            If m.Index > lastPos Then
                Dim normalPart As String = itemText.Substring(lastPos, m.Index - lastPos)
                Dim szNormal = TextRenderer.MeasureText(e.Graphics, normalPart, e.Item.Font, New Size(Integer.MaxValue, e.Bounds.Height), TextFormatFlags.NoPadding)
                TextRenderer.DrawText(e.Graphics, normalPart, e.Item.Font, New Rectangle(currentX, e.Bounds.Y, szNormal.Width, e.Bounds.Height), foreColor, TextFormatFlags.VerticalCenter Or TextFormatFlags.NoPadding)
                currentX += szNormal.Width
            End If

            ' 畫高亮背景 (黃色)
            Dim matchPart As String = m.Value
            Dim szMatch = TextRenderer.MeasureText(e.Graphics, matchPart, e.Item.Font, New Size(Integer.MaxValue, e.Bounds.Height), TextFormatFlags.NoPadding)
            Using highlightBrush As New SolidBrush(Color.Yellow)
                e.Graphics.FillRectangle(highlightBrush, New Rectangle(currentX, e.Bounds.Y + 2, szMatch.Width, e.Bounds.Height - 4))
            End Using
            ' 畫高亮文字 (黑色)
            TextRenderer.DrawText(e.Graphics, matchPart, e.Item.Font, New Rectangle(currentX, e.Bounds.Y, szMatch.Width, e.Bounds.Height), Color.Black, TextFormatFlags.VerticalCenter Or TextFormatFlags.NoPadding)

            currentX += szMatch.Width
            lastPos = m.Index + m.Length
        Next

        ' 畫剩餘文字
        If lastPos < itemText.Length Then
            Dim remainingPart As String = itemText.Substring(lastPos)
            Dim szRemaining = TextRenderer.MeasureText(e.Graphics, remainingPart, e.Item.Font, New Size(Integer.MaxValue, e.Bounds.Height), TextFormatFlags.NoPadding)
            TextRenderer.DrawText(e.Graphics, remainingPart, e.Item.Font, New Rectangle(currentX, e.Bounds.Y, szRemaining.Width, e.Bounds.Height), foreColor, TextFormatFlags.VerticalCenter Or TextFormatFlags.NoPadding)
        End If
    End Sub
    Private Sub lvwDebug_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles lvwDebug.MouseDoubleClick
        ' 雙擊：複製該行完整文字到剪貼簿，並重算與配對 Begin/End 的時間差
        If e.Button <> MouseButtons.Left OrElse e.Clicks <> 2 Then Return

        Dim selectedItem As ListViewItem = sender.GetItemAt(e.X, e.Y)
        If selectedItem Is Nothing Then Return

        ' ✅ 複製該行完整文字（Tab 分隔三欄）到剪貼簿
        Dim fullText As String = String.Join(vbTab,
            selectedItem.SubItems.Cast(Of ListViewItem.ListViewSubItem)().Select(Function(s) s.Text))
        Clipboard.SetText(fullText)

        ' ✅ 重算時間差：優先找配對的 Begin/End，否則與上一行比
        If selectedItem.Index > 0 Then
            Dim t2 As Date = selectedItem.Tag
            Dim pair As ListViewItem = FindSimilarPair(selectedItem)
            Dim t1 As Date = If(pair IsNot Nothing, pair.Tag, lvwDebug.Items(selectedItem.Index - 1).Tag)
            Dim sp As TimeSpan = t2 - t1
            selectedItem.SubItems(2).Text = sp.TotalMilliseconds.ToString("#,##0.000  ")
        End If
    End Sub
    Private Sub lvwDebug_KeyDown(sender As Object, e As KeyEventArgs) Handles lvwDebug.KeyDown
        ' Ctrl+C：複製所有已選取的行（Tab 分隔欄位，vbNewLine 分隔行）
        ' D5 2026-03-23：補實作多行複製，需 MultiSelect=True（於 DebugForm_Load 設定）
        If e.Control AndAlso e.KeyCode = Keys.C Then
            Dim selected = lvwDebug.SelectedItems.Cast(Of ListViewItem)().ToList()
            If selected.Count = 0 Then Return
            Dim lines = selected.Select(Function(item)
                                            Return String.Join(vbTab,
                                                item.SubItems.Cast(Of ListViewItem.ListViewSubItem)().Select(Function(s) s.Text))
                                        End Function)
            Clipboard.SetText(String.Join(Environment.NewLine, lines))
            e.Handled = True
        End If
    End Sub
#End Region

#Region "■ 06 輔助函數"
    Private Function ParseSearchKeywords(searchText As String) As List(Of String)
        ' 使用 Regex 拆分搜尋關鍵字，支援雙引號括起來的片語
        ' Regex 模式：(?:""(?<q>[^""]*)""|(?<w>\S+))
        Dim keywords As New List(Of String)()
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
    Private Function FindSimilarPair(selectedItem As ListViewItem) As ListViewItem
        ' FindSimilarPair：尋找與指定 End: 行配對的 Begin: 行（雙擊重算時間差用）
        ' D9 2026-03-23：修正 comment（原本誤寫成「Ctrl+C 複製」）
        If Not selectedItem.Text.Contains("End:") Then Return Nothing
        Dim currentContent As String = RemoveBeginEnd(selectedItem.Text)

        For i As Integer = selectedItem.Index - 1 To 0 Step -1
            Dim existingItem As ListViewItem = lvwDebug.Items(i)
            If existingItem.Text.Contains("Begin:") AndAlso
               IsContentSimilar(currentContent, RemoveBeginEnd(existingItem.Text)) Then
                Return existingItem
            End If
        Next
        Return Nothing
    End Function
    Private Function RemoveBeginEnd(content As String) As String
        ' RemoveBeginEnd：去除行號前綴與 Begin:/End: 標記，用於相似度比對
        ' 1. 去除前面三碼行號  2. 去除 Begin/End 標記  3. 去除前後空白
        Return content.Substring(3).Replace("開始：", "").Replace("結束：", "").Trim()
    End Function
    Private Function IsContentSimilar(content1 As String, content2 As String) As Boolean
        ' IsContentSimilar：判斷兩段文字是否相似（完全相符 or 包含關係）
        If content1 = content2 Then Return True
        Return content1.Contains(content2) OrElse content2.Contains(content1)
    End Function
#End Region

#Region "■ 99 舊版備用 (勿刪)"
    Public Sub AddMessage(Optional strA As String = "", Optional strB As String = "")
        ' 添加項目的方法
        Static lineCount As Integer : lineCount += 1
        Dim callingMethod As String = GetActualCallingMethod()

        Dim newItem As New ListViewItem($"{lineCount.ToString("000")} {strA} {callingMethod} ({strB})")
        'newItem.Tag = Now
        lvwDebug.Items.Add(newItem)
        TriggerItemAddedEvent(newItem)
    End Sub
    Public Sub AddMessage2(Optional strA As String = "", Optional strB As String = "")
        ' =======================================================
        ' Claude AI 的優化建議：將所有欄位準備好後再一次性加入 ListView，減少重繪次數，提升性能。
        ' 合併了AddMessage 和 UpdateListViewItem 的邏輯，避免了不必要的事件觸發和跨線程調用，簡化了代碼結構。
        ' TriggerItemAddedEvent 和 OnItemAddedAsync 也可以整個刪掉，邏輯已合併進來。
        ' 2026/3/6, by Claude.ai
        ' =======================================================

        Static lineCount As Integer : lineCount += 1
        Dim callingMethod As String = GetActualCallingMethod()

        'Dim currentTimestamp As Date = Now
        'Dim timeSpan As TimeSpan = currentTimestamp - previousTimestamp
        'previousTimestamp = currentTimestamp

        '' ✅ 先把所有欄位準備好，再一次性加入並重繪
        'Dim newItem As New ListViewItem($"{lineCount.ToString("000")} {strA} {callingMethod} ({strB})")
        'newItem.SubItems.Add(currentTimestamp.ToString("HH:mm:ss.fff"))
        'newItem.SubItems.Add(timeSpan.TotalMilliseconds.ToString("#,##0.000  "))
        'newItem.Tag = currentTimestamp

        '' ✅ BeginUpdate 包住整個 Add，只重繪一次
        'lvwDebug.BeginUpdate()
        'lvwDebug.Items.Add(newItem)
        'newItem.EnsureVisible()
        'lvwDebug.EndUpdate()

    End Sub
    Private Function GetActualCallingMethod() As String
        ' 獲取實際的調用方法名稱
        Dim stackTrace As New StackTrace()
        For i As Integer = 1 To stackTrace.FrameCount - 1
            Dim frame As StackFrame = stackTrace.GetFrame(i)
            Dim method As MethodBase = frame.GetMethod()
            If method.DeclaringType IsNot GetType(DebugForm) AndAlso method.Name <> "MoveNext" Then Return $"{method.DeclaringType.Name}.{method.Name}"
        Next
        Return "Unknown Method"
    End Function
    Private Sub TriggerItemAddedEvent(newItem As ListViewItem)
        ' 觸發項目新增事件
        Task.Run(Async Function()
                     Await OnItemAddedAsync(newItem)
                 End Function)
    End Sub
    Private Async Function OnItemAddedAsync(newItem As ListViewItem) As Task
        ' todo: 處理項目新增的非同步方法
        Await Task.Yield()

        'Dim currentTimestamp As Date = Now
        'Dim timeSpan As TimeSpan = currentTimestamp - previousTimestamp

        'If Me.InvokeRequired Then
        '    Me.Invoke(New Action(Sub() UpdateListViewItem(newItem, currentTimestamp, timeSpan)))
        'Else
        '    UpdateListViewItem(newItem, currentTimestamp, timeSpan)
        'End If

        'previousTimestamp = currentTimestamp
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
            '.SubItems(0).Tag = currentTimestamp
            '.SubItems(1).Tag = timeSpan
        End With
        lvwDebug.EndUpdate()
    End Sub
#End Region

End Class
