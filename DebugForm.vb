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
    ' ==============================================================

#Region "■ 01 Win32 API & 常數"

    <Runtime.InteropServices.DllImport("user32.dll")>
    Private Shared Function SendMessage(hWnd As IntPtr, msg As Integer, wParam As IntPtr, lParam As IntPtr) As IntPtr
    End Function

    Private Const LVM_SETEXTENDEDLISTVIEWSTYLE As Integer = &H1036
    Private Const LVS_EX_DOUBLEBUFFER As Integer = &H10000

#End Region

#Region "■ 02 成員變數"

    Private _previousTimestamp As Date

#End Region

#Region "■ 03 表單生命週期"
    Private Sub DebugForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' D1 2026-03-23：改名 Form2_Load → DebugForm_Load
        _previousTimestamp = Now

        ' ✅ 啟用 ListView 雙緩衝，減少閃爍
        SendMessage(lvwDebug.Handle, LVM_SETEXTENDEDLISTVIEWSTYLE, New IntPtr(LVS_EX_DOUBLEBUFFER), New IntPtr(LVS_EX_DOUBLEBUFFER))

        ' ✅ 覆寫 Designer 的 MultiSelect=False，支援 Ctrl+C 多選複製
        ' D5 2026-03-23：啟用多選才能讓 Ctrl+C KeyDown 複製多行
        lvwDebug.MultiSelect = True

        ' 設置 ListView 欄位
        With lvwDebug.Columns
            .Clear()
            .Add("Debug Message", 640, HorizontalAlignment.Left)
            .Add("Timestamp", 115, HorizontalAlignment.Center)
            .Add("Time Span", 85, HorizontalAlignment.Right)
        End With

        ' 事件註冊
        AddHandler lvwDebug.ItemSelectionChanged, AddressOf lvwDebug_ItemSelectionChanged
    End Sub
    Private Sub DebugForm_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form1.CheckDebug.Checked = False
    End Sub
    Private Sub DebugForm_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        If lvwDebug.Items.Count > 1 Then lvwDebug.Items(lvwDebug.Items.Count - 1).EnsureVisible()
    End Sub
#End Region

#Region "■ 04 訊息寫入"
    Public Sub AddMessage3(Optional strA As String = "", Optional strB As String = "",
                           Optional forcedCaller As String = "")
        ' AddMessage3：主要入口，由 Form1.Dbg() 呼叫
        ' forcedCaller: Form1.WhoCallsMe() 預先解析好的呼叫者字串（避免 stack trace 在 DebugForm 裡走不回去）
        ' 2026/3/22 by Grok.ai

        Static lineCount As Integer : lineCount += 1
        Dim callingMethod As String =
            If(forcedCaller <> "", forcedCaller, WhoCallsMe(1))

        Dim currentTimestamp As Date = Now
        Dim timeSpan As TimeSpan = currentTimestamp - _previousTimestamp
        _previousTimestamp = currentTimestamp

        Dim newItem As New ListViewItem($"{lineCount.ToString("000")} {strA} {callingMethod} ({strB})")
        newItem.SubItems.Add(currentTimestamp.ToString("HH:mm:ss.fff"))
        newItem.SubItems.Add(timeSpan.TotalMilliseconds.ToString("#,##0.000 "))
        newItem.Tag = currentTimestamp

        lvwDebug.BeginUpdate()
        lvwDebug.Items.Add(newItem)
        newItem.EnsureVisible()
        lvwDebug.EndUpdate()
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
            Clipboard.SetText(String.Join(vbNewLine, lines))
            e.Handled = True
        End If
    End Sub
#End Region

#Region "■ 06 輔助函數"
    ' FindSimilarPair：尋找與指定 End: 行配對的 Begin: 行（雙擊重算時間差用）
    ' D9 2026-03-23：修正 comment（原本誤寫成「Ctrl+C 複製」）
    Private Function FindSimilarPair(selectedItem As ListViewItem) As ListViewItem
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
        Return content.Substring(3).Replace("Begin:", "").Replace("End:", "").Trim()
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
        ' 處理項目新增的非同步方法
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
            .SubItems.Add(currentTimestamp.ToString("HH:mm:ss.fff"))
            '.SubItems.Add(timeSpan.TotalSeconds.ToString("F6"))            ' 顯示秒, 到小數點後六位
            .SubItems.Add(timeSpan.TotalMilliseconds.ToString("#,##0.000  ")) ' 顯示ms, 到小數點後三位
            .EnsureVisible()
            .Tag = currentTimestamp
            '.SubItems(0).Tag = currentTimestamp
            '.SubItems(1).Tag = timeSpan
        End With
        lvwDebug.EndUpdate()
    End Sub
#End Region

End Class
