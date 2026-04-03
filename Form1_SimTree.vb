Imports System.ComponentModel

Public Class Form1_SimTree

    ' ==============================================================
    ' SimTree.vb  —  支援多選的自訂 TreeView 控制項
    ' ==============================================================
    ' 設計目標:
    '   單選       : 行為與原生 TreeView 完全一致，AfterSelect 在滑鼠放開後才觸發
    '   Ctrl+Click : 切換單一節點的選取狀態（加入或移除），MouseUp 後觸發 AfterSelect
    '   Shift+Click: 從上一個選取節點到目前節點範圍全選，MouseUp 後觸發 AfterSelect
    '   方向鍵     : Up/Down 移動，Left 收攏/上移，Right 展開/下移
    '   Space      : 切換目前節點的展開/收攏狀態
    '   右鍵       : 不改變選取狀態，由 Form1 的 MouseClick 處理右鍵選單
    '   失焦       : 選取節點改成淡灰色（與 Windows 檔案總管一致）
    '   得焦       : 還原成深藍 Highlight 色（只重新上色，不改變選取狀態）
    '   Hover      : 只顯示淡灰色，不改變選取狀態
    '
    ' 核心設計決策:
    '   1. 選取狀態完全由 SimTree 自行管理，OnBeforeSelect 永遠 Cancel=True 阻止原生機制
    '   2. OnAfterSelect 不呼叫基類，所有 AfterSelect 事件統一由 FireAfterSelect 手動觸發
    '   3. MouseDown 只記錄目標節點，MouseUp 才執行選取 + 觸發事件（避免 Ctrl+Click 在按下就觸發統計）
    '   4. Shift+Click 範圍選取使用 NextVisibleNode 遍歷，演算法來自 MSTreeview.vb，支援跨層級
    '   5. Form1 不可直接修改 SelectedNode；請使用 AddSelectedNode / SetSelectedNode / ClearSelectedNodes
    '   6. SelectedNode (Shadows) Get 回傳 _lastClickedNode；Set 等同 SelectSingleNode（不觸發 AfterSelect）
    '   7. OnGotFocus 只重新上色，不改變選取狀態（2026-03-18 修正：避免 Await 競爭條件覆蓋 inbox 選取）
    '   8. Space 鍵必須用 e.SuppressKeyPress=True（否則 KeyPress 會再執行一次展開/收攏）
    '
    ' 整合來源:
    '   MSTreeview.vb (WindowsFormsControlLibrary1) — SelectNode / Shift 範圍選取 / OnKeyDown 完整實作
    '   TreeViewMS.vb (WindowsControlLibrary1)      — paintSelectedNodes / removePaintFromNodes 分離清晰
    '   SimTree.vb    (原版)                        — 失焦/得焦高亮保留、SelectedNode Shadows、方向鍵
    ' ==============================================================

    Inherits TreeView

#Region "■ 01 私有狀態"
    Private _selectedNodes As New List(Of TreeNode)()   ' 目前所有選定的節點清單
    Private _lastClickedNode As TreeNode = Nothing      ' 最後一次被點選的節點，Shift+Click 的起始點
    Private _pendingMouseUpNode As TreeNode = Nothing   ' MouseDown 記錄的目標節點，等到 MouseUp 才真正執行選取
    ' 目前是否套用 Highlight 色（True=得焦深藍，False=失焦淡灰）
    ' S2 2026-03-23：從 OnLostFocus 上方移至此私有狀態區，集中管理
    Private _isHighlightActive As Boolean = True
#End Region

#Region "■ 02 公共屬性"
    Public ReadOnly Property SelectedNodes As List(Of TreeNode)
        ' SelectedNodes：回傳目前所有選定節點的清單（唯讀）
        ' 多選時，Form1 應改用這個屬性取得完整清單
        Get
            Return _selectedNodes
        End Get
    End Property
    <Browsable(False)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Shadows Property SelectedNode As TreeNode
        ' SelectedNode：Shadows 原生屬性
        ' [Fix by AntiGravity, 2026/03/26]
        ' 為什麼顯示紅色？
        '   因為 Shadows 了基類的屬性，WinForms Designer 會試圖序列化它。
        '   但 TreeNode 類型無法直接序列化，導致設計工具報錯（未設定其屬性內容序列化）。
        ' 解決方法：
        '   添加 <Browsable(False)> 隱藏於屬性視窗，
        '   添加 <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)> 告訴設計工具不要序列化此屬性。
        '
        '   Get: 單選時回傳選定節點；多選時回傳 _lastClickedNode（最後被點選的那個）
        '        讓 Form1 的 EnsureVisible / ExpandTreeToDefaultInbox 等程式碼可以繼續正常使用
        '   Set: 清除所有選取，只選定指定節點（等同 SelectSingleNode，不觸發 AfterSelect）
        '        讓 ExpandTreeToDefaultInbox 的 treeview.SelectedNode = node 正常運作
        Get
            Return _lastClickedNode     ' 單選時 = 唯一選取節點；多選時 = 最後被點選的節點
        End Get
        Set(value As TreeNode)
            SelectSingleNode(value)     ' 清除所有選取，只選定這一個（不觸發 AfterSelect）
        End Set
    End Property
#End Region

#Region "■ 03 建構子"
    Public Sub New()
        ' VB.NET 繼承時自動生成預設建構子，此處明確宣告以方便未來擴充（例如加入初始化邏輯）
        MyBase.New()

    End Sub
#End Region

#Region "■ 04 滑鼠事件"
    Protected Overrides Sub OnMouseDown(e As MouseEventArgs)
        ' ✅ 右鍵：完全不改變選取狀態，交給 Form1 的 MouseClick 處理右鍵選單
        If e.Button = MouseButtons.Right Then
            MyBase.OnMouseDown(e) : Return
        End If
        ' 記錄 MouseDown 的目標節點，等到 MouseUp 才執行選取 + 觸發 AfterSelect
        ' 原因：避免 Ctrl+Click 每次 MouseDown 就觸發統計，應在使用者放開滑鼠後才統計
        Dim nodeUnderMouse As TreeNode = GetNodeAt(e.X, e.Y)
        _pendingMouseUpNode = nodeUnderMouse    ' 可能是 Nothing（點到空白處）
        ' ✅ 呼叫基類讓 TreeView 處理展開/收攏圖示的點擊（+/-圖示）
        MyBase.OnMouseDown(e)

    End Sub
    Protected Overrides Sub OnMouseUp(e As MouseEventArgs)
        MyBase.OnMouseUp(e)
        ' 右鍵或點到空白處不處理
        If e.Button <> MouseButtons.Left Then Return
        If _pendingMouseUpNode Is Nothing Then Return
        ' 確認 MouseUp 的位置跟 MouseDown 的節點一致（避免拖曳後的誤觸）
        Dim nodeUnderMouse As TreeNode = GetNodeAt(e.X, e.Y)
        If nodeUnderMouse IsNot _pendingMouseUpNode Then
            _pendingMouseUpNode = Nothing : Return
        End If
        ' ✅ MouseUp 才真正執行選取邏輯 + 觸發 AfterSelect
        SelectNode(_pendingMouseUpNode)
        _pendingMouseUpNode = Nothing

    End Sub
#End Region

#Region "■ 05 原生行為攔截（BeforeSelect / AfterSelect）"
    Protected Overrides Sub OnBeforeSelect(e As TreeViewCancelEventArgs)
        ' ✅ 永遠取消原生選取，防止基類自行改變 BackColor / ForeColor
        ' 選取邏輯完全由 SelectNode 負責，不依賴原生機制
        MyBase.OnBeforeSelect(e)
        e.Cancel = True

    End Sub
    Protected Overrides Sub OnAfterSelect(e As TreeViewEventArgs)
        ' ✅ 不呼叫 MyBase.OnAfterSelect，避免基類的 AfterSelect 和 FireAfterSelect 重複觸發
        '    Form1 的 Handles SimTree2.AfterSelect 由 FireAfterSelect 主動觸發

    End Sub
#End Region

#Region "■ 06 鍵盤事件"
    Protected Overrides Sub OnKeyDown(e As KeyEventArgs)
        ' 處理方向鍵導覽、Space 展開收攏、PageUp/Down、Home/End
        MyBase.OnKeyDown(e)
        If e.KeyCode = Keys.ShiftKey Then Return    ' 單獨按 Shift 不做任何事
        ' 沒有任何選取時，先選第一個可見節點（使用者 Tab 到 SimTree 後第一次按鍵時觸發）
        If _lastClickedNode Is Nothing AndAlso TopNode IsNot Nothing Then
            SelectSingleNode(TopNode) : Return
        End If
        If _lastClickedNode Is Nothing Then Return
        Dim bShift As Boolean = (ModifierKeys = Keys.Shift)
        Select Case e.KeyCode
            Case Keys.Up
                ' 向上移動（Shift+Up 擴展範圍選取）
                Dim prev As TreeNode = _lastClickedNode.PrevVisibleNode
                If prev IsNot Nothing Then
                    If bShift Then AppendToSelection(prev) Else SelectSingleNode(prev)
                    prev.EnsureVisible() : FireAfterSelect(prev)
                End If
                e.Handled = True
            Case Keys.Down
                ' 向下移動（Shift+Down 擴展範圍選取）
                Dim nxt As TreeNode = _lastClickedNode.NextVisibleNode
                If nxt IsNot Nothing Then
                    If bShift Then AppendToSelection(nxt) Else SelectSingleNode(nxt)
                    nxt.EnsureVisible() : FireAfterSelect(nxt)
                End If
                e.Handled = True
            Case Keys.Left
                ' 已展開 → 收攏；已收攏 → 移到父節點（單選，不擴展範圍）
                If _lastClickedNode.IsExpanded Then
                    _lastClickedNode.Collapse()
                ElseIf _lastClickedNode.Parent IsNot Nothing Then
                    SelectSingleNode(_lastClickedNode.Parent)
                    _lastClickedNode.EnsureVisible() : FireAfterSelect(_lastClickedNode)
                End If
                e.Handled = True
            Case Keys.Right
                ' 已收攏 → 展開；已展開 → 移到第一個子節點（單選，不擴展範圍）
                If Not _lastClickedNode.IsExpanded Then
                    _lastClickedNode.Expand()
                ElseIf _lastClickedNode.Nodes.Count > 0 Then
                    SelectSingleNode(_lastClickedNode.Nodes(0))
                    _lastClickedNode.EnsureVisible() : FireAfterSelect(_lastClickedNode)
                End If
                e.Handled = True
            Case Keys.Space
                ' 切換目前節點的展開/收攏狀態
                ' ✅ 必須用 SuppressKeyPress=True（單用 Handled=True 無法阻止 KeyPress，會執行兩次）
                If _lastClickedNode.IsExpanded Then _lastClickedNode.Collapse() Else _lastClickedNode.Expand()
                e.SuppressKeyPress = True
            Case Keys.Home
                ' 移到最頂端第一個根節點（Shift+Home：從目前到頂端全選）
                If Nodes.Count > 0 Then
                    If bShift Then SelectRange(_lastClickedNode, Nodes(0)) Else SelectSingleNode(Nodes(0))
                    Nodes(0).EnsureVisible() : FireAfterSelect(_lastClickedNode)
                End If
                e.Handled = True
            Case Keys.End
                ' 移到最後一個可見節點（Shift+End：從目前到底端全選）
                Dim lastVisible As TreeNode = GetLastVisibleNode()
                If lastVisible IsNot Nothing Then
                    If bShift Then SelectRange(_lastClickedNode, lastVisible) Else SelectSingleNode(lastVisible)
                    lastVisible.EnsureVisible() : FireAfterSelect(_lastClickedNode)
                End If
                e.Handled = True
            Case Keys.PageUp
                ' 向上捲動一頁
                Dim nCount As Integer = VisibleCount
                Dim ndCurrent As TreeNode = _lastClickedNode
                While nCount > 0 AndAlso ndCurrent.PrevVisibleNode IsNot Nothing
                    ndCurrent = ndCurrent.PrevVisibleNode : nCount -= 1
                End While
                SelectSingleNode(ndCurrent) : ndCurrent.EnsureVisible() : FireAfterSelect(ndCurrent)
                e.Handled = True
            Case Keys.PageDown
                ' 向下捲動一頁
                Dim nCount As Integer = VisibleCount
                Dim ndCurrent As TreeNode = _lastClickedNode
                While nCount > 0 AndAlso ndCurrent.NextVisibleNode IsNot Nothing
                    ndCurrent = ndCurrent.NextVisibleNode : nCount -= 1
                End While
                SelectSingleNode(ndCurrent) : ndCurrent.EnsureVisible() : FireAfterSelect(ndCurrent)
                e.Handled = True
        End Select

    End Sub
#End Region

#Region "■ 07 失焦 / 得焦"
    Protected Overrides Sub OnLostFocus(e As EventArgs)
        ' 失焦：改成淡灰色，與原生 TreeView 的失焦行為一致
        MyBase.OnLostFocus(e)
        _isHighlightActive = False
        For Each node As TreeNode In _selectedNodes
            node.BackColor = Form1.ThemeColors.MercuryGray
            node.ForeColor = SystemColors.InactiveCaptionText
        Next

    End Sub
    Protected Overrides Sub OnGotFocus(e As EventArgs)
        ' 得焦：還原 Highlight 深藍色
        ' ✅ 2026-03-18 修正：OnGotFocus 只做重新上色，不改變選取狀態（不自動選 TopNode）
        '    原本加入「_selectedNodes 為空時自動選 TopNode」是為了 Tab 鍵初始選取，
        ' 但造成副作用：
        '    ExpandTreeToDefaultInbox 設好 inbox 後，多個 Await Task.Yield 之間若 _selectedNodes 被清空，
        '    BeginInvoke(SimTree2.Focus()) 觸發 OnGotFocus 時會覆蓋 inbox 選取。
        ' 改為：
        '    「第一次按鍵時若無選取則選 TopNode」，由 OnKeyDown 開頭處理，效果相同但不受 Await 競爭條件影響。
        MyBase.OnGotFocus(e)
        _isHighlightActive = True
        PaintSelectedNodes()    ' 重新上 Highlight 色，不改選取狀態

    End Sub
#End Region

#Region "■ 08 核心選取邏輯"
    ' SelectNode：根據目前的 ModifierKeys 決定選取方式
    '   無修飾鍵 → 單選（清除其他選取，只選這個節點）
    '   Ctrl     → 切換這個節點的選取狀態（加入或移除）
    '   Shift    → 從 _lastClickedNode 到目前節點範圍全選
    Private Sub SelectNode(node As TreeNode)
        If node Is Nothing Then Return
        BeginUpdate()   ' ✅ 批次更新，避免逐一上色造成閃爍
        Try
            Select Case ModifierKeys
                Case Keys.Control
                    ' Ctrl+Click：切換單一節點的選取狀態
                    ToggleSingleNode(node, Not _selectedNodes.Contains(node))
                    _lastClickedNode = node
                Case Keys.Shift
                    ' Shift+Click：從 _lastClickedNode 到 node 範圍全選
                    ' _lastClickedNode 不更新（下一次 Shift+Click 仍從同一起點開始）
                    If _lastClickedNode IsNot Nothing Then
                        SelectRange(_lastClickedNode, node)
                    Else
                        SelectSingleNode(node)  ' 還沒有起點時，退化為單選
                    End If
                Case Else
                    ' 一般點擊：單選
                    SelectSingleNode(node)
            End Select
        Finally
            EndUpdate()
        End Try
        ' ✅ 選取完成後才觸發 AfterSelect，讓 Form1 知道選取已改變，可以開始統計
        FireAfterSelect(node)

    End Sub
#End Region

#Region "■ 09 輔助選取方法（Private）"
    Private Sub SelectSingleNode(node As TreeNode)
        ' SelectSingleNode：清除所有選取，只選定一個節點（不觸發 AfterSelect）
        ' 供鍵盤導覽和初始化使用，由呼叫端決定是否呼叫 FireAfterSelect
        If node Is Nothing Then Return
        ClearSelectedNodes()
        ToggleSingleNode(node, True)
        _lastClickedNode = node
        node.EnsureVisible()

    End Sub
    Private Sub AppendToSelection(node As TreeNode)
        ' AppendToSelection：在不清除現有選取的情況下新增一個節點（Shift+方向鍵用）
        If node Is Nothing Then Return
        ToggleSingleNode(node, True)
        _lastClickedNode = node

    End Sub
    Private Sub SelectRange(startNode As TreeNode, endNode As TreeNode)
        ' SelectRange：選取從 startNode 到 endNode 之間所有可見節點
        ' 使用 NextVisibleNode 遍歷，支援跨層級的選取
        ' 演算法來自 MSTreeview.vb (WindowsFormsControlLibrary1) 的 SelectNode Shift 段落
        If startNode Is Nothing OrElse endNode Is Nothing Then Return
        ' 判斷兩個節點在可見順序中的先後:
        '   方法：把兩個節點提升到同一層的共同祖先，比較祖先的 Index 決定上下順序
        '   這個方法取自 MSTreeview.vb，正確處理跨層級選取的情況
        Dim ndStartP As TreeNode = startNode
        Dim ndEndP As TreeNode = endNode
        Dim startDepth As Integer = Math.Min(ndStartP.Level, ndEndP.Level)
        ' 把比較深的節點提升到共同深度
        While ndStartP.Level > startDepth : ndStartP = ndStartP.Parent : End While
        While ndEndP.Level > startDepth : ndEndP = ndEndP.Parent : End While
        ' 繼續往上找到同一個父節點
        While ndStartP.Parent IsNot ndEndP.Parent
            ndStartP = ndStartP.Parent : ndEndP = ndEndP.Parent
        End While
        ' 確保 topNode 在上方（Index 較小或層級較淺）
        Dim topNode As TreeNode = startNode
        Dim bottomNode As TreeNode = endNode
        If ndStartP.Index > ndEndP.Index OrElse
           (ndStartP.Index = ndEndP.Index AndAlso startNode.Level > endNode.Level) Then
            topNode = endNode : bottomNode = startNode  ' 交換，讓 topNode 永遠在上面
        End If
        ' 清除舊選取，從上往下沿 NextVisibleNode 選取範圍內的所有節點
        ClearSelectedNodes()
        Dim current As TreeNode = topNode
        While current IsNot Nothing
            ToggleSingleNode(current, True)
            If current Is bottomNode Then Exit While
            current = current.NextVisibleNode
        End While
        _lastClickedNode = endNode  ' 範圍選取後，焦點節點是滑鼠點選的那個（非起點）

    End Sub
    Private Sub ToggleSingleNode(node As TreeNode, selectIt As Boolean)
        ' ToggleSingleNode：設定或清除單一節點的選取狀態和高亮顏色
        If node Is Nothing Then Return
        If selectIt Then
            If Not _selectedNodes.Contains(node) Then _selectedNodes.Add(node)
            node.BackColor = SystemColors.Highlight
            node.ForeColor = SystemColors.HighlightText
        Else
            _selectedNodes.Remove(node)
            node.BackColor = Me.BackColor
            node.ForeColor = Me.ForeColor
        End If
        ' ✅ 同步更新基類的 SelectedNode（單選時讓原生 API 也能讀到正確值）
        MyBase.SelectedNode = If(_selectedNodes.Count = 1, _selectedNodes(0), Nothing)

    End Sub
    Private Sub PaintSelectedNodes()
        ' PaintSelectedNodes：重新套用選取高亮顏色（得焦時使用）
        For Each node As TreeNode In _selectedNodes
            node.BackColor = SystemColors.Highlight
            node.ForeColor = SystemColors.HighlightText
        Next

    End Sub
#End Region

#Region "■ 10 公共輔助方法（供 Form1 呼叫）"
    Public Sub ClearSelectedNodes()
        ' ClearSelectedNodes：清除所有選取（顏色還原 + 清空清單）
        ' Form1 在切換模式時呼叫，避免殘留不正確的高亮
        For Each node As TreeNode In _selectedNodes
            node.BackColor = Me.BackColor
            node.ForeColor = Me.ForeColor
        Next
        _selectedNodes.Clear()
        _lastClickedNode = Nothing
        MyBase.SelectedNode = Nothing

    End Sub
    Public Sub FireAfterSelect(node As TreeNode)
        ' FireAfterSelect：手動觸發 AfterSelect 事件，通知 Form1 選取已改變
        ' 傳入最後被操作的節點作為 TreeViewEventArgs.Node
        ' Form1 的「Handles SimTree2.AfterSelect」透過這裡收到事件
        If node Is Nothing Then Return
        MyBase.OnAfterSelect(New TreeViewEventArgs(node))

    End Sub
    Public Sub AddSelectedNode(node As TreeNode)
        ' AddSelectedNode：從外部新增一個選定節點
        ' 不清除其他選取，不觸發 AfterSelect
        ' 例如 ExpandTreeToDefaultInbox 初始化選取後使用
        If node Is Nothing Then Return
        ToggleSingleNode(node, True)
        _lastClickedNode = node

    End Sub
    Public Sub SetSelectedNode(node As TreeNode)
        ' SetSelectedNode：等同 SelectedNode = node
        ' （保留相容舊版呼叫，不觸發 AfterSelect）
        SelectSingleNode(node)

    End Sub
#End Region

#Region "■ 11 私有輔助"
    Private Function GetLastVisibleNode() As TreeNode
        ' GetLastVisibleNode：取得樹狀結構中最後一個可見節點（End 鍵使用）
        If Nodes.Count = 0 Then Return Nothing
        Dim last As TreeNode = Nodes(Nodes.Count - 1)
        While last.IsExpanded AndAlso last.LastNode IsNot Nothing
            last = last.LastNode
        End While
        Return last

    End Function
#End Region

End Class
