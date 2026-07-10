Imports System.ComponentModel
Imports System.Runtime.InteropServices

Public Class SimTree

    Inherits TreeView

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
    '   MSTreeview.vb (WindowsFormsControlLibrary1) — SelectNodeInternal / Shift 範圍選取 / OnKeyDown 完整實作
    '   TreeViewMS.vb (WindowsControlLibrary1)      — paintSelectedNodes / removePaintFromNodes 分離清晰
    '   SimTree.vb    (原版)                        — 失焦/得焦高亮保留、SelectedNode Shadows、方向鍵
    ' ==============================================================

#Region "■ 00 私有狀態"
    Private _selectedNodes As New List(Of TreeNode)(16)     ' 目前所有選定的節點清單
    Private _lastClickedNode As TreeNode = Nothing          ' 最後一次被點選的節點，Shift+Click 的起始點
    Private _LastHoverNodeInternal As TreeNode = Nothing    ' 記錄最後懸停的節點 (原 Form1 的 _lastHoveredNode)
    Private _pendingMouseUpNode As TreeNode = Nothing       ' MouseDown 記錄的目標節點，等到 MouseUp 才真正執行選取
    ' 目前是否套用 Highlight 色（True=得焦深藍，False=失焦淡灰）
    ' S2 2026-03-23：從 OnLostFocus 上方移至此私有狀態區，集中管理
    Private _isHighlightActive As Boolean = True

    ' Win32 常數與 API (用於 SuppressAutoHScroll：保留水平捲軸、抑制自動水平位移)
    ' 2026/07/10 by Simon/Claude: 取代舊 HideHorizontalScrollBar（TVS_NOHSCROLL 只藏捲軸、擋不住內容位移，已移除）
    Private Const WM_HSCROLL As Integer = &H114             ' 水平捲動訊息
    Private Const WM_KEYDOWN As Integer = &H100             ' 鍵盤按下訊息
    Private Const WM_LBUTTONDOWN As Integer = &H201         ' 滑鼠左鍵按下訊息
    Private Const WM_LBUTTONDBLCLK As Integer = &H203       ' 滑鼠左鍵雙擊訊息
    Private Const TVM_ENSUREVISIBLE As Integer = &H1114     ' TreeView EnsureVisible 訊息 (TVM_FIRST + 20)
    Private Const TVM_SELECTITEM As Integer = &H110B        ' TreeView 選取節點訊息 (TVM_FIRST + 11)
    Private Const SB_HORZ As Integer = 0                    ' 水平捲軸標記
    Private Const SB_THUMBPOSITION As Integer = 4           ' WM_HSCROLL 通知碼：捲動到指定位置

    <DllImport("user32.dll")>
    Private Shared Function GetScrollPos(hWnd As IntPtr, nBar As Integer) As Integer
    End Function
    <DllImport("user32.dll")>
    Private Shared Function SendMessage(hWnd As IntPtr, msg As Integer, wParam As IntPtr, lParam As IntPtr) As IntPtr
    End Function
#End Region

#Region "■ 01 建構子"
    Public Sub New()
        ' VB.NET 繼承時自動生成預設建構子，此處明確宣告以方便未來擴充（例如加入初始化邏輯）
        MyBase.New()

        ' 2026/5/14 by simon/Gemini: 將mouse hover作成內建功能
        If EnableHoverHighlight Then Me.DrawMode = TreeViewDrawMode.OwnerDrawText   ' 開啟自訂繪製，由我們自己接管文字與背景的渲染
    End Sub
#End Region

#Region "■ 02 公共屬性"
    Public ReadOnly Property SelectedNodes As List(Of TreeNode)
        ' SelectedNodes：回傳目前所有選定節點的清單（唯讀）
        ' [Fix by Gemini 3.0 Flash, 2026/04/17]
        ' 增加安全性檢查：過濾掉因為 Nodes.Clear() 而脫離樹狀結構的殘留節點
        Get
            _selectedNodes.RemoveAll(Function(n) n.TreeView IsNot Me)
            Return _selectedNodes
        End Get
    End Property
    <Browsable(False)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Shadows Property SelectedNode As TreeNode
        ' SelectedNode：Shadows 原生屬性
        ' [Fix by Gemini, 2026/03/26]
        ' 為什麼顯示紅色？
        '   因為 Shadows 了基類的屬性，WinForms Designer 會試圖序列化它。
        '   但 TreeNode 類型無法直接序列化，導致設計工具報錯（未設定其屬性內容序列化）。
        ' 解決方法：
        '   添加 <Browsable(False)> 隱藏於屬性視窗，
        '   添加 <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)> 告訴設計工具不要序列化此屬性。
        '
        '   Get: 單選時回傳選定節點；多選時回傳 _lastClickedNode（最後被點選的那個）
        '        讓 Form1 的 EnsureVisible / GotoDefaultInbox 等程式碼可以繼續正常使用
        '   Set: 清除所有選取，只選定指定節點（等同 SelectSingleNode，不觸發 AfterSelect）
        '        讓 GotoDefaultInbox 的 treeview.SelectedNode = node 正常運作
        Get
            Return _lastClickedNode     ' 單選時 = 唯一選取節點；多選時 = 最後被點選的節點
        End Get
        Set(value As TreeNode)
            SelectSingleNode(value)     ' 清除所有選取，只選定這一個（不觸發 AfterSelect）
        End Set
    End Property

    ' SuppressAutoHScroll：保留水平捲軸（使用者仍可手動捲動），但抑制選取節點 / EnsureVisible
    ' 造成的自動水平位移（原生 TreeView 遇到長文字節點會自動往右對齊，無屬性可關）
    ' 2026/07/10 by Simon/Claude
    <Category("Behavior"), DefaultValue(False)>
    Public Property SuppressAutoHScroll As Boolean = False

    <Category("Appearance"), Description("是否啟用滑鼠懸停高亮效果"), DefaultValue(True)>
    Public Property EnableHoverHighlight As Boolean = True

    <Category("Appearance"), Description("懸停高亮顏色"), DefaultValue(GetType(Color), "240, 240, 240")>
    Public Property HoverColor As Color = Form1.ThemeColors.MercuryGray ' 淡灰色
#End Region

#Region "■ 03 核心訊息處理（捲軸控制）"
    Protected Overrides Sub WndProc(ByRef m As Message)
        ' WndProc：SuppressAutoHScroll 啟用時，攔截所有可能引發「自動水平位移」的訊息
        ' 2026/07/10 by Simon/Claude
        ' 原生 TreeView 在選取/EnsureVisible 長文字節點時會自動水平捲動對齊，comctl32 寫死、無樣式可關。
        ' 對策：讓訊息正常處理（垂直捲動、選取、展開收攏都保留），事後若發現水平位置被動過，就捲回原位。
        '   TVM_ENSUREVISIBLE : node.EnsureVisible()（鍵盤導航、Form1 GotoDefaultInbox 等）
        '   TVM_SELECTITEM    : MyBase.SelectedNode = ...（SelectSingleNode 等）
        '   WM_LBUTTONDOWN/DBLCLK, WM_KEYDOWN : comctl32 內部 ensure-visible（不經上述訊息）
        If SuppressAutoHScroll AndAlso IsHandleCreated Then
            Select Case m.Msg
                Case TVM_ENSUREVISIBLE, TVM_SELECTITEM, WM_LBUTTONDOWN, WM_LBUTTONDBLCLK, WM_KEYDOWN
                    Dim oldPos As Integer = GetScrollPos(Me.Handle, SB_HORZ)
                    MyBase.WndProc(m)
                    If GetScrollPos(Me.Handle, SB_HORZ) <> oldPos Then
                        ' SB_THUMBPOSITION 走控制項自己的捲動邏輯，內容與捲軸同步還原
                        SendMessage(Me.Handle, WM_HSCROLL, New IntPtr((oldPos << 16) Or SB_THUMBPOSITION), IntPtr.Zero)
                    End If
                    Return
            End Select
        End If
        MyBase.WndProc(m)
    End Sub
#End Region

#Region "■ 04 原生行為攔截（GotFocus / LostFocus / BeforeSelect / AfterSelect）"
    Protected Overrides Sub OnGotFocus(e As EventArgs)
        ' 得焦：還原 Highlight 深藍色
        ' ✅ 2026-03-18 修正：OnGotFocus 只做重新上色，不改變選取狀態（不自動選 TopNode）
        '    原本加入「_selectedNodes 為空時自動選 TopNode」是為了 Tab 鍵初始選取，
        ' 但造成副作用：
        '    GotoDefaultInbox 設好 inbox 後，多個 Await Task.Yield 之間若 _selectedNodes 被清空，
        '    BeginInvoke(SimTree2.Focus()) 觸發 OnGotFocus 時會覆蓋 inbox 選取。
        ' 改為：
        '    「第一次按鍵時若無選取則選 TopNode」，由 OnKeyDown 開頭處理，效果相同但不受 Await 競爭條件影響。
        MyBase.OnGotFocus(e)
        _isHighlightActive = True
        PaintSelectedNodes()    ' 重新上 Highlight 色，不改選取狀態

    End Sub
    Protected Overrides Sub OnLostFocus(e As EventArgs)
        ' 失焦：改成淡灰色，與原生 TreeView 的失焦行為一致
        MyBase.OnLostFocus(e)
        _isHighlightActive = False
        For Each node As TreeNode In _selectedNodes
            node.BackColor = Form1.ThemeColors.AltoGray
            node.ForeColor = SystemColors.InactiveCaptionText
        Next
        ' 2026/5/14 by simon/Gemini: 將mouse hover作成內建功能, 去除上方的color設置, 只用這行me.invalidate
        If EnableHoverHighlight Then Me.Invalidate()
    End Sub
    Protected Overrides Sub OnBeforeSelect(e As TreeViewCancelEventArgs)
        ' ✅ 永遠取消原生選取，防止基類自行改變 BackColor / ForeColor
        ' 選取邏輯完全由 SelectNodeInternal 負責，不依賴原生機制
        MyBase.OnBeforeSelect(e)
        e.Cancel = True

    End Sub
    Protected Overrides Sub OnAfterSelect(e As TreeViewEventArgs)
        ' ✅ 不呼叫 MyBase.OnAfterSelect，避免基類的 AfterSelect 和 FireAfterSelect 重複觸發
        '    Form1 的 Handles SimTree2.AfterSelect 由 FireAfterSelect 主動觸發

    End Sub
#End Region

#Region "■ 05 鍵盤滑鼠事件"
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
        SelectNodeInternal(_pendingMouseUpNode)
        _pendingMouseUpNode = Nothing

    End Sub
    Protected Overrides Sub OnMouseMove(e As MouseEventArgs)
        ' 覆寫 OnMouseMove 處理內建 Hover
        ' 2026/5/14 by simon/Gemini: 將mouse hover作成內建功能
        MyBase.OnMouseMove(e)
        If Not EnableHoverHighlight Then Return

        ' 取得滑鼠下的節點
        Dim hitInfo = Me.HitTest(e.Location)
        Dim currentNode = hitInfo.Node

        ' 若滑鼠移動到新節點
        If currentNode IsNot _LastHoverNodeInternal Then
            Dim oldNode = _LastHoverNodeInternal
            _LastHoverNodeInternal = currentNode

            ' 通知系統重繪舊節點與新節點的區域即可，絕對不要去改 node.BackColor
            If oldNode IsNot Nothing Then Me.Invalidate(oldNode.Bounds)
            If _LastHoverNodeInternal IsNot Nothing Then Me.Invalidate(_LastHoverNodeInternal.Bounds)
        End If
    End Sub
    Protected Overrides Sub OnMouseLeave(e As EventArgs)
        ' 當滑鼠離開時清除 Hover 狀態
        ' 2026/5/14 by simon/Gemini: 將mouse hover作成內建功能
        MyBase.OnMouseLeave(e)
        If _LastHoverNodeInternal IsNot Nothing Then
            Dim oldNode = _LastHoverNodeInternal
            _LastHoverNodeInternal = Nothing
            Me.Invalidate(oldNode.Bounds)
        End If
    End Sub
    Protected Overrides Async Sub OnKeyDown(e As KeyEventArgs)
        ' 處理方向鍵導覽、Space 展開收攏、PageUp/Down、Home/End
        MyBase.OnKeyDown(e)

        ' 2026/05/17 by Simon/Claude:
        ' 若外部事件處理器 (如 Form1 的 F5) 已標記 e.Handled = True，立即退出，
        ' 避免 async/await 讓出控制權後，_lastClickedNode 為 Nothing 觸發SelectSingleNode(TopNode)，汙染 _selectedNodes 造成統計結果錯誤。
        If e.Handled Then Return
        ' 時序問題根因：ClearSelectedNodes() 把 _lastClickedNode 設為 Nothing，ForceTvRefresh 遇到第一個 Await 就把控制權還給 OnKeyDown，
        ' 此時 If _lastClickedNode Is Nothing 判斷成立，誤選了 TopNode。

        If e.KeyCode = Keys.ShiftKey Then Return    ' 單獨按 Shift 不做任何事

        ' 沒有任何選取時，先選第一個可見節點（使用者 Tab 到 SimTree 後第一次按鍵時觸發）
        If _lastClickedNode Is Nothing AndAlso TopNode IsNot Nothing Then
            SelectSingleNode(TopNode) : Return
        End If
        If _lastClickedNode Is Nothing Then Return

        Dim bShift As Boolean = (ModifierKeys = Keys.Shift)
        Select Case e.KeyCode
            Case Keys.A
                If ModifierKeys = Keys.Control AndAlso Nodes.Count > 0 Then
                    ' by Gemini 3.1 Pro, 2026/05/10
                    ' 必須在 Await 之前設定 Handled 和 SuppressKeyPress，
                    ' 因為 Async Sub 遇到 Await 會立即將控制權交還給呼叫端 (WinForms)，
                    ' 若在 Await 之後才設定，WinForms 早已認定此按鍵未被處理而發出「咚」聲。
                    e.Handled = True
                    e.SuppressKeyPress = True

                    Dim lastNode As TreeNode = GetLastVisibleNode()
                    If lastNode IsNot Nothing Then
                        Me.BeginUpdate()
                        SelectRange(Nodes(0), lastNode)
                        Me.EndUpdate()
                        Await Task.Yield
                        FireAfterSelect(_lastClickedNode)
                    End If
                End If
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

#Region "■ 06 輔助選取方法（Private）"
    Private Sub SelectNodeInternal(node As TreeNode)
        ' SelectNodeInternal：根據目前的 ModifierKeys 決定選取方式
        '   無修飾鍵 → 單選（清除其他選取，只選這個節點）
        '   Ctrl     → 切換這個節點的選取狀態（加入或移除）
        '   Shift    → 從 _lastClickedNode 到目前節點範圍全選
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
        ' 演算法來自 MSTreeview.vb (WindowsFormsControlLibrary1) 的 SelectNodeInternal Shift 段落

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
        If ndStartP.Index > ndEndP.Index OrElse (ndStartP.Index = ndEndP.Index AndAlso startNode.Level > endNode.Level) Then
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
        ' [Fix by Gemini 3.0 Flash, 2026/04/17] 歸屬檢查：若節點不屬於目前樹則不動作
        If node Is Nothing OrElse node.TreeView IsNot Me Then Return

        If selectIt Then
            If Not _selectedNodes.Contains(node) Then _selectedNodes.Add(node)
            If Not EnableHoverHighlight Then node.BackColor = SystemColors.Highlight
            If Not EnableHoverHighlight Then node.ForeColor = SystemColors.HighlightText
        Else
            _selectedNodes.Remove(node)
            If Not EnableHoverHighlight Then node.BackColor = Me.BackColor
            If Not EnableHoverHighlight Then node.ForeColor = Me.ForeColor
        End If
        ' 2026/5/14 by simon/Gemini: 將mouse hover作成內建功能, 去除上方的color設置, 只用這行me.invalidate
        If EnableHoverHighlight Then Me.Invalidate(node.Bounds)

        ' ✅ 同步更新基類的 SelectedNode（單選時讓原生 API 也能讀到正確值）
        MyBase.SelectedNode = If(_selectedNodes.Count = 1, _selectedNodes(0), Nothing)
    End Sub
    Private Sub PaintSelectedNodes()
        ' PaintSelectedNodes：重新套用選取高亮顏色（得焦時使用）
        For Each node As TreeNode In _selectedNodes
            If Not EnableHoverHighlight Then node.BackColor = SystemColors.Highlight
            If Not EnableHoverHighlight Then node.ForeColor = SystemColors.HighlightText
        Next
        ' 2026/5/14 by simon/Gemini: 將mouse hover作成內建功能, 去除上方的color設置, 只用這行me.invalidate
        If EnableHoverHighlight Then Me.Invalidate()
    End Sub
    Protected Overrides Sub OnDrawNode(e As DrawTreeNodeEventArgs)
        ' 2026/05/24 by Claude Sonnet 4.6: 修正含 & 字元節點的背景與文字截斷問題
        ' 根因: OwnerDrawText 模式 e.Bounds.Width 以「& 為前綴」計算，比 NoPrefix 實際寬度短
        '       需自行測量實際寬度，同時用 BackColor 先清 e.Bounds 防 native 選取框殘留
        '
        ' 2026/07/10 by Simon/Claude [效能決策紀錄 — 決定不再優化，別再重新研究]:
        '   曾評估把下方每次 New SolidBrush 改成欄位級快取。結論不做：OnDrawNode 只對「可見節點」觸發
        '   （實務上一屏 30~100 個），一次完整重繪僅產生 ~百餘個微秒級短命 GDI 物件，Using 已正確 Dispose，
        '   量不出體感差異；快取反而要處理 BackColor/HoverColor 變更時重建 + 控制項 Dispose 的生命週期。結案不做。

        Dim fontToUse As Font = If(e.Node.NodeFont, Me.Font)

        ' 1. 測量 NoPrefix 實際文字寬度
        Dim actualW As Integer = TextRenderer.MeasureText(e.Graphics, e.Node.Text, fontToUse, New Size(Integer.MaxValue, e.Bounds.Height), TextFormatFlags.NoPrefix Or TextFormatFlags.NoPadding).Width

        ' 2. 有效繪製矩形（不超出控制項右緣）
        Dim drawW As Integer = Math.Min(actualW + 6, Me.ClientSize.Width - e.Bounds.Left)
        Dim drawRect As New Rectangle(e.Bounds.Left, e.Bounds.Top, drawW, e.Bounds.Height)

        ' 3. 決定背景色
        Dim bgColor As Color
        If _selectedNodes.Contains(e.Node) Then
            bgColor = If(_isHighlightActive, SystemColors.Highlight, Form1.ThemeColors.MercuryGray)
        ElseIf e.Node Is _LastHoverNodeInternal AndAlso EnableHoverHighlight AndAlso Me.Cursor <> Cursors.WaitCursor Then
            bgColor = HoverColor
        Else
            bgColor = Me.BackColor
        End If

        ' 4. 先清 e.Bounds（抹除 native mouse-down 選取框），再填正確寬度的背景
        Using b As New SolidBrush(Me.BackColor) : e.Graphics.FillRectangle(b, e.Bounds) : End Using
        Using b As New SolidBrush(bgColor) : e.Graphics.FillRectangle(b, drawRect) : End Using

        ' 5. 決定文字顏色
        Dim txtColor As Color = If(_selectedNodes.Contains(e.Node), If(_isHighlightActive, SystemColors.HighlightText, SystemColors.InactiveCaptionText), Me.ForeColor)

        ' 6. SetClip 防止 GDI 文字超出控制項產生殘影，drawRect 同時作為文字繪製範圍
        e.Graphics.SetClip(drawRect)
        TextRenderer.DrawText(e.Graphics, e.Node.Text, fontToUse, drawRect, txtColor,
        TextFormatFlags.VerticalCenter Or TextFormatFlags.NoPrefix)
        e.Graphics.ResetClip()
    End Sub
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

#Region "■ 07 公共輔助方法（供 Form1 呼叫）"
    Public Sub ClearSelectedNodes()
        ' ClearSelectedNodes：清除所有選取（顏色還原 + 清空清單）
        ' Form1 在切換模式時呼叫，避免殘留不正確的高亮
        ' by Gemini, 2026/04/07: 清空大量選取時加上 Begin/End Update 避免背景色重複重繪導致閃爍
        ' by Gemini 3.0 Flash, 2026/04/17: 再次確認重置效力，解決 Nodes.Clear 導致的 stale 引用問題
        Me.BeginUpdate()
        Try
            For Each node As TreeNode In _selectedNodes
                If Not EnableHoverHighlight Then node.BackColor = Me.BackColor
                If Not EnableHoverHighlight Then node.ForeColor = Me.ForeColor
            Next
        Finally
            Me.EndUpdate()
        End Try

        _selectedNodes.Clear()
        ' 2026/5/14 by simon/Gemini: 將mouse hover作成內建功能, 去除上方的color設置, 只用這行me.invalidate
        If EnableHoverHighlight Then Me.Invalidate()

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
        ' 例如 GotoDefaultInbox 初始化選取後使用
        If node Is Nothing Then Return
        ToggleSingleNode(node, True)
        _lastClickedNode = node
    End Sub
    Public Sub SetSelectedNode(node As TreeNode)
        ' SetSelectedNode：等同 SelectedNode = node
        ' （保留相容舊版呼叫，不觸發 AfterSelect）
        SelectSingleNode(node)
    End Sub
    Public Function GetDedupedSelection() As List(Of TreeNode)
        ' GetDedupedSelection：父子去重，回傳目前選取節點中排除掉「祖先也在選取清單內」的節點
        ' 用途：使用者同時選中父資料夾與其子孫節點時，避免呼叫端重複計算統計（父的統計本就已含子孫）
        ' 2026/07/10 by Simon/Claude: 從 Form1_MainTab12.vb 的 GetDeDupedNodes 搬入，改讀 Me.SelectedNodes，讓其他 Tab 可直接重用

        Dim nodes As List(Of TreeNode) = Me.SelectedNodes
        Dim selectedSet As New HashSet(Of TreeNode)(nodes)
        Dim dedupedNodes As New List(Of TreeNode)(nodes.Count)

        For Each node As TreeNode In nodes
            Dim isDescendantOfSelected As Boolean = False
            Dim ancestor As TreeNode = node.Parent
            While ancestor IsNot Nothing
                If selectedSet.Contains(ancestor) Then isDescendantOfSelected = True : Exit While
                ancestor = ancestor.Parent
            End While
            If Not isDescendantOfSelected Then dedupedNodes.Add(node)
        Next
        Return dedupedNodes
    End Function

    ' =========================================================================
    ' 整合後的路徑導覽與狀態管理 API (2026/05/14 重構)
    ' 解決了 FindNodeByPath, CollectExpandedPaths 等分散函數，統一由控制項內部管理
    ' 2026/07/10: ContainsPath(fullPath) 已補上薄 wrapper (見 GetNode 下方)，目前專案內尚無呼叫端，備用
    ' =========================================================================
    Public Function GetNode(fullPath As String, Optional searchOnlyExpanded As Boolean = True) As TreeNode
        ''' <summary>
        ''' [核心尋路引擎 ] 簡化物件回傳版本，TryGetNode 的回傳值包裝。
        ''' 適用於呼叫端只需要節點物件、不需要成功/失敗布林值的場景。找不到時回傳 Nothing。
        ''' </summary>
        ''' 
        ''' 範例:
        '''   Dim n = SimTree1.GetNode("\\Personal Folders\Inbox")
        '''   If n IsNot Nothing Then n.Expand()
        ''' 
        ''' 2026/05/21 by Simon/Claude
        ''' <param stateNAme="fullPath">完整路徑，例如 \\Personal Folders\Inbox\SubFolder</param>
        ''' <param stateNAme="searchOnlyExpanded">True = 只在已展開節點中搜尋（預設）；False = 允許搜尋未展開節點</param>
        ''' <returns>找到的 TreeNode；找不到時回傳 Nothing</returns>
        Dim n As TreeNode = Nothing
        Return If(TryGetNode(fullPath, n, searchOnlyExpanded), n, Nothing)
        ' 2026/07/10: 曾評估加 path→node Dictionary 快取，實測後結案不做（理由見 RestoreTreeState 的效能決策紀錄）
    End Function
    Public Function ContainsPath(fullPath As String, Optional searchOnlyExpanded As Boolean = True) As Boolean
        ''' <summary>
        ''' 判斷指定路徑的節點是否存在（TryGetNode 的布林薄包裝，不回傳節點物件）。
        ''' 2026/07/10 by Simon/Claude: 依 Todo 補上；目前專案內尚無任何呼叫端，備用。
        ''' </summary>
        Dim n As TreeNode = Nothing
        Return TryGetNode(fullPath, n, searchOnlyExpanded)
    End Function
    Public Function TryGetNode(fullPath As String, ByRef returnNode As TreeNode,
                               Optional searchOnlyExpanded As Boolean = True, Optional expandAlongTheWay As Boolean = False) As Boolean
        ''' <summary>
        ''' [核心尋路引擎] 路徑切段尋路法 (Path-Segment Routing)尋找指定路徑的節點。
        ''' </summary>
        ''' <param stateNAme="fullPath">目標的完整路徑，例如 "\\Personal Folders\Inbox\SubFolder"</param>
        ''' <param stateNAme="returnNode">ByRef 回傳找到的目標節點 TreeNode；找不到時為 Nothing</param>
        ''' <param stateNAme="searchOnlyExpanded">True = 只在已展開的節點中搜尋 (效能最佳)；False=允許搜尋未展開節點</param>
        ''' <param stateNAme="expandAlongTheWay">True = 沿途 Expand 未展開節點，觸發 BeforeExpand 載入真實子節點</param>
        ''' <param stateNAme="selectAndFire">True = 找到節點後，自動清除舊選取、將其設為唯一選取並觸發 AfterSelect</param>
        ''' <param stateNAme="ensureVisible">True = 找到後呼叫 EnsureVisible 捲動SimTree, 確保節點在畫面可見範圍（預設 False）</param>
        ''' <returns>True = 成功找到並填入 returnNode；False = 路徑任一段找不到</returns>
        ''' 
        ''' 設計背景:
        '''   原 FindNodeByFullPath 使用 TreeNodeCollection.Find() 比對 node.Name，但 Tab1～5 的所有 PST 節點 .Name 均為空字串
        '''   （LoadStoreToTreeView / LoadSubFolderToTreeView 建立節點時從未設定 .Name），導致原函數對 Tab1～5 實際上完全無效。
        '''   本函數改為比對 node.Text，與 Outlook FolderPath 各段名稱一致，正確可靠。
        ''' 
        ''' 路徑格式：Outlook FolderPath 標準格式，例如 \\Personal Folders\Inbox\2024
        '''   以反斜線切段，RemoveEmptyEntries 自動去除開頭雙反斜線產生的空字串段。
        ''' 
        ''' 比對策略：node.Text 大小寫不分（OrdinalIgnoreCase），跳過 ::: 佔位節點。
        ''' 
        ''' expandAlongTheWay 與 searchOnlyExpanded 的優先關係:
        '''   expandAlongTheWay = True  → 沿途主動 Expand()，觸發 BeforeExpand Lazy Load；
        '''                               此時 searchOnlyExpanded 被忽略（條件互斥）
        '''   expandAlongTheWay = False → 以 searchOnlyExpanded 決定是否進入未展開節點
        '''
        ''' 2026/05/21 by Simon/Claude

        returnNode = Nothing
        If String.IsNullOrEmpty(fullPath) Then Return False

        ' ── Step 1: 路徑切段 ────────────────────────────────────────────
        ' "\\Personal Folders\Inbox\Sub" → ["Personal Folders", "Inbox", "Sub"]
        ' 切割後如果有空字串，RemoveEmptyEntries 自動去除開頭雙反斜線產生的空字串part
        Dim parts() As String = fullPath.Split("\"c, StringSplitOptions.RemoveEmptyEntries)
        If parts.Length = 0 Then Return False

        ' ── Step 2: 逐段比對節點 ────────────────────────────────────────
        Dim currentNodes As TreeNodeCollection = Me.Nodes
        Dim lastFoundNode As TreeNode = Nothing

        For i As Integer = 0 To parts.Length - 1
            Dim part As String = parts(i)
            Dim isLastSegment As Boolean = (i = parts.Length - 1)
            Dim found As Boolean = False

            ' 在目前層找符合當前路徑段的節點（跳過 ::: 佔位符）
            For Each node As TreeNode In currentNodes
                If node.Text = ":::" Then Continue For     ' 跳過 Lazy Load 佔位節點

                If String.Equals(node.Text, part, StringComparison.OrdinalIgnoreCase) Then
                    lastFoundNode = node
                    found = True

                    ' ── Step 3: 非末段，處理子層展開與通行判斷 ──────────
                    If Not isLastSegment Then
                        ' expandAlongTheWay 優先：主動展開，觸發 BeforeExpand → LoadSubFolderToTreeView
                        If expandAlongTheWay AndAlso Not node.IsExpanded Then node.Expand()

                        ' searchOnlyExpanded 防護（僅在不沿路展開時生效，兩者互斥）
                        If searchOnlyExpanded AndAlso Not expandAlongTheWay AndAlso Not node.IsExpanded Then Return False    ' 不允許進入未展開節點

                        ' 展開後子節點仍只有 ::: → Lazy Load 未觸發或資料夾確實為空，路徑中斷
                        If node.Nodes.Count = 0 OrElse (node.Nodes.Count = 1 AndAlso node.Nodes(0).Text = ":::") Then Return False

                        currentNodes = node.Nodes
                    End If
                    Exit For    ' 本段已找到，跳出內層迴圈繼續處理下一段
                End If
            Next

            If Not found Then Return False  ' 本段路徑斷開
        Next

        ' ── Step 4: 找到目標，回傳 ──────────────────────────────────────────
        If lastFoundNode Is Nothing Then Return False
        returnNode = lastFoundNode
        Return True

    End Function
    Public Function SelectNode(fullPath As String, Optional selectAndFire As Boolean = True, Optional expandTarget As Boolean = False) As Boolean
        ''' <summary>
        ''' 萬用導覽函數：根據路徑展開並選取節點。(呼叫TryGetNode)
        ''' 涵蓋了舊版 SelectNode 與 SelectNodeByPathRecursive 的功能。
        ''' </summary>
        ''' <param stateNAme="fullPath">目標資料夾的完整路徑</param>
        ''' <param stateNAme="selectAndFire">是否觸發 AfterSelect 事件 (預設 True)</param>
        ''' <param stateNAme="expandTarget">是否展開目標節點 (預設 False)</param>
        ''' by Gemini 3.5 Flash, 2026/05/21: 重構以完全使用 TryGetNode 核心引擎
        Dim targetNode As TreeNode = Nothing
        If TryGetNode(fullPath, targetNode, searchOnlyExpanded:=False, expandAlongTheWay:=True) Then
            ClearSelectedNodes()
            AddSelectedNode(targetNode)
            targetNode.EnsureVisible()

            If expandTarget Then targetNode.Expand()
            If selectAndFire Then FireAfterSelect(targetNode)

            Return True
        End If
        Return False
    End Function
#End Region

#Region "■ 08 視圖狀態堆疊 (公用，Save/Restore Tree State)"
    ' =========================================================================
    ' SaveTreeNodeSnap / RestoreTreeNodeSnap — 具名插槽式快照
    ' =========================================================================
    ' 用途：Tab4「資料夾模式 <--> 結果模式」的快速切換。
    '   比 SaveTreeStateByPath/RestoreTreeState（路徑字串）快，直接存 TreeNode 物件參考。
    '   語意更清晰，加上可選項名稱，可同時存下多個不同用途的插槽。
    '
    ' 重要：_selectedNodes 必須在 Nodes.Clear() 之前讀取（直接存物件參考）。
    '       SelectedNodes getter 有 RemoveAll(n.TreeView IsNot Me)，
    '       Clear() 後 n.TreeView = Nothing，走 getter 會把節點全部過濾掉。
    '       SaveTreeStateByPath() 必須在 "開始搜尋之後, 顯示結果之前" 的時機呼叫，確保讀到正確的選取狀態。
    '
    ' 2026/05/23 by Simon/Claude: 封裝 Tab4「資料夾↔結果模式」快速切換所需的狀態
    ' 注意：_selectedNodes 必須在 Nodes.Clear() 之前讀取（存物件參考，非走 getter）
    '       因為 Clear() 後 n.TreeView = Nothing，SelectedNodes getter 會自動過濾掉節點
    ' =========================================================================
    ' ■ 08-A  Node 物件快照（SaveTreeNodeSnap / RestoreTreeNodeSnap）
    ' 用途：Tab4「資料夾模式 ↔ 結果模式」來回切換，節點物件直接存，不過 Outlook COM
    ' ■ 08-B  路徑字串快照（SaveTreeStateByPath / RestoreTreeState）
    ' 用途：F5 強制刷新，存路徑字串，刷新後觸發 LazyLoad 重讀 Outlook COM
    '        若資料夾在 Outlook 已消失，重展開時天然不出現，不需另寫 diff 邏輯
    ' 2026/05/25 by Simon/Claude
    ' =========================================================================
    Private _nodeSnapshots As New Dictionary(Of String, TreeNodeSnap)()
    Private Class TreeNodeSnap
        Friend DetachedNodes As List(Of TreeNode)   ' 從樹上拔下的節點物件清單（含子樹，原封不動）
        Friend SelectedRefs As List(Of TreeNode)    ' 選取狀態（直接存物件參考，不存路徑）
        Friend LastClickedRef As TreeNode           ' _lastClickedNode 的物件參考
    End Class
    Public Sub SaveTreeNodeSnap(Optional stateName As String = "default")
        ''' <summary>
        ''' 將目前整棵樹的節點與選取狀態快照到具名插槽，並清空 TreeView。
        ''' 適用於需要暫時切換到另一個視圖、之後再還原的情境（如 Tab4 資料夾模式 vs 搜尋結果模式 的快速來回切換）。
        ''' (ps. 等同於將目前的畫面凍結並移到後台)
        ''' </summary>
        ''' <param stateName="stateName">插槽名稱，預設 "default"；可同時存多份供不同切換情境使用</param>
        ' ★ 必須先讀 _selectedNodes / _lastClickedNode，再呼叫 ClearSelectedNodes()
        ' 建立一個新的NodeSnapshot物件來存節點參考（不透過 getter），封裝狀態確保快照中包含完整的選取
        Dim snap As New TreeNodeSnap()
        snap.DetachedNodes = New List(Of TreeNode)(Me.Nodes.Cast(Of TreeNode)())
        snap.SelectedRefs = New List(Of TreeNode)(_selectedNodes)   ' 直接讀私有欄位，繞過 getter 的自動過濾
        snap.LastClickedRef = _lastClickedNode
        _nodeSnapshots(stateName) = snap                            ' 覆蓋式寫入，同名插槽直接更新

        ' 清空樹，準備迎接新資料（節點物件已被 snap.DetachedNodes 持有，不會被 GC 回收）
        Me.BeginUpdate()
        ClearSelectedNodes()                                        ' 同時清掉 _lastClickedNode
        Me.Nodes.Clear()
        Me.EndUpdate()
    End Sub
    Public Function RestoreTreeNodeSnap(Optional stateName As String = "default") As Boolean
        ''' <summary>
        ''' 從具名插槽還原節點與選取狀態。還原後自動捲動到上次選取的節點。
        ''' </summary>
        ''' <param stateName="stateName">插槽名稱，預設 "default"</param>
        ''' <returns>True = 還原成功；False = 插槽不存在（呼叫端應自行 Fallback，例如重新載入資料夾樹）</returns>
        Dim snap As TreeNodeSnap = Nothing
        If Not _nodeSnapshots.TryGetValue(stateName, snap) Then Return False
        _nodeSnapshots.Remove(stateName)    ' 用完即清，避免殘留舊快照

        Try
            ' 清空目前的視圖以及 _lastClickedNode
            Me.BeginUpdate()
            Me.Nodes.Clear()
            ClearSelectedNodes()

            ' 1. 還原節點（含子樹與展開狀態，原封不動，IsExpanded 狀態會自然保留）
            If snap.DetachedNodes.Count > 0 Then Me.Nodes.AddRange(snap.DetachedNodes.ToArray())

            ' 2. 還原選取（只加回仍屬於本樹的節點，跳過已失效的參考）
            For Each node As TreeNode In snap.SelectedRefs
                If node.TreeView Is Me AndAlso Not _selectedNodes.Contains(node) Then _selectedNodes.Add(node)
            Next

            ' 3. 還原 _lastClickedNode
            _lastClickedNode = If(snap.LastClickedRef IsNot Nothing AndAlso
                                  snap.LastClickedRef.TreeView Is Me,
                                  snap.LastClickedRef, Nothing)

            ' 4. 同步基類的 SelectedNode（讓原生 API 也能讀到正確值）
            MyBase.SelectedNode = If(_selectedNodes.Count = 1, _selectedNodes(0), Nothing)

            ' 5. 捲動到上次選取的節點
            _lastClickedNode?.EnsureVisible()

            ' 6. 重繪（OwnerDraw 模式需要手動觸發）
            If EnableHoverHighlight Then Me.Invalidate()

        Finally
            Me.EndUpdate()
        End Try
        Return True
    End Function
    Public Class SimTreeState
        Friend ExpandedPaths As List(Of String)
        Friend SelectedPaths As List(Of String)
    End Class
    Public Function SaveTreeStateByPath() As SimTreeState
        ' ---------------------------------------------------------------
        ' SaveTreeStateByPath — 快照目前展開與選取路徑（路徑字串，Nodes.Clear 後仍有效）
        ' 2026/05/25 by Simon/Claude
        ' ---------------------------------------------------------------
        Dim state As New SimTreeState()
        state.ExpandedPaths = New List(Of String)(64)
        state.SelectedPaths = New List(Of String)(64)

        SaveExpandedPathsInternal(Me.Nodes, state.ExpandedPaths)

        For Each node As TreeNode In _selectedNodes
            Dim p As String = SafeGetPathFromNode(node)
            If Not String.IsNullOrEmpty(p) Then state.SelectedPaths.Add(p)
        Next

        Return state
    End Function
    Public Sub RestoreTreeState(state As SimTreeState, Optional selectAndFire As Boolean = True)
        ' ---------------------------------------------------------------
        ' RestoreTreeState — 按路徑字串重展開並還原選取
        ' 展開動作觸發 BeforeExpand → LoadSubFolderToTreeView → GetSortedSubFolders，
        ' 天然重讀 Outlook COM；若資料夾已消失，節點不出現（不需另寫 diff）
        ' FireAfterSelect 在 EndUpdate 之後呼叫，確保 Layout 結算完畢再觸發統計
        ' 2026/05/25 by Simon/Claude: 取代舊版 LoadTreeState + RefreshTreeState
        ' ---------------------------------------------------------------
        ' 2026/07/10 by Simon/Claude [效能決策紀錄 — 決定不再優化，別再重新研究]:
        '   曾評估兩案：① TryGetNode 加 path→node Dictionary 快取 ② 多條路徑共享父層前綴的 Trie 式重用。
        '   實測（300 與 780 個資料夾的兩個 Profile）：F5 全樹重刷首次約 0.25~0.35s，再次重刷降到 0.15~0.2s。
        '   耗時大頭是 Expand() 觸發 BeforeExpand → LoadSubFolderToTreeView 的 COM/DB 載入（毫秒級/節點），
        '   TryGetNode 的純記憶體字串比對只佔毫秒級零頭 — 兩案都省不到體感，卻要背快取失效維護成本。結案不做。
        ' ---------------------------------------------------------------
        If state Is Nothing Then Return

        Me.BeginUpdate()
        Try
            ' 1. 還原展開（短路徑先，確保父節點先 Expand 再載入子節點）
            ' 注意：TryGetNode expandAlongTheWay=True 只展開路徑中間節點，不展開最後節點本身。
            ' 因為 ExpandedPaths 的每條路徑對應的節點都應該是展開狀態，找到後必須再 Expand 一次。
            ' 2026/05/26 by Simon/Claude: 修正還原展開狀態失敗的 bug
            If state.ExpandedPaths IsNot Nothing Then
                For Each path In state.ExpandedPaths.OrderBy(Function(p) p.Length)
                    Dim found As TreeNode = Nothing
                    If TryGetNode(path, found, searchOnlyExpanded:=False, expandAlongTheWay:=True) Then
                        If found IsNot Nothing AndAlso Not found.IsExpanded Then found.Expand()
                    End If
                Next
            End If

            ' 2. 還原選取
            ClearSelectedNodes()
            If state.SelectedPaths IsNot Nothing Then
                For Each path In state.SelectedPaths
                    Dim found As TreeNode = Nothing
                    If TryGetNode(path, found, searchOnlyExpanded:=True) Then AddSelectedNode(found)
                Next
            End If

            _lastClickedNode?.EnsureVisible()
        Finally
            Me.EndUpdate()
        End Try

        ' 3. AfterSelect 在 EndUpdate 之後，避免 Layout 未結算就開始計算統計
        If selectAndFire AndAlso _lastClickedNode IsNot Nothing Then FireAfterSelect(_lastClickedNode)
    End Sub
#End Region

#Region "■ 09 私有輔助"
    ' ── 以下為支撐上述 API 的 Private 核心邏輯 ──
    Private Sub SaveExpandedPathsInternal(nodes As TreeNodeCollection, paths As List(Of String))
        For Each n As TreeNode In nodes
            If n.IsExpanded Then
                Dim p As String = SafeGetPathFromNode(n)
                If Not String.IsNullOrEmpty(p) Then paths.Add(p)
                SaveExpandedPathsInternal(n.Nodes, paths)
            End If
        Next
    End Sub
    Private Function SafeGetPathFromNode(node As TreeNode) As String
        ' 假設你的 node.Tag 存放的是 Outlook.Folder，請將原本 Form1 的 SafeGetPath 邏輯搬進來或簡化
        ' 如果有存 Folder，取 Folder.FolderPath；否則取 Node.FullPath
        If node.Tag Is Nothing Then Return node.FullPath

        ' ObjectToFolderPath 本身有 Try/Catch，失敗時回傳 ""
        ' 不管 Tag 是什麼型別都試著叫它報路徑，成功就用，失敗才 fallback
        Dim path As String = ObjectToFolderPath(node.Tag)
        Return If(Not String.IsNullOrEmpty(path), path, node.FullPath)
    End Function
    Private Function ObjectToFolderPath(obj As Object) As String
        Try
            Return CStr(Microsoft.VisualBasic.Interaction.CallByName(obj, "FolderPath", Microsoft.VisualBasic.CallType.Get))
        Catch ex As Exception
            Return ""
        End Try
    End Function
#End Region

End Class
