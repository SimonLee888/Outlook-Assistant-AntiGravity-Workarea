Partial Class Form1

#Region "■ 01 全域宣告"
#Region "  ├ Win32 API 宣告"
    ' ── 函數宣告 ────────────────────────────────────────────────────────────────
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
    Private Shared Function ShowWindow(ByVal hWnd As IntPtr, ByVal nCmdShow As Integer) As Boolean

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
    Private Shared Function GetWindowLong(hWnd As IntPtr, nIndex As Integer) As Integer

    End Function
    <Runtime.InteropServices.DllImport("user32.dll")>
    Private Shared Function SetWindowLong(hWnd As IntPtr, nIndex As Integer, dwNewLong As Integer) As Integer

    End Function

    Private Const GWL_STYLE As Integer = -16
    Private Const WS_TABSTOP As Integer = &H10000
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
    Private Const RDW_ALLCHILDREN As Integer = &H80                 ' debugForm resize 時閃爍, 改手動redraw
    Private Const RDW_INVALIDATE As Integer = &H1                   ' debugForm resize 時閃爍, 改手動redraw
    Private Const RDW_UPDATENOW As Integer = &H100                  ' debugForm resize 時閃爍, 改手動redraw
    Private Const RDW_ERASE As Integer = &H4                        ' 2026/03/28 by AntiGravity: 補上缺失定義
    Private Const RDW_FRAME As Integer = &H400                      ' 2026/03/28 by AntiGravity: 補上缺失定義
    Private Const WM_SETREDRAW As Integer = &HB                     ' 2026/3/26 by AntiGravity
    ' ↓ 新增 (2026-03-20) ListView1 進入資料夾用
    Private Const TVM_SELECTITEM As Integer = &H110B                ' = &H1100 + 11
    Private Const TVGN_CARET As Integer = &H9                       ' SendMessage 選取 Treeview 游標節點
    Private Const LVM_SETITEMCOUNT As Integer = &H1000 + 47         ' = &H102F '
#End Region
#End Region

End Class
