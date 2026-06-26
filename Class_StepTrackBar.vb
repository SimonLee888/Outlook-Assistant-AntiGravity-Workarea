Imports System.Windows.Forms
Imports System.Runtime.InteropServices

' StepTrackBar — 自訂 TrackBar，修正兩個原生行為：
'   (1) 點擊滑軌空白處：原生會直接跳到點擊位置，改為只移動一格。
'   (2) 鍵盤方向反置：
'         原生 PageUp = 往左(值-LargeChange)、PageDown = 往右(值+LargeChange)，
'         ↑↓ 箭頭與直覺相同(↑=+1、↓=-1)，但 PageUp/PageDown 與直覺相反。
'         本類別將 PageUp/PageDown 對調，使 PageUp = 值+LargeChange (往高)，
'         PageDown = 值-LargeChange (往低)，↑/↓ 亦對調與之一致。
'         -- by Claude Sonnet 4.6, 2026/06/18
'
' 實作方式：
'   (1) WM_LBUTTONDOWN：記錄按下前 Value，讓原生處理後限制在 ±SmallChange 範圍。
'   (2) WM_KEYDOWN：攔截 PageUp/PageDown/↑/↓，手動調整 Value 後 return，
'       不交給原生處理，避免原生再跑一次反向邏輯。
'   by Claude Sonnet 4.6, 2026/06/18

Public Class StepTrackBar
    Inherits TrackBar

    Private Const WM_LBUTTONDOWN As Integer = &H201
    Private Const WM_KEYDOWN As Integer = &H100

    ' 虛擬鍵碼
    Private Const VK_PRIOR As Integer = &H21  ' Page Up
    Private Const VK_NEXT As Integer = &H22   ' Page Down
    Private Const VK_UP As Integer = &H26     ' ↑
    Private Const VK_DOWN As Integer = &H28   ' ↓

    ' 記錄滑鼠按下時的 Value，用來計算允許移動的範圍
    Private _valueBeforeClick As Integer = -1
    Private _isClickInProgress As Boolean = False

    Protected Overrides Sub WndProc(ByRef m As Message)

        ' ── 鍵盤對調：PageUp=+LargeChange、PageDown=-LargeChange；↑=+1、↓=-1 ──
        ' 原生 TrackBar 水平方向：PageUp 往左(值減)、PageDown 往右(值增)，與直覺相反。
        ' 此處攔截後自行調整 Value，不交回原生，避免原生再跑一次反向邏輯。
        ' by Claude Sonnet 4.6, 2026/06/18
        If m.Msg = WM_KEYDOWN Then
            Dim key As Integer = m.WParam.ToInt32()
            Dim delta As Integer = 0
            Select Case key
                Case VK_PRIOR : delta = LargeChange   ' PageUp  → 往高
                Case VK_NEXT : delta = -LargeChange  ' PageDown → 往低
                Case VK_UP : delta = SmallChange   ' ↑ → 往高
                Case VK_DOWN : delta = -SmallChange  ' ↓ → 往低
            End Select

            If delta <> 0 Then
                Me.Value = Math.Max(Me.Minimum, Math.Min(Me.Maximum, Me.Value + delta))
                Return  ' 已處理，不交給原生
            End If
        End If

        ' ── 滑鼠點擊滑軌：限制最多移動一格 ──
        If m.Msg = WM_LBUTTONDOWN Then
            _valueBeforeClick = Me.Value
            _isClickInProgress = True

            MyBase.WndProc(m)   ' 讓原生 TrackBar 處理（拇指可能跳位）

            If _isClickInProgress AndAlso Me.Value <> _valueBeforeClick Then
                Dim corrected As Integer = If(Me.Value > _valueBeforeClick, _valueBeforeClick + 1, _valueBeforeClick - 1)
                Me.Value = Math.Max(Me.Minimum, Math.Min(Me.Maximum, corrected))
            End If

            _isClickInProgress = False
            Return
        End If

        MyBase.WndProc(m)
    End Sub

End Class
