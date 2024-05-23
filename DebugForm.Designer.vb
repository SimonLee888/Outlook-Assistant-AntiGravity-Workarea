<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DebugForm
    Inherits System.Windows.Forms.Form

    'Form 覆寫 Dispose 以清除元件清單。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    '為 Windows Form 設計工具的必要項
    Private components As System.ComponentModel.IContainer

    '注意: 以下為 Windows Form 設計工具所需的程序
    '可以使用 Windows Form 設計工具進行修改。
    '請勿使用程式碼編輯器進行修改。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DebugForm))
        lvwDebug = New ListView()
        ColumnHeader1 = New ColumnHeader()
        ColumnHeader2 = New ColumnHeader()
        ColumnHeader3 = New ColumnHeader()
        SuspendLayout()
        ' 
        ' lvwDebug
        ' 
        lvwDebug.AutoArrange = False
        lvwDebug.Columns.AddRange(New ColumnHeader() {ColumnHeader1, ColumnHeader2, ColumnHeader3})
        resources.ApplyResources(lvwDebug, "lvwDebug")
        lvwDebug.FullRowSelect = True
        lvwDebug.GridLines = True
        lvwDebug.MultiSelect = False
        lvwDebug.Name = "lvwDebug"
        lvwDebug.ShowGroups = False
        lvwDebug.TabStop = False
        lvwDebug.UseCompatibleStateImageBehavior = False
        lvwDebug.View = View.Details
        ' 
        ' ColumnHeader1
        ' 
        resources.ApplyResources(ColumnHeader1, "ColumnHeader1")
        ' 
        ' ColumnHeader2
        ' 
        resources.ApplyResources(ColumnHeader2, "ColumnHeader2")
        ' 
        ' ColumnHeader3
        ' 
        resources.ApplyResources(ColumnHeader3, "ColumnHeader3")
        ' 
        ' DebugForm
        ' 
        resources.ApplyResources(Me, "$this")
        AutoScaleMode = AutoScaleMode.Font
        AutoValidate = AutoValidate.EnableAllowFocusChange
        Controls.Add(lvwDebug)
        Name = "DebugForm"
        SizeGripStyle = SizeGripStyle.Show
        ResumeLayout(False)
    End Sub

    Friend WithEvents lvwDebug As ListView
    Friend WithEvents ColumnHeader1 As ColumnHeader
    Friend WithEvents ColumnHeader2 As ColumnHeader
    Friend WithEvents ColumnHeader3 As ColumnHeader
End Class
