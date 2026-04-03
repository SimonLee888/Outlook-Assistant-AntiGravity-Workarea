<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1

    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.

    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim ChartArea1 As System.Windows.Forms.DataVisualization.Charting.ChartArea = New DataVisualization.Charting.ChartArea()
        Dim Legend1 As System.Windows.Forms.DataVisualization.Charting.Legend = New DataVisualization.Charting.Legend()
        Dim Series1 As System.Windows.Forms.DataVisualization.Charting.Series = New DataVisualization.Charting.Series()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        TabControl1 = New TabControl()
        TabPage1 = New TabPage()
        SplitContainer1 = New SplitContainer()
        TreeView1 = New TreeView()
        ListView1 = New ListView()
        ColumnHeader11 = New ColumnHeader()
        ColumnHeader12 = New ColumnHeader()
        ColumnHeader13 = New ColumnHeader()
        ColumnHeader14 = New ColumnHeader()
        ColumnHeader15 = New ColumnHeader()
        TabPage2 = New TabPage()
        SplitContainer2 = New SplitContainer()
        CheckSubFolder2 = New CheckBox()
        ListView2 = New ListView()
        ColumnHeader21 = New ColumnHeader()
        ColumnHeader22 = New ColumnHeader()
        ColumnHeader23 = New ColumnHeader()
        Chart2 = New DataVisualization.Charting.Chart()
        TabPage3 = New TabPage()
        SplitContainer3 = New SplitContainer()
        TreeView3 = New TreeView()
        CheckSubFolder3 = New CheckBox()
        Button3_Stop = New Button()
        Button3 = New Button()
        GroupBox3 = New GroupBox()
        CountMax = New NumericUpDown()
        CountMin = New NumericUpDown()
        CheckAttCount = New CheckBox()
        Label3 = New Label()
        ListView3 = New ListView()
        ColumnHeader31 = New ColumnHeader()
        ColumnHeader32 = New ColumnHeader()
        ColumnHeader34 = New ColumnHeader()
        ColumnHeader33 = New ColumnHeader()
        ColumnHeader35 = New ColumnHeader()
        ColumnHeader36 = New ColumnHeader()
        GroupBox1 = New GroupBox()
        TextBox3 = New TextBox()
        CheckAttachName = New CheckBox()
        GroupBox2 = New GroupBox()
        UnitMax = New ComboBox()
        UnitMin = New ComboBox()
        NumberMax = New NumericUpDown()
        NumberMin = New NumericUpDown()
        CheckSize = New CheckBox()
        Label1 = New Label()
        TabPage4 = New TabPage()
        SplitContainer4 = New SplitContainer()
        TreeView4 = New TreeView()
        ListView4 = New ListView()
        ColumnHeader9 = New ColumnHeader()
        ColumnHeader10 = New ColumnHeader()
        Button4 = New Button()
        TabPage5 = New TabPage()
        SplitContainer5 = New SplitContainer()
        TreeView5 = New TreeView()
        lstEmails = New ListBox()
        TextBox2 = New TextBox()
        TextBox1 = New TextBox()
        Button5 = New Button()
        Label2 = New Label()
        TabPage6 = New TabPage()
        LoadCache = New Button()
        SaveCache = New Button()
        checkIncludeAllFolders = New CheckBox()
        CheckRDO = New CheckBox()
        buttonClearCache = New Button()
        CheckDebug = New CheckBox()
        ToolStripStatusLabel1 = New ToolStripStatusLabel()
        ProgressBar1 = New ToolStripStatusLabel()
        ProgressBar2 = New ToolStripStatusLabel()
        ToolStripProgressBar1 = New ToolStripStatusLabel()
        StatusStrip1 = New StatusStrip()
        TabControl1.SuspendLayout()
        TabPage1.SuspendLayout()
        CType(SplitContainer1, ComponentModel.ISupportInitialize).BeginInit()
        SplitContainer1.Panel1.SuspendLayout()
        SplitContainer1.Panel2.SuspendLayout()
        SplitContainer1.SuspendLayout()
        TabPage2.SuspendLayout()
        CType(SplitContainer2, ComponentModel.ISupportInitialize).BeginInit()
        SplitContainer2.Panel2.SuspendLayout()
        SplitContainer2.SuspendLayout()
        CType(Chart2, ComponentModel.ISupportInitialize).BeginInit()
        TabPage3.SuspendLayout()
        CType(SplitContainer3, ComponentModel.ISupportInitialize).BeginInit()
        SplitContainer3.Panel1.SuspendLayout()
        SplitContainer3.Panel2.SuspendLayout()
        SplitContainer3.SuspendLayout()
        GroupBox3.SuspendLayout()
        CType(CountMax, ComponentModel.ISupportInitialize).BeginInit()
        CType(CountMin, ComponentModel.ISupportInitialize).BeginInit()
        GroupBox1.SuspendLayout()
        GroupBox2.SuspendLayout()
        CType(NumberMax, ComponentModel.ISupportInitialize).BeginInit()
        CType(NumberMin, ComponentModel.ISupportInitialize).BeginInit()
        TabPage4.SuspendLayout()
        CType(SplitContainer4, ComponentModel.ISupportInitialize).BeginInit()
        SplitContainer4.Panel1.SuspendLayout()
        SplitContainer4.Panel2.SuspendLayout()
        SplitContainer4.SuspendLayout()
        TabPage5.SuspendLayout()
        CType(SplitContainer5, ComponentModel.ISupportInitialize).BeginInit()
        SplitContainer5.Panel1.SuspendLayout()
        SplitContainer5.Panel2.SuspendLayout()
        SplitContainer5.SuspendLayout()
        TabPage6.SuspendLayout()
        StatusStrip1.SuspendLayout()
        SuspendLayout()
        ' 
        ' TabControl1
        ' 
        TabControl1.Controls.Add(TabPage1)
        TabControl1.Controls.Add(TabPage2)
        TabControl1.Controls.Add(TabPage3)
        TabControl1.Controls.Add(TabPage4)
        TabControl1.Controls.Add(TabPage5)
        TabControl1.Controls.Add(TabPage6)
        TabControl1.Dock = DockStyle.Fill
        TabControl1.Font = New Font("Microsoft JhengHei UI", 10F, FontStyle.Regular, GraphicsUnit.Point, CByte(136))
        TabControl1.Location = New Point(0, 0)
        TabControl1.Margin = New Padding(4)
        TabControl1.Name = "TabControl1"
        TabControl1.Padding = New Point(6, 6)
        TabControl1.SelectedIndex = 0
        TabControl1.Size = New Size(1006, 1319)
        TabControl1.TabIndex = 2
        ' 
        ' TabPage1
        ' 
        TabPage1.Controls.Add(SplitContainer1)
        TabPage1.Location = New Point(4, 37)
        TabPage1.Margin = New Padding(4)
        TabPage1.Name = "TabPage1"
        TabPage1.Padding = New Padding(4)
        TabPage1.Size = New Size(998, 1278)
        TabPage1.TabIndex = 0
        TabPage1.Text = "資料夾統計"
        TabPage1.UseVisualStyleBackColor = True
        ' 
        ' SplitContainer1
        ' 
        SplitContainer1.Dock = DockStyle.Fill
        SplitContainer1.Location = New Point(4, 4)
        SplitContainer1.Margin = New Padding(4)
        SplitContainer1.Name = "SplitContainer1"
        ' 
        ' SplitContainer1.Panel1
        ' 
        SplitContainer1.Panel1.Controls.Add(TreeView1)
        ' 
        ' SplitContainer1.Panel2
        ' 
        SplitContainer1.Panel2.Controls.Add(ListView1)
        SplitContainer1.Size = New Size(990, 1270)
        SplitContainer1.SplitterDistance = 301
        SplitContainer1.SplitterWidth = 5
        SplitContainer1.TabIndex = 4
        SplitContainer1.TabStop = False
        ' 
        ' TreeView1
        ' 
        TreeView1.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        TreeView1.Font = New Font("Microsoft JhengHei UI", 10F)
        TreeView1.HideSelection = False
        TreeView1.Location = New Point(0, 0)
        TreeView1.Margin = New Padding(4)
        TreeView1.Name = "TreeView1"
        TreeView1.Size = New Size(300, 1244)
        TreeView1.TabIndex = 2
        ' 
        ' ListView1
        ' 
        ListView1.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        ListView1.Columns.AddRange(New ColumnHeader() {ColumnHeader11, ColumnHeader12, ColumnHeader13, ColumnHeader14, ColumnHeader15})
        ListView1.Font = New Font("Microsoft JhengHei UI", 10F)
        ListView1.FullRowSelect = True
        ListView1.Location = New Point(0, 0)
        ListView1.Margin = New Padding(4)
        ListView1.Name = "ListView1"
        ListView1.Size = New Size(612, 1244)
        ListView1.TabIndex = 3
        ListView1.UseCompatibleStateImageBehavior = False
        ListView1.View = View.Details
        ' 
        ' ColumnHeader11
        ' 
        ColumnHeader11.Text = "資料夾名稱"
        ColumnHeader11.Width = 150
        ' 
        ' ColumnHeader12
        ' 
        ColumnHeader12.Text = "郵件數量"
        ColumnHeader12.TextAlign = HorizontalAlignment.Right
        ColumnHeader12.Width = 100
        ' 
        ' ColumnHeader13
        ' 
        ColumnHeader13.Text = "資料夾數量"
        ColumnHeader13.TextAlign = HorizontalAlignment.Right
        ColumnHeader13.Width = 100
        ' 
        ' ColumnHeader14
        ' 
        ColumnHeader14.Text = "郵件總計"
        ColumnHeader14.TextAlign = HorizontalAlignment.Right
        ColumnHeader14.Width = 100
        ' 
        ' ColumnHeader15
        ' 
        ColumnHeader15.Text = "資料夾大小"
        ColumnHeader15.TextAlign = HorizontalAlignment.Right
        ColumnHeader15.Width = 200
        ' 
        ' TabPage2
        ' 
        TabPage2.Controls.Add(SplitContainer2)
        TabPage2.Location = New Point(4, 37)
        TabPage2.Margin = New Padding(4)
        TabPage2.Name = "TabPage2"
        TabPage2.Padding = New Padding(4)
        TabPage2.Size = New Size(998, 1278)
        TabPage2.TabIndex = 1
        TabPage2.Text = "統計圖表"
        TabPage2.UseVisualStyleBackColor = True
        ' 
        ' SplitContainer2
        ' 
        SplitContainer2.Dock = DockStyle.Fill
        SplitContainer2.Location = New Point(4, 4)
        SplitContainer2.Margin = New Padding(4)
        SplitContainer2.Name = "SplitContainer2"
        ' 
        ' SplitContainer2.Panel2
        ' 
        SplitContainer2.Panel2.Controls.Add(CheckSubFolder2)
        SplitContainer2.Panel2.Controls.Add(ListView2)
        SplitContainer2.Panel2.Controls.Add(Chart2)
        SplitContainer2.Size = New Size(990, 1270)
        SplitContainer2.SplitterDistance = 301
        SplitContainer2.SplitterWidth = 5
        SplitContainer2.TabIndex = 6
        SplitContainer2.TabStop = False
        ' 
        ' CheckSubFolder2
        ' 
        CheckSubFolder2.AutoSize = True
        CheckSubFolder2.FlatStyle = FlatStyle.System
        CheckSubFolder2.Font = New Font("Microsoft JhengHei UI", 9.5F, FontStyle.Regular, GraphicsUnit.Point, CByte(136))
        CheckSubFolder2.Location = New Point(4, 400)
        CheckSubFolder2.Margin = New Padding(4)
        CheckSubFolder2.Name = "CheckSubFolder2"
        CheckSubFolder2.Size = New Size(120, 25)
        CheckSubFolder2.TabIndex = 3
        CheckSubFolder2.Text = "含子資料夾"
        CheckSubFolder2.UseVisualStyleBackColor = True
        ' 
        ' ListView2
        ' 
        ListView2.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        ListView2.Columns.AddRange(New ColumnHeader() {ColumnHeader21, ColumnHeader22, ColumnHeader23})
        ListView2.Font = New Font("Microsoft JhengHei UI", 10F)
        ListView2.FullRowSelect = True
        ListView2.Location = New Point(0, 0)
        ListView2.Margin = New Padding(4)
        ListView2.Name = "ListView2"
        ListView2.Size = New Size(612, 395)
        ListView2.TabIndex = 2
        ListView2.UseCompatibleStateImageBehavior = False
        ListView2.View = View.Details
        ' 
        ' ColumnHeader21
        ' 
        ColumnHeader21.Text = "年份"
        ColumnHeader21.Width = 200
        ' 
        ' ColumnHeader22
        ' 
        ColumnHeader22.Text = "郵件數量"
        ColumnHeader22.TextAlign = HorizontalAlignment.Right
        ColumnHeader22.Width = 100
        ' 
        ' ColumnHeader23
        ' 
        ColumnHeader23.Text = ""
        ' 
        ' Chart2
        ' 
        Chart2.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        Chart2.BackGradientStyle = DataVisualization.Charting.GradientStyle.TopBottom
        ChartArea1.Name = "ChartArea1"
        Chart2.ChartAreas.Add(ChartArea1)
        Legend1.Name = "Legend1"
        Chart2.Legends.Add(Legend1)
        Chart2.Location = New Point(0, 400)
        Chart2.Margin = New Padding(4)
        Chart2.Name = "Chart2"
        Series1.ChartArea = "ChartArea1"
        Series1.Legend = "Legend1"
        Series1.Name = "Series1"
        Chart2.Series.Add(Series1)
        Chart2.Size = New Size(581, 857)
        Chart2.TabIndex = 4
        Chart2.Text = "Chart1"
        ' 
        ' TabPage3
        ' 
        TabPage3.Controls.Add(SplitContainer3)
        TabPage3.Location = New Point(4, 37)
        TabPage3.Margin = New Padding(4)
        TabPage3.Name = "TabPage3"
        TabPage3.Padding = New Padding(4)
        TabPage3.Size = New Size(998, 1278)
        TabPage3.TabIndex = 2
        TabPage3.Text = "尋找附件檔案"
        TabPage3.UseVisualStyleBackColor = True
        ' 
        ' SplitContainer3
        ' 
        SplitContainer3.Dock = DockStyle.Fill
        SplitContainer3.Location = New Point(4, 4)
        SplitContainer3.Margin = New Padding(4)
        SplitContainer3.Name = "SplitContainer3"
        ' 
        ' SplitContainer3.Panel1
        ' 
        SplitContainer3.Panel1.Controls.Add(TreeView3)
        ' 
        ' SplitContainer3.Panel2
        ' 
        SplitContainer3.Panel2.Controls.Add(CheckSubFolder3)
        SplitContainer3.Panel2.Controls.Add(Button3_Stop)
        SplitContainer3.Panel2.Controls.Add(Button3)
        SplitContainer3.Panel2.Controls.Add(GroupBox3)
        SplitContainer3.Panel2.Controls.Add(ListView3)
        SplitContainer3.Panel2.Controls.Add(GroupBox1)
        SplitContainer3.Panel2.Controls.Add(GroupBox2)
        SplitContainer3.Size = New Size(990, 1270)
        SplitContainer3.SplitterDistance = 301
        SplitContainer3.SplitterWidth = 5
        SplitContainer3.TabIndex = 31
        SplitContainer3.TabStop = False
        ' 
        ' TreeView3
        ' 
        TreeView3.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        TreeView3.Font = New Font("Microsoft JhengHei UI", 10F)
        TreeView3.HideSelection = False
        TreeView3.Location = New Point(0, 0)
        TreeView3.Margin = New Padding(4)
        TreeView3.Name = "TreeView3"
        TreeView3.Size = New Size(300, 1244)
        TreeView3.TabIndex = 2
        ' 
        ' CheckSubFolder3
        ' 
        CheckSubFolder3.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        CheckSubFolder3.AutoSize = True
        CheckSubFolder3.FlatStyle = FlatStyle.System
        CheckSubFolder3.Location = New Point(522, 66)
        CheckSubFolder3.Margin = New Padding(4)
        CheckSubFolder3.Name = "CheckSubFolder3"
        CheckSubFolder3.Size = New Size(126, 27)
        CheckSubFolder3.TabIndex = 32
        CheckSubFolder3.Text = "含子資料夾"
        CheckSubFolder3.TextAlign = ContentAlignment.MiddleCenter
        CheckSubFolder3.UseVisualStyleBackColor = True
        ' 
        ' Button3_Stop
        ' 
        Button3_Stop.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        Button3_Stop.FlatStyle = FlatStyle.System
        Button3_Stop.Font = New Font("Microsoft JhengHei UI", 10F, FontStyle.Bold, GraphicsUnit.Point, CByte(136))
        Button3_Stop.Location = New Point(535, 90)
        Button3_Stop.Margin = New Padding(4)
        Button3_Stop.Name = "Button3_Stop"
        Button3_Stop.Size = New Size(96, 42)
        Button3_Stop.TabIndex = 33
        Button3_Stop.Text = "STOP"
        Button3_Stop.UseVisualStyleBackColor = True
        Button3_Stop.Visible = False
        ' 
        ' Button3
        ' 
        Button3.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        Button3.FlatStyle = FlatStyle.System
        Button3.Font = New Font("Microsoft JhengHei UI", 10F, FontStyle.Bold, GraphicsUnit.Point, CByte(136))
        Button3.Location = New Point(562, 16)
        Button3.Margin = New Padding(4)
        Button3.Name = "Button3"
        Button3.Size = New Size(96, 42)
        Button3.TabIndex = 31
        Button3.Text = "Go"
        Button3.UseVisualStyleBackColor = True
        ' 
        ' GroupBox3
        ' 
        GroupBox3.Controls.Add(CountMax)
        GroupBox3.Controls.Add(CountMin)
        GroupBox3.Controls.Add(CheckAttCount)
        GroupBox3.Controls.Add(Label3)
        GroupBox3.FlatStyle = FlatStyle.Flat
        GroupBox3.Font = New Font("Microsoft JhengHei UI", 9.5F)
        GroupBox3.Location = New Point(517, 4)
        GroupBox3.Margin = New Padding(4)
        GroupBox3.Name = "GroupBox3"
        GroupBox3.Padding = New Padding(4)
        GroupBox3.Size = New Size(239, 100)
        GroupBox3.TabIndex = 23
        GroupBox3.TabStop = False
        GroupBox3.Text = "GroupBox3"
        GroupBox3.Visible = False
        ' 
        ' CountMax
        ' 
        CountMax.Enabled = False
        CountMax.Location = New Point(103, 60)
        CountMax.Margin = New Padding(4)
        CountMax.Maximum = New Decimal(New Integer() {1000, 0, 0, 0})
        CountMax.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        CountMax.Name = "CountMax"
        CountMax.Size = New Size(71, 28)
        CountMax.TabIndex = 20
        CountMax.TextAlign = HorizontalAlignment.Right
        CountMax.Value = New Decimal(New Integer() {2, 0, 0, 0})
        ' 
        ' CountMin
        ' 
        CountMin.Enabled = False
        CountMin.Location = New Point(8, 60)
        CountMin.Margin = New Padding(4)
        CountMin.Maximum = New Decimal(New Integer() {1000, 0, 0, 0})
        CountMin.Name = "CountMin"
        CountMin.Size = New Size(71, 28)
        CountMin.TabIndex = 18
        CountMin.TextAlign = HorizontalAlignment.Right
        CountMin.Value = New Decimal(New Integer() {1, 0, 0, 0})
        ' 
        ' CheckAttCount
        ' 
        CheckAttCount.AutoSize = True
        CheckAttCount.FlatStyle = FlatStyle.System
        CheckAttCount.Location = New Point(8, 28)
        CheckAttCount.Margin = New Padding(4)
        CheckAttCount.Name = "CheckAttCount"
        CheckAttCount.Size = New Size(104, 25)
        CheckAttCount.TabIndex = 9
        CheckAttCount.Text = "附件個數"
        CheckAttCount.UseVisualStyleBackColor = True
        ' 
        ' Label3
        ' 
        Label3.AutoSize = True
        Label3.Location = New Point(82, 62)
        Label3.Margin = New Padding(4, 0, 4, 0)
        Label3.Name = "Label3"
        Label3.Size = New Size(21, 20)
        Label3.TabIndex = 22
        Label3.Text = "~"
        ' 
        ' ListView3
        ' 
        ListView3.AllowColumnReorder = True
        ListView3.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        ListView3.Columns.AddRange(New ColumnHeader() {ColumnHeader31, ColumnHeader32, ColumnHeader34, ColumnHeader33, ColumnHeader35, ColumnHeader36})
        ListView3.Font = New Font("Microsoft JhengHei UI", 10F)
        ListView3.FullRowSelect = True
        ListView3.Location = New Point(0, 111)
        ListView3.Margin = New Padding(4)
        ListView3.Name = "ListView3"
        ListView3.Size = New Size(612, 1133)
        ListView3.TabIndex = 8
        ListView3.UseCompatibleStateImageBehavior = False
        ListView3.View = View.Details
        ' 
        ' ColumnHeader31
        ' 
        ColumnHeader31.Text = "郵件主旨"
        ColumnHeader31.Width = 200
        ' 
        ' ColumnHeader32
        ' 
        ColumnHeader32.Text = "郵件大小"
        ColumnHeader32.TextAlign = HorizontalAlignment.Right
        ColumnHeader32.Width = 80
        ' 
        ' ColumnHeader34
        ' 
        ColumnHeader34.Text = "收到日期"
        ColumnHeader34.TextAlign = HorizontalAlignment.Center
        ColumnHeader34.Width = 90
        ' 
        ' ColumnHeader33
        ' 
        ColumnHeader33.Text = "寄件者"
        ColumnHeader33.Width = 85
        ' 
        ' ColumnHeader35
        ' 
        ColumnHeader35.Text = "附件個數"
        ColumnHeader35.TextAlign = HorizontalAlignment.Center
        ColumnHeader35.Width = 65
        ' 
        ' ColumnHeader36
        ' 
        ColumnHeader36.Text = "EntryID"
        ColumnHeader36.Width = 100
        ' 
        ' GroupBox1
        ' 
        GroupBox1.Controls.Add(TextBox3)
        GroupBox1.Controls.Add(CheckAttachName)
        GroupBox1.FlatStyle = FlatStyle.Flat
        GroupBox1.Font = New Font("Microsoft JhengHei UI", 9.5F)
        GroupBox1.Location = New Point(4, 4)
        GroupBox1.Margin = New Padding(4)
        GroupBox1.Name = "GroupBox1"
        GroupBox1.Padding = New Padding(4)
        GroupBox1.Size = New Size(167, 100)
        GroupBox1.TabIndex = 18
        GroupBox1.TabStop = False
        GroupBox1.Text = "GroupBox1"
        ' 
        ' TextBox3
        ' 
        TextBox3.Location = New Point(8, 60)
        TextBox3.Margin = New Padding(4)
        TextBox3.Name = "TextBox3"
        TextBox3.Size = New Size(151, 28)
        TextBox3.TabIndex = 13
        TextBox3.Text = "pdf"
        ' 
        ' CheckAttachName
        ' 
        CheckAttachName.AutoSize = True
        CheckAttachName.FlatStyle = FlatStyle.System
        CheckAttachName.Font = New Font("Microsoft JhengHei UI", 9.5F, FontStyle.Regular, GraphicsUnit.Point, CByte(136))
        CheckAttachName.Location = New Point(8, 28)
        CheckAttachName.Margin = New Padding(4)
        CheckAttachName.Name = "CheckAttachName"
        CheckAttachName.Size = New Size(150, 25)
        CheckAttachName.TabIndex = 12
        CheckAttachName.Text = "附件名稱 (最慢)"
        CheckAttachName.UseVisualStyleBackColor = True
        ' 
        ' GroupBox2
        ' 
        GroupBox2.Controls.Add(UnitMax)
        GroupBox2.Controls.Add(UnitMin)
        GroupBox2.Controls.Add(NumberMax)
        GroupBox2.Controls.Add(NumberMin)
        GroupBox2.Controls.Add(CheckSize)
        GroupBox2.Controls.Add(Label1)
        GroupBox2.FlatStyle = FlatStyle.Flat
        GroupBox2.Font = New Font("Microsoft JhengHei UI", 9.5F)
        GroupBox2.Location = New Point(179, 4)
        GroupBox2.Margin = New Padding(4)
        GroupBox2.Name = "GroupBox2"
        GroupBox2.Padding = New Padding(4)
        GroupBox2.Size = New Size(330, 100)
        GroupBox2.TabIndex = 19
        GroupBox2.TabStop = False
        GroupBox2.Text = "GroupBox2"
        ' 
        ' UnitMax
        ' 
        UnitMax.FlatStyle = FlatStyle.System
        UnitMax.FormattingEnabled = True
        UnitMax.Items.AddRange(New Object() {"KB", "MB", "GB"})
        UnitMax.Location = New Point(244, 60)
        UnitMax.Margin = New Padding(4)
        UnitMax.Name = "UnitMax"
        UnitMax.Size = New Size(57, 28)
        UnitMax.TabIndex = 21
        UnitMax.Text = "MB"
        ' 
        ' UnitMin
        ' 
        UnitMin.FlatStyle = FlatStyle.System
        UnitMin.FormattingEnabled = True
        UnitMin.Items.AddRange(New Object() {"KB", "MB", "GB"})
        UnitMin.Location = New Point(86, 60)
        UnitMin.Margin = New Padding(4)
        UnitMin.Name = "UnitMin"
        UnitMin.Size = New Size(57, 28)
        UnitMin.TabIndex = 19
        UnitMin.Text = "KB"
        ' 
        ' NumberMax
        ' 
        NumberMax.Location = New Point(166, 60)
        NumberMax.Margin = New Padding(4)
        NumberMax.Maximum = New Decimal(New Integer() {1000, 0, 0, 0})
        NumberMax.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        NumberMax.Name = "NumberMax"
        NumberMax.Size = New Size(71, 28)
        NumberMax.TabIndex = 20
        NumberMax.TextAlign = HorizontalAlignment.Right
        NumberMax.Value = New Decimal(New Integer() {10, 0, 0, 0})
        ' 
        ' NumberMin
        ' 
        NumberMin.Location = New Point(8, 60)
        NumberMin.Margin = New Padding(4)
        NumberMin.Maximum = New Decimal(New Integer() {1000, 0, 0, 0})
        NumberMin.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        NumberMin.Name = "NumberMin"
        NumberMin.Size = New Size(71, 28)
        NumberMin.TabIndex = 18
        NumberMin.TextAlign = HorizontalAlignment.Right
        NumberMin.Value = New Decimal(New Integer() {200, 0, 0, 0})
        ' 
        ' CheckSize
        ' 
        CheckSize.AutoSize = True
        CheckSize.Checked = True
        CheckSize.CheckState = CheckState.Checked
        CheckSize.FlatStyle = FlatStyle.System
        CheckSize.Location = New Point(8, 28)
        CheckSize.Margin = New Padding(4)
        CheckSize.Name = "CheckSize"
        CheckSize.Size = New Size(104, 25)
        CheckSize.TabIndex = 9
        CheckSize.Text = "附件大小"
        CheckSize.UseVisualStyleBackColor = True
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.Location = New Point(145, 62)
        Label1.Margin = New Padding(4, 0, 4, 0)
        Label1.Name = "Label1"
        Label1.Size = New Size(21, 20)
        Label1.TabIndex = 22
        Label1.Text = "~"
        ' 
        ' TabPage4
        ' 
        TabPage4.Controls.Add(SplitContainer4)
        TabPage4.Location = New Point(4, 37)
        TabPage4.Margin = New Padding(4)
        TabPage4.Name = "TabPage4"
        TabPage4.Padding = New Padding(4)
        TabPage4.Size = New Size(998, 1278)
        TabPage4.TabIndex = 3
        TabPage4.Text = "尋找系列郵件"
        TabPage4.UseVisualStyleBackColor = True
        ' 
        ' SplitContainer4
        ' 
        SplitContainer4.Dock = DockStyle.Fill
        SplitContainer4.Location = New Point(4, 4)
        SplitContainer4.Margin = New Padding(4)
        SplitContainer4.Name = "SplitContainer4"
        ' 
        ' SplitContainer4.Panel1
        ' 
        SplitContainer4.Panel1.Controls.Add(TreeView4)
        ' 
        ' SplitContainer4.Panel2
        ' 
        SplitContainer4.Panel2.Controls.Add(ListView4)
        SplitContainer4.Panel2.Controls.Add(Button4)
        SplitContainer4.Size = New Size(990, 1270)
        SplitContainer4.SplitterDistance = 301
        SplitContainer4.SplitterWidth = 5
        SplitContainer4.TabIndex = 11
        ' 
        ' TreeView4
        ' 
        TreeView4.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        TreeView4.Font = New Font("Microsoft JhengHei UI", 10F)
        TreeView4.HideSelection = False
        TreeView4.Location = New Point(0, 0)
        TreeView4.Margin = New Padding(4)
        TreeView4.Name = "TreeView4"
        TreeView4.Size = New Size(300, 1244)
        TreeView4.TabIndex = 2
        ' 
        ' ListView4
        ' 
        ListView4.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        ListView4.Columns.AddRange(New ColumnHeader() {ColumnHeader9, ColumnHeader10})
        ListView4.Font = New Font("Microsoft JhengHei UI", 10F)
        ListView4.FullRowSelect = True
        ListView4.Location = New Point(0, 111)
        ListView4.Margin = New Padding(4)
        ListView4.Name = "ListView4"
        ListView4.Size = New Size(612, 1133)
        ListView4.TabIndex = 4
        ListView4.UseCompatibleStateImageBehavior = False
        ListView4.View = View.Details
        ' 
        ' ColumnHeader9
        ' 
        ColumnHeader9.Text = "郵件主旨"
        ColumnHeader9.Width = 100
        ' 
        ' ColumnHeader10
        ' 
        ColumnHeader10.Text = "重複數量"
        ColumnHeader10.TextAlign = HorizontalAlignment.Right
        ColumnHeader10.Width = 100
        ' 
        ' Button4
        ' 
        Button4.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        Button4.Font = New Font("Microsoft JhengHei UI", 10F, FontStyle.Bold, GraphicsUnit.Point, CByte(136))
        Button4.Location = New Point(516, 15)
        Button4.Margin = New Padding(4)
        Button4.Name = "Button4"
        Button4.Size = New Size(96, 42)
        Button4.TabIndex = 3
        Button4.Text = "Go"
        Button4.UseVisualStyleBackColor = True
        ' 
        ' TabPage5
        ' 
        TabPage5.Controls.Add(SplitContainer5)
        TabPage5.Controls.Add(Label2)
        TabPage5.Location = New Point(4, 37)
        TabPage5.Margin = New Padding(4)
        TabPage5.Name = "TabPage5"
        TabPage5.Padding = New Padding(4)
        TabPage5.Size = New Size(998, 1278)
        TabPage5.TabIndex = 4
        TabPage5.Text = "尋找重複郵件"
        TabPage5.UseVisualStyleBackColor = True
        ' 
        ' SplitContainer5
        ' 
        SplitContainer5.Dock = DockStyle.Fill
        SplitContainer5.Location = New Point(4, 4)
        SplitContainer5.Margin = New Padding(4)
        SplitContainer5.Name = "SplitContainer5"
        ' 
        ' SplitContainer5.Panel1
        ' 
        SplitContainer5.Panel1.Controls.Add(TreeView5)
        ' 
        ' SplitContainer5.Panel2
        ' 
        SplitContainer5.Panel2.Controls.Add(lstEmails)
        SplitContainer5.Panel2.Controls.Add(TextBox2)
        SplitContainer5.Panel2.Controls.Add(TextBox1)
        SplitContainer5.Panel2.Controls.Add(Button5)
        SplitContainer5.Panel2Collapsed = True
        SplitContainer5.Size = New Size(990, 1270)
        SplitContainer5.SplitterDistance = 334
        SplitContainer5.SplitterWidth = 5
        SplitContainer5.TabIndex = 4
        ' 
        ' TreeView5
        ' 
        TreeView5.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        TreeView5.Font = New Font("Microsoft JhengHei UI", 10F, FontStyle.Regular, GraphicsUnit.Point, CByte(136))
        TreeView5.Location = New Point(0, 0)
        TreeView5.Margin = New Padding(4)
        TreeView5.Name = "TreeView5"
        TreeView5.Size = New Size(989, 1244)
        TreeView5.TabIndex = 0
        ' 
        ' lstEmails
        ' 
        lstEmails.FormattingEnabled = True
        lstEmails.Location = New Point(4, 275)
        lstEmails.Margin = New Padding(4)
        lstEmails.Name = "lstEmails"
        lstEmails.Size = New Size(745, 598)
        lstEmails.TabIndex = 4
        ' 
        ' TextBox2
        ' 
        TextBox2.Location = New Point(373, 41)
        TextBox2.Margin = New Padding(4)
        TextBox2.Multiline = True
        TextBox2.Name = "TextBox2"
        TextBox2.Size = New Size(215, 341)
        TextBox2.TabIndex = 1
        ' 
        ' TextBox1
        ' 
        TextBox1.Location = New Point(152, 41)
        TextBox1.Margin = New Padding(4)
        TextBox1.Multiline = True
        TextBox1.Name = "TextBox1"
        TextBox1.Size = New Size(212, 341)
        TextBox1.TabIndex = 0
        ' 
        ' Button5
        ' 
        Button5.Location = New Point(651, 4)
        Button5.Margin = New Padding(4)
        Button5.Name = "Button5"
        Button5.Size = New Size(96, 29)
        Button5.TabIndex = 2
        Button5.Text = "Button5"
        Button5.UseVisualStyleBackColor = True
        ' 
        ' Label2
        ' 
        Label2.AutoSize = True
        Label2.Location = New Point(384, 343)
        Label2.Margin = New Padding(4, 0, 4, 0)
        Label2.Name = "Label2"
        Label2.Size = New Size(63, 22)
        Label2.TabIndex = 3
        Label2.Text = "Label2"
        ' 
        ' TabPage6
        ' 
        TabPage6.Controls.Add(LoadCache)
        TabPage6.Controls.Add(SaveCache)
        TabPage6.Controls.Add(checkIncludeAllFolders)
        TabPage6.Controls.Add(CheckRDO)
        TabPage6.Controls.Add(buttonClearCache)
        TabPage6.Controls.Add(CheckDebug)
        TabPage6.Location = New Point(4, 37)
        TabPage6.Margin = New Padding(4)
        TabPage6.Name = "TabPage6"
        TabPage6.Padding = New Padding(4)
        TabPage6.Size = New Size(998, 1278)
        TabPage6.TabIndex = 5
        TabPage6.Text = "Debug"
        TabPage6.UseVisualStyleBackColor = True
        ' 
        ' LoadCache
        ' 
        LoadCache.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        LoadCache.Font = New Font("Microsoft JhengHei UI", 10F, FontStyle.Bold, GraphicsUnit.Point, CByte(136))
        LoadCache.Location = New Point(851, 193)
        LoadCache.Name = "LoadCache"
        LoadCache.Size = New Size(129, 63)
        LoadCache.TabIndex = 12
        LoadCache.Text = "Load Cache from Disk"
        LoadCache.UseVisualStyleBackColor = True
        ' 
        ' SaveCache
        ' 
        SaveCache.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        SaveCache.Font = New Font("Microsoft JhengHei UI", 10F, FontStyle.Bold, GraphicsUnit.Point, CByte(136))
        SaveCache.Location = New Point(851, 107)
        SaveCache.Name = "SaveCache"
        SaveCache.Size = New Size(129, 63)
        SaveCache.TabIndex = 11
        SaveCache.Text = "Save Cache to Disk"
        SaveCache.UseVisualStyleBackColor = True
        ' 
        ' checkIncludeAllFolders
        ' 
        checkIncludeAllFolders.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        checkIncludeAllFolders.Appearance = Appearance.Button
        checkIncludeAllFolders.FlatStyle = FlatStyle.System
        checkIncludeAllFolders.Font = New Font("Microsoft JhengHei UI", 10F, FontStyle.Bold, GraphicsUnit.Point, CByte(136))
        checkIncludeAllFolders.Location = New Point(706, 107)
        checkIncludeAllFolders.Margin = New Padding(4)
        checkIncludeAllFolders.Name = "checkIncludeAllFolders"
        checkIncludeAllFolders.Size = New Size(129, 63)
        checkIncludeAllFolders.TabIndex = 10
        checkIncludeAllFolders.Text = "Include All Folders"
        checkIncludeAllFolders.TextAlign = ContentAlignment.MiddleCenter
        checkIncludeAllFolders.UseVisualStyleBackColor = True
        ' 
        ' CheckRDO
        ' 
        CheckRDO.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        CheckRDO.Appearance = Appearance.Button
        CheckRDO.FlatStyle = FlatStyle.System
        CheckRDO.Font = New Font("Microsoft JhengHei UI", 10F, FontStyle.Bold, GraphicsUnit.Point, CByte(136))
        CheckRDO.Location = New Point(706, 193)
        CheckRDO.Margin = New Padding(4)
        CheckRDO.Name = "CheckRDO"
        CheckRDO.Size = New Size(129, 63)
        CheckRDO.TabIndex = 9
        CheckRDO.Text = "Load Redemption"
        CheckRDO.TextAlign = ContentAlignment.MiddleCenter
        CheckRDO.UseVisualStyleBackColor = True
        ' 
        ' buttonClearCache
        ' 
        buttonClearCache.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        buttonClearCache.Font = New Font("Microsoft JhengHei UI", 10F, FontStyle.Bold, GraphicsUnit.Point, CByte(136))
        buttonClearCache.Location = New Point(851, 20)
        buttonClearCache.Name = "buttonClearCache"
        buttonClearCache.Size = New Size(129, 63)
        buttonClearCache.TabIndex = 6
        buttonClearCache.Text = "Clear Caches Memory"
        buttonClearCache.UseVisualStyleBackColor = True
        ' 
        ' CheckDebug
        ' 
        CheckDebug.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        CheckDebug.Appearance = Appearance.Button
        CheckDebug.FlatStyle = FlatStyle.System
        CheckDebug.Font = New Font("Microsoft JhengHei UI", 10F, FontStyle.Bold, GraphicsUnit.Point, CByte(136))
        CheckDebug.Location = New Point(706, 20)
        CheckDebug.Margin = New Padding(4)
        CheckDebug.Name = "CheckDebug"
        CheckDebug.Size = New Size(129, 63)
        CheckDebug.TabIndex = 5
        CheckDebug.Text = "Debug Window"
        CheckDebug.TextAlign = ContentAlignment.MiddleCenter
        CheckDebug.UseVisualStyleBackColor = True
        ' 
        ' ToolStripStatusLabel1
        ' 
        ToolStripStatusLabel1.AutoSize = False
        ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        ToolStripStatusLabel1.Size = New Size(6, 24)
        ToolStripStatusLabel1.Text = "   "
        ' 
        ' ProgressBar1
        ' 
        ProgressBar1.AutoSize = False
        ProgressBar1.DisplayStyle = ToolStripItemDisplayStyle.Text
        ProgressBar1.ForeColor = Color.DimGray
        ProgressBar1.Name = "ProgressBar1"
        ProgressBar1.Size = New Size(300, 24)
        ProgressBar1.Text = "ProgressBar1"
        ProgressBar1.TextAlign = ContentAlignment.MiddleLeft
        ' 
        ' ProgressBar2
        ' 
        ProgressBar2.AutoSize = False
        ProgressBar2.DisplayStyle = ToolStripItemDisplayStyle.Text
        ProgressBar2.ForeColor = Color.DimGray
        ProgressBar2.Name = "ProgressBar2"
        ProgressBar2.Size = New Size(480, 24)
        ProgressBar2.Text = "ProgressBar2"
        ProgressBar2.TextAlign = ContentAlignment.MiddleLeft
        ' 
        ' ToolStripProgressBar1
        ' 
        ToolStripProgressBar1.Name = "ToolStripProgressBar1"
        ToolStripProgressBar1.Size = New Size(0, 24)
        ' 
        ' StatusStrip1
        ' 
        StatusStrip1.ImageScalingSize = New Size(20, 20)
        StatusStrip1.Items.AddRange(New ToolStripItem() {ToolStripStatusLabel1, ProgressBar1, ProgressBar2, ToolStripProgressBar1})
        StatusStrip1.Location = New Point(0, 1289)
        StatusStrip1.Name = "StatusStrip1"
        StatusStrip1.Padding = New Padding(1, 0, 18, 0)
        StatusStrip1.Size = New Size(1006, 30)
        StatusStrip1.TabIndex = 6
        StatusStrip1.Text = "StatusStrip1"
        ' 
        ' Form1
        ' 
        AutoScaleDimensions = New SizeF(9F, 19F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(1006, 1319)
        Controls.Add(StatusStrip1)
        Controls.Add(TabControl1)
        DoubleBuffered = True
        Icon = CType(resources.GetObject("$this.Icon"), Icon)
        Margin = New Padding(4)
        Name = "Form1"
        StartPosition = FormStartPosition.CenterScreen
        Text = "Outlook Assistant"
        TabControl1.ResumeLayout(False)
        TabPage1.ResumeLayout(False)
        SplitContainer1.Panel1.ResumeLayout(False)
        SplitContainer1.Panel2.ResumeLayout(False)
        CType(SplitContainer1, ComponentModel.ISupportInitialize).EndInit()
        SplitContainer1.ResumeLayout(False)
        TabPage2.ResumeLayout(False)
        SplitContainer2.Panel2.ResumeLayout(False)
        SplitContainer2.Panel2.PerformLayout()
        CType(SplitContainer2, ComponentModel.ISupportInitialize).EndInit()
        SplitContainer2.ResumeLayout(False)
        CType(Chart2, ComponentModel.ISupportInitialize).EndInit()
        TabPage3.ResumeLayout(False)
        SplitContainer3.Panel1.ResumeLayout(False)
        SplitContainer3.Panel2.ResumeLayout(False)
        SplitContainer3.Panel2.PerformLayout()
        CType(SplitContainer3, ComponentModel.ISupportInitialize).EndInit()
        SplitContainer3.ResumeLayout(False)
        GroupBox3.ResumeLayout(False)
        GroupBox3.PerformLayout()
        CType(CountMax, ComponentModel.ISupportInitialize).EndInit()
        CType(CountMin, ComponentModel.ISupportInitialize).EndInit()
        GroupBox1.ResumeLayout(False)
        GroupBox1.PerformLayout()
        GroupBox2.ResumeLayout(False)
        GroupBox2.PerformLayout()
        CType(NumberMax, ComponentModel.ISupportInitialize).EndInit()
        CType(NumberMin, ComponentModel.ISupportInitialize).EndInit()
        TabPage4.ResumeLayout(False)
        SplitContainer4.Panel1.ResumeLayout(False)
        SplitContainer4.Panel2.ResumeLayout(False)
        CType(SplitContainer4, ComponentModel.ISupportInitialize).EndInit()
        SplitContainer4.ResumeLayout(False)
        TabPage5.ResumeLayout(False)
        TabPage5.PerformLayout()
        SplitContainer5.Panel1.ResumeLayout(False)
        SplitContainer5.Panel2.ResumeLayout(False)
        SplitContainer5.Panel2.PerformLayout()
        CType(SplitContainer5, ComponentModel.ISupportInitialize).EndInit()
        SplitContainer5.ResumeLayout(False)
        TabPage6.ResumeLayout(False)
        StatusStrip1.ResumeLayout(False)
        StatusStrip1.PerformLayout()
        ResumeLayout(False)
        PerformLayout()

    End Sub

    Friend WithEvents TabControl1 As TabControl
    Friend WithEvents SplitContainer1 As SplitContainer
    Friend WithEvents SplitContainer2 As SplitContainer
    Friend WithEvents SplitContainer3 As SplitContainer
    Friend WithEvents SplitContainer4 As SplitContainer
    Friend WithEvents SplitContainer5 As SplitContainer
    Friend WithEvents ToolStripStatusLabel1 As ToolStripStatusLabel
    Friend WithEvents ProgressBar1 As ToolStripStatusLabel
    Friend WithEvents ProgressBar2 As ToolStripStatusLabel
    Friend WithEvents ToolStripProgressBar1 As ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As StatusStrip

    'Friend WithEvents tmrPreCache As Timer
    'Friend WithEvents Timer2 As Timer
    'Friend WithEvents Timer3 As Timer
    'Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    'Friend WithEvents BackgroundWorker2 As System.ComponentModel.BackgroundWorker
    'Friend WithEvents BackgroundWorker3 As System.ComponentModel.BackgroundWorker

    Friend WithEvents TabPage1 As TabPage
    Friend WithEvents TreeView1 As TreeView
    Friend WithEvents ListView1 As ListView
    Friend WithEvents ColumnHeader11 As ColumnHeader
    Friend WithEvents ColumnHeader12 As ColumnHeader
    Friend WithEvents ColumnHeader13 As ColumnHeader
    Friend WithEvents ColumnHeader14 As ColumnHeader

    Friend WithEvents TabPage2 As TabPage
    Friend WithEvents Chart2 As DataVisualization.Charting.Chart
    Friend WithEvents ListView2 As ListView
    Friend WithEvents ColumnHeader21 As ColumnHeader
    Friend WithEvents ColumnHeader22 As ColumnHeader
    Friend WithEvents ColumnHeader23 As ColumnHeader
    Friend WithEvents CheckSubFolder2 As CheckBox
    Friend WithEvents TextBox2 As TextBox
    Friend WithEvents Label2 As Label

    Friend WithEvents TabPage3 As TabPage
    Friend WithEvents TreeView3 As TreeView
    Friend WithEvents ListView3 As ListView
    Friend WithEvents ColumnHeader31 As ColumnHeader
    Friend WithEvents ColumnHeader32 As ColumnHeader
    Friend WithEvents ColumnHeader34 As ColumnHeader
    Friend WithEvents ColumnHeader33 As ColumnHeader
    Friend WithEvents ColumnHeader35 As ColumnHeader
    Friend WithEvents ColumnHeader36 As ColumnHeader
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents GroupBox3 As GroupBox
    Friend WithEvents NumberMax As NumericUpDown
    Friend WithEvents NumberMin As NumericUpDown
    Friend WithEvents UnitMax As ComboBox
    Friend WithEvents UnitMin As ComboBox
    Friend WithEvents CountMax As NumericUpDown
    Friend WithEvents CountMin As NumericUpDown
    Friend WithEvents CheckAttCount As CheckBox
    Friend WithEvents CheckAttachName As CheckBox
    Friend WithEvents CheckSubFolder3 As CheckBox
    Friend WithEvents Label3 As Label
    Friend WithEvents Button3 As Button
    Friend WithEvents Button3_Stop As Button

    Friend WithEvents TabPage4 As TabPage
    Friend WithEvents TreeView4 As TreeView
    Friend WithEvents ListView4 As ListView
    Friend WithEvents ColumnHeader9 As ColumnHeader
    Friend WithEvents ColumnHeader10 As ColumnHeader
    Friend WithEvents Button4 As Button
    Friend WithEvents CheckSize As CheckBox

    Friend WithEvents TabPage5 As TabPage
    Friend WithEvents TreeView5 As TreeView
    Friend WithEvents Button5 As Button
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents TextBox3 As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents ColumnHeader15 As ColumnHeader
    Friend WithEvents lstEmails As ListBox

    Friend WithEvents TabPage6 As TabPage
    Friend WithEvents checkIncludeAllFolders As CheckBox
    Friend WithEvents CheckDebug As CheckBox
    Friend WithEvents CheckRDO As CheckBox
    Friend WithEvents buttonClearCache As Button
    Friend WithEvents LoadCache As Button
    Friend WithEvents SaveCache As Button

End Class