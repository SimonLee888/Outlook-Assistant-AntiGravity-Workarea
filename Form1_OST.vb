Imports System.Runtime.InteropServices
Imports System.Text
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Outlook

' ==============================================================
' Form1_OST.vb  — Tab7: OST / PST 解析
' ==============================================================
' 架構三層:
'   Layer1 UI  : 按鈕事件、AfterSelect → 叫 Layer2
'   Layer2 流程: LoadOstToTree / LoadPstToTree / BuildOstFolderTree
'   Layer3 資料: C# ost2pst.FM (讀 OST) / Outlook OOM (讀 PST)
'
' Phase 1 (2026/04/19): OST/PST 目錄樹顯示 (SimTreeOST / SimTreePST)
' Phase 2 (2026/04/22): 點選資料夾 → ListView 顯示郵件清單
'   - OST: LTP. → TC row data → OstMailRow
'   - PST: GetMailInfoOOM (OOM，Form1_Outlook.vb)
' Phase 3 (待實作): CopyFolder / MoveFolder → OOM 寫入目標 PST
'
' 注意:
'   OST 讀取：使用 C# ost2pst library (ost2pst.FM)，只使用其讀取路徑
'   PST 讀取：使用 Outlook OOM AddStore，與 Tab1 相同路徑
'   PST 寫入 (Phase 3)：也使用 OOM，不使用 ost2pst PST writer (已知損壞)
'   FM.StatusMsg delegate：橋接 C# 內部進度訊息到 ProgressBar2
'   2026/04/19 by Claude / Phase 2 2026/04/22 by Claude
'
' 註記 by Gemini 3.0 Flash, 2026/04/23:
'   目前的 UI 佈局實作 (EnsureTab7Phase2UI) 與 Designer 設計不符。
'   原本設計應為：左側放 TreeView (OST/PST)，右側放 ListView (OST/PST)。
'   但目前代碼會動態建立 SplitContainer 把 TreeView 區域切成上下兩半，
'   導致 TreeView 與 ListView 擠在同一側，這部分需要修正。
' ==============================================================

Partial Class Form1

#Region "■ 11 Tab7: OST/PST 解析"

    ' ── Phase 2 模組級欄位 ────────────────────────────────────────────────
    ' 2026/04/22 by Claude
    Private _tab7Initialized As Boolean = False     ' 標記 EnsureTab7Phase2UI 是否已執行過，避免重複初始化
    Private _ostLoaded As Boolean = False           ' 標記 OST 是否已成功開啟，控制 FM.srcFile 的可用性
    Private _tab7StatusSw As New Stopwatch()        ' by Gemini 3.0 Flash, 2026/04/23: 用於計算 CopyFolder 速度與 ETA
    Private _currentOstFilePath As String = ""      ' by Gemini 3.1 Pro, 2026/04/24: 記錄當前 OST 檔案路徑，用於在同目錄建立 Temp PST

    ' ── 排序狀態追蹤 (by Gemini 3 Flash, 2026/04/24) ────────────────────────
    Private _lvOstLastSortColumn As Integer = -1
    Private _lvOstSortOrder As SortOrder = SortOrder.Ascending
    Private _lvPstLastSortColumn As Integer = -1
    Private _lvPstSortOrder As SortOrder = SortOrder.Ascending

    ' OST 郵件列資料結構（純 .NET 值型別，不持有 COM 物件）
    ' 用於從 OST 的 TableContext 中解出關鍵欄位，不佔用 COM 資源
    Private Structure OstMailRow
        Dim Subject As String           ' 郵件主旨
        Dim SizeBytes As Long           ' 郵件大小 (PR_MESSAGE_SIZE，單位 bytes)
        Dim ReceivedTime As DateTime    ' 收到時間 (PidTagMessageDeliveryTime)
        Dim SenderName As String        ' 寄件者名稱
        Dim HasAttachments As Boolean   ' 是否有附件 (PidTagHasAttachments)
        Dim IsRead As Boolean           ' 是否已讀 (PidTagMessageFlags bit 0)
        Dim Nid As UInteger             ' 確實儲存 dwRowID 作為內部 NID 識別
        Dim EntryID As String           ' 郵件的 EntryID (或 NID 轉換)
    End Structure

    ' ── 常用 MAPI 屬性 Tag ID（低 16 bits = Property ID）──────────────────
    ' 用於 OstPropStr / OstPropDT / OstPropI32 搜尋 Property.id 的比對值
    Private Const PROP_SUBJECT As Integer = &H37            ' PidTagSubjectW (Unicode)
    Private Const PROP_DELIVERY_TIME As Integer = &HE06     ' PidTagMessageDeliveryTime (FILETIME)
    Private Const PROP_SENDER_NAME As Integer = &HC1A       ' PidTagSenderName (ANSI)
    Private Const PROP_SENDER_NAME_W As Integer = &HC1B     ' PidTagSenderNameW (Unicode)
    Private Const PROP_MSG_SIZE As Integer = &HE08          ' PidTagMessageSize (PT_LONG)
    Private Const PROP_HAS_ATTACH As Integer = &HE1B        ' PidTagHasAttachments (PT_BOOLEAN)
    Private Const PROP_MSG_FLAGS As Integer = &HE07         ' PidTagMessageFlags (PT_LONG)
    Private Const PROP_ENTRYID As Integer = &HFFF           ' PidTagEntryId (PT_BINARY)

#Region "  ├ Layer1 UI 事件"
    Private Sub InitTab7UI()
        ' ── Phase 2 & 3 UI 初始化 ────────────────────────────────────────────────
        ' by Gemini 3.0 Flash, 2026/04/23
        ' 整合原 EnsureTab7Phase2UI 與 InitTab7Layout，確保 UI 佈局符合 Designer 原始設計。
        _dbg("InitTab7UI", $"初始化開始 (旗標={_tab7Initialized})")
        If _tab7Initialized Then Return
        _tab7Initialized = True

        ' 1. 初始化 ListView 欄位與外觀
        For Each lv In {LvOST, LvPST}
            lv.View = System.Windows.Forms.View.Details
            lv.FullRowSelect = True
            lv.GridLines = False
            lv.Font = New Font("Microsoft JhengHei UI", 10.0F)
            ' 開啟雙緩衝
            GetType(ListView).GetProperty("DoubleBuffered", Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic)?.SetValue(lv, True, Nothing)

            lv.Columns.Clear()
            ' by Gemini 3 Flash, 2026/04/23: 參考 Tab4 修改欄位順序與格式
            lv.Columns.Add("主旨", 300)
            lv.Columns.Add("郵件大小", 100, HorizontalAlignment.Right)
            lv.Columns.Add("收到日期", 100, HorizontalAlignment.Center)
            lv.Columns.Add("寄件者", 140)
            lv.Columns.Add("EntryID", 0) ' 隱藏欄位，對齊 Tab4
        Next

        ' 2. 綁定事件 (2026/07/05 by Simon/Claude: OST/PST 改掛共用處理器，函式內依 sender 分派)
        For Each lv In {LvOST, LvPST}
            AddHandler lv.DoubleClick, AddressOf HandleLvOstPst_DoubleClick
            AddHandler lv.ColumnClick, AddressOf HandleLvOstPst_ColumnClick
            AddHandler lv.KeyDown, AddressOf HandleLvOstPst_KeyDown
        Next

        ' 3. 綁定縮放事件 (實現上下各半)
        Dim parentTab = SimTreeOST.Parent
        If parentTab IsNot Nothing Then AddHandler parentTab.Resize, AddressOf AdjustTab7Layout

        AdjustTab7Layout(Nothing, Nothing) ' 立即執行一次初始對齊
        _dbg("InitTab7UI", "初始化完成")

    End Sub
    Private Sub AdjustTab7Layout(sender As Object, e As EventArgs)

        ' by Gemini 3.0 Flash, 2026/04/23: 手動計算上下各半的比例，參考 Tab3/Tab4 美學間距
        Dim parent = SimTreeOST.Parent
        If parent Is Nothing Then Return

        parent.SuspendLayout()
        Try
            ' ── 佈局參數 ──
            Dim margin As Integer = 4           ' 離 TabPage 邊緣的距離
            Dim spacingH As Integer = 4         ' 左右控制項之間的間距
            Dim spacingV As Integer = 4         ' 上下控制項之間的間距
            Dim btnAreaWidth As Integer = 92    ' 右側按鈕預留寬度

            ' 計算基礎數值
            Dim clientW = parent.ClientSize.Width
            Dim clientH = parent.ClientSize.Height

            ' 可用寬高
            Dim totalWidth As Integer = clientW - btnAreaWidth - (margin * 2) - spacingH
            Dim totalHeight As Integer = clientH - (margin * 2) - spacingV

            ' 上下高度各半
            Dim halfHeight As Integer = (totalHeight \ 2)
            Dim ostY As Integer = margin
            Dim pstY As Integer = margin + halfHeight + spacingV

            ' 左右比例 (TreeView 佔 28%, ListView 佔剩餘)
            Dim treeWidth As Integer = CInt(totalWidth * 0.28)
            Dim lvWidth As Integer = totalWidth - treeWidth
            Dim startX As Integer = margin
            Dim midX As Integer = startX + treeWidth + spacingH

            ' --- 上層 (OST) ---
            SimTreeOST.Anchor = AnchorStyles.None
            LvOST.Anchor = AnchorStyles.None
            SimTreeOST.SetBounds(startX, ostY, treeWidth, halfHeight)
            LvOST.SetBounds(midX, ostY, lvWidth, halfHeight)

            ' --- 下層 (PST) ---
            SimTreePST.Anchor = AnchorStyles.None
            LvPST.Anchor = AnchorStyles.None
            SimTreePST.SetBounds(startX, pstY, treeWidth, halfHeight)
            LvPST.SetBounds(midX, pstY, lvWidth, halfHeight)

            ' --- 右側按鈕對齊 (可選) ---
            Dim btnX As Integer = clientW - btnAreaWidth
            For Each btn In {LoadOST, LoadPST, CopyFolder}
                btn.Anchor = AnchorStyles.Top Or AnchorStyles.Right
                btn.Left = btnX
            Next

            ' --- 恢復 Anchor 以支援視窗極速縮放時的補間 ---
            SimTreeOST.Anchor = AnchorStyles.Top Or AnchorStyles.Left
            LvOST.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
            SimTreePST.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left
            LvPST.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right

        Catch ex As System.Exception
            _dbg("AdjustTab7Layout 錯誤", ex.Message)
        Finally
            parent.ResumeLayout()
        End Try
    End Sub

    Private Sub LoadOST_Click(sender As Object, e As EventArgs) Handles LoadOST.Click

        ' 彈出 FileDialog 選擇 OST 檔，再呼叫 Layer2 解析
        InitTab7UI() ' by Gemini 3.0 Flash, 2026/04/23: 整合後的 UI 初始化

        Using dlg As New OpenFileDialog() With {
            .Title = "選擇要解析的 OST 檔案",
            .Filter = "OST 檔案 (*.ost)|*.ost|所有檔案 (*.*)|*.*",
            .InitialDirectory = My.Application.Info.DirectoryPath}
            If dlg.ShowDialog() <> DialogResult.OK Then Return

            'dlg.FileName = "D:\Users\Simon\Dropbox\私人文件\Visual Studio\Visual Studio 18 (2026)\Outlook Assistant - (AntiGravity測試區)" &
            '               "\bin\Debug\net10.0-windows10.0.17763.0\Inbox_2011_GLI_OST.ost"
            'dlg.FileName = "F:\Inbox_2011_GLI_OST.ost"

            LoadOstToTree(dlg.FileName, SimTreeOST)
        End Using

    End Sub
    Private Sub LoadPST_Click(sender As Object, e As EventArgs) Handles LoadPST.Click

        ' 彈出 FileDialog 選擇 PST 檔，再呼叫 Layer2 以 OOM 載入
        InitTab7UI() ' by Gemini 3.0 Flash, 2026/04/23

        Using dlg As New OpenFileDialog() With {
            .Title = "選擇要載入的 PST 檔案",
            .Filter = "PST 檔案 (*.pst)|*.pst|所有檔案 (*.*)|*.*",
            .InitialDirectory = My.Application.Info.DirectoryPath}
            If dlg.ShowDialog() <> DialogResult.OK Then Return

            'dlg.FileName = "D:\Users\Simon\Dropbox\私人文件\Visual Studio\Visual Studio 18 (2026)\Outlook Assistant - (AntiGravity測試區)" &
            '               "\bin\Debug\net10.0-windows10.0.17763.0\New PST for Test.pst"
            'dlg.FileName = "F:\New PST for Test.pst"

            LoadPstToTree(dlg.FileName, SimTreePST)
        End Using

    End Sub

    Private Async Sub SimTreeOST_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles SimTreeOST.AfterSelect
        ' Phase 2: 選取 OST 資料夾 → 讀取 Contents Table → 顯示郵件清單
        ' 運作路徑: AfterSelect → ParseOstContentsL3 (透過 Task.Run 執行純檔案 I/O)
        '           → 解析內容表格 (TC row data) → 轉換為 OstMailRow → 最後回 UI 執行緒 ShowLvOstItems
        ' 2026/04/22 by Claude
        ' 2026/04/23 by Gemini 3.0 Flash
        If e.Node Is Nothing Then Return
        _dbg("AfterSelect 偵測 Tag", $"Node={e.Node.Text}, TagType={If(e.Node.Tag Is Nothing, "Nothing", e.Node.Tag.GetType().FullName)}")

        Dim ostFolder = TryCast(e.Node.Tag, ost2pst.Folder)
        If ostFolder Is Nothing Then
            _dbg("    ⚠️ Tag 轉換為 ost2pst.Folder 失敗！") : Return
        End If
        _dbg("AfterSelect 觸發", $"資料夾: {ostFolder.name}, _ostLoaded: {_ostLoaded}, srcFile: {If(ost2pst.FM.srcFile Is Nothing, "Nothing", "Open")}")

        ' OST 尚未載入（或載入失敗）時只更新狀態列，不讀內容
        If Not _ostLoaded OrElse ost2pst.FM.srcFile Is Nothing Then
            PgrsBar2.Text = $"OST 資料夾: {ostFolder.path}" : Return
        End If

        _dbg("開始", ostFolder.name)
        LvOST.Items.Clear()
        PgrsBar1.Text = "正在讀取郵件清單..." : PgrsBar2.Text = ostFolder.path
        Cursor = Cursors.WaitCursor

        Try
            ' ost2pst 讀取 OST 是純粹的 File I/O（不透過 Outlook COM），因此可以安全放在 Task.Run 中平行執行
            ' 注意：FM.srcFile.stream 是 FileStream，雖然在 Task 中執行，但必須確保同時間只有一個讀取動作。
            ' Dim items As List(Of OstMailRow) = Await Task.Run(Function() ParseOstContentsL3(sourceFolderOST))

            _dbg("準備呼叫 ParseOstContentsL3 (同步模式)")
            Dim items As List(Of OstMailRow) = ParseOstContentsL3(ostFolder)
            _dbg("同步讀取結束", $"取得 {items.Count} 筆")

            ShowLvOstItems(items)
            PgrsBar1.Text = $"共 {items.Count:N0} 封 — {ostFolder.name}"

        Catch ex As System.Exception
            _dbg("錯誤", ex.Message)
            PgrsBar1.Text = "讀取失敗: " & ex.Message
        Finally
            Cursor = Cursors.Default
            _dbg("結束", ostFolder.name)
        End Try

    End Sub
    Private Async Sub SimTreePST_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles SimTreePST.AfterSelect
        ' Phase 2: 選取 PST 資料夾 → 使用 OOM GetTable → 顯示郵件清單
        '   OOM 呼叫必須在 UI 執行緒（STA）；使用 Form1_Outlook.vb 的GetFolderMailInfosL3，它已有快取機制與 cToken 支援。
        ' 2026/04/22 by Claude

        Dim folder = TryCast(e.Node?.Tag, Outlook.Folder)
        If folder Is Nothing Then Return

        _dbg("開始", folder.Name)
        LvPST.Items.Clear()
        PgrsBar1.Text = "正在讀取 PST 郵件清單..." : PgrsBar2.Text = SafeGetPath(folder)
        Cursor = Cursors.WaitCursor

        Try
            Dim cToken As System.Threading.CancellationToken = OkayNowYouHaveToken()
            ' needTopic:=False：Tab7 不需要 Conversation Topic，省去讀 PR_CONVERSATION_TOPIC 開銷
            Dim rows = Await GetMailInfo(folder, needTopic:=False, cToken:=cToken)
            ShowLvPstItems(rows.Select(Function(r) r.Mail).ToList())
            PgrsBar1.Text = $"共 {rows.Count:N0} 封 — {folder.Name}"

        Catch ex As OperationCanceledException
            _dbg("中斷", "ESC") : PgrsBar1.Text = "由使用者中斷"
        Catch ex As System.Exception
            _dbg("錯誤", ex.Message) : PgrsBar1.Text = "讀取失敗: " & ex.Message
        Finally
            Cursor = Cursors.Default
            _dbg("結束", folder.Name)
        End Try

    End Sub
    Private Sub ShowLvOstItems(items As List(Of OstMailRow))
        ' 把 OstMailRow 清單渲染到 _newLvOST；未讀郵件以粗體顯示

        LvOST.BeginUpdate()
        LvOST.Items.Clear()

        If items IsNot Nothing AndAlso items.Count > 0 Then
            ' by Gemini 3 Flash, 2026/04/23: 改用 AddRange 並對齊 Tab4 欄位順序與格式
            Dim lvItems As New List(Of ListViewItem)
            For Each item In items
                ' 欄位順序: 主旨, 郵件大小(Bytes), 收到日期(yyyy/MM/dd), 寄件者, EntryID(NID)
                Dim lvi As New ListViewItem(item.Subject)
                lvi.SubItems.Add(item.SizeBytes.ToString("N0"))
                lvi.SubItems.Add(If(item.ReceivedTime > DateTime.MinValue, item.ReceivedTime.ToString("yyyy/MM/dd"), ""))
                lvi.SubItems.Add(item.SenderName)
                lvi.SubItems.Add(item.Nid.ToString())

                If Not item.IsRead Then lvi.Font = New Font(LvOST.Font, FontStyle.Bold)
                lvi.Tag = item ' 存入 Tag 以便開啟
                lvItems.Add(lvi)
            Next
            LvOST.Items.AddRange(lvItems.ToArray())
        End If

        LvOST.EndUpdate()

    End Sub
    Private Sub ShowLvPstItems(mails As List(Of MailItemInfo))
        ' 把 MailItemInfo 清單（來自 GetMailInfoOOM）渲染到 _newLvPST

        ' by Gemini 3 Flash, 2026/04/23: 改用 AddRange 並對齊 Tab4 欄位順序與格式
        LvPST.BeginUpdate()
        LvPST.Items.Clear()

        If mails IsNot Nothing AndAlso mails.Count > 0 Then
            Dim lvItems As New List(Of ListViewItem)
            For Each mail In mails
                ' 欄位順序: 主旨, 郵件大小(Bytes), 收到日期(yyyy/MM/dd), 寄件者, EntryID
                Dim lvi As New ListViewItem(mail.Subject)
                lvi.SubItems.Add(mail.Size.ToString("N0"))
                lvi.SubItems.Add(If(mail.RcvTime > DateTime.MinValue, mail.RcvTime.ToString("yyyy/MM/dd"), ""))
                lvi.SubItems.Add(mail.SenderName)
                lvi.SubItems.Add(mail.EntryID)

                lvi.Tag = mail ' 存入 Tag 支援雙擊開啟
                lvItems.Add(lvi)
            Next
            LvPST.Items.AddRange(lvItems.ToArray())
        End If

        LvPST.EndUpdate()

    End Sub
    Private Async Sub CopyFolder_Click(sender As Object, e As EventArgs) Handles CopyFolder.Click

        ' ===========================================================================================
        ' 設計意圖: OST 來源資料夾 (SimTreeOST.SelectedNode) → OOM 目標資料夾 (SimTreePST.SelectedNode)
        '            用 OOM MailItem.Copy() 逐封複製，不使用 ost2pst PST writer
        ' MessageBox.Show("Copy Folder 功能待實作 (Phase 3)" & vbCrLf & vbCrLf &
        '                 "計畫：從 SimTreeOST 選取的 OST 資料夾，" & vbCrLf &
        '                 "複製郵件到 SimTreePST 選取的 Outlook 目標資料夾。",
        '                 "Phase 3 預留", MessageBoxButtons.OK, MessageBoxIcon.Information)
        ' ===========================================================================================

        Dim nodeOST = SimTreeOST.SelectedNode
        Dim nodePST = SimTreePST.SelectedNode
        If nodeOST Is Nothing OrElse nodePST Is Nothing Then
            MessageBox.Show("請先在左側選取來源 OST 資料夾，並在下方選取目標 PST 資料夾。", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information) : Return
        End If

        Dim sourceFolderOST = TryCast(nodeOST.Tag, ost2pst.Folder)
        Dim targetFolderPST = TryCast(nodePST.Tag, Outlook.Folder)
        If sourceFolderOST Is Nothing OrElse targetFolderPST Is Nothing Then
            MessageBox.Show("資料夾選擇無效。", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error) : Return
        End If

        ' by Gemini 3.1 Pro, 2026/04/24: 防呆檢查目標 PST 是否已有同名資料夾，避免 MAPI COM Exception
        Try
            ' 優化第六點：提取 Folders 集合以利釋放 (by Gemini 3 Flash, 2026/05/05)
            Dim subFolders As Outlook.Folders = targetFolderPST.Folders
            Try
                For Each f As Outlook.Folder In subFolders
                    If String.Compare(f.Name, sourceFolderOST.name, StringComparison.OrdinalIgnoreCase) = 0 Then
                        MessageBox.Show($"目標 PST 資料夾 [{targetFolderPST.Name}] 已經存在名為 [{sourceFolderOST.name}] 的子資料夾！" & vbCrLf &
                                        "為避免 Outlook 複製衝突，請先在目標端刪除或重新命名該資料夾。", "同名資料夾衝突", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Return
                    End If
                Next
            Finally
                TryMarshalRelease(subFolders)
            End Try
        Catch ex As System.Exception
            _dbg("檢查同名資料夾", ex.Message)
        End Try

        Dim dr = MessageBox.Show($"確定要複製 OST 資料夾 [{sourceFolderOST.name}] 到 PST 資料夾 [{targetFolderPST.Name}] 嗎？" & vbCrLf &
                                 "這可能需要一點時間。", "確認複製", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If dr <> DialogResult.Yes Then Return

        PgrsBar1.Text = "正在背景匯出 OST 資料夾至暫存 PST..." : PgrsBar2.Text = sourceFolderOST.name
        Cursor = Cursors.WaitCursor : CopyFolder.Enabled = False

        ' by Gemini 3.0 Flash, 2026/04/23: 啟動計時並設定進度增強委派 (Regex 解析核心庫訊息)
        _tab7StatusSw.Restart()
        Dim lastUiUpdateTimeMs As Long = 0 ' by Gemini 3.1 Pro: 效能極限優化 (節流閥)
        ost2pst.FM.StatusMsg = Sub(msg As String)
                                   If String.IsNullOrEmpty(msg) Then Return

                                   Dim currentMs = _tab7StatusSw.ElapsedMilliseconds

                                   ' 匹配核心庫格式: "Converting OST NBT entry 4500 out of 7613"
                                   Dim m = System.Text.RegularExpressions.Regex.Match(msg, "(\d+)\s+out\s+of\s+(\d+)")
                                   If m.Success Then
                                       Dim cur = Long.Parse(m.Groups(1).Value)
                                       Dim total = Long.Parse(m.Groups(2).Value)
                                       Dim elapsedSec = _tab7StatusSw.Elapsed.TotalSeconds

                                       ' 節流機制：如果不是最後一筆，且距離上次更新 UI 不到 250ms，則跳過更新。
                                       ' 背景執行緒每秒 500+ 次更新 WinForms 屬性會造成嚴重的 Lock 競爭和跨執行緒開銷，這是卡速主因！
                                       If cur < total AndAlso (currentMs - lastUiUpdateTimeMs) < 100 Then Return

                                       lastUiUpdateTimeMs = currentMs

                                       If elapsedSec > 0.1 Then
                                           Dim speed = cur / elapsedSec
                                           Dim etaString = ""
                                           If total > 50 AndAlso speed > 0 Then
                                               Dim remSec = CInt(Math.Max(0, (total - cur) / speed))
                                               If remSec > 2 Then etaString = $" (剩餘 {remSec \ 60:D2}:{remSec Mod 60:D2})"
                                           End If
                                           ' 使用 BeginInvoke 確保 UI 安全，避免背景線程與 UI 線程爭奪 Handle 導致鎖死
                                           Dim finalMsg = $"{msg} ({speed:F0} 筆/秒{etaString})"
                                           Me.BeginInvoke(Sub() PgrsBar2.Text = finalMsg)
                                           Return
                                       End If
                                   Else
                                       ' 非進度數字的訊息，也限制至少 100ms 更新一次避免洗頻
                                       If (currentMs - lastUiUpdateTimeMs) < 100 Then Return
                                       lastUiUpdateTimeMs = currentMs
                                   End If
                                   Dim statusMsg = msg
                                   Me.BeginInvoke(Sub() PgrsBar2.Text = statusMsg)
                               End Sub

        ' by Gemini 3.1 Pro, 2026/04/24: 效能極限優化 - 將 Temp PST 建立在與來源 OST 同一個目錄底下。
        ' 這樣如果使用者將 OST 放進 RAM Disk，Temp PST 也會寫在 RAM Disk 裡，解除 C 槽 SSD 的 I/O 瓶頸！
        Dim targetDir As String = System.IO.Path.GetTempPath()
        If Not String.IsNullOrEmpty(_currentOstFilePath) AndAlso System.IO.Directory.Exists(System.IO.Path.GetDirectoryName(_currentOstFilePath)) Then
            targetDir = System.IO.Path.GetDirectoryName(_currentOstFilePath)
        End If
        Dim tempPstPath As String = System.IO.Path.Combine(targetDir, "temp_export_" & Guid.NewGuid().ToString("N") & ".pst")
        Dim success As Boolean = False
        Dim errMsg As String = ""

        Try ' 步驟 1: 背景讀出 OST (純 File I/O)
            Await Task.Run(Sub()
                               ' by Gemini 3.0 Flash, 2026/04/23: 
                               ' 1. 備份原始狀態 (NBTs 與 folders)
                               '    備份 folders 是為了優化加速，避免底層 FindIndex 遍歷數千個無效資料夾。
                               Dim originalNBTs As List(Of ost2pst.NBTENTRY) = Nothing
                               Dim originalFolders As List(Of ost2pst.Folder) = Nothing
                               SyncLock ost2pst.FM.srcFile
                                   originalNBTs = New List(Of ost2pst.NBTENTRY)(ost2pst.FM.srcFile.NBTs)
                                   originalFolders = New List(Of ost2pst.Folder)(ost2pst.FM.folders)
                               End SyncLock

                               Try
                                   SyncLock ost2pst.FM.srcFile
                                       ' 2. 加速優化：精簡 folders 清單 (解決速度問題 3)
                                       '    底層 ToBeExported 會對每個節點執行 folders.FindIndex。
                                       '    如果 folders 有 7600 筆，這會非常慢。
                                       '    我們預先標記並只保留「確定要匯出」的資料夾，將搜尋範圍從 7600 降至極低。
                                       ost2pst.FM.MessagesToExportNIDs.Clear()

                                       ' 找出所有需要匯出的資料夾 NID (遞迴收集)
                                       Dim exportNids As New HashSet(Of UInteger)()
                                       Dim collectNids As Action(Of ost2pst.Folder) = Nothing
                                       collectNids = Sub(f)
                                                         If f Is Nothing Then Return
                                                         exportNids.Add(f.nid.dwValue)
                                                         ' 搜尋所有子資料夾
                                                         For Each subF In originalFolders
                                                             If subF.parent Is f AndAlso subF IsNot f Then collectNids(subF)
                                                         Next
                                                     End Sub
                                       collectNids(sourceFolderOST)

                                       ' 暫時替換 FM.folders 為精簡版
                                       ost2pst.FM.folders = originalFolders.Where(Function(f) exportNids.Contains(f.nid.dwValue)).ToList()
                                       _dbg("加速優化", $"folders 已從 {originalFolders.Count} 筆精簡至 {ost2pst.FM.folders.Count} 筆")

                                       If ost2pst.FM.CreatPstFile(tempPstPath) Then
                                           ' 呼叫 Niv2023 的匯出功能
                                           ost2pst.FM.CopySourceDatablocksToPST(sourceFolderOST.nid.dwValue, System.IO.Path.GetFileName(tempPstPath))
                                           ost2pst.FM.exportNBTnodes()
                                           ost2pst.FM.exportBBTnodes()
                                           ost2pst.FM.updateNidHighWaterMarks()
                                           ost2pst.FM.CloseOutputFile()
                                           success = True
                                       Else
                                           errMsg = "無法建立暫存 PST 檔案。"
                                       End If
                                   End SyncLock
                               Catch ex As System.Exception
                                   errMsg = ex.Message
                               Finally
                                   ' 3. 還原原始狀態，確保 UI TreeView 與後續點選正常
                                   SyncLock ost2pst.FM.srcFile
                                       If originalNBTs IsNot Nothing Then
                                           ost2pst.FM.srcFile.NBTs.Clear()
                                           ost2pst.FM.srcFile.NBTs.AddRange(originalNBTs)
                                       End If
                                       If originalFolders IsNot Nothing Then ost2pst.FM.folders = originalFolders
                                   End SyncLock
                               End Try
                           End Sub)

            If Not success Then
                MessageBox.Show("匯出暫存 PST 失敗: " & errMsg, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error) : Return
            End If

            ' 步驟 2: 透過 OOM 掛載 TempPST
            Dim ns As Outlook.NameSpace = _olApp.GetNamespace("MAPI")
            ns.AddStore(tempPstPath)
            PgrsBar1.Text = "正在掛載並複製資料夾至目標 PST..."
            Await Task.Yield() ' 刷新 UI

            ' 找到剛掛載的 Temp PST Store
            Dim tempStore As Outlook.Store = Nothing
            ' 優化第六點：提取 Stores 集合以利釋放 (by Gemini 3 Flash, 2026/05/05)
            Dim allStores As Outlook.Stores = ns.Stores
            Try
                For Each st As Outlook.Store In allStores
                    If st.FilePath IsNot Nothing AndAlso String.Compare(st.FilePath, tempPstPath, StringComparison.OrdinalIgnoreCase) = 0 Then
                        tempStore = st : Exit For
                    End If
                Next
            Finally
                TryMarshalRelease(allStores)
            End Try

            If tempStore Is Nothing Then
                MessageBox.Show("無法掛載暫存 PST 到 Outlook。", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error) : Return
            End If

            ' 步驟 3: 取得匯出的資料夾並複製 (從OST to TempPST)
            Dim tempRootPST As Outlook.Folder = tempStore.GetRootFolder()
            Dim sourceFolderPST As Outlook.Folder = Nothing
            ' 優化第六點：提取 Folders 集合以利釋放 (by Gemini 3 Flash, 2026/05/05)
            Dim rootFolders As Outlook.Folders = tempRootPST.Folders
            Try
                ' Niv2023 通常會把資料夾建立在 "Top of Personal Folders" 下面
                Dim topFolderPST As Outlook.Folder = Nothing
                Try : topFolderPST = rootFolders.Item("Top of Personal Folders") : Catch : End Try

                If topFolderPST IsNot Nothing Then
                    Dim topFolders As Outlook.Folders = topFolderPST.Folders
                    Try : sourceFolderPST = topFolders.Item(sourceFolderOST.name)
                    Finally : TryMarshalRelease(topFolders) : TryMarshalRelease(topFolderPST)
                    End Try
                End If
            Catch ex As System.Exception
                _dbg("尋找資料夾", "Top of Personal Folders 尋找失敗: " & ex.Message)
            End Try

            If sourceFolderPST Is Nothing Then   ' 若找不到，試著在根目錄直接找
                Try : sourceFolderPST = tempRootPST.Folders(sourceFolderOST.name)
                Catch : End Try
            End If

            If sourceFolderPST Is Nothing Then
                ' 暴力搜尋第一層所有的子資料夾
                ' 優化第六點：提取 Folders 集合以利釋放 (by Gemini 3 Flash, 2026/05/05)
                Try
                    For Each f As Outlook.Folder In rootFolders
                        Dim fFolders As Outlook.Folders = f.Folders
                        Try
                            sourceFolderPST = fFolders.Item(sourceFolderOST.name)
                            If sourceFolderPST IsNot Nothing Then Exit For
                        Catch
                        Finally
                            TryMarshalRelease(fFolders)
                        End Try
                    Next
                Finally
                    ' 注意：rootFolders 在此迴圈結束後釋放
                End Try
            End If
            ' 釋放 rootFolders (因前面步驟 3 也用到，故在此統一釋放)
            TryMarshalRelease(rootFolders)

            ' 步驟 4: 開始複製到目標資料夾 (從 TempPST to 目標 PST)
            If sourceFolderPST IsNot Nothing Then
                sourceFolderPST.CopyTo(targetFolderPST)

                ' by Gemini 3.1 Pro, 2026/04/24: 複製成功後，自動刷新目標 PST 的目錄樹，並選回原本的資料夾
                Try
                    Dim targetPstPath As String = targetFolderPST.Store.FilePath
                    Dim targetFolderPath As String = SafeGetPath(targetFolderPST) ' 記住目前的完整路徑
                    If Not String.IsNullOrEmpty(targetPstPath) Then
                        _dbg("自動重新整理目標 PST", targetPstPath)
                        LoadPstToTree(targetPstPath, SimTreePST)

                        ' 重新整理後，自動尋找並選回剛才的資料夾節點
                        ' by Gemini 3.5 Flash, 2026/05/21: 改用 SimTreePST.GetNodeIn 高效尋路引擎，取代舊有的暴力遞迴 FindNodeByPath
                        Dim foundNode = SimTreePST.GetNode(targetFolderPath, searchOnlyExpanded:=False)
                        If foundNode IsNot Nothing Then
                            SimTreePST.SelectedNode = foundNode
                            foundNode.EnsureVisible()
                        End If
                    End If
                Catch ex As System.Exception
                    _dbg("刷新目標 PST 失敗", ex.Message)
                End Try

                PgrsBar1.Text = "複製完成！"
                MessageBox.Show("複製完成！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("在暫存 PST 中找不到剛匯出的資料夾。", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If

            ' 步驟 5: 卸載 Temp PST 並嘗試刪除實體檔案
            ' (Outlook 可能還鎖著它，所以忽略錯誤，依賴系統 Temp 回收)
            ns.RemoveStore(tempRootPST)
            Try : System.IO.File.Delete(tempPstPath)
            Catch : End Try

        Catch ex As System.Exception
            MessageBox.Show("複製過程中發生錯誤: " & ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' by Gemini 3.0 Flash, 2026/04/23: 還原為簡約版進度委派
            ost2pst.FM.StatusMsg = Sub(msg As String)
                                       If Not String.IsNullOrEmpty(msg) Then PgrsBar2.Text = msg
                                   End Sub
            _tab7StatusSw.Stop()

            Cursor = Cursors.Default
            CopyFolder.Enabled = True
            If PgrsBar1.Text.StartsWith("正在") Then PgrsBar1.Text = "操作完成"
            PgrsBar2.Text = ""
        End Try

    End Sub

    ' ── ListView 事件處理 (Double Click / Enter) by Gemini 3.0 Flash, 2026/04/23 ───────────────────────────
    ' 2026/07/05 by Simon/Claude: OST/PST 四個處理器合併為兩個共用版，依 sender 分派 (同 HandleLv3Lv4Lv5 模式)
    Private Sub HandleLvOstPst_DoubleClick(sender As Object, e As EventArgs)
        _dbg("開始", DirectCast(sender, ListView).Name)
        If sender Is LvOST Then OpenSelectedOstMail()
        If sender Is LvPST Then OpenSelectedPstMail()
    End Sub
    Private Sub HandleLvOstPst_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            _dbg("開始", $"Enter 觸發 ({DirectCast(sender, ListView).Name})")
            If sender Is LvOST Then OpenSelectedOstMail()
            If sender Is LvPST Then OpenSelectedPstMail()
            e.Handled = True
        End If
    End Sub
    Private Sub HandleLvOstPst_ColumnClick(sender As Object, e As ColumnClickEventArgs)
        ' by Gemini 3 Flash, 2026/04/24: 處理 Tab7 ListView 的欄位標題點選排序
        ' 2026/07/05 by Simon/Claude: 排序方向切換改用共用 GetNewSortOrder，移除兩份手寫的三元邏輯
        Dim lv = DirectCast(sender, ListView)
        _dbg("開始", $"Column={e.Column}, ListView={lv.Name}")

        ' 判斷是哪個 ListView 並更新其狀態
        Dim order As SortOrder
        If lv Is LvOST Then
            _lvOstSortOrder = GetNewSortOrder(e.Column, _lvOstLastSortColumn, _lvOstSortOrder)
            _lvOstLastSortColumn = e.Column
            order = _lvOstSortOrder
        ElseIf lv Is LvPST Then
            _lvPstSortOrder = GetNewSortOrder(e.Column, _lvPstLastSortColumn, _lvPstSortOrder)
            _lvPstLastSortColumn = e.Column
            order = _lvPstSortOrder
        Else
            Return
        End If

        ' 設定比較器並執行排序
        lv.ListViewItemSorter = New Tab7LviComparer(e.Column, order)
        lv.Sort()
    End Sub
#End Region
#Region "  ├ Layer2 PST/OST 解析讀取載入"
    Private Async Sub LoadOstToTree(filePath As String, tv As SimTree)
        ' ---------------------------------------------------------------
        ' LoadOstToTree — 使用 C# ost2pst.FM 解析 OST 目錄結構
        '
        ' Phase 1 修改 (2026/04/19 by Claude)：
        '   ① FM.StatusMsg 橋接：C# 內部的 FM.StatusMsg(msg) 走到 ProgressBar2
        '   ② FM.OpenSourceFile() → 解析 NBT/BBT B-Tree
        '   ③ FM.GetFolderList()  → 建立 FM.folders (List(Of ost2pst.Folder))
        '   ④ BuildOstFolderTree() → 把平坦清單轉成 TreeView 節點

        ' Phase 2 修改 (2026/04/22 by Claude)：
        '   - 成功後不呼叫 CloseSourceFile()（保持 FM.srcFile.stream 開啟供 AfterSelect 讀郵件）
        '   - 失敗時仍呼叫 CloseSourceFile()（確保 handle 釋放）
        '   - 載入新檔前先關閉舊檔（_ostLoaded 旗標控制）
        '
        ' FM.CloseSourceFile() 在 Finally 確保 OST 檔 handle 釋放
        ' 需在以下情況呼叫：
        '   ① 載入新 OST 前（此處處理）
        '   ② 載入失敗時（Catch 中處理）
        '   ③ 程式關閉時（請在 Form1_FormClosing 追加：
        '      If _ostLoaded Then ost2pst.FM.CloseSourceFile()）
        ' ---------------------------------------------------------------
        _dbg("開始", filePath)

        ' 若已有 OST 檔開啟，先關閉（避免 handle 衝突）
        If _ostLoaded Then
            ost2pst.FM.CloseSourceFile()
            _ostLoaded = False
            _dbg("    ├ 關閉舊 OST")
        End If

        Cursor = Cursors.WaitCursor
        tv.Nodes.Clear()
        LvOST.Items.Clear()
        PgrsBar1.Text = "正在解析 OST..." : PgrsBar2.Text = ""
        Await Task.Yield()  ' 讓 UI 先刷新再開始耗時操作

        Try
            ' FM.StatusMsg delegate 橋接（待 C# DLL 重新編譯後可還原）：
            ost2pst.FM.StatusMsg = Sub(msg As String)
                                       If Not String.IsNullOrEmpty(msg) Then PgrsBar2.Text = msg
                                   End Sub

            ' ① 開啟 OST 檔（C# 端解析 Header + NBT/BBT B-Tree，約 0.5~2 秒）
            ' by Gemini 3.0 Flash, 2026/04/23: 增加密碼重置容錯機制
            ' 重構清理冗餘邏輯與 GoTo by Gemini 3.1 Pro, 2026/04/23
            If Not ost2pst.FM.OpenSourceFile(filePath) Then
                ' 如果開啟失敗，嘗試自動重置密碼再試一次
                _dbg("    ├ 載入失敗，嘗試執行密碼重置...")
                If Not (ResetOstPassword(filePath) AndAlso ost2pst.FM.OpenSourceFile(filePath)) Then
                    MessageBox.Show("無法開啟 OST 檔案，可能是檔案已被鎖定或格式不支援。", "錯誤")
                    Return
                End If
                _dbg("    └ 密碼重置後成功開啟！")
            End If

            ' 2026/04/23 by Gemini 3.1 Pro: 
            ' ost2pst 函式庫的 TableContext.RowData 建構子會強制呼叫 FM.NextUniqueID()，
            ' 若 outFile 未初始化會發生 NullReferenceException。
            ' 因此在讀取模式下，我們必須建立一個暫存的假 PST 來滿足這個依賴。
            Dim dummyPstPath As String = System.IO.Path.GetTempFileName() & ".pst"
            _dbg("    ├ 建立 Dummy PST 以滿足函式庫依賴", dummyPstPath)
            ost2pst.FM.CreatPstFile(dummyPstPath)

            ' ② 取得資料夾清單
            '    FM.GetFolderList() 走遍所有 NORMAL_FOLDER NID，讀取 PidTagDisplayName
            '    結果存在 FM.folders: List(Of ost2pst.Folder)
            '    每個 Folder 含: .name / .path / .parent (物件參考) / .level / .nbtIndex / .nid
            ost2pst.FM.GetFolderList()

            Dim folderList = ost2pst.FM.folders
            If folderList Is Nothing OrElse folderList.Count = 0 Then
                MessageBox.Show("OST 檔案內找不到任何資料夾。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
                ost2pst.FM.CloseSourceFile()
                Return
            End If

            ' ④ 把平坦清單轉成 TreeView（BFS 建樹，父節點一定比子節點先建）
            tv.BeginUpdate()
            BuildOstFolderTree(folderList, tv)
            tv.EndUpdate()
            If tv.Nodes.Count > 0 Then tv.Nodes(0).Expand()

            ' Phase 2: 成功後標記 _ostLoaded=True，不在 Finally 關閉
            _ostLoaded = True
            _currentOstFilePath = filePath ' by Gemini 3.1 Pro, 2026/04/24: 記錄當前 OST 檔案路徑
            PgrsBar1.Text = $"OST 解析完成：共 {folderList.Count} 個資料夾，請點選資料夾查看郵件"
            PgrsBar2.Text = filePath
            _dbg("結束", $"{folderList.Count} 個資料夾")

            ' By Gemini 3.0 Flash: 背景非同步更新資料夾內的郵件數量
            UpdateOstMaliCountsAsync(tv, folderList)

        Catch ex As System.Exception
            MessageBox.Show("解析 OST 時發生錯誤：" & vbCrLf & ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
            _dbg("錯誤", ex.Message)
            ost2pst.FM.CloseSourceFile()    ' 失敗時確保關閉 handle
            _ostLoaded = False
        Finally
            ' Phase 2: 注意不要在 Finally 關閉 OST，保持開啟供 AfterSelect 讀郵件
            'ost2pst.FM.StatusMsg = Sub(msg As String)  ' 清除 callback（待 DLL 更新） ' by Claude 2026/04/22, 用不到, 註解掉
            Cursor = Cursors.Default
        End Try

    End Sub
    Private Sub BuildOstFolderTree(folders As List(Of ost2pst.Folder), tv As SimTree)
        ' ---------------------------------------------------------------
        ' BuildOstFolderTree — 把 FM.folders 平坦清單 (含 .parent 物件指標) 轉成 TreeView
        '
        ' FM.folders 的父子關係：
        '   根節點：f.parent Is f (自己指向自己) 或 f.parent Is Nothing
        '   子節點：f.parent 指向父 Folder 物件
        '
        ' 演算法：BFS 多輪建樹
        '   第一輪：建立所有根節點並記入 nodeMap
        '   後續輪：每輪嘗試把 parent 已在 nodeMap 的子節點掛上去
        '   最多 50 輪（防無限迴圈，OST 通常不超過 15 層）
        '   孤兒節點（parent 始終找不到）：掛到第一個根節點下並記 _dbg 警告
        '
        ' TreeNode.Tag = ost2pst.Folder 物件（Phase 2 AfterSelect 使用 .nid 找 Contents Table）
        ' 2026/04/19 by Claude
        ' ---------------------------------------------------------------
        If folders Is Nothing OrElse folders.Count = 0 Then Return

        ' 用物件參考作 Dictionary key（ost2pst.Folder 可能覆寫 Equals，要用參考比較）
        Dim nodeMap As New Dictionary(Of ost2pst.Folder, TreeNode)(ReferenceEqualityComparer.Instance)
        ' by Gemini 3.1 Pro: 紀錄已被過濾的節點，用於阻斷子樹
        Dim filteredNodes As New HashSet(Of ost2pst.Folder)(ReferenceEqualityComparer.Instance)

        ' ── 第一步：找並建立根節點 ──────────────────────────────────
        ' 根節點條件：parent 指向自己，或 parent Is Nothing
        Dim rootFolders = folders.Where(Function(f) f.parent Is f OrElse f.parent Is Nothing).ToList()

        ' Fallback：找 level = 0 的（部分 OST 格式根節點不自指）
        If rootFolders.Count = 0 Then rootFolders = folders.Where(Function(f) f.level = 0).ToList()

        ' 最後手段：把第一筆當根
        If rootFolders.Count = 0 Then rootFolders = New List(Of ost2pst.Folder) From {folders(0)}

        For Each root In rootFolders
            Dim displayName As String = If(String.IsNullOrEmpty(root.name), "OST Root", root.name)

            ' by Gemini 3.1 Pro: 判斷是否過濾
            If IsFolderToHide(displayName) Then
                filteredNodes.Add(root) : Continue For
            End If

            Dim rootNode As New TreeNode(displayName) With {.Tag = root, .Name = root.name}
            tv.Nodes.Add(rootNode)
            nodeMap(root) = rootNode
        Next

        ' ── 第二步：BFS 多輪把子節點掛上去 ─────────────────────────
        Dim pending = folders.Where(
            Function(f) Not nodeMap.ContainsKey(f) AndAlso Not rootFolders.Contains(f, ReferenceEqualityComparer.Instance)).ToList()

        Dim maxRounds As Integer = 50
        Do While pending.Count > 0 AndAlso maxRounds > 0
            maxRounds -= 1
            Dim stillPending As New List(Of ost2pst.Folder)
            For Each f In pending
                If f.parent IsNot Nothing Then
                    If nodeMap.ContainsKey(f.parent) Then
                        ' 父節點已建立，檢查自己是否要被過濾
                        Dim displayName As String = If(String.IsNullOrEmpty(f.name), "(未命名)", f.name)

                        If IsFolderToHide(displayName) Then
                            filteredNodes.Add(f)
                            _dbg("    ├ 阻斷子樹 (子)", displayName)
                            Continue For
                        End If

                        Dim childNode As New TreeNode(displayName) With {.Tag = f, .Name = f.name}
                        nodeMap(f.parent).Nodes.Add(childNode)
                        nodeMap(f) = childNode
                    ElseIf filteredNodes.Contains(f.parent) Then
                        filteredNodes.Add(f)    ' 父節點已被過濾，子樹阻斷：將自己也標記為過濾，並且不加入 stillPending (直接丟棄)
                    Else
                        stillPending.Add(f)     ' 父節點尚未處理到，或是真正的孤兒，留待下一輪
                    End If
                Else
                    stillPending.Add(f)
                End If
            Next
            ' 若這輪一個都沒能掛上去，代表剩餘的都是孤兒，直接跳出
            If stillPending.Count = pending.Count Then Exit Do
            pending = stillPending
        Loop

        ' ── 第三步：處理孤兒節點（掛到第一個根節點下，記警告）──────
        If tv.Nodes.Count > 0 Then
            For Each orphan In pending
                Dim displayName As String = If(String.IsNullOrEmpty(orphan.name), "(孤兒)", orphan.name)
                ' 對孤兒也進行過濾檢查
                If IsFolderToHide(displayName) Then Continue For

                _dbg("⚠️ OST 孤兒資料夾", $"path={orphan.path} parent={orphan.parent?.name}")
                Dim orphanNode As New TreeNode(displayName) With {.Tag = orphan, .Name = orphan.name}
                tv.Nodes(0).Nodes.Add(orphanNode)
            Next
        End If
    End Sub
    Private Async Sub UpdateOstMaliCountsAsync(tv As TreeView, folders As List(Of ost2pst.Folder))
        ' ---------------------------------------------------------------
        ' UpdateOstMaliCountsAsync — 背景非同步更新 OST TreeView 的郵件數量
        ' 2026/04/23 by Gemini 3.0 Flash
        '
        ' 運作邏輯：
        '   1. 在背景線程遍歷資料夾清單。
        '   2. 鎖定 FM.srcFile 避免同時讀取衝突。
        '   3. 計算內容表 NID (Type 14) 並從 NBT 找出數據偏移。
        '   4. 讀取 TableContext (TC) 的 RowMatrixCount 取得郵件數量。
        '   5. 透過 Invoke 回 UI 線程更新 TreeView 節點文字。
        ' ---------------------------------------------------------------
        Try
            Await Task.Run(Sub()
                               For Each f In folders
                                   Try
                                       Dim count As Integer = 0
                                       SyncLock ost2pst.FM.srcFile
                                           ' 計算內容表 NID (MS-PST: bits[4:0] 為 0x0E(14) 代表 SubMessages/Contents Table)
                                           Dim contentNid As UInteger = (f.nid.dwValue And Not 31UI) Or 14UI
                                           Dim nbtIdx = ost2pst.FM.srcFile.NBTs.FindIndex(Function(n) n.nid.dwValue = contentNid)
                                           If nbtIdx >= 0 Then
                                               Dim nbt = ost2pst.FM.srcFile.NBTs(nbtIdx)
                                               Dim tc = ost2pst.LTP.ReadTCs_and_rowdata(ost2pst.FM.srcFile.stream, nbt)
                                               ' 僅讀取 TC row data 以獲取數量，不解析具體屬性以求最快速度
                                               If tc IsNot Nothing AndAlso tc.tcRowMatrix IsNot Nothing Then count = tc.tcRowMatrix.Count
                                           End If
                                       End SyncLock

                                       ' 若有資料且 UI 控制項仍有效，則更新節點文字
                                       If count > 0 AndAlso tv.IsHandleCreated Then
                                           tv.Invoke(Sub()
                                                         ' 在樹中尋找對應的資料夾節點 (使用 f.name 作為 Key)
                                                         Dim nodes = tv.Nodes.Find(f.name, True)
                                                         For Each n In nodes
                                                             If n.Tag Is f Then n.Text = $"{f.name} ({count})"
                                                         Next
                                                     End Sub)
                                       End If
                                   Catch
                                       ' 忽略單一資料夾讀取失敗，繼續處理下一個
                                   End Try
                               Next
                           End Sub)
        Catch ex As System.Exception
            _dbg("UpdateOstMaliCountsAsync 錯誤", ex.Message)
        End Try
    End Sub

    Private Sub LoadPstToTree(filePath As String, tv As SimTree)
        ' ---------------------------------------------------------------
        ' LoadPstToTree — 使用 Outlook OOM AddStore 載入 PST 目錄結構
        '
        ' 與 Tab1 的 LoadStoreToTreeView 架構相同，差異是：
        '   ① 先檢查 PST 是否已掛入（避免重複 AddStore 報錯）
        '   ② 找到對應 Store 後遞迴展開全部子資料夾（不用 lazy load，Tab7 目的是一次看全）
        '   ③ TreeNode.Tag = Outlook.Folder（Phase 2 AfterSelect / Phase 3 CopyFolder 直接使用）
        '
        ' 注意：ns 在 Finally 釋放，rootF 所屬 Store 仍由 _olNS 管理，不另外 Release
        ' 2026/04/19 by Claude
        ' ---------------------------------------------------------------
        _dbg("開始", filePath)
        Cursor = Cursors.WaitCursor
        tv.Nodes.Clear()
        LvPST.Items.Clear()
        PgrsBar1.Text = "載入 PST..." : PgrsBar2.Text = ""

        Dim ns As Outlook.NameSpace = Nothing
        Try
            ns = _olApp.GetNamespace("MAPI")

            ' ── 檢查是否已掛入（比對 FilePath，避免重複 AddStore）──
            Dim targetStore As Outlook.Store = Nothing
            For Each store As Outlook.Store In ns.Stores
                If store.FilePath IsNot Nothing AndAlso String.Compare(store.FilePath, filePath, StringComparison.OrdinalIgnoreCase) = 0 Then
                    targetStore = store : Exit For
                End If
            Next

            If targetStore Is Nothing Then
                ns.AddStore(filePath)
                ' AddStore 完成後再搜尋一次
                For Each store As Outlook.Store In ns.Stores
                    If store.FilePath IsNot Nothing AndAlso String.Compare(store.FilePath, filePath, StringComparison.OrdinalIgnoreCase) = 0 Then
                        targetStore = store : Exit For
                    End If
                Next
            End If

            If targetStore Is Nothing Then
                MessageBox.Show("加入 PST 後仍找不到對應的 Store。" & vbCrLf &
                                "可能是 PST 格式損毀或版本不支援。", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' ── 建樹（遞迴展開全部資料夾，Tab7 需要一次看全）──────
            Dim rootF As Outlook.Folder = targetStore.GetRootFolder()
            Dim rootNode As New TreeNode(rootF.Name) With {.Tag = rootF}
            tv.Nodes.Add(rootNode)

            tv.BeginUpdate()
            LoadPstSubFoldersRecursive(rootF, rootNode)
            tv.EndUpdate()
            rootNode.Expand()

            Dim totalNodes As Integer = CountAllNodes(tv.Nodes)
            PgrsBar1.Text = $"PST 載入完成：{rootF.Name}，共 {totalNodes} 個資料夾，請點選資料夾查看郵件"
            PgrsBar2.Text = filePath
            _dbg("結束", $"{rootF.Name} | {totalNodes} 個節點")

        Catch ex As System.Exception
            MessageBox.Show("載入 PST 時發生錯誤：" & vbCrLf & ex.Message,
                            "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
            _dbg("錯誤", ex.Message)
        Finally
            TryMarshalRelease(ns)
            Cursor = Cursors.Default
        End Try
    End Sub
    Private Sub LoadPstSubFoldersRecursive(folder As Outlook.Folder, parentNode As TreeNode)
        ' 遞迴載入 PST 所有子資料夾（OOM，UI 執行緒）
        ' Tag = Outlook.Folder，Phase 2 AfterSelect / Phase 3 CopyFolder 可直接用
        Try
            ' 優化第六點：提取 Folders 集合以利釋放 (by Gemini 3 Flash, 2026/05/05)
            Dim subFolders As Outlook.Folders = folder.Folders
            Try
                For Each subF As Outlook.Folder In subFolders
                    ' by Gemini 3.0 Flash, 2026/04/23: 
                    ' 根據使用者要求，PST 不進行過濾，完整顯示。
                    Dim node As New TreeNode(subF.Name) With {.Tag = subF}
                    parentNode.Nodes.Add(node)
                    ' 遞迴檢查子資料夾數量也提取變數
                    Dim subSubs As Outlook.Folders = subF.Folders
                    Try
                        If subSubs.Count > 0 Then LoadPstSubFoldersRecursive(subF, node)
                    Finally
                        TryMarshalRelease(subSubs)
                    End Try
                Next
            Finally
                TryMarshalRelease(subFolders)
            End Try
        Catch ex As System.Exception
            _dbg("LoadPstSubFoldersRecursive 錯誤", ex.Message)
        End Try
    End Sub
    Private Function IsFolderToHide(name As String) As Boolean
        ' ---------------------------------------------------------------
        ' IsFolderToHide — 判斷資料夾是否應該被過濾隱藏
        ' 2026/04/23 by Gemini 3.1 Pro
        '
        ' 過濾清單意圖 (根據使用者提供之圖片進行補全):
        '   - 系統內部資料夾 (如 NON_IPM_SUBTREE, Drizzle 等)
        '   - 以 ~ 開頭的隱藏資料夾 (如 ~MAPISP)
        '   - 根目錄下無用的導航節點 (如 捷徑, 尋找工具, 一般檢視方式 等)
        '   - 注意: 絕對不可過濾 IPM_SUBTREE，這是主信箱。被擋住的公用資料夾會由樹狀阻斷處理。
        ' ---------------------------------------------------------------
        If String.IsNullOrEmpty(name) Then Return False

        ' 1. 精確匹配過濾 (包含圖片中藍色標記的部分)
        Dim filterList As String() = {"NON_IPM_SUBTREE",
                                      "根資料夾 - 公用",
                                      "共用的資料",
                                      "ItemProcSearch",
                                      "SPAM Search Folder 2",
                                      "Conversation Action Settings",
                                      "~MAPISP(Internal)",
                                      "Drizzle",
                                      "Finder", "尋找工具", "捷徑", "檢視", "一般檢視方式"} ' 常見的 Outlook 系統隱藏資料夾

        If filterList.Contains(name, StringComparer.OrdinalIgnoreCase) Then Return True

        '' 2. 特殊字元前綴過濾
        'If name.StartsWith("~") Then Return True

        Return False
    End Function
    Private Function CountAllNodes(nodes As TreeNodeCollection) As Integer
        ' 遞迴計算 TreeView 節點總數（供狀態列顯示用）
        Dim count As Integer = nodes.Count
        For Each n As TreeNode In nodes
            count += CountAllNodes(n.Nodes)
        Next
        Return count
    End Function

    ' ── OST 郵件開啟邏輯 ──────────────────────────────────────────────────
    ' by Gemini 3.0 Flash, 2026/04/23
    ' 處理流程：雙擊 ListView -> 提取 NID -> 匯出極小 PST -> Outlook Display
    Private Sub OpenSelectedOstMail()
        ''' <summary>
        ''' 響應 LvOST 的雙擊或 Enter 事件。
        ''' 從選取的項目中提取 OST 內部 NID，並發起開啟請求。
        ''' </summary>
        _dbg("OpenSelectedOstMail", $"觸發 (選取數={LvOST.SelectedItems.Count})")
        If LvOST.SelectedItems.Count = 0 Then Return

        ' 取得當前選取的 OST 資料夾資訊
        Dim ostNode = SimTreeOST.SelectedNode
        If ostNode Is Nothing Then _dbg("  ❌ 失敗", "未選取 TreeView 節點") : Return
        Dim ostFolder = TryCast(ostNode.Tag, ost2pst.Folder)
        If ostFolder Is Nothing Then _dbg("  ❌ 失敗", "TreeView Tag 非 ost2pst.Folder") : Return

        ' 1. 收集所有選取郵件的 NID (用於過濾匯出)
        Dim nids As New List(Of UInteger)
        For Each item As ListViewItem In LvOST.SelectedItems
            ' by Gemini 3.0 Flash, 2026/04/23: 改用 TryCast 避免 GetType 在特定編譯條件下失效
            If item.Tag IsNot Nothing Then
                Try
                    ' 嘗試直接轉型 (OstMailRow 是值型別，需要小心處理)
                    Dim rowData = DirectCast(item.Tag, OstMailRow)
                    If rowData.Nid <> 0 Then
                        nids.Add(rowData.Nid)
                    End If
                Catch ex As System.Exception
                    _dbg("  ⚠️ Tag 轉換失敗", ex.Message)
                End Try
            End If
        Next

        ' 2. 執行開啟流程
        If nids.Count > 0 Then
            _dbg("  🚀 執行開啟 OST 郵件", $"Count={nids.Count}, 來源資料夾={ostFolder.name}")
            ' 呼叫背景轉檔開啟邏輯
            OpenSelectedOstMailViaTempPST(nids, ostFolder.nid.dwValue, ostFolder.name)
        Else
            _dbg("  ❌ 失敗", "未取得有效 NID (nids.Count=0)")
        End If
    End Sub
    Private Sub OpenSelectedPstMail()
        If LvPST.SelectedItems.Count = 0 Then Return
        _dbg("OpenSelectedPstMail", "觸發")
        ' PST 模式直接呼叫現有的 GetSelectedEntryIDs
        OpenMailByEntryID(GetSelectedEntryIDs(LvPST))
    End Sub
    Private Async Sub OpenSelectedOstMailViaTempPST(nids As List(Of UInteger), ostFolderNid As UInteger, folderName As String)
        PgrsBar1.Text = "正在背景準備開啟 OST 郵件..."
        Cursor = Cursors.WaitCursor

        Dim tempPstPath As String = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "temp_open_" & Guid.NewGuid().ToString("N") & ".pst")
        Dim success As Boolean = False
        Dim errMsg As String = ""

        Try
            ' 步驟 1: 背景匯出 Temp PST (過濾只包含這些特定郵件)
            Await Task.Run(
                Sub()
                    Try
                        SyncLock ost2pst.FM.srcFile
                            ' 設定單一/多封郵件過濾器
                            ost2pst.FM.MessagesToExportNIDs.Clear()
                            ost2pst.FM.MessagesToExportNIDs.AddRange(nids)

                            If ost2pst.FM.CreatPstFile(tempPstPath) Then
                                ' 匯出該資料夾，但因為 ToBeExported 的過濾，只會匯出我們指定的郵件
                                ost2pst.FM.CopySourceDatablocksToPST(ostFolderNid, System.IO.Path.GetFileName(tempPstPath))
                                ost2pst.FM.exportNBTnodes()
                                ost2pst.FM.exportBBTnodes()
                                ost2pst.FM.updateNidHighWaterMarks()
                                ost2pst.FM.CloseOutputFile()
                                success = True
                            Else
                                errMsg = "無法建立暫存 PST 檔案。"
                            End If
                            ' 清除過濾器
                            ost2pst.FM.MessagesToExportNIDs.Clear()
                        End SyncLock
                    Catch ex As System.Exception
                        errMsg = ex.Message
                    End Try
                End Sub)

            If Not success Then
                MessageBox.Show("匯出暫存郵件失敗: " & errMsg, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            PgrsBar1.Text = "正在掛載並開啟郵件..."
            Await Task.Yield()

            Dim ns As Outlook.NameSpace = _olApp.GetNamespace("MAPI")
            ns.AddStore(tempPstPath)

            Dim tempStore As Outlook.Store = Nothing
            ' 優化第六點：提取 Stores 集合以利釋放 (by Gemini 3 Flash, 2026/05/05)
            Dim allStores As Outlook.Stores = ns.Stores
            Try
                For Each store As Outlook.Store In allStores
                    If store.FilePath IsNot Nothing AndAlso String.Compare(store.FilePath, tempPstPath, StringComparison.OrdinalIgnoreCase) = 0 Then
                        tempStore = store
                        Exit For
                    End If
                Next
            Finally
                TryMarshalRelease(allStores)
            End Try

            If tempStore Is Nothing Then
                MessageBox.Show("無法掛載暫存 PST。", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            Dim tempRoot As Outlook.Folder = tempStore.GetRootFolder()

            ' 取得匯出的資料夾
            Dim exportedFolder As Outlook.Folder = Nothing
            Try
                Dim topFolder = tempRoot.Folders("Top of Personal Folders")
                exportedFolder = topFolder.Folders(folderName)
            Catch
            End Try
            If exportedFolder Is Nothing Then
                Try
                    exportedFolder = tempRoot.Folders(folderName)
                Catch
                End Try
            End If
            If exportedFolder Is Nothing Then
                ' 優化第六點：提取 Folders 集合以利釋放 (by Gemini 3 Flash, 2026/05/05)
                Dim rootFolders As Outlook.Folders = tempRoot.Folders
                Try
                    For Each f As Outlook.Folder In rootFolders
                        Dim fFolders As Outlook.Folders = f.Folders
                        Try
                            exportedFolder = fFolders.Item(folderName)
                            If exportedFolder IsNot Nothing Then Exit For
                        Catch
                        Finally
                            TryMarshalRelease(fFolders)
                        End Try
                    Next
                Finally
                    TryMarshalRelease(rootFolders)
                End Try
            End If

            If exportedFolder IsNot Nothing Then
                ' 顯示所有匯出的郵件
                Dim itemCount = 0
                ' 優化第六點：提取 Items 集合以利釋放 (by Gemini 3 Flash, 2026/05/05)
                Dim mailItems As Outlook.Items = exportedFolder.Items
                Try
                    For Each item As Object In mailItems
                        If TypeOf item Is Outlook.MailItem Then
                            Dim mail As Outlook.MailItem = DirectCast(item, Outlook.MailItem)
                            mail.Display()
                            itemCount += 1
                        ElseIf TypeOf item Is Outlook.ContactItem Then
                            Dim contact As Outlook.ContactItem = DirectCast(item, Outlook.ContactItem)
                            contact.Display()
                            itemCount += 1
                        End If
                    Next
                Finally
                    TryMarshalRelease(mailItems)
                End Try
                PgrsBar1.Text = $"成功開啟 {itemCount} 封郵件。"
            Else
                MessageBox.Show("在暫存 PST 中找不到剛匯出的郵件。", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If

            ' 延遲卸載 PST，讓 Outlook 有時間打開視窗並讀取資源 (MailItem.Display 是非同步且依賴 PST 存在)
            ' 既然只是看信，讓它保持掛載一段時間或直到下次操作，這裡先 Delay 3 秒後嘗試卸載。
            ' 更穩定的作法是不卸載，或者在 Form 關閉時清理，但考量到使用者只是一瞥，我們先延遲卸載。
            Dim cleanupTask = Task.Run(Async Sub()
                                           Await Task.Delay(5000)
                                           Try
                                               ns.RemoveStore(tempRoot)
                                               System.IO.File.Delete(tempPstPath)
                                           Catch
                                           End Try
                                       End Sub)

        Catch ex As System.Exception
            MessageBox.Show("開啟過程中發生錯誤: " & ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Cursor = Cursors.Default
            PgrsBar2.Text = ""
        End Try
    End Sub
#End Region
#Region "  └ Layer3 OST 郵件解析核心"
    Private Function ResetOstPassword(filePath As String) As Boolean
        ''' <summary>
        ''' 強制重置 OST 檔案的密碼旗標。
        ''' 藉由將 Header (offset 0x42) 的 dwPassword 歸零來繞過密碼檢查。
        ''' by Gemini 3.0 Flash, 2026/04/23
        ''' </summary>
        Try
            ' 1. 開啟檔案執行讀寫
            Using fs As New System.IO.FileStream(filePath, System.IO.FileMode.Open, System.IO.FileAccess.ReadWrite)

                ' 2. 檢查是否為 Unicode PST/OST (Magic number '!BDN' at offset 0)
                Dim magic(3) As Byte
                fs.Read(magic, 0, 4)
                If magic(0) <> &H21 OrElse magic(1) <> &H42 OrElse magic(2) <> &H44 OrElse magic(3) <> &H4E Then
                    _dbg("ResetOstPassword", "非有效的 PST/OST 檔案格式")
                    Return False
                End If

                ' 3. 跳到 dwPassword 欄位 (Unicode 偏移為 66 / 0x42)
                '    注意：若是極老舊的 ANSI 格式則是 0x40，目前暫以 Unicode 為主
                fs.Position = &H42
                Dim zero4(3) As Byte ' {0, 0, 0, 0}
                fs.Write(zero4, 0, 4)

                ' 4. 確保加密模式設為 PERMUTE (offset 0x41)
                '    如果原本是 CYCLIC (0x02)，也把它降級為 PERMUTE (0x01)
                fs.Position = &H41
                fs.WriteByte(&H1)
                _dbg("ResetOstPassword 成功", "已抹除密碼雜湊並強制設定為 Permute 混淆模式")
            End Using
            Return True
        Catch ex As System.Exception
            _dbg("ResetOstPassword 失敗", ex.Message)
            Return False
        End Try
    End Function
    Private Function ParseOstContentsL3(ostFolder As ost2pst.Folder) As List(Of OstMailRow)
        ' ---------------------------------------------------------------
        ' ParseOstContentsL3 — 讀取 OST 資料夾的郵件清單（純 file I/O，可在 Task.Run）
        ' 流程：
        '   1. 計算 CONTENTS_TABLE NID：保留 nid 高 27 bits，低 5 bits 改為 4 (CONTENTS_TABLE)
        '   2. 在 FM.srcFile.NBTs 搜尋該 NID 的 NBTENTRY
        '   3. 呼叫 LTP. → TC 含 tcRowMatrix (各 row 已解出 Properties)
        '   4. 逐 Row 用 OstPropStr / OstPropDT / OstPropI32 解出欄位
        '
        ' NID 結構 (MS-PST spec):
        '   bits[4:0]  = nidType    (NORMAL_FOLDER=1, HIERARCHY_TABLE=3, CONTENTS_TABLE=4)
        '   bits[31:5] = nidIndex   (資料夾獨有序號)
        '   → Contents Table NID =  (folderNid And Not 31UI) Or 4UI
        '
        ' 2026/04/22 by Claude
        ' 2026/04/23 by Gemini 3.0 Flash
        ' ---------------------------------------------------------------

        Dim result As New List(Of OstMailRow)()
        _dbg("===> 進入 ParseOstContentsL3", $"資料夾: {ostFolder.name}, NID: {ostFolder.nid.dwValue:X}")

        If ost2pst.FM.srcFile Is Nothing Then
            _dbg("===> 失敗: FM.srcFile 為 Nothing")
            Return result
        End If

        Try
            ' ── Step 1: 計算CONTENTS_TABLE NID ─────────────────────────
            ' NID 結構: bits[4:0] 是類型，4 代表內容清單
            ' Dim contentNid As UInteger = (sourceFolderOST.nid.dwValue And Not 31UI) Or 4UI
            '
            ' 2026/04/23 by Gemini 3.0 Flash: 根據 libpff 規格修正
            ' Type 0x0E (14) = SUB_MESSAGES (也就是真正的 Contents Table)
            ' Type 0x04 (4)  = MESSAGE (單封郵件節點)
            Dim contentNid As UInteger = (ostFolder.nid.dwValue And Not 31UI) Or 14UI
            _dbg("    ├ 原始 Folder NID", ostFolder.nid.dwValue.ToString("X"))
            _dbg("    ├ 計算內容表 NID", contentNid.ToString("X"))

            ' ── Step 2: 在 NBT 找對應 NBTENTRY ───────────────────────────
            ' NBT (Node B-Tree) 儲存了所有 Node 的位址與大小
            Dim nbtIdx As Integer = ost2pst.FM.srcFile.NBTs.FindIndex(Function(n) n.nid.dwValue = contentNid)
            _dbg("    ├ NBT 搜尋結果索引", nbtIdx.ToString())

            If nbtIdx < 0 Then
                _dbg("    ❌ 找不到 CONTENTS_TABLE (可能為空資料夾或沒有內容)", $"{ostFolder.name} (nid={ostFolder.nid.dwValue:X})")
                Return result
            End If
            Dim nbt As ost2pst.NBTENTRY = ost2pst.FM.srcFile.NBTs(nbtIdx)
            _dbg("    ├ 找到 NBTENTRY", $"bidData={nbt.bidData:X}, bidSub={nbt.bidSub:X}")

            ' ── Step 3: 讀取 TableContext（TC inline row data，非逐封讀完整 PC）─
            ' ReadTCs_and_rowdata：直接讀 TC 儲存的行資料，速度快；
            Dim tc As ost2pst.TableContext = Nothing
            _dbg("    ├ 正在呼叫 ReadTCs_and_rowdata...")
            Try
                tc = ost2pst.LTP.ReadTCs_and_rowdata(ost2pst.FM.srcFile.stream, nbt)
            Catch ex As System.Exception
                _dbg("    ❌ 呼叫 ReadTCs_and_rowdata 失敗", ex.Message)
                Return result
            End Try

            If tc Is Nothing Then
                _dbg("    ⚠️ TableContext 為 Nothing", ostFolder.name)
                Return result
            End If
            _dbg("    ├ TableContext 讀取成功", $"Rows={If(tc.tcRowMatrix Is Nothing, 0, tc.tcRowMatrix.Count)}")
            If tc Is Nothing OrElse tc.tcRowMatrix Is Nothing Then Return result

            ' ── Step 4: 逐 Row 解出屬性 ──────────────────────────────────
            For i As Integer = 0 To tc.tcRowMatrix.Count - 1
                Dim row = tc.tcRowMatrix(i)
                Dim props = row.Props
                If props Is Nothing Then Continue For

                Dim item As New OstMailRow()
                item.Subject = OstPropStr(props, PROP_SUBJECT)
                item.ReceivedTime = OstPropDT(props, PROP_DELIVERY_TIME)

                ' 優先讀 Unicode 版寄件者名稱，沒有再用 ANSI 版
                item.SenderName = OstPropStr(props, PROP_SENDER_NAME_W)
                If String.IsNullOrEmpty(item.SenderName) Then
                    item.SenderName = OstPropStr(props, PROP_SENDER_NAME)
                End If
                item.SizeBytes = OstPropI32(props, PROP_MSG_SIZE)

                ' PidTagHasAttachments: PtypBoolean = 1 byte，非零 = True
                Dim hasAttProp = props.FirstOrDefault(Function(p) CInt(p.id) = PROP_HAS_ATTACH)
                item.HasAttachments = hasAttProp IsNot Nothing AndAlso
                                      hasAttProp.data IsNot Nothing AndAlso
                                      hasAttProp.data.Length > 0 AndAlso
                                      hasAttProp.data(0) <> 0

                item.IsRead = (OstPropI32(props, PROP_MSG_FLAGS) And 1) <> 0

                ' 抓取 EntryID (如果存在)
                Dim eidProp = props.FirstOrDefault(Function(p) CInt(p.id) = PROP_ENTRYID)
                If eidProp IsNot Nothing AndAlso eidProp.data IsNot Nothing Then
                    ' EntryID 是 Binary，轉換為 Hex String (Outlook 常用格式)
                    item.EntryID = BitConverter.ToString(eidProp.data).Replace("-", "")
                End If

                ' By Gemini 3.0 Flash: 總是存下 NID 以供後續匯出單封郵件使用
                ' RowData 沒有 dwRowID，實際的 dwRowID 存在平行的 tcRowIndexes 陣列中
                If tc.tcRowIndexes IsNot Nothing AndAlso i < tc.tcRowIndexes.Count Then
                    item.Nid = tc.tcRowIndexes(i).dwRowID
                    If String.IsNullOrEmpty(item.EntryID) Then item.EntryID = item.Nid.ToString()
                End If
                result.Add(item)
            Next

        Catch ex As System.Exception
            _dbg("ParseOstContentsL3 錯誤", ex.Message)
        End Try

        Return result
    End Function

    ' ── 屬性解碼輔助函數（低 16 bits 比對 Property.id 列舉值）───────────
    ' 設計: CInt(p.id) 把 EpropertyId 列舉轉為整數後與 MAPI tag 比對
    '       不依賴 EpropertyId 的具名常數，避免 enum 未涵蓋的屬性找不到
    Private Function OstPropStr(props As List(Of ost2pst.Property), tagId As Integer) As String
        ' 讀取 Unicode 或 ANSI 字串屬性；失敗或不存在回傳空字串
        For Each p In props
            If CInt(p.id) <> tagId Then Continue For
            If p.data Is Nothing OrElse p.data.Length = 0 Then Return ""
            Try
                Select Case p.type
                    Case ost2pst.EpropertyType.PtypString : Return Encoding.Unicode.GetString(p.data)
                    Case ost2pst.EpropertyType.PtypString8 : Return Encoding.Default.GetString(p.data)
                End Select
            Catch
            End Try
        Next
        Return ""
    End Function
    Private Function OstPropDT(props As List(Of ost2pst.Property), tagId As Integer) As DateTime
        ' 讀取 PtypTime 屬性（8 bytes，Windows FILETIME，UTC）；失敗回 DateTime.MinValue
        For Each p In props
            If CInt(p.id) <> tagId Then Continue For
            If p.data Is Nothing OrElse p.data.Length < 8 Then Return DateTime.MinValue
            Try : Return DateTime.FromFileTimeUtc(BitConverter.ToInt64(p.data, 0)).ToLocalTime()
            Catch : Return DateTime.MinValue
            End Try
        Next
        Return DateTime.MinValue
    End Function
    Private Function OstPropI32(props As List(Of ost2pst.Property), tagId As Integer) As Integer
        ' 讀取 PtypInteger32 屬性（4 bytes，little-endian）；失敗回 0
        For Each p In props
            If CInt(p.id) <> tagId Then Continue For
            If p.data Is Nothing OrElse p.data.Length < 4 Then Return 0
            Try : Return BitConverter.ToInt32(p.data, 0)
            Catch : Return 0
            End Try
        Next
        Return 0
    End Function
#End Region
#End Region

    Private Class Tab7LviComparer   ' ── Tab7 專用排序比較器 (by Gemini 3 Flash, 2026/04/24) ─────────────────
        Implements IComparer
        Private ReadOnly _col As Integer
        Private ReadOnly _order As SortOrder
        Public Sub New(column As Integer, order As SortOrder)
            _col = column
            _order = order
        End Sub
        Public Function Compare(x As Object, y As Object) As Integer Implements IComparer.Compare
            Dim itemX = DirectCast(x, ListViewItem)
            Dim itemY = DirectCast(y, ListViewItem)
            Dim res As Integer = 0

            Select Case _col
                Case 1 ' 郵件大小: 從 Tag 或解析字串 (移除千分位)
                    Dim valX As Long = 0, valY As Long = 0
                    If TypeOf itemX.Tag Is OstMailRow Then
                        valX = DirectCast(itemX.Tag, OstMailRow).SizeBytes
                        valY = DirectCast(itemY.Tag, OstMailRow).SizeBytes
                    ElseIf TypeOf itemX.Tag Is MailItemInfo Then
                        valX = DirectCast(itemX.Tag, MailItemInfo).Size
                        valY = DirectCast(itemY.Tag, MailItemInfo).Size
                    Else
                        Long.TryParse(itemX.SubItems(1).Text.Replace(",", ""), valX)
                        Long.TryParse(itemY.SubItems(1).Text.Replace(",", ""), valY)
                    End If
                    res = valX.CompareTo(valY)

                Case 2 ' 收到日期
                    Dim dateX As DateTime = DateTime.MinValue, dateY As DateTime = DateTime.MinValue
                    If TypeOf itemX.Tag Is OstMailRow Then
                        dateX = DirectCast(itemX.Tag, OstMailRow).ReceivedTime
                        dateY = DirectCast(itemY.Tag, OstMailRow).ReceivedTime
                    ElseIf TypeOf itemX.Tag Is MailItemInfo Then
                        dateX = DirectCast(itemX.Tag, MailItemInfo).RcvTime
                        dateY = DirectCast(itemY.Tag, MailItemInfo).RcvTime
                    Else
                        DateTime.TryParse(itemX.SubItems(2).Text, dateX)
                        DateTime.TryParse(itemY.SubItems(2).Text, dateY)
                    End If
                    res = dateX.CompareTo(dateY)

                Case Else ' 文字欄位: 主旨 (0), 寄件者 (3), EntryID (4)
                    res = String.Compare(itemX.SubItems(_col).Text, itemY.SubItems(_col).Text, StringComparison.CurrentCultureIgnoreCase)
            End Select

            Return If(_order = SortOrder.Ascending, res, -res)
        End Function
    End Class

End Class
