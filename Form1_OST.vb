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
'   - OST: LTP.ReadTCs_and_rowdata → TC row data → OstMailRow
'   - PST: GetFolderBasicMailInfosL3 (OOM，Form1_Outlook.vb)
' Phase 3 (待實作): CopyFolder / MoveFolder → OOM 寫入目標 PST
'
' 注意:
'   OST 讀取：使用 C# ost2pst library (ost2pst.FM)，只使用其讀取路徑
'   PST 讀取：使用 Outlook OOM AddStore，與 Tab1 相同路徑
'   PST 寫入 (Phase 3)：也使用 OOM，不使用 ost2pst PST writer (已知損壞)
'   FM.StatusMsg delegate：橋接 C# 內部進度訊息到 ProgressBar2
'   2026/04/19 by Claude / Phase 2 2026/04/22 by Claude
' ==============================================================

Partial Class Form1

#Region "■ 10 Tab7: OST/PST 解析"

    ' ── Phase 2 模組級欄位 ────────────────────────────────────────────────
    ' 2026/04/22 by Claude
    Private _ostLoaded As Boolean = False        ' True = OST 已成功開啟，FM.srcFile 可用
    Private WithEvents _lvOST As New ListView()  ' OST 郵件清單（動態建立）
    Private WithEvents _lvPST As New ListView()  ' PST 郵件清單（動態建立）
    Private _tab7Initialized As Boolean = False  ' EnsureTab7Phase2UI 是否已執行過

    ' OST 郵件列資料結構（純 .NET 值型別，不持有 COM 物件）
    Private Structure OstMailRow
        Dim Subject As String
        Dim ReceivedTime As DateTime
        Dim SenderName As String
        Dim SizeBytes As Long           ' PR_MESSAGE_SIZE（bytes）
        Dim HasAttachments As Boolean   ' PidTagHasAttachments
        Dim IsRead As Boolean           ' PidTagMessageFlags bit 0
    End Structure

    ' ── 常用 MAPI 屬性 Tag ID（低 16 bits = Property ID）──────────────────
    ' 用於 OstPropStr / OstPropDT / OstPropI32 搜尋 Property.id 的比對值
    Private Const PROP_SUBJECT As Integer = &H37        ' PidTagSubjectW (Unicode)
    Private Const PROP_DELIVERY_TIME As Integer = &HE06 ' PidTagMessageDeliveryTime (FILETIME)
    Private Const PROP_SENDER_NAME As Integer = &HC1A   ' PidTagSenderName (ANSI)
    Private Const PROP_SENDER_NAME_W As Integer = &HC1B ' PidTagSenderNameW (Unicode)
    Private Const PROP_MSG_SIZE As Integer = &HE08      ' PidTagMessageSize (PT_LONG)
    Private Const PROP_HAS_ATTACH As Integer = &HE1B    ' PidTagHasAttachments (PT_BOOLEAN)
    Private Const PROP_MSG_FLAGS As Integer = &HE07     ' PidTagMessageFlags (PT_LONG)

#Region "  ├ Layer1 UI 事件"
    Private Sub LoadOST_Click(sender As Object, e As EventArgs) Handles LoadOST.Click
        ' 彈出 FileDialog 選擇 OST 檔，再呼叫 Layer2 解析
        ' Phase 2 先確保 ListView 已建立
        Using dlg As New OpenFileDialog() With {
            .Title = "選擇要解析的 OST 檔案",
            .Filter = "OST 檔案 (*.ost)|*.ost|所有檔案 (*.*)|*.*",
            .InitialDirectory = My.Application.Info.DirectoryPath
        }
            If dlg.ShowDialog() <> DialogResult.OK Then Return
            EnsureTab7Phase2UI()     ' Phase 2: 確保 ListView 已建立並完成版面配置
            LoadOstToTree(dlg.FileName, SimTreeOST)
        End Using
    End Sub

    Private Sub LoadPST_Click(sender As Object, e As EventArgs) Handles LoadPST.Click
        ' 彈出 FileDialog 選擇 PST 檔，再呼叫 Layer2 以 OOM 載入
        ' Phase 2 先確保 ListView 已建立
        Using dlg As New OpenFileDialog() With {
            .Title = "選擇要載入的 PST 檔案",
            .Filter = "PST 檔案 (*.pst)|*.pst|所有檔案 (*.*)|*.*",
            .InitialDirectory = My.Application.Info.DirectoryPath
        }
            If dlg.ShowDialog() <> DialogResult.OK Then Return
            EnsureTab7Phase2UI()     ' Phase 2
            LoadPstToTree(dlg.FileName, SimTreePST)
        End Using
    End Sub

    Private Sub CopyFolder_Click(sender As Object, e As EventArgs) Handles CopyFolder.Click
        ' Phase 3 預留
        ' 設計意圖: OST 來源資料夾 (SimTreeOST.SelectedNode) → OOM 目標資料夾 (SimTreePST.SelectedNode)
        '            用 OOM MailItem.Copy() 逐封複製，不使用 ost2pst PST writer
        MessageBox.Show("Copy Folder 功能待實作 (Phase 3)" & vbCrLf & vbCrLf &
                        "計畫：從 SimTreeOST 選取的 OST 資料夾，" & vbCrLf &
                        "複製郵件到 SimTreePST 選取的 Outlook 目標資料夾。",
                        "Phase 3 預留", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub MoveFolder_Click(sender As Object, e As EventArgs) Handles MoveFolder.Click
        ' Phase 3 預留 (與 CopyFolder 相同架構，完成後刪除來源)
        MessageBox.Show("Move Folder 功能待實作 (Phase 3)",
                        "Phase 3 預留", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Async Sub SimTreeOST_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles SimTreeOST.AfterSelect
        ' Phase 2: 選取 OST 資料夾 → 讀取 Contents Table → 顯示郵件清單
        '
        ' 路徑: AfterSelect → ReadOstFolderContentsL3 (Task.Run，純 I/O 非 COM)
        '        → 解 TC row data → OstMailRow → ShowOstItems (UI thread)
        '
        ' 2026/04/22 by Claude
        Dim ostFolder = TryCast(e.Node?.Tag, ost2pst.Folder)
        If ostFolder Is Nothing Then Return

        ' OST 尚未載入（或載入失敗）時只更新狀態列，不讀內容
        If Not _ostLoaded OrElse ost2pst.FM.srcFile Is Nothing Then
            ProgressBar2.Text = $"OST 資料夾: {ostFolder.path}"
            Return
        End If

        _dbg("開始", ostFolder.name)
        _lvOST.Items.Clear()
        ProgressBar1.Text = "正在讀取郵件清單..." : ProgressBar2.Text = ostFolder.path
        Cursor = Cursors.WaitCursor

        Try
            ' ost2pst 讀 OST 是純 file I/O（非 COM），可安全放 Task.Run
            ' 注意：FM.srcFile.stream 是 FileStream，不支援並行存取；
            '        Await Task.Run(...) 確保此時 UI 執行緒不再存取 stream，無競爭條件。
            Dim items As List(Of OstMailRow) = Await Task.Run(
                Function() ReadOstFolderContentsL3(ostFolder))
            ShowOstItems(items)
            ProgressBar1.Text = $"共 {items.Count:N0} 封 — {ostFolder.name}"
        Catch ex As System.Exception
            _dbg("錯誤", ex.Message)
            ProgressBar1.Text = "讀取失敗: " & ex.Message
        Finally
            Cursor = Cursors.Default
            _dbg("結束", ostFolder.name)
        End Try
    End Sub

    Private Async Sub SimTreePST_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles SimTreePST.AfterSelect
        ' Phase 2: 選取 PST 資料夾 → 使用 OOM GetTable → 顯示郵件清單
        '
        ' OOM 呼叫必須在 UI 執行緒（STA）；使用 Form1_Outlook.vb 的
        ' GetFolderBasicMailInfosL3，它已有快取機制與 cToken 支援。
        '
        ' 2026/04/22 by Claude
        Dim folder = TryCast(e.Node?.Tag, Outlook.Folder)
        If folder Is Nothing Then Return

        _dbg("開始", folder.Name)
        _lvPST.Items.Clear()
        ProgressBar1.Text = "正在讀取 PST 郵件清單..." : ProgressBar2.Text = folder.FolderPath
        Cursor = Cursors.WaitCursor

        Try
            Dim cToken As System.Threading.CancellationToken = OkayNowYouHaveToken()
            ' needTopic:=False：Tab7 不需要 Conversation Topic，省去讀 PR_CONVERSATION_TOPIC 開銷
            Dim rows = Await GetFolderBasicMailInfos(folder, needTopic:=False, ct:=cToken)
            ShowPstItems(rows.Select(Function(r) r.Mail).ToList())
            ProgressBar1.Text = $"共 {rows.Count:N0} 封 — {folder.Name}"
        Catch ex As OperationCanceledException
            _dbg("中斷", "ESC") : ProgressBar1.Text = "已中斷。"
        Catch ex As System.Exception
            _dbg("錯誤", ex.Message) : ProgressBar1.Text = "讀取失敗: " & ex.Message
        Finally
            Cursor = Cursors.Default
            _dbg("結束", folder.Name)
        End Try
    End Sub
#End Region

#Region "  ├ Layer2 OST 解析流程"
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
        _lvOST.Items.Clear()
        ProgressBar1.Text = "正在解析 OST..." : ProgressBar2.Text = ""
        Await Task.Yield()  ' 讓 UI 先刷新再開始耗時操作

        ' FM.StatusMsg delegate 橋接（待 C# DLL 重新編譯後可還原）：
        ost2pst.FM.StatusMsg = Sub(msg As String)
                                   If Not String.IsNullOrEmpty(msg) Then ProgressBar2.Text = msg
                               End Sub
        Try
            ' ① 開啟 OST 檔（C# 端解析 Header + NBT/BBT B-Tree，約 0.5~2 秒）
            If Not ost2pst.FM.OpenSourceFile(filePath) Then
                MessageBox.Show("無法開啟 OST 檔案：" & vbCrLf & filePath, "錯誤",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' ② 取得資料夾清單
            '    FM.GetFolderList() 走遍所有 NORMAL_FOLDER NID，讀取 PidTagDisplayName
            '    結果存在 FM.folders: List(Of ost2pst.Folder)
            '    每個 Folder 含: .name / .path / .parent (物件參考) / .level / .nbtIndex / .nid
            ost2pst.FM.GetFolderList()

            Dim folderList = ost2pst.FM.folders
            If folderList Is Nothing OrElse folderList.Count = 0 Then
                MessageBox.Show("OST 檔案內找不到任何資料夾。", "提示",
                                MessageBoxButtons.OK, MessageBoxIcon.Information)
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
            ProgressBar1.Text = $"OST 解析完成：共 {folderList.Count} 個資料夾，請點選資料夾查看郵件"
            ProgressBar2.Text = filePath
            _dbg("結束", $"{folderList.Count} 個資料夾")

        Catch ex As System.Exception
            MessageBox.Show("解析 OST 時發生錯誤：" & vbCrLf & ex.Message,
                            "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
            _dbg("錯誤", ex.Message)
            ' 失敗時確保關閉 handle
            ost2pst.FM.CloseSourceFile()
            _ostLoaded = False
        Finally
            ' Phase 2: 不在 Finally 關閉 OST，保持開啟供 AfterSelect 讀郵件
            'ost2pst.FM.StatusMsg = Sub(msg As String) _ = msg  ' 清除 callback（待 DLL 更新）
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
        Dim nodeMap As New Dictionary(Of ost2pst.Folder, TreeNode)(
            ReferenceEqualityComparer.Instance)

        ' ── 第一步：找並建立根節點 ──────────────────────────────────
        ' 根節點條件：parent 指向自己，或 parent Is Nothing
        Dim rootFolders = folders.Where(
            Function(f) f.parent Is f OrElse f.parent Is Nothing).ToList()

        If rootFolders.Count = 0 Then
            ' Fallback：找 level = 0 的（部分 OST 格式根節點不自指）
            rootFolders = folders.Where(Function(f) f.level = 0).ToList()
        End If
        If rootFolders.Count = 0 Then
            ' 最後手段：把第一筆當根
            rootFolders = New List(Of ost2pst.Folder) From {folders(0)}
        End If

        For Each root In rootFolders
            Dim displayName As String = If(String.IsNullOrEmpty(root.name), "OST Root", root.name)
            Dim rootNode As New TreeNode(displayName) With {.Tag = root}
            tv.Nodes.Add(rootNode)
            nodeMap(root) = rootNode
        Next

        ' ── 第二步：BFS 多輪把子節點掛上去 ─────────────────────────
        Dim pending = folders.Where(
            Function(f) Not nodeMap.ContainsKey(f) AndAlso
                        Not rootFolders.Contains(f, ReferenceEqualityComparer.Instance)).ToList()

        Dim maxRounds As Integer = 50
        Do While pending.Count > 0 AndAlso maxRounds > 0
            maxRounds -= 1
            Dim stillPending As New List(Of ost2pst.Folder)
            For Each f In pending
                If f.parent IsNot Nothing AndAlso nodeMap.ContainsKey(f.parent) Then
                    Dim displayName As String = If(String.IsNullOrEmpty(f.name), "(未命名)", f.name)
                    Dim childNode As New TreeNode(displayName) With {.Tag = f}
                    nodeMap(f.parent).Nodes.Add(childNode)
                    nodeMap(f) = childNode
                Else
                    stillPending.Add(f)   ' parent 尚未建立，下一輪再試
                End If
            Next
            ' 若這輪一個都沒能掛上去，代表有孤兒，直接跳出避免無限迴圈
            If stillPending.Count = pending.Count Then Exit Do
            pending = stillPending
        Loop

        ' ── 第三步：處理孤兒節點（掛到第一個根節點下，記警告）──────
        For Each orphan In pending
            _dbg("⚠️ OST 孤兒資料夾", $"path={orphan.path} parent={orphan.parent?.name}")
            Dim displayName As String = If(String.IsNullOrEmpty(orphan.name), "(孤兒)", orphan.name)
            Dim orphanNode As New TreeNode(displayName) With {.Tag = orphan}
            If tv.Nodes.Count > 0 Then
                tv.Nodes(0).Nodes.Add(orphanNode)
            Else
                tv.Nodes.Add(orphanNode)
            End If
        Next
    End Sub
#End Region

#Region "  ├ Layer3 OST 郵件讀取"
    Private Function ReadOstFolderContentsL3(ostFolder As ost2pst.Folder) As List(Of OstMailRow)
        ' ---------------------------------------------------------------
        ' ReadOstFolderContentsL3 — 讀取 OST 資料夾的郵件清單（純 file I/O，可在 Task.Run）
        '
        ' 流程：
        '   1. 計算 CONTENTS_TABLE NID：保留 nid 高 27 bits，低 5 bits 改為 4 (CONTENTS_TABLE)
        '   2. 在 FM.srcFile.NBTs 搜尋該 NID 的 NBTENTRY
        '   3. 呼叫 LTP.ReadTCs_and_rowdata → TC 含 tcRowMatrix (各 row 已解出 Properties)
        '   4. 逐 Row 用 OstPropStr / OstPropDT / OstPropI32 解出欄位
        '
        ' NID 結構 (MS-PST spec):
        '   bits[4:0]  = nidType  (NORMAL_FOLDER=1, HIERARCHY_TABLE=3, CONTENTS_TABLE=4)
        '   bits[31:5] = nidIndex (資料夾獨有序號)
        '   → Contents Table NID = (folderNid And Not 31UI) Or 4UI
        '
        ' 2026/04/22 by Claude
        ' ---------------------------------------------------------------
        Dim result As New List(Of OstMailRow)()
        If ost2pst.FM.srcFile Is Nothing Then Return result

        Try
            ' ── Step 1: 計算 CONTENTS_TABLE NID ─────────────────────────
            Dim contentNid As UInteger = (ostFolder.nid.dwValue And Not 31UI) Or 4UI

            ' ── Step 2: 在 NBT 找對應 NBTENTRY ───────────────────────────
            Dim nbtIdx As Integer = ost2pst.FM.srcFile.NBTs.FindIndex(
                Function(n) n.nid.dwValue = contentNid)
            If nbtIdx < 0 Then
                _dbg("    ├ CONTENTS_TABLE 不存在", $"{ostFolder.name} (nid={ostFolder.nid.dwValue:X})")
                Return result
            End If
            Dim nbt As ost2pst.NBTENTRY = ost2pst.FM.srcFile.NBTs(nbtIdx)

            ' ── Step 3: 讀取 TableContext（TC inline row data，非逐封讀完整 PC）─
            ' ReadTCs_and_rowdata：直接讀 TC 儲存的行資料，速度快；
            '   不同於 ReadTCs_new_rowdata（逐封讀完整 PC，適合 PST 轉換但太慢）
            Dim tc As ost2pst.TableContext = ost2pst.LTP.ReadTCs_and_rowdata(
                ost2pst.FM.srcFile.stream, nbt)
            If tc Is Nothing OrElse tc.tcRowMatrix Is Nothing Then Return result

            ' ── Step 4: 逐 Row 解出屬性 ──────────────────────────────────
            For Each row In tc.tcRowMatrix
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

                ' PidTagMessageFlags: bit 0 (MSGFLAG_READ = 0x01)，1 = 已讀
                item.IsRead = (OstPropI32(props, PROP_MSG_FLAGS) And 1) <> 0

                result.Add(item)
            Next

        Catch ex As System.Exception
            _dbg("ReadOstFolderContentsL3 錯誤", ex.Message)
        End Try

        Return result
    End Function

    ' ── 屬性解碼輔助函數（低 16 bits 比對 Property.id 列舉值）───────────
    ' 設計: CInt(p.id) 把 EpropertyId 列舉轉為整數後與 MAPI tag 比對
    '       不依賴 EpropertyId 的具名常數，避免 enum 未涵蓋的屬性找不到
    '
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

#Region "  ├ Phase 2 UI 初始化 & 渲染"
    Private Sub EnsureTab7Phase2UI()
        ' ---------------------------------------------------------------
        ' EnsureTab7Phase2UI — 建立並定位 Tab7 的兩個 ListView（只初始化一次）
        '
        ' 策略: 取 SimTreeOST.Parent 與 SimTreePST.Parent，
        '       各自用 ArrangeTab7ListView 插入水平 SplitContainer，
        '       把 TreeView 移到 Panel1 (上)、ListView 移到 Panel2 (下)。
        '
        ' 相容性: 無論 Designer 把 TreeView 放在 TabPage、Panel 或 SplitterPanel 裡
        '         都能正確處理（有 Try/Catch fallback）。
        '
        ' 2026/04/22 by Claude
        ' ---------------------------------------------------------------
        If _tab7Initialized Then Return
        _tab7Initialized = True

        ' ── 設定兩個 ListView 的欄位與共用外觀 ─────────────────────────
        For Each lv In {_lvOST, _lvPST}
            lv.View = System.Windows.Forms.View.Details
            lv.FullRowSelect = True
            lv.GridLines = False
            lv.Font = New Font("Microsoft Jhenghei", 9.5F)
            ' 開啟 .NET 反射雙緩衝（與 InitListView 相同做法）
            GetType(ListView).GetProperty("DoubleBuffered",
                Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic)?.
                SetValue(lv, True, Nothing)
            lv.Columns.Add("主旨", 300)
            lv.Columns.Add("收到時間", 135, HorizontalAlignment.Center)
            lv.Columns.Add("寄件者", 140)
            lv.Columns.Add("大小 (KB)", 70, HorizontalAlignment.Right)
            lv.Columns.Add("附件", 35, HorizontalAlignment.Center)
        Next

        ' ── 定位：把 TreeView 和 ListView 組合進垂直 SplitContainer ─────
        ArrangeTab7ListView(SimTreeOST, _lvOST)
        ArrangeTab7ListView(SimTreePST, _lvPST)
    End Sub

    Private Sub ArrangeTab7ListView(treeView As TreeView, lv As ListView)
        ' 把 treeView 換進垂直 SplitContainer 的 Panel1，lv 放 Panel2。
        ' 若 treeView 的 Parent 已是 SplitterPanel（即已在 SplitContainer 裡），
        ' 也能正確處理，因為 SplitterPanel 是 Panel 的子類，可直接 Add 子控制項。
        Dim parent As Control = treeView.Parent
        If parent Is Nothing Then Return

        parent.SuspendLayout()
        Try
            Dim origBounds As Rectangle = treeView.Bounds
            Dim origAnchor As AnchorStyles = treeView.Anchor
            Dim origDock As DockStyle = treeView.Dock

            parent.Controls.Remove(treeView)

            Dim sc As New SplitContainer() With {
                .Orientation = Orientation.Horizontal,
                .Panel1MinSize = 80,
                .Panel2MinSize = 60
            }
            ' 繼承 TreeView 原本的位置/大小/Dock/Anchor
            If origDock <> DockStyle.None Then
                sc.Dock = origDock
            Else
                sc.Location = origBounds.Location
                sc.Size = origBounds.Size
                sc.Anchor = origAnchor
            End If

            treeView.Dock = DockStyle.Fill
            sc.Panel1.Controls.Add(treeView)
            lv.Dock = DockStyle.Fill
            sc.Panel2.Controls.Add(lv)
            parent.Controls.Add(sc)

            ' SplitContainer 建立後才能設 SplitterDistance（需 Handle 存在）
            ' 用 BeginInvoke 確保 Layout 結算後再設
            sc.BeginInvoke(Sub()
                               Try
                                   If sc.Height > 0 Then sc.SplitterDistance = CInt(sc.Height * 0.55)
                               Catch
                               End Try
                           End Sub)
        Catch ex As System.Exception
            _dbg("ArrangeTab7ListView 失敗，使用 Fallback", ex.Message)
            ' Fallback: 把 TreeView 加回原 parent，ListView 固定在下方
            Try
                parent.Controls.Add(treeView)
                lv.Dock = DockStyle.Bottom
                lv.Height = 200
                parent.Controls.Add(lv)
            Catch
            End Try
        Finally
            parent.ResumeLayout()
        End Try
    End Sub

    Private Sub ShowOstItems(items As List(Of OstMailRow))
        ' 把 OstMailRow 清單渲染到 _lvOST；未讀郵件以粗體顯示
        _lvOST.BeginUpdate()
        _lvOST.Items.Clear()
        For Each item In items
            Dim lvi As New ListViewItem(item.Subject)
            lvi.SubItems.Add(If(item.ReceivedTime > DateTime.MinValue,
                               item.ReceivedTime.ToString("yyyy/MM/dd HH:mm"), ""))
            lvi.SubItems.Add(item.SenderName)
            lvi.SubItems.Add(If(item.SizeBytes > 0, (item.SizeBytes \ 1024L).ToString("N0"), ""))
            lvi.SubItems.Add(If(item.HasAttachments, "Y", ""))   ' 用 "Y" 取代 emoji 避免字型問題
            If Not item.IsRead Then lvi.Font = New Font(_lvOST.Font, FontStyle.Bold)
            _lvOST.Items.Add(lvi)
        Next
        _lvOST.EndUpdate()
    End Sub

    Private Sub ShowPstItems(mails As List(Of MailItemInfo))
        ' 把 MailItemInfo 清單（來自 GetFolderBasicMailInfosL3）渲染到 _lvPST
        _lvPST.BeginUpdate()
        _lvPST.Items.Clear()
        For Each mail In mails
            Dim lvi As New ListViewItem(mail.Subject)
            lvi.SubItems.Add(If(mail.ReceivedTime > DateTime.MinValue,
                               mail.ReceivedTime.ToString("yyyy/MM/dd HH:mm"), ""))
            lvi.SubItems.Add(mail.SenderName)
            lvi.SubItems.Add(If(mail.Size > 0, (mail.Size \ 1024L).ToString("N0"), ""))
            lvi.SubItems.Add(If(mail.AttachCount > 0, "Y", ""))
            _lvPST.Items.Add(lvi)
        Next
        _lvPST.EndUpdate()
    End Sub
#End Region

#Region "  └ Layer2 PST 載入流程 (OOM)"
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
        _lvPST.Items.Clear()
        ProgressBar1.Text = "載入 PST..." : ProgressBar2.Text = ""

        Dim ns As Outlook.NameSpace = Nothing
        Try
            ns = _olApp.GetNamespace("MAPI")

            ' ── 檢查是否已掛入（比對 FilePath，避免重複 AddStore）──
            Dim targetStore As Outlook.Store = Nothing
            For Each store As Outlook.Store In ns.Stores
                If store.FilePath IsNot Nothing AndAlso
                   String.Compare(store.FilePath, filePath,
                                  StringComparison.OrdinalIgnoreCase) = 0 Then
                    targetStore = store : Exit For
                End If
            Next

            If targetStore Is Nothing Then
                ns.AddStore(filePath)
                ' AddStore 完成後再搜尋一次
                For Each store As Outlook.Store In ns.Stores
                    If store.FilePath IsNot Nothing AndAlso
                       String.Compare(store.FilePath, filePath,
                                      StringComparison.OrdinalIgnoreCase) = 0 Then
                        targetStore = store : Exit For
                    End If
                Next
            End If

            If targetStore Is Nothing Then
                MessageBox.Show("加入 PST 後仍找不到對應的 Store。" & vbCrLf &
                                "可能是 PST 格式損毀或版本不支援。",
                                "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
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
            ProgressBar1.Text = $"PST 載入完成：{rootF.Name}，共 {totalNodes} 個資料夾，請點選資料夾查看郵件"
            ProgressBar2.Text = filePath
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
            For Each subF As Outlook.Folder In folder.Folders
                Dim node As New TreeNode(subF.Name) With {.Tag = subF}
                parentNode.Nodes.Add(node)
                If subF.Folders.Count > 0 Then LoadPstSubFoldersRecursive(subF, node)
            Next
        Catch ex As System.Exception
            _dbg("LoadPstSubFoldersRecursive 錯誤", ex.Message)
        End Try
    End Sub

    Private Function CountAllNodes(nodes As TreeNodeCollection) As Integer
        ' 遞迴計算 TreeView 節點總數（供狀態列顯示用）
        Dim count As Integer = nodes.Count
        For Each n As TreeNode In nodes
            count += CountAllNodes(n.Nodes)
        Next
        Return count
    End Function
#End Region

#End Region

End Class
