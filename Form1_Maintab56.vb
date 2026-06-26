Imports System.Runtime.InteropServices
Imports System.Threading
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Outlook

Partial Class Form1

#Region "■ 01 全域宣告"
    ' ── Tab5 SimTree 多選控制項 (2026/05/01 by Claude: 取代舊版 TreeView5，對齊 Tab1~4 操作行為) ──
    'Private Listview6 As ListView = Nothing                    ' by Gemini 3 Flash, 2026/04/20: 動態建立的統計列表
    Private rbExactMatch, rbFuzzyMatch As New RadioButton()     ' tab5 用到的radio button
    Private _includeSubTab5 As Boolean = True                   ' Tab5 是否含子資料夾，由 CheckSubFolder5 CheckBox 控制，預設 True
    Private _tv5PrevSearchMode As Boolean = True                ' by Gemini 3 Flash, 2026/05/06: 記憶最後一次掃描的模式
    Private _lv5LastSortColumn As Integer = -1                  ' by Simon/Claude, 2026/05/10: Tab5 欄位排序狀態
    Private _lv5SortOrder As SortOrder = SortOrder.Ascending    ' by Simon/Claude, 2026/05/10: Tab5 欄位排序狀態

    Private _lv5PrevGroupResults As Dictionary(Of String, List(Of MailItemInfo)) = Nothing  ' by Gemini 3 Flash, 2026/05/06: 記憶 Tab5 掃描結果，供刪除後重新渲染使用
    Private _lv5FuzzyScoreMap As Dictionary(Of String, Double) = Nothing                    ' 2026/06/17 by Simon/Claude Opus 4.8: Tab5 Fuzzy EntryID→對群代表的 body bigram Jaccard，供 RenderLv5Group 及排序/刪除重渲染查表顯示(不重讀 body)
    Private Structure RefreshStats
        ' 2026/06/14 by Simon/Claude Opus 4.8: 一次刷新作業的統計，供狀態列回饋
        Dim Updated As Integer
        Dim NotFound As Integer
        Dim Errored As Integer
    End Structure
#End Region

#Region "■ 08 Tab5: 重複郵件"
#Region "  ├ Layer1 UI事件層"
    Private Async Sub Bt5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        ' ---------------------------------------------------------------
        ' Bt5_Click — 掃描重複郵件 (Layer1，約 25 行)
        ' 2026/05/05 by Claude: 重構拆分
        '   Layer2: ScanMailsToGroupDictAsync / RenderLv5Group
        '   Helper: BuildMailGroupKey（含 Message-ID 主鍵 + Exact 容忍分桶）
        ' ---------------------------------------------------------------
        _dbg("開始")
        Dim selectedNodes As List(Of TreeNode) = SimTree5.SelectedNodes
        If selectedNodes Is Nothing OrElse selectedNodes.Count = 0 Then
            MessageBox.Show("請先在左側選取要掃描的資料夾或 PST。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information) : Return
        End If

        Dim cToken As CancellationToken = OkayNowYouHaveToken()
        Dim isExactMode As Boolean = rbExactMatch.Checked
        Dim includeSub As Boolean = _includeSubTab5
        Button5.Enabled = False : Cursor = Cursors.WaitCursor
        ListView5.BeginUpdate() : ListView5.Items.Clear() : ListView5.EndUpdate() : _refreshedList.Clear()
        PgrsBar1.Text = "正在準備" : PgrsBar2.Text = "展開資料夾結構..."
        Dim sw As Stopwatch = Stopwatch.StartNew()  ' by Claude Sonnet 4.6, 2026/06/07
        Dim progress5 As IProgress(Of ProgressReport) = New Progress(Of ProgressReport)(Sub(p) PgrsBar2.Text = p.Message)

        Try
            Dim folderList = Await GetUniqueFolderList(selectedNodes, includeSub:=includeSub, cToken:=cToken, progress:=progress5)
            If folderList.Count = 0 Then Return

            ' 2026/06/17 by Simon/Claude Opus 4.8: Exact 維持原主鍵分桶；Fuzzy 改走 SimHash+bigram Jaccard 內文比對管線(S3→S6)
            Dim groupDict As Dictionary(Of String, List(Of MailItemInfo))
            If isExactMode Then
                groupDict = Await ScanMailsToGroupDictAsync(folderList, True, progress5, cToken)
                _lv5FuzzyScoreMap = Nothing                                                          ' Exact 不用 scoreMap
            Else
                ' 沿用 ScanMailsToGroupDictAsync 完成資料夾列舉 + L2.5 快取，攤平成全體郵件(主旨分桶鍵在 Fuzzy 不再使用)
                Dim scanned = Await ScanMailsToGroupDictAsync(folderList, False, progress5, cToken)
                Dim allMails = scanned.Values.SelectMany(Function(x) x).ToList()
                Dim targetT As Double = GetFuzzyTargetT()                                           ' S8 改接trackbar控制項參數 (低/中/高/極高)
                Await PreComputeFuzzySimHashAsync(allMails, progress5, cToken)                      ' S3 build pass(暖快取跳過已算)
                Dim cand = Await GenerateFuzzyCandidatePairs(allMails, targetT, cToken)             ' S4 size 視窗 + Hamming 一階
                Dim filt = Await FilterCandidatesByJaccardAsync(cand, targetT, progress5, cToken)   ' S5 候選 body Jaccard 精算
                Dim fuzzy = BuildFuzzyGroups(filt.Pairs, filt.Sets)                                 ' S6 union-find 分群 + scoreMap
                _lv5FuzzyScoreMap = fuzzy.ScoreMap
                groupDict = fuzzy.GroupDict
            End If
            _lv5PrevGroupResults = groupDict ' by Gemini 3 Flash, 2026/05/06: 儲存結果以供動態刪除
            _tv5PrevSearchMode = isExactMode
            Dim counts = RenderLv5Group(groupDict, isExactMode)

            sw.Stop()
            PgrsBar1.Text = $"找到 {counts.GroupCount} 組 ({counts.MailCount} 封) / 耗時 {sw.Elapsed.TotalSeconds:0.00} 秒"
            PgrsBar2.Text = ""
        Catch ex As OperationCanceledException
            _dbg("結束", "ESC 中斷") : PgrsBar1.Text = "由使用者中斷。" : PgrsBar2.Text = ""
        Catch ex As System.Exception
            MessageBox.Show("掃描重複郵件時發生錯誤: " & ex.Message, "錯誤") : _dbg("錯誤", ex.Message)
        Finally
            Button5.Enabled = True : Cursor = Cursors.Default : _dbg("結束")
        End Try

    End Sub
    Private Sub Lv5_KeyDown(sender As Object, e As KeyEventArgs) Handles ListView5.KeyDown
        ''' <summary>
        ''' by Gemini 3 Flash, 2026/05/06: 處理 ListView5 的專屬快捷鍵 (Delete)
        ''' </summary>
        _dbg("開始", e.KeyCode.ToString())
        If e.KeyCode = Keys.Delete Then
            _dbg("快捷鍵", "偵測到 Delete (呼叫 HandleLv5Delete)")
            HandleLv5Delete(DirectCast(sender, ListView))
            e.Handled = True
        End If
    End Sub
    Private Sub Lv5_ColumnClick(sender As Object, e As ColumnClickEventArgs)
        ' 2026/05/10 by Simon/Claude: Tab5 群組感知排序
        ' 點擊同一欄 → 反向；點擊新欄 → 升冪
        _lv5SortOrder = GetNewSortOrder(e.Column, _lv5LastSortColumn, _lv5SortOrder)    ' 2026/05/30 by Gemini/Simon: 抽取共用函式 GetNewSortOrder，簡化排序狀態切換邏輯
        _lv5LastSortColumn = e.Column
        If _lv5PrevGroupResults IsNot Nothing Then RenderLv5Group(_lv5PrevGroupResults, _tv5PrevSearchMode)
    End Sub
#End Region
#Region "  ├ Layer2 流程協調層"
    Private Async Function ScanMailsToGroupDictAsync(folderList As List(Of (Folder As Folder, fPath As String)), isExact As Boolean, progress As IProgress(Of ProgressReport),
                                                    cToken As CancellationToken) As Task(Of Dictionary(Of String, List(Of MailItemInfo)))
        ' ---------------------------------------------------------------
        ' ScanMailsToGroupDictAsync — 改用 GetBasicMailInfo L2.5 快取（Tab4/Tab5 共用）
        ' 2026/05/06 by Claude: 原版直接 GetTable COM 掃描已移除，改走 L2.5 快取代理層
        '   ① 記憶體快取命中 → 0 COM call（Tab4 掃過即共享）
        '   ② SSD lazy load → 僅需 snapshot 驗證
        '   ③ COM fallback → GetBasicMailInfoL3，結果存入快取
        '   MsgIDhash / SenderEmail 已整合至 MailItemInfo，BuildMailGroupKey 直接使用
        ' ---------------------------------------------------------------
        Dim groupDict As New Dictionary(Of String, List(Of MailItemInfo))(StringComparer.OrdinalIgnoreCase)
        Dim totalFolders As Integer = folderList.Count
        Dim totalProcessed As Integer = 0
        Dim swThrottle As Stopwatch = Stopwatch.StartNew()  ' by Claude Sonnet 4.6, 2026/06/07
        Dim swTotal As Stopwatch = Stopwatch.StartNew()     ' 2026/05/10 by Simon/Claude: 供 ETA 計算使用; refactored by Claude Sonnet 4.6, 2026/06/07

        ' 2026/05/11 by Simon/Claude: SSD 批次預讀，將 DB 中的 basic_maillist 一次拉入記憶體
        Await PreLoadBasicMailCacheAsync(folderList, cToken)

        For i As Integer = 0 To folderList.Count - 1
            Dim folder As Folder = folderList(i).Folder
            Dim fPath As String = folderList(i).fPath

            Try
                ' 透過 L2.5 取得（含快取），needTopic:=False (Tab5 不需要)
                Dim rows = Await GetBasicMailInfo(folder, needTopic:=False, cToken, fPath)
                For Each row In rows
                    Dim m As MailItemInfo = row.Mail
                    ' SenderEmail 優先，無則 fallback SenderName（與原版邏輯一致）
                    Dim senderKey As String = If(Not String.IsNullOrEmpty(m.SenderEmail), m.SenderEmail, m.SenderName)
                    Dim hashKey As String = BuildMailGroupKey(m.MsgIDhash, m.Subject, senderKey, m.Size, m.RcvTime, isExact)
                    If Not groupDict.ContainsKey(hashKey) Then groupDict(hashKey) = New List(Of MailItemInfo)()
                    groupDict(hashKey).Add(m)
                Next
            Catch ex As System.Exception
                _dbg("錯誤", $"{ExtractFolderName(fPath)}: {ex.Message}")
            End Try

            totalProcessed += 1
            Await SmartThrottle(swThrottle, cToken:=cToken, ThrottleFreq.Hii,
                                Sub()   ' 新版 (2026/05/10 by Simon/Claude: 加入 ETA 顯示，對齊 Tab3 做法)
                                    Dim eta = CalculateSpeedAndETA(totalFolders, totalProcessed, swTotal.Elapsed.TotalSeconds)
                                    progress?.Report(New ProgressReport With {.Message = $"掃描中: {totalProcessed}/{totalFolders} 個資料夾 ({eta.Speed:F0} 個/秒{eta.EtaString})"})
                                End Sub)
        Next
        Return groupDict
    End Function
    Private Function BuildMailGroupKey(msgId As String, subject As String, senderEmail As String, size As Long, recvTime As DateTime, isExact As Boolean) As String
        ' ---------------------------------------------------------------
        ' BuildMailGroupKey — 依模式產生分組用 hashKey
        '
        ' Exact 模式（兩層）:
        '   一階: 若 Message-ID 不為空 → 直接用 "MID:<msgid>" 作為 Key
        '         Message-ID 由發件 MTA 產生，跨 PST/client/relay 不變，最可靠
        '   二階: Message-ID 為空時的 fallback → 容忍版 Key
        '         - 主旨（去除 Re:/Fw: 前綴後）
        '         - 寄件者 email
        '         - 時間桶（10 分鐘為一格）：吸收 MTA relay / Outlook 規則移動造成的時間偏差
        '         - 大小桶（512 bytes 為一格）：吸收 Received: header relay 累積造成的大小偏差
        '
        ' Fuzzy 模式:
        '   主旨前 20 字（去前綴、去空白、大寫化）+ 大小（不分桶，精確）
        '   後續由 RenderLv5Group 的 Jaccard 二次過濾淘汰假陽性
        '   ── [預留] 日後加入 SimHash body 一階篩選時，此函數不需修改 ──
        '
        ' 2026/05/05 by Claude
        ' ---------------------------------------------------------------
        If isExact Then
            ' 一階: Message-ID 精確命中
            Dim mid As String = msgId?.Trim()
            If Not String.IsNullOrEmpty(mid) Then Return "MID:" & mid

            ' 二階: 容忍版 fallback
            Dim cleanSubj As String = GetCleanSubject(subject)
            Dim timeBucket As Long = recvTime.Ticks \ (TimeSpan.TicksPerMinute * 10L)   ' 時間容忍：10 分鐘一格（TimeSpan.TicksPerMinute * 10）
            Dim sizeBucket As Long = (size \ 512L) * 512L                               ' 大小容忍：512 bytes 一格，吸收 relay header ±200 bytes 的累積偏差
            Return $"FB:{cleanSubj}|{senderEmail}|{timeBucket}|{sizeBucket}"
        Else
            ' Fuzzy 模式：主旨前 20 字 + 精確大小
            Dim cleanSubj As String = GetCleanSubject(subject).Replace(" ", "").ToUpper()
            If cleanSubj.Length > 20 Then cleanSubj = cleanSubj.Substring(0, 20)
            Return $"FZ:{cleanSubj}|{size}"
        End If
    End Function
    Private Function RenderLv5Group(groupDict As Dictionary(Of String, List(Of MailItemInfo)), isExact As Boolean) As (GroupCount As Integer, MailCount As Integer)
        ' ---------------------------------------------------------------
        ' RenderLv5Group — 將分組結果渲染至 ListView5
        ' Exact 模式：直接顯示，simScore 固定 100%
        ' Fuzzy 模式：Jaccard 主旨相似度二次過濾（門檻 0.6）
        '
        ' ── [預留架構] SimHash 內文比對 ──
        '   日後將加入 SimHash + Hamming Distance ≤ 5 作為一階快速篩選，
        '   再以 Jaccard body similarity ≥ 0.8 做二階精細比對。
        '   屆時 isExactMode = False 的二階邏輯由此函數擴充， 呼叫端無需修改。
        '
        ' 2026/05/05 by Claude: 使用 AddRange 批次寫入，避免逐行 Add 觸發多次 UI 更新
        ' 2026/05/10 by Simon/Claude: 加入群組感知排序
        ' ---------------------------------------------------------------
        ' 2026/06/17 by Simon/Claude Opus 4.8: 原 [預留架構] SimHash 內文比對已實作於 Bt5_Click 的 Fuzzy 管線
        '   (S3 Precompute→S4 候選→S5 Jaccard 精算→S6 union-find 分群)，過濾在管線內完成，此處僅顯示。
        '   舊描述「主旨 Jaccard 0.6 / Hamming≤5 / body≥0.8」為舊設計，已汰除。
        '
        ' Exact 模式：直接顯示，simScore 固定 100%
        ' Fuzzy 模式：查 _lv5FuzzyScoreMap(S3→S6 預算的 body bigram Jaccard %)顯示，本函式不讀 body
        ' ---------------------------------------------------------------

        ListView5.Tag = groupDict ' by Gemini 3 Flash, 2026/05/06: 將資料來源掛載至 Tag 供 HandleLv5Delete 使用
        ListView5.BeginUpdate()
        ListView5.Items.Clear()
        _lv5OrphanedList.Clear()   ' 2026/06/18 by Simon/Claude Opus 4.8: Q4 重渲染(重搜/F5)時清掉孤兒紅字標記

        Dim groupID As Integer = 1 : Dim totalMails As Integer = 0
        Dim ascending As Boolean = (_lv5SortOrder = SortOrder.Ascending)

        ' ── Step 1: Jaccard 過濾並計算 simScores，產生有效群組清單 ──────
        ' 先過濾出有效群組 (含 simScores)，再排序，避免排序後 index 與 simScores 錯位
        Dim validGroupList As New List(Of (Key As String, Items As List(Of MailItemInfo), Scores As List(Of Double)))(groupDict.Count)
        For Each kvp In groupDict
            If kvp.Value.Count <= 1 Then Continue For

            Dim simScores As New List(Of Double)(kvp.Value.Count)
            ' Dim isValidGroup As Boolean = True   ' 2026/06/17 註: 新 Fuzzy 不在此過濾(S5 已過門檻)，此旗標目前恆為 True (贅留, 待 Simon 決定是否清)

            If isExact Then
                ' 2026/05/10 by Simon/Claude: Exact 模式不做 Jaccard，全部填 100%
                simScores.Add(1.0) ' 第一封基準
                For i As Integer = 1 To kvp.Value.Count - 1 : simScores.Add(1.0) : Next
            Else
                ' 2026/06/17 by Simon/Claude Opus 4.8: Fuzzy 改消費 S6 預算的 body bigram Jaccard scoreMap(對群代表的 %, 代表=100%)
                '   分數由 S3→S6 管線算好，此同步函式只查表顯示(嚴禁讀 body)；群組已過 S5 targetT 門檻 → 不在此重複過濾
                '   原「主旨 Jaccard 0.6 二次過濾」+「[預留] SimHash 內文比對」整段汰除 (body 模糊已是完整獨立分類, 見 memory D8)
                For i As Integer = 0 To kvp.Value.Count - 1
                    simScores.Add(If(_lv5FuzzyScoreMap Is Nothing, 1.0, _lv5FuzzyScoreMap.GetValueOrDefault(kvp.Value(i).EntryID, 1.0)))
                Next
            End If
            ' If Not isValidGroup Then Continue For
            validGroupList.Add((kvp.Key, kvp.Value, simScores))
        Next

        ' ── Step 2: 依欄位對群組排序 (以群組內極值作為代表排序鍵) ──────
        ' by Gemini 3.0 flash, 2026/05/10: 修正群組間排序時使用 First() 造成的隨機代表值問題。
        ' 因為組內排序在 Step 3 才做，First() 取到的不一定是最大/最小，導致群組交界處出現視覺上的「大小顛倒」，讓使用者誤以為是依文字排序。
        ' 改用 Min()/Max() 來取得群組內真正的極值，使群組間的排序邏輯與群組內一致。
        Dim sortedGroups As IEnumerable(Of (Key As String, Items As List(Of MailItemInfo), Scores As List(Of Double)))
        Select Case _lv5LastSortColumn
            Case 0 : sortedGroups = If(ascending, validGroupList.OrderBy(Function(g) g.Items.Min(Function(m) m.Subject)),
                                                  validGroupList.OrderByDescending(Function(g) g.Items.Max(Function(m) m.Subject)))     ' 主旨
            Case 1 : sortedGroups = If(ascending, validGroupList.OrderBy(Function(g) g.Items.Min(Function(m) m.Size)),
                                                  validGroupList.OrderByDescending(Function(g) g.Items.Max(Function(m) m.Size)))        ' 郵件大小
            Case 2 : sortedGroups = If(ascending, validGroupList.OrderBy(Function(g) g.Items.Min(Function(m) m.RcvTime)),
                                                  validGroupList.OrderByDescending(Function(g) g.Items.Max(Function(m) m.RcvTime)))     ' 收到日期
            Case 3 : sortedGroups = If(ascending, validGroupList.OrderBy(Function(g) g.Items.Min(Function(m) m.SenderName)),
                                                  validGroupList.OrderByDescending(Function(g) g.Items.Max(Function(m) m.SenderName)))  ' 寄件者
            Case 4, 5 : sortedGroups = validGroupList   ' 2026/05/10: 群組/相似度欄不排序，保持原序
            Case Else : sortedGroups = validGroupList
        End Select

        ' ── Step 3: 逐群組渲染 ──────────────────────────────────────────
        For Each grp In sortedGroups
            Dim groupColor As Color = If(groupID Mod 2 = 0, Color.FromArgb(240, 248, 255), Color.White)

            ' 組內排序：將 items 與 simScores 配對後一起排，確保相似度欄位不錯位
            Dim pairs = grp.Items.Zip(grp.Scores, Function(m, s) (Mail:=m, Score:=s)).ToList()
            Dim sortedPairs As IEnumerable(Of (Mail As MailItemInfo, Score As Double))
            Select Case _lv5LastSortColumn
                Case 0 : sortedPairs = If(ascending, pairs.OrderBy(Function(p) p.Mail.Subject), pairs.OrderByDescending(Function(p) p.Mail.Subject))
                Case 1 : sortedPairs = If(ascending, pairs.OrderBy(Function(p) p.Mail.Size), pairs.OrderByDescending(Function(p) p.Mail.Size))
                Case 2 : sortedPairs = If(ascending, pairs.OrderBy(Function(p) p.Mail.RcvTime), pairs.OrderByDescending(Function(p) p.Mail.RcvTime))
                Case 3 : sortedPairs = If(ascending, pairs.OrderBy(Function(p) p.Mail.SenderName), pairs.OrderByDescending(Function(p) p.Mail.SenderName))
                Case Else : sortedPairs = pairs ' 4(群組), 5(相似), 6(EntryID) 均保持原序
            End Select

            ' ── 建立 ListViewItem 清單，一次 AddRange ──
            Dim lvItems As New List(Of ListViewItem)(grp.Items.Count)
            For Each pair In sortedPairs
                Dim m As MailItemInfo = pair.Mail
                Dim simText As String = $"{CInt(pair.Score * 100)}%"
                lvItems.Add(New ListViewItem({m.Subject, (m.Size \ 1024L).ToString("N0") & "KB",
                                              m.RcvTime.ToString("yyyy/MM/dd HH:mm:ss"),
                                              m.SenderName, "G" & groupID.ToString(), simText, m.EntryID}) With {.BackColor = groupColor, .Tag = m})
                totalMails += 1
            Next
            ListView5.Items.AddRange(lvItems.ToArray())
            groupID += 1
        Next

        ListView5.EndUpdate()
        Return (groupID - 1, totalMails)

    End Function
    Private Sub HandleLv5Delete(lv As ListView)
        ''' <summary>
        ''' by Gemini 3 Flash, 2026/05/06: 處理重複郵件刪除與 UI 連動，行為仿照 HandleLv4Delete
        ''' 2026/06/18 by Simon/Claude Opus 4.8: Q4 改「原地刪除」— 不再整表重畫(避免配對信消失、游標跳回頂端)。
        '''   刪除後：被刪列直接移除、同群其餘「配對信」加入 _lv5OrphanedList 由 DrawSubItem 標紅、游標留在被刪列下一列。
        '''   (Lv5 為 OwnerDraw，文字色由 DrawSubItem 決定，不能靠 item.ForeColor)
        '''   殘留的單封孤兒群不在此清除；待使用者按 F5(RefreshAllLvItems 末端 RenderLv5Group 丟棄 Count<=1 群並清紅字) 或重新搜尋時才整理。
        ''' </summary>
        _dbg("開始")
        Dim selCount As Integer = lv.SelectedItems.Count
        If selCount = 0 Then Return

        If MessageBox.Show($"確定要將選中的 {selCount} 封郵件移到「刪除郵件」資料夾嗎？", "確認刪除",
                           MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then

            ' 取得快取的資料字典
            Dim groupDict As Dictionary(Of String, List(Of MailItemInfo)) = _lv5PrevGroupResults
            If groupDict Is Nothing Then Return

            ' 2026/06/18 by Simon/Claude Opus 4.8: 刪除前固定「待移除的列」與游標錨點(最小選取索引)，供原地移除與游標還原
            Dim selectedItems As ListViewItem() = lv.SelectedItems.Cast(Of ListViewItem)().ToArray()
            Dim anchorIndex As Integer = lv.SelectedIndices.Cast(Of Integer)().Min()

            ' 2026/5/11 simon: 收集受影響的資料夾路徑，供後續清理快取與DB使用
            Dim affectedPaths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            ' 收集選中項目的 EntryID 並從資料源中移除
            Dim entryIDs As New List(Of String)(selCount)
            For Each item As ListViewItem In selectedItems
                If TypeOf item.Tag Is MailItemInfo Then
                    Dim info = DirectCast(item.Tag, MailItemInfo)
                    entryIDs.Add(info.EntryID)

                    ' 從 groupDict 中移除該封信 (遍歷所有群組尋找)
                    For Each kvp In groupDict
                        ' 找到並移除後，如果該群組只剩 1 封或 0 封，在重複郵件邏輯中視為不再重複，可選擇保留或由渲染器過濾
                        If kvp.Value.RemoveAll(Function(m) m.EntryID = info.EntryID) > 0 Then
                            ' 2026/06/18 by Simon/Claude Opus 4.8: 移除後該群其餘成員即「失去配對的另一封」→ 加入孤兒集合，由 DrawSubItem 標紅
                            For Each other In kvp.Value : _lv5OrphanedList.Add(other.EntryID) : Next
                            Exit For
                        End If
                    Next

                    ' 2026/5/11 simon: 收集受影響的資料夾路徑，供後續清理快取與DB使用
                    If Not String.IsNullOrEmpty(info.FolderPath) Then affectedPaths.Add(info.FolderPath)
                End If
            Next

            If entryIDs.Count > 0 Then
                For Each fPath In affectedPaths
                    InvalidateBasicMailCache(fPath)     ' 2026/5/11 by Simon: 刪除後手動清理快取資料，避免殘留已刪除郵件的資訊
                    DbDeleteBasicMailInfoByPath(fPath)  ' 2026/5/11 by Simon: 刪除後手動清理 DB 資料，避免殘留已刪除郵件的資訊
                Next
                MoveMailsToRecycle(entryIDs)            ' 實體移動

                ' 2026/06/18 by Simon/Claude Opus 4.8: 原地更新 UI(取代整表 RenderLv5Group)，保留捲動位置與其餘列
                lv.BeginUpdate()
                For Each it In selectedItems : lv.Items.Remove(it) : Next          ' ① 移除被刪列，其餘列原地不動
                If lv.Items.Count > 0 Then                                         ' ② 游標留在原位(被刪列下一列)，不跳頂端
                    Dim newIdx As Integer = Math.Min(anchorIndex, lv.Items.Count - 1)
                    lv.Items(newIdx).Selected = True : lv.Items(newIdx).Focused = True : lv.Items(newIdx).EnsureVisible()
                End If
                lv.Invalidate()                                                    ' ③ 強制重繪→孤兒信經 DrawSubItem 上紅字(OwnerDraw 資料未變不會自動重畫)
                lv.EndUpdate()

                PgrsBar2.Text = $"已移動 {selCount} 封郵件至刪除郵件資料夾"
            End If
        End If
        _dbg("結束")
    End Sub
#End Region
#Region "  ├ Fuzzy 模糊比對專用區塊 (SimHash + bigram Jaccard)"
    ' 原有: Private Const MIN_BIGRAM_FOR_FUZZY As Integer = 5   ' 內文 bigram 少於此值(極短/空白信)不納入模糊比對，避免無意義的雜訊群
    ' Q1-A 2026/06/18 by Simon/Claude Opus 4.8:
    '   MIN_BIGRAM_FOR_FUZZY          : 由 5 提高至 16 (保守起點, 實測再調)。內文 distinct bigram 少於此值(極短/純符號牆 >>>/空白信) 整封不納入模糊比對。屬「逐封自身長度」下限，與 Q1-C 互補
    '   MIN_SHARED_BIGRAM_FOR_FUZZY   : S5 最終閘門「共有內容量」下限 |A∩B|>=此值。Jaccard 比例對規模無感 (短信剛好全中→100% 但實質空洞)，另加交集絕對數把關。屬「逐對共有量」下限，與 Q1-A 互補
    ' Private Const MIN_BIGRAM_FOR_FUZZY As Integer = 16
    ' Private Const MIN_SHARED_BIGRAM_FOR_FUZZY As Integer = MIN_BIGRAM_FOR_FUZZY * 2
    ' 改後:
    '   Q1 連動滑桿 2026/06/18 by Simon/Claude Opus 4.8:
    '       原 MIN_SHARED = MIN_BIGRAM*2 固定值，改為隨檔位連動(見下方 MinSharedBigramFor)。MIN_BIGRAM_FOR_FUZZY 由 16→25 作 1 倍基準(低檔)。
    '       短信全中假陽性的成因是「共有量不足」，故連動打在共有量(C)上；S4 池子閘與 S5 最終閘共用 MinSharedBigramFor(targetT)，無死區(S4 只放進能過 S5 的信)。
    '       基準 25 為保守起點，看 _dbg("S5閘門") yield 再微調。
    Private Const MIN_BIGRAM_FOR_FUZZY As Integer = 45
    ' 2026/6/19 by Simon/Claude Opus 4.8
    ' MIN_BIGRAM_FOR_FUZZY 的閾值控制: 把過關結果裡 inter 最小的那幾群打開看——
    '   如果 [25,50) 那段大多是「請查收附件謝謝」這種客套話 → 40，甚至可往 50 推
    '   如果那段藏著不同來源的真同文（只是短）→ 退回 30。

    ' 2026/06/17 by Simon/Claude Opus 4.8: Tab5 Fuzzy 相似度檔位表。TrackBar1.Value 1~5 → 低/中/高/極高/完全一致。
    '   targetT 同時驅動 size 視窗(1/T)、Hamming 一階(HammingThresholdFor)、S5 最終 Jaccard 門檻(s >= targetT)，一個旋鈕全管。
    Private Shared ReadOnly _fuzzyTierT As Double() = New Double() {0, 0.87, 0.92, 0.95, 0.98, 1}   ' index 0 佔位(trackbar 從 1 起)
    Private Shared ReadOnly _fuzzyTierName As String() = New String() {"", "低", "中", "高", "極高", "完全一致"}

    Private Async Function PreComputeFuzzySimHashAsync(mails As List(Of MailItemInfo), progress As IProgress(Of ProgressReport), cToken As CancellationToken) As Task
        ' 對「_cacheSimHash 沒有的」郵件讀 body 算 simhash+bigram_count，寫回獨立 db 與記憶體快取。已算過的(暖快取)直接跳過。
        '   ※ 故意走 GetMailBodyL3(直接 L3)而非 GetMailBody：build pass 一次掃數萬封，若每封都進 _cacheMailBody 會撐爆記憶體；
        '     simhash 算完 body 即丟。候選的 body 才在 S5 經 GetMailBody 讀(只快取那少數幾封)。若你偏好走 L2.5 改這一行即可。

        ' 2026/06/23 by Simon/Claude Opus 4.8:
        ' ※ 原直呼 GetMailBodyL3 繞過 L2.5; 改 skipCache 後路由統一, RDO 提速涵蓋此熱路徑。
        '    走 L2.5 GetMailBody 但帶 skipCache:=True：build pass 一次掃數萬封, skipCache 跳過 _cacheMailBody 讀寫避免撐爆記憶體; 同時仍經 L2.5 分派吃到 _rdo2 store-scoped RDO 提速。

        LoadDbMail()
        Dim todo = mails.Where(Function(m) Not _cacheSimHash.ContainsKey(m.EntryID)).ToList()
        If todo.Count = 0 Then Return

        Dim totalBodyChars As Long = 0   ' 估算mailbody總容量用探針(算完可移除): 累計 body 字元數，估算全快取記憶體footprint
        ' 2026/06/25 by Gemini 3.1 Pro: 將 Batch Size 提升至 3000，大幅降低磁碟寫入次數與 I/O 停頓
        Dim batch As New List(Of (EntryID As String, SimHash As Long, BigramCount As Integer))(3072)
        Dim swEta As Stopwatch = Stopwatch.StartNew()  ' 2026/06/17 by Simon/Claude: 供進度速度與 ETA 計算
        Dim swThrottle As Stopwatch = Stopwatch.StartNew() ' 2026/06/25 by Gemini 3.1 Pro: 用於雙重節流的時間閘門
        For i As Integer = 0 To todo.Count - 1
            cToken.ThrowIfCancellationRequested()
            Dim id As String = todo(i).EntryID
            ' 2026/06/18 估算mailbody總容量用探針(todo: 算完可移除)
            ' Dim setB = BuildBigramSet(GetMailBodyL3(id))            ' COM 讀取(昂貴) → 拆 bigram 集合

            ' todo: 這裡真的需要skipCache:=True嗎? 區區幾百MB的純文字mailbody快取會撐爆記憶體??
            Dim body As String = GetMailBody(id, todo(i).FolderPath, skipCache:=True)   ' 2026/06/23 by Simon/Claude Opus 4.8:
            totalBodyChars += body.Length
            Dim setB = BuildBigramSet(body)
            Dim sh As Long = ComputeSimHashFromSet(setB)
            _cacheSimHash(id) = (sh, setB.Count)                            ' 立即進記憶體(本 session 即可用，即使尚未 flush)

            ' 2026/06/25 by Gemini 3.1 Pro: 批次寫入門檻提高至 3000
            batch.Add((id, sh, setB.Count))
            If batch.Count >= 3000 Then SaveDbMail(batch) : batch.Clear()   ' 分批 flush，兼作斷點(中斷後可續算)

            ' 2026/06/25 by Gemini 3.1 Pro: 雙重閘門優化。外層擋掉 63/64 的檢查，內層確保時間到了才更新 UI (消除超高頻刷新浪費)。
            If (i And 63) = 0 AndAlso swThrottle.ElapsedMilliseconds >= ThrottleFreq.Mid Then
                Dim eta = CalculateSpeedAndETA(todo.Count, i + 1, swEta.Elapsed.TotalSeconds)   ' 2026/06/17 by Simon/Claude: 加入速度與 ETA 顯示，對齊 Tab3/Tab4 做法
                progress?.Report(New ProgressReport With {.Message = $"計算內文指紋: {i + 1}/{todo.Count} ({eta.Speed:F0} 個/秒{eta.EtaString})"})
                ' 2026/06/25 by Gemini 3.1 Pro: 不帶入 onThrottled 委派，避免產生 Closure 記憶體配置，純享受其 Delay 與 OCE
                Await SmartThrottle(swThrottle, cToken, ThrottleFreq.Hii)
            End If
        Next
        ' 2026/06/18 估算mailbody總容量用探針(算完可移除)
        _dbg("[Probe][MailBodySize]]", $"S3 讀 {todo.Count} 封, body 累計 {totalBodyChars:N0} 字元 ≈ {totalBodyChars * 2 / 1048576:F0} MB(純UTF-16) + 約 {todo.Count * 26 / 1048576:F0} MB(字串物件開銷)。此即全進 _cacheMailBody 的概估量")

        If batch.Count > 0 Then SaveDbMail(batch)
    End Function
    Private Async Function GenerateFuzzyCandidatePairs(mails As List(Of MailItemInfo), targetT As Double, cToken As CancellationToken) As Task(Of List(Of (A As MailItemInfo, B As MailItemInfo)))
        ' size 1/T 滑動視窗收斂 O(n²) + Hamming 一階篩。純 CPU、無 COM → 放 Task.Run 不凍 UI。
        Dim hThr As Integer = HammingThresholdFor(targetT)
        Dim maxRatio As Double = 1.0 / targetT
        Dim minBigram As Integer = MinSharedBigramFor(targetT)   ' Q1 連動 2026/06/18 by Simon/Claude: S4 池子閘改用檔位值(同 S5，避免放進注定死在 S5 的信)

        Return Await Task.Run(
            Function() As List(Of (A As MailItemInfo, B As MailItemInfo))
                ' 取出有指紋且 bigram 數達標者，依 bigram_count 升冪排序(作為 size 視窗的排序鍵)
                Dim items = mails.
                    Where(Function(m) _cacheSimHash.ContainsKey(m.EntryID) AndAlso _cacheSimHash(m.EntryID).BigramCount >= minBigram).
                    Select(Function(m) (Mail:=m, SH:=_cacheSimHash(m.EntryID).SimHash, Cnt:=_cacheSimHash(m.EntryID).BigramCount)).
                    OrderBy(Function(x) x.Cnt).ToList()

                Dim result As New List(Of (A As MailItemInfo, B As MailItemInfo))()
                For i As Integer = 0 To items.Count - 1
                    If (i And 1023) = 0 Then cToken.ThrowIfCancellationRequested()
                    For j As Integer = i + 1 To items.Count - 1
                        If items(j).Cnt > items(i).Cnt * maxRatio Then Exit For   ' size 1/T 上界：升冪→超界後 j 全部出局，收尾視窗
                        If GetHammingDistance(items(i).SH, items(j).SH) <= hThr Then result.Add((items(i).Mail, items(j).Mail))
                    Next
                Next
                Return result
            End Function, cToken)
    End Function
    Private Async Function FilterCandidatesByJaccardAsync(candidates As List(Of (A As MailItemInfo, B As MailItemInfo)), targetT As Double, progress As IProgress(Of ProgressReport), cToken As CancellationToken) _
            As Task(Of (Pairs As List(Of (A As MailItemInfo, B As MailItemInfo, Score As Double)), Sets As Dictionary(Of String, HashSet(Of Integer))))
        ' 只對通過 S4 的少數候選讀 body、建 bigram 集合、算精確 bigram Jaccard。集合回傳給 S6 算群代表分數，免重讀。
        Dim sets As New Dictionary(Of String, HashSet(Of Integer))()
        Dim uniqueIds = candidates.SelectMany(Function(p) New String() {p.A.EntryID, p.B.EntryID}).Distinct().ToList()

        ' Phase1: 候選 body 讀取(COM, UI 執行緒, 少量) → bigram 集合。這裡走 GetMailBody(會快取這少數幾封)
        Dim swEta As Stopwatch = Stopwatch.StartNew()  ' 2026/06/17 by Simon/Claude: 供進度速度與 ETA 計算
        For k As Integer = 0 To uniqueIds.Count - 1
            cToken.ThrowIfCancellationRequested()
            Dim id = uniqueIds(k)
            If Not sets.ContainsKey(id) Then sets(id) = BuildBigramSet(GetMailBody(id))
            If (k And 15) = 0 Then
                ' 2026/06/17 by Simon/Claude: 加入速度與 ETA 顯示，對齊 Tab3/Tab4 做法
                Dim eta = CalculateSpeedAndETA(uniqueIds.Count, k + 1, swEta.Elapsed.TotalSeconds)
                progress?.Report(New ProgressReport With {.Message = $"開始過濾候選內文: {k + 1}/{uniqueIds.Count} ({eta.Speed:F0} 個/秒{eta.EtaString})"})
                Await Task.Delay(1, cToken)
            End If
        Next

        ' Phase2: Jaccard 精算(純 CPU)。候選少，序列即可；日後量大可改 Parallel.For(對齊 Tab4)
        Dim minShared As Integer = MinSharedBigramFor(targetT)   ' Q1 連動 2026/06/18 by Simon/Claude: S5 共有量下限改用檔位值
        Dim pairs = Await Task.Run(
            Function() As List(Of (A As MailItemInfo, B As MailItemInfo, Score As Double))
                Dim r As New List(Of (A As MailItemInfo, B As MailItemInfo, Score As Double))()
                For Each p In candidates
                    ' Q1-C 2026/06/18 by Simon/Claude Opus 4.8: 先取交集絕對數，不足門檻直接淘汰(擋短信比例 100% 假陽性)；達標再由 inter 導出 Jaccard，免重算交集
                    Dim setA = sets(p.A.EntryID), setB = sets(p.B.EntryID)
                    Dim inter As Integer = BigramIntersectionCount(setA, setB)
                    If inter < minShared Then Continue For

                    Dim union As Integer = setA.Count + setB.Count - inter
                    Dim s As Double = If(union = 0, 0.0, inter / union)   ' size 1/T 界線 S4 已保證，此處不重複早退
                    If s >= targetT Then r.Add((p.A, p.B, s))
                Next
                Return r
            End Function, cToken)
        Return (pairs, sets)
    End Function
    Private Function BuildFuzzyGroups(similar As List(Of (A As MailItemInfo, B As MailItemInfo, Score As Double)), sets As Dictionary(Of String, HashSet(Of Integer))) _
            As (GroupDict As Dictionary(Of String, List(Of MailItemInfo)), ScoreMap As Dictionary(Of String, Double))
        ' 通過 Jaccard 門檻的配對 → union-find 連通分量 → G1,G2…；每群選 bigram_count 最大者為代表，每封顯示「對代表的 bigram Jaccard %」。
        Dim parent As New Dictionary(Of String, String)()
        Dim mailById As New Dictionary(Of String, MailItemInfo)()
        For Each p In similar
            For Each m In {p.A, p.B}
                If Not parent.ContainsKey(m.EntryID) Then parent(m.EntryID) = m.EntryID : mailById(m.EntryID) = m
            Next
        Next
        For Each p In similar : Uf_Union(parent, p.A.EntryID, p.B.EntryID) : Next

        ' 依 root 收群
        Dim byRoot As New Dictionary(Of String, List(Of MailItemInfo))()
        For Each id In parent.Keys
            Dim root = Uf_Find(parent, id)
            If Not byRoot.ContainsKey(root) Then byRoot(root) = New List(Of MailItemInfo)()
            byRoot(root).Add(mailById(id))
        Next

        Dim groupDict As New Dictionary(Of String, List(Of MailItemInfo))()
        Dim scoreMap As New Dictionary(Of String, Double)()
        Dim gid As Integer = 1
        For Each grp In byRoot.Values
            If grp.Count < 2 Then Continue For
            ' 代表 = bigram_count 最大(內容最完整)那封；chain 成員對代表可能偏低(union-find 傳遞性)，照實顯示
            Dim rep = grp.OrderByDescending(Function(m) sets(m.EntryID).Count).First()
            Dim repSet = sets(rep.EntryID)
            For Each m In grp : scoreMap(m.EntryID) = BigramJaccardSimilarity(sets(m.EntryID), repSet) : Next
            groupDict("G" & gid.ToString()) = grp : gid += 1
        Next
        Return (groupDict, scoreMap)
    End Function

    ' 2026/06/17 by Simon/Claude Opus 4.8: SimHash / bigram Jaccard 三件組 (取代 ComputeSimHash_core.vb 初版)
    '   ① 逐 bigram 的 64-bit hash 改用 System.IO.Hashing.XxHash64.HashToUInt64：
    '      庫內部處理溢位 → 不需開專案 RemoveIntegerChecks (本專案預設溢位檢查為開, 手刻 FNV/splitmix 的 ULong 乘法會拋 OverflowException)；
    '      XxHash64 雪崩本就強 → 不再需要 splitmix64。餵重用的 4-byte buffer, 零堆積配置。Simon 原即指定 XxHash64。
    '   ② SimHash 改以「唯一 bigram 集合」投票 (非逐次出現)，與最終 set-based bigram Jaccard 語意一致
    '      → Hamming 成為 Jaccard 更準的代理；且 simhash 與 bigram_count(=集合大小) 同一趟算出, 不重複建集合。
    '   ③ bigram 用 packed Integer ((前字<<16) Or 後字)：精確、零碰撞、零配置、免 hash；<< 與 Or 不觸發溢位檢查。
    Private Function GetHammingDistance(hash1 As Long, hash2 As Long) As Integer
        ' 利用位元互斥或 (XOR) 找出不同的 bit，再計算有幾個 1 (PopCount)
        ' 原本：
        ' Return System.Numerics.BitOperations.PopCount(CULng(hash1 Xor hash2))
        ' 改為：
        ' 2026/06/17 by Simon/Claude Opus 4.8: simhash 含 bit63 時 XOR 結果為負 Long，CULng(負Long) 在溢位檢查開啟下會拋 OverflowException(超出 ULong 範圍)。
        '   先遮掉 bit63 安全轉 ULong，再依原 bit63 補回 → 位元保留、不觸發溢位檢查；PopCount 只數 bit 數，結果不變。
        Dim x As Long = hash1 Xor hash2
        Return System.Numerics.BitOperations.PopCount(CULng(x And &H7FFFFFFFFFFFFFFFL) Or If(x < 0L, &H8000000000000000UL, 0UL))
    End Function
    Private Function BuildBigramSet(text As String) As HashSet(Of Integer)
        ' 字串 → 相鄰兩字打包成 32-bit 的唯一集合，供 SimHash 投票 / bigram_count / Jaccard 共用
        Dim setB As New HashSet(Of Integer)()
        If String.IsNullOrEmpty(text) OrElse text.Length < 2 Then Return setB

        ' 2026/06/17 by Simon/Claude Opus 4.8: VB.NET 不支援 ReadOnlySpan Item 索引子(ByRef return)→ 直接索引 String, 等價且零配置
        For i As Integer = 0 To text.Length - 2 : setB.Add((AscW(text(i)) << 16) Or AscW(text(i + 1))) : Next
        Return setB
    End Function
    Private Function ComputeSimHash(text As String) As Long
        ' 便利進入點: 字串直接算 SimHash。build pass 請改呼叫 BuildBigramSet 一次, 同時取 simhash 與 .Count, 避免重建集合
        Return ComputeSimHashFromSet(BuildBigramSet(text))
    End Function
    Private Function ComputeSimHashFromSet(bigramSet As HashSet(Of Integer)) As Long
        ' 對唯一 bigram 集合算 64-bit SimHash (Charikar 2002, 集合式)
        '   每 bigram 以 XxHash64 取 64-bit hash → 對 64 個累積器投票 (bit=1 加, 0 減)；正者設位組成指紋
        '   比對用既有 GetHammingDistance (XOR+PopCount)；相似度 ≈ 1 − Hamming/64
        If bigramSet.Count = 0 Then Return 0L

        Dim v(63) As Integer
        Dim buf(3) As Byte                                  ' 重用 4-byte buffer (packed bigram), 零 per-bigram 配置
        For Each bg As Integer In bigramSet
            System.Buffers.Binary.BinaryPrimitives.WriteInt32LittleEndian(buf.AsSpan(), bg)
            Dim h As ULong = System.IO.Hashing.XxHash64.HashToUInt64(buf.AsSpan())
            For bit As Integer = 0 To 63 : v(bit) += If((h And (1UL << bit)) <> 0UL, 1, -1) : Next
        Next

        Dim fp As ULong = 0UL
        For bit As Integer = 0 To 63
            If v(bit) > 0 Then fp = fp Or (1UL << bit)
        Next
        ' 原本：
        ' Return CLng(fp)                                     ' ULong→Long 位元樣式不變, 供 SQLite INTEGER / GetHammingDistance 使用
        ' 改為：
        ' 2026/06/17 by Simon/Claude Opus 4.8: VB 的 CLng(ULong) 為「檢查式」轉換，fp 設了 bit63(≥2^63) 會超過 Long.MaxValue → 拋 OverflowException(本專案溢位檢查為開)。
        '   故先遮掉 bit63 安全轉 Long，再依原 bit63 補上 Long.MinValue(=bit63)，等同 C# unchecked((long)fp)：位元樣式不變、零配置、不觸發溢位檢查。比對端 GetHammingDistance 以 CULng 還原即一致。
        Return CLng(fp And &H7FFFFFFFFFFFFFFFUL) Or If((fp And &H8000000000000000UL) <> 0UL, Long.MinValue, 0L)

    End Function
    Private Function BigramJaccardSimilarity(setA As HashSet(Of Integer), setB As HashSet(Of Integer)) As Double
        ' set-based bigram Jaccard: |A∩B|/|A∪B|；D1 定案, 取代字元集 JaccardSimilarity 作 Fuzzy 內文相似度顯示值
        ' 2026/06/17 by Simon/Claude Opus 4.8: 與 SimHash 特徵(bigram) 及 size 1/T 界線數學一致
        If setA.Count = 0 AndAlso setB.Count = 0 Then Return 1.0
        If setA.Count = 0 OrElse setB.Count = 0 Then Return 0.0
        Dim small = setA : Dim big = setB                   ' 走較小集合計交集 (沿用 JaccardSimilarity 既有優化手法)
        If small.Count > big.Count Then small = setB : big = setA
        Dim inter As Integer = 0
        For Each bg In small : If big.Contains(bg) Then inter += 1
        Next
        Dim union As Integer = setA.Count + setB.Count - inter
        If union = 0 Then Return 0.0
        Return inter / union
    End Function
    Private Function BigramIntersectionCount(setA As HashSet(Of Integer), setB As HashSet(Of Integer)) As Integer
        ' Q1-C 2026/06/18 by Simon/Claude Opus 4.8: 回傳兩 bigram 集合交集絕對數 |A∩B|，供 S5「共有內容量下限」閘門。
        '   與 BigramJaccardSimilarity 同走較小集合計交集；分出此函式讓 S5 可由 inter 同時導出 Jaccard，避免重算交集
        If setA.Count = 0 OrElse setB.Count = 0 Then Return 0
        Dim small = setA : Dim big = setB                   ' 走較小集合計交集(沿用既有優化手法)
        If small.Count > big.Count Then small = setB : big = setA
        Dim inter As Integer = 0
        For Each bg In small : If big.Contains(bg) Then inter += 1
        Next
        Return inter
    End Function

    ' 輔助函數
    Private Function GetFuzzyTargetT() As Double
        Return _fuzzyTierT(Math.Clamp(TrackBar1.Value, 1, 5))   ' TrackBar1.Value(1~5)→targetT，越界夾住保險
    End Function
    ' 2026/06/17 by Simon/Claude Opus 4.8: D6 Hamming 一階門檻「起始值」表
    '   ── 微調原因：SimHash Hamming 對應的是特徵向量夾角(cosine)、非 Jaccard，且 64-bit 量化有噪音，無法用公式定死。
    '   ── 微調方式：寧鬆勿緊(誤選 OK, 後面還有 Jaccard 把關 >> 但漏掉真重複就不 OK 了)。
    '                   下方 _dbg 探針會記錄「Hamming 過關配對數 vs Jaccard 過關數」，實際上機看 yield rate：太低就收緊、疑似漏抓就放寬。v1.1 依實測定案。Jaccard(S5) 才是準確閘門。
    Private Function HammingThresholdFor(targetT As Double) As Integer
        ' Hamming 門檻對應表 (64-bit SimHash):
        ' Hamming   targetT 你的 E[d]=64(1−T)	SD=√(64·T(1−T))	    E+~2SD	Claude設定門檻
        ' 0 bit     0.999	    0.064	            0.25	        ~0.6	     2
        ' 1 bit     0.985	    1.28	            1.12	        ~3.5	     4
        ' 2 bit     0.953	    3.20                1.74	        ~6.7	     7
        ' 4 bit     0.922	    5.12	            2.17	        ~9.5	    10
        ' 8 bit     0.875	    8.00                2.65	        ~13.3	    14
        If targetT >= 0.99 Then Return 2   ' 2026/06/17 by Simon/Claude Opus 4.8: 完全一致檔。仍寧鬆勿緊(近乎相同內文 SimHash Hamming 多落 0~2)，最終由 S5 Jaccard>0.999 收斂
        If targetT >= 0.98 Then Return 4
        If targetT >= 0.95 Then Return 7
        If targetT >= 0.92 Then Return 10
        Return 14   ' 0.87(低檔)；0.9275(中)會先命中上一行的 >=0.92→10
    End Function
    ' Q1 連動滑桿 2026/06/18 by Simon/Claude Opus 4.8: 共有內容量下限的檔位連動表(對齊 HammingThresholdFor/GetFuzzyTargetT 邊界)
    '   越嚴(高 T)→要求兩封共有越多真實內容才算重複。倍率 低1/中2/高3/極高4/完全一致5，乘上基準 MIN_BIGRAM_FOR_FUZZY(=25)。
    Private Function MinSharedBigramFor(targetT As Double) As Integer
        If targetT >= 0.99 Then Return MIN_BIGRAM_FOR_FUZZY * 5   ' 完全一致 125
        If targetT >= 0.98 Then Return MIN_BIGRAM_FOR_FUZZY * 4   ' 極高 100
        If targetT >= 0.95 Then Return MIN_BIGRAM_FOR_FUZZY * 3   ' 高 75
        If targetT >= 0.92 Then Return MIN_BIGRAM_FOR_FUZZY * 2   ' 中 50
        Return MIN_BIGRAM_FOR_FUZZY                               ' 低 0.87 (1×) 25
    End Function
    Private Function Uf_Find(parent As Dictionary(Of String, String), x As String) As String
        Dim root As String = x
        While parent(root) <> root : root = parent(root) : End While
        While parent(x) <> root : Dim nxt = parent(x) : parent(x) = root : x = nxt : End While   ' 路徑壓縮
        Return root
    End Function
    Private Sub Uf_Union(parent As Dictionary(Of String, String), a As String, b As String)
        Dim ra = Uf_Find(parent, a), rb = Uf_Find(parent, b)
        If ra <> rb Then parent(ra) = rb
    End Sub
#End Region
#Region "  └ Tab3/Tab4/Tab5 共用事件函數"
    ' by Gemini 3.1 Pro, 2026/04/21: 邏輯整合 (Tab3/Tab4/Tab5)，完整統一行為。
    ' 理由: Tab3 與 Tab4 的 ListView 皆為「搜尋結果」，行為高度一致 (Enter/雙擊/連動與路徑顯示)。
    ' 整合後可減少冗餘代碼，並確保滑鼠與熱鍵行為絕對一致。
    ' --------------------------------------------------------------
    Private Sub InitLv3Lv4Lv5ContextMenu()
        ' 2026/06/15 by Simon/Claude Opus 4.8: Lv3/4/5 共用右鍵選單；冪等，重複呼叫只建一次 (確保三個 LV 共用同一實例)
        If ctxMenuLv3Lv4Lv5 IsNot Nothing Then Return Else ctxMenuLv3Lv4Lv5 = New ContextMenuStrip()

        Dim mnuRefresh As New ToolStripMenuItem("重刷選取項目(&R)")
        Dim mnuDelete As New ToolStripMenuItem("刪除選取項目(&D)")
        ctxMenuLv3Lv4Lv5.Items.Add(mnuRefresh)
        ctxMenuLv3Lv4Lv5.Items.Add(mnuDelete)

        ' 2026/06/21 by Simon/Claude Opus 4.8: 共用選單新增「刪除選取項目」，依 SourceControl 分派至各 LV 既有刪除處理器
        AddHandler mnuDelete.Click, Sub(sender, e)
                                        Dim lv = TryCast(ctxMenuLv3Lv4Lv5.SourceControl, ListView)
                                        If lv Is ListView3 Then HandleLv3Delete(lv)
                                        If lv Is Listview4 Then HandleLv4Delete(lv)
                                        If lv Is ListView5 Then HandleLv5Delete(lv)
                                    End Sub

        ' 2026/06/14 by Simon/Claude Opus 4.8: 右鍵選單「強制刷新選取的郵件」點擊
        AddHandler mnuRefresh.Click, Async Sub(sender, e)
                                         Dim lv = TryCast(ctxMenuLv3Lv4Lv5.SourceControl, ListView)
                                         If lv IsNot Nothing Then Await RefreshSelectedLvItems(lv)
                                     End Sub

        ' 沒有選取項就不顯示選單 (虛擬模式用 SelectedIndices，實體模式用 SelectedItems)
        AddHandler ctxMenuLv3Lv4Lv5.Opening, Sub(s, ev)
                                                 Dim lv = TryCast(ctxMenuLv3Lv4Lv5.SourceControl, ListView)
                                                 Dim cnt = If(lv Is Nothing, 0, If(lv.VirtualMode, lv.SelectedIndices.Count, lv.SelectedItems.Count))
                                                 If cnt = 0 Then ev.Cancel = True
                                             End Sub
    End Sub
    Private Sub HandleLv3Lv4Lv5_DrawItem(sender As Object, e As DrawListViewItemEventArgs)
        ' by Gemini 3 Flash, 2026/05/05: 為 OwnerDraw 模式提供基礎渲染
        ' 在 Details 視圖下，大部分工作由 DrawSubItem 完成，此處僅確保基本行為正確。
        If e.Item.Selected Then e.DrawDefault = True ' 選取狀態交由系統繪製，確保藍色高亮正確
    End Sub
    Private Sub HandleLv3Lv4Lv5_DrawSubItem(sender As Object, e As DrawListViewSubItemEventArgs)
        ' by Gemini 3 Flash, 2026/04/26: ListView3, Listview4, ListView5 共用的 OwnerDraw 繪製邏輯
        ' 2026/05/09 by Gemini 3 Flash: Resize 期間暫停繪製，避免文字滑動殘影
        'If _isResizingLv Then Return
        ' 2026/05/05 by Gemini 3 Flash: 修正非懸停狀態下的背景繪製，確保保留群組背景色 (BackColor)

        ' 1. 決定底色與文字色
        Dim backColor As Color = e.Item.BackColor
        ' 2026/06/15 by Simon/Claude Opus 4.8: Lv4/5 實體模式的刷新狀態集中於此判斷 (從 item.Tag 取 EntryID)；要加粗體/斜體等屬性只需在此擴充
        '   註：Lv3 為虛擬模式，非懸停時不走 DrawSubItem，其刷新藍色於 Lv3_RetrieveVirtualItem 設定
        Dim isRefreshed As Boolean = TypeOf e.Item.Tag Is MailItemInfo AndAlso _refreshedList.Contains(DirectCast(e.Item.Tag, MailItemInfo).EntryID)
        ' 2026/06/18 by Simon/Claude Opus 4.8: Q4 刪除後失去配對的孤兒信標紅(OwnerDraw 不吃 item.ForeColor，與藍字同機制集中於此)；紅字優先於藍字
        Dim isOrphaned As Boolean = TypeOf e.Item.Tag Is MailItemInfo AndAlso _lv5OrphanedList.Contains(DirectCast(e.Item.Tag, MailItemInfo).EntryID)
        Dim foreColor As Color = If(isOrphaned, Color.Red, If(isRefreshed, Color.Blue, SystemColors.WindowText))

        If _lastHoveredLvItem IsNot Nothing AndAlso e.ItemIndex = _lastHoveredLvItem.Index AndAlso Not e.Item.Selected Then
            ' 懸停中且未被選取：使用懸停灰色
            backColor = ThemeColors.MercuryGray
            foreColor = SystemColors.InactiveCaptionText
        ElseIf e.Item.Selected Then
            ' 已選取：讓系統處理選取藍色
            e.DrawDefault = True
            Return
        End If

        ' 2. 繪製背景 (使用 SolidBrush 繪製 BackColor，解決懸停消失問題)
        Using bgBrush As New SolidBrush(backColor)
            e.Graphics.FillRectangle(bgBrush, e.Bounds)
        End Using

        ' 3. 繪製文字 (使用 TextRenderer 確保對齊與抗鋸齒)
        Dim textRect As Rectangle = e.Bounds
        Dim flags As TextFormatFlags = TextFormatFlags.VerticalCenter Or TextFormatFlags.EndEllipsis Or TextFormatFlags.SingleLine Or TextFormatFlags.PreserveGraphicsClipping

        ' 依照欄位對齊方式設定旗標與微調位移 (位移量根據 USER 實測反饋調整)
        If e.ColumnIndex = 0 Then
            textRect.X += 2 : textRect.Width -= 2 ' 避免第一欄文字貼著邊線
            flags = flags Or TextFormatFlags.Left
        ElseIf e.Header.TextAlign = HorizontalAlignment.Right Then
            flags = flags Or TextFormatFlags.Right
        ElseIf e.Header.TextAlign = HorizontalAlignment.Center Then
            flags = flags Or TextFormatFlags.HorizontalCenter
            textRect.X += 1 ' 修正往左偏移的問題
        Else
            flags = flags Or TextFormatFlags.Left
            textRect.X += 2 ' by USER, 2026/04/26: 寄件者之後的欄位再多補 1px (共 2px)
        End If

        TextRenderer.DrawText(e.Graphics, e.SubItem.Text, e.Item.Font, textRect, foreColor, flags)
    End Sub
    Private Sub HandleLv3Lv4Lv5_MouseDown(sender As Object, e As MouseEventArgs)
        ' 2026/06/14 by Simon/Claude Opus 4.8: 右鍵點在某項目上但未選取時，先選取它 (供右鍵選單刷新)；已選取則保留多選
        If e.Button <> MouseButtons.Right Then Return

        Dim lv = DirectCast(sender, ListView)
        Dim it = lv.GetItemAt(e.X, e.Y)
        If it Is Nothing Then Return

        Dim multi = My.Computer.Keyboard.CtrlKeyDown OrElse My.Computer.Keyboard.ShiftKeyDown
        If lv.VirtualMode Then
            If Not lv.SelectedIndices.Contains(it.Index) Then
                If Not multi Then lv.SelectedIndices.Clear()
                lv.SelectedIndices.Add(it.Index)
            End If
        Else
            If Not it.Selected Then
                If Not multi Then lv.SelectedItems.Clear()
                it.Selected = True
            End If
        End If
    End Sub
    Private Sub HandleLv3Lv4Lv5_MouseClick(sender As Object, e As MouseEventArgs)
        ''' <summary>
        ''' 共通滑鼠點擊: 複製主旨與路徑預覽
        ''' </summary>
        Dim lv = DirectCast(sender, ListView)
        Dim item As ListViewItem = lv.GetItemAt(e.X, e.Y)

        'If item IsNot Nothing AndAlso e.Button = MouseButtons.Left Then
        '    ' 單擊左鍵複製主旨到剪貼簿，這原本是 Listview4 獨有的方便設計，現在擴展到 Tab3 共用 (by Gemini 3.1 Pro, 2026/04/21)
        '    Clipboard.SetText(item.SubItems(0).Text)
        'End If
        '' 路徑更新邏輯統一由 ShowLv3Lv4Lv5PathToPgrsBar 接管
        'ShowLv3Lv4Lv5PathToPgrsBar(sender, e)
    End Sub
    Private Sub HandleLv3Lv4Lv5_DoubleClick(sender As Object, e As EventArgs)
        ''' <summary>
        ''' 共通雙擊開啟
        ''' </summary>
        OpenMailByEntryID(GetSelectedEntryIDs(DirectCast(sender, ListView)))
    End Sub
    Private Sub ShowLv3Lv4Lv5PathToPgrsBar(sender As Object, e As EventArgs)
        ''' <summary>
        ''' 將目前ListviewItem 的 FolderPath 顯示於 ProgressBar2
        ''' </summary>
        Dim lv = DirectCast(sender, ListView)
        If lv.SelectedIndices.Count = 0 Then Return

        Dim path As String = ""
        If lv.VirtualMode Then
            ' Tab3 模式: 從 _lv3MailList 根據 SelectedIndices 取 FolderPath
            Dim idx = lv.SelectedIndices(0)
            If idx >= 0 AndAlso idx < _lv3MailList.Count Then path = _lv3MailList(idx).FolderPath
        Else
            ' Tab4 模式: 從 SelectedItems(0).Tag (MailItemInfo) 取 FolderPath
            If TypeOf lv.SelectedItems(0).Tag Is MailItemInfo Then path = DirectCast(lv.SelectedItems(0).Tag, MailItemInfo).FolderPath
        End If

        If Not String.IsNullOrEmpty(path) Then PgrsBar2.Text = path
    End Sub
    Private Async Sub HandleLv3Lv4Lv5_KeyDown(sender As Object, e As KeyEventArgs)
        ''' <summary>
        ''' 共通鍵盤按鍵 (Enter: 開啟, ESC: 目錄焦點歸位, Ctrl+A: 全選)
        ''' </summary>
        Dim lv = DirectCast(sender, ListView)

        If e.KeyCode = Keys.Enter Then
            OpenMailByEntryID(GetSelectedEntryIDs(lv))
            e.Handled = True : e.SuppressKeyPress = True

        ElseIf e.KeyCode = Keys.Escape Then
            If lv.VirtualMode Then lv.SelectedIndices.Clear() Else lv.SelectedItems.Clear()
            ' 對應不同的 TreeView 給予控制權
            If lv Is ListView3 Then SimTree3.Focus()
            If lv Is Listview4 Then Lv4Topic.Focus()   ' 2026/5/29 by Simon/Claude: 拆分SimTree4的雙重模式後，將 Tab4 ESC 焦點從 SimTree4 調整到 Lv4Topic
            If lv Is ListView5 Then SimTree5.Focus() ' 2026/05/03 by Gemini 3.1 Pro: 新增 Tab5 ESC 焦點歸位
            e.Handled = True

        ElseIf e.Control AndAlso e.KeyCode = Keys.A Then
            LviSelectAll(lv, e)

        ElseIf e.KeyCode = Keys.F5 Then
            ' 2026/06/14 by Simon/Claude Opus 4.8: F5 統一在此分派 (Lv4 已移除其專屬 F5 分支，避免雙重觸發)
            e.Handled = True : e.SuppressKeyPress = True
            If lv Is ListView3 OrElse lv Is ListView5 Then
                Await RefreshAllLvItems(lv)    ' Lv3/Lv5 全體刷新 (依數量自動 A/B)
            ElseIf lv Is Listview4 Then
                Await RefreshLv4Result(lv)      ' Lv4 沿用：重讀目前系列郵件
            End If
        ElseIf e.Control AndAlso e.KeyCode = Keys.A Then
            LviSelectAll(lv, e)
        End If
    End Sub

    ' Form1_Refresh.vb  —  郵件實體資訊強制刷新 (Lv3/Lv4/Lv5)
    ' ==============================================================
    ' 功能：
    '   (A) Lv3/Lv5 按 F5 → 強制刷新「目前顯示清單」內所有郵件的實體資訊 (Subject/Size/RcvTime/SenderName)
    '   (B) Lv3/Lv4/Lv5 右鍵選單「強制刷新選取的郵件」→ 只刷選取項，並額外更新 AttachCount
    '
    ' 設計重點：
    '   1. 跳過所有 cache (_cacheXXX / SSD)，直接打 COM 讀真值；讀完寫回顯示清單，並 patch 記憶體 cache。
    '   2. A/B 路徑「依數量」自動選擇 (與哪個 LV 無關)：
    '        targetList.Count <  REFRESH_BATCH_THRESHOLD(42) → 方法A：逐封 RefreshMailInfoL3
    '        targetList.Count >= 門檻                        → 方法B：依資料夾 GetTable+GetArray 批次
    '   3. 只刷「已在記憶體中的項目」：顯示清單 + 已含該 EntryID 的 cache；
    '      尚未 lazy load 的欄位 (附件數/檔名) 不主動額外讀取 (全體 F5 一律略過 AttachCount)。
    '   4. AttachCount 由「操作別」決定 (readAttachCount)：全體F5=False、右鍵刷新=True；
    '      但方法B 結構上讀不到附件數 → 即使右鍵選 >=42 封，該批 AttachCount 也不會更新 (效能優先的邊界取捨)。
    '   5. SSD 不在此逐封碎寫，交給正常存檔流程；snapshot 計數不動 (沒有增減郵件)。
    '   6. 失效郵件 (EntryID 找不到) 目前一律「保留舊資料 + 記錄」；日後若要移除/標記，只需改呼叫端政策。
    ' 2026/06/14 by Simon/Claude Opus 4.8
    ' ==============================================================
    Private Async Function RefreshSelectedLvItems(lv As ListView) As Task
        ' 2026/06/14 by Simon/Claude Opus 4.8: 右鍵單筆/複數刷新 (Lv3/4/5) — 只刷選取項 (readAttachCount:=True 更新附件數)
        '   注意：選取 >=42 會落到方法B，方法B讀不到附件數 → 該批 AttachCount 不更新 (效能優先的邊界取捨)
        _dbg("開始")
        Dim target = GetLviMailTarget(lv)
        If target.Count = 0 Then Return

        _isUserBusy = True : Cursor = Cursors.WaitCursor
        Try
            Dim stats = Await RefreshLviCore(target, readAttachCount:=True, ct:=CancellationToken.None)
            If lv Is ListView3 Then ListView3.Invalidate()
            If lv Is Listview4 Then RenderLv4Result(_tv4SelectedTopicMailList)
            If lv Is ListView5 Then RenderLv5Group(_lv5PrevGroupResults, _tv5PrevSearchMode)
            PgrsBar1.Text = $"已刷新選取的 {stats.Updated} 封 (失效 {stats.NotFound}, 錯誤 {stats.Errored})。" : PgrsBar2.Text = ""
        Catch ex As OperationCanceledException
            PgrsBar1.Text = "刷新已取消。"
        Catch ex As System.Exception
            _dbg("選取刷新錯誤", ex.Message)
        Finally
            _isUserBusy = False : Cursor = Cursors.Default
            _dbg("結束")
        End Try
    End Function
    Private Async Function RefreshAllLvItems(lv As ListView) As Task
        ' 2026/06/14 by Simon/Claude Opus 4.8: 全體 F5 (Lv3/Lv5) — 蒐集底層所有郵件成 targetList，呼叫核心 (readAttachCount:=False，全體不讀附件數)
        _dbg("開始")
        Dim targetList As New List(Of (lst As List(Of MailItemInfo), idx As Integer))
        If lv Is ListView3 Then
            For i = 0 To _lv3MailList.Count - 1 : targetList.Add((_lv3MailList, i)) : Next
        ElseIf lv Is ListView5 Then
            If _lv5PrevGroupResults Is Nothing Then Return
            For Each kv In _lv5PrevGroupResults
                Dim inner = kv.Value
                For j = 0 To inner.Count - 1 : targetList.Add((inner, j)) : Next
            Next
        Else
            Return
        End If
        If targetList.Count = 0 Then Return

        _isUserBusy = True : Cursor = Cursors.WaitCursor
        Try
            Dim stats = Await RefreshLviCore(targetList, readAttachCount:=False, ct:=CancellationToken.None)
            If lv Is ListView3 Then ListView3.Invalidate() Else RenderLv5Group(_lv5PrevGroupResults, _tv5PrevSearchMode)
            PgrsBar1.Text = $"已刷新 {stats.Updated} 封 (失效 {stats.NotFound}, 錯誤 {stats.Errored})。" : PgrsBar2.Text = ""
        Catch ex As OperationCanceledException
            PgrsBar1.Text = "刷新已取消。"
        Catch ex As System.Exception
            _dbg("全體刷新錯誤", ex.Message)
        Finally
            _isUserBusy = False : Cursor = Cursors.Default
            _dbg("結束")
        End Try
    End Function
    Private Async Function RefreshLviCore(target As List(Of (lst As List(Of MailItemInfo), idx As Integer)), readAttachCount As Boolean, ct As CancellationToken) As Task(Of RefreshStats)
        ' 2026/06/14 by Simon/Claude Opus 4.8: 刷新核心分派器 — 給定一組 (清單,索引) targetList，依數量選 A/B 重讀 COM 並寫回 + patch 記憶體 cache
        '   count <  門檻 → 方法A：逐封 RefreshMailInfoL3 (readAttachCount 決定是否讀附件數)
        '   count >= 門檻 → 方法B：依資料夾 GetTable+GetArray 批次 (讀不到附件數，一律略過)
        Dim stats As New RefreshStats
        If target Is Nothing OrElse target.Count = 0 Then Return stats

        Dim swThrottle As Stopwatch = Stopwatch.StartNew()
        Dim total As Integer = target.Count

        If total < REFRESH_BATCH_THRESHOLD Then
            ' ── 方法A：逐封開信 ──
            For i As Integer = 0 To total - 1
                ct.ThrowIfCancellationRequested()
                Dim s = target(i)
                Dim m As MailItemInfo = s.lst(s.idx)
                Select Case RefreshMailInfoL3(m, readAttachCount)
                    Case RefreshResult.Updated : s.lst(s.idx) = m : UpdateMailCaches(m, readAttachCount) : stats.Updated += 1 : _refreshedList.Add(m.EntryID)  ' 2026/06/15 by Simon/Claude Opus 4.8: 標記為已刷新
                    Case RefreshResult.NotFound : stats.NotFound += 1   ' 保留舊資料不動 (移除政策由呼叫端日後決定)
                    Case Else : stats.Errored += 1
                End Select
                Await SmartThrottle(swThrottle, ct, ThrottleFreq.Hii, Sub() PgrsBar2.Text = $"刷新郵件 (逐封): {i + 1} / {total}...")
            Next
        Else
            ' ── 方法B：依資料夾批次 GetTable+GetArray ──
            Dim byFolder = target.GroupBy(Function(s) s.lst(s.idx).FolderPath)
            Dim done As Integer = 0
            For Each grp In byFolder
                ct.ThrowIfCancellationRequested()
                Dim fieldDict As Dictionary(Of String, MailItemInfo) = Await GetFolderBasicByEntryIDL3(grp.Key, ct)
                If fieldDict Is Nothing Then
                    ' 資料夾解析/掃描失敗 → 該組退回逐封 (確保不整批漏掉)
                    For Each s In grp
                        ct.ThrowIfCancellationRequested()
                        Dim m As MailItemInfo = s.lst(s.idx)
                        Select Case RefreshMailInfoL3(m, readAttachCount)
                            Case RefreshResult.Updated : s.lst(s.idx) = m : UpdateMailCaches(m, readAttachCount) : stats.Updated += 1 : _refreshedList.Add(m.EntryID)  ' 2026/06/15 by Simon/Claude Opus 4.8: 標記為已刷新
                            Case RefreshResult.NotFound : stats.NotFound += 1
                            Case Else : stats.Errored += 1
                        End Select
                        done += 1
                        Await SmartThrottle(swThrottle, ct, ThrottleFreq.Hii, Sub() PgrsBar2.Text = $"刷新郵件 (退回逐封): {done} / {total}...")
                    Next
                    Continue For
                End If

                For Each s In grp
                    Dim m As MailItemInfo = s.lst(s.idx)
                    Dim fresh As MailItemInfo = Nothing
                    If fieldDict.TryGetValue(m.EntryID, fresh) Then
                        ' 只搬基本欄位，保留既有 AttachCount (方法B不碰附件數)
                        m.Subject = fresh.Subject : m.Size = fresh.Size : m.RcvTime = fresh.RcvTime : m.SenderName = fresh.SenderName
                        s.lst(s.idx) = m : UpdateMailCaches(m, readAttachCount:=False) : stats.Updated += 1 : _refreshedList.Add(m.EntryID)  ' 2026/06/15 by Simon/Claude: 標記為已刷新
                    Else
                        stats.NotFound += 1   ' 該 EntryID 已不在資料夾 (移動/刪除)
                    End If
                    done += 1
                Next
                Await SmartThrottle(swThrottle, ct, ThrottleFreq.Hii, Sub() PgrsBar2.Text = $"刷新郵件 (批次): {done} / {total}...")
            Next
        End If

        Return stats
    End Function
    Private Function GetLviMailTarget(lv As ListView) As List(Of (lst As List(Of MailItemInfo), idx As Integer))
        ' 2026/06/14 by Simon/Claude Opus 4.8: 將選取項對應回底層 List 的 (清單,索引)
        '   Lv3 虛擬 → SelectedIndices 直接是 _lv3MailList 索引；Lv4/Lv5 實體 → 用 item.Tag 的 EntryID 反查 (render 會重排，不能用顯示索引)
        Dim targetList As New List(Of (lst As List(Of MailItemInfo), idx As Integer))
        If lv Is ListView3 Then
            For Each i As Integer In lv.SelectedIndices
                If i >= 0 AndAlso i < _lv3MailList.Count Then targetList.Add((_lv3MailList, i))
            Next

        ElseIf lv Is Listview4 Then
            If _tv4SelectedTopicMailList Is Nothing Then Return targetList
            For Each it As ListViewItem In lv.SelectedItems
                If TypeOf it.Tag Is MailItemInfo Then
                    Dim eid = DirectCast(it.Tag, MailItemInfo).EntryID
                    Dim j = _tv4SelectedTopicMailList.FindIndex(Function(x) x.EntryID = eid)
                    If j >= 0 Then targetList.Add((_tv4SelectedTopicMailList, j))
                End If
            Next

        ElseIf lv Is ListView5 Then
            If _lv5PrevGroupResults Is Nothing Then Return targetList
            Dim wantIDs As New HashSet(Of String)(StringComparer.Ordinal)
            For Each it As ListViewItem In lv.SelectedItems
                If TypeOf it.Tag Is MailItemInfo Then wantIDs.Add(DirectCast(it.Tag, MailItemInfo).EntryID)
            Next
            For Each kv In _lv5PrevGroupResults
                Dim inner = kv.Value
                For j As Integer = 0 To inner.Count - 1
                    If wantIDs.Contains(inner(j).EntryID) Then targetList.Add((inner, j))
                Next
            Next
        End If
        Return targetList
    End Function
    Private Sub UpdateMailCaches(mail As MailItemInfo, readAttachCount As Boolean)
        ' 2026/06/14 by Simon/Claude Opus 4.8: 只 patch「現有 in-memory cache 中已含此 EntryID」的項目；絕不新建 key、不觸發掃描/lazy load
        '   內層 List 是參考型別，原地改元素即可，dict 與 snapshot 不動

        ' ① Tab3 附件清單快取
        Dim t3 As FolderCacheTab3 = Nothing
        If _cacheAttachMailList.TryGetValue(mail.FolderPath, t3) AndAlso t3.AttachMailList IsNot Nothing Then
            Dim lst = t3.AttachMailList
            Dim j As Integer = lst.FindIndex(Function(x) x.EntryID = mail.EntryID)
            If j >= 0 Then
                Dim c = lst(j)
                c.Subject = mail.Subject : c.Size = mail.Size : c.RcvTime = mail.RcvTime : c.SenderName = mail.SenderName
                If readAttachCount Then c.AttachCount = mail.AttachCount   ' 只有逐封(方法A)才有可信附件數
                lst(j) = c
            End If
        End If

        ' ② Tab4 基本資訊快取 (Mails 是 List(Of (Mail, Topic)))
        Dim t4 As (Mails As List(Of (Mail As MailItemInfo, Topic As String)), Snap As Long) = Nothing
        If _cacheBasicMailInfo.TryGetValue(mail.FolderPath, t4) AndAlso t4.Mails IsNot Nothing Then
            Dim lst = t4.Mails
            Dim j As Integer = lst.FindIndex(Function(x) x.Mail.EntryID = mail.EntryID)
            If j >= 0 Then
                Dim e = lst(j)
                e.Mail.Subject = mail.Subject : e.Mail.Size = mail.Size : e.Mail.RcvTime = mail.RcvTime : e.Mail.SenderName = mail.SenderName
                If readAttachCount Then e.Mail.AttachCount = mail.AttachCount
                lst(j) = e
            End If
        End If

        ' ③ 附件檔名快取：只在逐封(readAttachCount=True)時失效該筆，避免日後拿到過時檔名 (不主動重讀)
        If readAttachCount Then
            Dim dummy As List(Of String) = Nothing
            _cacheAttachFilename.TryRemove(mail.EntryID, dummy)
        End If
    End Sub
    Private Function GetNewSortOrder(clickedColumn As Integer, lastColumn As Integer, currentOrder As SortOrder) As SortOrder
        ''' <summary>
        ''' 共用排序方向切換邏輯
        ''' </summary>
        Return If(clickedColumn = lastColumn AndAlso currentOrder = SortOrder.Ascending, SortOrder.Descending, SortOrder.Ascending)
    End Function
    Private Function SortMailList(sourceList As List(Of MailItemInfo), columnIndex As Integer, order As SortOrder) As List(Of MailItemInfo)
        ''' <summary>
        ''' 共用 MailItemInfo 清單排序邏輯 (供 Tab3, Tab4 右側使用)
        ''' 2026/05/30 by Gemini/Simon
        ''' </summary>
        If sourceList?.Count = 0 Then Return sourceList

        Select Case columnIndex
            Case 0 : Return If(order = SortOrder.Ascending, sourceList.OrderBy(Function(x) x.Subject).ToList(), sourceList.OrderByDescending(Function(x) x.Subject).ToList())           ' 主旨
            Case 1 : Return If(order = SortOrder.Ascending, sourceList.OrderBy(Function(x) x.Size).ToList(), sourceList.OrderByDescending(Function(x) x.Size).ToList())                 ' 大小
            Case 2 : Return If(order = SortOrder.Ascending, sourceList.OrderBy(Function(x) x.RcvTime).ToList(), sourceList.OrderByDescending(Function(x) x.RcvTime).ToList()) ' 收到日期
            Case 3 : Return If(order = SortOrder.Ascending, sourceList.OrderBy(Function(x) x.SenderName).ToList(), sourceList.OrderByDescending(Function(x) x.SenderName).ToList())     ' 寄件者
            Case 4 : Return If(order = SortOrder.Ascending, sourceList.OrderBy(Function(x)  ' 附件數 (Tab3 專用，依賴全域 _cacheAttachFilename)
                                                                                   Dim files As List(Of String) = Nothing
                                                                                   Return If(_cacheAttachFilename.TryGetValue(x.EntryID, files), files.Count, 0)
                                                                               End Function).ToList(),
                                                            sourceList.OrderByDescending(Function(x)
                                                                                             Dim files As List(Of String) = Nothing
                                                                                             Return If(_cacheAttachFilename.TryGetValue(x.EntryID, files), files.Count, 0)
                                                                                         End Function).ToList())
            Case Else
                Return sourceList
        End Select
    End Function
#End Region
#End Region

#Region "■ 09 Tab6: Setting & Debug 設定/測試"
#Region "  ├ Setting 設定"
    Private Async Sub SaveCache_Click(sender As Object, e As EventArgs) Handles SaveCache.Click
        Await SaveCachesToDB()
        RefreshLv6DbStats()
    End Sub
    Private Async Sub LoadCache_Click(sender As Object, e As EventArgs) Handles LoadCache.Click
        Await LoadCachesFromDB()
        Dim st = GetDBSummary()
        PgrsBar2.Text = $"DB 統計 — folder_stats:{st.fc} 筆 / attach_maillist:{st.mb} 筆 / attach_filenames:{st.at} 筆 / year_counts:{st.yc} 筆 / month_counts:{st.mc} 筆 / {st.kb} KB"

    End Sub
    Private Async Sub ClearCache_Click(sender As Object, e As EventArgs) Handles ClearCache.Click
        ' ---------------------------------------------------------------
        ' ClearCache_Click — [透明化控制] 分流清理記憶體或 SSD 快取
        ' by Gemini, 2026/04/10: 實作三路自選對話框
        ' ---------------------------------------------------------------
        _dbg("開始")
        Dim st = GetDBSummary()
        Dim lastTimeStr As String = st.lastTs

        ' 1. 準備訊息文字
        Dim msg As String = $"【快取清理選項】" & vbCrLf & vbCrLf &
                            $"--- 目前 SSD 快取現況 ---" & vbCrLf &
                            $"最後儲存時間：{lastTimeStr}" & vbCrLf &
                            $"資料夾統計：{st.fc} 筆" & vbCrLf &
                            $"附件郵件：{st.mb} 筆" & vbCrLf &
                            $"檔案大小：{st.kb} KB" & vbCrLf & vbCrLf &
                            $"請選擇你要清理的範圍："

        ' 2. 使用動態 Form 實作三按鈕對話框 (為了精確符合使用者需求)
        Using f As New Form()
            f.Text = "清理快取" : f.Size = New Size(450, 280)
            f.StartPosition = FormStartPosition.CenterParent : f.FormBorderStyle = FormBorderStyle.FixedDialog
            f.MaximizeBox = False : f.MinimizeBox = False : f.BackColor = Color.White
            f.Font = _fontDefault

            Dim lbl As New Label() With {.Text = msg, .Location = New Point(20, 20), .Size = New Size(400, 150)}
            f.Controls.Add(lbl)

            Dim btnMem As New Button() With {.Text = "僅記憶體", .DialogResult = DialogResult.Yes, .Location = New Point(20, 180), .Size = New Size(120, 40), .BackColor = Color.LightBlue}
            Dim btnSSD As New Button() With {.Text = "僅 SSD (重建)", .DialogResult = DialogResult.No, .Location = New Point(155, 180), .Size = New Size(120, 40), .BackColor = Color.MistyRose}
            Dim btnBoth As New Button() With {.Text = "兩者皆清", .DialogResult = DialogResult.Retry, .Location = New Point(290, 180), .Size = New Size(120, 40), .BackColor = Color.Orange}

            f.Controls.AddRange({btnMem, btnSSD, btnBoth})
            f.AcceptButton = btnMem

            Dim result = f.ShowDialog()

            ' 3. 根據選擇執行處置
            Select Case result
                    ' 僅記憶體
                Case DialogResult.Yes
                    ClearMemoryCachesCore()
                    PgrsBar2.Text = "已完成：僅清除記憶體快取 (SSD 保留)"
                    _dbg("清理", "僅記憶體")

                    ' 僅 SSD
                Case DialogResult.No
                    If MessageBox.Show("【安全提示】這將把目前的 SSD 快取檔更名備份 ( .zip) 並重新建立空白資料表。" & vbCrLf & "這可以解決 Schema 不相容問題且具備救援機制，確定嗎？", "重置 SSD 快取", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = DialogResult.OK Then
                        Await ZipAndRebuildDB()
                        If chkClearSimHash.Checked Then DeleteDbMail() ' 2026/06/17 by Simon/Claude Opus 4.8: 勾選 chkClearSimHash 時連同 SSD 一併清除 SimHash 獨立 db (DeleteDbMail 內含關連線/刪檔/清記憶體/重建空表)
                        PgrsBar2.Text = "已完成：SSD 資料庫已備份並重新初始化" & If(chkClearSimHash.Checked, " (含 SimHash 清除)", "")
                        _dbg("清理", "僅 SSD (已備份)" & If(chkClearSimHash.Checked, " + SimHash", ""))
                    End If

                    ' 兩者皆清
                Case DialogResult.Retry
                    If MessageBox.Show("確定要清除記憶體並備份重置 SSD 快取嗎？", "最後確認", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) = DialogResult.OK Then
                        ClearMemoryCachesCore()
                        Await ZipAndRebuildDB()
                        If chkClearSimHash.Checked Then DeleteDbMail() ' 2026/06/17 by Simon/Claude Opus 4.8: 勾選 chkClearSimHash 時連同 SSD 一併清除 SimHash 獨立 db
                        PgrsBar2.Text = "已完成：記憶體與 SSD 快取已全數歸零 (舊 SSD 檔已備份)" & If(chkClearSimHash.Checked, " + SimHash 已清除", "")
                        _dbg("清理", "FULL CLEAN (已備份)" & If(chkClearSimHash.Checked, " + SimHash", ""))
                    End If
            End Select
        End Using

        RefreshLv6DbStats()
        _dbg("結束")

    End Sub
    Private Async Sub RenewCache_Click(sender As Object, e As EventArgs) Handles RenewCache.Click
        ' 2026/04/09 重構: 原本只做孤兒清除，現在改呼叫完整的 RenewCacheToDB
        '   RenewCacheToDB 內含: Phase1 BFS → Phase2 snapshot 比對 → Phase3 dirty 重算
        '                         Phase4 ancestor 聚合清除 → Phase5 month_counts DB 清除
        '                         Phase6 CleanupOrphan + SaveCachesToDB
        '   RenewIncludeSize 勾選時才重算 folder_size (GetTable 遍歷，大資料夾較慢)
        ' 2026/6/7: by simon/Gemini: 直接在這裡計時顯示整體耗時, 去除原本在 RenewCacheToDB 內的多段計時, 避免重構後的邏輯分散導致耗時統計不完整或混亂

        Dim sw As Stopwatch = Stopwatch.StartNew()
        Try
            Await RenewCacheToDB(RenewFolderSize.Checked)
            Await DbVacuumIfNeeded()    ' 2026/06/16 by Claude Sonnet 4.6: RenewCache 完成後，視碎片比例決定是否執行 VACUUM (freelist_count / page_count > 5% 才執行，避免每次都白等)
            RefreshLv6DbStats()
            Await RefreshAllTreeViews() ' by Gemini 3.0 flash, 2026/04/24: 更新完成後，執行非同步 UI 刷新，確保新資料夾能立即顯示

        Catch ex As OperationCanceledException
            _dbg(" ├ 中斷", "使用者已取消快取更新")
        Finally
            PgrsBar1.Text = $"RenewCache 完成 — 耗時: {sw.Elapsed.TotalSeconds:0.000} 秒"
        End Try
    End Sub
    Private Async Sub RefreshLv6DbStats()
        ' ---------------------------------------------------------------
        ' RefreshLv6DbStats — 切換到 Setting 頁時呼叫，更新 txtDatabaseStats / Listview6
        '
        ' 2026/04/20 重構要點 (by Gemini 3 Flash):
        '   1. 改為 Async Sub，使用 Task.Run 取得資料庫摘要，基礎解決 Tab 切換卡頓。
        '   2. 動態將 txtDatabaseStats 替換為 ListView，改用 Noto Sans TC 字型。
        '   3. 使用 ListView 的雙欄結構，完美達成靠右對齊，且文字渲染較優美。
        ' 2026/5/10 by simon, 刪除txtDatabaseStats, 去除動態生成_lvStat，簡化架構改用 ListView6 顯示統計資料
        ' 2026/6/2 by Gemini: 將計算zip檔案大小的功能和填充統計項目的高級內嵌寫法抽離
        ' ---------------------------------------------------------------

        _dbg("開始")
        Try
            ' ── 步驟 2: 非同步讀取資料庫摘要 (解決卡頓核心) ──
            ' 將耗時的 SQL COUNT(*) 移至背景執行緒
            Dim st = Await Task.Run(Function() GetDBSummary())
            ListView6.BeginUpdate()
            ListView6.Items.Clear()

            '' 輔助方法：填入統計項目 (VB.NET Lambda 不支援 Optional 參數，故移除並於呼叫處補齊)
            ' 2026/6/2 by Gemini: 將高級的內嵌寫法抽離成 AddLv6StatLine 函式並統一格式與樣式
            'Dim AddStat = Sub(label As String, val As String, isHeader As Boolean)
            '                  Dim itm = New ListViewItem(label)
            '                  itm.SubItems.Add(val)
            '                  itm.ForeColor = If(isHeader, Color.DarkRed, ThemeColors.DarkerDimGray)
            '                  itm.Font = If(isHeader, _fontHeader, _fontDefault)
            '                  ListView6.Items.Add(itm)
            '              End Sub

            ' ── 步驟 3: 填充 Memory 數據 ──
            AddLv6StatLine("═══ Memory 快取 ════", "", isHeader:=True)
            AddLv6StatLine("_cacheFolderTree", _cacheFolderTree.Count.ToString("N0") & " 筆")
            AddLv6StatLine("_cacheFolderIDs", _cacheFolderIDs.Count.ToString("N0") & " 筆")
            AddLv6StatLine("_cacheSubTreeList", _cacheSubTreeList.Count.ToString("N0") & " 筆")
            AddLv6StatLine("_cacheIsMailFolder", _cacheIsMailFolder.Count.ToString("N0") & " 筆")
            AddLv6StatLine("", "", isHeader:=False) ' 間隔
            AddLv6StatLine("_cacheMailCount", _cacheMailCount.Count.ToString("N0") & " 筆")
            AddLv6StatLine("_cacheMailCountAll", _cacheMailCountAll.Count.ToString("N0") & " 筆")
            AddLv6StatLine("_cacheFolderCount", _cacheFolderCount.Count.ToString("N0") & " 筆")
            AddLv6StatLine("_cacheFolderCountAll", _cacheFolderCountAll.Count.ToString("N0") & " 筆")
            AddLv6StatLine("_cacheYearCounts", _cacheYearCounts.Count.ToString("N0") & " 筆")
            AddLv6StatLine("_cacheMonthCounts", _cacheMonthCounts.Count.ToString("N0") & " 筆")
            AddLv6StatLine("_cacheAttachMailList", _cacheAttachMailList.Count.ToString("N0") & " 筆")
            AddLv6StatLine("_cacheAttachFilename", _cacheAttachFilename.Count.ToString("N0") & " 筆")
            AddLv6StatLine("_cacheFolderSize", _cacheFolderSize.Count.ToString("N0") & " 筆")
            AddLv6StatLine("_cacheFolderSizeAll", _cacheFolderSizeAll.Count.ToString("N0") & " 筆")
            AddLv6StatLine("", "", isHeader:=False) ' 間隔

            ' ── 步驟 4: 填充 SQLite 數據 ──
            ' 拆分日期與時間 (壓縮寫法)
            Dim parts = st.lastTs.Split(" "c)
            Dim datePart = If(st.lastTs.Contains(" "c), parts(0), st.lastTs)
            Dim timePart = If(st.lastTs.Contains(" "c), parts(1), "N/A")

            AddLv6StatLine("════ SQLite 快取 ════", "", True)
            ' 2026/06/21 by Simon/Claude: DB 檔案大小改雙檔並列；下方依 db 分組(順序不變，OLAcacheMail.db 的兩張表移到區塊末)
            AddLv6StatLine("DB 檔案大小", $"{ (st.kb / 1024.0).ToString(If(st.kb < 10240, "F1", "F0")) } + { (st.kbMail / 1024.0).ToString(If(st.kbMail < 10240, "F1", "F0")) } MB")
            AddLv6StatLine("──── OLAcache.db ────", "", True)
            AddLv6StatLine("folder_stats", st.fc.ToString("N0") & " 筆")
            AddLv6StatLine("senders", st.senders.ToString("N0") & " 筆")         ' 2026/06/14 by Simon/Claude Opus 4.8: 補上 senders，與 DbShowDbFileStat 順序一致
            AddLv6StatLine("basic_maillist", st.basic.ToString("N0") & " 筆")    ' by Gemini 3 Flash, 2026/04/22
            AddLv6StatLine("year_counts", st.yc.ToString("N0") & " 筆")
            AddLv6StatLine("month_counts", st.mc.ToString("N0") & " 筆")
            AddLv6StatLine("attach_maillist", st.mb.ToString("N0") & " 筆")
            AddLv6StatLine("──── OLAcacheMail.db ────", "", True)   ' 2026/06/21 by Simon/Claude: attach_filenames/mail_simhash 住此檔
            AddLv6StatLine("attach_filenames", st.at.ToString("N0") & " 筆")
            AddLv6StatLine("mail_simhash", st.sh.ToString("N0") & " 筆")   ' 2026/06/21 by Simon/Claude: 新增
            AddLv6StatLine("最後更新日期", datePart)
            AddLv6StatLine("最後更新時間", timePart)

            ' ── 步驟 5: 填充 ZIP 備份數據 ── (2026/06/01: added by Claude, 6/2: 抽離函式 by Gemini)
            Dim zipStats = GetFileStats(_dbCachePath, "*.zip")
            AddLv6StatLine($"備份 ZIP 檔總計 ({zipStats.Count}個)", $"{zipStats.TotalMB:N0} MB")

        Catch ex As System.Exception
            ListView6.Items.Clear()
            ListView6.Items.Add(New ListViewItem("❌ 讀取統計失敗: " & ex.Message))
        Finally
            ListView6.EndUpdate()
            PgrsBar1.Text = "已更新Cache / SQL DB 統計資料。"
            _dbg("結束")
        End Try
    End Sub

    Private Sub Lv6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListView6.SelectedIndexChanged
        ''' <summary>
        ''' Tab6 狀態顯示欄 (ListView6) 選擇項目改變時的事件處理常式
        ''' 點擊特定快取資料表時，會在右側 Debug 表單即時輸出該表的 Schema 與底層空間分布明細
        ''' </summary>
        ' 1. 安全防護：確保當前確實有選中項目，避免滑鼠點擊空白處或 ListView 重新整理（Clear）時引發錯誤
        If ListView6.SelectedItems.Count = 0 Then Return

        Try
            ' 2. 擷取使用者點擊的項目名稱 (第一欄的 Label 文字)
            Dim selectedLabel As String = ListView6.SelectedItems(0).Text
            _dbg("開始", $"🔍{selectedLabel}")
            If String.IsNullOrEmpty(selectedLabel) Then Return

            ' 3. 關鍵字過濾與對應：
            ' 由於 ListView6 包含 "DB 檔案大小" 或 "備份 ZIP" 等非實體資料表項目，我們透過 Contains 進行模糊識別，確保精準抓出使用者想看的是哪一張快取表。
            Dim targetTableName As String = ""
            If selectedLabel = "DB 檔案大小" Then               ' 2026/06/13 by Simon/Claude Opus 4.8: 新增對 DB 檔案大小 的特殊識別，觸發專門的空間分布分析
                Dim unused = DbShowDbFileStat()                 ' 明確的 fire-and-forget，編譯器知道你是故意的
            ElseIf selectedLabel.Contains("folder_stats") Then  ' 2026/06/12 by Simon/Claude Opus 4.8: 補上缺漏的分支
                targetTableName = "folder_stats"
            ElseIf selectedLabel.Contains("basic_maillist") Then
                targetTableName = "basic_maillist"
            ElseIf selectedLabel.Contains("attach_maillist") Then
                targetTableName = "attach_maillist"
            ElseIf selectedLabel.Contains("attach_filenames") Then
                targetTableName = "attach_filenames"
            ElseIf selectedLabel.Contains("year_counts") Then   ' 2026/06/12 by Simon/Claude Opus 4.8: 補上缺漏的分支
                targetTableName = "year_counts"
            ElseIf selectedLabel.Contains("month_counts") Then  ' 2026/06/12 by Simon/Claude Opus 4.8: 修正 typo (month_stats → month_counts)
                targetTableName = "month_counts"
            ElseIf selectedLabel.Contains("mail_simhash") Then  ' 2026/06/21 by Simon/Claude: 新增 mail_simhash 分支(DbShowTableStat 內部會路由到 _dbMail)
                targetTableName = "mail_simhash"
            Else
                ' 2026/06/13 by Simon/Claude Opus 4.8: 未來要加上例外續集也可以一行搞定，確保即使該功能發生錯誤也不會影響 UI 穩定性，並將錯誤訊息導向除錯視窗
                ' Me.DbShowDbFileStat().ContinueWith(
                '    Sub(t) _dbg(" ├ 錯誤", $"DbShowDbFileStat task faulted: {t.Exception?.Message}"), TaskContinuationOptions.OnlyOnFaulted Or TaskContinuationOptions.ExecuteSynchronously)
            End If

            ' 4. 根據比對結果執行對應的 Debug 輸出
            ' 【快取資料表分支】呼叫您在 Form1_SQLite2.vb 中實作好的深度空間診斷函數
            '   提示：因為您的 SQLite 持久層同屬 Form1 的 Partial Class，此處可直接利用 Me 呼叫
            If Not String.IsNullOrEmpty(targetTableName) Then
                Dim unused = DbShowTableStat(targetTableName)
            Else
                ' 【一般統計項目分支】如果點選的是一般資訊 (如檔案大小)，在右側除錯視窗同步留下一行簡單的軌跡提示
                '_dbg(, $"[🔍{selectedLabel}]")
            End If

        Catch ex As System.Exception
            ' 5. 全域異常攔截：防止任何 UI 層級的未知異常導致主視窗當掉，並將錯誤導向除錯視窗
            _dbg("❌ UI 事件異常", $"ListView6_SelectedIndexChanged 發生錯誤: {ex.Message}")
        End Try
    End Sub
    Private Sub Lv6_DoubleClick(sender As Object, e As EventArgs) Handles ListView6.DoubleClick
        If ListView6.SelectedItems.Count = 0 Then Return

        ' 取得被雙擊的項目文字
        Dim clickedText = ListView6.SelectedItems(0).Text

        ' 判斷是否為需要開啟資料夾的特定項目
        If clickedText = "DB 檔案大小" OrElse clickedText.StartsWith("備份 ZIP") Then
            Dim dbDir = IO.Path.GetDirectoryName(_dbCachePath)
            If IO.Directory.Exists(dbDir) Then Process.Start("explorer.exe", dbDir)
        End If
    End Sub
    Private Sub Lv6_KeyDown(sender As Object, e As KeyEventArgs) Handles ListView6.KeyDown
        ' F5 強制刷新
        If e.KeyCode = Keys.F5 Then RefreshLv6DbStats()
    End Sub
    Private Sub CheckDebug_CheckedChanged(sender As Object, e As EventArgs) Handles CheckDebug.CheckedChanged
        _isDebugMode = CheckDebug.Checked
        _dbg("開始", _isDebugMode.ToString)
        Dim offset As Integer = If(CheckDebug.Checked, -240, 240)
        Me.Left += offset
        System.Windows.Forms.Cursor.Position = New Point(System.Windows.Forms.Cursor.Position.X + offset,
                                                         System.Windows.Forms.Cursor.Position.Y) ' 2026/3/28 by Gemini: 滑鼠游標跟著表單偏移
        ' 2026/3/26 by Gemini: 先同步位置與大小再顯示，確保第一次 Load 時就能抓到正確的視窗寬度
        If CheckDebug.Checked Then
            SyncDebugFormResize()
            If Not DebugForm.Visible Then DebugForm.Show(Me) ' 2026/3/27 by Gemini: 設定 Owner 確保點選 Form1 時 DebugForm 一起回到前面
        Else
            DebugForm.Hide()
        End If
        _dbg("結束")

    End Sub
    Private Sub AddLv6StatLine(label As String, val As String, Optional isHeader As Boolean = False)
        ' ---------------------------------------------------------------
        ' 建立並格式化 ListViewItem，將項目加入 ListView6
        ' ---------------------------------------------------------------
        'Dim isLink = (label = "DB 檔案大小" OrElse label.StartsWith("備份 ZIP"))
        Dim itm = New ListViewItem(label) With {.ForeColor = If(isHeader, Color.DarkRed, ThemeColors.DarkerDimGray),
                                                .Font = If(isHeader, _fontHeader, _fontDefault)}
        itm.SubItems.Add(val)
        ListView6.Items.Add(itm)
    End Sub
    Private Function GetFileStats(dbPath As String, Optional fileType As String = "*.zip") As (Count As Integer, TotalMB As Double)
        ' ---------------------------------------------------------------
        ' 掃描資料庫目錄下的 ZIP 備份檔案，回傳檔案總數與總大小 (MB)
        ' ---------------------------------------------------------------
        Dim zipDir = If(Not String.IsNullOrEmpty(dbPath), IO.Path.GetDirectoryName(dbPath), "")
        If String.IsNullOrEmpty(zipDir) OrElse Not IO.Directory.Exists(zipDir) Then Return (0, 0)

        Dim zipFiles = IO.Directory.GetFiles(zipDir, fileType)
        Return (zipFiles.Length, zipFiles.Sum(Function(f) New IO.FileInfo(f).Length) / 1048576) ' 1024^2 = 1048576
    End Function
#End Region
#Region "  ├ Debug 測試區"
    Private Async Sub DebugButton_Click(sender As Object, e As EventArgs) Handles DebugButton.Click

        ''' 測試 DASL 是否能在 GetTable 直接濾出含有特定附檔名的信件
        ''Dim folder As Folder = TryCast(SimTree3.SelectedNode.Tag, Folder) : Dim keyword As String = "2025" ' 測試關鍵字
        ''If folder Is Nothing Then MessageBox.Show("請先選擇資料夾") : Return
        ''' 寫法 A: 使用 LIKE (不支援索引的情況)
        ''Dim filterLike As String = $"@SQL=""urn:schemas:httpmail:attachmentfilename"" LIKE '%{keyword}%'"
        ''' 寫法 B: 使用 CI_PHRASEMATCH (依賴 Windows Search 索引，速度極快)
        ''Dim filterCI As String = $"@SQL=""urn:schemas:httpmail:attachmentfilename"" CI_PHRASEMATCH '{keyword}'"
        ''' 這裡您可切換 filterLike 或 filterCI 測試
        ''Dim table As Outlook.Table = Nothing
        ''Try
        ''    table = folder.GetTable(filterLike)
        ''    MessageBox.Show($"測試成功！GetTable 直接過濾出 {table.GetRowCount()} 筆包含 {keyword} 的郵件。")
        ''    ' 印出前幾筆的主旨驗證
        ''    table.Columns.RemoveAll() : table.Columns.Add("Subject") : Dim count As Integer = 0
        ''    While Not table.EndOfTable AndAlso count < 5
        ''        Dim row As Outlook.Row = table.GetNextRow() : _dbg($"郵件: {row("Subject")}") : count += 1
        ''    End While
        ''Catch ex As System.Exception : MessageBox.Show($"DASL 過濾失敗: {ex.Message}")
        ''Finally : If table IsNot Nothing Then Marshal.ReleaseComObject(table)
        ''End Try

        'Await ScanAndMoveRpmsgRdo()
        ' Await SpikeResolveFormCompare()
        ' Await SpikeBodyResolveCompare()

        SpikeSubtreeWalkCompare()

    End Sub

    ' 2026/06/22 by Simon/Claude Opus 4.8: IRM 保護信隔離夾名稱 (方案 Y: 每顆 PST 各建一個同名夾, 同 store 內搬)
    Private Const QUARANTINE_NAME As String = "_IRM_Protected"
    Private Async Function ScanAndMoveRpmsgRdo() As Task
        ' ============================================================================
        ' 2026/06/22 by Simon/Claude Opus 4.8: 【一次性工具】scan-and-move — 把 message.rpmsg 保護信隔離
        '   作法: 依 SimTree3 選定節點掃整棵子樹, 命中(任一附件 .rpmsg)就用 RDO 把該信 Move 到
        '         「同一顆 PST 的 _IRM_Protected 夾」(方案 Y, 同 store 內搬, 避開跨 store 不確定性)。
        '   為何 scan-and-move 而非餵 EntryID: 搬移後 EntryID 會變, 來回 rebind 脆; 掃描當下手上就有 live
        '         RDOMail, 就地搬最穩, 且搬前再驗一次 .rpmsg 防呆。全程走 RDO 不會觸發授權 modal。
        '   ⚠ 破壞性: 信會離開原夾。搬完那些來源夾 + 隔離夾的 SQLite 快照會 stale, 需自行對受影響夾跑 RenewCache。
        '   ⚠ 完整性: 請先把「所有可能含 rpmsg 的 PST」都選進 SimTree3 再執行, 才能一次搬乾淨。
        ' ============================================================================

        ' ── 0. 確保 RDO 已載入 ──
        If _rdo Is Nothing Then Await InitRdoSessionWithoutEULA()
        If _rdo Is Nothing Then _dbg("RDO隔離", "Redemption 初始化失敗, 中止") : Return

        ' ── 1. UI 執行緒抽出選定節點 (EntryID, StoreID, 名稱) ──
        Dim selectedNodes As List(Of TreeNode) = SimTree3.SelectedNodes
        If selectedNodes Is Nothing OrElse selectedNodes.Count = 0 Then _dbg("RDO隔離", "SimTree3 未選取任何 PST/資料夾") : Return
        Dim roots As New List(Of (eid As String, sid As String, name As String))(selectedNodes.Count)
        For Each node As TreeNode In selectedNodes
            Dim f As Folder = TryCast(node.Tag, Folder)
            If f IsNot Nothing Then roots.Add((f.EntryID, f.StoreID, f.Name))
        Next
        If roots.Count = 0 Then _dbg("RDO隔離", "選取節點皆非有效資料夾") : Return

        ' ── 破壞性動作, 先確認 ──
        Dim dr As DialogResult = MessageBox.Show(
            $"即將掃描 {roots.Count} 個根節點, 把所有 message.rpmsg 保護信搬到各自 PST 的「{QUARANTINE_NAME}」夾。" & vbCrLf & vbCrLf &
            "此動作會改變封存結構且不易復原, 確定執行?", "確認隔離搬移", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
        If dr <> DialogResult.Yes Then _dbg("RDO隔離", "使用者取消") : Return
        _dbg("RDO隔離 開始", $"掃描 {roots.Count} 個根節點 ...")

        ' ── 2. 背景 scan-and-move (RDO free-threaded) ──
        Dim moves As New List(Of String)
        Dim movedCount As Integer = 0, failCount As Integer = 0
        Dim scanned As Long = 0
        Dim quarantineCache As New Dictionary(Of String, Redemption.RDOFolder)   ' key: store EntryID → 該 store 的隔離夾

        Await Task.Run(
            Sub()
                For Each r In roots
                    Dim rdoRoot As Redemption.RDOFolder = Nothing
                    Try
                        rdoRoot = _rdo.GetFolderFromID(r.eid, r.sid)
                        Dim folderList As List(Of Redemption.RDOFolder) = GetSubtreeToListL3_Rdo(rdoRoot, includeSubF:=True)
                        For Each rdoF As Redemption.RDOFolder In folderList
                            Dim fName As String = "" : Try : fName = rdoF.Name : Catch : End Try
                            If String.Equals(fName, QUARANTINE_NAME, StringComparison.OrdinalIgnoreCase) Then Continue For   ' 不掃隔離夾自己
                            Dim items = Nothing
                            Try
                                items = rdoF.Items
                                Dim cnt As Integer = items.Count
                                ' 由後往前 (Move 會把命中信移出本夾, 降序迭代不會影響尚未處理的索引)
                                For i As Integer = cnt To 1 Step -1
                                    Dim m As Redemption.RDOMail = TryCast(items.Item(i), Redemption.RDOMail)
                                    If m Is Nothing Then Continue For
                                    Try
                                        scanned += 1
                                        If scanned Mod 5000 = 0 Then _dbg("RDO隔離 進行中", $"已掃 {scanned}, 已搬 {movedCount} ...")

                                        ' 偵測: 任一附件 .rpmsg 即命中 (搬前再驗, 防呆)
                                        Dim matched As String = Nothing
                                        For k As Integer = 1 To m.Attachments.Count
                                            Dim att As Redemption.RDOAttachment = m.Attachments.Item(k)
                                            Try
                                                Dim afn As String = att.FileName
                                                If afn IsNot Nothing AndAlso afn.EndsWith(".rpmsg", StringComparison.OrdinalIgnoreCase) Then matched = afn : Exit For
                                            Finally : TryMarshalRelease(att)
                                            End Try
                                        Next
                                        If matched Is Nothing Then Continue For   ' Finally 會釋放 m

                                        ' 命中: 先取所屬 store, get-or-create 該 store 的隔離夾
                                        Dim st As Redemption.RDOStore = m.Store
                                        Dim stKey As String = st.EntryID
                                        Dim qf As Redemption.RDOFolder = Nothing
                                        If Not quarantineCache.TryGetValue(stKey, qf) Then
                                            qf = GetOrCreateQuarantineRdo(st)
                                            quarantineCache(stKey) = qf
                                        End If

                                        ' 搬移前先擷取資訊 (Move 後 m 會失效、EntryID 會變)
                                        Dim rcv As String = "" : Try : rcv = m.ReceivedTime.ToString("yyyy/MM/dd HH:mm") : Catch : End Try
                                        Dim subj As String = "" : Try : subj = m.Subject : Catch : End Try
                                        Dim sndr As String = "" : Try : sndr = m.SenderName : Catch : End Try
                                        Dim eidOld As String = "" : Try : eidOld = m.EntryID : Catch : End Try
                                        Dim stName As String = "" : Try : stName = st.Name : Catch : End Try

                                        m.Move(qf)   ' ← 搬到隔離夾
                                        movedCount += 1
                                        _dbg($"搬移 #{movedCount}", $"{rcv} | {sndr} | {subj}")
                                        moves.Add(String.Join(vbTab, {$"#{movedCount}", rcv, "寄件:" & sndr, "主旨:" & subj, "原夾:" & fName, "PST:" & stName, "舊EntryID:" & eidOld}))
                                        TryMarshalRelease(st)
                                    Catch ex As System.Exception
                                        failCount += 1
                                        _dbg("RDO隔離 搬移失敗", ex.Message)
                                    Finally
                                        TryMarshalRelease(m)
                                    End Try
                                Next
                            Catch ex As System.Exception
                                _dbg("RDO隔離 略過夾", $"{fName} | {ex.Message}")
                            Finally
                                TryMarshalRelease(items)
                            End Try
                        Next
                    Catch ex As System.Exception
                        _dbg("RDO隔離 根節點失敗", $"{r.name} | {ex.Message}")
                    Finally
                        TryMarshalRelease(rdoRoot)
                    End Try
                Next
            End Sub)

        For Each kv In quarantineCache : TryMarshalRelease(kv.Value) : Next

        ' ── 3. 寫搬移紀錄檔 (與 OLAcache.db 同目錄) ──
        Dim logPath As String = ""
        Try
            Dim baseDir As String = If(String.IsNullOrEmpty(_dbCachePath), My.Application.Info.DirectoryPath, System.IO.Path.GetDirectoryName(_dbCachePath))
            logPath = System.IO.Path.Combine(baseDir, $"RpmsgMoved_{DateTime.Now:yyyyMMdd_HHmmss}.log")
            Dim header As New List(Of String) From {
                $"# RDO 保護信隔離搬移   {DateTime.Now:yyyy/MM/dd HH:mm:ss}",
                $"# 已掃 {scanned} 封, 搬移 {movedCount} 封, 失敗 {failCount} 封 → 各 PST 的 {QUARANTINE_NAME} 夾",
                ""}
            System.IO.File.WriteAllLines(logPath, header.Concat(moves), System.Text.Encoding.UTF8)
        Catch ex As System.Exception
            _dbg("RDO隔離 寫檔失敗", ex.Message)
        End Try

        _dbg("RDO隔離 完成", $"掃 {scanned} | 搬 {movedCount} | 失敗 {failCount} | log: {logPath}")
    End Function
    Private Function GetOrCreateQuarantineRdo(st As Redemption.RDOStore) As Redemption.RDOFolder
        ' 2026/06/22 by Simon/Claude Opus 4.8: 取得(或建立)指定 store 頂層的 _IRM_Protected 隔離夾
        Dim root As Redemption.RDOFolder = st.IPMRootFolder   ' store 的可見頂層夾 (PST 適用)
        Try
            Dim subs = root.Folders
            For i As Integer = 1 To subs.Count
                Dim f As Redemption.RDOFolder = subs.Item(i)
                If String.Equals(f.Name, QUARANTINE_NAME, StringComparison.OrdinalIgnoreCase) Then Return f   ' 已存在
            Next
            Return subs.Add(QUARANTINE_NAME)   ' ★ 唯一沒在文件逐字確認的 API (鏡像 OOM Folders.Add); 不編譯就是這行
        Finally
            TryMarshalRelease(root)
        End Try
    End Function

    Private Async Function SpikeRdoIndependentSession() As Task
        ' 2026/06/19 by Simon/Claude: 拋棄式 spike — 驗證 RDO 獨立 session 三件事
        '   (1) Outlook 已掛載 PST 時，獨立 RDOSession 能否 Logon (PST 共享鎖)
        '   (2) 該獨立 session 能否讀到 RdoTest 內信件的附件檔名
        '   (3) 獨立 session 給的 EntryID，能否用 OOM _olNS.GetItemFromID 還原
        ' 測完即可整段刪除。請暫時掛到一個測試按鈕呼叫。
        ' ============================================================
        Dim log As New List(Of String)
        Dim firstEntryID As String = ""

        ' ── 先取 OOM 端 Gmail_2022 的 StoreID，供步驟3b比對用 ──
        Dim oomStoreId As String = ""
        Try
            For Each st As Outlook.Store In _olNS.Stores
                If st.DisplayName = "Gmail_2022" Then oomStoreId = st.StoreID : Exit For
            Next
        Catch ex As System.Exception
            log.Add("取 OOM StoreID 失敗: " & ex.Message)
        End Try

        ' ── 步驟1+2：背景執行緒用「獨立 session」讀取 ──
        Await Task.Run(Sub()
                           Dim sess As Redemption.RDOSession = Nothing
                           Try
                               sess = New Redemption.RDOSession()
                               ' ⚠ 確認點A：Logon 參數請依你的 Redemption 版本確認
                               '   目標 = 不沿用 Outlook session，建立獨立 MAPI session、用預設 profile、不彈窗
                               sess.Logon("", "", False, True)   ' (ProfileName, Pwd, ShowDialog, NewSession)
                               log.Add("步驟1 OK：獨立 session Logon 成功 (PST 共享鎖未擋住)")

                               ' ── 導覽到 \\Gmail_2022\收件匣\RdoTest ──
                               Dim store As Redemption.RDOStore = Nothing
                               For i As Integer = 1 To sess.Stores.Count
                                   If sess.Stores.Item(i).Name = "Gmail_2022" Then store = sess.Stores.Item(i) : Exit For
                               Next
                               If store Is Nothing Then log.Add("步驟2 失敗：獨立 session 找不到 Gmail_2022 store") : Return

                               ' ⚠ 確認點B：收件匣/RdoTest 確為 IPMRootFolder 下的層級
                               Dim inbox = store.IPMRootFolder.Folders.Item("收件匣")
                               Dim testFolder = inbox.Folders.Item("RdoTest")
                               log.Add($"步驟2 導覽 OK：RdoTest 共 {testFolder.Items.Count} 項")

                               Dim n As Integer = 0
                               For i As Integer = 1 To testFolder.Items.Count
                                   Dim msg = TryCast(testFolder.Items.Item(i), Redemption.RDOMail)
                                   If msg Is Nothing Then Continue For
                                   Dim names As New List(Of String)
                                   For a As Integer = 1 To msg.Attachments.Count
                                       names.Add(msg.Attachments.Item(a).FileName)
                                   Next
                                   If firstEntryID = "" Then firstEntryID = msg.EntryID
                                   n += 1
                                   log.Add($"  信{n}: 附件{names.Count}個 [{String.Join(", ", names)}]")
                               Next
                               log.Add($"步驟2 OK：成功讀出 {n} 封信的附件檔名")
                           Catch ex As System.Exception
                               log.Add("步驟1/2 例外: " & ex.Message)
                           Finally
                               Try : If sess IsNot Nothing Then sess.Logoff()
                               Catch : End Try
                               If sess IsNot Nothing Then TryMarshalRelease(sess)
                           End Try
                       End Sub)

        ' ── 步驟3：回 UI 執行緒，用 OOM 還原「獨立 session 給的」EntryID ──
        If firstEntryID = "" Then
            log.Add("步驟3 跳過：沒有取得任何 EntryID")
        Else
            ' 3a：單參數
            Try
                Dim m1 = TryCast(_olNS.GetItemFromID(firstEntryID), Outlook.MailItem)
                log.Add(If(m1 IsNot Nothing, "步驟3a OK：單參數還原成功 → " & m1.Subject,
                                         "步驟3a 失敗：單參數回傳 Nothing"))
            Catch ex As System.Exception
                log.Add("步驟3a 例外: " & ex.Message)
            End Try
            ' 3b：帶 OOM StoreID
            If oomStoreId <> "" Then
                Try
                    Dim m2 = TryCast(_olNS.GetItemFromID(firstEntryID, oomStoreId), Outlook.MailItem)
                    log.Add(If(m2 IsNot Nothing, "步驟3b OK：帶StoreID還原成功 → " & m2.Subject,
                                             "步驟3b 失敗：帶StoreID回傳 Nothing"))
                Catch ex As System.Exception
                    log.Add("步驟3b 例外: " & ex.Message)
                End Try
            End If
        End If

        MessageBox.Show(String.Join(vbCrLf, log), "RDO Spike 結果")
    End Function ' 2026/6/19~20 獨立 session 給的 EntryID，能否用 OOM _olNS.GetItemFromID 還原
    Private Async Function SpikeParallelReadBenchmark() As Task
        ' 2026/06/22 by Simon/Claude Opus 4.8: 拋棄式 spike P3 — 量測「同 profile 多獨立 session、各讀不同 PST」
        '   的真實平行加速。回答整輪調查唯一未解的問題: K 條 session 跨 PST 並行讀取, wall-clock 是否
        '   勝過序列, 還是被 MSPST provider / 實體磁碟 I/O 序列化。
        '   ★ 兩種 workload 分別計時(附件檔名 vs 內文), 因 Tab3/Tab5 負載特性可能不同。
        '   ★ 公平性: 每個 (workload,K) 各讀「獨立的冷 block」(不同信), 避免暖快取讓後跑的 config 假性變快。
        '   ★ 用與 production 同一支 API: sess.GetMessageFromID(EntryID) + rdoMsg.Attachments/.Body
        Const N As Integer = 2000      ' 每 PST 每個 block 的冷讀信數(想要更穩可調 2000, 時間約翻倍)
        Const M As Integer = 4         ' 取幾個「夠大」的 PST 當標的(K=4 時每 worker 各 1 個)
        Const BLOCKS As Integer = 6    ' 2 workload × 3 K-config; 每 PST 需 >= BLOCKS*N 封冷信

        If _rdo Is Nothing Then Await InitRdoSessionWithoutEULA()
        If _rdo Is Nothing Then _dbg("P3", "Redemption 初始化失敗, 中止") : Return

        Dim profileName As String = ""
        Try : profileName = CStr(CallByName(_rdo, "ProfileName", CallType.Get)) : Catch : End Try
        If profileName = "" Then _dbg("P3", "取不到 _rdo.ProfileName, 中止") : Return
        _dbg("P3", $"===== 平行讀取量測 開始 (profile=[{profileName}], N={N}, M={M}) =====")

        ' ── 1. 收集階段: 臨時一條 session 走訪, 挑 M 個有 >= BLOCKS*N 封的 PST, 各收 BLOCKS*N 個 EntryID ──
        '    (EntryID 是字串、跨 session 通用, 收一次給所有 worker 重用; RDOStore 物件不可跨 session 持有)
        Dim need As Integer = BLOCKS * N
        Dim pstEntryIds As New List(Of (pst As String, pstPath As String, ids As List(Of String)))()
        Dim swCollect As New Stopwatch() : swCollect.Start()
        Await Task.Run(Sub()
                           Dim sess As Redemption.RDOSession = Nothing
                           Try
                               sess = New Redemption.RDOSession()
                               sess.Logon(profileName, "", False, True)
                               For si As Integer = 1 To sess.Stores.Count
                                   If pstEntryIds.Count >= M Then Exit For
                                   Dim st = sess.Stores.Item(si)
                                   Dim nm As String = "" : Try : nm = st.Name : Catch : End Try
                                   Dim pp As String = "" : Try : pp = CStr(CallByName(st, "PstPath", CallType.Get)) : Catch : End Try   ' (c)store-scoped 需 PstPath 去 FindStoreByPath 開 store
                                   Dim ids As New List(Of String)()
                                   Try
                                       Dim stk As New Stack(Of Redemption.RDOFolder)()
                                       stk.Push(st.IPMRootFolder)
                                       Do While stk.Count > 0 AndAlso ids.Count < need
                                           Dim fld = stk.Pop()
                                           Dim cnt As Integer = fld.Items.Count
                                           For ii As Integer = 1 To cnt
                                               If ids.Count >= need Then Exit For
                                               Try
                                                   Dim mm = TryCast(fld.Items.Item(ii), Redemption.RDOMail)
                                                   If mm IsNot Nothing Then ids.Add(mm.EntryID)
                                               Catch : End Try
                                           Next
                                           For fi As Integer = 1 To fld.Folders.Count
                                               stk.Push(fld.Folders.Item(fi))
                                           Next
                                       Loop
                                   Catch : End Try
                                   If ids.Count >= need AndAlso pp <> "" Then
                                       pstEntryIds.Add((nm, pp, ids))
                                       _dbg(" │收集", $"採用 PST [{nm}] (收到 {ids.Count} EntryID, PstPath=[{pp}])")
                                   End If
                               Next
                           Catch ex As System.Exception
                               _dbg(" │收集", "例外: " & ex.GetBaseException().Message)
                           Finally
                               If sess IsNot Nothing Then
                                   Try : sess.Logoff() : Catch : End Try
                                   TryMarshalRelease(sess)
                               End If
                           End Try
                       End Sub)
        swCollect.Stop()
        If pstEntryIds.Count < M Then
            _dbg(" │✗", $"只湊到 {pstEntryIds.Count} 個夠大的 PST(需 {M}, 每個需 >= {need} 封)。請降低 N 或確認在 Work profile。中止。")
            Return
        End If
        _dbg(" │收集", $"完成: {pstEntryIds.Count} 個 PST, 各 {need} EntryID, 耗時 {swCollect.Elapsed.TotalSeconds:F1}s (不計入吞吐量)")

        ' ── 2. 對 2 種 workload × K=1/2/4 量測 ──
        Dim workloads = {"附件檔名", "內文"}
        Dim kConfigs = {1, 2, 4}
        Dim summary As New List(Of String)()

        For w As Integer = 0 To workloads.Length - 1
            Dim isBody As Boolean = (w = 1)
            For kc As Integer = 0 To kConfigs.Length - 1
                Dim K As Integer = kConfigs(kc)
                Dim blockIdx As Integer = w * 3 + kc            ' 0..5, 每 config 取不同冷 block
                Dim lo As Integer = blockIdx * N

                ' 把 M 個 PST round-robin 分給 K 個 worker
                Dim groups As New List(Of List(Of (pst As String, pstPath As String, ids As List(Of String))))()
                For g As Integer = 0 To K - 1 : groups.Add(New List(Of (pst As String, pstPath As String, ids As List(Of String)))()) : Next
                For pi As Integer = 0 To pstEntryIds.Count - 1 : groups(pi Mod K).Add(pstEntryIds(pi)) : Next

                Dim bag As New System.Collections.Concurrent.ConcurrentBag(Of (logonMs As Double, rs As Double, re As Double, mails As Integer, fails As Integer, payload As Long, storeMs As Double, withAttach As Integer))()
                Dim swWall As New Stopwatch() : swWall.Start()

                Dim tasks As New List(Of Task)()
                For g As Integer = 0 To K - 1
                    Dim myGroup = groups(g)
                    tasks.Add(Task.Run(Sub()
                                           Dim sess As Redemption.RDOSession = Nothing
                                           Dim mails As Integer = 0, fails As Integer = 0
                                           Dim bodyChars As Long = 0
                                           Dim payload As Long = 0          ' 附件: 總附件數; 內文: 總字元數 — 揪空轉用
                                           Dim withAttach As Integer = 0    ' 有附件(Count>0)的信數 — 確認取樣是否多為無附件信
                                           Dim storeMs As Double = 0        ' 本 worker 累計開 store(FindStoreByPath)耗時
                                           Dim swLogon As New Stopwatch() : swLogon.Start()
                                           Try
                                               sess = New Redemption.RDOSession()
                                               sess.Logon(profileName, "", False, True)
                                           Catch ex As System.Exception
                                               _dbg(" │✗", $"K={K} worker logon 失敗: {ex.GetBaseException().Message}") : Return
                                           End Try
                                           swLogon.Stop()
                                           Dim rs As Double = swWall.Elapsed.TotalSeconds
                                           Try
                                               For Each pe In myGroup
                                                   ' (c)store-scoped: 每個 PST 在本 worker session 內開一次 store, 之後該 PST 所有信都用 store.GetMessageFromID
                                                   ' (P4 已驗: 跨 session 單參數會 MAPI_E_UNKNOWN_ENTRYID, store-scoped 則 10/10)
                                                   Dim swStore As New Stopwatch() : swStore.Start()
                                                   Dim stStore As Redemption.RDOStore = FindStoreByPath(sess, pe.pstPath)
                                                   swStore.Stop() : storeMs += swStore.Elapsed.TotalMilliseconds   ' A: 開 store 耗時
                                                   If stStore Is Nothing Then fails += N : Continue For    ' 此 PST 在本 session 找不到 → 整塊計失敗
                                                   For idx As Integer = lo To lo + N - 1
                                                       Dim eid As String = pe.ids(idx)
                                                       Try
                                                           Dim rm = TryCast(stStore.GetMessageFromID(eid), Redemption.RDOMail)
                                                           If rm Is Nothing Then fails += 1 : Continue For
                                                           If isBody Then
                                                               Dim b As String = rm.Body
                                                               If b IsNot Nothing Then bodyChars += b.Length : payload += b.Length   ' 強制讀取內文 + 計字元(揪空轉)
                                                           Else
                                                               Dim ac As Integer = rm.Attachments.Count
                                                               If ac > 0 Then withAttach += 1                ' A: 這封真的有附件
                                                               For a As Integer = 1 To ac
                                                                   Dim fn As String = rm.Attachments.Item(a).FileName
                                                                   payload += 1                              ' A: 真讀到的附件檔名數
                                                               Next
                                                           End If
                                                           mails += 1
                                                       Catch
                                                           fails += 1
                                                       End Try
                                                   Next
                                               Next
                                           Catch ex As System.Exception
                                               _dbg(" │✗", $"K={K} worker 讀取例外: {ex.GetBaseException().Message}")
                                           Finally
                                               Dim re As Double = swWall.Elapsed.TotalSeconds
                                               bag.Add((swLogon.Elapsed.TotalMilliseconds, rs, re, mails, fails, payload, storeMs, withAttach))
                                               If sess IsNot Nothing Then
                                                   Try : sess.Logoff() : Catch : End Try
                                                   TryMarshalRelease(sess)
                                               End If
                                           End Try
                                       End Sub))
                Next
                Await Task.WhenAll(tasks)
                swWall.Stop()

                ' 聚合(手動迴圈, 不依賴 LINQ import)
                Dim arr = bag.ToArray()
                Dim totMails As Integer = 0, totFails As Integer = 0
                Dim sumLogon As Double = 0
                Dim readStart As Double = Double.MaxValue, readEnd As Double = 0
                For Each x In arr
                    totMails += x.mails : totFails += x.fails : sumLogon += x.logonMs
                    If x.rs < readStart Then readStart = x.rs
                    If x.re > readEnd Then readEnd = x.re
                Next
                If arr.Length = 0 Then readStart = 0 : readEnd = 0
                Dim avgLogon As Double = If(arr.Length > 0, sumLogon / arr.Length, 0)
                Dim wallRead As Double = Math.Max(0.001, readEnd - readStart)
                Dim thru As Double = totMails / wallRead
                ' A: 額外彙整 — 開store耗時、實際讀取量(揪空轉)、worker 重疊度
                Dim totPayload As Long = 0, totStoreMs As Double = 0, totWithAttach As Integer = 0
                Dim sumReadSpan As Double = 0   ' 各 worker 純讀取時間(rs..re)總和; 與 wallRead 比即重疊度
                For Each x In arr
                    totPayload += x.payload : totStoreMs += x.storeMs : totWithAttach += x.withAttach
                    sumReadSpan += (x.re - x.rs)
                Next
                Dim overlap As Double = sumReadSpan / wallRead   ' ≈K 表完全重疊平行; ≈1 表幾乎沒重疊
                Dim payloadDesc As String = If(w = 1, $"內文{totPayload}字元", $"附件{totPayload}個(有附件信{totWithAttach}/{totMails})")
                Dim line As String = $"[{workloads(w)}] K={K}: 讀 {totMails} 封, 讀取wall={wallRead:F1}s, 吞吐={thru:F0} 封/s, {payloadDesc}, 開store均={totStoreMs / Math.Max(1, arr.Length):F0}ms, 重疊={overlap:F2}x, logon均={avgLogon:F0}ms, resolve失敗={totFails}"
                _dbg(" │量測", line)
                summary.Add(line)
            Next
        Next

        _dbg("P3", "===== 量測結束, 摘要(看 K=2/4 吞吐相對 K=1 有沒有上去) =====")
        For Each s In summary : _dbg(" │摘要", s) : Next
        _dbg("P3", "===== 請把本段全部貼回 =====")
    End Function ' 2026/06/22 P3量測「同 profile 多獨立 session、各讀不同 PST」的真實平行加速
    Private Async Function SpikeResolveFormCompare() As Task
        ' 2026/06/22 by Simon/Claude Opus 4.8: 拋棄式 spike B — 釘死「P3 附件 K=1 達 5589 封/s, 但 production_1/_2 只有 200 多」這 25 倍矛盾。空轉假設已被推翻(本批信 ~55% 有附件), 剩三混淆變數:
        '     (a)resolve 形式  (b)session 種類  (c)取樣信 vs sourceList 不同 ←本 spike 用「同批信讀三遍」消掉 c. 單執行緒(純比 per-call 成本, 不平行), 同一批信讀三種形式:
        '     (1)共用_rdo 單參數      = 現行 production
        '     (2)共用_rdo store-scoped → (1)vs(2)= resolve 形式效應(同一 session)
        '     (3)獨立session store-scoped = P3 → (2)vs(3)= session 種類效應
        '   依賴: FindStoreByPath(寫 P4 時放的 class-level 函數)。前提: Outlook 切 Work profile。測完即可整段刪除。
        Const N As Integer = 2000      ' 取樣信數(單執行緒, 同一批讀三遍; 夠大讓 封/s 穩定)
        If _rdo Is Nothing Then Await InitRdoSessionWithoutEULA()
        If _rdo Is Nothing Then _dbg("B", "Redemption 初始化失敗, 中止") : Return
        Dim profileName As String = ""
        Try : profileName = CStr(CallByName(_rdo, "ProfileName", CallType.Get)) : Catch : End Try
        _dbg("B", $"===== resolve 形式對照 (profile=[{profileName}], N={N}, 單執行緒) =====")

        Await Task.Run(Sub()
                           ' ── 1. 用共用 _rdo 走訪頭部湊 N 封, 記每封所屬 pstPath(供 store-scoped 分組) ──
                           Dim sample As New List(Of (eid As String, pstPath As String))()
                           Try
                               For si As Integer = 1 To _rdo.Stores.Count
                                   If sample.Count >= N Then Exit For
                                   Dim st = _rdo.Stores.Item(si)
                                   Dim pp As String = "" : Try : pp = CStr(CallByName(st, "PstPath", CallType.Get)) : Catch : End Try
                                   If pp = "" Then Continue For
                                   Try
                                       Dim stk As New Stack(Of Redemption.RDOFolder)() : stk.Push(st.IPMRootFolder)
                                       Do While stk.Count > 0 AndAlso sample.Count < N
                                           Dim f = stk.Pop()
                                           For ii As Integer = 1 To f.Items.Count
                                               If sample.Count >= N Then Exit For
                                               Dim mm = TryCast(f.Items.Item(ii), Redemption.RDOMail)
                                               If mm IsNot Nothing Then sample.Add((mm.EntryID, pp))
                                           Next
                                           For fi As Integer = 1 To f.Folders.Count : stk.Push(f.Folders.Item(fi)) : Next
                                       Loop
                                   Catch : End Try
                               Next
                           Catch ex As System.Exception
                               _dbg(" │收集✗", ex.GetBaseException().Message)
                           End Try
                           If sample.Count = 0 Then _dbg(" │✗", "沒取到信, 中止") : Return

                           ' 按 pstPath 分組(供 (2)(3) store-scoped 重用 store; 手動建, 不依賴 LINQ import)
                           Dim groups As New Dictionary(Of String, List(Of String))()
                           For Each s In sample
                               Dim lst As List(Of String) = Nothing
                               If Not groups.TryGetValue(s.pstPath, lst) Then lst = New List(Of String)() : groups(s.pstPath) = lst
                               lst.Add(s.eid)
                           Next
                           _dbg(" │收集", $"取樣 {sample.Count} 封(跨 {groups.Count} 個 PST)")

                           ' 小工具: resolve 後讀附件檔名數(回 -1 表 resolve 失敗)
                           Dim readAttach = Function(rm As Redemption.RDOMail) As Integer
                                                If rm Is Nothing Then Return -1
                                                Dim c As Integer = rm.Attachments.Count
                                                For a As Integer = 1 To c : Dim fn As String = rm.Attachments.Item(a).FileName : Next
                                                Return c
                                            End Function

                           ' ── (1) 共用 _rdo 單參數 (現行 production) ──
                           Dim sw1 As New Stopwatch() : sw1.Start()
                           Dim att1 As Long = 0, fail1 As Integer = 0
                           For Each s In sample
                               Try
                                   Dim c = readAttach(TryCast(_rdo.GetMessageFromID(s.eid), Redemption.RDOMail))
                                   If c < 0 Then fail1 += 1 Else att1 += c
                               Catch : fail1 += 1
                               End Try
                           Next
                           sw1.Stop()
                           _dbg(" │(1)", $"共用_rdo 單參數: {sample.Count / Math.Max(0.001, sw1.Elapsed.TotalSeconds):F0} 封/s ({sw1.Elapsed.TotalSeconds:F1}s, 附件{att1}, 失敗{fail1})")

                           ' ── (2) 共用 _rdo, store-scoped (只換 resolve 形式, 同一 session) ──
                           Dim sw2 As New Stopwatch() : sw2.Start()
                           Dim att2 As Long = 0, fail2 As Integer = 0
                           For Each kv In groups
                               Dim store = FindStoreByPath(_rdo, kv.Key)
                               If store Is Nothing Then fail2 += kv.Value.Count : Continue For
                               For Each eid In kv.Value
                                   Try
                                       Dim c = readAttach(TryCast(store.GetMessageFromID(eid), Redemption.RDOMail))
                                       If c < 0 Then fail2 += 1 Else att2 += c
                                   Catch : fail2 += 1
                                   End Try
                               Next
                           Next
                           sw2.Stop()
                           _dbg(" │(2)", $"共用_rdo store-scoped: {sample.Count / Math.Max(0.001, sw2.Elapsed.TotalSeconds):F0} 封/s ({sw2.Elapsed.TotalSeconds:F1}s, 附件{att2}, 失敗{fail2})")

                           ' ── (3) 獨立 session, store-scoped (= P3 形式) ──
                           Dim sess As Redemption.RDOSession = Nothing
                           Try
                               sess = New Redemption.RDOSession()
                               sess.Logon(profileName, "", False, True)
                               Dim sw3 As New Stopwatch() : sw3.Start()
                               Dim att3 As Long = 0, fail3 As Integer = 0
                               For Each kv In groups
                                   Dim store = FindStoreByPath(sess, kv.Key)
                                   If store Is Nothing Then fail3 += kv.Value.Count : Continue For
                                   For Each eid In kv.Value
                                       Try
                                           Dim c = readAttach(TryCast(store.GetMessageFromID(eid), Redemption.RDOMail))
                                           If c < 0 Then fail3 += 1 Else att3 += c
                                       Catch : fail3 += 1
                                       End Try
                                   Next
                               Next
                               sw3.Stop()
                               _dbg(" │(3)", $"獨立session store-scoped: {sample.Count / Math.Max(0.001, sw3.Elapsed.TotalSeconds):F0} 封/s ({sw3.Elapsed.TotalSeconds:F1}s, 附件{att3}, 失敗{fail3})")
                           Catch ex As System.Exception
                               _dbg(" │(3)✗", ex.GetBaseException().Message)
                           Finally
                               If sess IsNot Nothing Then
                                   Try : sess.Logoff() : Catch : End Try
                                   TryMarshalRelease(sess)
                               End If
                           End Try
                       End Sub)
        _dbg("B", "===== 對照結束, 請貼回(三個附件數應一致才公平) =====")
    End Function    ' 2026/6/23, 修改P3, 開始比較獨立session 形式對效能的影響倍數, 與平行度效能吞吐量測試
    Private Async Function SpikeBodyResolveCompare() As Task
        ' 2026/06/22 by Simon/Claude Opus 4.8: 拋棄式 spike B-內文版 — 驗證「內文讀取換獨立 session 是否也有 ~10×」。
        '   注意: 內文 production 路徑(GetMailBodyL3 第2190行)走 OOM, 不是 _rdo, 故基準與附件版不同, 測三條:
        '     (1) OOM _olNS.GetItemFromID + .Body  = 內文現行 production 基準(你說的 70~80 封/s 來源)
        '     (2) 共用 _rdo store-scoped + .Body    → (2)vs(3) 對照「共用 vs 獨立 session」這條槓桿在內文是否成立
        '     (3) 獨立 session store-scoped + .Body = 目標形式
        '   防 IRM: 取樣時用 RDO 預掃 MessageClass, 跳過 IPM.Note.* 受保護(rpmsg)信, 避免 OOM .Body 卡死授權 modal。
        '   ★全程 UI/STA 緒同步跑: OOM COM 不可進 Task.Run; N=1000 單執行緒, UI 短暫凍結可接受。
        '   依賴: FindStoreByPath(P4 放的)。前提: Outlook 切 Work profile。測完即整段刪除。
        Const N As Integer = 1000
        If _rdo Is Nothing Then Await InitRdoSessionWithoutEULA()
        If _rdo Is Nothing Then _dbg("B內文", "Redemption 初始化失敗, 中止") : Return
        Dim profileName As String = ""
        Try : profileName = CStr(CallByName(_rdo, "ProfileName", CallType.Get)) : Catch : End Try
        _dbg("B內文", $"===== 內文 resolve 形式對照 (profile=[{profileName}], N={N}, 單執行緒/UI緒) =====")

        Await Task.Run(Sub()
                           ' ── 1. 用共用 _rdo 走訪頭部湊 N 封, 防 IRM: 跳過受保護信(MessageClass 含 .rpmsg 或非 IPM.Note 之保護類) ──
                           Dim sample As New List(Of (eid As String, pstPath As String))()
                           Dim skipIrm As Integer = 0
                           Try
                               For si As Integer = 1 To _rdo.Stores.Count
                                   If sample.Count >= N Then Exit For
                                   Dim st = _rdo.Stores.Item(si)
                                   Dim pp As String = "" : Try : pp = CStr(CallByName(st, "PstPath", CallType.Get)) : Catch : End Try
                                   If pp = "" Then Continue For
                                   Try
                                       Dim stk As New Stack(Of Redemption.RDOFolder)() : stk.Push(st.IPMRootFolder)
                                       Do While stk.Count > 0 AndAlso sample.Count < N
                                           Dim f = stk.Pop()
                                           For ii As Integer = 1 To f.Items.Count
                                               If sample.Count >= N Then Exit For
                                               Dim mm = TryCast(f.Items.Item(ii), Redemption.RDOMail)
                                               If mm Is Nothing Then Continue For
                                               Dim mc As String = "" : Try : mc = CStr(mm.MessageClass) : Catch : End Try
                                               ' IRM/RMS 保護信外層 MessageClass 多為 IPM.Note.SMIME 或含 rpmsg; 保守只收純 IPM.Note
                                               If mc.StartsWith("IPM.Note", StringComparison.OrdinalIgnoreCase) AndAlso
                               mc.IndexOf("rpmsg", StringComparison.OrdinalIgnoreCase) < 0 AndAlso
                               mc.IndexOf("SMIME", StringComparison.OrdinalIgnoreCase) < 0 Then
                                                   sample.Add((mm.EntryID, pp))
                                               Else
                                                   skipIrm += 1
                                               End If
                                           Next
                                           For fi As Integer = 1 To f.Folders.Count : stk.Push(f.Folders.Item(fi)) : Next
                                       Loop
                                   Catch : End Try
                               Next
                           Catch ex As System.Exception
                               _dbg(" │收集✗", ex.GetBaseException().Message)
                           End Try
                           If sample.Count = 0 Then _dbg(" │✗", "沒取到信, 中止") : Return

                           ' 按 pstPath 分組(供 (2)(3) store-scoped 重用 store)
                           Dim groups As New Dictionary(Of String, List(Of String))()
                           For Each s In sample
                               Dim lst As List(Of String) = Nothing
                               If Not groups.TryGetValue(s.pstPath, lst) Then lst = New List(Of String)() : groups(s.pstPath) = lst
                               lst.Add(s.eid)
                           Next
                           _dbg(" │收集", $"取樣 {sample.Count} 封(跨 {groups.Count} 個 PST), 跳過疑似IRM {skipIrm} 封")

                           ' ── (2) 共用 _rdo store-scoped + .Body ──
                           Dim sw2 As New Stopwatch() : sw2.Start()
                           Dim chars2 As Long = 0, fail2 As Integer = 0
                           For Each kv In groups
                               Dim store = FindStoreByPath(_rdo, kv.Key)
                               If store Is Nothing Then fail2 += kv.Value.Count : Continue For
                               For Each eid In kv.Value
                                   Try
                                       Dim rm = TryCast(store.GetMessageFromID(eid), Redemption.RDOMail)
                                       If rm Is Nothing Then fail2 += 1 : Continue For
                                       Dim b As String = rm.Body : If b IsNot Nothing Then chars2 += b.Length
                                   Catch : fail2 += 1
                                   End Try
                               Next
                           Next
                           sw2.Stop()
                           _dbg(" │(2)", $"共用_rdo .Body: {sample.Count / Math.Max(0.001, sw2.Elapsed.TotalSeconds):F0} 封/s ({sw2.Elapsed.TotalSeconds:F1}s, 字元{chars2}, 失敗{fail2})")

                           ' ── (3) 獨立 session store-scoped + .Body (RDO 在背景緒 OK, 但本支求一致仍在 UI 緒同步跑) ──
                           Dim sess As Redemption.RDOSession = Nothing
                           Try
                               sess = New Redemption.RDOSession()
                               sess.Logon(profileName, "", False, True)
                               Dim sw3 As New Stopwatch() : sw3.Start()
                               Dim chars3 As Long = 0, fail3 As Integer = 0
                               For Each kv In groups
                                   Dim store = FindStoreByPath(sess, kv.Key)
                                   If store Is Nothing Then fail3 += kv.Value.Count : Continue For
                                   For Each eid In kv.Value
                                       Try
                                           Dim rm = TryCast(store.GetMessageFromID(eid), Redemption.RDOMail)
                                           If rm Is Nothing Then fail3 += 1 : Continue For
                                           Dim b As String = rm.Body : If b IsNot Nothing Then chars3 += b.Length
                                       Catch : fail3 += 1
                                       End Try
                                   Next
                               Next
                               sw3.Stop()
                               _dbg(" │(3)", $"獨立session .Body: {sample.Count / Math.Max(0.001, sw3.Elapsed.TotalSeconds):F0} 封/s ({sw3.Elapsed.TotalSeconds:F1}s, 字元{chars3}, 失敗{fail3})")
                           Catch ex As System.Exception
                               _dbg(" │(3)✗", ex.GetBaseException().Message)
                           Finally
                               If sess IsNot Nothing Then
                                   Try : sess.Logoff() : Catch : End Try
                                   TryMarshalRelease(sess)
                               End If
                           End Try
                           _dbg("B內文", "===== 對照結束, 請貼回(三個字元數應相近才公平) =====")

                       End Sub)
    End Function    ' 驗證「內文讀取換獨立 session 效能與平行度效能吞吐量測試」
    Private Sub DumpResolve(tag As String, sess As Redemption.RDOSession, store As Redemption.RDOStore, eids As List(Of String), storeEid As String)
        Dim okA As Integer = 0, okB As Integer = 0, okC As Integer = 0
        Dim eA As String = "", eB As String = "", eC As String = ""
        For Each eid As String In eids
            Try
                If TryCast(sess.GetMessageFromID(eid), Redemption.RDOMail) IsNot Nothing Then okA += 1
            Catch ex As System.Exception
                If eA = "" Then eA = ex.GetBaseException().Message
            End Try
            Try
                If TryCast(sess.GetMessageFromID(eid, storeEid), Redemption.RDOMail) IsNot Nothing Then okB += 1
            Catch ex As System.Exception
                If eB = "" Then eB = ex.GetBaseException().Message
            End Try
            Try
                If store IsNot Nothing AndAlso TryCast(store.GetMessageFromID(eid), Redemption.RDOMail) IsNot Nothing Then okC += 1
            Catch ex As System.Exception
                If eC = "" Then eC = ex.GetBaseException().Message
            End Try
        Next
        _dbg($" │{tag}", $"(a)單參數={okA}/{eids.Count} [{eA}]　(b)雙參數={okB}/{eids.Count} [{eB}]　(c)store-scoped={okC}/{eids.Count} store={store IsNot Nothing} [{eC}]")
    End Sub ' P4 輔助: 對同一批 EntryID 試三種 resolve 形式, 各記成功數與首個例外
    Private Function FindStoreByPath(sess As Redemption.RDOSession, path As String) As Redemption.RDOStore
        If path = "" Then Return Nothing
        For i As Integer = 1 To sess.Stores.Count
            Dim pp As String = ""
            Try : pp = CStr(CallByName(sess.Stores.Item(i), "PstPath", CallType.Get)) : Catch : End Try
            If String.Equals(pp, path, StringComparison.OrdinalIgnoreCase) Then Return sess.Stores.Item(i)
        Next
        Return Nothing
    End Function ' P4 輔助: 用 PstPath 在指定 session 找 RDOStore

    ' 2026/06/23 驗證獨立 session _rdo2
    Private Async Function SpikeResolveFormOnRdo2() As Task
        ' =================================================================
        ' 2026/06/23 by Simon/Claude: 探針 — 驗證獨立 session _rdo2 的 resolve 形式
        '   目的: 用 OOM 取得的 (EntryID, OOM StoreID, FolderPath) 在 _rdo2 上分別試三種
        '         resolve, 決定 production 該走「雙參數」還是「store-scoped」。
        '   判讀: 看哪種形式 resolve 成功率高、且 Subject 對得上 (= 真解到, 非空 handle)。
        '   ※ 純診斷, 不動 production; 用完即可整段刪除。
        ' =================================================================
        If _rdo2 Is Nothing Then Await InitRdoSessionWithoutEULA()
        If _rdo2 Is Nothing Then _dbg("探針中止", "_rdo2 初始化失敗") : Return

        ' ── 1. 印 _rdo2 身分 (確認登對 profile、看得到哪些 store) ──
        Dim storeNames As New List(Of String)
        Try
            For i As Integer = 1 To _rdo2.Stores.Count : storeNames.Add(_rdo2.Stores.Item(i).Name) : Next
        Catch ex As System.Exception
            _dbg("探針", $"列舉 _rdo2.Stores 失敗: {ex.Message}")
        End Try
        _dbg("探針 _rdo2", $"ProfileName=[{_rdo2.ProfileName}] Stores={storeNames.Count}")
        _dbg("探針 _rdo2 stores", String.Join(" | ", storeNames))

        ' ── 2. 從 OOM 採樣: 最多 3 個 PST、每 PST 最多 4 封, 合計上限 ~12 ──
        Dim samples As New List(Of (eid As String, sid As String, fpath As String, subj As String))
        Dim storeTaken As Integer = 0
        For si As Integer = 1 To _olNS.Stores.Count
            If storeTaken >= 3 Then Exit For
            Dim st As Outlook.Store = Nothing
            Try
                st = _olNS.Stores.Item(si)
                If String.IsNullOrEmpty(st.FilePath) Then Continue For   ' 跳過無檔 store (iCloud 等)
                Dim grabbed As Integer = HarvestFromStore(st, st.StoreID, samples, 4)
                If grabbed > 0 Then storeTaken += 1
            Catch ex As System.Exception
                _dbg("探針採樣", $"store#{si} 失敗: {ex.Message}")
            Finally
                TryMarshalRelease(st)
            End Try
        Next
        _dbg("探針採樣", $"共取得 {samples.Count} 封樣本 (跨 {storeTaken} 個 PST)")
        If samples.Count = 0 Then _dbg("探針中止", "採樣 0 封") : Return

        ' ── 3. 三種形式逐封測試 ──
        Dim ok1, ok2, ok3, match1, match2, match3 As Integer
        Dim err1 As String = "", err2 As String = "", err3 As String = ""
        For Each s In samples
            ' (1) 單參數 (預期跨 session 失敗, 當 baseline)
            Dim m1 As Redemption.RDOMail = Nothing
            Try
                m1 = TryCast(_rdo2.GetMessageFromID(s.eid), Redemption.RDOMail)
                If m1 IsNot Nothing Then ok1 += 1 : If m1.Subject = s.subj Then match1 += 1
            Catch ex As System.Exception
                If err1 = "" Then err1 = ex.Message
            Finally
                TryMarshalRelease(m1)
            End Try
            ' (2) 雙參數 + OOM StoreID
            Dim m2 As Redemption.RDOMail = Nothing
            Try
                m2 = TryCast(_rdo2.GetMessageFromID(s.eid, s.sid), Redemption.RDOMail)
                If m2 IsNot Nothing Then ok2 += 1 : If m2.Subject = s.subj Then match2 += 1
            Catch ex As System.Exception
                If err2 = "" Then err2 = ex.Message
            Finally
                TryMarshalRelease(m2)
            End Try
            ' (3) store-scoped (依 FolderPath 取 store 名, 在 _rdo2.Stores 找 RDOStore)
            Dim m3 As Redemption.RDOMail = Nothing
            Dim rstore As Redemption.RDOStore = Nothing
            Try
                Dim wantName As String = GetStoreNameFromPath(s.fpath)
                For i As Integer = 1 To _rdo2.Stores.Count
                    Dim cand As Redemption.RDOStore = _rdo2.Stores.Item(i)
                    If cand.Name = wantName Then rstore = cand : Exit For
                    TryMarshalRelease(cand)
                Next
                If rstore IsNot Nothing Then
                    m3 = TryCast(rstore.GetMessageFromID(s.eid), Redemption.RDOMail)
                    If m3 IsNot Nothing Then ok3 += 1 : If m3.Subject = s.subj Then match3 += 1
                Else
                    If err3 = "" Then err3 = $"_rdo2.Stores 找不到 [{wantName}]"
                End If
            Catch ex As System.Exception
                If err3 = "" Then err3 = ex.Message
            Finally
                TryMarshalRelease(m3)
                TryMarshalRelease(rstore)
            End Try
        Next

        ' ── 4. 總結 ──
        Dim n As Integer = samples.Count
        _dbg("探針結果 (1)單參數", $"resolve {ok1}/{n}, subject吻合 {match1}/{n}{If(err1 = "", "", " | err: " & err1)}")
        _dbg("探針結果 (2)雙參數+OOM StoreID", $"resolve {ok2}/{n}, subject吻合 {match2}/{n}{If(err2 = "", "", " | err: " & err2)}")
        _dbg("探針結果 (3)store-scoped", $"resolve {ok3}/{n}, subject吻合 {match3}/{n}{If(err3 = "", "", " | err: " & err3)}")
    End Function   ' 驗證獨立 session _rdo2 的 resolve 形式
    Private Async Function SpikeResolveFolderOnRdo2() As Task
        ' =================================================================
        ' 2026/06/23 by Simon/Claude Opus 4.8: 探針 — 驗證 _rdo2 的 FOLDER resolve 形式
        '   目的: 用 OOM 取得的 (folder EntryID, OOM StoreID, FolderPath) 在 _rdo2 上試三種 resolve,
        '         決定 GetMailCountRdo/GetFolderCountRdo 該走「store-scoped 單參數」還是「雙參數」。
        '   判讀: 看哪種 resolve 成功率高、且 .Name 對得上 (= 真解到 folder, 非空 handle)。
        '   ※ 純診斷, 不動 production; 用完即可整段刪除。(對照 SpikeResolveFormOnRdo2 的 message 版)
        ' =================================================================
        If _rdo2 Is Nothing Then Await InitRdoSessionWithoutEULA()
        If _rdo2 Is Nothing Then _dbg("Folder探針中止", "_rdo2 初始化失敗") : Return

        ' ── 1. 從 OOM 採樣 folder: 最多 3 個 PST、每 PST 最多 4 個夾, 合計上限 ~12 ──
        Dim samples As New List(Of (eid As String, sid As String, fpath As String, name As String))
        Dim storeTaken As Integer = 0
        For si As Integer = 1 To _olNS.Stores.Count
            If storeTaken >= 3 Then Exit For
            Dim st As Outlook.Store = Nothing
            Try
                st = _olNS.Stores.Item(si)
                If String.IsNullOrEmpty(st.FilePath) Then Continue For
                Dim grabbed As Integer = HarvestFoldersFromStore(st, st.StoreID, samples, 4)
                If grabbed > 0 Then storeTaken += 1
            Catch ex As System.Exception
                _dbg("Folder探針採樣", $"store#{si} 失敗: {ex.Message}")
            Finally
                TryMarshalRelease(st)
            End Try
        Next
        _dbg("Folder探針採樣", $"共取得 {samples.Count} 個夾 (跨 {storeTaken} 個 PST)")
        If samples.Count = 0 Then _dbg("Folder探針中止", "採樣 0 個夾") : Return

        ' ── 2. 三種形式逐夾測試 ──
        Dim ok1, ok2, ok3, match1, match2, match3 As Integer
        Dim err1 As String = "", err2 As String = "", err3 As String = ""
        For Each s In samples
            ' (1) 單參數 session 級 (預期跨 session 失敗, baseline)
            Dim f1 As Redemption.RDOFolder = Nothing
            Try
                f1 = TryCast(_rdo2.GetFolderFromID(s.eid), Redemption.RDOFolder)
                If f1 IsNot Nothing Then ok1 += 1 : If f1.Name = s.name Then match1 += 1
            Catch ex As System.Exception
                If err1 = "" Then err1 = ex.Message
            Finally
                Dim o As Object = f1 : TryMarshalRelease(o)
            End Try
            ' (2) 雙參數 + OOM StoreID
            Dim f2 As Redemption.RDOFolder = Nothing
            Try
                f2 = TryCast(_rdo2.GetFolderFromID(s.eid, s.sid), Redemption.RDOFolder)
                If f2 IsNot Nothing Then ok2 += 1 : If f2.Name = s.name Then match2 += 1
            Catch ex As System.Exception
                If err2 = "" Then err2 = ex.Message
            Finally
                Dim o As Object = f2 : TryMarshalRelease(o)
            End Try
            ' (3) store-scoped (依 FolderPath 取 store, store.GetFolderFromID(eid)) — production 目標路徑
            Dim f3 As Redemption.RDOFolder = Nothing
            Dim rstore As Redemption.RDOStore = GetRdoStore(s.fpath)
            Try
                If rstore IsNot Nothing Then
                    f3 = TryCast(rstore.GetFolderFromID(s.eid), Redemption.RDOFolder)
                    If f3 IsNot Nothing Then ok3 += 1 : If f3.Name = s.name Then match3 += 1
                Else
                    If err3 = "" Then err3 = $"GetRdo2Store 找不到 store for [{s.fpath}]"
                End If
            Catch ex As System.Exception
                If err3 = "" Then err3 = ex.Message
            Finally
                Dim o As Object = f3 : TryMarshalRelease(o)   ' rstore 為 byName 參考,不在此釋放
            End Try
        Next

        ' ── 3. 總結 ──
        Dim n As Integer = samples.Count
        _dbg("Folder探針 (1)單參數", $"resolve {ok1}/{n}, name吻合 {match1}/{n}{If(err1 = "", "", " | err: " & err1)}")
        _dbg("Folder探針 (2)雙參數+OOM StoreID", $"resolve {ok2}/{n}, name吻合 {match2}/{n}{If(err2 = "", "", " | err: " & err2)}")
        _dbg("Folder探針 (3)store-scoped", $"resolve {ok3}/{n}, name吻合 {match3}/{n}{If(err3 = "", "", " | err: " & err3)}")
    End Function ' 驗證 _rdo2 的 FOLDER resolve 形式
    Private Function HarvestFromStore(st As Outlook.Store, sid As String, samples As List(Of (eid As String, sid As String, fpath As String, subj As String)), maxN As Integer) As Integer
        ' 探針輔助: 從單一 OOM store BFS 抓最多 maxN 封 (只讀 EntryID/Subject, 不碰 .Body/.Attachments 故不撞 IRM)
        Dim taken As Integer = 0
        Dim root As Outlook.Folder = Nothing
        Dim queue As New Queue(Of Outlook.Folder)()
        Try
            root = TryCast(st.GetRootFolder(), Outlook.Folder)
            If root Is Nothing Then Return 0
            queue.Enqueue(root) : root = Nothing      ' 交給 queue 統一釋放
            Dim visited As Integer = 0
            While queue.Count > 0 AndAlso taken < maxN AndAlso visited < 60
                Dim f As Outlook.Folder = queue.Dequeue()
                visited += 1
                Try
                    Dim items As Outlook.Items = f.Items
                    Dim cnt As Integer = items.Count
                    Dim fpath As String = f.FolderPath
                    Dim i As Integer = 1
                    While i <= cnt AndAlso taken < maxN
                        Dim it As Object = items.Item(i)
                        Try
                            Dim eid As String = CStr(CallByName(it, "EntryID", CallType.Get))
                            Dim subj As String = CStr(CallByName(it, "Subject", CallType.Get))
                            If Not String.IsNullOrEmpty(eid) Then samples.Add((eid, sid, fpath, subj)) : taken += 1
                        Catch
                            ' 非郵件項目或讀取失敗, 略過
                        Finally
                            TryMarshalRelease(it)
                        End Try
                        i += 1
                    End While
                    For sfi As Integer = 1 To f.Folders.Count : queue.Enqueue(f.Folders.Item(sfi)) : Next
                    TryMarshalRelease(items)
                Catch
                    ' 該夾讀取失敗, 略過
                Finally
                    TryMarshalRelease(f)
                End Try
            End While
        Catch ex As System.Exception
            _dbg("探針採樣", $"HarvestFromStore 失敗: {ex.Message}")
        Finally
            TryMarshalRelease(root)
            While queue.Count > 0 : TryMarshalRelease(queue.Dequeue()) : End While   ' 排空殘留子夾
        End Try
        Return taken
    End Function
    Private Function HarvestFoldersFromStore(st As Outlook.Store, sid As String, samples As List(Of (eid As String, sid As String, fpath As String, Name As String)), maxN As Integer) As Integer
        ' 探針輔助: 從單一 OOM store BFS 抓最多 maxN 個子夾 (只讀 EntryID/Name, 不碰 Items 故極輕量)
        Dim taken As Integer = 0
        Dim root As Outlook.Folder = Nothing
        Dim queue As New Queue(Of Outlook.Folder)()
        Try
            root = TryCast(st.GetRootFolder(), Outlook.Folder)
            If root Is Nothing Then Return 0
            queue.Enqueue(root) : root = Nothing
            Dim visited As Integer = 0
            While queue.Count > 0 AndAlso taken < maxN AndAlso visited < 60
                Dim f As Outlook.Folder = queue.Dequeue()
                visited += 1
                Try
                    Try
                        If Not String.IsNullOrEmpty(f.EntryID) Then samples.Add((f.EntryID, sid, f.FolderPath, f.Name)) : taken += 1
                    Catch
                    End Try
                    Dim subs As Outlook.Folders = f.Folders
                    Try
                        For Each sf As Outlook.Folder In subs
                            If queue.Count < 60 Then queue.Enqueue(sf) Else TryMarshalRelease(sf)
                        Next
                    Finally
                        TryMarshalRelease(subs)
                    End Try
                Finally
                    TryMarshalRelease(f)
                End Try
            End While
        Catch ex As System.Exception
            _dbg("HarvestFolders", $"{ex.Message}")
        End Try
        Return taken
    End Function

    Private Async Function SpikeFolderVisibilityCompare() As Task
        ' 探針一: SpikeFolderVisibilityCompare — RDO vs OOM 全枚舉夾清單差集 + 隱藏判據 dump
        ' 2026/06/23 by Simon/Claude Opus 4.8: 補 _rdoFastPath 的 visibility 技術債。
        '   目的: 找出 RDO 枚舉多撈、但 OOM 看不到的夾(實測曾 27 vs 22)，並 dump 其判據
        '         (Kind / PR_CONTAINER_CLASS / PR_ATTR_HIDDEN)，決定 isRDO 旗標的判斷規則。
        '   非破壞性: 只枚舉讀取，不寫任何快取、不改任何夾。測完可整段刪除。
        '   前提: 跑前把 Outlook 切到要測的 profile (Work 27 PST)。RDO 用獨立 _rdo2 不污染 _rdo。
        ' ════════════════════════════════════════════════════════════════════════
        Const PR_CONTAINER_CLASS As String = "http://schemas.microsoft.com/mapi/proptag/0x3613001E"
        Const PR_ATTR_HIDDEN As String = "http://schemas.microsoft.com/mapi/proptag/0x10F4000B"

        If _rdo2 Is Nothing Then Await InitRdoSessionWithoutEULA()  ' ← 若你的 _rdo2 初始化函數名不同，改這行
        If _rdo2 Is Nothing Then _dbg("VisCmp", "_rdo2 初始化失敗, 中止") : Return
        If _olNS Is Nothing Then _dbg("VisCmp", "_olNS 為空, 中止") : Return

        _dbg("VisCmp", "═════ RDO vs OOM 全枚舉差集 開始 ═════")

        Await Task.Run(
            Sub()
                ' ── 1. OOM 端: 逐 store BFS 枚舉 .Folders，收 FolderPath 集合 ──
                Dim oomPaths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                Dim oomByStore As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
                Try
                    For Each st As Outlook.Store In _olNS.Stores
                        Dim stName As String = "" : Try : stName = st.DisplayName : Catch : End Try
                        Dim root As Outlook.Folder = Nothing
                        Try : root = TryCast(st.GetRootFolder(), Outlook.Folder) : Catch : End Try
                        If root Is Nothing Then Continue For
                        Dim before As Integer = oomPaths.Count
                        Dim stk As New Stack(Of Outlook.Folder)() : stk.Push(root)
                        Do While stk.Count > 0
                            Dim f = stk.Pop()
                            Dim p As String = "" : Try : p = f.FolderPath : Catch : End Try
                            If p <> "" Then oomPaths.Add(p)
                            Try
                                For i As Integer = 1 To f.Folders.Count
                                    stk.Push(TryCast(f.Folders.Item(i), Outlook.Folder))
                                Next
                            Catch : End Try
                        Loop
                        oomByStore(stName) = oomPaths.Count - before
                    Next
                Catch ex As System.Exception
                    _dbg("VisCmp", "OOM 枚舉例外: " & ex.GetBaseException().Message)
                End Try
                _dbg("VisCmp", $"OOM 可見夾總數 = {oomPaths.Count}")
                For Each kv In oomByStore : _dbg(" │OOM", $"[{kv.Key}] {kv.Value} 夾") : Next

                ' ── 2. RDO 端(_rdo2): 逐 store BFS 枚舉 .Folders，收 FolderPath 集合 ──
                '    同時記下每夾的判據, 供差集 dump
                ' ── 2. RDO 端(_rdo2): 逐 store BFS 枚舉 .Folders，收 FolderPath 集合 ──
                '    2026/06/23 by Simon/Claude: 改用 IPMRootFolder(IPM 樹根)當起點。
                '      假設: search folder/系統夾在 IPM 樹外, 用 IPMRootFolder 枚舉天生不會撈到,
                '      集合應 = OOM 可見的 822。若差集歸零即證實「從源頭用 IPMRootFolder」可行。
                Dim rdoPaths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                Dim rdoInfo As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) ' path → 判據字串
                Try
                    For si As Integer = 1 To _rdo2.Stores.Count
                        Dim st = _rdo2.Stores.Item(si)
                        Dim stName As String = "" : Try : stName = st.Name : Catch : End Try
                        Dim root As Redemption.RDOFolder = Nothing
                        Try : root = st.IPMRootFolder : Catch : End Try    ' ← 改: RootFolder → IPMRootFolder
                        If root Is Nothing Then Continue For
                        Dim stk As New Stack(Of Redemption.RDOFolder)() : stk.Push(root)
                        Do While stk.Count > 0
                            Dim f = stk.Pop()
                            Dim p As String = "" : Try : p = f.FolderPath : Catch : End Try
                            If p <> "" Then
                                rdoPaths.Add(p)
                                ' dump 判據(讀法修正): Kind 直接取列舉轉 Integer, 不套 CallByName+CStr
                                Dim kind As String = "?"
                                Try : kind = CInt(f.Kind).ToString() : Catch : kind = "?" : End Try
                                Dim cclass As String = "" : Try : cclass = CStr(f.Fields(PR_CONTAINER_CLASS)) : Catch : cclass = "" : End Try
                                Dim hidden As String = "?" : Try : hidden = CStr(f.Fields(PR_ATTR_HIDDEN)) : Catch : hidden = "?" : End Try
                                rdoInfo(p) = $"Kind={kind}, Class=[{cclass}], Hidden={hidden}"
                            End If
                            Try
                                For i As Integer = 1 To f.Folders.Count
                                    stk.Push(f.Folders.Item(i))
                                Next
                            Catch : End Try
                        Loop
                    Next
                Catch ex As System.Exception
                    _dbg("VisCmp", "RDO 枚舉例外: " & ex.GetBaseException().Message)
                End Try
                _dbg("VisCmp", $"RDO(_rdo2, IPMRootFolder) 枚舉夾總數 = {rdoPaths.Count}")

                ' ── 3. 差集 ──
                Dim rdoOnly = rdoPaths.Where(Function(p) Not oomPaths.Contains(p)).OrderBy(Function(p) p).ToList()
                Dim oomOnly = oomPaths.Where(Function(p) Not rdoPaths.Contains(p)).OrderBy(Function(p) p).ToList()

                _dbg("VisCmp", $"═════ RDO-only(RDO有 OOM無) 共 {rdoOnly.Count} 個 ═════")
                For Each p In rdoOnly
                    Dim info As String = "" : rdoInfo.TryGetValue(p, info)
                    _dbg(" │RDO-only", $"{p}  ←  {info}")
                Next
                _dbg("VisCmp", $"═════ OOM-only(OOM有 RDO無) 共 {oomOnly.Count} 個 ═════")
                For Each p In oomOnly
                    _dbg(" │OOM-only", p)
                Next
                _dbg("VisCmp", "═════ 結束, 請貼回 RDO-only 清單與判據 ═════")
            End Sub)
    End Function ' RDO vs OOM 全枚舉夾清單差集 + 隱藏判據 dump
    Private Async Function SpikeFolderTableBenchmark() As Task
        ' 探針二: SpikeFolderTableBenchmark — 單夾 GetTable 的 OOM vs RDO parity + 分段計時
        '         + 平行 K=1/2/4 × {共用_rdo2 / 各自獨立session} 對照
        ' 2026/06/23 by Simon/Claude Opus 4.8: 改自 SpikeParallelReadBenchmark(P3)。
        '   回答三問: (A)單夾 GetTable 分段耗時瓶頸在哪 (B)RDO MAPITable 與 OOM GetTable 列數 parity
        '            (C)平行值不值得 + worker 共用一條 _rdo2 是否可行/掉速 vs 各自獨立 session(實測不預防)。
        '   非破壞性: 只讀不寫。測完可整段刪除。前提: Work profile, 已勾 CheckRDO 使 _rdo2 在。
        ' ════════════════════════════════════════════════════════════════════════
        Const PR_MESSAGE_SIZE As String = "http://schemas.microsoft.com/mapi/proptag/0x0E080003"
        Const PR_SENDER_EMAIL As String = "http://schemas.microsoft.com/mapi/proptag/0x0C1F001E"
        Const PR_INTERNET_MESSAGE_ID As String = "http://schemas.microsoft.com/mapi/proptag/0x1035001E"
        Dim cols = {"EntryID", "Subject", PR_MESSAGE_SIZE, "ReceivedTime", "SenderName", PR_INTERNET_MESSAGE_ID, PR_SENDER_EMAIL}

        If _rdo2 Is Nothing Then _dbg("TblBM", "_rdo2 為空(請先勾 CheckRDO), 中止") : Return
        If _olNS Is Nothing Then _dbg("TblBM", "_olNS 為空, 中止") : Return

        Dim profileName As String = ""
        Try : profileName = CStr(CallByName(_rdo2, "ProfileName", CallType.Get)) : Catch : End Try
        _dbg("TblBM", $"═════ 開始 (profile=[{profileName}]) ═════")

        ' ── 收集標的: OOM 走訪挑 >= MINROWS 封的夾, 收 (FolderPath, OOM Folder 物件) ──
        Const MINROWS As Integer = 500
        Const MAXFOLDERS As Integer = 8
        Dim targets As New List(Of (path As String, oomFolder As Outlook.Folder))()
        Try
            For Each st As Outlook.Store In _olNS.Stores
                If targets.Count >= MAXFOLDERS Then Exit For
                Dim root As Outlook.Folder = TryCast(st.GetRootFolder(), Outlook.Folder)
                If root Is Nothing Then Continue For
                Dim stk As New Stack(Of Outlook.Folder)() : stk.Push(root)
                Do While stk.Count > 0 AndAlso targets.Count < MAXFOLDERS
                    Dim f = stk.Pop()
                    Dim cnt As Integer = 0 : Try : cnt = f.Items.Count : Catch : End Try
                    If cnt >= MINROWS Then targets.Add((f.FolderPath, f))
                    Try
                        For i As Integer = 1 To f.Folders.Count : stk.Push(TryCast(f.Folders.Item(i), Outlook.Folder)) : Next
                    Catch : End Try
                Loop
            Next
        Catch : End Try
        If targets.Count = 0 Then _dbg("TblBM", $"找不到 >= {MINROWS} 封的夾, 中止") : Return
        _dbg("TblBM", $"標的夾 {targets.Count} 個 (每個 >= {MINROWS} 封)")

        ' ════ A: OOM GetTable 分段計時 ════
        _dbg("TblBM", "───── A: OOM GetTable 分段 ─────")
        For Each tg In targets
            Dim swPath As New Stopwatch(), swTable As New Stopwatch(), swArray As New Stopwatch()
            Dim rows As Integer = 0
            Try
                swPath.Start() : Dim p As String = tg.oomFolder.FolderPath : swPath.Stop()
                swTable.Start()
                Dim tbl As Outlook.Table = tg.oomFolder.GetTable("", Outlook.OlTableContents.olUserItems)
                tbl.Columns.RemoveAll()
                For Each c In cols : tbl.Columns.Add(c) : Next
                swTable.Stop()
                swArray.Start()
                Do While Not tbl.EndOfTable
                    Dim arr = tbl.GetArray(500)
                    If arr Is Nothing Then Exit Do
                    rows += arr.GetUpperBound(0) + 1
                Loop
                swArray.Stop()
            Catch ex As System.Exception
                _dbg(" │OOM", $"例外 {tg.path}: {ex.Message}") : Continue For
            End Try
            _dbg(" │OOM", $"[{ExtractFolderName(tg.path)}] rows={rows} | Path={swPath.ElapsedMilliseconds}ms Table={swTable.ElapsedMilliseconds}ms Array={swArray.ElapsedMilliseconds}ms")
        Next

        ' ════ B: RDO 列舉 Items(設 Columns 走 table 不開信) 分段計時 + parity ════
        _dbg("TblBM", "───── B: RDO 列舉 Items 分段 ─────")
        For Each tg In targets
            Dim swResolve As New Stopwatch(), swCols As New Stopwatch(), swRead As New Stopwatch()
            Dim rows As Integer = 0
            Try
                swResolve.Start()
                Dim rf As Redemption.RDOFolder = FolderPath2RdoFolder(_rdo2, tg.path)
                swResolve.Stop()
                If rf Is Nothing Then _dbg(" │RDO", $"解析失敗 {tg.path}") : Continue For

                Dim items As Redemption.RDOItems = rf.Items
                swCols.Start()
                ' 設 MAPITable.Columns: 設好後列舉 items 只讀這些欄、不開信 (官方 RDOItems 範例)
                Try
                    Dim mt As Object = items.MAPITable
                    mt.Columns.Clear()
                    For Each c In cols : mt.Columns.Add(c) : Next
                Catch exCol As System.Exception
                    _dbg(" │RDO", $"設 Columns 失敗 {tg.path}: {exCol.GetBaseException().Message}")
                End Try
                swCols.Stop()

                swRead.Start()
                For Each m As Redemption.RDOMail In items
                    Dim s As String = "" : Try : s = m.Subject : Catch : End Try   ' 觸發實際讀取(走 table)
                    rows += 1
                Next
                swRead.Stop()
            Catch ex As System.Exception
                _dbg(" │RDO", $"例外 {tg.path}: {ex.GetBaseException().Message}") : Continue For
            End Try
            _dbg(" │RDO", $"[{ExtractFolderName(tg.path)}] rows={rows} | Resolve={swResolve.ElapsedMilliseconds}ms Cols={swCols.ElapsedMilliseconds}ms Read={swRead.ElapsedMilliseconds}ms")
        Next


        ' ════ C: 平行 K=1/2/4 × {共用 _rdo2 / 各自獨立 session} ════
        _dbg("TblBM", "───── C: 平行對照 (workload=逐夾 列舉 Items 走 table) ─────")
        Dim allPaths = targets.Select(Function(t) t.path).ToList()
        For Each useShared In {True, False}
            Dim modeName As String = If(useShared, "共用_rdo2", "各自獨立session")
            For Each K In {1, 2, 4}
                Dim groups As New List(Of List(Of String))()
                For g = 0 To K - 1 : groups.Add(New List(Of String)) : Next
                For i = 0 To allPaths.Count - 1 : groups(i Mod K).Add(allPaths(i)) : Next

                Dim swWall As New Stopwatch() : swWall.Start()
                Dim tasks As New List(Of Task)()
                For g = 0 To K - 1
                    Dim myPaths = groups(g)
                    tasks.Add(Task.Run(
                        Sub()
                            Dim sess As Redemption.RDOSession = Nothing
                            Try
                                If useShared Then
                                    sess = _rdo2
                                Else
                                    sess = New Redemption.RDOSession()
                                    sess.Logon(profileName, "", False, True)
                                End If
                                For Each pth In myPaths
                                    Try
                                        Dim rf As Redemption.RDOFolder = FolderPath2RdoFolder(sess, pth)
                                        If rf Is Nothing Then Continue For
                                        Dim items As Redemption.RDOItems = rf.Items
                                        Try
                                            Dim mt As Object = items.MAPITable
                                            mt.Columns.Clear()
                                            For Each c In cols : mt.Columns.Add(c) : Next
                                        Catch : End Try
                                        For Each m As Redemption.RDOMail In items
                                            Dim s As String = "" : Try : s = m.Subject : Catch : End Try
                                        Next
                                    Catch : End Try
                                Next
                            Catch ex As System.Exception
                                _dbg(" │" & modeName, $"K={K} worker 例外: {ex.GetBaseException().Message}")
                            Finally
                                If Not useShared AndAlso sess IsNot Nothing Then
                                    Try : sess.Logoff() : Catch : End Try
                                    TryMarshalRelease(sess)
                                End If
                            End Try
                        End Sub))
                Next
                Await Task.WhenAll(tasks)
                swWall.Stop()
                _dbg(" │C", $"{modeName} K={K}: wall={swWall.ElapsedMilliseconds}ms ({allPaths.Count}夾)")
            Next
        Next
        _dbg("TblBM", "═════ 結束, 請貼回 ═════")

    End Function    ' 單夾 GetTable 的 OOM vs RDO parity + 分段計時
    Private Function FolderPath2RdoFolder(sess As Redemption.RDOSession, folderPath As String) As Redemption.RDOFolder
        ' ── 探針二專用小 helper: 在指定 session 上用 FolderPath 解出 RDOFolder ──
        ' 2026/06/23 by Simon/Claude: 拋棄式, 隨探針二刪除。
        '   策略: 先用 GetRdoStore 取 store(僅對 _rdo2 有效); 若傳入的是別條獨立 session,
        '   則退化為走訪該 session 的 Stores 找路徑開頭吻合者, 再 BFS 比對 FolderPath。
        Try
            ' 找 store: 路徑形如 \\store顯示名\夾\子夾, 取第一段比對 store.Name
            Dim trimmed As String = folderPath.TrimStart("\"c)
            Dim firstSeg As String = trimmed.Split("\"c)(0)
            Dim targetStore As Redemption.RDOStore = Nothing
            For si As Integer = 1 To sess.Stores.Count
                Dim st = sess.Stores.Item(si)
                Dim nm As String = "" : Try : nm = st.Name : Catch : End Try
                If String.Equals(nm, firstSeg, StringComparison.OrdinalIgnoreCase) Then targetStore = st : Exit For
            Next
            If targetStore Is Nothing Then Return Nothing
            ' 從 IPMRootFolder BFS 找 FolderPath 吻合
            Dim root As Redemption.RDOFolder = targetStore.IPMRootFolder
            Dim stk As New Stack(Of Redemption.RDOFolder)() : stk.Push(root)
            Do While stk.Count > 0
                Dim f = stk.Pop()
                Dim p As String = "" : Try : p = f.FolderPath : Catch : End Try
                If String.Equals(p, folderPath, StringComparison.OrdinalIgnoreCase) Then Return f
                Try
                    For i As Integer = 1 To f.Folders.Count : stk.Push(f.Folders.Item(i)) : Next
                Catch : End Try
            Loop
        Catch : End Try
        Return Nothing
    End Function ' 探針二專用小 helper: 在指定 session 上用 FolderPath 解出 RDOFolder

    ' 2026/06/24 by Simon/Claude Opus 4.8: 拋棄式探針 — 子樹階層走訪 OOM vs RDO批次 對拍
    '   本輪唯一目的: 先確認 API 讀法寫對 + 取得「暖快取」基準值(供 GetSubtreeListRdo 完工後比對是否有額外開銷)。
    '   標的: SimTree3.SelectedNodes 當 root(可多選逐一各跑;Simon 自行換不同深淺節點重跑)。
    '   對手(全單執行緒,全產出「子孫 path 集合」對拍):
    '     A  OOM        : current.Folders 逐夾 BFS(= GetSubtreeToListL3 去副作用版,基準)
    '     B  RDO-Enum   : rdoFolder.Folders For Each 逐夾(診斷: 隔離 RDO 層 vs OOM 層)
    '     C  RDO-Batch  : Folders.MAPITable.GetRows 整層批次,只對 PR_SUBFOLDERS=true 遞迴(候選)
    '     C+ RDO-Batch+CC: C 多撈 PR_CONTENT_COUNT(獨立計時,驗免費搭車且不污染 A/B/C)
    '   正確性對拍用 path 集合(最穩);EntryID 經 SpikeEidToHex 統一轉 hex 供遞迴。
    ' ============================================================================
    Private Sub SpikeSubtreeWalkCompare()
        Dim log As New List(Of String)
        If _rdo2 Is Nothing Then MessageBox.Show("_rdo2 未初始化,請先勾選 CheckRDO。") : Return
        Dim roots As List(Of TreeNode) = SimTree3.SelectedNodes
        If roots Is Nothing OrElse roots.Count = 0 Then MessageBox.Show("請先在 Tab3 的樹選定至少一個節點當 root。") : Return

        For Each node As TreeNode In roots
            Dim rootF As Folder = TryCast(node.Tag, Folder)
            If rootF Is Nothing Then Continue For
            Dim rootPath As String = SafeGetPath(rootF)
            Dim rootEid As String = "" : Try : rootEid = rootF.EntryID : Catch : End Try
            log.Add("══════ ROOT: " & ExtractFolderName(rootPath) & " ══════")
            log.Add("path = " & rootPath)

            Dim store As Redemption.RDOStore = GetRdoStore(rootPath)
            If store Is Nothing Then log.Add("✗ GetRdo2Store 失敗 → 跳過此 root 的 RDO 對手")

            ' ── 暖機一次(OOM)丟棄,讓後續對手吃同樣暖快取 ──
            Try : SpikeWalk_Oom(rootF, rootPath) : Catch : End Try

            Dim ra = SpikeWalk_Oom(rootF, rootPath)
            log.Add($"A  OOM        : {ra.paths.Count} 夾 | {ra.ms} ms")

            Dim rdoRoot As Redemption.RDOFolder = Nothing
            If store IsNot Nothing AndAlso rootEid <> "" Then
                Try : rdoRoot = TryCast(store.GetFolderFromID(rootEid), Redemption.RDOFolder)
                Catch ex As System.Exception : log.Add("✗ RDO root GetFolderFromID: " & ex.Message) : End Try
            End If

            If rdoRoot IsNot Nothing Then
                Try
                    Dim rb = SpikeWalk_RdoEnum(rdoRoot, rootPath)
                    log.Add($"B  RDO-Enum   : {rb.paths.Count} 夾 | {rb.ms} ms | vs A: {SpikeDiff(ra.paths, rb.paths)}")
                Catch ex As System.Exception : log.Add("✗ B 例外: " & ex.GetBaseException().Message) : End Try

                Try
                    Dim k As String = "?" : Dim cv As Integer = 0, ce As Integer = 0
                    Dim rc = SpikeWalk_RdoBatch(store, rdoRoot, rootPath, False, cv, ce, k)
                    log.Add($"C  RDO-Batch  : {rc.paths.Count} 夾 | {rc.ms} ms | vs A: {SpikeDiff(ra.paths, rc.paths)} | EntryID型別={k}")
                Catch ex As System.Exception : log.Add("✗ C 例外: " & ex.GetBaseException().Message) : End Try

                Try
                    Dim k As String = "?" : Dim cv As Integer = 0, ce As Integer = 0
                    Dim rcc = SpikeWalk_RdoBatch(store, rdoRoot, rootPath, True, cv, ce, k)
                    log.Add($"C+ RDO-Batch+CC: {rcc.paths.Count} 夾 | {rcc.ms} ms | PR_CONTENT_COUNT 有效={cv} 缺/錯={ce}")
                Catch ex As System.Exception : log.Add("✗ C+ 例外: " & ex.GetBaseException().Message) : End Try
            End If

            Dim o As Object = rdoRoot : TryMarshalRelease(o)
        Next

        For Each ln In log : _dbg("SubtreeSpike", ln) : Next
        MessageBox.Show(String.Join(vbCrLf, log), "子樹走訪對拍結果")
    End Sub

    Private Function SpikeDiff(a As HashSet(Of String), b As HashSet(Of String)) As String
        Dim ao = a.Where(Function(x) Not b.Contains(x)).Count()
        Dim bo = b.Where(Function(x) Not a.Contains(x)).Count()
        If ao = 0 AndAlso bo = 0 Then Return "一致✓"
        Return $"A獨有{ao}/此法獨有{bo}✗"
    End Function
    Private Function SpikeWalk_Oom(rootF As Folder, rootPath As String) As (paths As HashSet(Of String), ms As Long)
        Dim paths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim sw = Stopwatch.StartNew()
        Dim q As New Queue(Of (f As Folder, p As String))() : q.Enqueue((rootF, rootPath))
        While q.Count > 0
            Dim cur = q.Dequeue()
            Dim subs As Folders = Nothing
            Try
                subs = cur.f.Folders
                For Each sf As Folder In subs
                    Dim nm As String = "" : Try : nm = sf.Name : Catch : Continue For : End Try
                    Dim cp As String = cur.p & "\" & nm : paths.Add(cp) : q.Enqueue((sf, cp))
                Next
            Catch : End Try
            If subs IsNot Nothing Then TryMarshalRelease(subs)
        End While
        sw.Stop() : Return (paths, sw.ElapsedMilliseconds)
    End Function
    Private Function SpikeWalk_RdoEnum(rdoRoot As Redemption.RDOFolder, rootPath As String) As (paths As HashSet(Of String), ms As Long)
        Dim paths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim toRel As New List(Of Object)()
        Dim sw = Stopwatch.StartNew()
        Dim q As New Queue(Of (f As Redemption.RDOFolder, p As String))() : q.Enqueue((rdoRoot, rootPath))
        Try
            While q.Count > 0
                Dim cur = q.Dequeue()
                Dim subs = cur.f.Folders
                Try
                    For Each sf As Redemption.RDOFolder In subs
                        Dim nm As String = "" : Try : nm = sf.Name : Catch : Continue For : End Try
                        Dim cp As String = cur.p & "\" & nm : paths.Add(cp) : q.Enqueue((sf, cp)) : toRel.Add(sf)
                    Next
                Catch : End Try
                TryMarshalRelease(subs)
            End While
        Finally
            For Each o In toRel : Dim oo As Object = o : TryMarshalRelease(oo) : Next
        End Try
        sw.Stop() : Return (paths, sw.ElapsedMilliseconds)
    End Function
    Private Function SpikeWalk_RdoBatch(store As Redemption.RDOStore, rdoRoot As Redemption.RDOFolder, rootPath As String,
                                        withCC As Boolean, ByRef ccValid As Integer, ByRef ccErr As Integer, ByRef eidKind As String) As (paths As HashSet(Of String), ms As Long)
        Const DASL_SUB As String = "http://schemas.microsoft.com/mapi/proptag/0x360A000B"  ' PR_SUBFOLDERS
        Const DASL_CC As String = "http://schemas.microsoft.com/mapi/proptag/0x36020003"   ' PR_CONTENT_COUNT
        Dim cols As String = If(withCC, $"Name, EntryID, {DASL_SUB}, {DASL_CC}", $"Name, EntryID, {DASL_SUB}")
        Dim paths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim toRel As New List(Of Object)()
        Dim sw = Stopwatch.StartNew()
        Dim q As New Queue(Of (f As Redemption.RDOFolder, p As String))() : q.Enqueue((rdoRoot, rootPath))
        Try
            While q.Count > 0
                Dim cur = q.Dequeue()
                Try
                    Dim foldersCol = cur.f.Folders  ' 推斷型別,避免猜 RDOFolders 名稱
                    Dim tbl = foldersCol.MAPITable  ' 推斷型別,避免猜 MAPITable 名稱
                    Dim rc As Integer = CInt(tbl.RowCount)
                    If rc > 0 Then
                        tbl.Columns = cols : tbl.GoToFirst()
                        Dim rowsArr As Array = DirectCast(tbl.GetRows(rc), Array)
                        For i As Integer = rowsArr.GetLowerBound(0) To rowsArr.GetUpperBound(0)
                            Dim row As Array = DirectCast(rowsArr.GetValue(i), Array)
                            Dim lb As Integer = row.GetLowerBound(0)
                            Dim vName = row.GetValue(lb) : Dim vEid = row.GetValue(lb + 1) : Dim vSub = row.GetValue(lb + 2)
                            If eidKind = "?" AndAlso vEid IsNot Nothing Then eidKind = vEid.GetType().Name
                            Dim nm As String = If(TypeOf vName Is String, CStr(vName), "")
                            Dim cp As String = cur.p & "\" & nm : paths.Add(cp)
                            If withCC Then
                                Dim vCc = row.GetValue(lb + 3)
                                If TypeOf vCc Is Integer Then ccValid += 1 Else ccErr += 1
                            End If
                            Dim hasSub As Boolean = If(TypeOf vSub Is Boolean, CBool(vSub), True)  ' 未知→保守遞迴
                            If hasSub Then
                                Dim eidHex As String = SpikeEidToHex(vEid)
                                If eidHex <> "" Then
                                    Dim child As Redemption.RDOFolder = TryCast(store.GetFolderFromID(eidHex), Redemption.RDOFolder)
                                    If child IsNot Nothing Then q.Enqueue((child, cp)) : toRel.Add(child)
                                End If
                            End If
                        Next
                    End If
                    TryMarshalRelease(tbl) : TryMarshalRelease(foldersCol)
                Catch ex As System.Exception
                    Throw New System.Exception($"RdoBatch@{ExtractFolderName(cur.p)}: {ex.GetBaseException().Message}")  ' 探針: 明確報錯不靜默
                End Try
            End While
        Finally
            For Each o In toRel : Dim oo As Object = o : TryMarshalRelease(oo) : Next
        End Try
        sw.Stop() : Return (paths, sw.ElapsedMilliseconds)
    End Function
    Private Function SpikeEidToHex(v As Object) As String
        If v Is Nothing Then Return ""
        If TypeOf v Is String Then Return CStr(v)
        If TypeOf v Is Byte() Then Return BitConverter.ToString(DirectCast(v, Byte())).Replace("-", "")
        If TypeOf v Is Array Then
            Dim a As Array = DirectCast(v, Array)
            Dim sb As New System.Text.StringBuilder(a.Length * 2)
            For k As Integer = a.GetLowerBound(0) To a.GetUpperBound(0) : sb.Append(Convert.ToByte(a.GetValue(k)).ToString("X2")) : Next
            Return sb.ToString()
        End If
        Return ""
    End Function
#End Region
#End Region

End Class
