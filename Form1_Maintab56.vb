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

    Private _lv5FuzzyScoreMap As Dictionary(Of String, Double) = Nothing    ' 2026/06/17 by Simon/Claude Opus 4.8: Tab5 Fuzzy EntryID→對群代表的 body bigram Jaccard，供 RenderLv5Group 及排序/刪除重渲染查表顯示(不重讀 body)
    ' 原有: Private Const MIN_BIGRAM_FOR_FUZZY As Integer = 5   ' 內文 bigram 少於此值(極短/空白信)不納入模糊比對，避免無意義的雜訊群
    ' Q1-A 2026/06/18 by Simon/Claude Opus 4.8:
    '   MIN_BIGRAM_FOR_FUZZY          : 由 5 提高至 16 (保守起點, 實測再調)。內文 distinct bigram 少於此值(極短/純符號牆 >>>/空白信) 整封不納入模糊比對。屬「逐封自身長度」下限，與 Q1-C 互補
    '   MIN_SHARED_BIGRAM_FOR_FUZZY   : S5 最終閘門「共有內容量」下限 |A∩B|>=此值。Jaccard 比例對規模無感 (短信剛好全中→100% 但實質空洞)，另加交集絕對數把關。屬「逐對共有量」下限，與 Q1-A 互補
    ' 改後:
    '   Q1 連動滑桿 2026/06/18 by Simon/Claude Opus 4.8:
    '       原 MIN_SHARED = MIN_BIGRAM*2 固定值，改為隨檔位連動(見下方 MinSharedBigramFor)。MIN_BIGRAM_FOR_FUZZY 由 16→25 作 1 倍基準(低檔)。
    '       短信全中假陽性的成因是「共有量不足」，故連動打在共有量(C)上；S4 池子閘與 S5 最終閘共用 MinSharedBigramFor(targetT)，無死區(S4 只放進能過 S5 的信)。
    '       基準 25 為保守起點，看 _dbg("S5閘門") yield 再微調。
    ' 2026/6/19 by Simon/Claude Opus 4.8
    ' MIN_BIGRAM_FOR_FUZZY 的閾值控制: 把過關結果裡 inter 最小的那幾群打開看——
    '   如果 [25,50) 那段大多是「請查收附件謝謝」這種客套話 → 40，甚至可往 50 推
    '   如果那段藏著不同來源的真同文（只是短）→ 退回 30。
    Private Const MIN_BIGRAM_FOR_FUZZY As Integer = 45
    ' 2026/06/17 by Simon/Claude Opus 4.8: Tab5 Fuzzy 相似度檔位表。TrackBar1.Value 1~5 → 低/中/高/極高/完全一致。
    '   targetT 同時驅動 size 視窗(1/T)、Hamming 一階(HammingThresholdFor)、S5 最終 Jaccard 門檻(s >= targetT)，一個旋鈕全管。
    Private Shared ReadOnly _fuzzyTierT As Double() = New Double() {0, 0.87, 0.92, 0.95, 0.98, 1}   ' index 0 佔位(trackbar 從 1 起)
    Private Shared ReadOnly _fuzzyTierName As String() = New String() {"", "低", "中", "高", "極高", "完全一致"}
    Private ReadOnly _dbWriteLock As New Object()               ' 平行版多 worker 共用 _dbMail 連線寫入時的互斥鎖 (SqliteConnection 非執行緒安全，不可多執行緒同時 BeginTransaction)
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
        '   Layer2: ScanMailsToGroupDict / RenderLv5Group
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
                groupDict = Await ScanMailsToGroupDict(folderList, True, progress5, cToken)
                _lv5FuzzyScoreMap = Nothing ' Exact 不用 scoreMap
            Else
                ' 沿用 ScanMailsToGroupDict 完成資料夾列舉 + L2.5 快取，攤平成全體郵件(主旨分桶鍵在 Fuzzy 不再使用)
                Dim scanned = Await ScanMailsToGroupDict(folderList, False, progress5, cToken)
                Dim allMails = scanned.Values.SelectMany(Function(x) x).ToList()
                Dim targetT As Double = GetFuzzyTargetT()                                       ' S8 改接trackbar控制項參數 (低/中/高/極高)
                Dim thread_K As Integer = GetThreadCount()                                      ' 2026/07/07 by Simon/Claude: 讀 UI(numThread)一次，往下貫穿 S3/S4/S5 三處平行化
                Await PreComputeSimHash(allMails, progress5, cToken, thread_K)                  ' S3 build pass(暖快取跳過已算)
                Dim cand = Await PairFuzzyCandidates(allMails, targetT, thread_K, cToken)       ' S4 size 視窗 + Hamming 一階
                Dim filt = Await FilterCandidates(cand, targetT, progress5, thread_K, cToken)   ' S5 候選 body Jaccard 精算
                Dim fuzzy = BuildFuzzyGroups(filt.Pairs, filt.Sets)                             ' S6 union-find 分群 + scoreMap
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
            OkeyNowByeByeToken(cToken)   ' 2026/07/07 by Simon/Claude: 歸還 token — 運算中判定 token 化(見 OkayNowYouHaveToken/OkeyNowByeByeToken)
            Button5.Enabled = True : Cursor = Cursors.Default : _dbg("結束")
        End Try

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
    Private Async Function ScanMailsToGroupDict(folderList As List(Of (eid As String, sid As String, fPath As String)), isExact As Boolean, progress As IProgress(Of ProgressReport),   ' 2026/06/28 Stage2: 合約改 (eid,sid,fPath)
                                                     cToken As CancellationToken) As Task(Of Dictionary(Of String, List(Of MailItemInfo)))
        ' ---------------------------------------------------------------
        ' ScanMailsToGroupDict — 改用 GetMailInfo L2.5 快取（Tab4/Tab5 共用）
        ' 2026/05/06 by Claude: 原版直接 GetTable COM 掃描已移除，改走 L2.5 快取代理層
        '   ① 記憶體快取命中 → 0 COM call（Tab4 掃過即共享）
        '   ② SSD lazy load → 僅需 snapshot 驗證
        '   ③ COM fallback → GetMailInfoOOM，結果存入快取
        '   MsgIDhash / SenderEmail 已整合至 MailItemInfo，BuildMailGroupKey 直接使用
        ' ---------------------------------------------------------------
        Dim groupDict As New Dictionary(Of String, List(Of MailItemInfo))(StringComparer.OrdinalIgnoreCase)
        Dim totalFolders As Integer = folderList.Count
        Dim totalProcessed As Integer = 0
        Dim swThrottle As Stopwatch = Stopwatch.StartNew()  ' by Claude Sonnet 4.6, 2026/06/07
        Dim swTotal As Stopwatch = Stopwatch.StartNew()     ' 2026/05/10 by Simon/Claude: 供 ETA 計算使用; refactored by Claude Sonnet 4.6, 2026/06/07

        ' 2026/05/11 by Simon/Claude: SSD 批次預讀，將 DB 中的 mail_info 一次拉入記憶體
        Await PreLoadMailCacheAsync(folderList, cToken)

        For i As Integer = 0 To folderList.Count - 1
            ' 2026/06/29 by Simon/Claude [Stage2]: 改走 id-tuple，眼物化移除——folder 由免-folder 多載延後到 ③ 才建
            Dim eid As String = folderList(i).eid
            Dim sid As String = folderList(i).sid
            Dim fPath As String = folderList(i).fPath
            Try
                ' 透過 L2.5 取得（含快取），needTopic:=False (Tab5 不需要)
                Dim rows = Await GetMailInfo(fPath, eid, sid, needTopic:=False, cToken)
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

        ' 2026/07/04 by Simon/Claude: 筆數變化造成捲軸出現/消失不會觸發 Resize 事件，需主動重算欄寬一次 (同 RenderLv1 修法)
        CalculateLvColumnSize(ListView5)

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
        If Not ConfirmMailDelete(selCount) Then Return

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
            DeleteByMailList(entryIDs, affectedPaths, selCount)

            ' 2026/06/18 by Simon/Claude Opus 4.8: 原地更新 UI(取代整表 RenderLv5Group)，保留捲動位置與其餘列
            lv.BeginUpdate()
            For Each it In selectedItems : lv.Items.Remove(it) : Next          ' ① 移除被刪列，其餘列原地不動
            If lv.Items.Count > 0 Then                                         ' ② 游標留在原位(被刪列下一列)，不跳頂端
                Dim newIdx As Integer = Math.Min(anchorIndex, lv.Items.Count - 1)
                lv.Items(newIdx).Selected = True : lv.Items(newIdx).Focused = True : lv.Items(newIdx).EnsureVisible()
            End If
            lv.Invalidate()                                                    ' ③ 強制重繪→孤兒信經 DrawSubItem 上紅字(OwnerDraw 資料未變不會自動重畫)
            lv.EndUpdate()
            CalculateLvColumnSize(lv)                                          ' 2026/07/04 by Simon/Claude: 刪除後筆數變化可能造成捲軸出現/消失，同 RenderLv1 修法
        End If
        _dbg("結束")
    End Sub
#End Region
#Region "  ├ Fuzzy 模糊比對專用區塊 (SimHash + bigram Jaccard)"
    Private Async Function PreComputeSimHash(mails As List(Of MailItemInfo), progress As IProgress(Of ProgressReport), cToken As CancellationToken, thread_K As Integer) As Task
        ' 對「_cacheSimHash 沒有的」郵件讀 body 算 simhash+bigram_count，寫回獨立 db 與記憶體快取。已算過的(暖快取)直接跳過。
        ' 2026/07/03 by Simon/Claude [平行化SimHash]: PROBE_BODYPAR 實測(跨349資料夾/205616封母體樣本) numThread=8 平行讀 body 效率75%，
        '   35萬封估4.3分鐘(對比單執行緒約25分鐘)。_rdo2 在且 Outlook session 就緒時走平行版；否則(未勾CheckRDO)退回序列版。
        ' 2026/07/07 by Simon/Claude: numThread 改由呼叫端傳入(來源 numThread UI)，取代原本硬編碼的 SIMHASH_PARALLEL_K。
        LoadDbMail()
        Dim todo = mails.Where(Function(m) Not _cacheSimHash.ContainsKey(m.EntryID)).ToList()
        If todo.Count = 0 Then Return

        If _rdo2 IsNot Nothing AndAlso _olNS IsNot Nothing Then
            Await PreComputeSimHashParallel(todo, progress, thread_K, cToken)
        Else
            Await PreComputeSimHashSerial(todo, progress, cToken)
        End If
    End Function
    Private Async Function PreComputeSimHashSerial(todo As List(Of MailItemInfo), progress As IProgress(Of ProgressReport), cToken As CancellationToken) As Task
        ' 序列版(原實作)：未勾 CheckRDO 或 Outlook session 未就緒時的 fallback。走 L2.5 GetMailBody(內建 RDO/OOM 分派)。
        Dim totalBodyChars As Long = 0
        ' 2026/06/25 by Gemini 3.1 Pro: 將 Batch Size 提升至 3000，大幅降低磁碟寫入次數與 I/O 停頓
        Dim batch As New List(Of (EntryID As String, SimHash As Long, BigramCount As Integer))(3072)
        Dim swEta As Stopwatch = Stopwatch.StartNew()       ' 2026/06/17 by Simon/Claude: 供進度速度與 ETA 計算
        Dim swThrottle As Stopwatch = Stopwatch.StartNew()  ' 2026/06/25 by Gemini 3.1 Pro: 用於雙重節流的時間閘門
        For i As Integer = 0 To todo.Count - 1
            cToken.ThrowIfCancellationRequested()
            Dim id As String = todo(i).EntryID
            Dim body As String = GetMailBody(id, todo(i).FolderPath, skipCache:=True)
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
        _dbg("[SimHash]", $"序列版讀 {todo.Count} 封, body 累計 {totalBodyChars:N0} 字元 ≈ {totalBodyChars * 2 / 1048576:F0} MB(純UTF-16)")

        If batch.Count > 0 Then SaveDbMail(batch)
    End Function
    Private Async Function PreComputeSimHashParallel(todo As List(Of MailItemInfo), progress As IProgress(Of ProgressReport), thread_K As Integer, cToken As CancellationToken) As Task
        ' 平行版：每個 worker 在自己的 ThreadPool 執行緒內自建/自用/自 Logoff 一個獨立 RDOSession(不碰共用 _rdo2/UI 緒)，
        '   對齊 PROBE_BODYPAR 探針驗證過、確實能拿到平行加速的用法(memory_20260622_1846 §八: OOM 物件才必須留 UI 緒, RDO 獨立 session 可背景跑)。
        '   RDO 解析失敗的少數信(探針觀察約0.3%)收進回傳清單，等全部 worker 結束、續回本協調函數(繼承呼叫端的 UI 緒 SynchronizationContext)後，
        '   才用既有 GetMailBody(內建OOM fallback)逐封補算——OOM COM 物件只能在 UI 緒操作，背景 worker 絕不可呼叫 OOM。
        Dim profileName As String = _olNS.CurrentProfileName
        Dim numThread As Integer = Math.Min(thread_K, Math.Max(1, todo.Count))   ' 2026/07/07 by Simon/Claude: numThread 改用 numThread(來自 UI)，取代 SIMHASH_PARALLEL_K
        Dim chunks = SplitIntoChunks(todo, numThread)

        Dim swEta As Stopwatch = Stopwatch.StartNew()
        Dim processedCounter(0) As Integer     ' 單元素陣列：worker 們用 Interlocked 共用累加同一儲存格
        Dim lastReportMs(0) As Long
        Dim reportGate As New Object()
        Dim totalCount As Integer = todo.Count

        Dim tasks As New List(Of Task(Of List(Of MailItemInfo)))()
        For w As Integer = 0 To numThread - 1
            Dim idx As Integer = w
            tasks.Add(Task.Run(Function() SimHashParallelWorker(profileName, chunks(idx), processedCounter, totalCount, swEta, lastReportMs, reportGate, progress, numThread, cToken)))
        Next
        Dim rdoFailedLists = Await Task.WhenAll(tasks)
        Dim rdoFailed = rdoFailedLists.SelectMany(Function(x) x).ToList()

        _dbg("[SimHash]", $"平行化(K={numThread}) 完成 {todo.Count - rdoFailed.Count} 封, RDO失敗待OOM補算 {rdoFailed.Count} 封")

        ' RDO 解析失敗的少數信，回到 UI 緒用既有 GetMailBody(內建OOM fallback) 逐封補算
        If rdoFailed.Count > 0 Then
            Dim fallbackBatch As New List(Of (EntryID As String, SimHash As Long, BigramCount As Integer))(rdoFailed.Count)
            For Each m In rdoFailed
                cToken.ThrowIfCancellationRequested()
                Dim body As String = GetMailBody(m.EntryID, m.FolderPath, skipCache:=True)
                Dim setB = BuildBigramSet(body)
                Dim sh = ComputeSimHashFromSet(setB)
                _cacheSimHash(m.EntryID) = (sh, setB.Count)
                fallbackBatch.Add((m.EntryID, sh, setB.Count))
            Next
            SaveDbMail(fallbackBatch)
        End If
    End Function
    Private Function SimHashParallelWorker(profileName As String, subset As List(Of MailItemInfo), processedCounter As Integer(), totalCount As Integer, swEta As Stopwatch,
                                           lastReportMs As Long(), reportGate As Object, progress As IProgress(Of ProgressReport), thread_K As Integer, cToken As CancellationToken) As List(Of MailItemInfo)
        ' 平行版 worker：自建獨立 session 掃 subset 算 simhash，寫入 _cacheSimHash(ConcurrentDictionary,執行緒安全)與 _dbMail(自帶鎖)。
        '   RDO 解析失敗(store找不到/GetMessageFromID失敗/.Body失敗)的信收進回傳清單，交還協調函數做OOM補算(worker背景緒不可碰OOM)。
        Dim rdoFailed As New List(Of MailItemInfo)()
        Dim session As Redemption.RDOSession = Nothing
        Dim storeByName As New Dictionary(Of String, Redemption.RDOStore)()
        Try
            Try
                session = New Redemption.RDOSession()
                session.Logon(ProfileName:=profileName, Password:="", ShowDialog:=False, NewSession:=True)
            Catch ex As System.Exception
                _dbg("SimHashParallelWorker Logon失敗", ex.Message)
                Return subset   ' 這個 worker 的整批信全部交還協調函數用 OOM 補算
            End Try

            storeByName = BuildRdoStoreByNameDict(session)

            Dim batch As New List(Of (EntryID As String, SimHash As Long, BigramCount As Integer))(1024)
            For i As Integer = 0 To subset.Count - 1
                If (i And 127) = 0 Then cToken.ThrowIfCancellationRequested()
                Dim m As MailItemInfo = subset(i)
                Dim store As Redemption.RDOStore = Nothing
                Dim body As String = Nothing
                If storeByName.TryGetValue(GetStoreNameFromPath(m.FolderPath), store) Then
                    Dim rm As Redemption.RDOMail = Nothing
                    Try
                        Try : rm = TryCast(store.GetMessageFromID(m.EntryID), Redemption.RDOMail) : Catch : End Try
                        If rm IsNot Nothing Then Try : body = rm.Body : Catch : End Try
                    Finally
                        Dim o As Object = rm : TryMarshalRelease(o)
                    End Try
                End If

                If body Is Nothing Then
                    rdoFailed.Add(m)
                Else
                    ' 2026/07/03 by Simon/Claude: 指紋必須用正規化 body 算 — 序列版(GetMailBodyRdo)與 OOM 補算路徑都先過 NormalizeMailBody,
                    '   worker 若拿生 rm.Body 算,同一封信兩路徑指紋不一致,且會永久污染 OLAsimHash.db(Regex 為 Shared 預編譯實例,跨緒安全)
                    Dim setB = BuildBigramSet(NormalizeMailBody(body))
                    Dim sh = ComputeSimHashFromSet(setB)
                    _cacheSimHash(m.EntryID) = (sh, setB.Count)
                    batch.Add((m.EntryID, sh, setB.Count))
                    If batch.Count >= 1000 Then
                        SyncLock _dbWriteLock : SaveDbMail(batch) : End SyncLock
                        batch.Clear()
                    End If
                End If

                Dim done As Integer = Interlocked.Increment(processedCounter(0))
                If (done And 63) = 0 Then
                    Dim nowMs As Long = swEta.ElapsedMilliseconds
                    If nowMs - Interlocked.Read(lastReportMs(0)) >= ThrottleFreq.Mid AndAlso Monitor.TryEnter(reportGate) Then
                        Try
                            If nowMs - lastReportMs(0) >= ThrottleFreq.Mid Then   ' 進鎖後再驗一次，避免重複回報(雙重檢查)
                                lastReportMs(0) = nowMs
                                Dim eta = CalculateSpeedAndETA(totalCount, done, swEta.Elapsed.TotalSeconds)
                                progress?.Report(New ProgressReport With {.Message = $"計算內文指紋(平行化(K={thread_K}): {done}/{totalCount} ({eta.Speed:F0} 個/秒{eta.EtaString})"})
                            End If
                        Finally
                            Monitor.Exit(reportGate)
                        End Try
                    End If
                End If
            Next
            If batch.Count > 0 Then SyncLock _dbWriteLock : SaveDbMail(batch) : End SyncLock
            Return rdoFailed
        Finally
            For Each kv In storeByName : Dim o As Object = kv.Value : TryMarshalRelease(o) : Next
            If session IsNot Nothing Then
                Try : session.Logoff() : Catch : End Try
                Dim so As Object = session : TryMarshalRelease(so)
            End If
        End Try
    End Function
    Private Function MailBodyParallelWorker(profileName As String, subset As List(Of MailItemInfo), processedCounter As Integer(), totalCount As Integer, dbHitCount As Integer, swEta As Stopwatch,
                                            lastReportMs As Long(), reportGate As Object, progress As IProgress(Of ProgressReport), thread_K As Integer, cToken As CancellationToken) _
                                            As (Sets As Dictionary(Of String, HashSet(Of Integer)), RdoFailed As List(Of MailItemInfo))
        ' 2026/07/08 by Simon/Claude: S5-1b 平行版 worker — 對齊 SimHashParallelWorker(S3)的手法：
        '   自建獨立 RDOSession(不碰共用 _rdo2/UI 緒)，逐封讀 body → 正規化 → 建 bigram 集合(不算 simhash，S5 只需要集合)。
        '   RDO 解析失敗的少數信收進回傳清單，交還協調函數用既有 GetMailBody(內建OOM fallback)逐封補算——OOM COM 物件只能在 UI 緒操作。
        Dim localSets As New Dictionary(Of String, HashSet(Of Integer))(subset.Count)
        Dim rdoFailed As New List(Of MailItemInfo)()
        Dim session As Redemption.RDOSession = Nothing
        Dim storeByName As New Dictionary(Of String, Redemption.RDOStore)()
        Try
            Try
                session = New Redemption.RDOSession()
                session.Logon(ProfileName:=profileName, Password:="", ShowDialog:=False, NewSession:=True)
            Catch ex As System.Exception
                _dbg("FuzzyBodyParallelWorker Logon失敗", ex.Message)
                Return (localSets, subset)   ' 這個 worker 的整批信全部交還協調函數用 OOM 補算
            End Try

            storeByName = BuildRdoStoreByNameDict(session)

            For i As Integer = 0 To subset.Count - 1
                If (i And 127) = 0 Then cToken.ThrowIfCancellationRequested()
                Dim m As MailItemInfo = subset(i)
                Dim store As Redemption.RDOStore = Nothing
                Dim body As String = Nothing
                If storeByName.TryGetValue(GetStoreNameFromPath(m.FolderPath), store) Then
                    Dim rm As Redemption.RDOMail = Nothing
                    Try
                        Try : rm = TryCast(store.GetMessageFromID(m.EntryID), Redemption.RDOMail) : Catch : End Try
                        If rm IsNot Nothing Then Try : body = rm.Body : Catch : End Try
                    Finally
                        Dim o As Object = rm : TryMarshalRelease(o)
                    End Try
                End If

                If body Is Nothing Then
                    rdoFailed.Add(m)
                Else
                    ' 同 S3 SimHashParallelWorker: 集合必須用正規化 body 建 — 序列版(GetMailBodyRdo)與 OOM 補算路徑都先過 NormalizeMailBody,
                    '   worker 若拿生 rm.Body 建,同一封信兩路徑集合不一致,且會永久污染 bigram_set BLOB(Regex 為 Shared 預編譯實例,跨緒安全)
                    localSets(m.EntryID) = BuildBigramSet(NormalizeMailBody(body))
                End If

                Dim done As Integer = Interlocked.Increment(processedCounter(0))
                If (done And 63) = 0 Then
                    Dim nowMs As Long = swEta.ElapsedMilliseconds
                    If nowMs - Interlocked.Read(lastReportMs(0)) >= ThrottleFreq.Mid AndAlso Monitor.TryEnter(reportGate) Then
                        Try
                            If nowMs - lastReportMs(0) >= ThrottleFreq.Mid Then   ' 進鎖後再驗一次，避免重複回報(雙重檢查)
                                lastReportMs(0) = nowMs
                                Dim eta = CalculateSpeedAndETA(totalCount, done, swEta.Elapsed.TotalSeconds)
                                progress?.Report(New ProgressReport With {.Message = $"開始過濾候選內文(平行K={thread_K}): {done}/{totalCount} (DB命中 {dbHitCount:N0}) ({eta.Speed:F0} 個/秒{eta.EtaString})"})
                            End If
                        Finally
                            Monitor.Exit(reportGate)
                        End Try
                    End If
                End If
            Next
            Return (localSets, rdoFailed)
        Finally
            For Each kv In storeByName : Dim o As Object = kv.Value : TryMarshalRelease(o) : Next
            If session IsNot Nothing Then
                Try : session.Logoff() : Catch : End Try
                Dim so As Object = session : TryMarshalRelease(so)
            End If
        End Try
    End Function
    Private Async Function FilterCandidates(candidates As List(Of (A As MailItemInfo, B As MailItemInfo)), targetT As Double, progress As IProgress(Of ProgressReport), thread_K As Integer, cToken As CancellationToken) _
                                            As Task(Of (Pairs As List(Of (A As MailItemInfo, B As MailItemInfo, Score As Double)), Sets As Dictionary(Of String, HashSet(Of Integer))))

        ' 只對通過 S4 的少數候選讀 body、建 bigram 集合、算精確 bigram Jaccard。集合回傳給 S6 算群代表分數，免重讀。
        Dim sets As New Dictionary(Of String, HashSet(Of Integer))()

        ' 2026/07/05 by Simon/Claude: 去重改保留整個 MailItemInfo(原本只留 EntryID) — GetMailBody 不帶 folderPath 會落 OOM 慢路徑,帶上才走 RDO 快路徑
        Dim uniqueMails = candidates.SelectMany(Function(p) New MailItemInfo() {p.A, p.B}).GroupBy(Function(m) m.EntryID).Select(Function(g) g.First()).ToList()

        ' Phase1a: 先從 mail_simhash.bigram_set 整批載回已存的候選集合(核心版 BLOB 快取, 純 SQLite+CPU → 背景執行緒)
        ' 2026/07/07 by Simon/Claude: 只有進過 S5 的候選才有 BLOB(見 SaveDbMailSets), 是「最可能相似配對」的核心族群。
        '   實測全庫平均每封 set 僅 ~2.4KB, 數萬候選反序列化約 1~2 秒, 取代原本 20~50 秒的 RDO 重讀(22k/20s, 53k/50s)。
        progress?.Report(New ProgressReport With {.Message = $"載入候選指紋集合: {uniqueMails.Count:N0} 封 (DB)..."})
        Dim idList = uniqueMails.Select(Function(m) m.EntryID).ToList()
        Dim dbSets = Await Task.Run(Function()
                                        SyncLock _dbWriteLock
                                            Return LoadDbMailSets(idList)
                                        End SyncLock
                                    End Function, cToken)
        For Each kv In dbSets : sets(kv.Key) = kv.Value : Next

        ' Phase1b: DB 沒有的才讀 body → 建集合。
        ' 2026/07/08 by Simon/Claude: 原本序列版單線 GetMailBody 派工實測約 500 封/秒(大量首見候選時要等好幾分鐘)。
        '   比照 S3 PreComputeSimHashParallel 的多 RDOSession worker 手法接上 numThread：_rdo2 在且 Outlook
        '   session 就緒時走平行版；否則(未勾 CheckRDO)退回序列版 GetMailBody(內建 RDO/OOM 分派，會快取這少數幾封)。
        Dim missing = uniqueMails.Where(Function(m) Not sets.ContainsKey(m.EntryID)).ToList()
        Dim newRows As New List(Of (EntryID As String, SetBytes As Byte()))(missing.Count)

        If missing.Count > 0 AndAlso _rdo2 IsNot Nothing AndAlso _olNS IsNot Nothing Then
            Dim profileName As String = _olNS.CurrentProfileName
            Dim wK As Integer = Math.Min(thread_K, Math.Max(1, missing.Count))
            Dim chunks = SplitIntoChunks(missing, wK)
            Dim swEta As Stopwatch = Stopwatch.StartNew()
            Dim processedCounter(0) As Integer
            Dim lastReportMs(0) As Long
            Dim reportGate As New Object()
            Dim dbHitCount As Integer = dbSets.Count

            Dim tasks As New List(Of Task(Of (Sets As Dictionary(Of String, HashSet(Of Integer)), RdoFailed As List(Of MailItemInfo))))()
            For w As Integer = 0 To wK - 1
                Dim idx As Integer = w
                tasks.Add(Task.Run(Function() MailBodyParallelWorker(profileName, chunks(idx), processedCounter, missing.Count, dbHitCount, swEta, lastReportMs, reportGate, progress, wK, cToken)))
            Next
            Dim results = Await Task.WhenAll(tasks)

            Dim rdoFailed As New List(Of MailItemInfo)()
            For Each r In results
                For Each kv In r.Sets
                    sets(kv.Key) = kv.Value
                    newRows.Add((kv.Key, BigramSetToBytes(kv.Value)))
                Next
                rdoFailed.AddRange(r.RdoFailed)
            Next
            _dbg("[Fuzzy]", $"S5-1b 平行(K={wK}) 完成 {missing.Count - rdoFailed.Count} 封, RDO失敗待OOM補算 {rdoFailed.Count} 封")

            ' RDO 解析失敗的少數信，回到 UI 緒用既有 GetMailBody(內建OOM fallback) 逐封補算——OOM COM 物件只能在 UI 緒操作
            If rdoFailed.Count > 0 Then
                For Each m In rdoFailed
                    cToken.ThrowIfCancellationRequested()
                    Dim setB = BuildBigramSet(GetMailBody(m.EntryID, m.FolderPath))
                    sets(m.EntryID) = setB
                    newRows.Add((m.EntryID, BigramSetToBytes(setB)))
                Next
            End If
        ElseIf missing.Count > 0 Then
            ' 序列版 fallback(原實作)：未勾 CheckRDO 或 Outlook session 未就緒時使用。走 L2.5 GetMailBody(會快取這少數幾封)
            Dim swEta As Stopwatch = Stopwatch.StartNew()  ' 2026/06/17 by Simon/Claude: 供進度速度與 ETA 計算
            For k As Integer = 0 To missing.Count - 1
                cToken.ThrowIfCancellationRequested()
                Dim m = missing(k)
                Dim setB = BuildBigramSet(GetMailBody(m.EntryID, m.FolderPath))
                sets(m.EntryID) = setB
                newRows.Add((m.EntryID, BigramSetToBytes(setB)))   ' 2026/07/07: 收集待回寫 BLOB
                If (k And 15) = 0 Then
                    ' 2026/06/17 by Simon/Claude: 加入速度與 ETA 顯示，對齊 Tab3/Tab4 做法
                    Dim eta = CalculateSpeedAndETA(missing.Count, k + 1, swEta.Elapsed.TotalSeconds)
                    progress?.Report(New ProgressReport With {.Message = $"開始過濾候選內文: {k + 1}/{missing.Count} (DB命中 {dbSets.Count:N0}) ({eta.Speed:F0} 個/秒{eta.EtaString})"})
                    Await Task.Delay(1, cToken)
                End If
            Next
        End If

        ' Phase1c: 新算的集合批次回寫 DB, 下次冷搜尋這批候選直接走 Phase1a 免讀 body
        If newRows.Count > 0 Then
            progress?.Report(New ProgressReport With {.Message = $"回寫候選指紋集合: {newRows.Count:N0} 封 (DB)..."})
            Await Task.Run(Sub()
                               SyncLock _dbWriteLock
                                   SaveDbMailSets(newRows)
                               End SyncLock
                           End Sub, cToken)
        End If

        ' Phase2: Jaccard 精算(純 CPU)。
        ' 2026/07/07 by Simon/Claude: 「日後量大可改 Parallel.For」的日後到了 — 實測 95.8 萬對單線 2.56s，改平行。
        '   sets 字典此階段唯讀(Dictionary 並行讀取安全)；worker 收 (Idx, 結果) 合併後依 Idx 排序，輸出順序與單線版一致。
        '   MaxDegreeOfParallelism 改接 numThread(UI)，與 S3/S4 共用同一顆旋鈕。
        Dim minShared As Integer = MinSharedBigramFor(targetT)   ' Q1 連動 2026/06/18 by Simon/Claude: S5 共有量下限改用檔位值
        Dim pairs = Await Task.Run(
            Function() As List(Of (A As MailItemInfo, B As MailItemInfo, Score As Double))
                Dim collected As New List(Of (Idx As Integer, A As MailItemInfo, B As MailItemInfo, Score As Double))()
                Dim mergeLock As New Object()
                Dim po As New ParallelOptions With {.CancellationToken = cToken, .MaxDegreeOfParallelism = thread_K}
                Parallel.For(0, candidates.Count, po,
                    Function() New List(Of (Idx As Integer, A As MailItemInfo, B As MailItemInfo, Score As Double))(),
                    Function(k As Integer, state As ParallelLoopState, local As List(Of (Idx As Integer, A As MailItemInfo, B As MailItemInfo, Score As Double)))
                        Dim p = candidates(k)
                        ' Q1-C 2026/06/18 by Simon/Claude Opus 4.8: 先取交集絕對數，不足門檻直接淘汰(擋短信比例 100% 假陽性)；達標再由 inter 導出 Jaccard，免重算交集
                        Dim setA = sets(p.A.EntryID), setB = sets(p.B.EntryID)
                        Dim inter As Integer = BigramIntersectionCount(setA, setB)
                        If inter >= minShared Then
                            Dim union As Integer = setA.Count + setB.Count - inter
                            Dim s As Double = If(union = 0, 0.0, inter / union)   ' size 1/T 界線 S4 已保證，此處不重複早退
                            If s >= targetT Then local.Add((k, p.A, p.B, s))
                        End If
                        Return local
                    End Function,
                    Sub(local)
                        SyncLock mergeLock : collected.AddRange(local) : End SyncLock
                    End Sub)

                collected.Sort(Function(x, y) x.Idx.CompareTo(y.Idx))
                Return collected.Select(Function(x) (x.A, x.B, x.Score)).ToList()
            End Function, cToken)
        Return (pairs, sets)
    End Function
    Private Async Function PairFuzzyCandidates(mails As List(Of MailItemInfo), targetT As Double, thread_K As Integer, cToken As CancellationToken) As Task(Of List(Of (A As MailItemInfo, B As MailItemInfo)))
        ' size 1/T 滑動視窗收斂 O(n²) + Hamming 一階篩。純 CPU、無 COM → 放 Task.Run 不凍 UI。
        Dim hThr As Integer = HammingThresholdFor(targetT)
        Dim maxRatio As Double = 1.0 / targetT
        Dim minBigram As Integer = MinSharedBigramFor(targetT)   ' Q1 連動 2026/06/18 by Simon/Claude: S4 池子閘改用檔位值(同 S5，避免放進注定死在 S5 的信)

        Return Await Task.Run(
            Function() As List(Of (A As MailItemInfo, B As MailItemInfo))
                ' 取出有指紋且 bigram 數達標者，依 bigram_count 升冪排序(作為 size 視窗的排序鍵)
                Dim items = mails.Where(Function(m) _cacheSimHash.ContainsKey(m.EntryID) AndAlso _cacheSimHash(m.EntryID).BigramCount >= minBigram).
                                  Select(Function(m) (Mail:=m, SH:=_cacheSimHash(m.EntryID).SimHash, Cnt:=_cacheSimHash(m.EntryID).BigramCount)).
                                  OrderBy(Function(x) x.Cnt).ToList()

                ' 2026/07/07 by Simon/Claude: 外圈 Parallel.For 平行化 — 實測 178k 指紋/95.8萬對時單線 6.8~7.1s，是暖跑最大頭。
                '   純 CPU 零共享：worker 各自收 (i,j) 索引對(值型別)，合併後依 (i,j) 排序再物化，
                '   輸出順序與舊單線版完全一致(下游 union-find 群編號維持決定性)。
                '   取消改由 ParallelOptions.CancellationToken 負責(取代舊的 (i And 1023) 手動檢查)。
                '   MaxDegreeOfParallelism 改接 numThread(UI)，與 S3/S5-2 共用同一顆旋鈕。
                Dim pairsIdx As New List(Of (I As Integer, J As Integer))()
                Dim mergeLock As New Object()
                Dim po As New ParallelOptions With {.CancellationToken = cToken, .MaxDegreeOfParallelism = thread_K}
                Parallel.For(0, items.Count, po,
                    Function() New List(Of (I As Integer, J As Integer))(),
                    Function(i As Integer, state As ParallelLoopState, local As List(Of (I As Integer, J As Integer)))
                        For j As Integer = i + 1 To items.Count - 1
                            If items(j).Cnt > items(i).Cnt * maxRatio Then Exit For   ' size 1/T 上界：升冪→超界後 j 全部出局，收尾視窗
                            If GetHammingDistance(items(i).SH, items(j).SH) <= hThr Then local.Add((i, j))
                        Next
                        Return local
                    End Function,
                    Sub(local)
                        SyncLock mergeLock : pairsIdx.AddRange(local) : End SyncLock
                    End Sub)

                pairsIdx.Sort(Function(a, b) If(a.I <> b.I, a.I.CompareTo(b.I), a.J.CompareTo(b.J)))
                Dim result As New List(Of (A As MailItemInfo, B As MailItemInfo))(pairsIdx.Count)
                For Each p In pairsIdx : result.Add((items(p.I).Mail, items(p.J).Mail)) : Next
                Return result
            End Function, cToken)
    End Function
    Private Function BuildFuzzyGroups(similar As List(Of (A As MailItemInfo, B As MailItemInfo, Score As Double)), sets As Dictionary(Of String, HashSet(Of Integer))) As (GroupDict As Dictionary(Of String, List(Of MailItemInfo)), ScoreMap As Dictionary(Of String, Double))
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
        ' 2026/06/17 by Simon/Claude Opus 4.8: D6 Hamming 一階門檻「起始值」表
        '   ── 微調原因：SimHash Hamming 對應的是特徵向量夾角(cosine)、非 Jaccard，且 64-bit 量化有噪音，無法用公式定死。
        '   ── 微調方式：寧鬆勿緊(誤選 OK, 後面還有 Jaccard 把關 >> 但漏掉真重複就不 OK 了)。
        '                   下方 _dbg 探針會記錄「Hamming 過關配對數 vs Jaccard 過關數」，實際上機看 yield rate：太低就收緊、疑似漏抓就放寬。v1.1 依實測定案。Jaccard(S5) 才是準確閘門。
        Return _fuzzyTierT(Math.Clamp(TrackBar1.Value, 1, 5))   ' TrackBar1.Value(1~5)→targetT，越界夾住保險
    End Function
    Private Function HammingThresholdFor(targetT As Double) As Integer
        ' Q1 連動滑桿 2026/06/18 by Simon/Claude Opus 4.8: 共有內容量下限的檔位連動表(對齊 HammingThresholdFor/GetFuzzyTargetT 邊界)
        '   越嚴(高 T)→要求兩封共有越多真實內容才算重複。倍率 低1/中2/高3/極高4/完全一致5，乘上基準 MIN_BIGRAM_FOR_FUZZY(=25)。
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
    Private Function MinSharedBigramFor(targetT As Double) As Integer
        If targetT >= 0.99 Then Return MIN_BIGRAM_FOR_FUZZY * 5   ' 完全一致 125
        If targetT >= 0.98 Then Return MIN_BIGRAM_FOR_FUZZY * 4   ' 極高 100
        If targetT >= 0.95 Then Return MIN_BIGRAM_FOR_FUZZY * 3   ' 高 75
        If targetT >= 0.92 Then Return MIN_BIGRAM_FOR_FUZZY * 2   ' 中 50
        Return MIN_BIGRAM_FOR_FUZZY                               ' 低 0.87 (1×) 25
    End Function

    Private Function GetThreadCount() As Integer
        ' 2026/07/07 by Simon/Claude: Fuzzy 管線三處平行化(S3 SimHash worker 數 / S4 Hamming Parallel.For / S5-2 Jaccard Parallel.For)
        '   統一改讀 numThread(UI, layoutPanel5)，取代原本各自硬編碼的 SIMHASH_PARALLEL_K=8。
        '   在 Bt5_Click 開頭讀一次(UI 執行緒)、往下當參數傳，不在背景執行緒碰 UI 控制項。
        ' numThread.Value: 0 或未輸入 → 自動(Environment.ProcessorCount)；否則採使用者指定值(下限 1)。
        Dim value As Integer = CInt(numThreads.Value)
        Return If(value <= 0, Environment.ProcessorCount, value)
    End Function
    Private Function SplitIntoChunks(Of T)(list As List(Of T), k As Integer) As List(Of List(Of T))
        ' 把 list 切成 numThread 塊(連續切片, 最後一塊可能較短)。供平行版 SimHash 與 PROBE_BODYPAR 探針共用
        Dim result As New List(Of List(Of T))(k)
        Dim n As Integer = list.Count
        Dim per As Integer = CInt(Math.Ceiling(n / CDbl(k)))
        For i As Integer = 0 To k - 1
            Dim startIdx As Integer = i * per
            If startIdx >= n Then result.Add(New List(Of T)()) : Continue For
            result.Add(list.GetRange(startIdx, Math.Min(per, n - startIdx)))
        Next
        Return result
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
    Private Function ConfirmMailDelete(selCount As Integer) As Boolean
        ' 2026/07/05 by Simon/Claude: 從 HandleLv3/4/5Delete 抽出的共用確認對話框
        Return MessageBox.Show($"確定要將選中的 {selCount} 封郵件移到「刪除郵件」資料夾嗎？", "確認刪除",
                               MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes
    End Function
    Private Sub DeleteByMailList(entryIDs As List(Of String), affectedPaths As HashSet(Of String), selCount As Integer)
        ' 2026/07/05 by Simon/Claude: 從 HandleLv3/4/5Delete 抽出的共用刪除中段 — 清快取/DB、實體刪除、狀態列訊息
        For Each fPath In affectedPaths
            InvalidateMailCache(fPath)     ' 刪除後手動清理快取資料，避免殘留已刪除郵件的資訊
            ' 2026/07/06 by Simon/Claude Fable 5: 原 DbDeleteMailInfoByPath 只清 mail_info 一張表 —
            '   att_maillist 的 snap 比對是「att 列 snap vs folder_info.mc」兩邊同一次 SaveCache 寫入、同源必相等，
            '   year/month_counts 更是完全無驗證，三張表都會把已刪郵件從 DB lazy 原樣復活。改走統一失效入口一次清乾淨。
            DbPurgeFolderMailRows(fPath)
        Next
        ' 2026/07/06 by Simon/Claude Fable 5: 昂貴表 surgical — 刪除當下手上就有確切 entryID 清單，
        '   順手清 mail_simhash/att_filenames 對應列(含記憶體)，否則重複郵件/附件搜尋殘留幽靈要等下次 RenewCache 才消
        SimDbDeleteMailRowsByEntryIds(entryIDs, includeAttFilenames:=True)
        MoveMailsToRecycle(entryIDs)       ' 實體刪除 (移動到同 Store 的刪除郵件資料夾)
        PgrsBar2.Text = $"已移動 {selCount} 封郵件至刪除郵件資料夾"
    End Sub
    Private Sub InitLv3Lv4Lv5ContextMenu()
        ' 2026/06/15 by Simon/Claude Opus 4.8: Lv3/4/5 共用右鍵選單；冪等，重複呼叫只建一次 (確保三個 LV 共用同一實例)
        If ctxMenuLv3Lv4Lv5 IsNot Nothing Then Return Else ctxMenuLv3Lv4Lv5 = New ContextMenuStrip()

        Dim mnuOpen As New ToolStripMenuItem("開啟選取項目(&O)")
        Dim mnuPreview As New ToolStripMenuItem("快速預覽選取項目(&P)")
        Dim mnuRefresh As New ToolStripMenuItem("重刷選取項目(&R)")
        Dim mnuDelete As New ToolStripMenuItem("刪除選取項目(&D)")
        ctxMenuLv3Lv4Lv5.Items.Add(mnuOpen)
        ctxMenuLv3Lv4Lv5.Items.Add(mnuPreview)
        ctxMenuLv3Lv4Lv5.Items.Add(New ToolStripSeparator())
        ctxMenuLv3Lv4Lv5.Items.Add(mnuRefresh)
        ctxMenuLv3Lv4Lv5.Items.Add(mnuDelete)

        ' 2026/07/07 by Simon/Claude: 「開啟」= OpenSelectedMailsWithPreviewOffer(真正 Outlook Inspector，超過10封會問要不要改快速預覽)；「快速預覽」= ShowMailQuickPreview(HTMLBody+WebView2，ms 等級，無數量限制)
        AddHandler mnuOpen.Click, Sub(sender, e)
                                      Dim lv = TryCast(ctxMenuLv3Lv4Lv5.SourceControl, ListView)
                                      If lv IsNot Nothing Then OpenSelectedMailsWithPreviewOffer(lv)
                                  End Sub
        AddHandler mnuPreview.Click, Sub(sender, e)
                                         Dim lv = TryCast(ctxMenuLv3Lv4Lv5.SourceControl, ListView)
                                         If lv IsNot Nothing Then ShowMailQuickPreview(GetSelectedMailInfos(lv))
                                     End Sub

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
        Using bgBrush As New SolidBrush(backColor) : e.Graphics.FillRectangle(bgBrush, e.Bounds) : End Using

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

        ' 單擊左鍵複製主旨到剪貼簿，這原本是 Listview4 獨有的方便設計，現在擴展到 Tab3 共用 (by Gemini 3.1 Pro, 2026/04/21)
        'If item IsNot Nothing AndAlso e.Button = MouseButtons.Left Then  Clipboard.SetText(item.SubItems(0).Text)
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
            ' 2026/07/07 by Simon/Claude: 一般 Enter = 開啟(超過10封會問要不要改快速預覽)；Ctrl/Shift+Enter = 快速預覽(HTMLBody+WebView2，ms 等級，無數量限制)
            If e.Control OrElse e.Shift Then
                ShowMailQuickPreview(GetSelectedMailInfos(lv))
            Else
                OpenSelectedMailsWithPreviewOffer(lv)
            End If
            e.Handled = True : e.SuppressKeyPress = True

        ElseIf e.KeyCode = Keys.Delete Then
            ' 2026/07/05 by Simon/Claude Opus 4.8: 整合自 Lv3/Lv4/Lv5 原本各自獨立的 Delete 按鍵處理 (Lv3 原本沒有此快捷鍵，一併補上)
            If lv Is ListView3 Then HandleLv3Delete(lv)
            If lv Is Listview4 Then HandleLv4Delete(lv)
            If lv Is ListView5 Then HandleLv5Delete(lv)
            e.Handled = True

        ElseIf e.KeyCode = Keys.Escape Then
            If lv.VirtualMode Then lv.SelectedIndices.Clear() Else lv.SelectedItems.Clear()
            ' 對應不同的 TreeView 給予控制權
            If lv Is ListView3 Then SimTree3.Focus()
            If lv Is Listview4 Then Lv4Topic.Focus() : _lv4SimCToken?.Cancel()   ' 2026/5/29 by Simon/Claude: 拆分SimTree4的雙重模式後，將 Tab4 ESC 焦點從 SimTree4 調整到 Lv4Topic ' 2026/07/03 by Claude Sonnet 5: SelectedItems.Clear()觸發的SelectedIndexChanged在Count=0時會提前return不會自己取消，這裡補上明確取消
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
        End If
    End Sub

    ' Form1_Refresh.vb  —  郵件實體資訊強制刷新 (Lv3/Lv4/Lv5)
    ' ==============================================================
    ' 功能：
    '   (A) Lv3/Lv5 按 F5 → 強制刷新「目前顯示清單」內所有郵件的實體資訊 (Subject/Size/RcvTime/SenderName)
    '   (B) Lv3/Lv4/Lv5 右鍵選單「強制刷新選取的郵件」→ 只刷選取項
    '
    ' 設計重點：
    '   1. 跳過所有 cache (_cacheXXX / SSD)，直接打 COM 讀真值；讀完寫回顯示清單，並 patch 記憶體 cache。
    '   2. A/B 路徑「依數量」自動選擇 (與哪個 LV 無關)：
    '        targetList.Count <  REFRESH_BATCH_THRESHOLD(42) → 方法A：逐封 RefreshMailInfo
    '        targetList.Count >= 門檻                        → 方法B：依資料夾 GetTable+GetArray 批次
    '   3. 只刷「已在記憶體中的項目」：顯示清單 + 已含該 EntryID 的 cache；
    '      尚未 lazy load 的欄位 (附件數/檔名) 不主動額外讀取。
    '   4. AttCount 不再由這個流程讀取(全體F5/右鍵刷新一律不讀)：2026/07/04 探針證實 MAPI 無訊息層級批次欄位、
    '      逐封枚舉又是 Tab3 篩選路徑用不到的死碼(Tab3 實際讀 GetAttFilename().Count)。
    '      唯一保留的「操作別」訊號是 flushAttFilenameCache：右鍵單筆刷新時，順便清掉該筆過期的附件檔名快取
    '      (_cacheAttFilename)，讓 Tab3 下次查詢時重讀真值；全體 F5 / 批次 B 路徑不做這件事(避免大量失效過度昂貴)。
    '   5. SSD 不在此逐封碎寫，交給正常存檔流程；snapshot 計數不動 (沒有增減郵件)。
    '   6. 失效郵件 (EntryID 找不到) 顯示清單一律「保留舊資料 + 記錄」不動；
    '      但會觸發 SelfHealDeadEntryId(該資料夾) 毒化 DB snapshot，逼下次 RenewCache 對該夾強制全量重讀 + entry_id 清理
    '      (典型成因：PST 壓縮換 entry_id，或 copy→改→放回→刪原始 等 RenewCache 雙訊號都巧合沒變的邊界案例)。2026/07/04
    ' 2026/06/14 by Simon/Claude Opus 4.8
    ' ==============================================================
    Private Async Function RefreshSelectedLvItems(lv As ListView) As Task
        ' 2026/06/14 by Simon/Claude Opus 4.8: 右鍵單筆/複數刷新 (Lv3/4/5) — 只刷選取項
        ' 2026/07/04 by Simon/Claude: flushAttFilenameCache:=True — 順便清掉附件檔名快取(與 AttCount 讀取脫鉤)
        _dbg("開始")
        Dim target = GetLviMailTarget(lv)
        If target.Count = 0 Then Return

        _isUserBusy = True : Cursor = Cursors.WaitCursor
        Try
            Dim stats = Await RefreshLviCore(target, flushAttFilenameCache:=True, ct:=CancellationToken.None)
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
        ' 2026/06/14 by Simon/Claude Opus 4.8: 全體 F5 (Lv3/Lv5) — 蒐集底層所有郵件成 targetList，呼叫核心 (flushAttFilenameCache:=False)
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
            Dim stats = Await RefreshLviCore(targetList, flushAttFilenameCache:=False, ct:=CancellationToken.None)
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
    Private Async Function RefreshLviCore(target As List(Of (lst As List(Of MailItemInfo), idx As Integer)), flushAttFilenameCache As Boolean, ct As CancellationToken) As Task(Of RefreshStats)
        ' 2026/06/14 by Simon/Claude Opus 4.8: 刷新核心分派器 — 給定一組 (清單,索引) targetList，依數量選 A/B 重讀 COM 並寫回 + patch 記憶體 cache
        '   count <  門檻 → 方法A：逐封 RefreshMailInfo
        '   count >= 門檻 → 方法B：依資料夾 GetTable+GetArray 批次
        ' 2026/07/04 by Simon/Claude: 移除 readAttachCount(全體F5/右鍵刷新統一不讀附件數)；flushAttFilenameCache 純粹控制要不要順便清附件檔名快取，跟 AttCount 讀取無關。
        Dim stats As New RefreshStats
        If target Is Nothing OrElse target.Count = 0 Then Return stats

        Dim swThrottle As Stopwatch = Stopwatch.StartNew()
        Dim total As Integer = target.Count
        ' 2026/07/04 by Simon/Claude Fable 5: 每個資料夾在這趟刷新中只自癒一次 (SelfHealDeadEntryId 有 DB UPDATE，避免同夾多個 NotFound 重複打)
        Dim healedPaths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        If total < REFRESH_BATCH_THRESHOLD Then
            ' ── 方法A：逐封開信 ──
            For i As Integer = 0 To total - 1
                ct.ThrowIfCancellationRequested()
                Dim s = target(i)
                Dim m As MailItemInfo = s.lst(s.idx)
                Select Case RefreshMailInfo(m)
                    Case RefreshResult.Updated : s.lst(s.idx) = m : UpdateMailCaches(m, flushAttFilenameCache) : stats.Updated += 1 : _refreshedList.Add(m.EntryID)  ' 2026/06/15 by Simon/Claude Opus 4.8: 標記為已刷新
                    Case RefreshResult.NotFound   ' 保留舊資料不動；2026/07/04: 觸發該資料夾自癒 (毒化 DB snapshot，逼下次 RenewCache 強制重讀)
                        stats.NotFound += 1
                        If healedPaths.Add(m.FolderPath) Then SelfHealDeadEntryId(m.FolderPath)
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
                Dim fieldDict As Dictionary(Of String, MailItemInfo) = Await GetMailInfoAsDict(grp.Key, ct)
                If fieldDict Is Nothing Then
                    ' 資料夾解析/掃描失敗 → 該組退回逐封 (確保不整批漏掉)
                    For Each s In grp
                        ct.ThrowIfCancellationRequested()
                        Dim m As MailItemInfo = s.lst(s.idx)
                        Select Case RefreshMailInfo(m)
                            Case RefreshResult.Updated : s.lst(s.idx) = m : UpdateMailCaches(m, flushAttFilenameCache) : stats.Updated += 1 : _refreshedList.Add(m.EntryID)  ' 2026/06/15 by Simon/Claude Opus 4.8: 標記為已刷新
                            Case RefreshResult.NotFound
                                stats.NotFound += 1
                                If healedPaths.Add(m.FolderPath) Then SelfHealDeadEntryId(m.FolderPath)   ' 2026/07/04 by Simon/Claude Fable 5
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
                        ' 只搬基本欄位
                        m.Subject = fresh.Subject : m.Size = fresh.Size : m.RcvTime = fresh.RcvTime : m.SenderName = fresh.SenderName
                        s.lst(s.idx) = m : UpdateMailCaches(m, flushAttFilenameCache:=False) : stats.Updated += 1 : _refreshedList.Add(m.EntryID)  ' 2026/06/15 by Simon/Claude: 標記為已刷新 (方法B一律不清附件檔名快取)
                    Else
                        stats.NotFound += 1   ' 該 EntryID 已不在資料夾 (移動/刪除)；2026/07/04: 整批用同一個 grp.Key，資料夾層級自癒一次即可
                        If healedPaths.Add(grp.Key) Then SelfHealDeadEntryId(grp.Key)
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
            Case 4 : Return If(order = SortOrder.Ascending, sourceList.OrderBy(Function(x)  ' 附件數 (Tab3 專用，依賴全域 _cacheAttFilename)
                                                                                   Dim files As List(Of String) = Nothing
                                                                                   Return If(_cacheAttFilename.TryGetValue(x.EntryID, files), files.Count, 0)
                                                                               End Function).ToList(),
                                                            sourceList.OrderByDescending(Function(x)
                                                                                             Dim files As List(Of String) = Nothing
                                                                                             Return If(_cacheAttFilename.TryGetValue(x.EntryID, files), files.Count, 0)
                                                                                         End Function).ToList())
            Case Else
                Return sourceList
        End Select
    End Function
    Private Sub UpdateMailCaches(mail As MailItemInfo, flushAttFilenameCache As Boolean)
        ' 2026/06/14 by Simon/Claude Opus 4.8: 只 patch「現有 in-memory cache 中已含此 EntryID」的項目；絕不新建 key、不觸發掃描/lazy load
        '   內層 List 是參考型別，原地改元素即可，dict 與 snapshot 不動
        ' 2026/07/03 by Simon/Claude Fable 5: 原地 patch 不會經過 _cacheMailInfo(fPath)=... / _cacheAttMailList(fPath)=... 這種
        '   「整格重新賦值」的 dirty 追蹤標記點，SaveCache 的 dirty 過濾會因此漏掉這裡的異動。必須在此手動補標記，
        '   否則 F5 逐封刷新改到的 Subject/Size/SenderName 等欄位永遠不會被寫回 SQLite。
        ' 2026/07/04 by Simon/Claude: 移除 AttCount 回填(探針證實 Tab3 篩選路徑不讀這個欄位，逐封枚舉是死碼)。
        '   flushAttFilenameCache 改為獨立參數，只負責③附件檔名快取失效，跟 AttCount 讀取脫鉤。
        MarkMailFolderDirty(mail.FolderPath)

        ' ① Tab3 附件清單快取
        Dim t3 As FolderCacheTab3 = Nothing
        If _cacheAttMailList.TryGetValue(mail.FolderPath, t3) AndAlso t3.AttMailList IsNot Nothing Then
            Dim lst = t3.AttMailList
            Dim j As Integer = lst.FindIndex(Function(x) x.EntryID = mail.EntryID)
            If j >= 0 Then
                Dim c = lst(j)
                c.Subject = mail.Subject : c.Size = mail.Size : c.RcvTime = mail.RcvTime : c.SenderName = mail.SenderName
                lst(j) = c
            End If
        End If

        ' ② Tab4 基本資訊快取 (Mails 是 List(Of (Mail, Topic)))
        Dim t4 As (Mails As List(Of (Mail As MailItemInfo, Topic As String)), Snap As Long) = Nothing
        If _cacheMailInfo.TryGetValue(mail.FolderPath, t4) AndAlso t4.Mails IsNot Nothing Then
            Dim lst = t4.Mails
            Dim j As Integer = lst.FindIndex(Function(x) x.Mail.EntryID = mail.EntryID)
            If j >= 0 Then
                Dim e = lst(j)
                e.Mail.Subject = mail.Subject : e.Mail.Size = mail.Size : e.Mail.RcvTime = mail.RcvTime : e.Mail.SenderName = mail.SenderName
                lst(j) = e
            End If
        End If

        ' ③ 附件檔名快取：只在右鍵單筆刷新(flushAttFilenameCache=True)時失效該筆，避免日後拿到過期檔名 (不主動重讀)
        If flushAttFilenameCache Then
            Dim dummy As List(Of String) = Nothing
            _cacheAttFilename.TryRemove(mail.EntryID, dummy)
        End If
    End Sub
    Private Sub InvalidateFolderTreeCache(fPath As String)
        ' ---------------------------------------------------------------
        ' InvalidateFolderTreeCache — 宣告指定路徑的記憶體快取失效 (Layer 2.5)
        ' 目的: 隱藏快取鍵值的命名細節 (如 "|True", "|False")，避免 UI 層過度耦合
        ' 2026/5/31 by Simon/Gemini 3.1 Pro
        ' ---------------------------------------------------------------
        If String.IsNullOrEmpty(fPath) Then Return

        Dim isInSub = Function(k As String) k = fPath OrElse k.StartsWith(fPath & "\") OrElse k.StartsWith(fPath & "|")
        Dim dummyFolderList As List(Of Folder) = Nothing
        Dim dL As Long

        ' 1. 清除該資料夾與其所有子資料夾的快取
        For Each key In _cacheFolderTree.Keys.Where(isInSub).ToList() : _cacheFolderTree.TryRemove(key, dummyFolderList) : Next
        For Each key In _cacheMailCountAll.Keys.Where(isInSub).ToList() : _cacheMailCountAll.TryRemove(key, dL) : Next
        For Each key In _cacheFolderCountAll.Keys.Where(isInSub).ToList() : _cacheFolderCountAll.TryRemove(key, dL) : Next
        For Each key In _cacheMailCount.Keys.Where(isInSub).ToList() : _cacheMailCount.TryRemove(key, dL) : Next
        For Each key In _cacheFolderCount.Keys.Where(isInSub).ToList() : _cacheFolderCount.TryRemove(key, dL) : Next

        ' 【修復關鍵 2】補上身分證字典的清理, 2026/6/1 by Simon/Gemini 3.1 Pro (2026/07/03: _cacheIsMailFolder 已併入 _cacheFolderIDs，移除獨立清理)
        Dim dummyId As (eid As String, sid As String, isMail As Boolean, hasCh As Boolean) = Nothing
        For Each key In _cacheFolderIDs.Keys.Where(isInSub).ToList() : _cacheFolderIDs.TryRemove(key, dummyId) : Next

        ' 2. 清除祖先節點的快取 (因為子節點異動，祖先的「包含子目錄」加總也會跟著變)
        For Each ancestor In GetAncestors(fPath)
            For Each sfx In {"", "|True", "|False"}
                _cacheMailCountAll.TryRemove(ancestor & sfx, dL)
                _cacheFolderCountAll.TryRemove(ancestor & sfx, dL)
                _cacheFolderSizeAll.TryRemove(ancestor & sfx, dL)
            Next
        Next

        _dbg("快取清除", $"已清除 {fPath} 相關之記憶體快取")
    End Sub
    Private Sub InvalidateMailCache(fPath As String)
        ' ---------------------------------------------------------------
        ' InvalidateMailCache — 刪除郵件後，主動清除指定 fPath 的記憶體快取
        ' 只清 _cacheMailInfo 和 _cacheMailCount 兩個 key，不影響其他資料夾
        ' 配合 DbDeleteMailInfoByPath 一起呼叫，確保記憶體與 DB 兩層同步失效
        ' 2026/05/11 by Claude Sonnet 4.6
        ' 2026/05/12 by Simon/Claude: 擴充清除範圍至所有受影響的快取
        ' ---------------------------------------------------------------
        If String.IsNullOrEmpty(fPath) Then Return

        ' ── 層次一：該資料夾本身 ──────────────────────────────────────
        Dim dummy1 As (Mails As List(Of (Mail As MailItemInfo, Topic As String)), Snap As Long) = Nothing
        _cacheMailInfo.TryRemove(fPath, dummy1)

        Dim dummy2 As Long
        _cacheMailCount.TryRemove(fPath, dummy2)
        _cacheMailCountAll.TryRemove(fPath, dummy2)
        _cacheMailCountAll.TryRemove(fPath & "|True", dummy2)
        _cacheMailCountAll.TryRemove(fPath & "|False", dummy2)

        _cacheFolderSize.TryRemove(fPath, dummy2)
        _cacheFolderSizeAll.TryRemove(fPath, dummy2)
        _cacheFolderCommitMax.TryRemove(fPath, dummy2)   ' 2026/07/06 by Simon/Claude Fable 5: commit 基準一併失效，否則 SaveCache 的 COALESCE 會拿記憶體舊值把 PoisonFolderSnapDb 清掉的 commit_max 寫回

        _cacheYearCounts.TryRemove(fPath, Nothing)

        ' month_counts key 格式為 "fPath_YYYY"，不知道是哪年，清所有匹配的
        For Each mk In _cacheMonthCounts.Keys.Where(Function(k) k.StartsWith(fPath & "_")).ToList()
            _cacheMonthCounts.TryRemove(mk, Nothing)
        Next

        _cacheAttMailList.TryRemove(fPath, Nothing)

        ' ── 層次二：所有祖先路徑的聚合快取 ──────────────────────────
        For Each ancestor In GetAncestors(fPath)
            _cacheMailCountAll.TryRemove(ancestor, dummy2)
            _cacheMailCountAll.TryRemove(ancestor & "|True", dummy2)
            _cacheMailCountAll.TryRemove(ancestor & "|False", dummy2)
            _cacheFolderSizeAll.TryRemove(ancestor, dummy2)
        Next

        _dbg("結束", ExtractFolderName(fPath))
    End Sub
    Private Sub SelfHealDeadEntryId(fPath As String)
        ' ---------------------------------------------------------------
        ' SelfHealDeadEntryId — 快取 entry_id 對 GetItemFromID/RDO GetMessageFromID 解析失敗 (NotFound) 時呼叫
        ' 清記憶體 (InvalidateMailCache) + 毒化 DB snapshot (PoisonFolderSnapDb)，
        ' 逼下次 RenewCache 對此資料夾強制全量重讀 + surgical entry_id 清理，不必等數量/commit_max 剛好變動才觸發。
        ' 涵蓋 RenewCacheToDB 雙訊號 (count+commit_max) 都巧合沒變的邊界案例，典型成因：純 PST 壓縮換 entry_id。
        ' 不在此同步重讀 —— 呼叫端多半在逐封/批次迴圈中途，同步重掃整夾太貴；統計欄位交給下次 RenewCache 狀況A。
        ' 2026/07/04 by Simon/Claude Fable 5
        ' 2026/07/06 by Simon/Claude Fable 5: 光「清記憶體+毒化」擋不住 DB-lazy 立即復活死列 —— GetMailInfo ② 的
        '   比對是 mail_info.pr_count_snap vs GetMailCount(免-folder)=folder_info.mail_count，兩邊同一次 SaveCache
        '   寫入、同源必相等(壓縮不改 count)；且毒化寫的是 pr_count_snap 欄、免-folder ② 讀的是 mail_count 欄，
        '   整條 lazy 路徑根本讀不到毒。所以直接把整夾可疑 DB 列清掉：先撈 entryID 做兩張昂貴表 surgical
        '   (mail_info 一刪就再也查不到 folder→entryID 對應，順序不能反)，再 nuke 便宜表；
        '   下次 lazy 讀取 DB miss → RDO ms 級重建全新列，死 ID 立即絕跡，不必苦等 RenewCache。
        '   成本與 RenewCache 狀況A 的 surgical 步驟同級，且 healedPaths 保證每夾每趟只跑一次。
        ' ---------------------------------------------------------------
        If String.IsNullOrEmpty(fPath) Then Return
        InvalidateMailCache(fPath)
        PoisonFolderSnapDb(fPath)
        SimDbDeleteMailRowsByEntryIds(LazyGetFolderIdAsList(fPath), includeAttFilenames:=True)
        DbPurgeFolderMailRows(fPath)
    End Sub
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
        PgrsBar2.Text = $"DB 統計 — folder_info:{st.fc} 筆 / att_maillist:{st.mb} 筆 / att_filenames:{st.at} 筆 / year_counts:{st.yc} 筆 / month_counts:{st.mc} 筆 / {st.kb} KB"

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
        chkClearSimHash.CheckState = CheckState.Unchecked

        RefreshLv6DbStats()
        _dbg("結束")

    End Sub
    Private Async Sub RenewCache_Click(sender As Object, e As EventArgs) Handles RenewCache.Click
        ' 2026/04/09 重構: 原本只做孤兒清除，現在改呼叫完整的 RenewCacheToDB
        '   RenewCacheToDB 內含: Phase1 BFS → Phase2 snapshot 比對 → Phase3 dirty 重算
        '                        Phase4 ancestor 聚合清除 → Phase5 month_counts DB 清除
        '                        Phase6 CleanupOrphan + SaveCachesToDB
        '   RenewIncludeSize 勾選時才重算 folder_size (GetTable 遍歷，大資料夾較慢)
        ' 2026/6/7: by simon/Gemini: 直接在這裡計時顯示整體耗時, 去除原本在 RenewCacheToDB 內的多段計時, 避免重構後的邏輯分散導致耗時統計不完整或混亂

        Dim sw As Stopwatch = Stopwatch.StartNew()
        Dim renewSummary As String = ""     ' 2026/07/07 by Simon/Claude: 接住 RenewCacheToDB 既有的統計彙整字串
        Try
            renewSummary = Await RenewCacheToDB()
            Await DbVacuumIfNeeded()        ' 2026/06/16 by Claude Sonnet 4.6: RenewCache 完成後，視碎片比例決定是否執行 VACUUM (freelist_count / page_count > 5% 才執行，避免每次都白等)
            RefreshLv6DbStats()
            Await RefreshAllTreeViews()     ' by Gemini 3.0 flash, 2026/04/24: 更新完成後，執行非同步 UI 刷新，確保新資料夾能立即顯示

        Catch ex As OperationCanceledException
            _dbg(" ├ 中斷", "使用者已取消快取更新")
        Finally
            PgrsBar1.Text = $"RenewCache 完成 — 耗時: {sw.Elapsed.TotalSeconds:0.00} 秒"
            If renewSummary <> "" Then PgrsBar2.Text = renewSummary   ' 2026/07/07 by Simon/Claude: 秀出既有的新增/更新/刪除統計，純接線不加新邏輯
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
            AddLv6StatLine("_cacheFolderTree", $"{_cacheFolderTree.Count:N0} 筆")
            AddLv6StatLine("_cacheFolderIDs", $"{_cacheFolderIDs.Count:N0} 筆")
            AddLv6StatLine("_cacheSubTreeList", $"{_cacheSubTreeList.Count:N0} 筆")
            AddLv6StatLine("", "", isHeader:=False) ' 間隔
            AddLv6StatLine("_cacheMailCount", $"{_cacheMailCount.Count:N0} 筆")
            AddLv6StatLine("_cacheMailCountAll", $"{_cacheMailCountAll.Count:N0} 筆")
            AddLv6StatLine("_cacheFolderCount", $"{_cacheFolderCount.Count:N0} 筆")
            AddLv6StatLine("_cacheFolderCountAll", $"{_cacheFolderCountAll.Count:N0} 筆")
            AddLv6StatLine("_cacheYearCounts", $"{_cacheYearCounts.Count:N0} 筆")
            AddLv6StatLine("_cacheMonthCounts", $"{_cacheMonthCounts.Count:N0} 筆")
            AddLv6StatLine("_cacheAttMailList", $"{_cacheAttMailList.Count:N0} 筆")
            AddLv6StatLine("_cacheAttFilename", $"{_cacheAttFilename.Count:N0} 筆")
            AddLv6StatLine("_cacheFolderSize", $"{_cacheFolderSize.Count:N0} 筆")
            AddLv6StatLine("_cacheFolderSizeAll", $"{_cacheFolderSizeAll.Count:N0} 筆")
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
            AddLv6StatLine("folder_info", $"{st.fc:N0} 筆")
            AddLv6StatLine("senders", $"{st.senders:N0} 筆")     ' 2026/06/14 by Simon/Claude Opus 4.8: 補上 senders，與 DbShowDbFileStat 順序一致
            AddLv6StatLine("mail_info", $"{st.basic:N0} 筆")     ' by Gemini 3 Flash, 2026/04/22
            AddLv6StatLine("year_counts", $"{st.yc:N0} 筆")
            AddLv6StatLine("month_counts", $"{st.mc:N0} 筆")
            AddLv6StatLine("att_maillist", $"{st.mb:N0} 筆")
            AddLv6StatLine("─── OLAcacheMail.db ────", "", True) ' 2026/06/21 by Simon/Claude: att_filenames/mail_simhash 住此檔
            AddLv6StatLine("att_filenames", $"{st.at:N0} 筆")
            AddLv6StatLine("mail_simhash", $"{st.sh:N0} 筆")     ' 2026/06/21 by Simon/Claude: 新增
            AddLv6StatLine("bigram_set", $"{st.bs:N0} 筆")       ' 2026/07/07 by Simon/Claude: S5 候選集合 BLOB 回填量(非 NULL 筆數/淨容量)
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
            ElseIf selectedLabel.Contains("folder_info") Then   ' 2026/06/12 by Simon/Claude Opus 4.8: 補上缺漏的分支
                targetTableName = "folder_info"
            ElseIf selectedLabel.Contains("mail_info") Then
                targetTableName = "mail_info"
            ElseIf selectedLabel.Contains("att_maillist") Then
                targetTableName = "att_maillist"
            ElseIf selectedLabel.Contains("att_filenames") Then
                targetTableName = "att_filenames"
            ElseIf selectedLabel.Contains("year_counts") Then   ' 2026/06/12 by Simon/Claude Opus 4.8: 補上缺漏的分支
                targetTableName = "year_counts"
            ElseIf selectedLabel.Contains("month_counts") Then  ' 2026/06/12 by Simon/Claude Opus 4.8: 修正 typo (month_stats → month_counts)
                targetTableName = "month_counts"
            ElseIf selectedLabel.Contains("mail_simhash") Then  ' 2026/06/21 by Simon/Claude: 新增 mail_simhash 分支(DbShowTableStat 內部會路由到 _dbMail)
                targetTableName = "mail_simhash"
            ElseIf selectedLabel.Contains("bigram_set") Then    ' 2026/07/07 by Simon/Claude: bigram_set 是 mail_simhash 表內的欄位，筆數/大小跟全表不同，需獨立統計而非整表路由
                Dim unused2 = DbShowBigramSetStat()
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
    Private Sub DebugButton_Click(sender As Object, e As EventArgs) Handles DebugButton.Click
        ' 2026/07/07 by Simon/Claude: 快速預覽功能已提升為正式功能(見 ShowMailQuickPreview/GetSelectedMailInfos in Form1_MainTab34.vb，
        '   右鍵選單「快速預覽選取項目」與 Ctrl/Shift+Enter 皆可觸發)。這裡保留兩支純比較用診斷探針：
        '   平常點擊 = 手工 TextBox 陽春預覽(.Body 純文字，比較基準)；按住 Shift = 測試 RDO 獨立 session 自己的 Display()(已證實無效，見備忘)；
        '   按住 Ctrl = 呼叫正式版 ShowMailQuickPreview，方便跟前兩者同批次比較耗時


    End Sub

#End Region
#End Region

End Class
