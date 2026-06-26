Imports System.Buffers
Imports System.Threading
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Outlook

Partial Class Form1

#Region "■ 01 全域宣告"
    Private _lv3LastSortColumn As Integer = -1                  ' 儲存上一次點選的列索引
    Private _lv3SortOrder As SortOrder = SortOrder.Ascending    ' 設置初始排序方式為升序
    Private _lv4LastSortColumn As Integer = -1                  ' by Gemini 3 Flash, 2026/04/19: 加入 Listview4 專屬排序狀態 (避免與 LV3 共用變數互相干擾)
    Private _lv4SortOrder As SortOrder = SortOrder.Ascending    ' by Gemini 3 Flash, 2026/04/19: 加入 Listview4 專屬排序狀態 (避免與 LV3 共用變數互相干擾)
    Private _lv4TLastSortColumn As Integer = -1                 ' by Gemini/Simon, 2026/5/30: 紀錄 Lv4Topic 上次點擊的欄位
    Private _lv4TSortOrder As SortOrder = SortOrder.Ascending   ' by Gemini/Simon, 2026/5/30: 紀錄 Lv4Topic 目前是升冪或降冪

    Const REFRESH_BATCH_THRESHOLD As Integer = 42               ' 2026/06/14 by Simon/Claude Opus 4.8: <42 走A、>=42 走B (涵蓋 <41→A、>42→B，並補齊 41→A/42→B)
    Private ctxMenuLv3Lv4Lv5 As ContextMenuStrip = Nothing      ' 2026/06/14 by Simon/Claude Opus 4.8: Lv3/4/5 共用的右鍵刷新選單 (單一實例，初始化於 InitLv3Lv4Lv5ContextMenu)
    Private _refreshedList As New HashSet(Of String)(StringComparer.Ordinal)    ' 2026/06/15 by Simon/Claude Opus 4.8: 記錄本次（或累積數次）刷新成功的 EntryID，供 Lv3/4/5 以藍色字體標示；新搜尋開始時清除
    Private _lv5OrphanedList As New HashSet(Of String)(StringComparer.Ordinal)  ' 2026/06/18 by Simon/Claude Opus 4.8: Q4 Lv5 刪除後失去配對的「孤兒信」EntryID，供 DrawSubItem 標紅；RenderLv5Group 重渲染時清除

    Private layoutPanel3 As Panel
    Private _lv3MailList As New List(Of MailItemInfo)(4096)     ' by Gemini, 2026/04/10: Tab3 顯示資料庫 (虛擬模式核心)    ' 預分配容量為 4096，因應 Tab3 可能載入的大量郵件資訊，顯著降低記憶體配置開銷 (by Gemini 3 Flash, 2026/05/04)
    ' Private _isTab3_Stop As Boolean                           ' 2026/04/05 by Gemini: 已併入全域 _cancelRequested，不再單獨使用專屬旗標以簡化邏輯內容流程處理機制
    Private _lv4SimCToken As CancellationTokenSource = Nothing  ' by Gemini 3 Flash, 2026/04/26: 用於游標快速移動時, 取消前一次未完成的相似度計算任務
    ' _tab4FolderTreeNodesBackup / _tab4LastClickedFolderNode 已移除'   節點快照改由 SimTree4.SaveTreeNodeSnap("folder-view") 內部管理，不再需要 Form1 level 備份變數  ' 2026/05/23 by Simon/Claude
    ' Private _isTv4ResultMode As Boolean = False               ' ✅ 2026/04/20 by Gemini 2.0 Flash: 標記 Tab4 左側樹目前顯示的是搜尋結果模式 ' by Claude Sonnet 4.6, 2026/05/29: 已廢棄，改用 Lv4Topic.Visible 代替
    Private _tv4PrevSelection As New List(Of Folder)(32)        ' ✅ 2026/04/21 by Gemini 3.0 flash: 記憶最後一次搜尋的多個資料夾  ' 預分配容量為 32，足以涵蓋多數搜尋路徑結構，減少陣列頻繁 Resize 開銷 (by Gemini 3 Flash, 2026/05/04)
    Private _tv4SelectedTopicMailList As List(Of MailItemInfo) = Nothing    ' 2026/5/29 by Simon/Claude: 將SimTree4的雙重模式拆分，取代 SimTree4.SelectedNode.Tag 作為跨函數的資料橋樑, 供 F6 快速切換使用
    ' 2026/5/31 by Gemini/Simon: 徹底大掃除：清除所有 F6 與 Group 的殘留
    'Private _tv4GroupSortByCount As Boolean = True              ' ✅ 2026/04/20 by Gemini 2.0 Flash: 記錄排序方式 (True=數量, False=主旨)
    'Private _lv4GroupSortByCount As Boolean = False             ' by Gemini 3 Flash, 2026/04/20: 記錄 Tab4 Listview4 分組排序模式 (False:按主旨, True:按數量)
    'Private _tv4PrevTopicResults As Dictionary(Of String, List(Of MailItemInfo)) = Nothing  ' ✅ 2026/04/20 by Gemini 2.0 Flash: 記憶搜尋結果，供 F6 操作使用
    Private _lv4TopicList As New List(Of KeyValuePair(Of String, List(Of MailItemInfo)))(4096)  ' 2026/5/30 by Gemini, Lv4Topic 虛擬模式的資料來源
#End Region

#Region "■ 06 Tab3: 依附件條件搜尋"
    ' ===========================================================================================
    ' TabPage3 搜尋附件 — 架構重構演進
    ' ---------------------------------------------------
    ' [v2] by Claude, 2026/03/07
    '      策略: 建立雙階段搜尋 (Phase1 GetTable 快速掃描中繼資料, Phase2 GetItemFromID 讀取附件明細)。
    '      成效: 大幅減少對 MailItem 物件的依賴和操作，提升搜尋效率。
    '
    ' [v3] by Gemini, 2026/04/05 (現行架構)
    '      策略: 導入「管線化處理 (Pipeline)」與「SOLID 分層 (Layer1/Layer2.5/Layer3)」，徹底解耦 MAPI、業務與 UI。
    '      分層:
    '        ├─ Layer1   (UI/流程層) : Bt3_Click, ShowLv3Result
    '        ├─ Layer2   (商務過濾層): FilterBySize, FilterByAttachDetailsAsync
    '        ├─ Layer2.5 (快取層)    : GetAttachMailList (_cacheAttachMailList / _cacheAttachFilename)
    '        └─ Layer3   (MAPI操作層): GetAttachMailListL3
    '
    ' Bt3_Click 管線 (Pipeline) 步驟分解:
    '   Step 1. 前置驗證      → 檢查參數合法性 (UI)
    '   Step 2. BFS 遍歷      → GetSubtreeToList，取得目標資料夾樹 (COM)
    '   Step 3. 匯集資料全集  → 向 Layer2.5 索取候選郵件全集 (GetAttachMailList) → ① 記憶體 → ② DB (mail_basic) → ③ L3
    '   Step 4. 管線過濾 1    → 記憶體 LINQ 瞬間過濾大小限制 (FilterBySize)
    '   Step 5. 管線過濾 2    → 依據關鍵字與數量條件深層過濾，配合 Layer2.5 快取判定 (FilterByAttachDetailsAsync)
    '                         → PreloadAttachmentCacheRDOAsync (RDO 平行預載) → GetAttachFilename → ① 記憶體 → ② DB (mail_attachments) → ③ L3
    '   Step 6. UI 映射與顯示 → 將資料封裝為介面項目並顯示，無縫銜接">0"或真實統計 (ShowLv3Result)
    ' ===========================================================================================
#Region "  ├ Layer1 UI事件層"
    Private Async Sub Bt3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        _dbg("開始") ' by Gemini, 2026/04/10: UI 層級 Level 0
        Dim sw As Stopwatch = Stopwatch.StartNew()          ' by Claude Sonnet 4.6, 2026/06/07
        Dim swThrottle3 As Stopwatch = Stopwatch.StartNew() ' by Claude, 2026/04/11; refactored by Claude Sonnet 4.6, 2026/06/07

        ' by Gemini, 2026/04/09: 讀取 SimTree3.SelectedNodes 集合以支援多選 (取代原單一節點)
        Dim selectedNodes As List(Of TreeNode) = SimTree3.SelectedNodes
        If selectedNodes Is Nothing OrElse selectedNodes.Count = 0 Then Return

        ' ── 鎖定 UI ──
        PgrsBar1.Text = "準備中" : PgrsBar2.Text = "" : Cursor = Cursors.WaitCursor
        layoutPanel3.Enabled = False : SimTree3.Enabled = False
        ListView3.VirtualMode = True        ' by Gemini, 2026/04/10: 解決 ListView 萬筆資料 Clear() 造成 UI 卡頓 1.8 秒的效能瓶頸
        ListView3.VirtualListSize = 0       ' 切換至 VirtualMode 並清空 Size，不銷毀實體物件，速度為 0ms
        _lv3MailList.Clear() : _refreshedList.Clear() : Await Task.Yield ' 確保 UI 先更新狀態再進行後續耗時操作

        Dim cToken As CancellationToken = OkayNowYouHaveToken()  ' ✅ 取得新 Token，同時取消上一次未完成的操作
        Try
            ' ── Step 1: 驗證大小設定 (矛盾就提早返回，快取查詢在 Step3 做 LINQ 過濾) ──
            If CheckSize.Checked Then
                Dim minSize = CLng(NumberMin.Value * GetSizeMultiplier(UnitMin.SelectedItem.ToString))
                Dim maxSize = CLng(NumberMax.Value * GetSizeMultiplier(UnitMax.SelectedItem.ToString))
                If minSize > maxSize Then
                    _dbg("結束", "大小設定錯誤")
                    MessageBox.Show("大小設定錯誤: 最小值不能大於最大值。", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If
            End If

            ' ── Step 2: 收集目標資料夾清單 (支援多選防重複) ── ' high: 第一耗時, 52000筆資料花360ms
            Dim swStep As Stopwatch = Stopwatch.StartNew() ' by Gemini 3.0 flash, 2026/04/16: 暫時加入分段測速
            Dim progressTree = New Progress(Of ProgressReport)(Sub(p) PgrsBar2.Text = p.Message)
            Dim folderList = Await GetUniqueFolderList(selectedNodes, _includeSubTab3, cToken:=cToken, progress:=progressTree) ' 導入 HashSet(Of String) 來過濾跨父子節點重複選擇的資料夾
            Dim tStep2_UniqueList = swStep.Elapsed.TotalMilliseconds : swStep.Restart() ' by Gemini 3.0 flash, 2026/04/16: 改名以區分 (GetUniqueFolderList)

            If folderList.Count = 0 Then Return

            ' by Gemini, 2026/04/15: 提煉 FolderPath 本機陣列，大幅減少後續迴圈內的 COM 調用
            ' 2026/04/16 by Gemini: 這裡的 f.fPath 已經是 Tuple 屬性，完全無 COM 開銷
            Dim fPaths = folderList.Select(Function(f) f.fPath).ToList()

            ' by Gemini, 2026/04/09: 計算選定資料夾內的郵件總數 (用於進度顯示母數，參考 Tab2)
            Dim totalMailCount As Long = 0
            _dbg("開始step2迴圈")
            For i As Integer = 0 To folderList.Count - 1
                ' 2026/04/16 by Gemini: 指定使用 Tuple 內的 .Folder 與 .fPath
                Dim c As Integer = GetMailCount(folderList(i).folder, fPaths(i))    ' 從 400ms 降至近乎 0ms!
                If c > 0 Then totalMailCount += c
                Await SmartThrottle(swThrottle3, cToken:=cToken, ThrottleFreq.Hii) ' 2026/04/16 by Simon/Claude: 改用 ThrottleFreq.Hii + SmartThrottle
            Next
            _dbg("結束step2迴圈")
            Dim tStep2_MailCountLoop = swStep.Elapsed.TotalMilliseconds : swStep.Restart() ' by Gemini 3.0 flash, 2026/04/16: 改名以區分 (GetMailCount Loop)

            PgrsBar1.Text = "正在讀取..."
            PgrsBar2.Text = $"準備掃描 {folderList.Count:N0} 個資料夾 (相依包含共計 {totalMailCount:N0} 封信)..."
            Await Task.Yield()

            ' ── Step 3: 收集含附件的郵件清單 (透過 Layer2.5 快取) ──
            Dim progressPhase1 As IProgress(Of ProgressReport) = New Progress(Of ProgressReport)(Sub(p) PgrsBar2.Text = p.Message)
            Dim targetMails As New List(Of MailItemInfo)(4096)  ' 預分配容量為 4096，顯著降低掃描大量郵件時的記憶體配置開銷 (by Gemini 3 Flash, 2026/05/04)
            Try
                _dbg("開始step3迴圈")
                For i As Integer = 0 To folderList.Count - 1
                    Dim processed As Integer = i + 1
                    ' 2026/04/16 by Gemini: 使用 Tuple 中的 .Folder 與預錄好的 fPaths(i)
                    Dim folderResult = Await GetAttachMailList(folderList(i).folder, progressPhase1, fPaths(i), cToken:=cToken)
                    targetMails.AddRange(folderResult)
                    Await SmartThrottle(swThrottle3, cToken:=cToken, ThrottleFreq.Hii,
                                        Sub()
                                            Dim eta = CalculateSpeedAndETA(folderList.Count, processed, swStep.Elapsed.TotalSeconds)
                                            progressPhase1?.Report(New ProgressReport With {.CurrentCount = processed, .TotalCount = folderList.Count,
                                                                                            .Message = $"Phase 1 (載入資料夾清單): {processed} / {folderList.Count} 個資料夾 ({eta.Speed:F0} 個/秒{eta.EtaString})"})
                                        End Sub)
                Next
                _dbg("結束step3迴圈")
            Catch ex As OperationCanceledException
                ' by Gemini, 2026/04/12: 捕捉 ESC 中斷，結算目前已載入的部分郵件清單
                _dbg(" ├ 中斷", $"Step 3 已中斷，結算目前已載入的 {targetMails.Count:N0} 封")
                PgrsBar1.Text = "由使用者中斷"
            End Try
            Dim tStep3_AttachMailLoop = swStep.Elapsed.TotalMilliseconds : swStep.Restart() ' by Gemini 3.0 flash, 2026/04/16: 改名以區分 (GetAttachMailList Loop)

            ' ── Pipeline 過濾 1: 大小篩選 ──
            If CheckSize.Checked Then targetMails = FilterBySize(targetMails)   ' 這裡五萬筆資料只花<3ms

            ' ── Pipeline 過濾 2: 附件條件深層篩選 ──
            Dim hasKeyword = CheckAttachName.Checked AndAlso TextBox3.Text.Trim.Length > 0
            If hasKeyword OrElse CheckAttCount.Checked Then
                Dim progressPhase2 = New Progress(Of ProgressReport)(Sub(p) PgrsBar2.Text = p.Message)
                targetMails = Await FilterByAttachDetailsAsync(targetMails, progressPhase2, cToken:=cToken)
            End If
            Dim tStep5_DetailsFiltering = swStep.Elapsed.TotalMilliseconds : swStep.Stop() ' by Gemini 3.0 flash, 2026/04/16: 改名以區分 (Details Filtering)

            sw.Stop() ' ── 終極 Mapping 與顯示結果 ── ' by Gemini 3.0 flash, 2026/04/16: 將分段耗時拆分為多列顯示於 Debug ListView
            If _iLikeNoisy Then
                _dbg("⌛ 效能 (1/4) - GetUniqueFolderList", $"{tStep2_UniqueList:F0}ms")
                _dbg("⌛ 效能 (2/4) - GetMailCount", $"{tStep2_MailCountLoop:F0}ms")
                _dbg("⌛ 效能 (3/4) - GetAttachMailList", $"{tStep3_AttachMailLoop:F0}ms")
                _dbg("⌛ 效能 (4/4) - FilterByAttachDetailsAsync", $"{tStep5_DetailsFiltering:F0}ms")
                _dbg("⌛ 效能 (總計) - Total", $"{sw.Elapsed.TotalMilliseconds:F0}ms")
            End If

            ShowLv3Result(targetMails, sw.Elapsed.TotalSeconds)
        Catch ex As OperationCanceledException
            _dbg("結束", "ESC 中斷")
            PgrsBar1.Text = "由使用者中斷。" : PgrsBar2.Text = ""
        Catch ex As System.Exception
            MessageBox.Show("搜尋發生錯誤: " & ex.Message, "錯誤")
            _dbg("       ├ 錯誤", ex.Message) ' by Gemini, 2026/04/11: Level 3
        Finally
            ' ── 無論如何都解鎖 UI ──
            SimTree3.Enabled = True : Button3.Enabled = True
            layoutPanel3.Enabled = True : Cursor = Cursors.Default
            _dbg("結束") ' by Gemini, 2026/04/11: 修正對應開始層級 Level 0
        End Try

    End Sub
    Private Sub Lv3_RetrieveVirtualItem(sender As Object, e As RetrieveVirtualItemEventArgs) Handles ListView3.RetrieveVirtualItem
        ' --------------------------------------------------------------
        ' [封存] 下方為 2026/04/10 之前的實體排序邏輯，保留作為開發歷程參考
        ' 當時由 USER/Gemini 共同優化，利用 Tag 存取 Long 值來加速 O(1) 比較。
        ' 目前在 VirtualMode = True 時已不呼叫，但註解內部的 debug 改進歷程非常珍貴。
        ' Friend Class ListViewItemComparer ... (內容已完整保留在上方屬性區預定義中)
        ' --------------------------------------------------------------

        ' by Gemini, 2026/04/10: 虛擬模式核心事件 - 只有當某行「進入視野」時才會被觸發
        ' 此處利用記憶體中預載的 _lv3MailList 瞬間組裝 Item，完全避免掉傳統 AddRange 的巨量 Handle 配置時間
        If e.ItemIndex < 0 OrElse e.ItemIndex >= _lv3MailList.Count Then Return

        Dim mail = _lv3MailList(e.ItemIndex)
        Dim cachedFiles As List(Of String) = Nothing
        Dim displayName As String = ">0"
        If _cacheAttachFilename.TryGetValue(mail.EntryID, cachedFiles) Then displayName = cachedFiles.Count.ToString()

        ' 暫時建立一個 Item 回報給系統進行繪製
        ' by Gemini 3.0 Flash, 2026/04/20: 郵件大小改為 KB, 日期格式統一 yyyy/MM/dd (補零+置中需求)
        Dim lvi As New ListViewItem(mail.Subject)
        lvi.SubItems.Add((mail.Size \ 1024L).ToString("N0") & " KB")
        lvi.SubItems.Add(mail.RcvTime.ToString("yyyy/MM/dd HH:mm:ss"))
        lvi.SubItems.Add(mail.SenderName)
        lvi.SubItems.Add(displayName)
        lvi.SubItems.Add(mail.EntryID)
        ' 2026/06/15 by Simon/Claude Opus 4.8: Lv3 為虛擬模式，非懸停時系統 DrawDefault 直接讀 item 屬性繪製 (不走 DrawSubItem)，故刷新藍色須在此設定
        If _refreshedList.Contains(mail.EntryID) Then lvi.ForeColor = Color.Blue

        e.Item = lvi
    End Sub
    Private Sub Lv3_DrawItem(sender As Object, e As DrawListViewItemEventArgs) Handles ListView3.DrawItem
        ' ── ListView3 (Virtual Mode) OwnerDraw 實作 ──
        ' VirtualMode 必須比較 Index，不能比較 Item 參照 (因為是動態生成的)
        If _lastHoveredLvItem IsNot Nothing AndAlso e.ItemIndex = _lastHoveredLvItem.Index AndAlso Not e.Item.Selected Then
            ' 攔截繪製，由 DrawSubItem 處理灰底
        Else
            e.DrawDefault = True
        End If
    End Sub
    Private Sub Lv3_ColumnClick(sender As Object, e As ColumnClickEventArgs) Handles ListView3.ColumnClick
        ' ==============================================================
        ' by Gemini, 2026/04/10: 效能大躍進 — 虛擬模式排序
        ' 理由:
        '   當資料量達到 5 萬筆時，原有的 ListViewItemComparer 必須不斷從實體 Item 中取值，每秒只能比較約數千筆，會導致 UI 卡頓
        '   切換至 VirtualMode 後，我們直接在記憶體中對 _lv3MailList (List(Of MailItemInfo)) 進行 LINQ 排序，處理 5 萬筆數據僅需 10-30ms，達成瞬間排序效果
        ' --------------------------------------------------------------
        _dbg("開始", "虛擬列表排序") ' by Gemini, 2026/04/10: Level 0
        Dim sw As Stopwatch = Stopwatch.StartNew()  ' by Claude Sonnet 4.6, 2026/06/07

        ' 判斷是否點選的是同一個列標題, 如果是，則切換排序方式, 否則預設使用升序排序
        _lv3SortOrder = GetNewSortOrder(e.Column, _lv3LastSortColumn, _lv3SortOrder)    ' 2026/05/30 by Gemini/Simon: 抽取共用函式 GetNewSortOrder，簡化排序狀態切換邏輯
        _lv3LastSortColumn = e.Column  ' 儲存目前點選的列索引

        'ListView3.ListViewItemSorter = New ListViewItemComparer(e.Column, _lv3SortOrder)   ' by Gemini 2026/4/10, Listview3改virtual mode
        ListView3.BeginUpdate()
        Try
            ' 2. 直接對底層資料源進行排序，不操作 UI 物件
            _lv3MailList = SortMailList(_lv3MailList, e.Column, _lv3SortOrder)              ' 2026/5/31 by Gemini/Simon: 呼叫共用函式進行 LINQ 排序
            ListView3.Invalidate()  ' 💡 關鍵：Invalidate 會通知 ListView 重新按需索取資料，配合 EndUpdate 瞬間更新畫面
        Finally
            ListView3.EndUpdate()
        End Try

        sw.Stop()
        PgrsBar2.Text = $"虛擬排序 {_lv3MailList.Count:N0} 項，耗時 {sw.Elapsed.TotalSeconds:0.00} 秒"
        _dbg("結束", "排序列表") ' by Gemini, 2026/04/10

    End Sub
#End Region
#Region "  ├ Layer2 流程協調層"
    Private Function FilterBySize(sourceList As List(Of MailItemInfo)) As List(Of MailItemInfo)
        ' Pipeline 過濾 1: 大小篩選 (純記憶體 LINQ，速度極快)
        _dbg(" ├ 開始")
        Dim minSz = CLng(NumberMin.Value * GetSizeMultiplier(UnitMin.SelectedItem.ToString))
        Dim maxSz = CLng(NumberMax.Value * GetSizeMultiplier(UnitMax.SelectedItem.ToString))
        Dim resultList = sourceList.Where(Function(c) c.Size >= minSz AndAlso c.Size <= maxSz).ToList()
        _dbg(" ├ 結束", $"篩選後剩下 {resultList.Count:N0} 封")
        Return resultList
    End Function
    Private Async Function FilterByAttachDetailsAsync(sourceList As List(Of MailItemInfo), progress As IProgress(Of ProgressReport), cToken As CancellationToken) As Task(Of List(Of MailItemInfo))
        ' Pipeline 過濾 2: 逐一讀取附件明細，利用 _cacheAttachFilename 大幅降低 COM 存取
        _dbg(" ├ 開始", $"候選郵件: {sourceList.Count} 封")

        ' 2026/04/05 by Gemini: Layer2 業務層向 Layer2.5 請求平行預載快取。若 RDO 存在，便能在極短時間內把後續需要的資料全數載入記憶體快取供後續流程取用。
        ' 2026/06/19 by Simon/Claude: 新增RdoPreloadAttach_3, 每執行緒獨立 RDOSession，繞過 _olNS 共用 session 序列化，跨 PST 真平行
        ' 2025/06/23 by simon/Claude Opus 4.8: 正確導入Redemption獨立session加速附件檔名讀取, 直接高速讀取, 淘汰原有的PreLoad填充快取機制
        'If RDO_Parallel1.Checked Then     : Await RdoPreloadAttach_1(sourceList, progress, cToken:=cToken) ' by Parellel.ForEach 來平行讀取附件資料，適合 CPU 密集型的 MAPI 存取
        'ElseIf RDO_Parallel2.Checked Then : Await RdoPreloadAttach_2(sourceList, progress, cToken:=cToken) ' by Task.WhenAll 來平行讀取附件資料，適合 I/O 等待型的資料庫存取
        'ElseIf RDO_Parallel3.Checked Then : Await RdoPreloadAttach_3(sourceList, progress, cToken:=cToken) ' 2026/06/19 by Simon/Claude: 每執行緒獨立 RDOSession，繞過 _olNS 共用 session 序列化，跨 PST 真平行
        ' 假設您現在的流程是：「讀取大量郵件屬性 --> 與本地資料庫比對快取 --> 寫入資料庫」
        ' 1. 讀取 MAPI 資料 (CPU + 嚴格 Thread 限制)：適用 Parallel.ForEach。 → 因為您需要真實在多個核心上建立獨立的 RDOSession 來平行榨取硬碟與 MAPI 引擎的讀取速度
        ' 2. 查詢/寫入本地資料庫快取 (I/O 等待)：適用 Task.WhenAll + async。  → 如果您的底層資料庫驅動 (例如 SQLite-net 或 Entity Framework) 支援原生的非同步方法(ToListAsync, ExecuteAsync)
        ' 3. 跨多個不同PST檔讀取大量郵件資料：多PST獨立Session平行預載，獨立 RDOSession (自有 MAPI session) 才是 Redemption 真 free-threaded 的前提。
        '    設計為「每個 PST 一條獨立 session、組內循序、組間平行」，加速僅來自「跨多個 PST 同時讀取」 → 適用同時選取大量 PST 的整庫掃描，若只有少數 PST 時請改用 preLoad_1/_2 (共用熱 session 反而較快)。

        Dim swTotal As Stopwatch = Stopwatch.StartNew()     ' by Claude Sonnet 4.6, 2026/06/07
        Dim swThrottle As Stopwatch = Stopwatch.StartNew()  ' by Claude Sonnet 4.6, 2026/06/07

        Dim mustCountAttach As Boolean = CheckAttCount.Checked
        Dim minCount As Integer = If(mustCountAttach, CInt(CountMin.Value), 0)
        Dim maxCount As Integer = If(mustCountAttach, CInt(CountMax.Value), Integer.MaxValue)

        ' todo: 要在這裡過濾 "_IRM_Protected" 資料夾?? 2026/6/23 by simon
        'sourceList = sourceList.Where(Function(c) Not c.FolderPath.EndsWith("\_IRM_Protected", StringComparison.OrdinalIgnoreCase)).ToList()
        'sourceList = sourceList.Where(Function(c) Not c.FolderPath.Split("\"c).Last().Equals("_IRM_Protected", StringComparison.OrdinalIgnoreCase)).ToList()

        Dim processed As Integer = 0, total As Integer = sourceList.Count
        Dim resultList As New List(Of MailItemInfo)(1024)   ' 預分配容量為 1024，優化搜尋結果清單的填充速度 (by Gemini 3 Flash, 2026/05/04)
        Dim keyword As String = If(CheckAttachName.Checked, TextBox3.Text.Trim.ToLower(), "")
        Try
            For curMail As Integer = 0 To sourceList.Count - 1
                ' 2026/4/5, by Gemini: 將進度報告與 UI 釋放移至迴圈開頭，提早反饋處理進度
                ' 避免被下方的 Guard Clauses (Continue For) 略過而導致長時間霸佔主執行緒, 未更新UI進度反饋
                processed = curMail + 1
                ' 2026/04/16 by Gemini 3.0 flash: 改用 ThrottleFreq.Hii + SmartThrottle 與 onThrottled 委派
                Await SmartThrottle(swThrottle, cToken:=cToken, ThrottleFreq.Hii,
                                          Sub()
                                              Dim eta = CalculateSpeedAndETA(total, processed, swTotal.Elapsed.TotalSeconds)
                                              progress?.Report(New ProgressReport With {.CurrentCount = processed, .TotalCount = total,
                                                                                        .Message = $"Phase 2 (開始逐一開啟郵件讀取附件名稱): {processed} / {total}，已符合 {resultList.Count} 封 ({eta.Speed:F0} 封/秒{eta.EtaString})"})
                                          End Sub)

                Dim currentMail As MailItemInfo = sourceList(curMail)
                Dim cachedAttFilenames As List(Of String) = GetAttachFilename(currentMail, skipCache:=False)

                ' ── Guard Clause 0: 沒附件資料就不受理 ──
                If cachedAttFilenames Is Nothing Then Continue For

                ' ── Guard Clause 1: 數量過濾 ──
                If mustCountAttach AndAlso (cachedAttFilenames.Count < minCount OrElse cachedAttFilenames.Count > maxCount) Then Continue For

                ' ── Guard Clause 2: 檔名關鍵字過濾 (使用 LINQ Any 取代巢狀 For Each) ──
                If keyword.Length > 0 AndAlso Not cachedAttFilenames.Any(Function(fn) fn IsNot Nothing AndAlso
                                                                             fn.Contains(keyword, StringComparison.OrdinalIgnoreCase)) Then Continue For
                ' 通過所有安檢
                resultList.Add(currentMail)
            Next
        Catch ex As OperationCanceledException
            ' by Gemini, 2026/04/12: 捕捉 ESC 中斷，回傳已比對到的部分結果
            _dbg(" ├ 中斷", $"FilterByAttachDetailsAsync 已中斷，結算目前已符合的 {resultList.Count} 封")
        End Try
        _dbg(" ├ 結束", $"Phase 2 完成，篩選後共 {resultList.Count} 封")
        Return resultList

    End Function
    Private Sub ShowLv3Result(sourceList As List(Of MailItemInfo), elapsedSeconds As Double)
        _dbg("開始", sourceList.Count)
        ' by Gemini, 2026/04/10: 虛擬模式下僅需同步資料與設定 Size，完全不需建立物件
        _lv3MailList = sourceList
        ListView3.VirtualListSize = _lv3MailList.Count
        ' ListView3.Invalidate() ' 2026/05/09 by Gemini 3 Flash: 移除冗餘 Invalidate，設定 VirtualListSize 本身已會觸發重繪，且 Invalidate 可能引發 redundant Resize 事件。

        '' by Gemini 2026/4/10, Listview3改virtual mode, 下面的實體項目建立邏輯已完全移除，改由 RetrieveVirtualItem 事件按需生成
        '' ListView3.Items.Clear()
        '' ' 先告訴 ListView 總共會有幾筆，讓它一次配置好記憶體，不要每次 Add 都 realloc
        '' Dim lviCount As Integer = sourceList.Count
        '' If lviCount > 50 Then SendMessage(ListView3.Handle, LVM_SETITEMCOUNT, New IntPtr(lviCount), IntPtr.Zero)
        '' If lviCount > 10 Then ListView3.BeginUpdate()
        '' If lviCount > 0 Then
        ''     Dim items As New List(Of ListViewItem)(lviCount)
        ''     For Each m As MailItemInfo In sourceList
        ''         ' 如果快取區存有這封信真的附件數量，就顯示明確數量。如果沒有，代表完全沒有跑過 Phase 2，就顯示 ">0", 不需真的去讀COM物件確認，避免不必要的性能損耗
        ''         Dim cachedFiles As List(Of String) = Nothing
        ''         Dim displayName As String = ">0"
        ''         If _cacheAttachFilename.TryGetValue(m.EntryID, cachedFiles) Then displayName = cachedFiles.Count.ToString()
        ''         items.Add(New ListViewItem({m.Subject,
        ''                                     m.Size.ToString("###,###,##0"),
        ''                                     m.RcvTime.ToShortDateString(),
        ''                                     m.SenderName,
        ''                                     displayName,
        ''                                     m.EntryID}))
        ''     Next
        ''     ListView3.Items.AddRange(items.ToArray())
        '' Else
        ''     ListView3.Items.Add("找不到符合條件的郵件")
        '' End If
        '' ListView3.EndUpdate()

        ' lviCount 避免除以零
        Dim lviCount As Integer = sourceList.Count
        Dim speedText As String = ""
        If elapsedSeconds > 0 AndAlso lviCount > 0 Then speedText = $" ({CInt(lviCount / elapsedSeconds):N0}/sec)"
        PgrsBar1.Text = $"共找到 {lviCount} 封 / 耗時 {elapsedSeconds:0.00} 秒{speedText}"
        PgrsBar2.Text = ""
        _dbg("結束", $"{lviCount} 封 | {elapsedSeconds:0.00}s")

    End Sub
    Private Sub OpenMailByEntryID(entryIDs As List(Of String))
        ' 2026/3/20, by Claude.ai, 建立獨立執行緒fire-and-forget
        ' 讓作業系統跟outlook.exe 自己去做它們的事, 我們不用等它開啟完畢, 可以直接回到自己的程式介面

        ' 2026/04/16 by Gemini 3.1 Pro: 將多筆郵件開啟工作集中在一個 STA 執行緒內完成
        If entryIDs?.Count = 0 Then Return

        ' 2026/04/16 by Gemini 3.1 Pro: 處理多選開啟郵件，超過 10 封需確認 (避免彈出過多視窗)
        ' ✅ 2026/04/21 by Gemini 3.0 flash: 整合多選開啟確認邏輯，避免視窗爆炸
        If entryIDs.Count > 10 Then
            If MessageBox.Show($"確定要同時開啟 {entryIDs.Count:N0} 封郵件嗎？" & vbCrLf & "(一次開啟過多郵件可能會導致 Outlook 暫時無回應)",
                               "大量開啟確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then Return
        End If

        _dbg("開始", $"準備批次開啟 {entryIDs.Count} 封郵件")
        Dim nThread As New Thread(Sub()
                                      Dim nSpace As Outlook.NameSpace = Nothing
                                      Try
                                          nSpace = _olApp.GetNamespace("MAPI")
                                          For Each id In entryIDs
                                              ' by Gemini 3.0 flash, 2026/04/24: 修改為 Object 以支援 RSS (PostItem) 或其他類型項目開啟
                                              Dim olItem As Object = Nothing
                                              Try
                                                  olItem = nSpace.GetItemFromID(id)
                                                  If olItem IsNot Nothing Then
                                                      ' 優先嘗試轉型為常用類型以獲得 Intellisense 支援與早期綁定效能
                                                      Dim mail = TryCast(olItem, Outlook.MailItem)
                                                      If mail IsNot Nothing Then
                                                          mail.Display()
                                                      Else
                                                          Dim post = TryCast(olItem, Outlook.PostItem)
                                                          If post IsNot Nothing Then
                                                              post.Display()
                                                          Else
                                                              ' 2026/04/24: 若非郵件或貼文，嘗試透過晚期綁定呼叫 Display (適用於 MeetingItem, AppointmentItem 等)
                                                              Microsoft.VisualBasic.Interaction.CallByName(olItem, "Display", CallType.Method)
                                                          End If
                                                      End If
                                                  End If
                                              Catch ex As System.Exception
                                                  _dbg("錯誤", $"開啟郵件失敗 (ID: {id}): {ex.Message}")
                                              Finally
                                                  TryMarshalRelease(olItem)
                                              End Try
                                          Next
                                      Catch ex As System.Exception
                                          _dbg("錯誤", $"內部批次開啟邏輯發生錯誤: {ex.Message}")
                                      Finally
                                          TryMarshalRelease(nSpace)
                                      End Try
                                      _dbg("結束", "批次開啟作業完成")
                                  End Sub)
        nThread.SetApartmentState(ApartmentState.STA)    ' ✅ 維持 STA 合規性，避免 COM Marshalling 錯誤
        nThread.IsBackground = True
        nThread.Start()

    End Sub
    Private Sub HandleLv3Delete(lv As ListView)
        ' 2026/06/21 by Simon/Claude Opus 4.8: Tab3 虛擬模式刪除 (行為仿照 HandleLv4Delete，但操作 _lv3MailList + VirtualListSize)
        _dbg("開始")
        Dim selCount As Integer = lv.SelectedIndices.Count
        If selCount = 0 Then Return

        If MessageBox.Show($"確定要將選中的 {selCount} 封郵件移到「刪除郵件」資料夾嗎？", "確認刪除",
                           MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then

            ' 收集選取的 EntryID 與受影響資料夾路徑 (虛擬模式由 SelectedIndices 對應回 _lv3MailList)
            Dim affectedPaths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim entryIDs As New List(Of String)(selCount)
            Dim toRemove As New HashSet(Of String)(StringComparer.Ordinal)
            For Each idx As Integer In lv.SelectedIndices
                If idx >= 0 AndAlso idx < _lv3MailList.Count Then
                    Dim info = _lv3MailList(idx)
                    entryIDs.Add(info.EntryID) : toRemove.Add(info.EntryID)
                    If Not String.IsNullOrEmpty(info.FolderPath) Then affectedPaths.Add(info.FolderPath)
                End If
            Next

            If entryIDs.Count > 0 Then
                _lv3MailList.RemoveAll(Function(m) toRemove.Contains(m.EntryID))   ' 從虛擬資料源移除

                For Each fPath In affectedPaths
                    InvalidateBasicMailCache(fPath)     ' 刪除後手動清理快取資料，避免殘留已刪除郵件的資訊
                    DbDeleteBasicMailInfoByPath(fPath)  ' 刪除後手動清理 DB 資料，避免殘留已刪除郵件的資訊
                Next

                MoveMailsToRecycle(entryIDs)            ' 實體刪除 (移動到同 Store 的刪除郵件資料夾)
                lv.SelectedIndices.Clear()              ' 清選取，避免殘留索引超出新 Size
                lv.VirtualListSize = _lv3MailList.Count ' 設定 Size 即觸發重繪 (參照 ShowLv3Result)
                PgrsBar2.Text = $"已移動 {selCount} 封郵件至刪除郵件資料夾"
            End If
        End If
        _dbg("結束")
    End Sub
#End Region
#Region "  └ 輔助函數"
    Private Function BuildFilterAttachment() As String
        ' 2026-03-16: 大小篩選移到 Bt3_Click Step3b 的 LINQ，
        '   此函數保留但現在只回傳 hasattachment 基礎 filter (與 strFilterHasAttachment 一致)
        '   保留原有大小條件建構邏輯以備日後參考，但 Bt3_Click 已不呼叫此函數
        Dim q As String = Chr(34)
        Return "@SQL=" & q & "urn:schemas:httpmail:hasattachment" & q & " = True"
    End Function
    Private Function GetSizeMultiplier(sizeUnit As String, Optional base1024 As Boolean = False) As Integer
        ' 獲取大小單位的倍數
        Dim multi As Long = If(base1024, 1024, 1000)
        Select Case sizeUnit.ToLower()
            Case "kb" : Return multi
            Case "mb" : Return multi ^ 2
            Case "gb" : Return multi ^ 3
            Case Else : Return 1
        End Select
    End Function
    Private Function GetSelectedEntryIDs(lv As ListView) As List(Of String)
        ''' <summary>
        ''' 取得目前選取節點的 EntryID 清單 (自動判斷 VirtualMode)
        ''' </summary>
        Dim ids As New List(Of String)(32) ' 預設容量為 32，優化小批量選取的性能
        If lv.VirtualMode Then
            ' Tab3 模式: 使用 _lv3MailList
            For Each idx As Integer In lv.SelectedIndices
                If idx >= 0 AndAlso idx < _lv3MailList.Count Then ids.Add(_lv3MailList(idx).EntryID)
            Next
        Else
            ' Tab4 模式: 使用 Item.Tag (MailItemInfo)
            For Each item As ListViewItem In lv.SelectedItems
                If TypeOf item.Tag Is MailItemInfo Then ids.Add(DirectCast(item.Tag, MailItemInfo).EntryID)
            Next
        End If
        Return ids

    End Function
    Private Function GetItemFromPoint(pt As System.Drawing.Point) As MailItemInfo?
        ' 輔助方法：從點座標找出虛擬項目的資料
        Dim item = ListView3.GetItemAt(pt.X, pt.Y)
        If item IsNot Nothing AndAlso item.Index >= 0 AndAlso item.Index < _lv3MailList.Count Then
            Return _lv3MailList(item.Index)
        End If
        Return Nothing
    End Function
    Private Function CalculateSpeedAndETA(totalItems As Integer, processedItems As Integer, elapsedSec As Double) As (Speed As Double, EtaString As String)
        ''' <summary>
        ''' 計算處理速度與預估剩餘時間 (ETA)
        ''' </summary>
        ''' <param name="totalItems">總項目數</param>
        ''' <param name="processedItems">已處理項目數</param>
        ''' <param name="elapsedSec">經過時間 (秒)</param>
        ''' <param name="minTotalThreshold">計算 ETA 的最低總項目數門檻 (例如 10 或 500)</param>
        ''' <returns>Tuple 包含: Speed (項目數/秒) 與 EtaString (格式化字串)</returns>
        ''' by Gemini 3.1 Pro, 2026/05/11
        Dim safeElapsedSec As Double = Math.Max(elapsedSec, 0.001)
        Dim currentSpeed As Double = If(processedItems > 0, processedItems / safeElapsedSec, 0)

        Dim etaStr As String = ""
        If totalItems > 0 AndAlso currentSpeed > 0 Then
            Dim remainingSec As Integer = CInt(Math.Max(0, (totalItems - processedItems) / currentSpeed))
            If remainingSec > 1 Then etaStr = $"，預估剩餘 {remainingSec \ 60:D2}:{remainingSec Mod 60:D2}"
        End If
        Return (currentSpeed, etaStr)
    End Function
#End Region
#End Region

#Region "■ 07 Tab4: 系列郵件"
#Region "  ├ Layer1 UI事件層"
    Private Async Sub Bt4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        ' by Gemini 3 Flash, 2026/04/20: 修改搜尋來源為 Tab4 專屬的 SimTree4
        _dbg("開始")

        Dim cToken As CancellationToken = OkayNowYouHaveToken()
        Dim selectedFolders As New List(Of Folder)(32)
        For Each node In SimTree4.SelectedNodes
            Dim f = TryCast(node.Tag, Folder)
            If f IsNot Nothing Then selectedFolders.Add(f)
        Next

        ' ✅ 2026/04/21 by Gemini 3.0 flash: F5 強化邏輯 - 如果未選擇節點，嘗試使用最後一次搜尋的資料夾清單
        If selectedFolders.Count = 0 AndAlso _tv4PrevSelection.Count > 0 Then
            selectedFolders.AddRange(_tv4PrevSelection)
            _dbg("F5 刷新模式：引用歷史資料夾清單", selectedFolders.Count & " 個資料夾")
        End If

        If selectedFolders.Count = 0 Then Return

        Button4.Enabled = False : Cursor = Cursors.WaitCursor
        Listview4.Items.Clear() : _refreshedList.Clear()
        PgrsBar1.Text = "正在處理..." : PgrsBar2.Text = "開始掃描系列郵件..."
        _tv4PrevSelection = New List(Of Folder)(selectedFolders) ' 記憶最後成功的搜尋目標清單

        Dim sw As Stopwatch = Stopwatch.StartNew()          ' by Claude Sonnet 4.6, 2026/06/07
        Dim swThrottle As Stopwatch = Stopwatch.StartNew()  ' by Gemini, 2026/04/02: 重用秒錶做節流; refactored by Claude Sonnet 4.6, 2026/06/07
        Dim topicDict As New Dictionary(Of String, List(Of MailItemInfo))(StringComparer.OrdinalIgnoreCase)
        Dim progress4 As IProgress(Of ProgressReport) = New Progress(Of ProgressReport)(Sub(p) PgrsBar2.Text = p.Message)

        Try
            ' ✅ 2026/04/21 by Gemini 3.0 flash: 呼叫共用核心 GetUniqueFolderList (內含路徑去重與子資料夾展開)
            ' 2026/04/22 by Gemini 3.1 Pro: 如果在結果模式刷新，SelectedNodes裝的是話題不是Folder。用偽造的 TreeNode 清單包裝歷史 Folder 傳交給底層。
            Dim fakeNodes As New List(Of TreeNode)(32)
            For Each f In selectedFolders
                fakeNodes.Add(New TreeNode() With {.Tag = f})
            Next
            Dim targetTupleList = Await GetUniqueFolderList(fakeNodes, includeSub:=True, progress:=progress4, cToken:=cToken)
            Await PreLoadBasicMailCacheAsync(targetTupleList, cToken)   ' 2026/05/11 by Simon/Claude: SSD 批次預讀，將 DB 中的 basic_maillist 一次拉入記憶體

            Dim targetFolderList = targetTupleList.Select(Function(x) x.folder).ToList()
            Dim processed As Integer = 0
            For Each folder In targetFolderList
                ' by Gemini 3.0 Flash, 2026/04/19: 替換為統一的底層讀取方法 (升級 L2.5)
                Dim infoList = Await GetBasicMailInfo(folder, needTopic:=True, cToken:=cToken)
                For Each item In infoList
                    If item.Topic = "" Then Continue For ' 沒有 Conversation Topic 的信件略過
                    If Not topicDict.ContainsKey(item.Topic) Then topicDict(item.Topic) = New List(Of MailItemInfo)()
                    topicDict(item.Topic).Add(item.Mail)
                Next
                processed += 1

                ' 2026/04/16 by Gemini 3.0 flash: 改用 ThrottleFreq.Hii + SmartThrottle
                Await SmartThrottle(swThrottle, cToken:=cToken, ThrottleFreq.Hii,
                                    Sub()
                                        ' 新版 (2026/05/10 by Simon/Claude: 加入 ETA 顯示，對齊 Tab3 做法)
                                        Dim eta = CalculateSpeedAndETA(targetFolderList.Count, processed, sw.Elapsed.TotalSeconds)
                                        progress4?.Report(New ProgressReport With {.CurrentCount = processed, .TotalCount = targetFolderList.Count,
                                                                                   .Message = $"正在掃描系列郵件: {processed} / {targetFolderList.Count} 個資料夾 ({eta.Speed:F0} 個/秒{eta.EtaString})"})
                                    End Sub)
            Next

            ' ✅ 2026/05/23 by Simon/Claude: 改用 SimTree 內建快照，取代舊版手動備份節點清單，SaveTreeNodeSnap 會在 Nodes.Clear() 之前安全地儲存節點物件與選取狀態
            'If Not _isTv4ResultMode Then SimTree4.SaveTreeNodeSnap("folder-view")
            '_tv4PrevTopicResults = topicDict   ' ✅ 2026/04/20 by Gemini 2.0 Flash: 記憶結果並呼叫共用渲染函數 ' 2026/05/31 by Gemini/Simon: 徹底大掃除：清除所有 F6 與 Group 的殘留
            'RenderLv4Group(topicDict)
            RenderLv4Topic(topicDict)
            sw.Stop()
            PgrsBar1.Text = $"找到 {SimTree4.Nodes.Count} 個系列 / 耗時 {sw.Elapsed.TotalSeconds:0.00} 秒" : PgrsBar2.Text = ""
        Catch ex As System.Exception
            _dbg("結束", "ESC 中斷")
            PgrsBar1.Text = "由使用者中斷。" : PgrsBar2.Text = ""
        Finally
            Button4.Enabled = True
            Cursor = Cursors.Default
            _dbg("結束")
        End Try

    End Sub
    ' Tv4_AfterSelect, Tv4_KeyDown, RenderLv4Group() — 舊版資料夾樹選取連動事件全數移除
    '   因 SimTree4 雙軌模式拆分，結果選取已由 Lv4Topic 專職負責。點擊 SimTree4 只作搜尋參考。
    '   此事件已無業務邏輯需求，故整段註解保留，以備日後參考。
    ' 2026/05/29 by Claude Sonnet 4.6:
    ' ---------------------------------------------------------------------------------------------------------

    Private Sub Lv4Topic_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Lv4Topic.SelectedIndexChanged
        ' ---------------------------------------------------------------
        ' Lv4Topic_SelectedIndexChanged — 點選主旨列 → 右側顯示該系列的郵件清單
        ' 職責: 讀取選取項的 Tag（List(Of MailItemInfo)），
        '       更新 _tv4SelectedTopicMailList，
        '       重置排序狀態後呼叫 RenderLv4Result。
        '
        ' 設計說明：
        '   - 不處理雙擊開信（單擊即連動，雙擊不另外綁定）
        '   - 不做 Await，RenderLv4Result 是同步 UI render，不需要非同步
        '   - Tag 的型別保證來自 RenderLv4Topic，理論上不會 TryCast 失敗；但加 Guard 保護，防止 Clear() 後殘留事件觸發
        ' 2026/05/29 by Simon/Claude: Phase 1 新增，取代 Tv4_AfterSelect 的模式 B 邏輯
        ' 2026/05/30 by Gemini/Simon: 優化為虛擬模式，改用 _lv4TopicList 存底層資料，
        '   SelectedIndices 對應回資料陣列，避免直接綁定 List(Of MailItemInfo) 到 Tag 引起的記憶體壓力與潛在 COM 物件存活問題
        ' ---------------------------------------------------------------
        'If Lv4Topic.SelectedItems.Count = 0 Then Return
        'Dim mailList As List(Of MailItemInfo) = TryCast(Lv4Topic.SelectedItems(0).Tag, List(Of MailItemInfo))
        'If mailList Is Nothing Then Return
        '_tv4SelectedTopicMailList = mailList

        '' 重置排序狀態為預設（日期降冪），對齊原本 Tv4_AfterSelect 的行為
        '_lv4SortOrder = SortOrder.Descending: _lv4LastSortColumn = 2  ' 收到日期欄位 index
        'mailList.Sort(Function(a, b) b.RcvTime.CompareTo(a.RcvTime))
        'RenderLv4Result(mailList)

        ' 💡 虛擬模式下，改由 SelectedIndices 對應回底層資料陣列
        If Lv4Topic.SelectedIndices.Count = 0 Then Return

        Dim idx = Lv4Topic.SelectedIndices(0)
        If idx < 0 OrElse idx >= _lv4TopicList.Count Then Return

        Dim mailList As List(Of MailItemInfo) = _lv4TopicList(idx).Value
        _tv4SelectedTopicMailList = mailList

        ' 重置排序狀態為預設（日期降冪）
        _lv4SortOrder = SortOrder.Descending : _lv4LastSortColumn = 2
        mailList.Sort(Function(a, b) b.RcvTime.CompareTo(a.RcvTime))

        RenderLv4Result(mailList)
        _dbg("結束", $"顯示 {mailList.Count} 封系列郵件")
    End Sub
    Private Sub Lv4Topic_RetrieveVirtualItem(sender As Object, e As RetrieveVirtualItemEventArgs) Handles Lv4Topic.RetrieveVirtualItem
        ' 虛擬模式核心: 當項目進入視野時才動態組裝
        If e.ItemIndex < 0 OrElse e.ItemIndex >= _lv4TopicList.Count Then Return

        Dim kvp = _lv4TopicList(e.ItemIndex)
        Dim lvi As New ListViewItem($"{kvp.Key} ({kvp.Value.Count})")
        lvi.SubItems.Add(kvp.Value.Count.ToString())

        e.Item = lvi
    End Sub
    Private Sub Lv4Topic_GotFocus(sender As Object, e As EventArgs) Handles Lv4Topic.GotFocus
        ' 當 Lv4Topic 取得焦點時，若左側 Panel1 處於收合狀態，自動展開
        ' 適用情境：ESC 從 ListView4 退回 Lv4Topic、或任何其他讓 Lv4Topic 得焦的操作
        ' 2026/06/14 by Simon/Claude Opus 4.8
        Dim sc = GetCurrentSplitter()
        If sc IsNot Nothing AndAlso sc.SplitterDistance <= 20 Then SplitterToggle(sc)
    End Sub
    Private Sub Lv4Topic_KeyDown(sender As Object, e As KeyEventArgs) Handles Lv4Topic.KeyDown
        ' -----------------------------------------------------------------------------------------------------
        ' by Gemini 3.5 Flash, 2026/05/29:
        ' 說明：此事件處理器必須獨立於共通的 HandleLv3Lv4Lv5 處理器，理由如下：
        ' 1. 控制項角色與職責不同：
        '    共通的 HandleLv3Lv4Lv5 主要負責處理 Tab3/4/5 的「郵件詳細清單」ListView3、Listview4、ListView5（顯示個別郵件 MailItemInfo）。
        '    而 Lv4Topic 則是 Tab4 在結果模式下用來展示「郵件主題/系列群組主旨」的左側清單，其 Tag 綁定的是話題下所有郵件的 List(Of MailItemInfo)。
        ' 2. 按鍵互動邏輯截然不同：
        '    - Enter 鍵：共通處理器中按下 Enter 會呼叫 OpenMailByEntryID 直接開啟選中的郵件；而在 Lv4Topic 按下 Enter 是將焦點切換至右側的系列郵件清單 Listview4。
        '    - Escape 鍵：共通處理器中按下 ESC 僅做清除選取並聚焦回對應資料夾樹的簡單行為；
        '                 而在 Lv4Topic 按下 ESC 必須處理 UI 狀態的還原（隱藏 Lv4Topic、顯示 SimTree4，並將焦點移回 SimTree4 恢復資料夾樹模式）。
        ' 3. 維護單一職責與共通簡潔性：
        '    若將兩者強行合併，會在共通處理器中引入大量針對 Lv4Topic 的 If/Select 特判程式碼，違背單一職責原則，使得原本高度一致且簡潔的共通邏輯變得複雜難以維護。
        ' -----------------------------------------------------------------------------------------------------
        _dbg("開始", e.KeyCode.ToString())

        Select Case e.KeyCode
            Case Keys.Enter
                ' 在結果模式下按下 Enter 切換焦點到列表
                If Listview4.Items.Count > 0 Then Listview4.Focus() ' by Claude Sonnet 4.6, 2026/05/29: 移除 _isTv4ResultMode 改用 Lv4Topic.Visible 判定
                e.Handled = True : e.SuppressKeyPress = True

            Case Keys.F5
                ' 按下 F5 等同 Button4 (重新開始掃描系列郵件)
                Button4.PerformClick()
                e.Handled = True

            ' 2026/05/31 by Gemini/Simon: 徹底大掃除：清除所有 F6 與 Group 的殘留
            'Case Keys.F6
                'If _isTv4ResultMode AndAlso _tv4PrevTopicResults IsNot Nothing Then
                '    _tv4GroupSortByCount = Not _tv4GroupSortByCount
                '    RenderLv4Group(_tv4PrevTopicResults)
                '    _dbg("F6 按下：切換排序為", If(_tv4GroupSortByCount, "數量", "主旨"))
                '    e.Handled = True
                'End If

            Case Keys.Escape
                ' 按下 ESC：從結果模式恢復為資料夾模式
                If Not SimTree4.Visible Then
                    ' _isTv4ResultMode = False
                    ' RestoreTreeNodeSnap 或 LoadStoreToTreeView（同原本 Tv4_KeyDown ESC 邏輯）
                    ' by Claude Sonnet 4.6, 2026/05/29: 廢棄變數，直接由下面 Lv4Topic.Visible = False 控制

                    _tv4SelectedTopicMailList = Nothing
                    Listview4.Items.Clear()
                    Lv4Topic.Visible = False : SimTree4.Visible = True : SimTree4.Focus()

                    PgrsBar1.Text = "" : PgrsBar2.Text = ""
                    e.Handled = True : e.SuppressKeyPress = True
                End If
        End Select

    End Sub
    Private Sub Lv4Topic_ColumnClick(sender As Object, e As ColumnClickEventArgs) Handles Lv4Topic.ColumnClick
        If _lv4TopicList Is Nothing OrElse _lv4TopicList.Count = 0 Then Return
        _dbg("開始", $"點擊 Lv4Topic 欄位: {e.Column}")
        Dim sw As Stopwatch = Stopwatch.StartNew()  ' by Claude Sonnet 4.6, 2026/06/07

        ' 1. 決定排序方向
        _lv4TSortOrder = GetNewSortOrder(e.Column, _lv4TLastSortColumn, _lv4TSortOrder) ' 2026/05/30 by Gemini/Simon: 抽取共用函式 GetNewSortOrder，簡化排序狀態切換邏輯
        _lv4TLastSortColumn = e.Column

        Lv4Topic.BeginUpdate()
        Try
            ' 2. 直接對底層資料源進行 LINQ 排序 (結構不同，保持獨立寫法)
            Select Case e.Column
                Case 0 ' 主旨 (Key)
                    _lv4TopicList = If(_lv4TSortOrder = SortOrder.Ascending,
                                       _lv4TopicList.OrderBy(Function(x) x.Key).ToList(),
                                       _lv4TopicList.OrderByDescending(Function(x) x.Key).ToList())
                Case 1 ' 數量 (Value.Count)
                    _lv4TopicList = If(_lv4TSortOrder = SortOrder.Ascending,
                                       _lv4TopicList.OrderBy(Function(x) x.Value.Count).ThenBy(Function(x) x.Key).ToList(),
                                       _lv4TopicList.OrderByDescending(Function(x) x.Value.Count).ThenBy(Function(x) x.Key).ToList())
            End Select

            ' 3. 排序後，原本的選取索引會指錯資料，建議清空選取並重新選定第一筆
            Lv4Topic.SelectedIndices.Clear()
            If _lv4TopicList.Count > 0 Then Lv4Topic.SelectedIndices.Add(0)

            ' 4. 通知 ListView 重新索取畫面上的虛擬項目
            Lv4Topic.Invalidate()
        Finally
            Lv4Topic.EndUpdate()
        End Try

        sw.Stop()
        _dbg("結束", $"虛擬排序完成，耗時 {sw.Elapsed.TotalMilliseconds:F0}ms")
    End Sub
    Private Sub Lv4_DrawItem(sender As Object, e As DrawListViewItemEventArgs) Handles Listview4.DrawItem
        ' by Gemini 3.1 Pro, 2026/04/26: 針對被 Hover 但未選取的項目，交由 DrawSubItem 自行畫上灰底；其餘讓系統自己畫
        If e.Item Is _lastHoveredLvItem AndAlso Not e.Item.Selected Then
            ' 不設 DrawDefault = True，讓系統呼叫 DrawSubItem
        Else
            e.DrawDefault = True
        End If
    End Sub
    Private Sub Lv4_ColumnClick(sender As Object, e As ColumnClickEventArgs) Handles Listview4.ColumnClick

        ' by Gemini 3 Flash, 2026/04/19: Listview4 欄位排序 (實體模式，參考 ListView3 的虛擬模式做法)
        ' by Gemini 3.5 Flash, 2026/05/29: Phase 1 — 加入 Lv4Topic.Visible 防禦性 Guard，只在結果模式下啟用排序
        If Lv4Topic.Visible = False Then Return

        _dbg("開始", $"點擊欄位: {e.Column}")
        Dim sw As Stopwatch = Stopwatch.StartNew()  ' by Claude Sonnet 4.6, 2026/06/07

        ' 取得目前選取節點的郵件清單 (因為是從 SimTree4 選取的，資料存在 Tag 裡)
        ' by Gemini 3.5 Flash, 2026/05/29: Phase 1 — 改讀 _tv4SelectedTopicMailList，不再去碰 SimTree4.SelectedNode.Tag
        Dim mailList As List(Of MailItemInfo) = _tv4SelectedTopicMailList
        If mailList Is Nothing OrElse mailList.Count = 0 Then Return

        ' 切換排序方式
        _lv4SortOrder = GetNewSortOrder(e.Column, _lv4LastSortColumn, _lv4SortOrder)    ' 2026/05/30 by Gemini/Simon: 抽取共用函式 GetNewSortOrder，簡化排序狀態切換邏輯
        _lv4LastSortColumn = e.Column

        ' by Gemini 3.5 Flash, 2026/05/29: Phase 1 — 將排序後的新 List 寫回 _tv4SelectedTopicMailList，不再接觸 SimTree4
        ' 2. 根據點選的欄位, 呼叫共用函式排序 (2026/5/31 by Gemini/simon, 排序邏輯改為直接在這裡呼叫 SortMailList)
        _tv4SelectedTopicMailList = SortMailList(mailList, e.Column, _lv4SortOrder)
        RenderLv4Result(mailList)   ' 重新填入 ListView
        sw.Stop()
        _dbg("結束", "排序完成")

    End Sub
    Private Async Sub Lv4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Listview4.SelectedIndexChanged
        ' ---------------------------------------------------------------
        ' Lv4_UpdateSimilarity — 選定郵件後，非同步計算並更新全列表內文相似度欄 (Index 4)
        ' 2026/04/28 by Simon/Claude: 合併 Simon 架構 + Claude 的 L2.5/L3 分層與 NormalizeMailBody
        '
        ' 架構 (以 Simon 為主):
        '   ① Task.Delay(100) — 避開 SelectedItems 過渡期為空的 Windows 原生兩次觸發
        '   ② _lv4SimCToken 取消機制 — 游標快速移動時取消前次未完成的計算，不讓舊任務蓋掉新結果
        '   ③ 先同步標記全列表（基準=「Base」，其他=「...」），再逐封非同步計算
        '   ④ Jaccard 計算放入 Task.Run 背景執行緒，真正不阻塞 UI
        '   ⑤ EntryID 從 SubItems(5) 讀取（直接、輕量，不做 DirectCast）
        '   ⑥ Body 讀取透過 L2.5 GetMailBody（快取 → L3 COM），不跨群組殘留
        '   ⑦ 比對範圍：全列表（不限同組），方便跨群組發現高相似度郵件
        ' 2026/06/16 by Simon/Claude Opus 4.8: 三階段改「分批漸進式」— 抽出 ProcessLv4SimBatch helper，
        '   每 BATCH_SIZE 封做一輪「讀 body → 平行算 → 顯示」，首屏延遲從「全部 body」降為「第一批 body」；
        '   未算到的列維持下方預先標記的「計算中」
        ' ---------------------------------------------------------------
        Dim lv = DirectCast(sender, ListView)
        If lv.Items.Count = 0 OrElse lv.SelectedItems.Count = 0 Then Return ' 💡 提前檢查，避開 Windows 兩次觸發導致的重複日誌

        If lv.SelectedItems.Count <> 1 Then Return  ' 2026/5/11 by Simon: 僅在單選時才計算相似度, 多選會觸發大量SelectedIndexChanged
        _dbg("開始")

        ' 💡 關鍵修正 1：微小延遲確保選取狀態穩定，避開 SelectedItems 在這100ms內快速移動游標導致的 SelectedItems 為空狀態造成exception
        ' 2026/6/16 by Claude Opus 4.8: 改良版 debounce：先換 token，delay 可被後續選取立即取消
        _lv4SimCToken?.Cancel()
        _lv4SimCToken = New CancellationTokenSource()
        Dim token As CancellationToken = _lv4SimCToken.Token
        Try
            Await Task.Delay(100, token)   ' 游標再動 → 這個 delay 立刻被 cancel，不空等
        Catch ex As OperationCanceledException
            Return
        End Try
        If lv.SelectedItems.Count = 0 Then Return

        ' 取得基準項目的 EntryID（從 SubItems(5) 讀取，輕量直接）
        Dim baseItem As ListViewItem = lv.SelectedItems(0)
        Dim baseEntryID As String = If(baseItem.SubItems.Count > 5, baseItem.SubItems(5).Text, "")
        If String.IsNullOrEmpty(baseEntryID) Then Return

        ' 取得基準郵件正規化 Body（L2.5 快取優先）
        Dim baseBody As String = GetMailBody(baseEntryID)
        If String.IsNullOrEmpty(baseBody) Then Return

        ' 💡 關鍵修正 2：先同步標記全列表初始狀態（UI 執行緒批次寫入，不閃爍）
        lv.BeginUpdate()
        Dim lviCompareList = lv.Items.Cast(Of ListViewItem)().ToList()
        For Each item In lviCompareList
            If item.SubItems.Count <= 4 Then Continue For
            Dim thisID As String = If(item.SubItems.Count > 5, item.SubItems(5).Text, "")
            item.SubItems(4).Text = If(thisID = baseEntryID, "Base", "計算中")
        Next
        lv.EndUpdate()

        ' 2026/06/16 by Simon/Claude Opus 4.8: 分批漸進式主迴圈 — 每批呼叫 ProcessLv4SimBatch，
        '   首批 (BATCH_SIZE 封) 算完即顯示，其餘陸續填入；批次邊界檢查 token 可快速中止
        Const BATCH_SIZE As Integer = 16
        Try
            For Each chunk In lviCompareList.Chunk(BATCH_SIZE)
                If token.IsCancellationRequested Then Return
                Await ProcessLv4SimBatch(lv, chunk, baseBody, baseEntryID, token)
            Next
        Catch ex As OperationCanceledException
            ' 正常取消，不需處理
        End Try
        _dbg("結束")

    End Sub
    Private Sub Lv4_KeyDown(sender As Object, e As KeyEventArgs) Handles Listview4.KeyDown
        ' by Gemini 3.1 Pro, 2026/04/21: Tab4 專屬快捷鍵 (Delete)
        ' 2026/06/14 by Simon/Claude Opus 4.8: F5 分支已移至共用 HandleLv3Lv4Lv5_KeyDown 統一分派，此處移除 (連帶移除 Async)
        _dbg("開始", e.KeyCode.ToString())
        If e.KeyCode = Keys.Delete Then
            _dbg("快捷鍵", "偵測到 Delete (呼叫 HandleLv4Delete)")
            HandleLv4Delete(DirectCast(sender, ListView))
            e.Handled = True
        End If
    End Sub
#End Region
#Region "  ├ Layer2 流程協調層"
    Private Sub RenderLv4Topic(topicDict As Dictionary(Of String, List(Of MailItemInfo)))
        ' ---------------------------------------------------------------
        ' RenderLv4Topic — 將系列郵件掃描結果渲染到 Lv4Topic（左側主旨清單）
        ' 職責: 純 UI render，不計算，不碰 COM。
        '
        ' 與舊版 RenderLv4Group 的差異：
        '   舊版把結果塞進 SimTree4（TreeNode.Tag = mailList）
        '   新版把結果塞進 Lv4Topic（ListViewItem.Tag = mailList）
        '   SimTree4 保持資料夾模式不動，僅透過 Visible 切換顯示
        '
        ' 呼叫端: Bt4_Click 掃描完成後、F6 切換排序後
        ' 2026/05/29 by Simon/Claude: Phase 1 新增，取代 RenderLv4Group 對 SimTree4 的操作
        ' 2026/05/30 by Gemini/Simon: 優化為虛擬模式，改用 _lv4TopicList 存底層資料，RenderLv4Topic 只負責排序與刷新 ListView 的 VirtualListSize，實際項目由 RetrieveVirtualItem 動態組裝
        ' 2026/05/31 by Gemini/Simon: 徹底大掃除：清除所有 F6 與 Group 的殘留
        ' ---------------------------------------------------------------
        _dbg("開始")
        If topicDict Is Nothing Then Return

        ' ── 排序：拿掉 IF 判斷，直接預設：數量多的排前面，數量相同則按主旨排 ──
        Dim sortedItems = topicDict.Where(Function(kvp) kvp.Value.Count > 1).
                                    OrderByDescending(Function(kvp) kvp.Value.Count).ThenBy(Function(kvp) kvp.Key)

        ' ── 排序後存入 _lv4TopicList (取代實體 ListViewItem) ──
        _lv4TopicList = sortedItems.ToList()    ' 2026/5/30 by Gemini/Simon

        Lv4Topic.BeginUpdate()
        Lv4Topic.VirtualListSize = 0    ' 先歸零強制重置
        Lv4Topic.VirtualListSize = _lv4TopicList.Count
        Lv4Topic.EndUpdate()

        ' ── 切換顯示：Lv4Topic 上台，SimTree4 退後 ──
        SimTree4.Visible = False : Lv4Topic.Visible = True : Lv4Topic.Focus()

        ' (虛擬模式改用 SelectedIndices) ── ' 2026/5/30 by Gemini/Simon
        If _lv4TopicList.Count > 0 Then Lv4Topic.SelectedIndices.Add(0)

        ' 更新 ProgressBar，移除了排序狀態文字
        PgrsBar1.Text = $"找到 {_lv4TopicList.Count} 個系列"
        _dbg("結束", $"{_lv4TopicList.Count} 個系列")
    End Sub
    Private Sub RenderLv4Result(mailList As List(Of MailItemInfo))
        _dbg("開始")
        Listview4.Tag = mailList
        Listview4.BeginUpdate()
        Listview4.Items.Clear()
        'Listview4.Groups.Clear()

        If mailList Is Nothing OrElse mailList.Count = 0 Then
            Listview4.EndUpdate() : Return
        End If

        '' by Gemini 3 Flash, 2026/04/20: 實作智慧分組 (排除 Re:/Fw:) 與動態排序邏輯, 確保資料清單被記住，以便 F6 切換時使用
        '' 2026/05/31 by Gemini/Simon: 徹底大掃除：清除所有 F6 與 Group 的殘留
        '' 1. 執行分組 (LINQ GroupBy 智慧清理後的主旨)
        'Dim groups = mailList.GroupBy(Function(m) GetCleanSubject(m.Subject))
        '' 2. 依照排序模式對「組」進行排序
        'Dim sortedGroups = If(_lv4GroupSortByCount, groups.OrderByDescending(Function(g) g.Count()).ThenBy(Function(g) g.Key), groups.OrderBy(Function(g) g.Key))
        '' 3. 逐組渲染到 UI
        'For Each group In sortedGroups
        '    ' 建立組標題：主旨 (數量封)
        '    Dim groupHeader As String = $"{group.Key} ({group.Count} 封)"
        '    'Dim lvGroup As New ListViewGroup(group.Key, groupHeader)
        '    'Listview4.Groups.Add(lvGroup)

        ' ✅ 2026/04/20 by Gemini 2.0 Flash: 連動 Column Header 的點擊排序, 根據全域變數 _lv4LastSortColumn 對組內項目進行動態排序
        ' 2026/05/31 by Gemini/Simon: 徹底大掃除：清除所有 F6 與 Group 的殘留
        ' 1. 直接對整包 mailList 進行「全域排序」(徹底捨棄 GroupBy 邏輯)
        Dim sortedItems As IEnumerable(Of MailItemInfo) = mailList
        Select Case _lv4LastSortColumn
            Case 0          ' 主旨
                sortedItems = If(_lv4SortOrder = SortOrder.Ascending, sortedItems.OrderBy(Function(m) m.Subject), sortedItems.OrderByDescending(Function(m) m.Subject))
            Case 1          ' 郵件大小
                sortedItems = If(_lv4SortOrder = SortOrder.Ascending, sortedItems.OrderBy(Function(m) m.Size), sortedItems.OrderByDescending(Function(m) m.Size))
            Case 2          ' 收到日期
                sortedItems = If(_lv4SortOrder = SortOrder.Ascending, sortedItems.OrderBy(Function(m) m.RcvTime), sortedItems.OrderByDescending(Function(m) m.RcvTime))
            Case 3          ' 寄件者
                sortedItems = If(_lv4SortOrder = SortOrder.Ascending, sortedItems.OrderBy(Function(m) m.SenderName), sortedItems.OrderByDescending(Function(m) m.SenderName))
            Case Else       ' 預設 (或點擊其他無效欄位時)
                sortedItems = sortedItems.OrderByDescending(Function(m) m.RcvTime)
        End Select

        ' 收集該組的所有項目，再一次性 AddRange (by Gemini 3.0 flash, 2026/04/21)
        ' 預分配容量為 512，優化重複郵件掃描結果的 UI 清單組裝 (by Gemini 3 Flash, 2026/05/04)
        ' 2026/05/31 by Gemini/Simon: 徹底大掃除：清除所有 F6 與 Group 的殘留
        ' 2. 建立 ListViewItem 清單，一次性 AddRange 提升效能
        Dim itmToAdd As New List(Of ListViewItem)(mailList.Count)
        For Each mailItem In sortedItems
            ' by Gemini 3.0 Flash, 2026/04/20: 郵件大小改為位元組(精細), 日期格式統一 yyyy/MM/dd (補零+置中需求)
            Dim lvi As New ListViewItem({mailItem.Subject,
                                         mailItem.Size.ToString("N0"),
                                         mailItem.RcvTime.ToString("yyyy/MM/dd HH:mm:ss"),
                                         mailItem.SenderName, " - ",
                                         mailItem.EntryID})
            'lvi.Group = lvGroup
            lvi.Tag = mailItem ' by Gemini 3.0 flash, 2026/04/21: 直接存入物件避開 Index 錯位問題
            itmToAdd.Add(lvi)
        Next

        ' 3. 建立 ListViewItem 清單，一次性 AddRange 提升效能
        Listview4.Items.AddRange(itmToAdd.ToArray())
        Listview4.EndUpdate()

        ' 4. 更新狀態列反饋 (by Gemini 3 Flash, 2026/04/20)
        PgrsBar2.Text = $"系列選中：共 {mailList.Count:N0} 封郵件"
        _dbg("結束")

    End Sub
    Private Sub HandleLv4Delete(lv As ListView)
        ' by Gemini 3 Flash, 2026/04/20: 處理 Listview4 的刪除邏輯
        _dbg("開始")
        Dim selCount As Integer = lv.SelectedItems.Count
        If selCount = 0 Then Return

        If MessageBox.Show($"確定要將選中的 {selCount} 封郵件移到「刪除郵件」資料夾嗎？", "確認刪除",
                           MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then

            ' 2026/5/11 simon: 收集受影響的資料夾路徑，供後續清理快取與DB使用
            Dim affectedPaths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            Dim entryIDs As New List(Of String)(64) ' 預分配容量，優化批量刪除的 ID 收集 (by Gemini 3 Flash, 2026/05/04)
            ' 2026/04/28 by Gemini 3.1 Pro: 改由 ListView 本身的 Tag 取回清單，並以 EntryID 作為刪除比對基準，避免 Structure 預設比對失敗
            Dim mailList As List(Of MailItemInfo) = TryCast(lv.Tag, List(Of MailItemInfo))
            If mailList IsNot Nothing Then
                ' 先收集 ID 並從原始清單中移除
                For Each item As ListViewItem In lv.SelectedItems
                    If TypeOf item.Tag Is MailItemInfo Then
                        Dim info = DirectCast(item.Tag, MailItemInfo)
                        entryIDs.Add(info.EntryID)
                        mailList.RemoveAll(Function(m) m.EntryID = info.EntryID)

                        ' 2026/5/11 simon: 收集受影響的資料夾路徑，供後續清理快取與DB使用
                        If Not String.IsNullOrEmpty(info.FolderPath) Then affectedPaths.Add(info.FolderPath)
                    End If
                Next

                For Each fPath In affectedPaths
                    InvalidateBasicMailCache(fPath)     ' 2026/5/11 by Simon: 刪除後手動清理快取資料，避免殘留已刪除郵件的資訊
                    DbDeleteBasicMailInfoByPath(fPath)  ' 2026/5/11 by Simon: 刪除後手動清理 DB 資料，避免殘留已刪除郵件的資訊
                Next

                MoveMailsToRecycle(entryIDs)            ' 實體刪除 (移動到預設刪除資料夾)
                RenderLv4Result(mailList)               ' 重新整理 UI
                PgrsBar2.Text = $"已移動 {selCount} 封郵件至刪除郵件資料夾"
            End If
        End If
        _dbg("結束")

    End Sub
    Private Sub MoveMailsToRecycle(entryIDs As List(Of String))
        ' by Gemini 3 Flash, 2026/04/20: 核心移動邏輯 (Layer3)
        ' 建立背景執行緒執行移動，避免 UI 卡死

        ' 2026/04/30 by Simon/Claude: 修正跨 Store 移動失敗的問題
        '   原本使用 ns.GetDefaultFolder(olFolderDeletedItems) 取得預設帳號的刪除資料夾，
        '   但在多 PST 環境下，非主 Store 的郵件跨 Store Move() 會靜默失敗。
        '   修正：改用 m.Delete() 自動移入同一 Store 的刪除郵件資料夾，
        '   行為與 Outlook UI 按 Delete 鍵一致，跨 Store 問題徹底消除。

        ' 2026/05/06 by Gemini 3 Flash: 放寬型別限制，改用 Object 以支援 RSS (PostItem), 會議, 草稿等各種項目
        _dbg("開始")
        Dim th As New Thread(Sub()
                                 Dim ns As Outlook.NameSpace = Nothing
                                 Try
                                     ns = _olApp.GetNamespace("MAPI")
                                     For Each id In entryIDs
                                         Dim item As Object = Nothing ' 改用 Object 以相容多種 Outlook 型別
                                         Try
                                             item = ns.GetItemFromID(id)
                                             ' 利用 Late Binding 呼叫 Delete 方法 (大多數 Outlook Item 皆具備此方法)
                                             ' 這樣不論是 MailItem, PostItem (RSS), MeetingItem 都能處理
                                             item?.Delete()  ' ← 自動移至同 Store 的刪除郵件，不需指定 destFolder
                                         Catch ex As System.Exception
                                             _dbg("刪除失敗", $"ID: {id}, Error: {ex.Message}")
                                         Finally
                                             TryMarshalRelease(item)
                                         End Try
                                     Next
                                 Catch ex As System.Exception
                                     _dbg("移動郵件失敗", ex.Message)
                                 Finally
                                     TryMarshalRelease(ns)
                                 End Try
                             End Sub)
        th.SetApartmentState(ApartmentState.STA)
        th.IsBackground = True
        th.Start()

    End Sub
    Private Async Function ProcessLv4SimBatch(lv As ListView, chunk As ListViewItem(), baseBody As String, baseEntryID As String, token As CancellationToken) As Task
        ' ---------------------------------------------------------------
        ' ProcessLv4SimBatch — 處理單一批次的「讀 body → 平行算 Jaccard → 批次更新 UI」
        ' 2026/06/16 by Simon/Claude Opus 4.8: 自 Lv4_SelectedIndexChanged 抽出，三階段邏輯原樣不變，僅作用範圍縮為一批 (chunk)
        ' ---------------------------------------------------------------

        ' 💡 2026/04/30 by Gemini 3.1 Pro: 兩階段處理。第一階段在 UI Context 循序拿 Body，確保若是 Cache Miss 去讀 COM 的安全性。
        ' 預分配容量為本批封數，處理大量郵件內文比對時減少頻繁 Resize (by Gemini 3 Flash, 2026/05/04)
        ' 💡 2026/05/09 by Gemini 3.0 flash: 優化非同步頻率。每處理一批(例如 10 封)才 Yield 一次讓 UI 喘氣，減少頻繁切換的負擔
        Dim mBodyList As New List(Of (Item As ListViewItem, TargetBody As String))(chunk.Length)
        Dim processedCount As Integer = 0
        For Each item In chunk
            If token.IsCancellationRequested Then Return
            If item.SubItems.Count <= 4 Then Continue For

            Dim targetID As String = If(item.SubItems.Count > 5, item.SubItems(5).Text, "")
            If targetID = baseEntryID OrElse String.IsNullOrEmpty(targetID) Then Continue For

            Dim targetBody As String = GetMailBody(targetID)
            If String.IsNullOrEmpty(targetBody) Then
                item.SubItems(4).Text = "失敗" : Continue For
            End If
            mBodyList.Add((item, targetBody))
            processedCount += 1
            If processedCount Mod 10 = 0 Then Await PreciseDelay(1) ' 每 10 封釋放一次 UI 執行緒，兼顧流暢度與效率 (by Gemini 3.0 flash, 2026/05/09)
        Next

        If token.IsCancellationRequested Then Return
        If mBodyList.Count = 0 Then Return  ' 2026/06/16 by Simon/Claude Opus 4.8: 整批都是 Base/失敗時 Count=0，避開 results(-1) 越界

        ' 💡 2026/04/30 by Gemini 3.1 Pro: 第二階段純 CPU 運算。剝離 UI 與 COM，利用多核心火力全開
        ' 2026/6/18 by Claude Opus 4.8, Q9: 原有的字元集JaccardSimilarity 改用新建的bigram版本
        Dim results(mBodyList.Count - 1) As Double
        Dim baseSet = BuildBigramSet(baseBody)  ' base bigram 集合提到迴圈外建一次，避免每個 target 重建
        Await Task.Run(Sub()
                           Parallel.For(0, mBodyList.Count, Sub(i)
                                                                If token.IsCancellationRequested Then Return
                                                                Dim targetBody = mBodyList(i).TargetBody
                                                                ' results(i) = CalculateSimilarity(baseBody, targetBody)
                                                                ' results(i) = CharlesHash(baseBody, targetBody)
                                                                ' results(i) = JaccardSimilarity(baseBody, targetBody)   ' 舊: 字元集 Jaccard
                                                                ' 2026/6/18 by Claude Opus 4.8: 改 bigram Jaccard，與 Tab5 一致、中文更準
                                                                results(i) = BigramJaccardSimilarity(baseSet, BuildBigramSet(targetBody))
                                                            End Sub)
                       End Sub)
        If token.IsCancellationRequested Then Return

        ' 💡 2026/04/30 by Gemini 3.1 Pro: 第三階段批次更新 UI
        lv.BeginUpdate()
        For i = 0 To mBodyList.Count - 1
            Dim item = mBodyList(i).Item
            If item.ListView IsNot Nothing Then item.SubItems(4).Text = $"{CInt(results(i) * 100)}%"
        Next
        lv.EndUpdate()
    End Function
    Private Async Function RefreshLv4Result(lv As ListView) As Task
        ' by Gemini 3 Flash, 2026/04/20: 重新讀取目前系列郵件的最新資訊並更新 MailItemInfo
        ' ✅ by Gemini 3.0 flash, 2026/04/21: 修正控制項名稱為 SimTree4
        ' by Gemini 3.5 Flash, 2026/05/29: Phase 1 — 改讀 _tv4SelectedTopicMailList，不再去碰 SimTree4.SelectedNode.Tag
        _dbg("開始")
        Dim mailList As List(Of MailItemInfo) = _tv4SelectedTopicMailList
        If mailList Is Nothing OrElse mailList.Count = 0 Then Return

        _isUserBusy = True : Cursor = Cursors.WaitCursor
        Try
            ' 2026/06/14 by Simon/Claude Opus 4.8: 改走共用核心 RefreshLviCore (依數量自動 A/B)，移除原逐封 GetItemFromID 內聯
            Dim targetList As New List(Of (lst As List(Of MailItemInfo), idx As Integer))(mailList.Count)
            For i As Integer = 0 To mailList.Count - 1 : targetList.Add((mailList, i)) : Next
            Dim stats = Await RefreshLviCore(targetList, readAttachCount:=False, ct:=CancellationToken.None)

            RenderLv4Result(mailList)   ' 重新填寫列表 (保留目前排序狀態，資料是原地更新)
            PgrsBar1.Text = $"已重新讀取 {stats.Updated} 封 (失效 {stats.NotFound}, 錯誤 {stats.Errored})。" : PgrsBar2.Text = ""
        Catch ex As OperationCanceledException
            PgrsBar1.Text = "刷新已取消。"
        Catch ex As System.Exception
            _dbg("重新讀取發生錯誤", ex.Message)
        Finally
            _isUserBusy = False : Cursor = Cursors.Default
            _dbg("結束")
        End Try
    End Function
#End Region
#Region "  └ 輔助函數"
    Private Function CharlesHash(strA As String, strB As String) As Double
        ''' <summary>
        ''' 2006/06/22 Algorithm by Charles Wu / Translated & Optimized by Gemini 3 Flash, 2026/04/26
        ''' 基於字元雜湊統計的快速相似度比對 (O(N) 複雜度，適合長內文)
        ''' </summary>
        ' by Gemini 3 Flash, 2026/04/26: 傳承自 Charles Wu 2006 年的經典演算法，專攻常用字元相似度對比
        If String.IsNullOrEmpty(strA) OrElse String.IsNullOrEmpty(strB) Then Return 0

        ' 使用 Byte 陣列記錄 Unicode 字元出現狀態 (0:無, 1:A有, 2:B有, 3:兩者皆有)
        Dim charTable(65535) As Byte
        Dim lSum(3) As Integer

        ' 1. 處理字串 A
        For Each c In strA
            Dim code As Integer = AscW(c)
            If charTable(code) = 0 Then charTable(code) = 1
        Next

        ' 2. 處理字串 B
        For Each c In strB
            Dim code As Integer = AscW(c)
            If charTable(code) = 0 Then
                charTable(code) = 2
            ElseIf charTable(code) = 1 Then
                charTable(code) = 3
            End If
        Next

        ' 3. 統計範圍擴大：
        ' - ASCII 區段 (0 ~ 255): 包含英文、數字、常用符號
        ' - CJK 中日韓區段 (&H2E80 ~ &H9FBF): 包含繁簡中、日文、韓文
        ' 註：若要極致精準可掃描全表 (0 To 65535)，但掃描特定區段效能更好
        For i As Integer = 0 To 255 : lSum(charTable(i)) += 1 : Next            ' 統計 ASCII
        For i As Integer = &H2E80 To &H9FBF : lSum(charTable(i)) += 1 : Next    ' 統計 CJK

        ' 4. 計算相似度
        Dim denominator As Integer = lSum(1) + lSum(2) + lSum(3)
        If denominator = 0 Then Return 0
        Return lSum(3) / denominator
    End Function
    Private Function CharlesHash_Ultimate(strA As String, strB As String) As Double
        ' -----------------------------------------------------------------------------------------
        ' 保留原版 CharlesHash 的邏輯，但引入 System.Buffers.ArrayPool 來重複利用那 64KB 的陣列，做到零記憶體分配 (Zero Allocation)
        ' 2026/6/9 by Gemini 3.1 Pro: Ultimate 版本，使用 ArrayPool 重用陣列，徹底消除 GC 分配，適合大量相似度計算的場景
        '
        ' 極致效能建議:
        '   如果字串通常很短（幾百字內）: 使用 HashSet 會比較快，因為省下了掃描大陣列與配置大記憶體的時間。
        '   如果比對字串很長（長篇本文）: 處理量極大，兩者都有缺陷。原版 CharlesHash 會塞爆 GC，HashSet 版 CPU 運算太重。
        ' -----------------------------------------------------------------------------------------

        If String.IsNullOrEmpty(strA) OrElse String.IsNullOrEmpty(strB) Then Return 0

        ' 從共用池中借出一個至少 65536 大小的陣列 (避免每次 New 產生 GC 垃圾)
        Dim charTable As Byte() = ArrayPool(Of Byte).Shared.Rent(65536)
        Try
            ' ArrayPool 借出來的陣列可能殘留舊資料，必須清空前面我們會用到的區塊
            Array.Clear(charTable, 0, 65536)
            Dim lSum(3) As Integer

            ' 1. 處理字串 A
            For Each c In strA
                Dim code As Integer = AscW(c)
                If charTable(code) = 0 Then charTable(code) = 1
            Next

            ' 2. 處理字串 B
            For Each c In strB
                Dim code As Integer = AscW(c)
                If charTable(code) = 0 Then
                    charTable(code) = 2
                ElseIf charTable(code) = 1 Then
                    charTable(code) = 3
                End If
            Next

            ' 3. 統計範圍擴大：
            ' - ASCII 區段 (0 ~ 255): 包含英文、數字、常用符號
            ' - CJK 中日韓區段 (&H2E80 ~ &H9FBF): 包含繁簡中、日文、韓文
            ' 註：若要極致精準可掃描全表 (0 To 65535)，但掃描特定區段效能更好
            For i As Integer = 0 To 255 : lSum(charTable(i)) += 1 : Next            ' 統計 ASCII
            For i As Integer = &H2E80 To &H9FBF : lSum(charTable(i)) += 1 : Next    ' 統計 CJK

            ' 4. 計算相似度
            Dim denominator As Integer = lSum(1) + lSum(2) + lSum(3)
            If denominator = 0 Then Return 0
            Return lSum(3) / denominator

        Finally
            ' 確保一定會歸還陣列給共用池
            ArrayPool(Of Byte).Shared.Return(charTable)
        End Try
    End Function
    Private Function JaccardSimilarity(strA As String, strB As String) As Double

        ' ---------------------------------------------------------------
        ' JaccardSimilarity — 字元集 Jaccard 相似度（Charles Wu 算法現代化版）
        ' 2026/04/28 by Simon/Claude: 基於 GetSimRate_New (2006/06/22 Algorithm by Charles Wu)
        '   原版只統計 U+4E00~U+9FBF 中文字範圍；新版擴展至全 Unicode 字元集
        '   原版用 Integer 陣列 65536 格；新版用 HashSet(Of Char)，語意更清晰且省記憶體
        '   Simon 優化：先找較小的集合走迴圈，減少一半的迭代次數
        '
        ' 算法原理 (Jaccard similarity):
        '   相似度 = |A ∩ B| / |A ∪ B|
        '   其中 A、B 為各字串的唯一字元集合（重複字元只算一次）
        '   包含/被包含關係的郵件（轉寄引用）也能得到高相似度
        '
        ' 複雜度: O(n+m)，遠優於 Levenshtein 的 O(n×m)
        '         兩封幾千字的信也能在幾毫秒內完成
        ' ---------------------------------------------------------------

        If String.IsNullOrEmpty(strA) AndAlso String.IsNullOrEmpty(strB) Then Return 1.0
        If String.IsNullOrEmpty(strA) OrElse String.IsNullOrEmpty(strB) Then Return 0.0

        ' 將字串轉換為唯一字元的 Hash 集合
        Dim setA As New HashSet(Of Char)(strA)
        Dim setB As New HashSet(Of Char)(strB)

        '' 計算交集數量-1 (必須在 IntersectWith 之前記錄原始大小)
        'Dim countA As Integer = setA.Count
        'Dim countB As Integer = setB.Count
        'setA.IntersectWith(setB)                    ' 直接與setB 交集, setA 就會被修改成交集後的集合了
        'Dim intersectCount As Integer = setA.Count  ' 交集後直接讀取，O(1)

        ' 計算交集數量-2 (使用迴圈比呼叫 IntersectWith 更省記憶體與效能)
        '   因為這裡根本不需要修改 setA，只需要計數，原始迴圈反而是更精準的選擇
        '   IntersectWith() 的優勢是語意清楚、後續可直接使用縮減後的集合，但在純計數場景下並無好處。

        ' 找出較小的集合來走迴圈，性能更好
        If setA.Count > setB.Count Then
            Dim temp = setA : setA = setB : setB = temp
        End If

        Dim intersectCount As Integer = 0
        For Each c In setA
            If setB.Contains(c) Then intersectCount += 1
        Next

        ' 聯集數量 = A大小 + B大小 - 交集數量 (數學集合論公式，避免呼叫 UnionWith)
        Dim unionCount As Integer = setA.Count + setB.Count - intersectCount
        If unionCount = 0 Then Return 0
        Return intersectCount / unionCount

    End Function
#End Region
#End Region

End Class
