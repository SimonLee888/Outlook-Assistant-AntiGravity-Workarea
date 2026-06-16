Imports System.Buffers
Imports System.Collections.Concurrent
Imports System.Runtime.InteropServices
Imports System.Threading
Imports Microsoft.Data.Sqlite
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
    Private _lv5LastSortColumn As Integer = -1                  ' by Simon/Claude, 2026/05/10: Tab5 欄位排序狀態
    Private _lv5SortOrder As SortOrder = SortOrder.Ascending    ' by Simon/Claude, 2026/05/10: Tab5 欄位排序狀態

    Const REFRESH_BATCH_THRESHOLD As Integer = 42               ' 2026/06/14 by Simon/Claude Opus 4.8: <42 走A、>=42 走B (涵蓋 <41→A、>42→B，並補齊 41→A/42→B)
    Private ctxMenuRefresh As ContextMenuStrip = Nothing        ' 2026/06/14 by Simon/Claude Opus 4.8: Lv3/4/5 共用的右鍵刷新選單 (單一實例，初始化於 InitLv3Lv4Lv5RefreshMenu)
    Private _refreshedList As New HashSet(Of String)(StringComparer.Ordinal)    ' 2026/06/15 by Simon/Claude Opus 4.8: 記錄本次（或累積數次）刷新成功的 EntryID，供 Lv3/4/5 以藍色字體標示；新搜尋開始時清除

    Private pnlOptions_tab3 As Panel
    Private _lv3MailList As New List(Of MailItemInfo)(4096)     ' by Gemini, 2026/04/10: Tab3 顯示資料庫 (虛擬模式核心)    ' 預分配容量為 4096，因應 Tab3 可能載入的大量郵件資訊，顯著降低記憶體配置開銷 (by Gemini 3 Flash, 2026/05/04)
    ' Private _isTab3_Stop As Boolean                           ' 2026/04/05 by Gemini: 已併入全域 _cancelRequested，不再單獨使用專屬旗標以簡化邏輯內容流程處理機制
    Private _lv4SimCToken As CancellationTokenSource = Nothing  ' by Gemini 3 Flash, 2026/04/26: 用於游標快速移動時, 取消前一次未完成的相似度計算任務
    ' _tab4FolderTreeNodesBackup / _tab4LastClickedFolderNode 已移除'   節點快照改由 SimTree4.SaveTreeNodeSnap("folder-view") 內部管理，不再需要 Form1 level 備份變數  ' 2026/05/23 by Simon/Claude
    ' Private _isTv4ResultMode As Boolean = False               ' ✅ 2026/04/20 by Gemini 2.0 Flash: 標記 Tab4 左側樹目前顯示的是搜尋結果模式 ' by Claude Sonnet 4.6, 2026/05/29: 已廢棄，改用 Lv4Topic.Visible 代替
    Private _tv4PrevSelection As New List(Of Folder)(32)        ' ✅ 2026/04/21 by Gemini 3.0 flash: 記憶最後一次搜尋的多個資料夾  ' 預分配容量為 32，足以涵蓋多數搜尋路徑結構，減少陣列頻繁 Resize 開銷 (by Gemini 3 Flash, 2026/05/04)
    Private _tv4SelectedTopicMailList As List(Of MailItemInfo) = Nothing                        ' 2026/5/29 by Simon/Claude: 將SimTree4的雙重模式拆分，取代 SimTree4.SelectedNode.Tag 作為跨函數的資料橋樑, 供 F6 快速切換使用
    Private _lv4TopicData As New List(Of KeyValuePair(Of String, List(Of MailItemInfo)))(4096)  ' 2026/5/30 by Gemini, Lv4Topic 虛擬模式的資料來源
    ' ── Tab5 SimTree 多選控制項 (2026/05/01 by Claude: 取代舊版 TreeView5，對齊 Tab1~4 操作行為) ──
    'Private Listview6 As ListView = Nothing                    ' by Gemini 3 Flash, 2026/04/20: 動態建立的統計列表
    Private rbExactMatch, rbFuzzyMatch As New RadioButton()     ' tab5 用到的radio button
    Private _includeSubTab5 As Boolean = True                   ' Tab5 是否含子資料夾，由 CheckSubFolder5 CheckBox 控制，預設 True
    Private _tv5PrevSearchMode As Boolean = True                ' by Gemini 3 Flash, 2026/05/06: 記憶最後一次掃描的模式
    ' 2026/5/31 by Gemini/Simon: 徹底大掃除：清除所有 F6 與 Group 的殘留
    'Private _tv4GroupSortByCount As Boolean = True              ' ✅ 2026/04/20 by Gemini 2.0 Flash: 記錄排序方式 (True=數量, False=主旨)
    'Private _lv4GroupSortByCount As Boolean = False             ' by Gemini 3 Flash, 2026/04/20: 記錄 Tab4 Listview4 分組排序模式 (False:按主旨, True:按數量)
    'Private _tv4PrevTopicResults As Dictionary(Of String, List(Of MailItemInfo)) = Nothing  ' ✅ 2026/04/20 by Gemini 2.0 Flash: 記憶搜尋結果，供 F6 操作使用
    Private _lv5PrevGroupResults As Dictionary(Of String, List(Of MailItemInfo)) = Nothing  ' by Gemini 3 Flash, 2026/05/06: 記憶 Tab5 掃描結果，供刪除後重新渲染使用
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
        ProgressBar1.Text = "準備中" : ProgressBar2.Text = "" : Cursor = Cursors.WaitCursor
        pnlOptions_tab3.Enabled = False : SimTree3.Enabled = False
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
            Dim progressTree = New Progress(Of ProgressReport)(Sub(p) ProgressBar2.Text = p.Message)
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

            ProgressBar1.Text = "正在讀取..."
            ProgressBar2.Text = $"準備掃描 {folderList.Count:N0} 個資料夾 (相依包含共計 {totalMailCount:N0} 封信)..."
            Await Task.Yield()

            ' ── Step 3: 收集含附件的郵件清單 (透過 Layer2.5 快取) ──
            Dim progressPhase1 As IProgress(Of ProgressReport) = New Progress(Of ProgressReport)(Sub(p) ProgressBar2.Text = p.Message)
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
                                                                                            .Message = $"Phase 1 (載入郵件清單): {processed} / {folderList.Count} 個資料夾 ({eta.Speed:F0} 個/秒{eta.EtaString})"})
                                        End Sub)
                Next
                _dbg("結束step3迴圈")
            Catch ex As OperationCanceledException
                ' by Gemini, 2026/04/12: 捕捉 ESC 中斷，結算目前已載入的部分郵件清單
                _dbg(" ├ 中斷", $"Step 3 已中斷，結算目前已載入的 {targetMails.Count:N0} 封")
                ProgressBar1.Text = "由使用者中斷"
            End Try
            Dim tStep3_AttachMailLoop = swStep.Elapsed.TotalMilliseconds : swStep.Restart() ' by Gemini 3.0 flash, 2026/04/16: 改名以區分 (GetAttachMailList Loop)

            ' ── Pipeline 過濾 1: 大小篩選 ──
            If CheckSize.Checked Then targetMails = FilterBySize(targetMails)   ' 這裡五萬筆資料只花<3ms

            ' ── Pipeline 過濾 2: 附件條件深層篩選 ──
            Dim hasKeyword = CheckAttachName.Checked AndAlso TextBox3.Text.Trim.Length > 0
            If hasKeyword OrElse CheckAttCount.Checked Then
                Dim progressPhase2 = New Progress(Of ProgressReport)(Sub(p) ProgressBar2.Text = p.Message)
                targetMails = Await FilterByAttachDetailsAsync(targetMails, progressPhase2, cToken:=cToken)
            End If
            Dim tStep5_DetailsFiltering = swStep.Elapsed.TotalMilliseconds : swStep.Stop() ' by Gemini 3.0 flash, 2026/04/16: 改名以區分 (Details Filtering)

            ' ── 終極 Mapping 與顯示結果 ──
            sw.Stop()
            ' by Gemini 3.0 flash, 2026/04/16: 依照使用者要求，將分段耗時拆分為多列顯示於 Debug ListView
            _dbg("⌛ 效能 (1/4) - GetUniqueFolderList", $"{tStep2_UniqueList:F0}ms")
            _dbg("⌛ 效能 (2/4) - GetMailCount", $"{tStep2_MailCountLoop:F0}ms")
            _dbg("⌛ 效能 (3/4) - GetAttachMailList", $"{tStep3_AttachMailLoop:F0}ms")
            _dbg("⌛ 效能 (4/4) - FilterByAttachDetailsAsync", $"{tStep5_DetailsFiltering:F0}ms")
            _dbg("⌛ 效能 (總計) - Total", $"{sw.Elapsed.TotalMilliseconds:F0}ms")

            ShowLv3Result(targetMails, sw.Elapsed.TotalSeconds)
        Catch ex As OperationCanceledException
            _dbg("結束", "ESC 中斷")
            ProgressBar1.Text = "由使用者中斷。" : ProgressBar2.Text = ""
        Catch ex As System.Exception
            MessageBox.Show("搜尋發生錯誤: " & ex.Message, "錯誤")
            _dbg("       ├ 錯誤", ex.Message) ' by Gemini, 2026/04/11: Level 3
        Finally
            ' ── 無論如何都解鎖 UI ──
            SimTree3.Enabled = True : Button3.Enabled = True
            pnlOptions_tab3.Enabled = True : Cursor = Cursors.Default
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
        If _lastHoveredListItem IsNot Nothing AndAlso e.ItemIndex = _lastHoveredListItem.Index AndAlso Not e.Item.Selected Then
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
        ProgressBar2.Text = $"虛擬排序 {_lv3MailList.Count:N0} 項，耗時 {sw.Elapsed.TotalSeconds:0.00} 秒"
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

        ' by Gemini: Layer2 業務層向 Layer2.5 請求平行預載快取。若 RDO 存在，這行能在極短時間內把後續需要的資料全數載入記憶體。
        If RDO_Parallel1.Checked Then
            Await RdoPreloadAttach_1(sourceList, progress, cToken:=cToken)    ' by Parellel.ForEach 來平行讀取附件資料，適合 CPU 密集型的 MAPI 存取
        ElseIf RDO_Parallel2.Checked Then
            Await RdoPreloadAttach_2(sourceList, progress, cToken:=cToken)    ' by Task.WhenAll 來平行讀取附件資料，適合 I/O 等待型的資料庫存取
        End If
        ' 假設您現在的流程是：「讀取大量郵件屬性 --> 與本地資料庫比對快取 --> 寫入資料庫」
        ' 1. 讀取 MAPI 資料 (CPU + 嚴格 Thread 限制)：使用 Parallel.ForEach。
        '    因為您需要真實在多個核心上建立獨立的 RDOSession 來平行榨取硬碟與 MAPI 引擎的讀取速度
        ' 2. 查詢/寫入本地資料庫快取 (I/O 等待)：使用 Task.WhenAll + async。
        '    如果您的底層資料庫驅動 (例如 SQLite-net 或 Entity Framework) 支援原生的非同步方法
        '    (ToListAsync, ExecuteAsync)

        Dim swTotal As Stopwatch = Stopwatch.StartNew()     ' by Claude Sonnet 4.6, 2026/06/07
        Dim swThrottle As Stopwatch = Stopwatch.StartNew()  ' by Claude Sonnet 4.6, 2026/06/07

        Dim mustCountAttach As Boolean = CheckAttCount.Checked
        Dim minCount As Integer = If(mustCountAttach, CInt(CountMin.Value), 0)
        Dim maxCount As Integer = If(mustCountAttach, CInt(CountMax.Value), Integer.MaxValue)

        Dim processed As Integer = 0, total As Integer = sourceList.Count
        Dim resultList As New List(Of MailItemInfo)(4096)   ' 預分配容量為 4096，優化搜尋結果清單的填充速度 (by Gemini 3 Flash, 2026/05/04)
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
                                                                                        .Message = $"Phase 2 (開始比對郵件): {processed} / {total}，已符合 {resultList.Count} 封 ({eta.Speed:F0} 封/秒{eta.EtaString})"})
                                          End Sub)

                Dim currentMail As MailItemInfo = sourceList(curMail)
                Dim cachedAttFilenames As List(Of String) = GetAttachFilename(currentMail)

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
        ProgressBar1.Text = $"共找到 {lviCount} 封 / 耗時 {elapsedSeconds:0.00} 秒{speedText}"
        ProgressBar2.Text = ""
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
        ProgressBar1.Text = "正在處理..." : ProgressBar2.Text = "開始掃描系列郵件..."
        _tv4PrevSelection = New List(Of Folder)(selectedFolders) ' 記憶最後成功的搜尋目標清單

        Dim sw As Stopwatch = Stopwatch.StartNew()          ' by Claude Sonnet 4.6, 2026/06/07
        Dim swThrottle As Stopwatch = Stopwatch.StartNew()  ' by Gemini, 2026/04/02: 重用秒錶做節流; refactored by Claude Sonnet 4.6, 2026/06/07
        Dim topicDict As New Dictionary(Of String, List(Of MailItemInfo))(StringComparer.OrdinalIgnoreCase)
        Dim progress4 As IProgress(Of ProgressReport) = New Progress(Of ProgressReport)(Sub(p) ProgressBar2.Text = p.Message)

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
            ProgressBar1.Text = $"找到 {SimTree4.Nodes.Count} 個系列 / 耗時 {sw.Elapsed.TotalSeconds:0.00} 秒" : ProgressBar2.Text = ""
        Catch ex As System.Exception
            _dbg("結束", "ESC 中斷")
            ProgressBar1.Text = "由使用者中斷。" : ProgressBar2.Text = ""
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
    'Private Sub Tv4_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles SimTree4.AfterSelect

    '    ' ✅ 2026/04/20 by Gemini 2.0 Flash: 新增雙模式選取邏輯
    '    ' 模式 A: 資料夾模式 (目前的行為是選取後僅供搜尋參考，不執行連動)
    '    _dbg("開始 (A:資料夾模式)", e.Node.Text)
    '    If Not Lv4Topic.Visible Then Return ' by Claude Sonnet 4.6, 2026/05/29: 將 _isTv4ResultMode 替換為 Lv4Topic.Visible

    '    ' 模式 B: 主旨模式 (顯示主旨下的郵件清單)
    '    _dbg("開始 (B:主旨模式)", e.Node.Text)
    '    Dim mailList As List(Of MailItemInfo) = TryCast(e.Node.Tag, List(Of MailItemInfo))
    '    If mailList Is Nothing Then Return

    '    _lv4SortOrder = SortOrder.Descending    ' 每次點選新節點時，重置排序狀態為預設 (日期降冪)
    '    _lv4LastSortColumn = 2                  ' 收到日期所在的 index
    '    mailList.Sort(Function(a, b) b.RcvTime.CompareTo(a.RcvTime))  ' 排序: 依據時間遞減 (越新的在越前面)
    '    RenderLv4Result(mailList)
    '    _dbg("結束", $"顯示 {mailList.Count} 封系列郵件")

    'End Sub
    'Private Sub Tv4_KeyDown(sender As Object, e As KeyEventArgs) Handles SimTree4.KeyDown
    '    ' ✅ 2026/04/20 by Gemini 2.0 Flash: 處理 SimTree4 的快捷鍵與模式切換
    '    _dbg("開始", e.KeyCode.ToString())

    '    Select Case e.KeyCode
    '        Case Keys.Enter
    '            ' 在結果模式下按下 Enter 切換焦點到列表
    '            ' 2026/05/29 by Simon/Claude: 拆分SimTree4的雙重模式, 讓SimTree4回復到純粹的資料夾樹行為
    '            '   這裡的 Enter 只負責開始搜尋 (等同 Button4)，不再處理切換焦點的行為
    '            '   結果模式下的主旨選取改由 Lv4Topic 處理，
    '            ' todo: 其實只剩資料夾模式就可以合併回去原本的共用熱鍵處理函數了，已經沒有雙重模式的需求
    '            'If _isTv4ResultMode AndAlso Listview4.Items.Count > 0 Then Listview4.Focus()
    '            Button4.PerformClick()
    '            e.Handled = True

    '            'Case Keys.F5
    '            '    ' 按下 F5 等同 Button4 (重新開始掃描系列郵件)
    '            '    ' ✅ 2026/04/20: 在結果模式下按 F5 會自動引用上一資料夾重新掃描
    '            '    Button4.PerformClick()
    '            '    e.Handled = True

    '            'Case Keys.F6
    '            '    ' ✅ 2026/04/20 by Gemini 2.0 Flash: 切換左側樹排序方式 (數量/名稱)
    '            '    If _isTv4ResultMode AndAlso _tv4PrevTopicResults IsNot Nothing Then
    '            '        _tv4GroupSortByCount = Not _tv4GroupSortByCount
    '            '        RenderLv4Group(_tv4PrevTopicResults)
    '            '        _dbg("F6 按下：切換排序為", If(_tv4GroupSortByCount, "數量", "主旨"))
    '            '        e.Handled = True
    '            '    End If

    '            'Case Keys.Escape
    '            '    ' 按下 ESC：從結果模式恢復為資料夾模式
    '            '    If _isTv4ResultMode Then
    '            '        _dbg("ESC 按下：恢復資料夾模式 (PopNodeSnapshot)")
    '            '        _isTv4ResultMode = False
    '            '        Listview4.Items.Clear()

    '            '        ' ✅ 2026/05/23 by Simon/Claude: 改用 SimTree 內建快照還原，取代舊版手動重插節點
    '            '        '   RestoreTreeNodeSnap 內部處理：BeginUpdate/EndUpdate、節點插回、選取還原、EnsureVisible
    '            '        '   若插槽不存在（Fallback：重新載入資料夾樹）
    '            '        If Not SimTree4.RestoreTreeNodeSnap("folder-view") Then
    '            '            LoadStoreToTreeView(_pstStoreList, SimTree4)
    '            '            GotoDefaultInbox(SimTree4)
    '            '        End If

    '            '        ProgressBar1.Text = "已恢復資料夾樹模式。" : ProgressBar2.Text = ""
    '            '        SimTree4.Focus()
    '            '        e.Handled = True : e.SuppressKeyPress = True
    '            '    End If
    '    End Select

    'End Sub
    'Private Sub RenderLv4Group(topicDict As Dictionary(Of String, List(Of MailItemInfo)))
    '    ''' <summary>
    '    ''' ✅ 2026/04/20 by Gemini 2.0 Flash: 根據目前的排序模式渲染 Tab4 的主旨群組樹
    '    ''' </summary>

    '    _dbg("開始")
    '    If topicDict Is Nothing Then Return

    '    SimTree4.BeginUpdate()
    '    SimTree4.Nodes.Clear()
    '    ' _isTv4ResultMode = True ' by Claude Sonnet 4.6, 2026/05/29: 已廢棄，改用 Lv4Topic.Visible 代替

    '    _dbg("渲染系列清單", $"模式: {If(_tv4GroupSortByCount, "按數量", "按主旨")}")
    '    ' 根據旗標決定排序方式 (by Gemini 3 Flash, 2026/05/11: 改為 AddRange 模式以提升效能)
    '    Dim nodesArray = If(Not _tv4GroupSortByCount,
    '        topicDict.Where(Function(kvp) kvp.Value.Count > 1).
    '                  OrderBy(Function(kvp) kvp.Key).
    '                  Select(Function(kvp) New TreeNode($"{kvp.Key} ({kvp.Value.Count})") With {.Tag = kvp.Value}).ToArray(),
    '        topicDict.Where(Function(kvp) kvp.Value.Count > 1).
    '                  OrderByDescending(Function(kvp) kvp.Value.Count).
    '                  ThenBy(Function(kvp) kvp.Key).
    '                  Select(Function(kvp) New TreeNode($"{kvp.Key} ({kvp.Value.Count})") With {.Tag = kvp.Value}).ToArray())

    '    If nodesArray.Length > 0 Then SimTree4.Nodes.AddRange(nodesArray)
    '    SimTree4.EndUpdate()

    '    ' ✅ by Gemini 3.0 flash, 2026/04/21: 搜尋完成後，自動選取第一個結果並 Focus
    '    ' 💡 補充: 為了確保右側 Listview4 同步更新，手動呼叫事件處理器 (by Gemini 3.0 flash, 2026/04/21)
    '    If SimTree4.Nodes.Count > 0 Then
    '        Dim firstNode = SimTree4.Nodes(0)
    '        SimTree4.SelectedNode = firstNode
    '        SimTree4.Focus()
    '        Tv4_AfterSelect(SimTree4, New TreeViewEventArgs(firstNode))
    '    End If
    '    ProgressBar1.Text = $"找到 {SimTree4.Nodes.Count} 個系列 (排序: {If(_tv4GroupSortByCount, "數量", "主旨")})"
    '    _dbg("結束")

    'End Sub
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
        ' 2026/05/30 by Gemini/Simon: 優化為虛擬模式，改用 _lv4TopicData 存底層資料，
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
        If idx < 0 OrElse idx >= _lv4TopicData.Count Then Return

        Dim mailList As List(Of MailItemInfo) = _lv4TopicData(idx).Value
        _tv4SelectedTopicMailList = mailList

        ' 重置排序狀態為預設（日期降冪）
        _lv4SortOrder = SortOrder.Descending : _lv4LastSortColumn = 2
        mailList.Sort(Function(a, b) b.RcvTime.CompareTo(a.RcvTime))

        RenderLv4Result(mailList)
        _dbg("結束", $"顯示 {mailList.Count} 封系列郵件")
    End Sub
    Private Sub Lv4Topic_RetrieveVirtualItem(sender As Object, e As RetrieveVirtualItemEventArgs) Handles Lv4Topic.RetrieveVirtualItem
        ' 虛擬模式核心: 當項目進入視野時才動態組裝
        If e.ItemIndex < 0 OrElse e.ItemIndex >= _lv4TopicData.Count Then Return

        Dim kvp = _lv4TopicData(e.ItemIndex)
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

                    ProgressBar1.Text = "" : ProgressBar2.Text = ""
                    e.Handled = True : e.SuppressKeyPress = True
                End If
        End Select

    End Sub
    Private Sub Lv4Topic_ColumnClick(sender As Object, e As ColumnClickEventArgs) Handles Lv4Topic.ColumnClick
        If _lv4TopicData Is Nothing OrElse _lv4TopicData.Count = 0 Then Return
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
                    _lv4TopicData = If(_lv4TSortOrder = SortOrder.Ascending,
                                       _lv4TopicData.OrderBy(Function(x) x.Key).ToList(),
                                       _lv4TopicData.OrderByDescending(Function(x) x.Key).ToList())
                Case 1 ' 數量 (Value.Count)
                    _lv4TopicData = If(_lv4TSortOrder = SortOrder.Ascending,
                                       _lv4TopicData.OrderBy(Function(x) x.Value.Count).ThenBy(Function(x) x.Key).ToList(),
                                       _lv4TopicData.OrderByDescending(Function(x) x.Value.Count).ThenBy(Function(x) x.Key).ToList())
            End Select

            ' 3. 排序後，原本的選取索引會指錯資料，建議清空選取並重新選定第一筆
            Lv4Topic.SelectedIndices.Clear()
            If _lv4TopicData.Count > 0 Then Lv4Topic.SelectedIndices.Add(0)

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
        If e.Item Is _lastHoveredListItem AndAlso Not e.Item.Selected Then
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
        ' 2026/05/30 by Gemini/Simon: 優化為虛擬模式，改用 _lv4TopicData 存底層資料，RenderLv4Topic 只負責排序與刷新 ListView 的 VirtualListSize，實際項目由 RetrieveVirtualItem 動態組裝
        ' 2026/05/31 by Gemini/Simon: 徹底大掃除：清除所有 F6 與 Group 的殘留
        ' ---------------------------------------------------------------
        _dbg("開始")
        If topicDict Is Nothing Then Return

        ' ── 排序：拿掉 IF 判斷，直接預設：數量多的排前面，數量相同則按主旨排 ──
        Dim sortedItems = topicDict.Where(Function(kvp) kvp.Value.Count > 1).
                                    OrderByDescending(Function(kvp) kvp.Value.Count).ThenBy(Function(kvp) kvp.Key)

        ' ── 排序後存入 _lv4TopicData (取代實體 ListViewItem) ── 
        _lv4TopicData = sortedItems.ToList()    ' 2026/5/30 by Gemini/Simon

        Lv4Topic.BeginUpdate()
        Lv4Topic.VirtualListSize = 0    ' 先歸零強制重置
        Lv4Topic.VirtualListSize = _lv4TopicData.Count
        Lv4Topic.EndUpdate()

        ' ── 切換顯示：Lv4Topic 上台，SimTree4 退後 ──
        SimTree4.Visible = False : Lv4Topic.Visible = True : Lv4Topic.Focus()

        ' (虛擬模式改用 SelectedIndices) ── ' 2026/5/30 by Gemini/Simon
        If _lv4TopicData.Count > 0 Then Lv4Topic.SelectedIndices.Add(0)

        ' 更新 ProgressBar，移除了排序狀態文字
        ProgressBar1.Text = $"找到 {_lv4TopicData.Count} 個系列"
        _dbg("結束", $"{_lv4TopicData.Count} 個系列")
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
        ProgressBar2.Text = $"系列選中：共 {mailList.Count:N0} 封郵件"
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
                ProgressBar2.Text = $"已移動 {selCount} 封郵件至刪除郵件資料夾"
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
        Dim results(mBodyList.Count - 1) As Double
        Await Task.Run(Sub()
                           Parallel.For(0, mBodyList.Count, Sub(i)
                                                                If token.IsCancellationRequested Then Return
                                                                Dim targetBody = mBodyList(i).TargetBody
                                                                'results(i) = CalculateSimilarity(baseBody, targetBody)
                                                                'results(i) = CharlesHash(baseBody, targetBody)
                                                                results(i) = JaccardSimilarity(baseBody, targetBody)
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
            ProgressBar1.Text = $"已重新讀取 {stats.Updated} 封 (失效 {stats.NotFound}, 錯誤 {stats.Errored})。" : ProgressBar2.Text = ""
        Catch ex As OperationCanceledException
            ProgressBar1.Text = "刷新已取消。"
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
    Private Function GetHammingDistance(hash1 As Long, hash2 As Long) As Integer
        ' 利用位元互斥或 (XOR) 找出不同的 bit，再計算有幾個 1 (PopCount)
        Return System.Numerics.BitOperations.PopCount(CULng(hash1 Xor hash2))
    End Function
#End Region
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
        ProgressBar1.Text = "正在準備" : ProgressBar2.Text = "展開資料夾結構..."
        Dim sw As Stopwatch = Stopwatch.StartNew()  ' by Claude Sonnet 4.6, 2026/06/07
        Dim progress5 As IProgress(Of ProgressReport) = New Progress(Of ProgressReport)(Sub(p) ProgressBar2.Text = p.Message)

        Try
            Dim folderList = Await GetUniqueFolderList(selectedNodes, includeSub:=includeSub, cToken:=cToken, progress:=progress5)
            If folderList.Count = 0 Then Return

            Dim groupDict = Await ScanMailsToGroupDictAsync(folderList, isExactMode, progress5, cToken)
            _lv5PrevGroupResults = groupDict ' by Gemini 3 Flash, 2026/05/06: 儲存結果以供動態刪除
            _tv5PrevSearchMode = isExactMode
            Dim counts = RenderLv5Group(groupDict, isExactMode)

            sw.Stop()
            ProgressBar1.Text = $"找到 {counts.GroupCount} 組 ({counts.MailCount} 封) / 耗時 {sw.Elapsed.TotalSeconds:0.00} 秒"
            ProgressBar2.Text = ""
        Catch ex As OperationCanceledException
            _dbg("結束", "ESC 中斷") : ProgressBar1.Text = "由使用者中斷。" : ProgressBar2.Text = ""
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
                                Sub()
                                    ' 新版 (2026/05/10 by Simon/Claude: 加入 ETA 顯示，對齊 Tab3 做法)
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
        ListView5.Tag = groupDict ' by Gemini 3 Flash, 2026/05/06: 將資料來源掛載至 Tag 供 HandleLv5Delete 使用
        ListView5.BeginUpdate()
        ListView5.Items.Clear()

        Dim groupID As Integer = 1 : Dim totalMails As Integer = 0
        Dim ascending As Boolean = (_lv5SortOrder = SortOrder.Ascending)

        ' ── Step 1: Jaccard 過濾並計算 simScores，產生有效群組清單 ──────
        ' 先過濾出有效群組 (含 simScores)，再排序，避免排序後 index 與 simScores 錯位
        Dim validGroupList As New List(Of (Key As String, Items As List(Of MailItemInfo), Scores As List(Of Double)))(groupDict.Count)
        For Each kvp In groupDict
            If kvp.Value.Count <= 1 Then Continue For

            Dim simScores As New List(Of Double)(kvp.Value.Count)
            Dim isValidGroup As Boolean = True
            simScores.Add(1.0) ' 第一封基準

            If isExact Then
                ' 2026/05/10 by Simon/Claude: Exact 模式不做 Jaccard，全部填 100%
                For i As Integer = 1 To kvp.Value.Count - 1 : simScores.Add(1.0) : Next
            Else
                ' Fuzzy 模式：Jaccard 僅用於過濾，順帶記錄供顯示
                Dim firstSubject As String = kvp.Value(0).Subject
                For i As Integer = 1 To kvp.Value.Count - 1
                    Dim sim As Double = JaccardSimilarity(firstSubject, kvp.Value(i).Subject)
                    simScores.Add(sim)
                    ' 僅在模糊模式下才套用門檻過濾 (0.6)
                    If sim < 0.6 Then isValidGroup = False : Exit For
                Next
            End If

            ' ── [預留] SimHash 內文比對將在此插入 ──
            ' If isValidGroup Then isValidGroup = SimHashBodyFilter(kvp.Value)

            If Not isValidGroup Then Continue For

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
        ''' </summary>
        _dbg("開始")
        Dim selCount As Integer = lv.SelectedItems.Count
        If selCount = 0 Then Return

        If MessageBox.Show($"確定要將選中的 {selCount} 封郵件移到「刪除郵件」資料夾嗎？", "確認刪除",
                           MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then

            ' 取得快取的資料字典
            Dim groupDict As Dictionary(Of String, List(Of MailItemInfo)) = _lv5PrevGroupResults
            If groupDict Is Nothing Then Return

            ' 2026/5/11 simon: 收集受影響的資料夾路徑，供後續清理快取與DB使用
            Dim affectedPaths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            ' 收集選中項目的 EntryID 並從資料源中移除
            Dim entryIDs As New List(Of String)(selCount)
            For Each item As ListViewItem In lv.SelectedItems
                If TypeOf item.Tag Is MailItemInfo Then
                    Dim info = DirectCast(item.Tag, MailItemInfo)
                    entryIDs.Add(info.EntryID)

                    ' 從 groupDict 中移除該封信 (遍歷所有群組尋找)
                    For Each kvp In groupDict
                        ' 找到並移除後，如果該群組只剩 1 封或 0 封，在重複郵件邏輯中視為不再重複，可選擇保留或由渲染器過濾
                        If kvp.Value.RemoveAll(Function(m) m.EntryID = info.EntryID) > 0 Then Exit For
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
                RenderLv5Group(groupDict, _tv5PrevSearchMode) ' 重新渲染 UI
                ProgressBar2.Text = $"已移動 {selCount} 封郵件至刪除郵件資料夾"
            End If
        End If
        _dbg("結束")
    End Sub
#End Region
#Region "  ├ 共用事件函數"
    ' by Gemini 3.1 Pro, 2026/04/21: 邏輯整合 (Tab3/Tab4/Tab5)，完整統一行為。
    ' 理由: Tab3 與 Tab4 的 ListView 皆為「搜尋結果」，行為高度一致 (Enter/雙擊/連動與路徑顯示)。
    ' 整合後可減少冗餘代碼，並確保滑鼠與熱鍵行為絕對一致。
    ' --------------------------------------------------------------
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
                Await RefreshLviAllItems(lv)         ' Lv3/Lv5 全體刷新 (依數量自動 A/B)
            ElseIf lv Is Listview4 Then
                Await RefreshLv4Result(lv)      ' Lv4 沿用：重讀目前系列郵件
            End If
        ElseIf e.Control AndAlso e.KeyCode = Keys.A Then
            LviSelectAll(lv, e)
        End If
    End Sub
    Private Sub HandleLv3Lv4Lv5_MouseClick(sender As Object, e As MouseEventArgs)
        ''' <summary>
        ''' 共通滑鼠點擊: 複製主旨與路徑預覽
        ''' </summary>
        Dim lv = DirectCast(sender, ListView)
        Dim item As ListViewItem = lv.GetItemAt(e.X, e.Y)

        If item IsNot Nothing AndAlso e.Button = MouseButtons.Left Then
            ' 單擊左鍵複製主旨到剪貼簿，這原本是 Listview4 獨有的方便設計，現在擴展到 Tab3 共用 (by Gemini 3.1 Pro, 2026/04/21)
            Clipboard.SetText(item.SubItems(0).Text)
        End If
        ' 路徑更新邏輯統一由 ShowLv3Lv4Lv5PathToProgressBar 接管
        ShowLv3Lv4Lv5PathToProgressBar(sender, e)
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
    Private Sub HandleLv3Lv4Lv5_DoubleClick(sender As Object, e As EventArgs)
        ''' <summary>
        ''' 共通雙擊開啟
        ''' </summary>
        OpenMailByEntryID(GetSelectedEntryIDs(DirectCast(sender, ListView)))
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
        Dim foreColor As Color = If(isRefreshed, Color.Blue, SystemColors.WindowText)

        If _lastHoveredListItem IsNot Nothing AndAlso e.ItemIndex = _lastHoveredListItem.Index AndAlso Not e.Item.Selected Then
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
    Private Sub ShowLv3Lv4Lv5PathToProgressBar(sender As Object, e As EventArgs)
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

        If Not String.IsNullOrEmpty(path) Then ProgressBar2.Text = path
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
    Private Sub InitLv3Lv4Lv5RefreshMenu()
        ' 2026/06/15 by Simon/Claude Opus 4.8: Lv3/4/5 共用右鍵選單；冪等，重複呼叫只建一次 (確保三個 LV 共用同一實例)
        If ctxMenuRefresh IsNot Nothing Then Return

        Dim mnuItem As New ToolStripMenuItem("重新讀取選取的郵件(&R)")
        ctxMenuRefresh = New ContextMenuStrip()
        ctxMenuRefresh.Items.Add(mnuItem)

        ' 2026/06/14 by Simon/Claude Opus 4.8: 右鍵選單「強制刷新選取的郵件」點擊
        AddHandler mnuItem.Click, Async Sub(sender, e)
                                      Dim lv = TryCast(ctxMenuRefresh.SourceControl, ListView)
                                      If lv IsNot Nothing Then Await RefreshLviSelected(lv)
                                  End Sub

        ' 沒有選取項就不顯示選單 (虛擬模式用 SelectedIndices，實體模式用 SelectedItems)
        AddHandler ctxMenuRefresh.Opening, Sub(s, ev)
                                               Dim lv = TryCast(ctxMenuRefresh.SourceControl, ListView)
                                               Dim cnt = If(lv Is Nothing, 0, If(lv.VirtualMode, lv.SelectedIndices.Count, lv.SelectedItems.Count))
                                               If cnt = 0 Then ev.Cancel = True
                                           End Sub
    End Sub
    Private Async Function RefreshLviSelected(lv As ListView) As Task
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
            ProgressBar1.Text = $"已刷新選取的 {stats.Updated} 封 (失效 {stats.NotFound}, 錯誤 {stats.Errored})。" : ProgressBar2.Text = ""
        Catch ex As OperationCanceledException
            ProgressBar1.Text = "刷新已取消。"
        Catch ex As System.Exception
            _dbg("選取刷新錯誤", ex.Message)
        Finally
            _isUserBusy = False : Cursor = Cursors.Default
            _dbg("結束")
        End Try
    End Function
    Private Async Function RefreshLviAllItems(lv As ListView) As Task
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
            ProgressBar1.Text = $"已刷新 {stats.Updated} 封 (失效 {stats.NotFound}, 錯誤 {stats.Errored})。" : ProgressBar2.Text = ""
        Catch ex As OperationCanceledException
            ProgressBar1.Text = "刷新已取消。"
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
                Await SmartThrottle(swThrottle, ct, ThrottleFreq.Hii, Sub() ProgressBar2.Text = $"刷新郵件 (逐封): {i + 1} / {total}...")
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
                        Await SmartThrottle(swThrottle, ct, ThrottleFreq.Hii, Sub() ProgressBar2.Text = $"刷新郵件 (退回逐封): {done} / {total}...")
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
                Await SmartThrottle(swThrottle, ct, ThrottleFreq.Hii, Sub() ProgressBar2.Text = $"刷新郵件 (批次): {done} / {total}...")
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
#End Region
#Region "  └ 輔助函數"
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

#Region "■ 09 Tab6: Debug & 設定"
    Private Async Sub SaveCache_Click(sender As Object, e As EventArgs) Handles SaveCache.Click
        Await SaveCachesToDB()
        RefreshLv6DbStats()
    End Sub
    Private Async Sub LoadCache_Click(sender As Object, e As EventArgs) Handles LoadCache.Click
        Await LoadCachesFromDB()
        Dim st = GetDBSummary()
        ProgressBar2.Text = $"DB 統計 — folder_stats:{st.fc} 筆 / attach_maillist:{st.mb} 筆 / attach_filenames:{st.at} 筆 / year_counts:{st.yc} 筆 / month_counts:{st.mc} 筆 / {st.kb} KB"

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
                Case DialogResult.Yes ' 僅記憶體
                    ClearMemoryCachesCore()
                    ProgressBar2.Text = "已完成：僅清除記憶體快取 (SSD 保留)"
                    _dbg("清理", "僅記憶體")

                Case DialogResult.No ' 僅 SSD
                    If MessageBox.Show("【安全提示】這將把目前的 SSD 快取檔更名備份 ( .zip) 並重新建立空白資料表。" & vbCrLf & "這可以解決 Schema 不相容問題且具備救援機制，確定嗎？", "重置 SSD 快取", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = DialogResult.OK Then
                        Await ZipAndRebuildDB()
                        ProgressBar2.Text = "已完成：SSD 資料庫已備份並重新初始化"
                        _dbg("清理", "僅 SSD (已備份)")
                    End If

                Case DialogResult.Retry ' 兩者皆清
                    If MessageBox.Show("確定要清除記憶體並備份重置 SSD 快取嗎？", "最後確認", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) = DialogResult.OK Then
                        ClearMemoryCachesCore()
                        Await ZipAndRebuildDB()
                        ProgressBar2.Text = "已完成：記憶體與 SSD 快取已全數歸零 (舊 SSD 檔已備份)"
                        _dbg("清理", "FULL CLEAN (已備份)")
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
            Await RenewCacheToDB(RenewIncludeSize.Checked)
            Await DbVacuumIfNeeded()    ' 2026/06/16 by Claude Sonnet 4.6: RenewCache 完成後，視碎片比例決定是否執行 VACUUM (freelist_count / page_count > 5% 才執行，避免每次都白等)
            RefreshLv6DbStats()
            Await RefreshAllTreeViews() ' by Gemini 3.0 flash, 2026/04/24: 更新完成後，執行非同步 UI 刷新，確保新資料夾能立即顯示

        Catch ex As OperationCanceledException
            _dbg(" ├ 中斷", "使用者已取消快取更新")
        Finally
            ProgressBar1.Text = $"RenewCache 完成 — 耗時: {sw.Elapsed.TotalSeconds:0.000} 秒"
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
            AddLv6StatLine("_cacheSubTreeList", _cacheSubTreeList.Count.ToString("N0") & " 筆")
            AddLv6StatLine("_cacheIsMailFolder", _cacheIsMailFolder.Count.ToString("N0") & " 筆")
            AddLv6StatLine("_cacheFolderIDs", _cacheFolderIDs.Count.ToString("N0") & " 筆")
            AddLv6StatLine("", "", isHeader:=False) ' 間隔
            AddLv6StatLine("_cacheMailCount", _cacheMailCount.Count.ToString("N0") & " 筆")
            AddLv6StatLine("_cacheMailCountAll", _cacheMailCountAll.Count.ToString("N0") & " 筆")
            AddLv6StatLine("_cacheFolderCount", _cacheFolderCount.Count.ToString("N0") & " 筆")
            AddLv6StatLine("_cacheFolderCountAll", _cacheFolderCountAll.Count.ToString("N0") & " 筆")
            AddLv6StatLine("_cacheFolderSize", _cacheFolderSize.Count.ToString("N0") & " 筆")
            AddLv6StatLine("_cacheFolderSizeAll", _cacheFolderSizeAll.Count.ToString("N0") & " 筆")
            AddLv6StatLine("_cacheYearCounts", _cacheYearCounts.Count.ToString("N0") & " 筆")
            AddLv6StatLine("_cacheMonthCounts", _cacheMonthCounts.Count.ToString("N0") & " 筆")
            AddLv6StatLine("_cacheAttachMailList", _cacheAttachMailList.Count.ToString("N0") & " 筆")
            AddLv6StatLine("_cacheAttachFilename", _cacheAttachFilename.Count.ToString("N0") & " 筆")
            AddLv6StatLine("", "", isHeader:=False) ' 間隔

            ' ── 步驟 4: 填充 SQLite 數據 ──
            ' 拆分日期與時間 (壓縮寫法)
            Dim parts = st.lastTs.Split(" "c)
            Dim datePart = If(st.lastTs.Contains(" "c), parts(0), st.lastTs)
            Dim timePart = If(st.lastTs.Contains(" "c), parts(1), "N/A")

            AddLv6StatLine("════ SQLite 快取 ════", "", True)
            AddLv6StatLine("DB 檔案大小", (st.kb / 1024).ToString("F2") & " MB")
            AddLv6StatLine("folder_stats", st.fc.ToString("N0") & " 筆")
            AddLv6StatLine("senders", st.senders.ToString("N0") & " 筆")         ' 2026/06/14 by Simon/Claude Opus 4.8: 補上 senders，與 DbShowDbFileStat 順序一致
            AddLv6StatLine("basic_maillist", st.basic.ToString("N0") & " 筆")    ' by Gemini 3 Flash, 2026/04/22
            AddLv6StatLine("year_counts", st.yc.ToString("N0") & " 筆")
            AddLv6StatLine("month_counts", st.mc.ToString("N0") & " 筆")
            AddLv6StatLine("attach_maillist", st.mb.ToString("N0") & " 筆")
            AddLv6StatLine("attach_filenames", st.at.ToString("N0") & " 筆")
            AddLv6StatLine("最後更新日期", datePart)
            AddLv6StatLine("最後更新時間", timePart)

            ' ── 步驟 5: 填充 ZIP 備份數據 ── (2026/06/01: added by Claude, 6/2: 抽離函式 by Gemini)
            Dim zipStats = GetFileStats(_dbPath, "*.zip")
            AddLv6StatLine($"備份 ZIP 檔總計 ({zipStats.Count}個)", $"{zipStats.TotalMB:N0} MB")

        Catch ex As System.Exception
            ListView6.Items.Clear()
            ListView6.Items.Add(New ListViewItem("❌ 讀取統計失敗: " & ex.Message))
        Finally
            ListView6.EndUpdate()
            ProgressBar1.Text = "已更新Cache / SQL DB 統計資料。"
            _dbg("結束")
        End Try
    End Sub

    Private Sub ListView6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListView6.SelectedIndexChanged
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
            '    由於 ListView6 包含 "DB 檔案大小" 或 "備份 ZIP" 等非實體資料表項目，
            '    我們透過 Contains 進行模糊識別，確保精準抓出使用者想看的是哪一張快取表。
            Dim targetTableName As String = ""
            If selectedLabel.Contains("basic_maillist") Then
                targetTableName = "basic_maillist"
            ElseIf selectedLabel.Contains("attach_maillist") Then
                targetTableName = "attach_maillist"
            ElseIf selectedLabel.Contains("attach_filenames") Then
                targetTableName = "attach_filenames"
            ElseIf selectedLabel.Contains("month_counts") Then  ' 2026/06/12 by Simon/Claude Opus 4.8: 修正 typo (month_stats → month_counts)
                targetTableName = "month_counts"
            ElseIf selectedLabel.Contains("year_counts") Then   ' 2026/06/12 by Simon/Claude Opus 4.8: 補上缺漏的分支
                targetTableName = "year_counts"
            ElseIf selectedLabel.Contains("folder_stats") Then  ' 2026/06/12 by Simon/Claude Opus 4.8: 補上缺漏的分支
                targetTableName = "folder_stats"
            ElseIf selectedLabel = "DB 檔案大小" Then           ' 2026/06/13 by Simon/Claude Opus 4.8: 新增對 DB 檔案大小 的特殊識別，觸發專門的空間分布分析
                Dim unused = DbShowDbFileStat()                 ' 明確的 fire-and-forget，編譯器知道你是故意的
                ' 2026/06/13 by Simon/Claude Opus 4.8: 未來要加上例外續集也可以一行搞定，確保即使該功能發生錯誤也不會影響 UI 穩定性，並將錯誤訊息導向除錯視窗
                ' Me.DbShowDbFileStat().ContinueWith(
                '    Sub(t) _dbg(" ├ 錯誤", $"DbShowDbFileStat task faulted: {t.Exception?.Message}"), TaskContinuationOptions.OnlyOnFaulted Or TaskContinuationOptions.ExecuteSynchronously)
            End If

            ' 4. 根據比對結果執行對應的 Debug 輸出
            If Not String.IsNullOrEmpty(targetTableName) Then
                ' 【快取資料表分支】
                ' 呼叫您在 Form1_SQLite2.vb 中實作好的深度空間診斷函數
                ' 提示：因為您的 SQLite 持久層同屬 Form1 的 Partial Class，此處可直接利用 Me 呼叫
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
    Private Sub ListView6_DoubleClick(sender As Object, e As EventArgs) Handles ListView6.DoubleClick
        If ListView6.SelectedItems.Count = 0 Then Return

        ' 取得被雙擊的項目文字
        Dim clickedText = ListView6.SelectedItems(0).Text

        ' 判斷是否為需要開啟資料夾的特定項目
        If clickedText = "DB 檔案大小" OrElse clickedText.StartsWith("備份 ZIP") Then
            Dim dbDir = IO.Path.GetDirectoryName(_dbPath)
            If IO.Directory.Exists(dbDir) Then Process.Start("explorer.exe", dbDir)
        End If
    End Sub
    Private Sub ListView6_KeyDown(sender As Object, e As KeyEventArgs) Handles ListView6.KeyDown
        ' F5 強制刷新
        If e.KeyCode = Keys.F5 Then RefreshLv6DbStats()
    End Sub
    Private Sub DebugButton_Click(sender As Object, e As EventArgs) Handles DebugButton.Click

        ' 測試 DASL 是否能在 GetTable 直接濾出含有特定附檔名的信件
        Dim folder As Folder = TryCast(SimTree3.SelectedNode.Tag, Folder)
        If folder Is Nothing Then MessageBox.Show("請先選擇資料夾") : Return
        Dim keyword As String = "2025" ' 測試關鍵字

        ' 寫法 A: 使用 LIKE (不支援索引的情況)
        Dim filterLike As String = $"@SQL=""urn:schemas:httpmail:attachmentfilename"" LIKE '%{keyword}%'"

        ' 寫法 B: 使用 CI_PHRASEMATCH (依賴 Windows Search 索引，速度極快)
        Dim filterCI As String = $"@SQL=""urn:schemas:httpmail:attachmentfilename"" CI_PHRASEMATCH '{keyword}'"

        ' 這裡您可切換 filterLike 或 filterCI 測試
        Dim table As Outlook.Table = Nothing
        Try
            table = folder.GetTable(filterLike)
            MessageBox.Show($"測試成功！GetTable 直接過濾出 {table.GetRowCount()} 筆包含 {keyword} 的郵件。")

            ' 印出前幾筆的主旨驗證
            table.Columns.RemoveAll()
            table.Columns.Add("Subject")
            Dim count As Integer = 0
            While Not table.EndOfTable AndAlso count < 5
                Dim row As Outlook.Row = table.GetNextRow()
                _dbg($"郵件: {row("Subject")}")
                count += 1
            End While
        Catch ex As System.Exception
            MessageBox.Show($"DASL 過濾失敗: {ex.Message}")
        Finally
            If table IsNot Nothing Then Marshal.ReleaseComObject(table)
        End Try

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
            SyncDebugFormPosition()
            If Not DebugForm.Visible Then DebugForm.Show(Me) ' 2026/3/27 by Gemini: 設定 Owner 確保點選 Form1 時 DebugForm 一起回到前面
        Else
            DebugForm.Hide()
        End If
        _dbg("結束")

    End Sub

    ' 輔助函式 Helper Functions
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

End Class
