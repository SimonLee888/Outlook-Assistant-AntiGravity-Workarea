Imports System.Collections.Concurrent
Imports System.Runtime.InteropServices
Imports System.Threading
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Outlook

Partial Class Form1

#Region "■ 01 全域宣告"
    Private pnlOptions_tab3 As Panel
    Private lv3SortOrder As SortOrder = SortOrder.Ascending     ' 設置初始排序方式為升序
    Private lv3LastSortColumn As Integer = -1                   ' 儲存上一次點選的列索引
    Private _lv3MailList As New List(Of MailItemInfo)(4096)     ' by Gemini, 2026/04/10: Tab3 顯示資料庫 (虛擬模式核心)    ' 預分配容量為 4096，因應 Tab3 可能載入的大量郵件資訊，顯著降低記憶體配置開銷 (by Gemini 3 Flash, 2026/05/04)
    ' Private _isTab3_Stop As Boolean                           ' 2026/04/05 by Gemini: 已併入全域 _cancelRequested，不再單獨使用專屬旗標以簡化邏輯內容流程處理機制

    Private _currentTabIdx As Integer = 0
    Private _isTab4ShowingResults As Boolean = False                    ' ✅ 2026/04/20 by Gemini 2.0 Flash: 標記 Tab4 左側樹目前顯示的是搜尋結果模式
    Private _tab4SortGroupsByCount As Boolean = True                    ' ✅ 2026/04/20 by Gemini 2.0 Flash: 記錄排序方式 (True=數量, False=主旨)
    Private _tab4LastSearchFolders As New List(Of Outlook.Folder)(32)   ' ✅ 2026/04/21 by Gemini 3.0 flash: 記憶最後一次搜尋的多個資料夾  ' 預分配容量為 32，足以涵蓋多數搜尋路徑結構，減少陣列頻繁 Resize 開銷 (by Gemini 3 Flash, 2026/05/04)
    Private _tab4LastTopicResults As Dictionary(Of String, List(Of MailItemInfo)) = Nothing ' ✅ 2026/04/20 by Gemini 2.0 Flash: 記憶搜尋結果，供 F6 操作使用
    Private _tab4FolderTreeNodesBackup As New List(Of TreeNode)(64)     ' ✅ 2026/04/21 by Gemini 3.0 flash: 記憶資料夾模式下的節點狀態 (含展開狀態)   ' 預分配容量為 64，以減少大資料量下 UI 節點備份的 Resize 開銷 (by Gemini 3 Flash, 2026/05/04)
    Private _tab4LastClickedFolderNode As TreeNode = Nothing            ' ✅ 2026/04/21 by Gemini 3.0 flash: 記憶進入結果模式前的最後一個選中節點

    Private _lv4SortOrder As SortOrder = SortOrder.Ascending            ' by Gemini 3 Flash, 2026/04/19: 加入 ListView4 專屬排序狀態 (避免與 LV3 共用變數互相干擾)
    Private _lv4LastSortColumn As Integer = -1                          ' by Gemini 3 Flash, 2026/04/19: 加入 ListView4 專屬排序狀態 (避免與 LV3 共用變數互相干擾)
    Private _lv4LastHoverItem As ListViewItem = Nothing                 ' by Gemini 3 Flash, 2026/04/19: 自訂 ListView4 ToolTip 延遲顯示邏輯
    Private _lv4GroupSortByCount As Boolean = False                     ' by Gemini 3 Flash, 2026/04/20: 記錄 Tab4 ListView4 分組排序模式 (False:按主旨, True:按數量)
    Private _lv4BodyCache As New ConcurrentDictionary(Of String, String) ' by Gemini 3 Flash, 2026/04/26: Tab4 相似度計算用的 Body 快取 (session 級，避免重複讀取 Outlook mailitem.Body)
    Private _lv4SimCts As CancellationTokenSource = Nothing             ' by Gemini 3 Flash, 2026/04/26: 用於游標快速移動時, 取消前一次未完成的相似度計算任務

    ' ── Tab5 SimTree 多選控制項 (2026/05/01 by Claude: 取代舊版 TreeView5，對齊 Tab1~4 操作行為) ──
    ' Private WithEvents SimTree5 As SimTree = Nothing
    Private _includeSubTab5 As Boolean = True       ' Tab5 是否含子資料夾，由 CheckSubFolder5 CheckBox 控制，預設 True
    Private rbExactMatch As New RadioButton()       ' tab5 用到的radio button
    Private rbFuzzyMatch As New RadioButton()       ' tab5 用到的radio button
    Private CheckSubFolder5 As New CheckBox()       ' tab5 是否包含子資料夾 (暫用程式建立，日後改設計工具)
    Private _tab5LastGroupResults As Dictionary(Of String, List(Of MailItemInfo)) = Nothing ' by Gemini 3 Flash, 2026/05/06: 記憶 Tab5 掃描結果，供刪除後重新渲染使用
    Private _tab5LastIsExact As Boolean = True      ' by Gemini 3 Flash, 2026/05/06: 記憶最後一次掃描的模式

    Private _lvStats As ListView = Nothing          ' by Gemini 3 Flash, 2026/04/20: 動態建立的統計列表
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
        Dim sw As New Stopwatch : sw.Start()
        Dim swThrottle3 As New Stopwatch : swThrottle3.Start()  ' by Claude, 2026/04/11

        ' by Gemini, 2026/04/09: 讀取 SimTree3.SelectedNodes 集合以支援多選 (取代原單一節點)
        Dim selectedNodes As List(Of TreeNode) = SimTree3.SelectedNodes
        If selectedNodes Is Nothing OrElse selectedNodes.Count = 0 Then Return

        ' ── 鎖定 UI ──
        ProgressBar1.Text = "準備中" : ProgressBar2.Text = "" : Cursor = Cursors.WaitCursor
        pnlOptions_tab3.Enabled = False : SimTree3.Enabled = False
        ListView3.VirtualMode = True        ' by Gemini, 2026/04/10: 解決 ListView 萬筆資料 Clear() 造成 UI 卡頓 1.8 秒的效能瓶頸
        ListView3.VirtualListSize = 0       ' 切換至 VirtualMode 並清空 Size，不銷毀實體物件，速度為 0ms
        _lv3MailList.Clear() : Await Task.Yield ' 確保 UI 先更新狀態再進行後續耗時操作

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
            For i As Integer = 0 To folderList.Count - 1
                ' 2026/04/16 by Gemini: 指定使用 Tuple 內的 .Folder 與 .fPath
                Dim c As Integer = GetMailCount(folderList(i).folder, fPaths(i))    ' 從 400ms 降至近乎 0ms!
                If c > 0 Then totalMailCount += c
                Await SmartThrottle(swThrottle3, cToken:=cToken, ThrottleFreq.Hii) ' 2026/04/16 by Simon/Claude: 改用 ThrottleFreq.Hii + SmartThrottle
            Next
            Dim tStep2_MailCountLoop = swStep.Elapsed.TotalMilliseconds : swStep.Restart() ' by Gemini 3.0 flash, 2026/04/16: 改名以區分 (GetMailCount Loop)

            ProgressBar1.Text = "正在讀取..."
            ProgressBar2.Text = $"準備掃描 {folderList.Count:N0} 個資料夾 (相依包含共計 {totalMailCount:N0} 封信)..."
            Await Task.Yield()

            ' ── Step 3: 收集含附件的郵件清單 (透過 Layer2.5 快取) ──
            Dim progressPhase1 = New Progress(Of ProgressReport)(Sub(p) ProgressBar2.Text = p.Message)
            ' 預分配容量為 4096，顯著降低掃描大量郵件時的記憶體配置開銷 (by Gemini 3 Flash, 2026/05/04)
            Dim targetMails As New List(Of MailItemInfo)(4096)
            Try
                For i As Integer = 0 To folderList.Count - 1
                    ' 2026/04/16 by Gemini: 使用 Tuple 中的 .Folder 與預錄好的 fPaths(i)
                    Dim folderResult = Await GetAttachMailList(folderList(i).folder, progressPhase1, fPaths(i), cToken:=cToken)
                    targetMails.AddRange(folderResult)
                    Await SmartThrottle(swThrottle3, cToken:=cToken, ThrottleFreq.Hii) ' 2026/04/16 by Simon/Claude: 改用 ThrottleFreq.Hii + SmartThrottle
                Next
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
        lvi.SubItems.Add(mail.ReceivedTime.ToString("yyyy/MM/dd HH:mm:ss"))
        lvi.SubItems.Add(mail.SenderName)
        lvi.SubItems.Add(displayName)
        lvi.SubItems.Add(mail.EntryID)

        e.Item = lvi
    End Sub
    Private Sub Lv3_ColumnClick(sender As Object, e As ColumnClickEventArgs) Handles ListView3.ColumnClick
        ' ==============================================================
        ' by Gemini, 2026/04/10: 效能大躍進 — 虛擬模式排序
        ' 理由: 
        '   當資料量達到 5 萬筆時，原有的 ListViewItemComparer 必須不斷從實體 Item 中取值，每秒只能比較約數千筆，會導致 UI 卡頓
        '   切換至 VirtualMode 後，我們直接在記憶體中對 _lv3MailList (List(Of MailItemInfo)) 進行 LINQ 排序，處理 5 萬筆數據僅需 10-30ms，達成「瞬發」排序的效果
        ' --------------------------------------------------------------
        _dbg("開始", "虛擬列表排序") ' by Gemini, 2026/04/10: Level 0
        Dim sw As New Stopwatch : sw.Start()

        ' 判斷是否點選的是同一個列標題, 如果是，則切換排序方式, 否則預設使用升序排序
        lv3SortOrder = If(e.Column = lv3LastSortColumn AndAlso lv3SortOrder = SortOrder.Ascending, SortOrder.Descending, SortOrder.Ascending)
        lv3LastSortColumn = e.Column  ' 儲存目前點選的列索引

        ' by Gemini 2026/4/10, Listview3改virtual mode
        'ListView3.ListViewItemSorter = New ListViewItemComparer(e.Column, lv3SortOrder)    
        ListView3.BeginUpdate()
        Try
            ' 直接對底層資料源進行排序，不操作 UI 物件
            Select Case e.Column
                Case 0 ' 主旨
                    If lv3SortOrder = SortOrder.Ascending Then
                        _lv3MailList = _lv3MailList.OrderBy(Function(x) x.Subject).ToList()
                    Else
                        _lv3MailList = _lv3MailList.OrderByDescending(Function(x) x.Subject).ToList()
                    End If
                Case 1 ' 大小
                    If lv3SortOrder = SortOrder.Ascending Then
                        _lv3MailList = _lv3MailList.OrderBy(Function(x) x.Size).ToList()
                    Else
                        _lv3MailList = _lv3MailList.OrderByDescending(Function(x) x.Size).ToList()
                    End If
                Case 2 ' 時間
                    If lv3SortOrder = SortOrder.Ascending Then
                        _lv3MailList = _lv3MailList.OrderBy(Function(x) x.ReceivedTime).ToList()
                    Else
                        _lv3MailList = _lv3MailList.OrderByDescending(Function(x) x.ReceivedTime).ToList()
                    End If
                Case 3 ' 寄件者
                    If lv3SortOrder = SortOrder.Ascending Then
                        _lv3MailList = _lv3MailList.OrderBy(Function(x) x.SenderName).ToList()
                    Else
                        _lv3MailList = _lv3MailList.OrderByDescending(Function(x) x.SenderName).ToList()
                    End If
                Case 4 ' 附件數 (依 EntryID 從快取抓取，模擬原 Compare 邏輯)
                    If lv3SortOrder = SortOrder.Ascending Then
                        _lv3MailList = _lv3MailList.OrderBy(Function(x)
                                                                Dim files As List(Of String) = Nothing
                                                                Return If(_cacheAttachFilename.TryGetValue(x.EntryID, files), files.Count, 0)
                                                            End Function).ToList()
                    Else
                        _lv3MailList = _lv3MailList.OrderByDescending(Function(x)
                                                                          Dim files As List(Of String) = Nothing
                                                                          Return If(_cacheAttachFilename.TryGetValue(x.EntryID, files), files.Count, 0)
                                                                      End Function).ToList()
                    End If
            End Select
            ListView3.Invalidate()  ' 💡 關鍵：Invalidate 會通知 ListView 重新按需索取資料，配合 EndUpdate 瞬間更新畫面
        Finally
            ListView3.EndUpdate()
        End Try

        sw.Stop()
        ProgressBar2.Text = $"虛擬排序 {_lv3MailList.Count:N0} 項，耗時 {sw.Elapsed.TotalSeconds:0.00} 秒"
        _dbg("結束", "排序列表") ' by Gemini, 2026/04/10

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
    ' by Gemini 3.1 Pro, 2026/04/21: 邏輯整合 (Tab3 & Tab4)，完整統一行為。
    ' 理由: Tab3 與 Tab4 的 ListView 皆為「搜尋結果」，行為高度一致 (Enter/雙擊/連動與路徑顯示)。
    ' 整合後可減少冗餘代碼，並確保滑鼠與熱鍵行為絕對一致。
    ' --------------------------------------------------------------
    Private Sub HandleLv3Lv4Lv5_KeyDown(sender As Object, e As KeyEventArgs)
        ''' <summary>
        ''' 共通鍵盤按鍵 (Enter: 開啟, ESC: 目錄焦點歸位, Ctrl+A: 全選)
        ''' </summary>
        Dim lv = DirectCast(sender, ListView)

        If e.KeyCode = Keys.Enter Then
            OpenMailByEntryID(GetSelectedEntryIDs(lv))
            e.Handled = True
            e.SuppressKeyPress = True

        ElseIf e.KeyCode = Keys.Escape Then
            If lv.VirtualMode Then lv.SelectedIndices.Clear() Else lv.SelectedItems.Clear()
            ' 對應不同的 TreeView 給予控制權
            If lv Is ListView3 Then SimTree3.Focus()
            If lv Is ListView4 Then SimTree4.Focus()
            If lv Is ListView5 Then SimTree5.Focus() ' 2026/05/03 by Gemini 3.1 Pro: 新增 Tab5 ESC 焦點歸位
            e.Handled = True

        ElseIf e.Control AndAlso e.KeyCode = Keys.A Then
            lv.BeginUpdate()
            ' by Gemini 3 Flash, 2026/05/09: 修復虛擬模式全選當機問題
            If lv.VirtualMode Then
                For i As Integer = 0 To lv.Items.Count - 1 : lv.SelectedIndices.Add(i) : Next   ' 虛擬模式下不可枚舉 Items，改用索引循環或直接操作 SelectedIndices
            Else
                For Each item As ListViewItem In lv.Items : item.Selected = True : Next         ' 實體模式維持原樣
            End If
            lv.EndUpdate()
            e.Handled = True : e.SuppressKeyPress = True
        End If
    End Sub
    Private Sub HandleLv3Lv4Lv5_MouseClick(sender As Object, e As MouseEventArgs)
        ''' <summary>
        ''' 共通滑鼠點擊: 複製主旨與路徑預覽
        ''' </summary>
        Dim lv = DirectCast(sender, ListView)
        Dim item As ListViewItem = lv.GetItemAt(e.X, e.Y)

        If item IsNot Nothing AndAlso e.Button = MouseButtons.Left Then
            ' 單擊左鍵複製主旨到剪貼簿，這原本是 ListView4 獨有的方便設計，現在擴展到 Tab3 共用 (by Gemini 3.1 Pro, 2026/04/21)
            Clipboard.SetText(item.SubItems(0).Text)
        End If
        ' 路徑更新邏輯統一由 ShowLv3Lv4Lv5PathToProgressBar 接管
        ShowLv3Lv4Lv5PathToProgressBar(sender, e)
    End Sub
    Private Sub HandleLv3Lv4Lv5_DoubleClick(sender As Object, e As EventArgs)
        ''' <summary>
        ''' 共通雙擊開啟
        ''' </summary>
        OpenMailByEntryID(GetSelectedEntryIDs(DirectCast(sender, ListView)))
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

        Dim swTotal As New Stopwatch() : swTotal.Start()
        Dim swThrottle As New Stopwatch() : swThrottle.Start()

        Dim mustCountAttach As Boolean = CheckAttCount.Checked
        Dim minCount As Integer = If(mustCountAttach, CInt(CountMin.Value), 0)
        Dim maxCount As Integer = If(mustCountAttach, CInt(CountMax.Value), Integer.MaxValue)

        Dim processed As Integer = 0, total As Integer = sourceList.Count
        ' 預分配容量為 4096，優化搜尋結果清單的填充速度 (by Gemini 3 Flash, 2026/05/04)
        Dim resultList As New List(Of MailItemInfo)(4096)
        Dim keyword As String = If(CheckAttachName.Checked, TextBox3.Text.Trim.ToLower(), "")
        Try
            For curMail As Integer = 0 To sourceList.Count - 1
                ' 2026/4/5, by Gemini: 將進度報告與 UI 釋放移至迴圈開頭，提早反饋處理進度
                ' 避免被下方的 Guard Clauses (Continue For) 略過而導致長時間霸佔主執行緒, 未更新UI進度反饋
                processed = curMail + 1
                ' 2026/04/16 by Gemini 3.0 flash: 改用 ThrottleFreq.Hii + SmartThrottle 與 onThrottled 委派
                Await SmartThrottle(swThrottle, cToken:=cToken, ThrottleFreq.Hii,
                                          Sub()
                                              Dim elapsedSec As Double = Math.Max(swTotal.Elapsed.TotalSeconds, 0.001)
                                              Dim speed As Double = processed / elapsedSec
                                              Dim etaString As String = ""
                                              If total > 500 AndAlso speed > 0 Then
                                                  Dim remainingSec As Integer = CInt(Math.Max(0, (total - processed) / speed))
                                                  If remainingSec > 3 Then etaString = $"，預估剩餘 {remainingSec \ 60:D2}:{remainingSec Mod 60:D2}"
                                              End If
                                              progress?.Report(New ProgressReport With {.CurrentCount = processed, .TotalCount = total,
                                                                                        .Message = $"Phase 2: {processed} / {total}，已符合 {resultList.Count} 封 ({speed:F0} 封/秒{etaString})"})
                                          End Sub)

                Dim currentMail As MailItemInfo = sourceList(curMail)
                Dim cachedAttFilenames As List(Of String) = GetAttachFilename(currentMail)

                ' ── Guard Clause 0: 沒附件資料就不受理 ──
                If cachedAttFilenames Is Nothing Then Continue For

                ' ── Guard Clause 1: 數量過濾 ──
                If mustCountAttach AndAlso
                    (cachedAttFilenames.Count < minCount OrElse cachedAttFilenames.Count > maxCount) Then Continue For

                ' ── Guard Clause 2: 檔名關鍵字過濾 (使用 LINQ Any 取代巢狀 For Each) ──
                If keyword.Length > 0 AndAlso Not cachedAttFilenames.Any(
                    Function(fn) fn IsNot Nothing AndAlso
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
        ''     For Each mail As MailItemInfo In sourceList
        ''         ' 如果快取區存有這封信真的附件數量，就顯示明確數量。如果沒有，代表完全沒有跑過 Phase 2，就顯示 ">0", 不需真的去讀COM物件確認，避免不必要的性能損耗
        ''         Dim cachedFiles As List(Of String) = Nothing
        ''         Dim displayName As String = ">0"
        ''         If _cacheAttachFilename.TryGetValue(mail.EntryID, cachedFiles) Then displayName = cachedFiles.Count.ToString()
        ''         items.Add(New ListViewItem({mail.Subject,
        ''                                     mail.Size.ToString("###,###,##0"),
        ''                                     mail.ReceivedTime.ToShortDateString(),
        ''                                     mail.SenderName,
        ''                                     displayName,
        ''                                     mail.EntryID}))
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
        If entryIDs Is Nothing OrElse entryIDs.Count = 0 Then Return

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
#End Region
#End Region

#Region "■ 07 Tab4: 系列郵件"
#Region "  ├ Layer1 UI事件層"
    Private Async Sub Bt4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        ' by Gemini 3 Flash, 2026/04/20: 修改搜尋來源為 Tab4 專屬的 SimTree4
        _dbg("開始")

        ' ✅ 2026/04/21 by Gemini 3.0 flash: 如果目前是資料夾模式，在清空前先備份節點及其展開狀態
        If Not _isTab4ShowingResults Then
            _tab4FolderTreeNodesBackup.Clear()
            _tab4LastClickedFolderNode = SimTree4.SelectedNode ' 📢 額上記下目前的焦點節點
            For Each node As TreeNode In SimTree4.Nodes
                _tab4FolderTreeNodesBackup.Add(node)
            Next
            _dbg("已備份資料夾樹節點狀態", _tab4FolderTreeNodesBackup.Count & " 個根節點")
        End If

        Dim cToken As CancellationToken = OkayNowYouHaveToken()
        Dim selectedFolders As New List(Of Folder)(32)
        For Each node In SimTree4.SelectedNodes
            Dim f = TryCast(node.Tag, Folder)
            If f IsNot Nothing Then selectedFolders.Add(f)
        Next

        ' ✅ 2026/04/21 by Gemini 3.0 flash: F5 強化邏輯 - 如果未選擇節點，嘗試使用最後一次搜尋的資料夾清單
        If selectedFolders.Count = 0 AndAlso _tab4LastSearchFolders.Count > 0 Then
            selectedFolders.AddRange(_tab4LastSearchFolders)
            _dbg("F5 刷新模式：引用歷史資料夾清單", selectedFolders.Count & " 個資料夾")
        End If

        If selectedFolders.Count = 0 Then Return

        Button4.Enabled = False : Cursor = Cursors.WaitCursor
        ListView4.Items.Clear()
        ProgressBar1.Text = "正在處理..." : ProgressBar2.Text = "開始掃描系列郵件..."
        _tab4LastSearchFolders = New List(Of Folder)(selectedFolders) ' 記憶最後成功的搜尋目標清單

        Dim sw As New Stopwatch() : sw.Start()
        Dim progress4 As IProgress(Of ProgressReport) = New Progress(Of ProgressReport)(Sub(p) ProgressBar2.Text = p.Message)
        Dim swThrottle As New Stopwatch() : swThrottle.Start() ' by Gemini, 2026/04/02: 重用秒錶做節流
        Dim topicDict As New Dictionary(Of String, List(Of MailItemInfo))(StringComparer.OrdinalIgnoreCase)

        Try
            ' ✅ 2026/04/21 by Gemini 3.0 flash: 呼叫共用核心 GetUniqueFolderList (內含路徑去重與子資料夾展開)
            ' 2026/04/22 by Gemini 3.1 Pro: 如果在結果模式刷新，SelectedNodes裝的是話題不是Folder。用偽造的 TreeNode 清單包裝歷史 Folder 傳交給底層。
            Dim fakeNodes As New List(Of TreeNode)(32)
            For Each f In selectedFolders
                fakeNodes.Add(New TreeNode() With {.Tag = f})
            Next
            Dim targetTupleList = Await GetUniqueFolderList(fakeNodes, includeSub:=True, progress:=progress4, cToken:=cToken)
            Dim targetFolderList = targetTupleList.Select(Function(x) x.Folder).ToList()
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
                                        Dim elapsedSec As Double = Math.Max(sw.Elapsed.TotalSeconds, 0.001)
                                        Dim speed As Double = If(processed > 0, processed / elapsedSec, 0)
                                        Dim etaString As String = ""
                                        If targetFolderList.Count > 10 AndAlso speed > 0 Then
                                            Dim remainingSec As Integer = CInt(Math.Max(0, (targetFolderList.Count - processed) / speed))
                                            If remainingSec > 3 Then etaString = $"，預估剩餘 {remainingSec \ 60:D2}:{remainingSec Mod 60:D2}"
                                        End If
                                        progress4?.Report(New ProgressReport With {.CurrentCount = processed, .TotalCount = targetFolderList.Count,
                                                                                   .Message = $"正在掃描系列郵件: {processed} / {targetFolderList.Count} 個資料夾 ({speed:F0} 個/秒{etaString})"})
                                    End Sub)
            Next

            ' ✅ 2026/04/20 by Gemini 2.0 Flash: 記憶結果並呼叫共用渲染函數
            _tab4LastTopicResults = topicDict
            RenderLv4Group(topicDict)

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
    Private Sub Tv4_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles SimTree4.AfterSelect

        ' ✅ 2026/04/20 by Gemini 2.0 Flash: 新增雙模式選取邏輯
        ' 模式 A: 資料夾模式 (目前的行為是選取後僅供搜尋參考，不執行連動)
        _dbg("開始 (A:資料夾模式)", e.Node.Text)
        If Not _isTab4ShowingResults Then Return

        ' 模式 B: 主旨模式 (顯示主旨下的郵件清單)
        _dbg("開始 (B:主旨模式)", e.Node.Text)
        Dim mailList As List(Of MailItemInfo) = TryCast(e.Node.Tag, List(Of MailItemInfo))
        If mailList Is Nothing Then Return

        _lv4SortOrder = SortOrder.Descending    ' 每次點選新節點時，重置排序狀態為預設 (日期降冪)
        _lv4LastSortColumn = 2                  ' 收到日期所在的 index
        mailList.Sort(Function(a, b) b.ReceivedTime.CompareTo(a.ReceivedTime))  ' 排序: 依據時間遞減 (越新的在越前面)
        ShowLv4Result(mailList)
        _dbg("結束", $"顯示 {mailList.Count} 封系列郵件")

    End Sub
    Private Sub Tv4_KeyDown(sender As Object, e As KeyEventArgs) Handles SimTree4.KeyDown
        ' ✅ 2026/04/20 by Gemini 2.0 Flash: 處理 SimTree4 的快捷鍵與模式切換
        _dbg("開始", e.KeyCode.ToString())

        Select Case e.KeyCode
            Case Keys.Enter
                ' 在結果模式下按下 Enter 切換焦點到列表
                If _isTab4ShowingResults AndAlso ListView4.Items.Count > 0 Then ListView4.Focus()
                e.Handled = True
                e.SuppressKeyPress = True

            Case Keys.F5
                ' 按下 F5 等同 Button4 (重新開始掃描系列郵件)
                ' ✅ 2026/04/20: 在結果模式下按 F5 會自動引用上一資料夾重新掃描
                Button4.PerformClick()
                e.Handled = True

            Case Keys.F6
                ' ✅ 2026/04/20 by Gemini 2.0 Flash: 切換左側樹排序方式 (數量/名稱)
                If _isTab4ShowingResults AndAlso _tab4LastTopicResults IsNot Nothing Then
                    _tab4SortGroupsByCount = Not _tab4SortGroupsByCount
                    RenderLv4Group(_tab4LastTopicResults)
                    _dbg("F6 按下：切換排序為", If(_tab4SortGroupsByCount, "數量", "主旨"))
                    e.Handled = True
                End If

            Case Keys.Escape
                ' 按下 ESC：從結果模式恢復為資料夾模式
                If _isTab4ShowingResults Then
                    _dbg("ESC 按下：恢復資料夾模式 (還原備份節點)")
                    _isTab4ShowingResults = False
                    SimTree4.BeginUpdate()
                    Try
                        SimTree4.Nodes.Clear()
                        ListView4.Items.Clear()

                        ' ✅ 2026/04/21 by Gemini 3.0 flash: 直接還原備份的節點，保持原有展開狀態與選取
                        If _tab4FolderTreeNodesBackup.Count > 0 Then
                            For Each node In _tab4FolderTreeNodesBackup
                                SimTree4.Nodes.Add(node)
                            Next

                            ' 回復最後選中的單一位置 (by Gemini 3.0 flash, 2026/04/21)
                            SimTree4.ClearSelectedNodes()
                            If _tab4LastClickedFolderNode IsNot Nothing Then
                                SimTree4.SelectedNode = _tab4LastClickedFolderNode
                                _tab4LastClickedFolderNode.EnsureVisible()
                            End If
                        Else
                            ' 萬一備份為空 (例如直接按 ESC)，才執行重新讀取
                            LoadStoreToTreeView(_pstStoreList, SimTree4)
                            ExpandTvToDefaultInbox(SimTree4)
                        End If
                    Finally
                        SimTree4.EndUpdate()
                    End Try

                    ProgressBar1.Text = "已恢復資料夾樹模式。" : ProgressBar2.Text = ""
                    SimTree4.Focus() ' 將焦點還給左側
                    e.Handled = True
                    e.SuppressKeyPress = True ' ✅ by Gemini 3.0 flash, 2026/04/21: 徹底攔截，避免 KeyPress 重複執行退回邏輯
                End If
        End Select

    End Sub
    Private Sub Lv4_DrawItem(sender As Object, e As DrawListViewItemEventArgs) Handles ListView4.DrawItem
        ' by Gemini 3.1 Pro, 2026/04/26: 針對被 Hover 但未選取的項目，交由 DrawSubItem 自行畫上灰底；其餘讓系統自己畫
        If e.Item Is _lastHoveredListItem AndAlso Not e.Item.Selected Then
            ' 不設 DrawDefault = True，讓系統呼叫 DrawSubItem
        Else
            e.DrawDefault = True
        End If
    End Sub
    Private Sub Lv4_ColumnClick(sender As Object, e As ColumnClickEventArgs) Handles ListView4.ColumnClick

        ' by Gemini 3 Flash, 2026/04/19: ListView4 欄位排序 (實體模式，參考 ListView3 的虛擬模式做法)
        _dbg("開始", $"點擊欄位: {e.Column}")
        Dim sw As New Stopwatch : sw.Start()

        ' 取得目前選取節點的郵件清單 (因為是從 SimTree4 選取的，資料存在 Tag 裡)
        Dim mailList As List(Of MailItemInfo) = TryCast(SimTree4.SelectedNode?.Tag, List(Of MailItemInfo))
        If mailList Is Nothing OrElse mailList.Count = 0 Then Return

        ' 切換排序方式
        _lv4SortOrder = If(e.Column = _lv4LastSortColumn AndAlso _lv4SortOrder = SortOrder.Ascending, SortOrder.Descending, SortOrder.Ascending)
        _lv4LastSortColumn = e.Column

        ' 根據點選的欄位進行排序 (by Gemini 2.0 Flash: 恢復先前誤刪的邏輯)
        Select Case e.Column
            Case 0 ' 主旨
                If _lv4SortOrder = SortOrder.Ascending Then
                    mailList = mailList.OrderBy(Function(x) x.Subject).ToList()
                Else
                    mailList = mailList.OrderByDescending(Function(x) x.Subject).ToList()
                End If
            Case 1 ' 郵件大小
                If _lv4SortOrder = SortOrder.Ascending Then
                    mailList = mailList.OrderBy(Function(x) x.Size).ToList()
                Else
                    mailList = mailList.OrderByDescending(Function(x) x.Size).ToList()
                End If
            Case 2 ' 收到日期
                If _lv4SortOrder = SortOrder.Ascending Then
                    mailList = mailList.OrderBy(Function(x) x.ReceivedTime).ToList()
                Else
                    mailList = mailList.OrderByDescending(Function(x) x.ReceivedTime).ToList()
                End If
            Case 3 ' 寄件者
                If _lv4SortOrder = SortOrder.Ascending Then
                    mailList = mailList.OrderBy(Function(x) x.SenderName).ToList()
                Else
                    mailList = mailList.OrderByDescending(Function(x) x.SenderName).ToList()
                End If
            Case Else
                _dbg("結束", "無效欄位")
                Return
        End Select

        SimTree4.SelectedNode.Tag = mailList    ' 💡 重要：因為 LINQ 的 .ToList() 會產生新清單，所以必須把排序後的清單再塞回 Tag
        ShowLv4Result(mailList)                  ' 重新填入 ListView
        sw.Stop()
        _dbg("結束", "排序完成")

    End Sub
    Private Async Sub Lv4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListView4.SelectedIndexChanged
        ' ---------------------------------------------------------------
        ' Lv4_UpdateSimilarity — 選定郵件後，非同步計算並更新全列表內文相似度欄 (Index 4)
        ' 2026/04/28 by Simon/Claude: 合併 Simon 架構 + Claude 的 L2.5/L3 分層與 NormalizeMailBody
        '
        ' 架構 (以 Simon 為主):
        '   ① Task.Delay(100) — 避開 SelectedItems 過渡期為空的 Windows 原生兩次觸發
        '   ② _lv4SimCts 取消機制 — 游標快速移動時取消前次未完成的計算，不讓舊任務蓋掉新結果
        '   ③ 先同步標記全列表（基準=「Base」，其他=「...」），再逐封非同步計算
        '   ④ Jaccard 計算放入 Task.Run 背景執行緒，真正不阻塞 UI
        '   ⑤ EntryID 從 SubItems(5) 讀取（直接、輕量，不做 DirectCast）
        '   ⑥ Body 讀取透過 L2.5 GetMailBody（快取 → L3 COM），不跨群組殘留
        '   ⑦ 比對範圍：全列表（不限同組），方便跨群組發現高相似度郵件
        ' ---------------------------------------------------------------
        Dim lv = DirectCast(sender, ListView)
        If lv.Items.Count = 0 Then Return

        ' 💡 提前檢查，避開 Windows 兩次觸發導致的重複日誌
        If lv.SelectedItems.Count = 0 Then Return
        _dbg("開始")

        ' 💡 關鍵修正 1：微小延遲確保選取狀態穩定，避開 SelectedItems 在這100ms內快速移動游標導致的 SelectedItems 為空狀態造成exception
        Await Task.Delay(100)
        If lv.SelectedItems.Count = 0 Then Return

        ' 取消前次未完成的計算任務 (還沒算完就快速移動游標的話，直接取消前次任務)
        _lv4SimCts?.Cancel()
        _lv4SimCts = New CancellationTokenSource()
        Dim token As CancellationToken = _lv4SimCts.Token

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

        ' 逐封取 Body、背景計算 Jaccard、即時更新欄位
        Try
            ' 💡 2026/04/30 by Gemini 3.1 Pro: 兩階段處理。第一階段在 UI Context 循序拿 Body，確保若是 Cache Miss 去讀 COM 的安全性。
            ' 預分配容量為 512，處理大量郵件內文比對時減少頻繁 Resize (by Gemini 3 Flash, 2026/05/04)
            ' 💡 2026/05/09 by Gemini 3.0 flash: 優化非同步頻率。每處理一批(例如 30 封)才 Yield 一次讓 UI 喘氣，減少頻繁切換的負擔
            Dim mBodyList As New List(Of (Item As ListViewItem, TargetBody As String))(512)
            Dim processedCount As Integer = 0
            For Each item In lviCompareList
                If token.IsCancellationRequested Then Exit For
                If item.SubItems.Count <= 4 Then Continue For

                Dim targetID As String = If(item.SubItems.Count > 5, item.SubItems(5).Text, "")
                If targetID = baseEntryID OrElse String.IsNullOrEmpty(targetID) Then Continue For

                Dim targetBody As String = GetMailBody(targetID)
                If String.IsNullOrEmpty(targetBody) Then
                    item.SubItems(4).Text = "失敗" : Continue For
                End If
                mBodyList.Add((item, targetBody))
                processedCount += 1
                If processedCount Mod 30 = 0 Then Await Task.Delay(1) ' 每 30 封釋放一次 UI 執行緒，兼顧流暢度與效率 (by Gemini 3.0 flash, 2026/05/09)
            Next

            If token.IsCancellationRequested Then Return

            ' 💡 2026/04/30 by Gemini 3.1 Pro: 第二階段純 CPU 運算。剝離 UI 與 COM，利用多核心火力全開
            Dim results((mBodyList.Count) - 1) As Double
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
                If item.ListView IsNot Nothing Then
                    item.SubItems(4).Text = $"{CInt(results(i) * 100)}%"
                End If
            Next
            lv.EndUpdate()
        Catch ex As OperationCanceledException
            ' 正常取消，不需處理
        End Try
        _dbg("結束")

    End Sub
    Private Async Sub Lv4_KeyDown(sender As Object, e As KeyEventArgs) Handles ListView4.KeyDown

        ' by Gemini 3.1 Pro, 2026/04/21: Tab4 專屬快捷鍵 (Delete, F5)
        ' ESC 等共通快捷鍵已被遷移至 InitListView 掛載的 HandleLv3Lv4Lv5_KeyDown
        _dbg("開始", e.KeyCode.ToString())
        If e.KeyCode = Keys.Delete Then
            _dbg("快捷鍵", "偵測到 Delete (呼叫 HandleListView4Delete)")
            HandleLv4Delete(DirectCast(sender, ListView))
            e.Handled = True

        ElseIf e.KeyCode = Keys.F5 Then
            ' 按下 F5 重新從 Outlook 讀取目前的郵件內容 (by Gemini 3 Flash, 2026/04/20)
            ' 💡 提示：若按下沒反應，請確認 ListView4 已獲得焦點。
            _dbg("快捷鍵", "偵測到 F5 (重新整理)")
            Await RefreshLv4MailsAsync(DirectCast(sender, ListView))
            e.Handled = True
        End If

    End Sub
#End Region
#Region "  ├ Layer2 流程協調層"
    Private Sub RenderLv4Group(topicDict As Dictionary(Of String, List(Of MailItemInfo)))
        ''' <summary>
        ''' ✅ 2026/04/20 by Gemini 2.0 Flash: 根據目前的排序模式渲染 Tab4 的主旨群組樹
        ''' </summary>

        _dbg("開始")
        If topicDict Is Nothing Then Return

        SimTree4.BeginUpdate()
        SimTree4.Nodes.Clear()
        _isTab4ShowingResults = True

        _dbg("渲染系列清單", $"模式: {If(_tab4SortGroupsByCount, "按數量", "按主旨")}")
        ' 根據旗標決定排序方式
        Dim sortedItems = If(_tab4SortGroupsByCount,
            topicDict.Where(Function(kvp) kvp.Value.Count > 1).OrderByDescending(Function(kvp) kvp.Value.Count).ThenBy(Function(kvp) kvp.Key),
            topicDict.Where(Function(kvp) kvp.Value.Count > 1).OrderBy(Function(kvp) kvp.Key))
        For Each kvp In sortedItems
            Dim node As New TreeNode($"{kvp.Key} ({kvp.Value.Count})") With {.Tag = kvp.Value}
            SimTree4.Nodes.Add(node)
        Next
        SimTree4.EndUpdate()

        ' ✅ by Gemini 3.0 flash, 2026/04/21: 搜尋完成後，自動選取第一個結果並 Focus
        ' 💡 補充: 為了確保右側 ListView4 同步更新，手動呼叫事件處理器 (by Gemini 3.0 flash, 2026/04/21)
        If SimTree4.Nodes.Count > 0 Then
            Dim firstNode = SimTree4.Nodes(0)
            SimTree4.SelectedNode = firstNode
            SimTree4.Focus()
            Tv4_AfterSelect(SimTree4, New TreeViewEventArgs(firstNode))
        End If
        ProgressBar1.Text = $"找到 {SimTree4.Nodes.Count} 個系列 (排序: {If(_tab4SortGroupsByCount, "數量", "主旨")})"
        _dbg("結束")

    End Sub
    Private Sub ShowLv4Result(mailList As List(Of MailItemInfo))

        ' by Gemini 3 Flash, 2026/04/20: 實作智慧分組 (排除 Re:/Fw:) 與動態排序邏輯
        ' 確保資料清單被記住，以便 F6 切換時使用
        _dbg("開始")
        ListView4.Tag = mailList
        ListView4.BeginUpdate()
        ListView4.Items.Clear()
        ListView4.Groups.Clear()

        If mailList Is Nothing OrElse mailList.Count = 0 Then
            ListView4.EndUpdate() : Return
        End If

        ' 1. 執行分組 (LINQ GroupBy 智慧清理後的主旨)
        Dim groups = mailList.GroupBy(Function(m) GetCleanSubject(m.Subject))

        ' 2. 依照排序模式對「組」進行排序
        Dim sortedGroups As IEnumerable(Of IGrouping(Of String, MailItemInfo))
        If _lv4GroupSortByCount Then
            ' 模式：按組內數量遞減排序
            sortedGroups = groups.OrderByDescending(Function(g) g.Count()).ThenBy(Function(g) g.Key)
        Else
            ' 模式：按主旨字母順序排序
            sortedGroups = groups.OrderBy(Function(g) g.Key)
        End If

        ' 3. 逐組渲染到 UI
        For Each group In sortedGroups
            ' 建立組標題：主旨 (數量封)
            Dim groupHeader As String = $"{group.Key} ({group.Count} 封)"
            Dim lvGroup As New ListViewGroup(group.Key, groupHeader)
            ListView4.Groups.Add(lvGroup)

            ' ✅ 2026/04/20 by Gemini 2.0 Flash: 連動 Column Header 的點擊排序
            ' 根據全域變數 _lv4LastSortColumn 對組內項目進行動態排序
            Dim sortedItems As IEnumerable(Of MailItemInfo)
            Select Case _lv4LastSortColumn
                Case 0 ' 主旨
                    sortedItems = If(_lv4SortOrder = SortOrder.Ascending, group.OrderBy(Function(m) m.Subject), group.OrderByDescending(Function(m) m.Subject))
                Case 1 ' 郵件大小
                    sortedItems = If(_lv4SortOrder = SortOrder.Ascending, group.OrderBy(Function(m) m.Size), group.OrderByDescending(Function(m) m.Size))
                Case 2 ' 收到日期
                    sortedItems = If(_lv4SortOrder = SortOrder.Ascending, group.OrderBy(Function(m) m.ReceivedTime), group.OrderByDescending(Function(m) m.ReceivedTime))
                Case 3 ' 寄件者
                    sortedItems = If(_lv4SortOrder = SortOrder.Ascending, group.OrderBy(Function(m) m.SenderName), group.OrderByDescending(Function(m) m.SenderName))
                Case Else
                    sortedItems = group.OrderByDescending(Function(m) m.ReceivedTime)
            End Select

            ' 收集該組的所有項目，再一次性 AddRange (by Gemini 3.0 flash, 2026/04/21)
            ' 預分配容量為 512，優化重複郵件掃描結果的 UI 清單組裝 (by Gemini 3 Flash, 2026/05/04)
            Dim groupItems As New List(Of ListViewItem)(512)
            For Each mailItem In sortedItems
                ' by Gemini 3.0 Flash, 2026/04/20: 郵件大小改為位元組(精細), 日期格式統一 yyyy/MM/dd (補零+置中需求)
                Dim lvi As New ListViewItem({mailItem.Subject,
                                             mailItem.Size.ToString("N0"),
                                             mailItem.ReceivedTime.ToString("yyyy/MM/dd HH:mm:ss"),
                                             mailItem.SenderName,
                                             " - ",
                                             mailItem.EntryID})
                lvi.Tag = mailItem ' by Gemini 3.0 flash, 2026/04/21: 直接存入物件避開 Index 錯位問題
                lvi.Group = lvGroup
                groupItems.Add(lvi)
            Next
            ListView4.Items.AddRange(groupItems.ToArray())
        Next
        ListView4.EndUpdate()

        ' 4. 更新狀態列反饋 (by Gemini 3 Flash, 2026/04/20)
        Dim sortModeStr = If(_lv4GroupSortByCount, "數量排序", "名稱排序")
        ProgressBar2.Text = $"系列選中：{ListView4.Groups.Count:N0} 個主題，共 {mailList.Count:N0} 封郵件 (目前：{sortModeStr})"
        _dbg("結束")

    End Sub
    Private Sub HandleLv4Delete(lv As ListView)
        ' by Gemini 3 Flash, 2026/04/20: 處理 ListView4 的刪除邏輯
        _dbg("開始")
        Dim selCount As Integer = lv.SelectedItems.Count
        If selCount = 0 Then Return

        If MessageBox.Show($"確定要將選中的 {selCount} 封郵件移到「刪除郵件」資料夾嗎？", "確認刪除",
                           MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then

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
                    End If
                Next
                MoveMailsToRecycle(entryIDs)    ' 實體刪除 (移動到預設刪除資料夾)
                ShowLv4Result(mailList)         ' 重新整理 UI
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
        '   修正：改用 mail.Delete() 自動移入同一 Store 的刪除郵件資料夾，
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
    Private Async Function RefreshLv4MailsAsync(lv As ListView) As Task

        ' by Gemini 3 Flash, 2026/04/20: 重新讀取目前系列郵件的最新資訊並更新 MailItemInfo
        _dbg("開始")
        ' ✅ by Gemini 3.0 flash, 2026/04/21: 修正控制項名稱為 SimTree4
        Dim mailList As List(Of MailItemInfo) = TryCast(SimTree4.SelectedNode?.Tag, List(Of MailItemInfo))
        If mailList Is Nothing OrElse mailList.Count = 0 Then Return

        _isUserBusy = True : Cursor = Cursors.WaitCursor
        Dim swThrottle As New Stopwatch : swThrottle.Start()
        Dim total As Integer = mailList.Count

        Try
            For i As Integer = 0 To total - 1
                Dim info As MailItemInfo = mailList(i)
                Dim mail As Outlook.MailItem = Nothing
                Try
                    ' 從 Outlook 讀取最新狀態
                    mail = CType(_olNS.GetItemFromID(info.EntryID), Outlook.MailItem)
                    If mail IsNot Nothing Then
                        info.Subject = mail.Subject
                        info.Size = mail.Size
                        info.ReceivedTime = mail.ReceivedTime
                        info.SenderName = mail.SenderName
                        mailList(i) = info ' 必須寫回，因為是 Structure (Value Type)
                    End If
                Catch ex As System.Exception
                    _dbg("讀取失敗", $"ID: {info.EntryID}, Error: {ex.Message}")
                Finally
                    TryMarshalRelease(mail)
                End Try

                ' 2026/04/16 by Gemini 3.0 flash: 改用 ThrottleFreq.Hii + SmartThrottle 與 onThrottled 委派
                Await SmartThrottle(swThrottle, cToken:=CancellationToken.None, ThrottleFreq.Hii, Sub() ProgressBar2.Text = $"正在重新讀取郵件資訊: {i + 1} / {total}...")
            Next

            ShowLv4Result(mailList)  ' 重新填寫列表 (保留目前的排序狀態，因為資料是原地更新)
            ProgressBar1.Text = $"已重新讀取 {total} 封郵件。" : ProgressBar2.Text = ""

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

        ' 統計 ASCII
        For i As Integer = 0 To 255
            lSum(charTable(i)) += 1
        Next
        ' 統計 CJK
        For i As Integer = &H2E80 To &H9FBF
            lSum(charTable(i)) += 1
        Next

        ' 4. 計算相似度
        Dim denominator As Integer = lSum(1) + lSum(2) + lSum(3)
        If denominator = 0 Then Return 0
        Return lSum(3) / denominator
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
        Dim isExact As Boolean = rbExactMatch.Checked
        Dim includeSub As Boolean = _includeSubTab5
        Button5.Enabled = False : Cursor = Cursors.WaitCursor
        ListView5.BeginUpdate() : ListView5.Items.Clear() : ListView5.EndUpdate()
        ProgressBar1.Text = "正在準備" : ProgressBar2.Text = "展開資料夾結構..."
        Dim sw As New Stopwatch() : sw.Start()
        Dim progress5 As IProgress(Of ProgressReport) = New Progress(Of ProgressReport)(Sub(p) ProgressBar2.Text = p.Message)

        Try
            Dim folderList = Await GetUniqueFolderList(selectedNodes, includeSub:=includeSub, cToken:=cToken, progress:=progress5)
            If folderList.Count = 0 Then Return

            Dim groupDict = Await ScanMailsToGroupDictAsync(folderList, isExact, progress5, cToken)
            _tab5LastGroupResults = groupDict ' by Gemini 3 Flash, 2026/05/06: 儲存結果以供動態刪除
            _tab5LastIsExact = isExact
            Dim counts = RenderLv5Group(groupDict, isExact)

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
        '   MessageID / SenderEmail 已整合至 MailItemInfo，BuildMailGroupKey 直接使用
        ' ---------------------------------------------------------------
        Dim groupDict As New Dictionary(Of String, List(Of MailItemInfo))(StringComparer.OrdinalIgnoreCase)
        Dim totalFolders As Integer = folderList.Count
        Dim totalProcessed As Integer = 0
        Dim swThrottle As New Stopwatch() : swThrottle.Start()
        Dim swTotal As New Stopwatch() : swTotal.Start()    ' 2026/05/10 by Simon/Claude: 供 ETA 計算使用

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
                    Dim hashKey As String = BuildMailGroupKey(m.MessageID, m.Subject, senderKey, m.Size, m.ReceivedTime, isExact)
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
                                    Dim elapsedSec As Double = Math.Max(swTotal.Elapsed.TotalSeconds, 0.001)
                                    Dim speed As Double = If(totalProcessed > 0, totalProcessed / elapsedSec, 0)
                                    Dim etaString As String = ""
                                    If totalFolders > 10 AndAlso speed > 0 Then
                                        Dim remainingSec As Integer = CInt(Math.Max(0, (totalFolders - totalProcessed) / speed))
                                        If remainingSec > 3 Then etaString = $"，預估剩餘 {remainingSec \ 60:D2}:{remainingSec Mod 60:D2}"
                                    End If
                                    progress?.Report(New ProgressReport With {.Message = $"掃描中: {totalProcessed}/{totalFolders} 個資料夾 ({speed:F0} 個/秒{etaString})"})
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
        '             ── [預留架構] SimHash 內文比對 ──
        '             日後將加入 SimHash + Hamming Distance ≤ 5 作為一階快速篩選，
        '             再以 Jaccard body similarity ≥ 0.8 做二階精細比對。
        '             屆時 isExact=False 的二階邏輯由此函數擴充，呼叫端無需修改。
        ' 使用 AddRange 批次寫入，避免逐行 Add 觸發多次 UI 更新
        ' 2026/05/05 by Claude
        ' ---------------------------------------------------------------
        ListView5.Tag = groupDict ' by Gemini 3 Flash, 2026/05/06: 將資料來源掛載至 Tag 供 HandleLv5Delete 使用
        ListView5.BeginUpdate()
        ListView5.Items.Clear()

        Dim groupID As Integer = 1
        Dim totalMails As Integer = 0

        For Each kvp In groupDict
            If kvp.Value.Count <= 1 Then Continue For

            ' ── 相似度計算 ──
            ' 2026/05/06 by Gemini 3 Flash: 不論模式一律計算相似度，但僅在 Fuzzy 模式下套用 0.6 門檻過濾
            Dim simScores As New List(Of Double)(512)
            Dim isValidGroup As Boolean = True
            Dim firstSubject As String = kvp.Value(0).Subject
            simScores.Add(1.0) ' 第一封永遠是基準 100%

            For i As Integer = 1 To kvp.Value.Count - 1
                Dim sim As Double = JaccardSimilarity(firstSubject, kvp.Value(i).Subject)
                simScores.Add(sim)

                ' 僅在模糊模式下才套用門檻過濾 (0.6)
                If Not isExact AndAlso sim < 0.6 Then isValidGroup = False : Exit For
            Next
            ' ── [預留] SimHash 內文比對將在此插入 ──
            ' If isValidGroup Then isValidGroup = SimHashBodyFilter(kvp.Value)

            If Not isValidGroup Then Continue For

            ' ── 建立 ListViewItem 清單，一次 AddRange ──
            Dim groupColor As Color = If(groupID Mod 2 = 0, Color.FromArgb(240, 248, 255), Color.White)
            Dim lvItems As New List(Of ListViewItem)(kvp.Value.Count)
            For idx As Integer = 0 To kvp.Value.Count - 1
                Dim m As MailItemInfo = kvp.Value(idx)
                Dim simText As String = If(idx < simScores.Count, $"{CInt(simScores(idx) * 100)}%", "-")
                lvItems.Add(New ListViewItem({m.Subject,
                                              (m.Size \ 1024L).ToString("N0") & "KB",
                                              m.ReceivedTime.ToString("yyyy/MM/dd HH:mm:ss"),
                                              m.SenderName,
                                              "G" & groupID.ToString(),
                                              simText,
                                              m.EntryID}) With {.BackColor = groupColor, .Tag = m})
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

            Dim entryIDs As New List(Of String)(selCount)
            ' 取得快取的資料字典
            Dim groupDict As Dictionary(Of String, List(Of MailItemInfo)) = _tab5LastGroupResults
            If groupDict Is Nothing Then Return

            ' 收集選中項目的 EntryID 並從資料源中移除
            For Each item As ListViewItem In lv.SelectedItems
                If TypeOf item.Tag Is MailItemInfo Then
                    Dim info = DirectCast(item.Tag, MailItemInfo)
                    entryIDs.Add(info.EntryID)

                    ' 從 groupDict 中移除該封信 (遍歷所有群組尋找)
                    For Each kvp In groupDict
                        ' 找到並移除後，如果該群組只剩 1 封或 0 封，在重複郵件邏輯中視為不再重複，可選擇保留或由渲染器過濾
                        If kvp.Value.RemoveAll(Function(m) m.EntryID = info.EntryID) > 0 Then Exit For
                    Next
                End If
            Next

            If entryIDs.Count > 0 Then
                MoveMailsToRecycle(entryIDs)    ' 實體移動
                RenderLv5Group(groupDict, _tab5LastIsExact) ' 重新渲染 UI
                ProgressBar2.Text = $"已移動 {selCount} 封郵件至刪除郵件資料夾"
            End If
        End If
        _dbg("結束")
    End Sub
#End Region
#Region "  └ 輔助函數"
#End Region
#End Region

#Region "■ 09 Tab6: Debug & 設定"
    Private Async Sub SaveCache_Click(sender As Object, e As EventArgs) Handles SaveCache.Click
        Await SaveCachesToDB()
        RefreshDatabaseStats()
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
            f.Font = New Font("Microsoft JhengHei UI", 10)

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

        RefreshDatabaseStats()
        _dbg("結束")

    End Sub
    Private Async Sub RenewCache_Click(sender As Object, e As EventArgs) Handles RenewCache.Click
        ' 2026/04/09 重構: 原本只做孤兒清除，現在改呼叫完整的 RenewCacheToDB
        '   RenewCacheToDB 內含: Phase1 BFS → Phase2 snapshot 比對 → Phase3 dirty 重算
        '                         Phase4 ancestor 聚合清除 → Phase5 month_counts DB 清除
        '                         Phase6 CleanupOrphan + SaveCachesToDB
        '   RenewIncludeSize 勾選時才重算 folder_size (GetTable 遍歷，大資料夾較慢) 
        Try
            Await RenewCacheToDB(RenewIncludeSize.Checked)

            ' by Gemini 3.0 flash, 2026/04/24: 更新完成後，執行非同步 UI 刷新，確保新資料夾能立即顯示
            Await RefreshAllTreeViews()

            RefreshDatabaseStats()
        Catch ex As OperationCanceledException
            _dbg(" ├ 中斷", "使用者已取消快取更新")
        End Try
    End Sub
    Private Async Sub RefreshDatabaseStats()
        ' ---------------------------------------------------------------
        ' RefreshDatabaseStats — 切換到 Setting 頁時呼叫，更新 txtDatabaseStats / _lvStats
        '
        ' 2026/04/20 重構要點 (by Gemini 3 Flash):
        '   1. 改為 Async Sub，使用 Task.Run 取得資料庫摘要，基礎解決 Tab 切換卡頓。
        '   2. 動態將 txtDatabaseStats 替換為 ListView，改用 Noto Sans TC 字型。
        '   3. 使用 ListView 的雙欄結構，完美達成靠右對齊，且文字渲染較優美。
        ' ---------------------------------------------------------------
        If txtDatabaseStats Is Nothing Then Return

        ' ── 步驟 1: 第一次執行時，動態建立並配置 ListView ──
        If _lvStats Is Nothing Then
            _lvStats = New ListView()
            _lvStats.View = System.Windows.Forms.View.Details
            _lvStats.FullRowSelect = True
            _lvStats.HeaderStyle = ColumnHeaderStyle.None ' 隱藏標題節省空間
            _lvStats.GridLines = False
            _lvStats.BorderStyle = BorderStyle.None
            _lvStats.Location = txtDatabaseStats.Location
            _lvStats.Size = txtDatabaseStats.Size
            _lvStats.Height += 100 ' ListView 在顯示列表時需要比單一文字框多一點高度空間
            _lvStats.Anchor = txtDatabaseStats.Anchor
            _lvStats.BackColor = Color.White
            _lvStats.Columns.Add("Item", 200)
            _lvStats.Columns.Add("Value", 120, HorizontalAlignment.Right)
            _lvStats.Font = New Font("Microsoft Jhenghei UI", 9.5F)

            txtDatabaseStats.Parent.Controls.Add(_lvStats)
            _lvStats.BringToFront()
            txtDatabaseStats.Visible = False ' 隱藏原本的 TextBox
        End If

        _lvStats.Items.Clear()
        _lvStats.Items.Add(New ListViewItem("📊 正在讀取統計資料..."))

        Try
            ' ── 步驟 2: 非同步讀取資料庫摘要 (解決卡頓核心) ──
            ' 將耗時的 SQL COUNT(*) 移至背景執行緒
            Dim st = Await Task.Run(Function() GetDBSummary())

            _lvStats.BeginUpdate()
            _lvStats.Items.Clear()

            ' 輔助方法：填入統計項目 (VB.NET Lambda 不支援 Optional 參數，故移除並於呼叫處補齊)
            Dim AddStat = Sub(label As String, val As String, isHeader As Boolean)
                              Dim itm = New ListViewItem(label)
                              itm.SubItems.Add(val)
                              If isHeader Then
                                  itm.ForeColor = Color.DarkBlue
                                  itm.Font = New Font(_lvStats.Font, FontStyle.Bold)
                              End If
                              _lvStats.Items.Add(itm)
                          End Sub

            ' ── 步驟 3: 填充 Memory 數據 ──
            AddStat("═══ Memory 快取 ════", "", True)
            AddStat("_cacheFolderTree", _cacheFolderTree.Count.ToString("N0") & " 筆", False)
            AddStat("_cacheSubTreeList", _cacheSubTreeList.Count.ToString("N0") & " 筆", False)
            AddStat("_cacheIsMailFolder", _cacheIsMailFolder.Count.ToString("N0") & " 筆", False)
            AddStat("_cacheFolderIDs", _cacheFolderIDs.Count.ToString("N0") & " 筆", False)
            AddStat("", "", False) ' 間隔
            AddStat("_cacheMailCount", _cacheMailCount.Count.ToString("N0") & " 筆", False)
            AddStat("_cacheMailCountAll", _cacheMailCountAll.Count.ToString("N0") & " 筆", False)
            AddStat("_cacheFolderCount", _cacheFolderCount.Count.ToString("N0") & " 筆", False)
            AddStat("_cacheFolderCountAll", _cacheFolderCountAll.Count.ToString("N0") & " 筆", False)
            AddStat("_cacheFolderSize", _cacheFolderSize.Count.ToString("N0") & " 筆", False)
            AddStat("_cacheFolderSizeAll", _cacheFolderSizeAll.Count.ToString("N0") & " 筆", False)
            AddStat("_cacheYearCounts", _cacheYearCounts.Count.ToString("N0") & " 筆", False)
            AddStat("_cacheMonthCounts", _cacheMonthCounts.Count.ToString("N0") & " 筆", False)
            AddStat("_cacheAttachMailList", _cacheAttachMailList.Count.ToString("N0") & " 筆", False)
            AddStat("_cacheAttachFilename", _cacheAttachFilename.Count.ToString("N0") & " 筆", False)
            AddStat("", "", False) ' 間隔

            ' ── 步驟 4: 填充 SQLite 數據 ──
            ' 拆分日期與時間
            Dim datePart As String = "N/A" : Dim timePart As String = "N/A"
            If st.lastTs.Contains(" "c) Then
                Dim parts = st.lastTs.Split(" "c)
                datePart = parts(0) : timePart = parts(1)
            Else
                datePart = st.lastTs
            End If
            AddStat("════ SQLite 快取 ════", "", True)
            AddStat("DB 檔案大小", st.kb.ToString("N0") & " KB", False)
            AddStat("folder_stats", st.fc.ToString("N0") & " 筆", False)
            AddStat("basic_maillist", st.basic.ToString("N0") & " 筆", False)   ' by Gemini 3 Flash, 2026/04/22
            AddStat("year_counts", st.yc.ToString("N0") & " 筆", False)
            AddStat("month_counts", st.mc.ToString("N0") & " 筆", False)
            AddStat("attach_maillist", st.mb.ToString("N0") & " 筆", False)
            AddStat("attach_filenames", st.at.ToString("N0") & " 筆", False)
            AddStat("最後更新日期", datePart, False)
            AddStat("最後更新時間", timePart, False)

            _lvStats.EndUpdate()
        Catch ex As System.Exception
            _lvStats.Items.Clear()
            _lvStats.Items.Add(New ListViewItem("❌ 讀取統計失敗: " & ex.Message))
        End Try
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
#End Region

End Class
