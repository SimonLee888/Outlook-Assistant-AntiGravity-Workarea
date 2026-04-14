Imports System.Collections.Concurrent
Imports System.Formats
Imports System.Globalization
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Windows.Forms.DataVisualization.Charting
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Outlook
Imports Redemption
Imports SQLitePCL
Imports Windows.UI.Composition

Partial Class Form1

#Region "■ 01 全域宣告"
    Private _fontDefault As New Font("Microsoft Jhenghei", 10.0F, System.Drawing.FontStyle.Regular, GraphicsUnit.Point, 0)
    Private _fontHeader As New Font("Microsoft Jhenghei", 10.0F, System.Drawing.FontStyle.Bold, GraphicsUnit.Point, 0)
    Private _fontRegular = System.Drawing.FontStyle.Regular
    Private _fontBold = System.Drawing.FontStyle.Bold
    Private _fontItalic = System.Drawing.FontStyle.Italic

    ' ── 全域勾選狀態變數 (by Gemini, 2026/04/10: 優化效能，避免頻繁讀取 UI) ──
    Private _includeSubTab2 As Boolean = False
    Private _includeSubTab3 As Boolean = False
    Private _showAllFolders As Boolean = False

    Private _lastHoveredTreeNode As TreeNode = Nothing
    Private _lastHoveredListItem As ListViewItem = Nothing
    Private _lastHoveredPointIndex As Integer = -1                  ' 記住上一個 hover 的點，-1 表示沒有

    Private _tab1SelectSeq As Integer = 0                           ' Tab1 快速點選防護序號
    Private _tab2SelectSeq As Integer = 0                           ' Tab2 快速點選防護序號
    Private _tab2MonthViewYear As Integer = 0                       ' 目前月份視圖顯示的是哪一年
    Private _tab2IsMonthView As Boolean = False                     ' 目前 ListView2 顯示的是月份視圖還是年度視圖
    Private _tab2FolderList As List(Of Outlook.Folder) = Nothing    ' 記住目前 Tab2 的資料夾清單，供月份展開使用

    Private _lv2DataYear As ConcurrentDictionary(Of Integer, Integer) = Nothing  ' Tab2 年度視圖 session 快取（已合併多資料夾），月份進出時直接 render 不重算
    Private _lv2DataMonth As ConcurrentDictionary(Of Integer, Integer) = Nothing ' Tab2 月份視圖 session 快取（已合併多資料夾）；_tab2MonthViewYear 記錄對應年份（方案A：單一變數）
    Private _lv3MailList As New List(Of MailItemInfo)()             ' by Gemini, 2026/04/10: Tab3 顯示資料庫 (虛擬模式核心)


    Private Structure ProgressReport
        ' by Gemini, 2026/04/02: 統一進度回報結構，用於 IProgress(Of T)
        Dim CurrentCount As Integer         ' 目前完成數 (郵件數、資料夾數或位元組)
        Dim TotalCount As Integer           ' 總數 (分母)
        Dim Message As String               ' 顯示在狀態列的文字
        Dim IsIndeterminate As Boolean      ' 是否為不確定的進度 (跑馬燈模式)
    End Structure
    Private Class FolderBfsEntry
        ' 候選待掃瞄剪枝的資料夾結構
        Public Folder As Outlook.Folder
        Public ParentIndex As Integer       ' -1 = rootFolder；>= 0 = 父節點在 allEntries 的索引
        Public DirectMailCount As Integer   ' 本層郵件數 (不含子孫) ，由 Layer3 填入
        Public TotalMailCount As Integer    ' 含子孫郵件總數，Layer2 底部向上彙總後填入
        Public TotalSubCount As Integer     ' 含子孫資料夾總數，Layer2 底部向上彙總後填入
        Public IsFromCache As Boolean       ' True = TotalMailCount/TotalSubCount 從快取取得，子樹已剪枝
    End Class

#End Region

#Region "■ 04 Tab1: 資料夾統計 — 重構後程式碼 v5 (最終版) ==="
    ' ==============================================================
    '   Layer1  TreeView1_AfterSelect         UI 事件層：序號防護 + 批次寫 ListView
    '   Layer2  ComputeFolderStatsAsync       流程協調層，拆成五個子函數：
    '             - BuildBfsFolderTree        BFS + 快取剪枝（2026-04-08 加入 DB lazy，不驗 snapshot）
    '             - FetchDirectMailCountsAsync呼叫 GetCachedMailCount（有記憶體+DB lazy+COM 三層）
    '             - SummarizeSubTreeBottomUp  純記憶體底部向上加總
    '             - UpdateFolderStatsCache    寫入 L2.5 快取字典
    '             - GetBfsResult              提取 root + 直屬子資料夾
    '   Layer3  GetMailCount / GetFolderCount 等 COM 底層
    '
    ' ── 版本演進摘要 ──────────────────────────────────────────────
    '
    '   原始版  循序 Await GetInfoForListview × N，各自等遞迴完成後才輪下一個
    '           A. 用 Task.Run 包 COM (STA 違規) + B. s4Task.Result 潛在 deadlock
    '           cache: 0.10~0.19s
    '
    '   v1      BFS 一次展開整棵子樹，GetMailCount 循序讀 PR_CONTENT_COUNT
    '           底部向上彙總後一次寫快取，之後點選子資料夾直接命中，架構最乾淨，
    '           但有 bug: root 快取命中時不展開子資料夾 → 第二次點選 ListView 只顯示 root 自身
    '           cache: 0.01s (最快，因為完全不碰 thread pool)
    '
    '   v2      Task.WhenAll 同時發起 N 個子資料夾的計算 (並行的並行)
    '           修掉 s4Task.Result deadlock, 1st read 明顯變快；
    '           但 cache 仍有 40 次 Task.Run dispatch overhead
    '           cache: 0.04~0.09s (因 Task.Run overhead 限制)
    '
    '   v3      BFS + Task.WhenAll 試圖合併 v1 + v2 優點
    '           但 ComputeFolderDisplayList 在 UI 執行緒循序走整棵子樹 → 更慢
    '   v3fix   修正 v3 過深遍歷問題，ComputeFolderDisplayList 只收 depth=0/1
    '           效能介於 v1 和 v2 之間，但仍有 Task.Run overhead
    '           cache: 0.05~0.08s
    '
    '   v4      v1 的 BFS 架構 + 一行 bug fix: root 永遠展開直屬子資料夾
    '           保留 v1 的所有效能優勢，同時修正第二次點選只顯示 root 的問題
    '           不引入 Task.WhenAll (實測 sequential BFS 比 parallel of parallel 快)
    '           cache: 0.01s (應當與 v1 相同)
    '
    '   v5 (本版)
    '           2026/04/04 by Gemini: 大幅重構 ComputeFolderStatsAsync，
    '           依「單一職責原則」拆分為五個子函數，確保各步驟隔離互不干擾。
    '
    ' ── 為什麼 v4 不用 Task.WhenAll？─────────────────────────────
    '
    '   v2/v3fix 的「並行的並行」看起來應該更快，但實測反而輸給 v1，
    '   原因: PST 的 PR_CONTENT_COUNT 讀取是 COM overhead 主導 (不是 I/O bottleneck)
    '   v1 的 BFS sequential: N 個資料夾 × 1 PR_CONTENT_COUNT call = O(N)，無其他 overhead
    '   v2/v3fix 的 Task.WhenAll: 20 子資料夾 × 2 Task.Run = 40 次 thread pool dispatch
    '            每次 dispatch ~1~2ms，40 次 = 40~80ms → 這就是 cache 0.05s 的來源
    '
    '   PST 是單一檔案，並行讀取可能造成 I/O 競爭，在慢速 HDD 上優勢也有限
    '   → v1 的 sequential BFS 在此場景下已是最優，不需要 Task.WhenAll  
    '   todo: 但我還是想要再嚐試看看, 我覺得上次測試結果應該不是這個原因
    '
    ' ── 分層架構 ──────────────────────────────────────────────────
    '   Layer1  TreeView1_AfterSelect   UI 事件層
    '       取得選中資料夾 → 呼叫 Layer2 → 批次更新 ListView1
    '       規則: 不做計算，不直接操作 COM，只傳達意圖與呈現結果
    '
    '   Layer2  ComputeFolderStatsAsync 流程協調層 (核心)
    '       BFS 展開整棵子樹 (root 永遠展開直屬子，其餘節點依快取決定)
    '       → 呼叫 Layer3 讀每個節點的直接郵件數
    '       → 底部向上彙總 (O(N)，無遞迴 stack overflow 風險)
    '       → 一次性寫快取 (整棵子樹預讀)
    '       → 回傳 root + 直屬子資料夾清單供 Layer1 顯示
    '       回呼 onProgress 讓 Layer1 更新進度，Layer2 自身不碰任何 UI 控制項
    '
    '   Layer3  GetMailCount            COM 資料層
    '       只讀單一資料夾的 PR_CONTENT_COUNT (本層郵件數，不含子孫)
    '       不遞迴，不展開子資料夾，最小化 COM 呼叫量
    '
    ' ── 快取策略 ──────────────────────────────────────────────────
    '   mailCountCache   → TotalMailCount (含子孫郵件總數)   Layer2 底部向上彙總後寫入，TryAdd 不覆蓋既有值
    '   folderCountCache → TotalSubCount (含子孫資料夾總數)  Layer2 底部向上彙總後寫入，TryAdd 不覆蓋既有值
    '   folderSizeCache  → 資料夾大小 (Lazy，由 ColumnClick / 右鍵觸發計算)
    '   folderTreeCache  → 子資料夾排序清單 (GetSortedSubFolders 負責維護)
    ' ─────────────────────────────────────────────────────────────
    ' FolderBfsEntry: BFS 過程中每個資料夾節點的容器
    ' 貫穿 Layer2 的所有步驟 (BFS 展開 → Layer3 讀取 → 底部向上彙總 → 快取寫入 → 回傳清單)
    ' ─────────────────────────────────────────────────────────────
    '   快取命中剪枝規則:
    '     root (BFS 起點)   → 永遠展開直屬子資料夾 (v4 bug fix 的核心)
    '     非 root 節點      → mailCountCache + folderCountCache 都命中 → IsFromCache=True → 不再往下展開
    '
    ' ── 使用說明 ──────────────────────────────────────────────────
    '   【替換以下函數】
    '     - TreeView1_AfterSelect   → 本檔 Layer1 取代
    '     - GetInfoForListview      → 由 Layer2/Layer3 取代，舊函數可刪除
    '     - GetFolderSizeLegacy     → 本檔修正版取代 (移除 Task.Run 包 COM)
    '
    '   【完全不動的函數】
    '     - GetMailCountByMAPINew    保留 (GetFolderSizeLegacy exception path 仍呼叫)
    '     - GetTotalFolderCountAsync 保留 (不再由 Tab1 主流程呼叫，但其他地方可能用到)
    '     - GetSortedSubFolders      不改 (Layer2 BFS 直接呼叫)
    '     - GetFolderByName, FindNodeByName (右鍵、雙擊功能用)
    '     - ListView1_ColumnClick, ComputeFolderSize, EnterFolderMenuItem
    '     - GetFolderSizeOld (問題資料夾的 fallback，新版 GetFolderSizeLegacy 仍呼叫)
    ' ==============================================================
#Region "  ├ Layer1 UI事件層"
    Private Async Sub SimTree1_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles SimTree1.AfterSelect
        ' ==============================================================
        ' === Layer 1 (UI 事件層) — SimTree1 多選版 ===
        ' 2026/04/13 by Simon/Claude: Tab1 由原生 TreeView1 升級為 SimTree1（多選控制項）
        '
        ' 職責: 讀取 SimTree1.SelectedNodes（多選清單）→ 對每個選中的資料夾呼叫 Layer2
        '       → 組裝「群組標題行 + 直屬子資料夾行」清單 → 批次寫入 ListView1
        '
        ' 顯示格式（統一，不論單選或多選）:
        '   ▸ 資料夾名稱（群組標題行，粗體淡藍底，含該資料夾的完整統計數字）
        '   - 子資料夾A
        '   - 子資料夾B  ...
        '   [多選時底部追加合計列]
        '
        ' Tag 結構（ValueTuple）:
        '   群組標題行 & 合計列 → Tag = Nothing（EnterFolder / ComputeSize 看到 Nothing 直接跳過）
        '   一般子資料夾行     → Tag = (SubFolder:=Outlook.Folder, ParentNode:=TreeNode)
        '                         ComputeSize 從 .SubFolder 取資料夾；EnterSelectedFolder 從 .ParentNode 找節點
        ' ==============================================================
        _dbg("開始") : Dim sw As New Stopwatch : sw.Start()

        Dim selectedNodes As List(Of TreeNode) = SimTree1.SelectedNodes
        If selectedNodes.Count = 0 Then Return

        ' ── 父子去重 (2026/04/13 by Simon/Claude) ──────────────────────────
        ' 問題: 使用者同時選父資料夾與其子孫節點時，父的 TotalMailCount 已含子孫，
        '       若再對子孫跑一次 BFS 則數字重複計算。
        ' 解法: 用 HashSet 快速查找，若某節點的任一祖先也在選中清單裡，就跳過該節點。
        '       只保留「沒有任何祖先被選中」的節點，確保每封郵件只被計算一次。
        Dim selectedSet As New HashSet(Of TreeNode)(selectedNodes)
        Dim dedupedNodes As New List(Of TreeNode)
        For Each node As TreeNode In selectedNodes
            Dim isDescendantOfSelected As Boolean = False
            Dim ancestor As TreeNode = node.Parent
            While ancestor IsNot Nothing
                If selectedSet.Contains(ancestor) Then isDescendantOfSelected = True : Exit While
                ancestor = ancestor.Parent
            End While
            If Not isDescendantOfSelected Then dedupedNodes.Add(node)
        Next
        Dim skippedCount As Integer = selectedNodes.Count - dedupedNodes.Count
        If skippedCount > 0 Then _dbg(" ├ 父子去重", $"移除 {skippedCount} 個子孫節點，實際處理 {dedupedNodes.Count} 個")
        ' ────────────────────────────────────────────────────────────────────

        Dim mySeq As Integer = System.Threading.Interlocked.Increment(_tab1SelectSeq)
        Dim cToken As CancellationToken = OkayNowYouHaveToken()

        _isUserBusy = True : Cursor = Cursors.WaitCursor : ListView1.Items.Clear()
        ProgressBar1.Text = "" : ProgressBar2.Text = ""

        Try
            Dim allItems As New List(Of ListViewItem)
            Dim grandTotalMail As Long = 0
            Dim grandTotalFolders As Integer = 0
            Dim multiMode As Boolean = dedupedNodes.Count > 1  ' 以去重後數量判斷

            For Each node As TreeNode In dedupedNodes
                Dim folder As Outlook.Folder = TryCast(node.Tag, Outlook.Folder)
                If folder Is Nothing Then Continue For

                ' Layer2: BFS 展開整棵子樹，快取命中剪枝，底部向上彙總
                Dim progressIndicator = New Progress(Of ProgressReport)(Sub(p) ProgressBar2.Text = p.Message)
                Dim rows As List(Of FolderBfsEntry) = Await ComputeFolderStatsAsync(folder, progressIndicator, cToken)

                ' 序號機制配對（在每個 Await 點後都要檢查，多選時任一中途切換就整批放棄）
                If _tab1SelectSeq <> mySeq Then Return

                If rows.Count = 0 Then Continue For

                ' rows(0) = root（選中的資料夾本身），統計已彙總到 TotalMailCount / TotalSubCount
                ' 用 root 建群組標題行；rows(1..) = 直屬子資料夾，逐一建資料列
                allItems.Add(BuildGroupHeaderItem_Tab1(rows(0), node))
                For i As Integer = 1 To rows.Count - 1
                    allItems.Add(BuildListViewItem_Tab1(rows(i), node))
                Next

                grandTotalMail += rows(0).TotalMailCount
                grandTotalFolders += rows(0).TotalSubCount
            Next

            ' 多選時底部追加跨資料夾合計列（用去重後數量）
            If multiMode Then allItems.Add(BuildSumRowItem_Tab1(dedupedNodes.Count, grandTotalFolders, grandTotalMail))

            ListView1.BeginUpdate()
            ListView1.Items.Clear()
            _lastHoveredListItem = Nothing   ' 2026/04/14 fix: 重建清單前清掉 stale 參照，避免第一次 hover 閃動
            ListView1.Items.AddRange(allItems.ToArray())
            ListView1.EndUpdate()

            ' 狀態列：多選顯示跨選資料夾合計；單選由 GetBfsResult 的 progress callback 填好了
            If multiMode Then ProgressBar2.Text = $"統計完成: 共選取 {dedupedNodes.Count} 個資料夾，合計 {grandTotalMail:###,###,##0} 封郵件。"

        Catch ex As OperationCanceledException
            _dbg("結束", "ESC 中斷") : ProgressBar1.Text = "已中斷。" : Return
        Catch ex As System.Exception
            _dbg("錯誤", ex.Message)
        End Try

        sw.Stop()
        ProgressBar1.Text = "統計花費 " & sw.Elapsed.TotalSeconds.ToString("0.00") & " 秒。"
        Cursor = Cursors.Default : _isUserBusy = False
        SimTree1.Enabled = True : SimTree1.Focus()
        _dbg("結束")

    End Sub
    Private Async Sub TreeView1_AfterSelect(sender As Object, e As TreeViewEventArgs)
        ' ★ 2026/04/13 by Simon/Claude: Tab1 升級為 SimTree1 後，TreeView1_AfterSelect 已不再是事件處理器
        '   （移除 Handles TreeView1.AfterSelect，SimTree1 完全接管 L1 角色）
        '   保留函數本體供 TriggerAfterSelect 的舊有呼叫路徑相容，不刪除以保留版本記錄。
        '   若確認無其他呼叫端，之後版本可安全刪除。
        Await Task.CompletedTask ' 避免 Async Function 無 Await 的編譯器警告
    End Sub
    Private Sub TreeView1_MouseClick(sender As Object, e As MouseEventArgs) Handles TreeView1.MouseClick
        ' 只為了第一次啟動時自動展開第一層資料夾, 點選之後就不再自動展開了, 以免干擾使用者操作
        'If e.Button = MouseButtons.Left AndAlso _isTabInitialized(0) = True Then _isTabInitialized(0) = False
        ' debug: 這行原本的作用是要保護什麼?? 現在還需要嗎?? 
    End Sub
    Private Sub ListView1_MouseClick(sender As Object, e As MouseEventArgs) Handles ListView1.MouseClick
        ' ✅ 直接顯示已初始化好的選單，不重複建立和 AddHandler
        If e.Button = MouseButtons.Right Then _ctxListView1.Show(System.Windows.Forms.Cursor.Position)
        ' 2026/3/6: 原有程式碼每次都會新建一個ContextMenuStrip, 每次都新建一個都要重新AddHandler會造成memory leak
        ' 現在改成只在initial的時候建立一次, 之後每次右鍵點擊的時候直接Show()就好, 不用再重複建立
        ' todo: 這裡只有一行程式碼, 改去addHandler 
    End Sub
    Private Sub ListView1_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles ListView1.MouseDoubleClick
        _dbg("開始")
        If e.Button = MouseButtons.Left AndAlso e.Clicks = 2 Then           ' Double-click就跳至該資料夾統計資料顯示
            Dim selectedItem As ListViewItem = sender.GetItemAt(e.X, e.Y)   ' 獲取點選的資料夾並進入
            If selectedItem Is Nothing Then
                _dbg("結束", "未選定項目")
                Exit Sub
            Else
                EnterSelectedFolder(selectedItem)
            End If
        End If
        _dbg("結束")

    End Sub
#End Region
#Region "  ├ Layer2 流程協調層"
    Private Async Function ComputeFolderStatsAsync(rootFolder As Outlook.Folder, progress As IProgress(Of ProgressReport), cToken As CancellationToken) As Task(Of List(Of FolderBfsEntry))
        ' ==============================================================
        ' === Layer 2 (流程協調層) ===
        ' 職責: BFS 廣度優先搜索，展開整棵子樹，管理快取剪枝，驅動 Layer3，底部向上彙總，回傳顯示清單
        ' 
        ' 2026/04/04 by Gemini 重構紀錄:
        ' v5: 原有的百行巨型函數已被依「單一職責原則」拆分為五個子函數，確保各步驟隔離互不干擾。
        '
        ' 拆分後的五個步驟 (Steps):
        '   Step 1. BuildBfsFolderTree          : BFS 展開，收集整棵子樹的所有節點；若快取命中(非root)則剪枝。
        '   Step 2. FetchDirectMailCountsAsync  : 對未快取節點逐一呼叫 GetCachedMailCount() 取本層郵件數。
        '                                         處理 progress 報告並支援 _cancelRequested 中斷。
        '   Step 3. SummarizeSubTreeBottomUp    : 利用 BFS「父索引 < 子索引」特性，從陣列尾端往前掃一次完成加總。
        '   Step 4. UpdateFolderStatsCache      : 將最新結果寫入 Layer2.5 的 _cacheMailCountAll 等字典。
        '   Step 5. GetBfsResult                : 從陣列中挑出 root 與直屬子資料夾 (ParentIndex=0) 並補讀快取。
        '
        ' 架構與效能考量:
        '   - allEntries 是 Reference Type，在此作為狀態載體在各子函數間傳遞，避免不必要的陣列複製。
        '   - 為防 BFS 索引錯亂，以 IReadOnlyList 宣告參數，確保子函數不可改變 allEntries 長度或顛倒內部順序。
        '   - v4 bug fix: BFS 剪枝規則為「root 永遠展開直屬子資料夾，不論快取」。
        ' ==============================================================
        _dbg(" ├ 開始", rootFolder.Name)

        ' ── Step 1: 負責展開樹狀結構與初步快取剪枝 (by Gemini, 2026/04/05 改為非同步以提升響應)
        ' pending B: 目前第二耗時, 占30~35%
        Dim allEntries As List(Of FolderBfsEntry) = Await BuildBfsFolderTree(rootFolder, cToken)

        ' ── Step 2: 負責與 COM 溝通，取得基本數據 
        ' pending A. 目前第一耗時, 占55~65%
        Await FetchDirectMailCountsAsync(allEntries, progress, cToken)

        ' ── Step 3 & 4: 純記憶體運算與快取更新
        SummarizeSubTreeBottomUp(allEntries)
        UpdateFolderStatsCache(allEntries)

        ' ── Step 5: 提取 UI 所需的結果並回報最終進度
        _dbg(" ├ 結束", rootFolder.Name)
        Return GetBfsResult(allEntries, progress)

    End Function
    Private Sub EnterSelectedFolder(selectedItem As ListViewItem)
        ' ★ 核心修正 (2026-03-20) by Claude.ai:
        ' TreeView 使用 lazy loading: 子節點未展開時 .Nodes 只有 ":::" 佔位節點。
        ' 因此在搜尋目標節點前，必須先確保 SelectedNode 已展開，讓真實子節點載入。
        '
        ' 展開 SelectedNode 只觸發一次 BeforeExpand → LoadSubFolderToTreeView，這是正確且必要的。
        ' 問題出在舊版後續的 TreeView1.SelectedNode = foundNode:
        '       WinForms setter 內部呼叫 Win32 TVM_ENSUREVISIBLE，TVM_ENSUREVISIBLE 沿祖先鏈逐一 Expand()，
        '       每個都觸發 BeforeExpand → LoadSubFolderToTreeView，即使 foundNode 是直屬子節點也如此 (Win32 層不知道它已可見) 。
        '
        ' 修正方案:
        '   ① SelectedNode.Expand()         → 只展開父節點，載入真實子節點 (一次 BeforeExpand，正確)
        '   ② 在真實子節點裡找 foundNode     → 不遞迴 (FindNodeByName 每個節點都 Expand()，已知錯誤)
        '   ③ foundNode.Tag判斷folder.count → 確認目標資料夾有子資料夾才進入
        '   ④ SendMessage TVM_SELECTITEM    → 直接在 Win32 層選取 foundNode，
        '      繞過 WinForms setter 的 EnsureVisible 路徑，不再展開任何額外節點。
        '      Win32 TVM_SELECTITEM 仍會發出 TVN_SELCHANGED，
        '      WinForms 收到後自動觸發 SimTree1_AfterSelect，行為與原本完全一致。
        '
        ' ★ 2026/04/13 by Simon/Claude: SimTree1 升級後，Tag 改為 ValueTuple:
        '   群組標題行 & 合計列 → Tag = Nothing → 直接 Return，不進入
        '   一般子資料夾行     → Tag = (SubFolder, ParentNode) → 從 ParentNode 直接取得父節點，不再依賴 TreeView1.SelectedNode
        _dbg(" ├ 開始", selectedItem.SubItems(0).Text)

        ' 群組標題行 / 合計列的 Tag 是 Nothing，不可進入
        If selectedItem.Tag Is Nothing Then Return

        ' 從 ValueTuple Tag 取得子資料夾與其父 TreeNode
        Dim t As (SubFolder As Outlook.Folder, ParentNode As TreeNode) =
            DirectCast(selectedItem.Tag, (SubFolder As Outlook.Folder, ParentNode As TreeNode))
        Dim parentNode As TreeNode = t.ParentNode
        If parentNode Is Nothing Then Return

        ' ① 確保父節點已展開 (若只有 ":::" 則展開一次，載入真實子節點)
        parentNode.Expand()

        ' ② 在直屬子節點裡找目標 (不遞迴，不呼叫任何 Expand)
        Dim subject As String = selectedItem.SubItems(0).Text.Replace(" - ", "")
        Dim foundNode As TreeNode = Nothing
        For Each node As TreeNode In parentNode.Nodes
            If node.Text.Replace(" - ", "") = subject Then
                foundNode = node : Exit For
            End If
        Next
        If foundNode Is Nothing Then Return

        ' ③ 確認目標資料夾有子資料夾才進入
        Dim targetFolder As Outlook.Folder = TryCast(foundNode.Tag, Outlook.Folder)
        If targetFolder Is Nothing OrElse GetCachedFolderCount(targetFolder) = 0 Then Return
        foundNode.EnsureVisible()

        ' ④ 用 Win32 直接選取，繞過 WinForms SelectedNode setter 的 EnsureVisible 路徑
        '    Win32 TVM_SELECTITEM 仍會發出 TVN_SELCHANGED → SimTree1.AfterSelect → SimTree1_AfterSelect
        SendMessage(SimTree1.Handle, TVM_SELECTITEM, New IntPtr(TVGN_CARET), foundNode.Handle)
        ListView1.Focus()
        If ListView1.Items.Count > 0 Then ListView1.Items(0).Selected = True
        _dbg("結束")

    End Sub
    Private Async Sub ComputeFolderSize(sender As Object, e As EventArgs)
        _isUserBusy = True
        _dbg(" ├ 開始", $"選取項目數: {ListView1.SelectedItems.Count}") ' by Gemini, 2026/04/10: 調整縮排層級為 Level 1

        Try
            Dim stopwatch As New Stopwatch : stopwatch.Start()
            Dim selectedItems As ListView.SelectedListViewItemCollection = ListView1.SelectedItems  ' 如果有選中項目, 獲取所選中的項目
            If selectedItems.Count > 0 Then
                Dim cToken As CancellationToken = OkayNowYouHaveToken() ' ✅ 取得新 Token
                For Each s As ListViewItem In selectedItems
                    'If s.Index = 0 Then Continue For ' 若選中本體目錄則跳過 (之前統計速度很慢的時候, 怕計算量太大跑太久)
                    If s.SubItems.Count > 4 Then s.SubItems(4).Text = "計算中..." Else s.SubItems.Add("計算中...")
                    ' 提高反應速度, 先占位 (如果已經有FolderSize的子項目就先把它改成「計算中...」, 如果還沒有就先加一個占位用的子項目)
                Next

                Dim swThrottle As New Stopwatch : swThrottle.Start() ' by Gemini, 2026/04/11
                Dim totalCount As Integer = selectedItems.Count
                Dim processedCount As Integer = 0

                For Each s As ListViewItem In selectedItems
                    'If s.Index = 0 Then Continue For ' 一樣, 若選中本體目錄則跳過 (之前統計速度還很慢的時候, 怕計算量太大跑太久)
                    ' 2026/04/13 by Simon/Claude: Tag 升級為 ValueTuple (SubFolder, ParentNode)；群組標題行 / 合計列 Tag=Nothing，直接跳過
                    If s.Tag Is Nothing Then Continue For
                    Dim t As (SubFolder As Outlook.Folder, ParentNode As TreeNode) =
                        DirectCast(s.Tag, (SubFolder As Outlook.Folder, ParentNode As TreeNode))
                    Dim folder As Outlook.Folder = t.SubFolder
                    If folder Is Nothing Then Continue For

                    Dim folderSize As Long = Await GetCachedFolderSizeAllAsync(folder)  ' 2026/3/29 by Gemini: 改為存取 Layer2.5 快取代理，第二次點擊同一資料夾直接命中快取

                    Dim strFolderSize As String
                    If folderSize < 0 Then strFolderSize = "計算失敗" Else strFolderSize = (folderSize / 1024).ToString("###,###,###,##0 KB")
                    If s.SubItems.Count > 4 Then s.SubItems(4).Text = strFolderSize Else s.SubItems.Add(strFolderSize)

                    processedCount += 1
                    ' by Gemini, 2026/04/11: 100ms 節流回報進度與檢查 ESC 中斷
                    If swThrottle.ElapsedMilliseconds >= 100 OrElse processedCount = totalCount Then
                        ProgressBar2.Text = $"正在計算資料夾大小: {processedCount} / {totalCount} ({folder.Name})"
                        swThrottle.Restart()
                        Await Task.Delay(1, cToken)
                    End If
                Next
            End If

            ProgressBar2.Text = "統計資料夾大小花費了 " & stopwatch.Elapsed.TotalSeconds.ToString("0.00") & " 秒。"
        Catch ex As OperationCanceledException
            ' by Gemini, 2026/04/11: 捕捉 ESC 中斷異常，優雅顯示訊息而不拋出錯誤視窗
            ProgressBar2.Text = "計算已由使用者中斷。"
            _dbg(" ├ 中斷", "ComputeFolderSize 已中斷")
        Catch ex As System.Exception
            ProgressBar2.Text = "發生錯誤: " & ex.Message
            _dbg(" ├ 錯誤", ex.Message)
        Finally
            _isUserBusy = False
            _dbg("結束")
        End Try

    End Sub
#End Region
#Region "  └ 輔助函數"
    ' 以下為 ComputeFolderStatsAsync 專用的拆分子函數 (Steps 1~5)
    Private Async Function BuildBfsFolderTree(rootFolder As Outlook.Folder, cToken As CancellationToken) As Task(Of List(Of FolderBfsEntry))
        ' 負責: 維護 Queue 執行 BFS，根據 Layer2.5 快取字典決定是否剪枝。
        ' 產出: 所有走訪過的資料夾陣列，每個元素皆紀錄了其 ParentIndex。
        _dbg(" ├ 開始", rootFolder.Name)

        Dim allEntries As New List(Of FolderBfsEntry)
        Dim queue As New Queue(Of (folderObj As Outlook.Folder, parentIdx As Integer))
        queue.Enqueue((rootFolder, -1))

        ' by Gemini, 2026/04/05: 每 100ms 主動讓出執行緒並檢查中斷，兼顧效能與靈敏度
        Dim swThrottle As New Stopwatch() : swThrottle.Start()
        Try
            Do While queue.Count > 0
                Dim curr = queue.Dequeue()
                Dim entry As New FolderBfsEntry With {.Folder = curr.folderObj, .ParentIndex = curr.parentIdx, .IsFromCache = False}
                Dim myIdx As Integer = allEntries.Count
                allEntries.Add(entry)

                ' 快取命中判斷: 兩個快取都有才算完整命中 (任一失效都重新計算，確保一致性)
                Dim cachedMail As Integer, cachedSub As Integer
                Dim fPath As String = curr.folderObj.FolderPath
                Dim isHit As Boolean = False
                If _cacheMailCountAll.TryGetValue(fPath, cachedMail) AndAlso _cacheFolderCountAll.TryGetValue(fPath, cachedSub) Then
                    isHit = True    ' ① 記憶體命中
                Else
                    ' ② DB lazy load：不驗 snapshot（剪枝決策可接受略舊資料，不需要額外 COM call）
                    ' 只在 mca 和 fca 兩個欄位都有值（非 NULL）時才算命中，確保顯示正確
                    Dim row = DbGetFolderStats(fPath)
                    If row IsNot Nothing AndAlso row.mca >= 0 AndAlso row.fca >= 0 Then
                        cachedMail = row.mca : cachedSub = row.fca
                        FillFolderCacheFromDbRow(fPath, row)   ' 一次填滿所有欄位到記憶體快取
                        isHit = True    ' ② DB 命中
                    End If
                End If

                If isHit Then
                    entry.TotalMailCount = cachedMail
                    entry.TotalSubCount = cachedSub
                    entry.IsFromCache = True

                    ' ★ v4 bug fix: root (parentIdx=-1) 即使快取命中，也要繼續展開直屬子資料夾
                    ' 只有非 root 節點才允許剪枝
                    If curr.parentIdx <> -1 Then Continue Do
                End If

                ' 未命中，或是 root (不論有無快取) → 展開直屬子資料夾
                For Each subFolder As Outlook.Folder In GetSortedSubFolders(curr.folderObj)
                    queue.Enqueue((subFolder, myIdx))
                Next

                ' by Gemini, 2026/04/11: 補回節流與中斷檢查
                If swThrottle.ElapsedMilliseconds >= 100 Then
                    swThrottle.Restart()
                    Await Task.Delay(1, cToken)
                End If
            Loop
        Catch ex As OperationCanceledException
            ' 2026/04/11 by Claude: 改為 re-throw，確保不完整的 BFS 樹不會繼續傳入
            ' UpdateFolderStatsCache，避免錯誤的中途統計結果汙染快取。
            ' (原本 catch 後繼續 Return allEntries，導致上層 ComputeFolderStatsAsync
            '  看不到中斷，仍執行 SummarizeSubTreeBottomUp + UpdateFolderStatsCache)
            _dbg("    ├ 中斷", $"BuildBfsFolderTree 已由使用者中斷")
            Throw
        End Try

        Dim total As Integer = allEntries.Count
        _dbg("    ├ 結束", $"節點總計: {total} (含快取命中剪枝)")
        Return allEntries

    End Function
    Private Async Function FetchDirectMailCountsAsync(allEntries As IReadOnlyList(Of FolderBfsEntry), progress As IProgress(Of ProgressReport), cToken As CancellationToken) As Task
        ' 負責: 對未快取節點打 COM (呼叫 GetCachedMailCount)，並負責 UI 節流 (Task.Yield) 與 ESC 中斷檢查。
        ' 2026/04/11 by Claude: 回傳值從 Task(Of Boolean) 改為 Task，原本的 Return True/False 均改為 re-throw。
        '   理由: 呼叫端 Await FetchDirectMailCountsAsync(...) 完全丟棄了 Task(Of Boolean) 的回傳值，
        '         等同 Return True 無效，ESC 後上層照樣執行 UpdateFolderStatsCache 污染快取。
        '         改為 Throw 後，OperationCanceledException 直接傳到 TreeView1_AfterSelect 的 catch 攔截。

        _dbg("    ├ 開始") ' by Gemini, 2026/04/10: 調整縮排層級為 Level 2
        Dim total As Integer = allEntries.Count, processed As Integer = 0
        Dim swThrottle As New Stopwatch() : swThrottle.Start()

        Try
            For fd As Integer = 0 To total - 1
                Dim entry As FolderBfsEntry = allEntries(fd)
                If Not entry.IsFromCache Then
                    entry.DirectMailCount = GetCachedMailCount(entry.Folder) ' 直接呼叫 Proxy 代理層
                    entry.TotalMailCount = entry.DirectMailCount             ' 初始值 = 本層，後面底部向上累加子孫
                    entry.TotalSubCount = 0                                  ' 初始為 0，後面累加子孫資料夾數
                End If
                processed += 1

                ' by Gemini, 2026/04/11: 補回節流與中斷檢查
                If swThrottle.ElapsedMilliseconds >= 100 Then
                    progress?.Report(New ProgressReport With {.CurrentCount = processed, .TotalCount = total,
                                                              .Message = $"正在統計郵件數: {processed} / {total} 個資料夾..."})
                    swThrottle.Restart()
                    Await Task.Delay(1, cToken)
                End If
            Next
        Catch ex As OperationCanceledException
            ' 2026/04/11 by Claude: 改為 re-throw，不再 Return True（上層丟棄了回傳值，Return True 等於無效）
            _dbg("    ├ 中斷", "FetchDirectMailCountsAsync 已由使用者中斷")
            Throw
        End Try

        _dbg(" ├ 結束", $"共讀取 {total} 個節點（非快取）") ' by Gemini, 2026/04/04: Issue 2 補上結束

    End Function
    Private Sub SummarizeSubTreeBottomUp(allEntries As IReadOnlyList(Of FolderBfsEntry))
        ' 負責: 底部向上彙總。利用 BFS 父節點索引必小於子節點的特性，反向遍歷即可。
        ' 備註: 純記憶體 O(N) 運算且無 COM 呼叫，無 StackOverflow 風險。
        For i As Integer = allEntries.Count - 1 To 1 Step -1
            Dim child As FolderBfsEntry = allEntries(i)
            Dim parent As FolderBfsEntry = allEntries(child.ParentIndex)

            ' ★ 重要: parent.IsFromCache=True 代表該節點的 TotalMailCount 已是含子孫的正確快取值
            ' 此時不能再疊加 child.TotalMailCount，只有 parent.IsFromCache=False 的節點才需累加
            If Not parent.IsFromCache Then
                parent.TotalMailCount += child.TotalMailCount
                parent.TotalSubCount += child.TotalSubCount + 1     ' +1 = child 這個資料夾本身也計入
            End If
        Next
    End Sub
    Private Sub UpdateFolderStatsCache(allEntries As IReadOnlyList(Of FolderBfsEntry))
        ' 負責: 將新計算的彙總結果寫入 Layer2.5 快取 
        ' 備註: TryAdd 不覆蓋既有值，避免污染快取
        For Each entry As FolderBfsEntry In allEntries
            If Not entry.IsFromCache Then
                Dim fPath As String = entry.Folder.FolderPath
                _cacheMailCountAll.TryAdd(fPath, entry.TotalMailCount)
                _cacheFolderCountAll.TryAdd(fPath, entry.TotalSubCount)
            End If
        Next
    End Sub
    Private Function GetBfsResult(allEntries As IReadOnlyList(Of FolderBfsEntry), progress As IProgress(Of ProgressReport)) As List(Of FolderBfsEntry)
        ' 負責: 找出 root 與直屬子資料夾 (ParentIndex=0) 組裝成新的 UI 呈現清單，並在最後發布結束進度。
        _dbg("開始")

        Dim result As New List(Of FolderBfsEntry)
        result.Add(allEntries(0))   ' index 0 = rootFolder 本身

        For i As Integer = 1 To allEntries.Count - 1
            Dim entry As FolderBfsEntry = allEntries(i)
            If entry.ParentIndex = 0 Then
                ' 若直屬子資料夾快取命中，補讀一下其本層郵件 (DirectMailCount)
                If entry.IsFromCache Then entry.DirectMailCount = GetCachedMailCount(entry.Folder)
                result.Add(entry)
            End If
        Next

        ' 若 root 自身快取命中，也補讀其本層郵件數
        If allEntries(0).IsFromCache Then allEntries(0).DirectMailCount = GetCachedMailCount(allEntries(0).Folder)

        Dim totalMail As Long = allEntries(0).TotalMailCount
        Dim totalFolder As Integer = allEntries(0).TotalSubCount
        progress?.Report(New ProgressReport With {.CurrentCount = allEntries.Count, .TotalCount = allEntries.Count,
                                                  .Message = $"統計完成: 共 {totalFolder} 個子資料夾，{totalMail:###,###,##0} 封郵件。"})

        _dbg("    ├ 結束", $"回傳 {result.Count} 列 (1 root + {result.Count - 1} 直屬子資料夾)") ' by Gemini, 2026/04/10
        Return result

    End Function
    Private Function BuildGroupHeaderItem_Tab1(rootEntry As FolderBfsEntry, parentNode As TreeNode) As ListViewItem
        ' 群組標題行：取代舊版的 isRoot=True 第一列
        ' 顯示選中資料夾本身的完整統計（TotalMailCount / TotalSubCount）
        ' 欄位: ▸ 資料夾名稱 / 郵件數量 / 資料夾數量 / 郵件總計 / 大小（5欄回归）
        ' Tag = Nothing：EnterSelectedFolder 與 ComputeFolderSize 看到 Nothing 直接跳過
        ' 2026/04/13 by Simon/Claude: B方案 — 統一格式，單選與多選皆顯示群組標題行
        ' 2026/04/13 v2: 移除「所屬父資料夾」欄（該欄內容永遠等於標題行本身，元余）
        Dim sizeStr As String = "- "
        Dim sizeVal As Long
        If _cacheFolderSizeAll.TryGetValue(rootEntry.Folder.FolderPath, sizeVal) AndAlso sizeVal > 0 Then
            sizeStr = (sizeVal \ 1024L).ToString("###,###,##0") & "KB "
        End If

        Dim directMailStr As String = rootEntry.DirectMailCount.ToString("###,###,##0") & " "
        Dim totalSubStr As String = rootEntry.TotalSubCount.ToString("###,###,##0") & " "
        Dim totalMailStr As String = rootEntry.TotalMailCount.ToString("###,###,##0") & " "

        ' 欄位順序: 名稱 / 郵件數量 / 資料夾數量 / 郵件總計 / 大小
        Dim lvi As New ListViewItem({"▸ " & rootEntry.Folder.Name, directMailStr, totalSubStr, totalMailStr, sizeStr})
        lvi.Font = New Font(ListView1.Font, _fontBold)
        lvi.BackColor = SystemColors.GradientInactiveCaption
        lvi.Tag = Nothing
        Return lvi

    End Function
    Private Function BuildListViewItem_Tab1(entry As FolderBfsEntry, parentNode As TreeNode) As ListViewItem
        ' 組裝 ListView1 的單一資料列（直屬子資料夾）
        ' 欄位: 資料夾名稱 / 郵件數量 / 資料夾數量 / 郵件總計 / 大小(Lazy)
        ' 2026/04/13 by Simon/Claude: B方案升級
        '   - 移除 isRoot 參數（群組標題行已獨立為 BuildGroupHeaderItem_Tab1）
        '   - 移除 parentFolderName 參數（該欄元余，已移除）
        '   - parentNode 存入 Tag ValueTuple，供 EnterSelectedFolder 使用
        '   - Tag 改為 ValueTuple(SubFolder, ParentNode)，ComputeFolderSize 同步更新
        ' 舊版註記: by Gemini, 2026/03/31: 視覺優化重構
        '   1. 資料夾名稱：还原開頭縮排空白 (" - ")，保持整齊。
        '   2. 防止切邊：斜體時字串結尾補一格空白。

        Dim isItalicFolder As Boolean = Not IsMailFolder(entry.Folder)
        Dim displayName As String = " - " & entry.Folder.Name
        If isItalicFolder Then displayName &= " "

        ' 大小: Lazy，從快取讀；未計算過則留空，等 ColumnClick 或右鍵選單觸發計算
        Dim sizeStr As String = "- "
        Dim sizeVal As Long
        If _cacheFolderSizeAll.TryGetValue(entry.Folder.FolderPath, sizeVal) AndAlso sizeVal > 0 Then sizeStr = (sizeVal \ 1024L).ToString("###,###,##0") & "KB"
        ' 統計數字字串化 (字串結尾一律補一格空白，確保斜體與正常字體對齊且不切邊)
        Dim directMailStr As String = entry.DirectMailCount.ToString("###,###,##0") & " "
        Dim totalSubStr As String = entry.TotalSubCount.ToString("###,###,##0") & " "
        Dim totalMailStr As String = entry.TotalMailCount.ToString("###,###,##0") & " "
        If sizeStr <> "- " Then sizeStr &= " "

        ' 欄位順序: 名稱 / 郵件數量 / 資料夾數量 / 郵件總計 / 大小
        Dim lvi As New ListViewItem({displayName, directMailStr, totalSubStr, totalMailStr, sizeStr})
        ' by Gemini, 2026/03/29: 特殊顯示非郵件資料夾 (斜體 + 灰色)
        If isItalicFolder Then
            lvi.ForeColor = Color.DarkGray
            lvi.Font = New Font(ListView1.Font, _fontItalic)
        End If

        ' 2026/04/13 by Simon/Claude: Tag 改為 ValueTuple，ComputeFolderSize 與 EnterSelectedFolder 同步更新
        lvi.Tag = (SubFolder:=entry.Folder, ParentNode:=parentNode)
        Return lvi

    End Function
    Private Function BuildSumRowItem_Tab1(selectedCount As Integer, totalSub As Integer, totalMail As Long) As ListViewItem
        ' 合計列：多選模式才插入，顯示跨資料夾加總。Tag = Nothing
        ' Tag = Nothing：同群組標題行，不可進入
        ' 2026/04/13 by Simon/Claude: B方案新增，5欄格式
        Dim totalMailStr As String = totalMail.ToString("###,###,##0") & " "
        Dim totalSubStr As String = totalSub.ToString("###,###,##0") & " "
        Dim lvi As New ListViewItem({"▶ 合計（" & selectedCount & " 個資料夾）", "", totalSubStr, totalMailStr, ""})
        lvi.Font = New Font(ListView1.Font, _fontBold)
        lvi.BackColor = Color.FromArgb(220, 235, 252)
        lvi.ForeColor = Color.FromArgb(0, 70, 140)
        lvi.Tag = Nothing
        Return lvi
    End Function
    ' ── ListView1 OwnerDraw handlers (2026/04/13 by Simon/Claude) ──────────────────
    ' 問題根因: Windows 在 ListView 的 hover/select 狀態下會覆蓋自訂 BackColor，
    '   導致群組標題行的淡藍底在滑鼠移上去或點擊後消失。
    ' 解法: ListView.OwnerDraw = True，只對 Tag=Nothing 的行自訂繪製，其餘一律 DrawDefault=True。
    ' 該方式不影響一般資料列的外觀和排序等功能。
    Private Sub ListView1_DrawColumnHeader(sender As Object, e As DrawListViewColumnHeaderEventArgs) Handles ListView1.DrawColumnHeader
        e.DrawDefault = True   ' 欄位標頭用預設繪製
    End Sub
    Private Sub ListView1_DrawItem(sender As Object, e As DrawListViewItemEventArgs) Handles ListView1.DrawItem
        ' Tag=Nothing（群組標題行 / 合計列）：自己填背景色，阻止 OS 在此階段畫 hover/select 高亮
        '   → DrawSubItem 再負責各欄文字
        ' 一般列：DrawDefault=True，OS 正常繪製（選取藍、hover 灰等維持原樣）
        ' 2026/04/14 fix by Gemini 3.1 Pro: 原本 DrawDefault=True，OS 仍會在 DrawItem 階段塗高亮背景，
        '   之後 DrawSubItem 雖覆蓋，但 hover 重繪有時只觸發 DrawItem 不重觸發 DrawSubItem，
        '   導致背景色被清掉。
        '   之前嘗試在此處填色，但這會導致沒有觸發 DrawSubItem 時，被背景色覆蓋掉文字而產生閃爍。
        '   最好的解法是什麼都不做（且不設 DrawDefault=True），讓畫面保留原本 DrawSubItem 畫好的狀態。
        If e.Item.Tag Is Nothing Then
            ' 不做任何事，保持原有的畫面像素
        Else
            e.DrawDefault = True
        End If
    End Sub
    Private Sub ListView1_DrawSubItem(sender As Object, e As DrawListViewSubItemEventArgs) Handles ListView1.DrawSubItem
        ' Tag=Nothing (群組標題行 / 合計列)：自訂繪製，防止 OS hover/select 顏色覆蓋我們設定的 BackColor
        ' 其餘一般列： DrawDefault=True 不影響任何現有功能
        '
        ' 2026/04/14 對齊修正: 原本 textRect.Inflate(-3, 0) 兩側各縮 3px，
        '   導致右對齊欄位比 DrawDefault 多偏左 3px，與一般列數字不對齊。
        '   修正為：右對齊欄位直接用 e.Bounds，trailing " " 本身就是視覺間距，不額外 Inflate。
        '   第一欄（左對齊）僅加左側 3px padding 避免文字貼邊。
        If e.Item.Tag Is Nothing Then
            Using bgBrush As New SolidBrush(e.Item.BackColor)
                e.Graphics.FillRectangle(bgBrush, e.Bounds)
            End Using
            Dim textRect As Rectangle = e.Bounds
            Dim align As StringAlignment
            If e.ColumnIndex = 0 Then
                textRect.X += 3 : textRect.Width -= 3  ' 第一欄左側加 3px padding，避免文字貼欄邊
                align = StringAlignment.Near
            Else
                align = StringAlignment.Far             ' 右對齊欄位直接用 e.Bounds，trailing " " 提供間距
            End If
            ' 2026/04/14 fix by Gemini 3.1 Pro: 增加 StringFormatFlags.MeasureTrailingSpaces 確保右對齊時包含空白間距
            Dim sf As New StringFormat() With {.Alignment = align, .LineAlignment = StringAlignment.Center,
                                              .FormatFlags = StringFormatFlags.NoWrap Or StringFormatFlags.NoClip Or StringFormatFlags.MeasureTrailingSpaces,
                                              .Trimming = StringTrimming.EllipsisCharacter}
            Using fgBrush As New SolidBrush(e.Item.ForeColor)
                e.Graphics.DrawString(e.SubItem.Text, e.Item.Font, fgBrush,
                                      RectangleF.FromLTRB(textRect.Left, textRect.Top, textRect.Right, textRect.Bottom), sf)
            End Using
        Else
            e.DrawDefault = True
        End If
    End Sub
    Private Sub ListView1_ItemSelectionChanged(sender As Object, e As ListViewItemSelectionChangedEventArgs) Handles ListView1.ItemSelectionChanged
        ' 群組標題行 / 合計列不可被選取，選中即強制取消。
        ' 進一步防止 OS 在取消選取時再覆蓋一次我們的 BackColor。
        If e.Item.Tag Is Nothing AndAlso e.IsSelected Then
            e.Item.Selected = False
        End If
    End Sub
    ' ─────────────────────────────────────────────────────────────────────────────────
#End Region
#End Region

#Region "■ 05 Tab2: 依日期統計"
    ' ==============================================================
    ' 重構目標: COM/UI/流程邏輯與業務分離清晰分層，去除全域狀態，優化快取機制
    ' 1. 分層架構: 將原本混在一起的程式碼重構成三個明確的層次
    '    - Layer 1 (UI 事件層)  : 回應使用者操作，組裝參數後交給 Layer2 執行，最後把結果交給顯示函數
    '    - Layer 2 (流程協調層) : BFS 遍歷 folderList，管理快取，驅動 Layer3 計算，合併結果，回報進度
    '    - Layer 3 (COM 資料層) : 對 Outlook 發出 COM 呼叫，回傳單一資料夾的年份郵件分佈
    ' 2. 去除全域狀態: 原本的 _intTotalMailCount 和 _intProcessedCount 全域變數已改成局部變數，避免多次點選時的計數錯亂
    ' 3. 優化快取機制: 快取的 key 改為純字串 FolderPath，避免 COM 物件當 key 導致 RCW 殘留問題；快取只存單一資料夾的結果，由 Layer2 負責合併
    ' 4. 進度回報改為 callback 機制: Layer2 執行統計時，透過 onProgress callback 回報已處理的郵件數和總郵件數，Layer1 負責更新 UI 顯示，保持分層乾淨
    ' by: Claude AI (2026/3/10)
    ' ==============================================================
    ' 替換說明:
    '   以下程式碼完整取代 Tab2 相關的所有邏輯函數。
    '   請同時刪除以下舊的函數與宣告:
    '     - Private _intTotalMailCount As Integer   (全域變數宣告，已改成局部)
    '     - Private _intProcessedCount As Integer   (全域變數宣告，已改成局部)
    '     - TreeView2_AfterSelect()                 (已重寫)
    '     - SimTree2_AfterSelect()                  (已重寫，不再 commented out)
    '     - CheckSubFolder2_CheckedChanged()        (已重寫)
    '     - GetYearCountsAsync_CL()                 (已由 CollectYearCounts 取代)
    '     - CountMailByYearAsync_CLayer2()          (已由 GetYearCountsForFolderAsync 取代)
    '     - UpdateCounterProgress()                 (已改成 callback 機制，函數可刪除)
    '     - ShowProgressTab2()                      (簽章已更改，請替換)(2026/4/12 重構 v2 已刪除)
    '
    ' ==============================================================
    ' 2026/04/12 重構 v2 (render層拆分+導覽函數整合):
    '   刪除: ShowYearView, ShowMonthView, ShowResultTab2, ShowProgressTab2
    '         UpdateChart2forYearView, UpdateChart2forMonthView
    '   新增: CollectMonthCounts            ← 月份資料收集 Layer2（純計算，不碰UI）
    '         GoToYearView()                ← 共用導覽：返回年度視圖（純 render from cache，DoubleClick/KeyPress 共用）
    '         GoToMonthView(year, cToken)   ← 共用導覽：進入月份視圖（方案A cache，DoubleClick/KeyPress 共用）
    '         RenderLvYearView              ← 年度 ListView render（純UI，不計算）
    '         RenderLvMonthView             ← 月份 ListView render（純UI，不計算）
    '         RenderCtYearView              ← 年度 Chart render（純UI，不計算）
    '         RenderCtMonthView             ← 月份 Chart render（純UI，不計算）
    '   更改: Form1.vb HandleListViewKeyPress 的 ShowYearView/ShowMonthView 改呼叫 GoToYearView/GoToMonthView
    ' ==============================================================
    ' 分層架構 (更新後):
    '   Layer 1 (UI 事件層)       : SimTree2_AfterSelect, CheckSubFolder2_CheckStateChanged
    '                                ListView2_MouseDoubleClick（只兩行，委派 GoToYearView/GoToMonthView）
    '   Layer 2 (流程協調層)      : CollectYearCounts, CollectMonthCounts
    '                                RenderLvYearView, RenderCtYearView
    '                                RenderLvMonthView, RenderCtMonthView
    '                                GoToYearView, GoToMonthView
    '   Layer 3 (COM 資料層)      : GetYearCountsForFolder, GetMonthCountsForYear (Form1_Outlook.vb，不動)
    ' ==============================================================
#Region "  ├ Layer1 UI事件層"
    Private Async Sub SimTree2_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles SimTree2.AfterSelect
        ' ---------------------------------------------------------------
        ' === Layer 1: UI 事件層 ===
        ' 職責: 回應使用者操作，組裝參數後交給 Layer2 執行，最後把結果交給顯示函數
        ' 規則: 不做業務計算，不直接碰 COM，只傳達意圖
        ' ---------------------------------------------------------------
        ' SimTree2_AfterSelect: 多選模式 SimTree2 的節點點選事件, 完整替換舊版
        ' 與 TreeView2_AfterSelect 對齊，補上月份展開所需的狀態賦值
        ' 支援 Ctrl+Click 或 Shift+Click 多選，每個選定節點各自 BFS 展開後合併統計
        '
        ' by Gemini, 2026/03/29: 移除 TreeView2_AfterSelect，由 SimTree2 完全取代。
        ' ---------------------------------------------------------------
        _dbg("開始") : Dim stopwatch As New Stopwatch() : stopwatch.Start()    ' 開始計時，初始化畫面狀態
        Cursor = Cursors.WaitCursor : ProgressBar1.Text = "" : ProgressBar2.Text = ""

        ' 序號機制: 每次點選遞增；計算完成後若序號已變，代表有更新的點選，丟棄本次結果
        Dim mySeq As Integer = System.Threading.Interlocked.Increment(_tab2SelectSeq)
        Dim cToken As CancellationToken = OkayNowYouHaveToken()  ' ✅ 取得新 Token，同時取消上一次未完成的操作

        ' 取得 SimTree2 多選清單 (SelectedNodes 是 SimTree 提供的 List(Of TreeNode))
        Dim selectedNodes As List(Of TreeNode) = SimTree2.SelectedNodes
        If selectedNodes Is Nothing OrElse selectedNodes.Count = 0 Then
            _dbg("結束", "無節點被選取")
            Cursor = Cursors.Default : Return           ' 選擇節點為空，直接結束
        End If

        Dim targetFolderList =                          ' 把所有已選 TreeNode 的 Tag 轉換成 Outlook.Folder，過濾掉無效節點
            selectedNodes.Select(Function(n) TryCast(n.Tag, Outlook.Folder)).Where(Function(f) f IsNot Nothing).ToList()
        If targetFolderList.Count = 0 Then
            _dbg("結束", "所有選定節點均無效資料夾")
            Cursor = Cursors.Default : Return           ' 如果沒有任何有效的資料夾 (List.Count=0) 就直接結束
        End If

        Try ' by Claude Opus, 2026/04/11: Try 上移，包住 GetSubtreeToList 的 Await，否則 ESC 時拋出的 OperationCanceledException 沒有被捕捉
            Dim folderList = Await GetUniqueFolderList(selectedNodes, _includeSubTab2, cToken:=cToken)
            _tab2FolderList = folderList    ' ✅ 記住本次統計的資料夾清單，供 GoToMonthView (CollectMonthCounts) 使用
            _tab2IsMonthView = False        ' 切換選取時，重置視圖狀態為年度視圖

            ''Dim totalMailCount As Integer =                                                   ' 計算所有選定根資料夾的郵件總數作為進度分母
            ''    If(CheckSub2.Checked, rootFolders.Sum(Function(f) GetMailCountRecursive(f)),  ' CheckSubFolder2.Checked = True  → 含子資料夾: 各自完整子樹的總和
            ''                          rootFolders.Sum(Function(f) GetMailCount(f)))           ' CheckSubFolder2.Checked = False → 只算選定的那一層
            ''' 2026/3/20, 重寫了底層GetMailCountAll() 效能還是比不過現在上面的遞迴版本
            '' 原因: 原版遞迴只走一遍 COM 資料夾樹，新版走了兩遍COM:
            '' 第一遍: GetSubtreeToList()  → BFS 遍歷，存取每個 folder.Folders
            '' 第二遍: For Each allFolders → GetMailCount() 再讀每個資料夾一次

            ' --- 計算所有選定根資料夾的郵件總數，作為 CollectYearCounts 進度條的分母
            ' todo: 現在這一段只為了顯示進度分母就重新計算郵件總數, 多花了一倍的時間好像不太划算?
            Dim totalMailCount As Long = 0
            Dim totalRoots As Integer = targetFolderList.Count
            Dim processedRoots As Integer = 0
            Dim swThrottleB As New Stopwatch : swThrottleB.Start()  ' by Claude, 2026/04/11
            For Each rf As Outlook.Folder In targetFolderList
                If _includeSubTab2 Then
                    ' GetMailCountAll 是 Async，不能像上面放在 LINQ Sum lambda 裡，改用明確的 For Each + Await (光是這二點, 效能也差一大截)
                    ' 2026/3/20, 再次嚐試把GetMailCountAll() 改平行處理+非同步, 效能回復到原有的遞迴函數, 但速度並不穩定
                    Dim c As Long = Await GetCachedMailCountAllAsync(rf)
                    If c > 0 Then totalMailCount += c       ' -1 表示讀取失敗，略過不累加
                    ' --- 序號校驗點 1 (預讀中) ---
                    If _tab2SelectSeq <> mySeq Then Return  ' _dbg("結束", "序號已不匹配，丟棄本次結果（預讀中斷）")
                Else
                    Dim c As Integer = GetCachedMailCount(rf)
                    If c > 0 Then totalMailCount += c       ' 只算本層，Layer3 同步函數，不需要 Await
                End If
                processedRoots += 1

                ' by Claude, 2026/04/11: todo 完成 - 加入 progressreport + throttle + ESC 中斷
                If swThrottleB.ElapsedMilliseconds >= 100 Then
                    swThrottleB.Restart()
                    ProgressBar2.Text = $"正在計算郵件總數: {processedRoots}/{totalRoots} 個根資料夾 (累計 {totalMailCount:N0} 封)..."
                    Await Task.Delay(1, cToken)  ' by Gemini, 2026/04/11: 讓消息泵有機會處理 ESC
                End If
            Next

            ' 呼叫 Layer2 流程協調層執行統計；結果存入 _lv2DataYear session 快取，GoToMonthView/GoToYearView 直接 render 不重算
            Dim progressYear = New Progress(Of ProgressReport)(Sub(p) ProgressBar2.Text = p.Message)
            _lv2DataYear = Await CollectYearCounts(folderList, totalMailCount, progressYear, cToken)

            ' --- 序號校驗點 2 (核心運算完成後) ---
            If _tab2SelectSeq <> mySeq Then Return          ' _dbg("結束", "序號已不匹配，丟棄本次結果（運算完畢中斷）")
            stopwatch.Stop()                                ' ✅ 統計完成後才停錶

            ' 2026/04/12: ShowResultTab2 + ShowProgressTab2 拆分為 Render 函數 + inline progress
            RenderLvYearView(_lv2DataYear)
            RenderCtYearView(_lv2DataYear)

            Dim _yTotal As Integer = _lv2DataYear.Values.Sum   ' Values.Sum 是最可靠的實際計數（含/不含子資料夾皆正確）
            Dim _ySpd As Double = If(stopwatch.Elapsed.TotalSeconds > 0, _yTotal / stopwatch.Elapsed.TotalSeconds, 0)
            ProgressBar1.Text = $"共 {_yTotal:###,###,##0} 封 / {stopwatch.Elapsed.TotalSeconds:0.00} 秒"
            ProgressBar2.Text = $"(年度統計完成 - 處理速度為 {_ySpd:###,##0}/sec)"
            sender.Enabled = True : sender.Focus() : Cursor = Cursors.Default
            _dbg("結束")
        Catch ex As OperationCanceledException
            _dbg("結束", "ESC 中斷")
            ProgressBar1.Text = "已中斷。" : ProgressBar2.Text = "" : Cursor = Cursors.Default
        Catch ex As System.Exception
            _dbg("錯誤", ex.Message) : Cursor = Cursors.Default
        End Try

    End Sub
    Private Async Sub ListView2_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles ListView2.MouseDoubleClick
        ' ---------------------------------------------------------------
        ' ListView2 雙擊事件（2026/04/12 重構：委派 GoToYearView / GoToMonthView，消除與 HandleListViewKeyPress 的重複）
        ' 年度視圖: 雙擊某一年 → GoToMonthView(selectedYear, cToken)
        ' 月份視圖: 雙擊「← 返回」→ GoToYearView()
        ' ✅ cToken 重構: 每次進入都 OkayNowYouHaveToken() 取得全新 cToken（同 2026/04/11 原設計）
        ' ---------------------------------------------------------------
        _dbg("開始")
        Dim clickedItem As ListViewItem = ListView2.GetItemAt(e.X, e.Y)
        If clickedItem Is Nothing Then Return

        Dim cToken As CancellationToken = OkayNowYouHaveToken()
        Try
            If _tab2IsMonthView AndAlso clickedItem.Tag?.ToString() = "BACK" Then
                Await GoToYearView() : Return
            End If
            Dim selectedYear As Integer = ParseYearFromText(clickedItem.Text)
            If selectedYear = 0 OrElse _tab2FolderList Is Nothing OrElse _tab2FolderList.Count = 0 Then Return
            Await GoToMonthView(selectedYear, cToken)
            _dbg("結束", $"{selectedYear} 年")
        Catch ex As OperationCanceledException
            _dbg("結束", "ESC 中斷")
            ProgressBar1.Text = "已中斷。" : ProgressBar2.Text = "" : Cursor = Cursors.Default
        Catch ex As System.Exception
            _dbg("錯誤", ex.Message)
        End Try

    End Sub
    Private Sub ListView2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListView2.SelectedIndexChanged
        ' ---------------------------------------------------------------
        ' ListView2 選取變更 ↔ Chart2 對應長條同步高亮
        ' 年度視圖: 選取某年 → 高亮 Chart2 中對應年份的長條
        ' 月份視圖: 選取某月 → 高亮 Chart2 中對應月份的長條
        ' 與 Chart2_MouseMove 共用 _lastHoveredPointIndex，確保兩者高亮互斥不累積
        ' 注意: Chart2_MouseLeave 會清掉 _lastHoveredPointIndex，
        '       滑鼠離開圖表後 ListView 的選取高亮也會消失
        ' 2026-03-18, by Claude.ai / 2026-04-04 by Gemini: 共用函數重構
        ' ---------------------------------------------------------------
        _dbg("開始")
        If ListView2.SelectedItems.Count = 0 Then Return
        If Chart2.Series.Count = 0 OrElse Chart2.Series(0).Points.Count = 0 Then Return  ' Chart 尚未載入資料，直接結束

        Dim selectedItem As ListViewItem = ListView2.SelectedItems(0)
        Dim selectedText As String = selectedItem.Text.Trim()

        ' ── 防護特殊控制列 ──
        If selectedItem.Tag?.ToString() = "BACK" Then Return
        If selectedText.Contains("──") Then Return

        ' ── 找出目標 DataPoint index ──
        Dim targetIndex As Integer = -1
        If Not _tab2IsMonthView Then
            ' 年度視圖: 解析文字中的年份
            Dim selectedYear As Integer = ParseYearFromText(selectedText)
            If selectedYear > 0 Then
                For i = 0 To Chart2.Series(0).Points.Count - 1
                    If CInt(Chart2.Series(0).Points(i).XValue) = selectedYear Then targetIndex = i : Exit For
                Next
            End If
        Else
            ' 月份視圖: 利用 ParseMonthFromText 解析數字
            Dim selectedMonth As Integer = ParseMonthFromText(selectedText)
            If selectedMonth > 0 Then targetIndex = selectedMonth - 1 ' RenderCtMonthView 依 1~12 月順序加入 DataPoints，月份N = index N-1
        End If

        If targetIndex < 0 OrElse targetIndex >= Chart2.Series(0).Points.Count Then Return

        ' ── 呼叫共用的圖表渲染函數套用高亮 ──
        BrushChartHoverState(Chart2, targetIndex)

    End Sub
    Private Sub Chart2_MouseClick(sender As Object, e As MouseEventArgs) Handles Chart2.MouseClick
        ' ---------------------------------------------------------------
        ' Chart2 點擊長條 → 同步高亮 ListView2 對應的年份或月份列
        ' 反向對應: ListView2_SelectedIndexChanged 負責 ListView → Chart2
        ' 設定 item.Selected = True 會觸發 ListView2_SelectedIndexChanged，
        ' 後者會再次把 Chart2 同一條塗紅 — 因為是同一條，行為是 idempotent 不會閃爍
        ' 2026-03-18, by Claude.ai / 2026-04-04 by Gemini: 共用函數重構
        ' ---------------------------------------------------------------
        _dbg("開始")
        If Chart2.Series.Count = 0 OrElse Chart2.Series(0).Points.Count = 0 Then Return

        Dim hit As HitTestResult = Chart2.HitTest(e.X, e.Y)
        If hit.ChartElementType <> ChartElementType.DataPoint Then Return

        ' ── 根據目前視圖找目標 ListViewItem ──
        Dim pt As DataPoint = Chart2.Series(0).Points(hit.PointIndex)
        Dim targetItem As ListViewItem = Nothing
        If Not _tab2IsMonthView Then
            targetItem = FindListViewItemByYear(CInt(pt.XValue))        ' 年度視圖: pt.XValue = 年份
        Else
            Dim monthNum As Integer = ParseMonthFromText(pt.AxisLabel)  ' 月份視圖: 呼叫共用函數解析
            If monthNum > 0 Then targetItem = FindListViewItemByMonth(monthNum)
        End If

        ' ── 開始在 ListView2 中選取目標列，並確保只有一列被選取 (Selected = True)，其他列都取消選取 (Selected = False)
        If targetItem Is Nothing Then Return
        For Each item As ListViewItem In ListView2.Items    ' ✅ 先清除所有現有選取，避免多次點擊累積多個 highlighted item
            item.Selected = False                           ' 改用逐一設 Selected = False，安全可靠
        Next                                                ' 不可用 ListView.SelectedItems.Clear() (會丟 NotSupportedException)

        ' ── 選取並捲動到目標列 (會觸發 SelectedIndexChanged 同步塗色) ──
        targetItem.Selected = True
        targetItem.Focused = True
        ListView2.Focus()
        targetItem.EnsureVisible()
        _dbg("結束")

    End Sub
    Private Sub Chart2_MouseMove(sender As Object, e As MouseEventArgs) Handles Chart2.MouseMove
        ' ✅ 用 MouseMove，滑鼠移動時持續觸發，才能追蹤到每個長條
        Dim chart As Chart = CType(sender, Chart)
        If chart.Series.Count = 0 OrElse chart.Series(0).Points.Count = 0 Then Return

        Dim hit As HitTestResult = chart.HitTest(e.X, e.Y)
        If hit.ChartElementType = ChartElementType.DataPoint Then
            BrushChartHoverState(chart, hit.PointIndex)
        Else
            ' 滑鼠離開所有長條，還原上一個點與標題
            ClearChartHoverState(chart)
        End If

    End Sub
    Private Sub Chart2_MouseLeave(sender As Object, e As EventArgs) Handles Chart2.MouseLeave
        ' 滑鼠離開 Chart2，還原上一個高亮點與標題
        Dim chart As Chart = CType(sender, Chart)
        ClearChartHoverState(chart)
    End Sub
    Private Sub CheckSubFolder2_CheckStateChanged(sender As Object, e As EventArgs) Handles CheckSubFolder2.CheckStateChanged
        _dbg("開始", _includeSubTab2.ToString)

        ' by Gemini, 2026/03/29: 合併為 SimTree2 單一操作路徑
        Dim selectedNodes As List(Of TreeNode) = SimTree2.SelectedNodes
        If selectedNodes IsNot Nothing AndAlso selectedNodes.Count > 0 Then
            SimTree2_AfterSelect(SimTree2, New TreeViewEventArgs(selectedNodes(0)))
        End If
        _dbg("結束")

    End Sub
#End Region
#Region "  ├ Layer2 流程協調層"
    Private Async Function CollectYearCounts(folderList As List(Of Outlook.Folder), totalMailCount As Integer, progress As IProgress(Of ProgressReport), cToken As CancellationToken) As Task(Of ConcurrentDictionary(Of Integer, Integer))
        ' ---------------------------------------------------------------
        ' === Layer 2: 流程協調層 ===
        ' 職責: BFS 遍歷 folderList，管理快取，驅動 Layer3 計算，合併結果，回報進度
        '       逐資料夾計算年份統計並合併，是 Tab2 所有統計流程的唯一入口
        ' 規則: 不直接碰 UI 控制項 (ProgressBar1 等) ，進度透過 onProgress callback 傳出, 自己不會知道上一層是單選還是多選，只知道接受傳入的 folderList 清單
        '
        ' 參數:
        '   folderList      : 由 Layer1 組裝好的目標資料夾清單 (已包含 BFS 展開結果)
        '   totalMailCount  : 總郵件數，用來計算進度百分比的分母
        '   onProgress      : 進度 callback，每處理完一個資料夾呼叫一次，回傳 (已處理, 總數)
        '   cToken          : CancellationToken，由 Layer1 透過 OkayNowYouHaveToken() 取得，ESC 時拋 OperationCanceledException
        ' ---------------------------------------------------------------
        _dbg(" ├ 開始", $"目標資料夾數: {folderList.Count}") ' by Gemini, 2026/04/10: 調整縮排層級為 Level 1

        Dim swThrottle As New Stopwatch() : swThrottle.Start() ' by Gemini, 2026/04/02
        Dim processedCount As Integer = 0       ' ✅ 局部計數器，取代全域的 _intProcessedCount 和 _intTotalMailCount, 不會被其他事件汙染，快速點選時不會計數錯亂
        Dim processedFolders As Integer = 0
        Dim totalFolders As Integer = folderList.Count

        Dim merged As New ConcurrentDictionary(Of Integer, Integer)
        Try
            For Each folder As Outlook.Folder In folderList

                ' ✅ 2026/04/10: 提前過濾沒有信件的資料夾 (by Gemini) 既然根本沒有信，就不必去查 DB 或打 COM，直接跳過
                If GetCachedMailCount(folder) <= 0 Then ' 放個空快取避免下次又查 (<= 0 也包含 -1 的情況視同沒信防護)
                    _cacheYearCounts(folder.FolderPath) = New ConcurrentDictionary(Of Integer, Integer)()
                    processedFolders += 1 : Continue For
                End If

                ' 只快取「單一資料夾」的結果，合併邏輯由本層負責
                Dim folderResult As ConcurrentDictionary(Of Integer, Integer)
                Dim cacheKey As String = folder.FolderPath, value As ConcurrentDictionary(Of Integer, Integer) = Nothing
                If _cacheYearCounts.TryGetValue(cacheKey, value) Then   ' ① 記憶體命中: 直接取結果，完全不再讀 COM
                    If _iLikeNoisy Then _dbg("    ├ Cache Hit: ", folder.Name)
                    folderResult = value
                Else
                    ' ② DB lazy load：先查 SQLite，命中就直接用，不走 COM
                    Dim dbResult = DbGetYearCountsForFolder(cacheKey)
                    If dbResult IsNot Nothing Then
                        _cacheYearCounts(cacheKey) = dbResult   ' 直接填入記憶體快取
                        folderResult = dbResult
                        If _iLikeNoisy Then _dbg("    ├ DB Hit: ", folder.Name)
                    Else                                                ' ③ 快取未命中: 呼叫 Layer3 COM 資料層，計算這個資料夾的年份分佈
                        If _iLikeNoisy Then _dbg("    ├ Cache miss: ", folder.Name)
                        folderResult = Await GetYearCountsForFolder(folder)
                        _cacheYearCounts(cacheKey) = folderResult       ' ✅ 計算完成後存入快取，下次點選同一資料夾直接命中
                        ' ✅ 用 "=" 賦值 (非 .Add()) ，有重複 key 時直接覆蓋，不拋例外
                    End If
                End If

                merged = MergeDictionaries(merged, folderResult)    ' 把這個資料夾的結果合併到總計 (純 .NET 運算，不碰 COM)
                processedCount += folderResult.Values.Sum()         ' 累加已處理郵件數，透過 callback 通知 Layer1 更新進度顯示
                processedFolders += 1

                ' by Gemini, 2026/04/02: 100ms 節流回報進度，且不輸出 _dbg()
                If swThrottle.ElapsedMilliseconds >= 100 OrElse processedCount >= totalMailCount Then
                    progress?.Report(New ProgressReport With {.CurrentCount = processedCount, .TotalCount = totalMailCount,
                                                              .Message = $"正在統計年度分佈: ({processedFolders}/{totalFolders})個資料夾 (已統計 {processedCount:###,###,##0} / {totalMailCount:###,###,##0} 封信)..."})
                    swThrottle.Restart()
                    Await Task.Delay(1, cToken) ' 這裡一定要保留至少 .delay(1) 才能讓 ESC 中斷生效 (simon, 2026/04/05)
                End If
            Next
        Catch ex As OperationCanceledException
            ' by Gemini, 2026/04/11: 捕捉 ESC 中斷，回傳已計算的部分結果而不拋出不常
            _dbg(" ├ 中斷", "ComputeYearCounts 已中斷")
        End Try
        _dbg(" ├ 結束", $"共 {merged.Count} 個年份 | 郵件總計: {merged.Values.Sum}") ' by Gemini, 2026/04/10
        Return merged

    End Function
    Private Async Function CollectMonthCounts(selectedYear As Integer, cToken As CancellationToken) As Task(Of ConcurrentDictionary(Of Integer, Integer))
        ' ---------------------------------------------------------------
        ' 月份資料收集 Layer2（2026/04/12 由 ShowMonthView 計算部分拆出）
        ' 職責: 遍歷 _tab2FolderList，對每個資料夾呼叫 GetMonthCountsForYear，合併結果，回報進度
        '       不碰 UI render（render 由 GoToMonthView 的 RenderLvMonthView / RenderCtMonthView 負責）
        '       cToken 與 CollectYearCounts 同理，都需要傳入以支援 ESC 中斷
        '       OperationCanceledException 由 caller（GoToMonthView → DoubleClick / HandleListViewKeyPress）的 Catch 攔截
        ' ---------------------------------------------------------------
        _dbg(" ├ 開始", selectedYear.ToString())

        Dim monthCounts As New ConcurrentDictionary(Of Integer, Integer)
        Dim totalFolders As Integer = _tab2FolderList.Count
        Dim processedFolders As Integer = 0
        Dim totalMailCount As Long = 0

        ' 先算信件總數作為進度顯示的分母（節流）
        Dim swThrottleA As New Stopwatch : swThrottleA.Start()
        For Each f In _tab2FolderList
            Dim c As Integer = GetCachedMailCount(f)
            If c > 0 Then totalMailCount += c
            If swThrottleA.ElapsedMilliseconds >= 100 Then
                swThrottleA.Restart() : Await Task.Delay(1, cToken)
            End If
        Next

        ' 逐資料夾取月份分布並合併
        Dim swThrottle As New Stopwatch() : swThrottle.Start()
        For Each folder As Outlook.Folder In _tab2FolderList
            processedFolders += 1
            Dim folderMonthCounts As ConcurrentDictionary(Of Integer, Integer) = Await GetMonthCountsForYear(folder, selectedYear)
            monthCounts = MergeDictionaries(monthCounts, folderMonthCounts)

            If swThrottle.ElapsedMilliseconds >= 100 OrElse processedFolders = totalFolders Then
                ProgressBar1.Text = "正在讀取..."
                ProgressBar2.Text = $"正在統計 {selectedYear} 年月份分佈: ({processedFolders}/{totalFolders})個資料夾 (相依包含共計 {totalMailCount:N0} 封信)。"
                swThrottle.Restart() : Await Task.Delay(1, cToken)
            End If
        Next
        _dbg(" ├ 結束", $"{selectedYear} 年 | 月份數: {monthCounts.Count}")
        Return monthCounts

    End Function
    Private Async Function GoToYearView() As Task
        ' ---------------------------------------------------------------
        ' 共用導覽：返回年度視圖（2026/04/12 取代 ShowYearView，供 DoubleClick 與 KeyPress 共用）
        ' 職責: 純 render from _lv2DataYear session 快取，完全不碰 COM / Layer2 計算層
        '       _tab2MonthViewYear 刻意不 reset，讓 _lv2DataMonth 方案A快取跨 back-and-forth 繼續有效
        ' ---------------------------------------------------------------
        _dbg(" ├ GoToYearView")
        Dim yearToRestore As Integer = _tab2MonthViewYear   ' 先記住要還原游標的年份
        _tab2IsMonthView = False
        Await Task.Yield()  ' 讓 UI 喘口氣，確保畫面流暢切換

        If _lv2DataYear IsNot Nothing AndAlso _tab2FolderList IsNot Nothing AndAlso _tab2FolderList.Count > 0 Then
            Cursor = Cursors.WaitCursor
            RenderLvYearView(_lv2DataYear)
            RenderCtYearView(_lv2DataYear)

            Dim _rTotal As Integer = _lv2DataYear.Values.Sum
            ProgressBar1.Text = $"共 {_rTotal:###,###,##0} 封"
            ProgressBar2.Text = "(返回年度統計)"
            Cursor = Cursors.Default
        End If

        ' 還原游標到進入月份前的那一年，讓使用者感覺回到剛才看的地方
        ' ✅ 2026-03-16 Bug fix: 移除此處多餘的 item.Selected = True（造成兩個 highlighted item 的根源）
        If yearToRestore > 0 AndAlso ListView2.Items.Count > 0 Then
            Dim tgt = FindListViewItemByYear(yearToRestore)
            If tgt IsNot Nothing Then
                tgt.Selected = True : tgt.Focused = True : tgt.EnsureVisible() : ListView2.Focus()
            End If
        End If
        _dbg(" ├ GoToYearView 完成")

    End Function
    Private Async Function GoToMonthView(selectedYear As Integer, cToken As CancellationToken) As Task
        ' ---------------------------------------------------------------
        ' 共用導覽：進入月份視圖（2026/04/12 取代 ShowMonthView，供 DoubleClick 與 KeyPress 共用）
        ' 職責: 方案A _lv2DataMonth 快取判斷 → 命中時純 render；未命中時 CollectMonthCounts → render
        '       _tab2MonthViewYear 同時作為「目前顯示年份」與「方案A快取 tag」
        ' OperationCanceledException 由 caller（DoubleClick / HandleListViewKeyPress）的 Catch 攔截
        ' ---------------------------------------------------------------
        _dbg(" ├ GoToMonthView", selectedYear.ToString())
        Dim swM As New Stopwatch() : swM.Start()
        ProgressBar1.Text = "" : ProgressBar2.Text = "" : Cursor = Cursors.WaitCursor

        If _lv2DataMonth IsNot Nothing AndAlso _tab2MonthViewYear = selectedYear Then
            ' ★ 快取命中：直接 render，完全不碰計算層（方案A：同一年份才命中）
            _dbg(" ├ _lv2DataMonth 快取命中", selectedYear.ToString())
            _tab2IsMonthView = True
            RenderLvMonthView(selectedYear, _lv2DataMonth)
            RenderCtMonthView(_lv2DataMonth, selectedYear)

            Dim _mHit As Integer = _lv2DataMonth.Values.Sum
            ProgressBar1.Text = $"共 {_mHit:###,###,##0} 封"
            ProgressBar2.Text = $"({selectedYear} 年月份分佈 - 按 ESC 或雙擊標題橫列可返回視圖) "
        Else
            ' ★ 快取未命中：CollectMonthCounts → _cacheMonthCounts 一定命中 → merge → render
            ' todo: 這裡就算未命中, 也是從 _cacheMonthCounts 取出, 為何那麼久?? 764 folder要一秒鐘??
            _dbg(" ├ _lv2DataMonth 快取未命中，開始計算", selectedYear.ToString())
            Dim mc As ConcurrentDictionary(Of Integer, Integer) = Await CollectMonthCounts(selectedYear, cToken)
            _lv2DataMonth = mc : _tab2MonthViewYear = selectedYear : _tab2IsMonthView = True
            swM.Stop()
            RenderLvMonthView(selectedYear, mc)
            RenderCtMonthView(mc, selectedYear)

            Dim _mMiss As Integer = mc.Values.Sum
            Dim _mSpd As Double = If(swM.Elapsed.TotalSeconds > 0, _mMiss / swM.Elapsed.TotalSeconds, 0)
            ProgressBar1.Text = $"共 {_mMiss:###,###,##0} 封 / {swM.Elapsed.TotalSeconds:0.00} 秒"
            ProgressBar2.Text = $"({selectedYear} 年月份分佈讀取完成 - 按 ESC 或雙擊標題橫列可返回視圖) "
        End If

        ' 確保 SimTree2 的選取節點保持可見
        Cursor = Cursors.Default
        If SimTree2.Visible Then
            Dim nodes As List(Of TreeNode) = SimTree2.SelectedNodes
            If nodes IsNot Nothing AndAlso nodes.Count > 0 Then nodes(0).EnsureVisible()
        End If
        _dbg(" ├ GoToMonthView 完成", selectedYear.ToString())

    End Function

    Private Sub RenderLvYearView(yearCounts As ConcurrentDictionary(Of Integer, Integer))
        ' ---------------------------------------------------------------
        ' 年度視圖 ListView2 渲染（2026/04/12 由 ShowResultTab2 拆出）
        ' 職責: 純 UI render，不做計算，不查快取，不碰 COM
        ' 對稱: RenderCtYearView 負責同一視圖的 Chart2 部分
        ' ---------------------------------------------------------------
        _dbg("開始", yearCounts?.Count)

        ListView2.Items.Clear()
        If yearCounts Is Nothing OrElse yearCounts.IsEmpty Then
            ClearChart2Series()     ' ★ 空資料夾時也要清除 Chart2，否則前一個資料夾的圖表會殘留
            ListView2.Items.Add(New ListViewItem("找不到郵件"))
        Else
            ListView2.BeginUpdate()     ' ✅ 批次更新，避免每次 Add 都觸發重繪
            Dim sortedYearCounts = yearCounts.OrderBy(Function(pair) pair.Key).ToList()
            For Each pair In sortedYearCounts
                ListView2.Items.Add(New ListViewItem({pair.Key, pair.Value.ToString("###,###,##0") & " "}))  ' 字串結尾一律補一格空白 (by Gemini, 2026/03/31)
            Next
            ListView2.EndUpdate()
        End If
        _dbg("結束")

    End Sub
    Private Sub RenderLvMonthView(selectedYear As Integer, monthCounts As ConcurrentDictionary(Of Integer, Integer))
        ' ---------------------------------------------------------------
        ' 月份視圖 ListView2 渲染（2026/04/12 由 ShowMonthView render 部分拆出）
        ' 職責: 純 UI render，不做計算，不查快取，不碰 COM
        ' 對稱: RenderCtMonthView 負責同一視圖的 Chart2 部分
        ' ---------------------------------------------------------------
        _dbg(" ├ 開始", selectedYear.ToString())
        ListView2.BeginUpdate()
        ListView2.Items.Clear()

        ' 第一行: 返回按鈕
        Dim backItem As New ListViewItem("← 返回年度統計")
        backItem.SubItems.Add("") : backItem.Tag = "BACK"
        backItem.ForeColor = Color.Gray
        backItem.Font = New Font(_fontDefault, _fontItalic)
        ListView2.Items.Add(backItem)

        ' 第二行: 年份標題
        Dim titleItem As New ListViewItem($"── {selectedYear} 年月份分佈 ──")
        titleItem.SubItems.Add($"共 {monthCounts.Values.Sum:###,###,##0}  封")  ' 字串結尾補上空白防止選取時切邊，與下方對齊
        titleItem.ForeColor = Color.DimGray
        titleItem.Font = New Font(_fontDefault, _fontBold)
        ListView2.Items.Add(titleItem)

        ' 逐月顯示（只顯示有郵件的月份）
        For month As Integer = 1 To 12
            Dim count As Integer = 0
            monthCounts.TryGetValue(month, count)
            If count > 0 Then
                Dim monthItem As New ListViewItem($"{selectedYear} /  {month:D2}月")
                monthItem.SubItems.Add(count.ToString("###,###,##0") & " ")  ' 字串結尾一律補一格空白
                ListView2.Items.Add(monthItem)
            End If
        Next
        ListView2.EndUpdate()
        _dbg(" ├ 結束", selectedYear.ToString())

    End Sub
    Private Sub RenderCtYearView(yearCounts As ConcurrentDictionary(Of Integer, Integer))
        ' ---------------------------------------------------------------
        ' 年度視圖 Chart2 渲染（2026/04/12 由 UpdateChart2forYearView 改名重構）
        ' 職責: 純 UI render；接受 ConcurrentDictionary，內部自行排序（原版由 caller 排序後傳 List，現改為自己排序讓介面更乾淨）
        ' 對稱: RenderLvYearView 負責同一視圖的 ListView2 部分
        ' ---------------------------------------------------------------
        _dbg("開始")
        ClearChart2Series()     ' 清除之前的統計結果，包括 Series Points、平均線 Series、平均值標籤 Annotation
        If yearCounts Is Nothing OrElse yearCounts.IsEmpty Then Return
        Dim sortedYearCounts = yearCounts.OrderBy(Function(p) p.Key).ToList()

        ' 添加數據到 Series, 在 Chart2 中顯示統計結果
        Dim series As Series = Chart2.Series(0)
        For Each pair In sortedYearCounts
            series.Points.AddXY(pair.Key, pair.Value)
        Next

        ' 依內容大小來設置 Chart2 的 X 軸上下限
        With Chart2.ChartAreas(0).AxisX
            .Minimum = sortedYearCounts.Min(Function(p) p.Key) - 0.5
            .Maximum = sortedYearCounts.Max(Function(p) p.Key) + 0.5
            .Interval = 1
            .IntervalOffset = 0                 ' ✅ 還原年度視圖的長條置中偏移
            .LabelStyle.Format = "####"         ' ✅ 還原年份格式
            .LabelStyle.Interval = 1
            .LabelStyle.IntervalOffset = 0.5    ' ✅ 校正還原上面 max/min 的 0.5 偏移
            .MajorTickMark.IntervalOffset = 0   ' ✅ 還原刻度偏移
        End With

        ' 添加一條代表平均值的線（獨立 Series 才能控制線型，StripLine 不支援虛線）
        ' 2026/3/6 by Claude Code；2026/04/12 移入 RenderCtYearView
        Dim average As Double = sortedYearCounts.Average(Function(pair) pair.Value)
        Dim xMin As Double = sortedYearCounts.Min(Function(pair) pair.Key)
        Dim xMax As Double = sortedYearCounts.Max(Function(pair) pair.Key)

        Dim avgSeries As New Series("平均線") With {.ChartType = SeriesChartType.Line,
                                                    .Color = ThemeColors.avgLineColor,
                                                    .BorderWidth = 2,
                                                    .BorderDashStyle = ChartDashStyle.Dash,  ' ✅ 虛線
                                                    .ChartArea = Chart2.ChartAreas(0).Name,
                                                    .IsVisibleInLegend = False}
        avgSeries.Points.AddXY(xMin - 1, average)  ' 0: 從 X 軸最小值往左延伸
        avgSeries.Points.AddXY(xMax, average)       ' 1: 圖表最右邊長條的確切 X 座標（錨定用）
        avgSeries.Points.AddXY(xMax + 1, average)  ' 2: 到 X 軸最大值往右延伸

        ' 用 TextAnnotation 顯示平均值標籤 (by Gemini, 2026/04/04 改用 DeepAmber 提升辨識度)
        Dim avgLabel As New TextAnnotation With {.Name = "平均值標籤",
                                                 .Text = "AVG: " & average.ToString("#,###,##0"),
                                                 .ForeColor = ThemeColors.avgLineColor,
                                                 .Font = New Font("Tahoma", 10.0F, System.Drawing.FontStyle.Bold),
                                                 .AnchorDataPoint = avgSeries.Points(1),           ' 錨定在最右側長條的中間點 X 座標
                                                 .AnchorAlignment = ContentAlignment.BottomCenter,  ' ★ 強制對齊點的正上方（避免 MS Chart 自動亂飄移）
                                                 .AnchorOffsetX = 0,    ' 保持置中
                                                 .AnchorOffsetY = -1,   ' 產生 1% 的空隙，確保不在線上
                                                 .BackColor = Color.Transparent,
                                                 .LineColor = Color.Transparent}
        Chart2.Series.Add(avgSeries)
        Chart2.Annotations.Add(avgLabel)
        Chart2.Invalidate()     ' 強制重新繪製圖表
        _dbg("結束")

    End Sub
    Private Sub RenderCtMonthView(monthCounts As ConcurrentDictionary(Of Integer, Integer), year As Integer)
        ' ---------------------------------------------------------------
        ' 月份視圖 Chart2 渲染（2026/04/12 由 UpdateChart2forMonthView 改名）
        ' 月份長條圖：只畫 1~12 月，X 軸標籤顯示「M月」，不畫平均線
        ' 完整替換 Chart2 的內容，與 RenderCtYearView 平行存在
        ' 對稱: RenderLvMonthView 負責同一視圖的 ListView2 部分
        ' ---------------------------------------------------------------
        _dbg("開始", year)
        ClearChart2Series()     ' 清除之前的所有圖表內容（同 RenderCtYearView 的清除邏輯）

        ' 把 1~12 月的資料全部加入（沒有郵件的月份補 0，讓 X 軸保持完整 12 格）
        Dim series As Series = Chart2.Series(0)
        For month As Integer = 1 To 12
            Dim count As Integer = 0
            monthCounts.TryGetValue(month, count)
            Dim pt As New DataPoint()
            pt.SetValueXY(month, count)
            pt.AxisLabel = $"{month}月"     ' ✅ 用月份名稱當 X 軸標籤，比純數字 1~12 更易讀
            series.Points.Add(pt)
            pt.IsVisibleInLegend = True
        Next

        ' X 軸固定顯示 1~12，不根據資料範圍自動縮放
        ' X 軸重置所有從 InitChart2 繼承的年度設定，改成月份專用設定
        With Chart2.ChartAreas(0).AxisX
            .Minimum = 0.5
            .Maximum = 12.5
            .Interval = 1
            .IntervalOffset = 0                 ' ✅ 清除 InitChart2 的 0.5 偏移量
            .LabelStyle.Format = ""             ' ✅ 清除 "####" 年份格式，讓 AxisLabel 屬性生效
            .LabelStyle.Interval = 1
            .LabelStyle.IntervalOffset = 0.5    ' ✅ 清除偏移
            .MajorTickMark.IntervalOffset = 0   ' ✅ 清除刻度偏移
        End With
        Chart2.Invalidate()
        _dbg("結束", year)

    End Sub
#End Region
#Region "  └ 輔助函數"
    Private Function BuildFilterDateRangeTab2(year As Integer, Optional mon1 As Integer = 1, Optional mon2 As Integer = 12) As String
        If year < 1974 Then Return Nothing
        'Const DATE_FORMAT As String = "yyyy/MM/dd HH:mm:ss"
        'Dim startDate As Date = Date.ParseExact($"{year}/01/01 00:00:00", DATE_FORMAT, Nothing) ' 建立當年的起始日期和結束日期
        'Dim endDate As Date = Date.ParseExact($"{year}/12/31 23:59:59", DATE_FORMAT, Nothing)   ' 設置結束日期的時間為23:59:59
        'Return $"[ReceivedTime] >= '{startDate}' AND [ReceivedTime] <= '{endDate}'"
        ' 2026/3/11, by Claude, 重構BuildFilterDateRangeTab2 函數: 增加了月份參數，並且直接用 Date 物件來建立日期範圍，避免字串格式問題
        Dim startDate As New Date(year, mon1, 1, 0, 0, 0)                                   ' ✅ 用 mon1/mon2 決定起訖月份，預設 1~12 代表整年
        Dim endDate As New Date(year, mon2, Date.DaysInMonth(year, mon2), 23, 59, 59)       ' mon2 的結束日用該月最後一天，避免硬寫 31 日造成 2 月等短月份抓不準
        Return $"[ReceivedTime] >= '{startDate}' AND [ReceivedTime] <= '{endDate}'"

    End Function
    Private Function Find1stYear(selectedFolder As Outlook.Folder) As Integer
        _dbg("開始", selectedFolder.Name)
        ' =============================================================
        ' 尋找資料夾中最早的郵件年份，作為統計的起點
        ' 2026/3/10, by Claude, 重構 Find1stYear 函數
        ' 改進: 多層try/catch加強錯誤處理、確保 COM 物件正確釋放，避免 RCW 殘留問題

        ' 2026/3/24 by Gemini:
        ' 改用 GetTable + GetArray 取代逐年 Restrict之後就用不到這個函數了，因為 GetTable 直接過濾掉 1974 年之前的郵件
        ' =============================================================
        Dim mail As Outlook.MailItem = Nothing
        Dim allItems As Outlook.Items = Nothing
        Dim validItems As Outlook.Items = Nothing

        ' 改用一層一層的 Try-Catch 包裹過濾，確保物件讀取失敗或類型轉換失敗都能被捕捉到
        Try
            ' 資料夾裡可能混有 MeetingRequest / ContactItem / Note 等, 這些物件沒有 ReceivedTime
            ' 透過 COM late binding 存取會拋 COMException 或 AccessViolationException (.NET 4+ 的 corrupted state exception) ，bare Catch 接不住
            ' ✅ 先 Restrict 過濾掉 null/零值 ReceivedTime 的壞項目，再升冪排序取最舊年份
            allItems = selectedFolder.Items : If allItems Is Nothing OrElse allItems.Count = 0 Then Return 1974
            validItems = allItems.Restrict("[ReceivedTime] > '1974/01/01'") : If validItems.Count = 0 Then Return 1974
            validItems.Sort("[ReceivedTime]", OlSortOrder.olDescending)
            Dim firstItem As Object = validItems.GetFirst() : If firstItem Is Nothing Then Return 1974
            mail = TryCast(firstItem, Outlook.MailItem) : If mail Is Nothing Then Return 1974
            Dim year As Integer = mail.ReceivedTime.Year : Return If(year <= 0 OrElse year > Date.Today.Year, 1974, year)
        Catch ex As System.Exception
            _dbg("錯誤", selectedFolder.Name & " - " & ex.Message)
            Return 1974
        Finally ' ✅ Finally 確保不管正常結束或例外都一定釋放，包括 Return 提前返回的情況
            TryMarshalRelease(mail)
            TryMarshalRelease(validItems)
            TryMarshalRelease(allItems)
            _dbg("結束")
        End Try

    End Function
    Private Function MergeDictionaries(dict1 As ConcurrentDictionary(Of Integer, Integer), dict2 As ConcurrentDictionary(Of Integer, Integer)) As ConcurrentDictionary(Of Integer, Integer)
        If dict1 Is Nothing Then dict1 = New ConcurrentDictionary(Of Integer, Integer)
        If dict2 Is Nothing Then Return dict1
        '' 逐一遍歷合併 dict2 的鍵值對到 dict1 中，如果 dict1 已經有相同的鍵，則將值相加
        'For Each kvp As KeyValuePair(Of Integer, Integer) In dict2
        '    If dict1.ContainsKey(kvp.Key) Then: dict1(kvp.Key) += kvp.Value
        '    Else:                               dict1.Add(kvp.Key, kvp.Value)
        '    End If
        'Next
        ' ✅ LINQ 改寫後, 效能更好，因為不需要每次都檢查 dict1 是否包含鍵，直接使用 GetValueOrDefault 來獲取值，如果鍵不存在則返回 0，然後加上 dict2 的值
        For Each kvp In dict2
            dict1(kvp.Key) = dict1.GetValueOrDefault(kvp.Key, 0) + kvp.Value
        Next
        Return dict1

    End Function

    Private Function ParseYearFromText(text As String) As Integer
        ' by Gemini, 2026/04/04: 從字串 (例如 "2024" 或包含年份的標題) 萃取出年份數字的共用邏輯
        If String.IsNullOrWhiteSpace(text) Then Return 0

        ' 處理 "── 2024 年月份分佈 ──" 這種標題格式，或是純數字 "2024"
        ' 如果字串包含 "年"，取 "年" 之前的數字
        Dim cleanText As String = text.Trim()
        Dim yearIdx As Integer = cleanText.IndexOf(CChar("年"))
        Dim numStr As String = ""

        If yearIdx >= 0 Then
            For k As Integer = yearIdx - 1 To 0 Step -1 ' 從 "年" 往前找連續數字
                If Char.IsDigit(cleanText(k)) Then numStr = cleanText(k) & numStr Else If numStr <> "" Then Exit For
            Next
        Else
            For Each c In cleanText ' 如果沒看到 "年"，試著看是不是純數字，或者擷取字串中第一組連續數字
                If Char.IsDigit(c) Then numStr &= c Else If numStr <> "" Then Exit For
            Next
        End If

        Dim resultYear As Integer = 0
        If Integer.TryParse(numStr, resultYear) AndAlso resultYear > 1900 AndAlso resultYear < 2100 Then Return resultYear
        Return 0

    End Function
    Private Function ParseMonthFromText(text As String) As Integer
        ' by Gemini, 2026/04/04: 從字串 (例如 "2024 / 03月" 或 "3月") 萃取出月份數字的共用邏輯
        Dim moonIdx As Integer = text.IndexOf(CChar("月"))
        If moonIdx < 0 Then Return 0

        Dim numStr As String = ""
        For k As Integer = moonIdx - 1 To 0 Step -1
            If Char.IsDigit(text(k)) Then numStr = text(k) & numStr Else Exit For
        Next

        Dim selectedMonth As Integer = 0
        If Integer.TryParse(numStr, selectedMonth) AndAlso selectedMonth >= 1 AndAlso selectedMonth <= 12 Then Return selectedMonth
        Return 0

    End Function
    Private Function FindListViewItemByYear(targetYear As Integer) As ListViewItem
        ' by Gemini, 2026/04/04: 根據年份尋找對應的 ListViewItem
        For Each item As ListViewItem In ListView2.Items
            If ParseYearFromText(item.Text) = targetYear Then Return item
        Next
        Return Nothing
    End Function
    Private Function FindListViewItemByMonth(targetMonth As Integer) As ListViewItem
        ' by Gemini, 2026/04/04: 根據月份數字尋找對應的 ListViewItem
        Dim monthStr As String = targetMonth.ToString("D2") & "月"  ' e.g. "03月"
        For Each item As ListViewItem In ListView2.Items
            If item.Tag IsNot Nothing AndAlso item.Tag.ToString() = "BACK" Then Continue For
            If item.Text.Contains(monthStr) Then Return item
        Next
        Return Nothing
    End Function

    Private Sub ClearChart2Series()
        ' 清除之前的統計結果, 包括 Series Points 和 平均線 Series 以及平均值標籤 Annotation (避免重複加入)
        Chart2.Series(0).Points.Clear()

        Dim existingAvg As Series = Chart2.Series.FindByName("平均線") ' 清除舊的平均線 Series (避免重複加入)
        If existingAvg IsNot Nothing Then Chart2.Series.Remove(existingAvg)

        Dim existingAnnotation = Chart2.Annotations.FindByName("平均值標籤")  ' 先清除舊的 Annotation (避免重複加入)
        If existingAnnotation IsNot Nothing Then Chart2.Annotations.Remove(existingAnnotation)

    End Sub
    Private Sub BrushChartHoverState(chart As Chart, pointIndex As Integer)
        ' by Gemini, 2026/04/04: 抽取共用的圖表高亮渲染邏輯
        If pointIndex = _lastHoveredPointIndex Then Return ' 如果跟上次是同一個點就不重複處理，避免閃爍

        ' ✅ 先把上一個點的顏色跟狀態清除，但在這裡不要馬上 Refresh 畫面，等新的屬性上完再一起重繪避免畫面閃爍
        ClearChartHoverState(chart, refreshChart:=False)

        ' ✅ 把目前這個點變成設定的高亮色
        chart.Series(0).Points(pointIndex).Color = ThemeColors.barHighlight
        _lastHoveredPointIndex = pointIndex

        ' ✅ 取得資料點並計算顯示名稱 (修正年度檢視 AxisLabel 為空的問題)
        Dim dP As DataPoint = chart.Series(0).Points(pointIndex)
        Dim xLabel As String = If(Not String.IsNullOrEmpty(dP.AxisLabel), dP.AxisLabel, dP.XValue.ToString("0000"))
        Dim headerText As String = If(xLabel.Contains(CChar("月")), "月份", "年份")

        ' ✅ 動態數據標籤
        Dim formattedX As String = If(xLabel.Contains(CChar("月")), xLabel, xLabel & "年")
        dP.Label = $"{formattedX}:{dP.YValues(0):##0}"
        dP.IsValueShownAsLabel = True
        chart.Series(0).ToolTip = $"{headerText}: {xLabel}, 數量: {dP.YValues(0):##0}"

        chart.Refresh()         ' ✅ 確保標籤與顏色即時更新

    End Sub
    Private Sub ClearChartHoverState(chart As Chart, Optional refreshChart As Boolean = True)
        ' by Gemini, 2026/04/04: 抽取共用的清除圖表高亮邏輯
        If _lastHoveredPointIndex >= 0 AndAlso chart.Series.Count > 0 AndAlso _lastHoveredPointIndex < chart.Series(0).Points.Count Then
            Dim prevPt = chart.Series(0).Points(_lastHoveredPointIndex)
            prevPt.Color = Color.Empty
            prevPt.Label = ""
            prevPt.IsValueShownAsLabel = False
            _lastHoveredPointIndex = -1

            chart.Series(0).ToolTip = String.Empty
            If refreshChart Then chart.Refresh()
        End If

    End Sub
#End Region
#End Region

#Region "■ 06 Tab3: 依附件條件搜尋"
    ' ===================================================
    ' TabPage3 搜尋附件 — 架構重構演進
    ' ---------------------------------------------------
    ' [v2] by Claude, 2026/03/07
    '      策略: 建立雙階段搜尋 (Phase1 GetTable 快速掃描中繼資料, Phase2 GetItemFromID 讀取附件明細)。
    '      成效: 大幅減少對 MailItem 物件的依賴和操作，提升搜尋效率。
    '
    ' [v3] by Gemini, 2026/04/05 (現行架構)
    '      策略: 導入「管線化處理 (Pipeline)」與「SOLID 分層 (Layer1/Layer2.5/Layer3)」，徹底解耦 MAPI、業務與 UI。
    '      分層:
    '        ├─ Layer1  (UI/流程層) : Button3_Click, ShowResultTab3
    '        ├─ Layer2  (商務過濾層): FilterBySize, FilterByAttachDetailsAsync
    '        ├─ Layer2.5 (快取層)  : GetCachedAttachMailList (_cacheAttachMailList / _cacheAttachFilename)
    '        └─ Layer3  (MAPI操作層): GetAttachMailList
    '
    ' Button3_Click 管線 (Pipeline) 步驟分解:
    '   Step 1. 前置驗證      → 檢查參數合法性 (UI)
    '   Step 2. BFS 遍歷      → GetSubtreeToList，取得目標資料夾樹 (COM)
    '   Step 3. 匯集資料全集  → 向 Layer2.5 索取候選郵件全集 (GetCachedAttachMailList) → ① 記憶體 → ② DB (mail_basic) → ③ L3
    '   Step 4. 管線過濾 1    → 記憶體 LINQ 瞬間過濾大小限制 (FilterBySize)
    '   Step 5. 管線過濾 2    → 依據關鍵字與數量條件深層過濾，配合 Layer2.5 快取判定 (FilterByAttachDetailsAsync)
    '                         → PreloadAttachmentCacheRDOAsync (RDO 平行預載) → GetCachedAttachFilename → ① 記憶體 → ② DB (mail_attachments) → ③ L3
    '   Step 6. UI 映射與顯示 → 將資料封裝為介面項目並顯示，無縫銜接">0"或真實統計 (ShowResultTab3)
    ' ===================================================
#Region "  ├ Layer1 UI事件層"
    Private Async Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
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

            ' ── Step 2: 收集目標資料夾清單 (支援多選防重複) ──
            Dim folderList = Await GetUniqueFolderList(selectedNodes, _includeSubTab3, cToken) ' 導入 HashSet(Of String) 來過濾跨父子節點重複選擇的資料夾
            If folderList.Count = 0 Then Return

            ' by Gemini, 2026/04/09: 計算選定資料夾內的郵件總數 (用於進度顯示母數，參考 Tab2)
            Dim totalMailCount As Long = 0
            For Each f In folderList    ' high: 現在是在 GetSubtreeToList() 460ms之後, FilterBySize() 之前, 第二耗時, 這個迴圈五萬筆資料花450ms
                Dim c As Integer = GetCachedMailCount(f)
                If c > 0 Then totalMailCount += c
                If swThrottle3.ElapsedMilliseconds >= 100 Then
                    swThrottle3.Restart()
                    Await Task.Delay(1, cToken)  ' by Claude, 2026/04/11
                End If
            Next
            ProgressBar1.Text = "正在讀取..."
            ProgressBar2.Text = $"準備掃描 {folderList.Count} 個資料夾 (相依包含共計 {totalMailCount:N0} 封信)..."
            Await Task.Yield()

            ' ── Step 3: 收集含附件的郵件清單 (透過 Layer2.5 快取) ──
            Dim progressPhase1 = New Progress(Of ProgressReport)(Sub(p) ProgressBar2.Text = p.Message)
            Dim targetMails As New List(Of MailItemInfo)
            Try
                For Each folder In folderList   ' high: 現在是在 GetSubtreeToList() 460ms之後, FilterBySize() 之前, 第一耗時, 這個迴圈五萬筆資料花1450ms
                    Dim folderResult = Await GetCachedAttachMailList(folder, progressPhase1)
                    targetMails.AddRange(folderResult)
                    If swThrottle3.ElapsedMilliseconds >= 100 Then
                        swThrottle3.Restart()
                        Await Task.Delay(1, cToken)  ' by Claude, 2026/04/11
                    End If
                Next
            Catch ex As OperationCanceledException
                ' by Gemini, 2026/04/12: 捕捉 ESC 中斷，結算目前已載入的部分郵件清單
                _dbg(" ├ 中斷", $"Step 3 已中斷，結算目前已載入的 {targetMails.Count} 封")
                ProgressBar1.Text = "已結算 (中斷)"
            End Try

            ' ── Pipeline 過濾 1: 大小篩選 ──
            If CheckSize.Checked Then targetMails = FilterBySize(targetMails)   ' 這裡五萬筆資料只花5ms

            ' ── Pipeline 過濾 2: 附件條件深層篩選 ──
            Dim hasKeyword = CheckAttachName.Checked AndAlso TextBox3.Text.Trim.Length > 0
            If hasKeyword OrElse CheckAttCount.Checked Then
                Dim progressPhase2 = New Progress(Of ProgressReport)(Sub(p) ProgressBar2.Text = p.Message)
                targetMails = Await FilterByAttachDetailsAsync(targetMails, progressPhase2, cToken)
            End If

            ' ── 終極 Mapping 與顯示結果 ──
            sw.Stop()
            ShowResultTab3(targetMails, sw.Elapsed.TotalSeconds)
        Catch ex As OperationCanceledException
            _dbg("結束", "ESC 中斷")
            ProgressBar1.Text = "已中斷。" : ProgressBar2.Text = ""
        Catch ex As System.Exception
            MessageBox.Show("搜尋發生錯誤: " & ex.Message, "錯誤")
            _dbg("       ├ 錯誤", ex.Message) ' by Gemini, 2026/04/11: Level 3
        Finally
            ' ── 無論如何都解鎖 UI ──
            Button3.Enabled = True ' : TextBox3.Enabled = CheckAttachName.Checked
            SimTree3.Enabled = True : pnlOptions_tab3.Enabled = True : Cursor = Cursors.Default
            _dbg("結束") ' by Gemini, 2026/04/11: 修正對應開始層級 Level 0
        End Try

    End Sub
    Private Sub ListView3_RetrieveVirtualItem(sender As Object, e As RetrieveVirtualItemEventArgs) Handles ListView3.RetrieveVirtualItem
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
        Dim lvi As New ListViewItem(mail.Subject)
        lvi.SubItems.Add(mail.Size.ToString("###,###,##0"))
        lvi.SubItems.Add(mail.ReceivedTime.ToShortDateString())
        lvi.SubItems.Add(mail.SenderName)
        lvi.SubItems.Add(displayName)
        lvi.SubItems.Add(mail.EntryID)

        e.Item = lvi
    End Sub
    Private Sub ListView3_ColumnClick(sender As Object, e As ColumnClickEventArgs) Handles ListView3.ColumnClick
        ' ==============================================================
        ' by Gemini, 2026/04/10: 效能大躍進 — 虛擬模式排序
        ' 理由: 
        '   當資料量達到 5 萬筆時，原有的 ListViewItemComparer 必須不斷從實體 Item 中取值，每秒只能比較約數千筆，會導致 UI 卡頓
        '   切換至 VirtualMode 後，我們直接在記憶體中對 _lv3MailList (List(Of MailItemInfo)) 進行 LINQ 排序，處理 5 萬筆數據僅需 10-30ms，達成「瞬發」排序的效果
        ' --------------------------------------------------------------
        _dbg("開始", "虛擬列表排序") ' by Gemini, 2026/04/10: Level 0
        Dim sw As New Stopwatch : sw.Start()

        ' 判斷是否點選的是同一個列標題, 如果是，則切換排序方式, 否則預設使用升序排序
        currentSortOrder = If(e.Column = previousColumnIndex AndAlso currentSortOrder = SortOrder.Ascending, SortOrder.Descending, SortOrder.Ascending)
        previousColumnIndex = e.Column  ' 儲存目前點選的列索引

        ' by Gemini 2026/4/10, Listview3改virtual mode
        'ListView3.ListViewItemSorter = New ListViewItemComparer(e.Column, currentSortOrder)    
        ListView3.BeginUpdate()
        Try
            ' 直接對底層資料源進行排序，不操作 UI 物件
            Select Case e.Column
                Case 0 ' 主旨
                    If currentSortOrder = SortOrder.Ascending Then _lv3MailList = _lv3MailList.OrderBy(Function(x) x.Subject).ToList() Else _lv3MailList = _lv3MailList.OrderByDescending(Function(x) x.Subject).ToList()
                Case 1 ' 大小
                    If currentSortOrder = SortOrder.Ascending Then _lv3MailList = _lv3MailList.OrderBy(Function(x) x.Size).ToList() Else _lv3MailList = _lv3MailList.OrderByDescending(Function(x) x.Size).ToList()
                Case 2 ' 時間
                    If currentSortOrder = SortOrder.Ascending Then _lv3MailList = _lv3MailList.OrderBy(Function(x) x.ReceivedTime).ToList() Else _lv3MailList = _lv3MailList.OrderByDescending(Function(x) x.ReceivedTime).ToList()
                Case 3 ' 寄件者
                    If currentSortOrder = SortOrder.Ascending Then _lv3MailList = _lv3MailList.OrderBy(Function(x) x.SenderName).ToList() Else _lv3MailList = _lv3MailList.OrderByDescending(Function(x) x.SenderName).ToList()
                Case 4 ' 附件數 (依 EntryID 從快取抓取，模擬原 Compare 邏輯)
                    If currentSortOrder = SortOrder.Ascending Then
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
        ProgressBar2.Text = $"虛擬排序 {_lv3MailList.Count} 項，耗時 {sw.Elapsed.TotalSeconds:0.00} 秒"
        _dbg("結束", "排序列表") ' by Gemini, 2026/04/10

    End Sub
    Private Sub ListView3_MouseClick(sender As Object, e As MouseEventArgs) Handles ListView3.MouseClick
        ' 單擊左鍵 → 複製郵件主旨到剪貼簿 (方便貼到搜尋欄或筆記)
        ' by Gemini, 2026/04/10: 虛擬模式下 index 对应 _lv3MailList
        Dim itemInfo As MailItemInfo? = GetItemFromPoint(e.Location)
        If itemInfo IsNot Nothing AndAlso e.Button = MouseButtons.Left Then Clipboard.SetText(itemInfo.Value.Subject)
    End Sub
    Private Sub ListView3_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles ListView3.MouseDoubleClick
        ' 雙擊左鍵 → 複製主旨 + 用 EntryID 在 Outlook 中打開郵件
        ' by Gemini, 2026/04/10: 虛擬模式下 index 对应 _lv3MailList
        Dim itemInfo As MailItemInfo? = GetItemFromPoint(e.Location)
        If itemInfo Is Nothing OrElse e.Button <> MouseButtons.Left Then Return

        ' Clipboard.SetText(itemInfo.Value.Subject) ' 先複製主旨 (第 1 欄 SubItems(0)
        OpenMailByEntryID(itemInfo.Value.EntryID) ' 用 EntryID 打開郵件 (第 6 欄 SubItems(5))
    End Sub
#End Region
#Region "  ├ Layer2 流程協調層"
    Private Function FilterBySize(sourceList As List(Of MailItemInfo)) As List(Of MailItemInfo)
        ' Pipeline 過濾 1: 大小篩選 (純記憶體 LINQ，速度極快)
        _dbg("開始")
        Dim minSz = CLng(NumberMin.Value * GetSizeMultiplier(UnitMin.SelectedItem.ToString))
        Dim maxSz = CLng(NumberMax.Value * GetSizeMultiplier(UnitMax.SelectedItem.ToString))
        Dim resultList = sourceList.Where(Function(c) c.Size >= minSz AndAlso c.Size <= maxSz).ToList()
        _dbg("結束", $"篩選後剩下 {resultList.Count} 封")
        Return resultList
    End Function
    Private Async Function FilterByAttachDetailsAsync(sourceList As List(Of MailItemInfo), progress As IProgress(Of ProgressReport), cToken As CancellationToken) As Task(Of List(Of MailItemInfo))
        ' Pipeline 過濾 2: 逐一讀取附件明細，利用 _cacheAttachFilename 大幅降低 COM 存取
        _dbg("開始", $"候選郵件: {sourceList.Count} 封")

        ' by Gemini: Layer2 業務層向 Layer2.5 請求平行預載快取。若 RDO 存在，這行能在極短時間內把後續需要的資料全數載入記憶體。
        If RDO_Parallel1.Checked Then
            Await PreloadAttachByRDOAsync1(sourceList, progress, cToken)    ' by Parellel.ForEach 來平行讀取附件資料，適合 CPU 密集型的 MAPI 存取
        ElseIf RDO_Parallel2.Checked Then
            Await PreloadAttachByRDOAsync2(sourceList, progress, cToken)    ' by Task.WhenAll 來平行讀取附件資料，適合 I/O 等待型的資料庫存取
        End If
        ' 假設您現在的流程是：「讀取大量郵件屬性 --> 與本地資料庫比對快取 --> 寫入資料庫」
        ' 1. 讀取 MAPI 資料 (CPU + 嚴格 Thread 限制)：使用 Parallel.ForEach。
        '    因為您需要真實在多個核心上建立獨立的 RDOSession 來平行榨取硬碟與 MAPI 引擎的讀取速度
        ' 2. 查詢/寫入本地資料庫快取 (I/O 等待)：使用 Task.WhenAll + async。
        '    如果您的底層資料庫驅動（例如 SQLite-net 或 Entity Framework）支援原生的非同步方法
        '    (ToListAsync, ExecuteAsync)

        Dim swTotal As New Stopwatch() : swTotal.Start()
        Dim swThrottle As New Stopwatch() : swThrottle.Start()

        Dim mustCountAttach As Boolean = CheckAttCount.Checked
        Dim minCount As Integer = If(mustCountAttach, CInt(CountMin.Value), 0)
        Dim maxCount As Integer = If(mustCountAttach, CInt(CountMax.Value), Integer.MaxValue)

        Dim processed As Integer = 0, total As Integer = sourceList.Count
        Dim resultList As New List(Of MailItemInfo)
        Dim keyword As String = If(CheckAttachName.Checked, TextBox3.Text.Trim.ToLower(), "")
        Try
            For curMail As Integer = 0 To sourceList.Count - 1
                ' 2026/4/5, by Gemini: 將進度報告與 UI 釋放移至迴圈開頭，提早反饋處理進度
                ' 避免被下方的 Guard Clauses (Continue For) 略過而導致長時間霸佔主執行緒, 未更新UI進度反饋
                processed = curMail + 1
                If swThrottle.ElapsedMilliseconds >= 100 OrElse processed = total Then
                    Dim elapsedSec As Double = Math.Max(swTotal.Elapsed.TotalSeconds, 0.001)
                    Dim speed As Double = processed / elapsedSec

                    Dim etaString As String = ""
                    If total > 1000 AndAlso speed > 0 Then
                        Dim remainingSec As Integer = CInt(Math.Max(0, (total - processed) / speed))
                        If remainingSec > 3 Then etaString = $"，預估剩餘 {remainingSec \ 60:D2}:{remainingSec Mod 60:D2}"
                    End If

                    progress?.Report(New ProgressReport With {.CurrentCount = processed, .TotalCount = total,
                                                              .Message = $"Phase 2: {processed} / {total}，已符合 {resultList.Count} 封 ({speed:F0} 封/秒{etaString})"})
                    swThrottle.Restart()
                    Await Task.Delay(1, cToken)
                End If

                Dim currentMail As MailItemInfo = sourceList(curMail)
                Dim cachedAttFilenames As List(Of String) = GetCachedAttachFilename(currentMail)

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
        _dbg("結束", $"Phase 2 完成，篩選後共 {resultList.Count} 封")
        Return resultList

    End Function
    Private Sub ShowResultTab3(sourceList As List(Of MailItemInfo), elapsedSeconds As Double)
        _dbg("開始", sourceList.Count)
        ' by Gemini, 2026/04/10: 虛擬模式下僅需同步資料與設定 Size，完全不需建立物件
        _lv3MailList = sourceList
        ListView3.VirtualListSize = _lv3MailList.Count
        ListView3.Invalidate() ' 強制重繪

        ' by Gemini 2026/4/10, Listview3改virtual mode, 下面的實體項目建立邏輯已完全移除，改由 RetrieveVirtualItem 事件按需生成
        ' ListView3.Items.Clear()
        ' ' 先告訴 ListView 總共會有幾筆，讓它一次配置好記憶體，不要每次 Add 都 realloc
        ' Dim lviCount As Integer = sourceList.Count
        ' If lviCount > 50 Then SendMessage(ListView3.Handle, LVM_SETITEMCOUNT, New IntPtr(lviCount), IntPtr.Zero)
        ' If lviCount > 10 Then ListView3.BeginUpdate()
        ' If lviCount > 0 Then
        '     Dim items As New List(Of ListViewItem)(lviCount)
        '     For Each mail As MailItemInfo In sourceList
        '         ' 如果快取區存有這封信真的附件數量，就顯示明確數量。如果沒有，代表完全沒有跑過 Phase 2，就顯示 ">0", 不需真的去讀COM物件確認，避免不必要的性能損耗
        '         Dim cachedFiles As List(Of String) = Nothing
        '         Dim displayName As String = ">0"
        '         If _cacheAttachFilename.TryGetValue(mail.EntryID, cachedFiles) Then displayName = cachedFiles.Count.ToString()
        '         items.Add(New ListViewItem({mail.Subject,
        '                                     mail.Size.ToString("###,###,##0"),
        '                                     mail.ReceivedTime.ToShortDateString(),
        '                                     mail.SenderName,
        '                                     displayName,
        '                                     mail.EntryID}))
        '     Next
        '     ListView3.Items.AddRange(items.ToArray())
        ' Else
        '     ListView3.Items.Add("找不到符合條件的郵件")
        ' End If
        ' ListView3.EndUpdate()

        ProgressBar2.Text = ""

        ' lviCount 避免除以零
        Dim lviCount As Integer = sourceList.Count
        Dim speedText As String = ""
        If elapsedSeconds > 0 AndAlso lviCount > 0 Then speedText = $" ({CInt(lviCount / elapsedSeconds):###,##0}/sec)"
        ProgressBar1.Text = $"共找到 {lviCount} 封 / 耗時 {elapsedSeconds:0.00} 秒{speedText}"
        _dbg("結束", $"{lviCount} 封 | {elapsedSeconds:0.00}s")

    End Sub
#End Region
#Region "  └ 輔助函數"
    Private Function BuildFilterAttachment() As String
        ' 2026-03-16: 大小篩選移到 Button3_Click Step3b 的 LINQ，
        '   此函數保留但現在只回傳 hasattachment 基礎 filter (與 strFilterHasAttachment 一致)
        '   保留原有大小條件建構邏輯以備日後參考，但 Button3_Click 已不呼叫此函數
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
    Private Function GetItemFromPoint(pt As System.Drawing.Point) As MailItemInfo?
        ' 輔助方法：從點座標找出虛擬項目的資料
        Dim item = ListView3.GetItemAt(pt.X, pt.Y)
        If item IsNot Nothing AndAlso item.Index >= 0 AndAlso item.Index < _lv3MailList.Count Then
            Return _lv3MailList(item.Index)
        End If
        Return Nothing
    End Function
    Private Sub OpenMailByEntryID(strEntryID As String)
        ' 依照傳入的Mailitem's EntryID, 呼叫NameSpace打開郵件再釋放object
        _dbg("開始", strEntryID)
        If strEntryID Is Nothing Then Return        '_dbg("結束", "EntryID 為空")

        'Dim currentMail As MailItem = Nothing
        'Try
        '    currentMail = CType(objNameSpace.GetItemFromID(strEntryID), MailItem)
        '    currentMail.Display()
        'Catch ex As System.Exception
        '    MessageBox.Show("無法開啟郵件: " & ex.Message)
        'Finally
        '    If currentMail IsNot Nothing Then Marshal.ReleaseComObject(currentMail)   ' ✅ MailItem 釋放
        'End Try

        ' 2026/3/20, by Claude.ai, 建立獨立執行緒fire-and-forget
        ' 讓作業系統跟outlook.exe 自己去做它們的事, 我們不用等它開啟完畢, 可以直接回到自己的程式介面
        Dim ns As Outlook.NameSpace = Nothing
        Dim mail As MailItem = Nothing
        Dim th As New Thread(Sub()
                                 Try
                                     ns = _olApp.GetNamespace("MAPI")
                                     mail = CType(ns.GetItemFromID(strEntryID), MailItem)
                                     mail.Display()
                                 Catch ex As System.Exception
                                     ' 開視窗失敗，靜默忽略 (或 BeginInvoke 到 UI 執行緒顯示 MessageBox)
                                     'MessageBox.Show("無法開啟郵件: " & ex.Message)
                                 Finally
                                     TryMarshalRelease(mail)
                                     TryMarshalRelease(ns)       ' ✅ 補上
                                 End Try
                             End Sub)

        th.SetApartmentState(ApartmentState.STA)    ' ✅ 新執行緒設 STA，COM 呼叫合法
        th.IsBackground = True                      ' ✅ 主程式關閉時不等這條執行緒
        th.Start()                                  ' ✅ fire-and-forget，直接 return，主程式UI 立刻恢復回應

        _dbg("結束", "開啟郵件執行緒已啟動")

    End Sub
#End Region
#End Region

#Region "■ 07 Tab4: 系列郵件"
    Private Async Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        _dbg("開始")
        Dim rootFolder As Outlook.Folder = TryCast(TreeView1.SelectedNode?.Tag, Outlook.Folder)
        If rootFolder Is Nothing Then
            _dbg("結束", "未選擇資料夾")
            MessageBox.Show("請先在左側 Tab1 選擇要掃描的資料夾", "提示")
            Return
        End If
        Button4.Enabled = False
        Cursor = Cursors.WaitCursor
        TreeView4.Nodes.Clear()
        ListView4.Items.Clear()
        ProgressBar1.Text = "正在處理..."
        ProgressBar2.Text = "開始掃描系列郵件..."
        Dim sw As New Stopwatch() : sw.Start()
        Dim progress4 As IProgress(Of ProgressReport) = New Progress(Of ProgressReport)(Sub(p) ProgressBar2.Text = p.Message)
        Dim swThrottle As New Stopwatch() : swThrottle.Start() ' by Gemini, 2026/04/02: 重用秒錶做節流
        Dim topicDict As New Dictionary(Of String, List(Of MailItemInfo))(StringComparer.OrdinalIgnoreCase)
        Try
            ' 取得所有子資料夾 (Layer3 展開)
            Dim targetFolderList As List(Of Outlook.Folder) = Await GetSubtreeToList(rootFolder, includeSubF:=True, progress:=progress4)  ' by Claude, 2026/04/11: Async 化後加 Await
            Dim processed As Integer = 0
            For Each folder In targetFolderList
                Dim table As Outlook.Table = Nothing
                Try
                    table = folder.GetTable()
                    Const PR_MESSAGE_SIZE As String = "http://schemas.microsoft.com/mapi/proptag/0x0E080003"
                    Const PR_CONVERSATION_TOPIC As String = "http://schemas.microsoft.com/mapi/proptag/0x0070001E"
                    table.Columns.RemoveAll()
                    table.Columns.Add("EntryID")
                    table.Columns.Add("Subject")
                    table.Columns.Add(PR_MESSAGE_SIZE)
                    table.Columns.Add("ReceivedTime")
                    table.Columns.Add("SenderName")
                    table.Columns.Add(PR_CONVERSATION_TOPIC)
                    Do While Not table.EndOfTable
                        Dim arr As Object = table.GetArray(1000)
                        If arr Is Nothing Then Exit Do
                        Dim data(,) As Object = DirectCast(arr, Object(,))
                        For r As Integer = 0 To data.GetUpperBound(0)
                            Dim topic As String = SafeGet(Of String)(data, r, 5, "")
                            If topic = "" Then Continue For ' 沒有 Conversation Topic 的信件略過
                            Dim entryID As String = SafeGet(Of String)(data, r, 0, "")
                            If entryID = "" Then Continue For
                            Dim info As New MailItemInfo With {.EntryID = entryID,
                                                               .Subject = SafeGet(Of String)(data, r, 1, ""),
                                                               .Size = SafeGet(Of Long)(data, r, 2, 0L),
                                                               .ReceivedTime = SafeGet(Of DateTime)(data, r, 3, DateTime.MinValue),
                                                               .SenderName = SafeGet(Of String)(data, r, 4, "")}
                            If Not topicDict.ContainsKey(topic) Then topicDict(topic) = New List(Of MailItemInfo)()
                            topicDict(topic).Add(info)
                        Next
                    Loop
                Catch ex As System.Exception
                    _dbg("錯誤", $"{folder.Name}: {ex.Message}") ' by Gemini, 2026/04/04: Issue 4 格式標準化
                Finally
                    TryMarshalRelease(table)
                End Try
                swThrottle.Start() ' by Gemini, 2026/04/02: 重用秒錶做節流
                processed += 1

                ' by Gemini, 2026/04/02: 100ms 節流回報 (標準化 IProgress Pattern)
                If progress4 IsNot Nothing AndAlso (swThrottle.ElapsedMilliseconds >= 100 OrElse processed = targetFolderList.Count) Then
                    progress4.Report(New ProgressReport With {.CurrentCount = processed, .TotalCount = targetFolderList.Count,
                                                              .Message = $"正在掃描系列郵件: {processed} / {targetFolderList.Count} 個資料夾..."})
                    swThrottle.Restart()
                    Await Task.Yield()
                End If
            Next

            ' 將結果加入 TreeView4 (只加數量 > 1 的)
            TreeView4.BeginUpdate()
            Dim nodesProcessed As Integer = 0
            For Each kvp In topicDict
                If kvp.Value.Count > 1 Then
                    Dim node As New TreeNode($"{kvp.Key} ({kvp.Value.Count})") With {.Tag = kvp.Value}  ' 存入 List(Of MailItemInfo)
                    TreeView4.Nodes.Add(node)
                    nodesProcessed += 1

                    ' by Gemini, 2026/04/02: 長列表建構時也需節流 (標準化 IProgress)
                    If swThrottle.ElapsedMilliseconds >= 100 Then
                        progress4.Report(New ProgressReport With {.Message = $"正在建立系列清單: {nodesProcessed} 組..."})
                        swThrottle.Restart()
                        Await Task.Delay(1)
                    End If
                End If
            Next
            TreeView4.EndUpdate()
            sw.Stop()
            ProgressBar1.Text = $"找到 {TreeView4.Nodes.Count} 個系列 / 耗時 {sw.Elapsed.TotalSeconds:0.00} 秒"
            ProgressBar2.Text = ""
        Catch ex As System.Exception
            MessageBox.Show("掃描系列郵件時發生錯誤: " & ex.Message, "錯誤")
            _dbg("錯誤", ex.Message)
        Finally
            Button4.Enabled = True
            Cursor = Cursors.Default
            _dbg("結束")
        End Try

    End Sub
    Private Sub TreeView4_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TreeView4.AfterSelect
        _dbg("開始", e.Node.Text)
        Dim mailList As List(Of MailItemInfo) = TryCast(e.Node.Tag, List(Of MailItemInfo))
        If mailList Is Nothing Then Return        '_dbg("結束", "標記資料為空")

        ' 排序: 依據時間遞減 (越新的在越前面)
        mailList.Sort(Function(a, b) b.ReceivedTime.CompareTo(a.ReceivedTime))
        ListView4.BeginUpdate()
        ListView4.Items.Clear()
        For Each mailItem In mailList
            Dim lvi As New ListViewItem({
                mailItem.Subject,
                (mailItem.Size \ 1024L).ToString("###,###,###,##0") & "KB",
                mailItem.ReceivedTime.ToString("yyyy/MM/dd HH:mm:ss"),
                mailItem.SenderName,
                mailItem.EntryID
            })
            ListView4.Items.Add(lvi)
        Next
        ListView4.EndUpdate()
        _dbg("結束", $"顯示 {mailList.Count} 封系列郵件")

    End Sub
    Private Sub ListView4_MouseClick(sender As Object, e As MouseEventArgs) Handles ListView4.MouseClick
        ' ── ListView4 滑鼠事件 (Tab4 系列郵件搜尋結果) ──
        ' 2026-03-16 新增: 單擊/雙擊都複製主旨到剪貼簿
        _dbg("開始")
        Dim item As ListViewItem = sender.GetItemAt(e.X, e.Y)
        If item IsNot Nothing AndAlso e.Button = MouseButtons.Left Then Clipboard.SetText(item.SubItems(0).Text)
        _dbg("結束")

    End Sub
    Private Sub ListView4_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles ListView4.MouseDoubleClick
        _dbg("開始")
        Dim item As ListViewItem = sender.GetItemAt(e.X, e.Y)
        If item IsNot Nothing AndAlso e.Button = MouseButtons.Left Then Clipboard.SetText(item.SubItems(0).Text)
        _dbg("結束")

    End Sub
#End Region

#Region "■ 08 Tab5: 重複郵件"
    Private Async Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        _dbg("開始")
        If _pstStoreList Is Nothing OrElse _pstStoreList.Count = 0 Then
            _dbg("結束", "PST 尚未載入")
            MessageBox.Show("PST 檔案庫尚未載入完成，請稍後再試", "提示")
            Return
        End If
        Button5.Enabled = False
        Cursor = Cursors.WaitCursor
        ListView5.BeginUpdate()
        ListView5.Items.Clear()
        ListView5.EndUpdate()
        ProgressBar1.Text = "正在準備"
        ProgressBar2.Text = "準備全信箱掃描重複郵件..."
        Dim sw As New Stopwatch() : sw.Start()
        Dim progress5 As IProgress(Of ProgressReport) = New Progress(Of ProgressReport)(Sub(p) ProgressBar2.Text = p.Message)
        Dim swThrottle As New Stopwatch() : swThrottle.Start() ' by Gemini, 2026/04/02
        Dim exactDict As New Dictionary(Of String, List(Of MailItemInfo))(StringComparer.OrdinalIgnoreCase)
        Dim isExact As Boolean = rbExactMatch.Checked
        Try
            ' 遍歷所有 Store

            ' 遍歷所有 Store
            Dim totalProcessed As Integer = 0
            For Each store In _pstStoreList
                If _cancelRequested Then Exit For
                Try
                    Dim rootFolder As Outlook.Folder = store.GetRootFolder()
                    Dim targetFolderList As List(Of Outlook.Folder) = Await GetSubtreeToList(rootFolder, includeSubF:=True, progress:=progress5)  ' by Claude, 2026/04/11: Async 化後加 Await
                    For Each folder In targetFolderList
                        If _cancelRequested Then Exit For
                        Dim table As Outlook.Table = Nothing
                        Try
                            table = folder.GetTable()
                            Const PR_MESSAGE_SIZE As String = "http://schemas.microsoft.com/mapi/proptag/0x0E080003"
                            table.Columns.RemoveAll()
                            table.Columns.Add("EntryID")
                            table.Columns.Add("Subject")
                            table.Columns.Add(PR_MESSAGE_SIZE)
                            table.Columns.Add("ReceivedTime")
                            table.Columns.Add("SenderName")
                            Do While Not table.EndOfTable
                                Dim arr As Object = table.GetArray(1000)
                                If arr Is Nothing Then Exit Do
                                Dim data(,) As Object = DirectCast(arr, Object(,))
                                For r As Integer = 0 To data.GetUpperBound(0)
                                    Dim entryID As String = SafeGet(Of String)(data, r, 0, "")
                                    If entryID = "" Then Continue For
                                    Dim subject As String = SafeGet(Of String)(data, r, 1, "")
                                    Dim size As Long = SafeGet(Of Long)(data, r, 2, 0L)
                                    Dim recvTime As DateTime = SafeGet(Of DateTime)(data, r, 3, DateTime.MinValue)
                                    Dim senderName As String = SafeGet(Of String)(data, r, 4, "")
                                    Dim info As New MailItemInfo With {.EntryID = entryID,
                                                                       .Subject = subject,
                                                                       .Size = size,
                                                                       .ReceivedTime = recvTime,
                                                                       .SenderName = senderName}
                                    Dim hashKey As String
                                    If isExact Then
                                        hashKey = $"{subject}|{size}|{recvTime:yyyyMMddHHmmss}|{senderName}"
                                    Else
                                        Dim cleanSubj As String = subject.ToUpper().Replace("RE:", "").Replace("FW:", "").Replace("回覆:", "").Replace("轉寄:", "").Replace(" ", "").Trim()
                                        If cleanSubj.Length > 20 Then cleanSubj = cleanSubj.Substring(0, 20)
                                        hashKey = $"{cleanSubj}|{size}"
                                    End If

                                    If Not exactDict.ContainsKey(hashKey) Then exactDict(hashKey) = New List(Of MailItemInfo)()
                                    exactDict(hashKey).Add(info)
                                Next
                                Await Task.Yield()
                            Loop
                        Catch ex As System.Exception
                            _dbg("錯誤", $"{folder.Name}: {ex.Message}") ' by Gemini, 2026/04/04: Issue 4 格式標準化
                        Finally
                            TryMarshalRelease(table)
                        End Try
                        totalProcessed += 1

                        ' by Gemini, 2026/04/02: 100ms 節流回報 (標準化 IProgress)
                        If swThrottle.ElapsedMilliseconds >= 100 Then
                            progress5.Report(New ProgressReport With {.Message = $"掃描中 ({store.DisplayName}): 已處理 {totalProcessed} 個資料夾..."})
                            swThrottle.Restart()
                            Await Task.Yield()
                        End If
                    Next
                Catch ex As System.Exception
                    _dbg("錯誤", $"{store.DisplayName}: {ex.Message}") ' by Gemini, 2026/04/04: Issue 4 格式標準化
                End Try
            Next
            ' 尋找符合條件的群組
            ListView5.BeginUpdate()
            Dim groupID As Integer = 1
            Dim totalDuplicateMails As Integer = 0
            Dim swThrottleBuild As New Stopwatch() : swThrottleBuild.Start() ' by Gemini, 2026/04/02
            For Each kvp In exactDict
                If kvp.Value.Count > 1 Then
                    Dim isValidGroup As Boolean = True
                    ' 若是 Fuzzy 模式，還需確認 Levenshtein 距離不超過門檻 (至少大於 0.8 相似度)
                    If Not isExact Then
                        Dim firstSubject As String = kvp.Value(0).Subject.ToUpper()
                        For i As Integer = 1 To kvp.Value.Count - 1
                            Dim sim As Double = CalculateSimilarity(firstSubject, kvp.Value(i).Subject.ToUpper())
                            If sim < 0.8 Then
                                isValidGroup = False
                                Exit For
                            End If
                        Next
                    End If
                    If isValidGroup Then
                        Dim groupColor As Color = If(groupID Mod 2 = 0, Color.FromArgb(240, 248, 255), Color.White)
                        For Each mailItem In kvp.Value
                            Dim lvi As New ListViewItem({mailItem.Subject,
                                                        (mailItem.Size \ 1024L).ToString("###,###,###,##0") & "KB",
                                                         mailItem.ReceivedTime.ToString("yyyy/MM/dd HH:mm:ss"),
                                                         mailItem.SenderName,
                                                         "群組 " & groupID.ToString(),
                                                         mailItem.EntryID}) With {.BackColor = groupColor}
                            ListView5.Items.Add(lvi)
                            totalDuplicateMails += 1
                        Next
                        groupID += 1

                        ' by Gemini, 2026/04/02: 100ms 節流回報 (標準化 IProgress)
                        If swThrottleBuild.ElapsedMilliseconds >= 100 Then
                            progress5.Report(New ProgressReport With {.Message = $"正在建立重複郵件清單: {groupID} 組..."})
                            swThrottleBuild.Restart()
                            Await Task.Delay(1)
                        End If
                    End If
                End If
            Next
            ListView5.EndUpdate()
            sw.Stop()
            ProgressBar1.Text = $"找到 {groupID - 1} 組 ({totalDuplicateMails} 封) / 耗時 {sw.Elapsed.TotalSeconds:0.00} 秒"
            ProgressBar2.Text = ""
        Catch ex As System.Exception
            MessageBox.Show("掃描重複郵件時發生錯誤: " & ex.Message, "錯誤")
            _dbg("錯誤", ex.Message)
        Finally
            Button5.Enabled = True
            Cursor = Cursors.Default
            _dbg("結束")
        End Try

    End Sub
    Private Function CalculateSimilarity(strA As String, strB As String) As Double
        ' by Gemini, 2026/04/04: Issue 1 移除 _dbg（Tab5 高頻呼叫，N封×2個函數=2N行輸出）
        ' 計算編輯距離
        Dim editDistance As Integer = LevenshteinDistance(strA, strB)

        ' 將編輯距離歸一化為範圍在 0 到 1 之間的值
        Dim maxLength As Integer = Math.Max(strA.Length, strB.Length)
        Dim similarity As Double = 1 - CDbl(editDistance) / maxLength
        Return similarity

    End Function
    Private Function LevenshteinDistance(strA As String, strB As String) As Integer
        ' by Gemini, 2026/04/04: Issue 1 移除 _dbg（Tab5 高頻呼叫，同上）
        ' 計算 Levenshtein 編輯距離的輔助函數
        Dim lenA As Integer = strA.Length
        Dim lenB As Integer = strB.Length
        Dim distance(lenA, lenB) As Integer
        For i As Integer = 0 To lenA : distance(i, 0) = i : Next
        For j As Integer = 0 To lenB : distance(0, j) = j : Next
        For j As Integer = 1 To lenB
            For i As Integer = 1 To lenA
                '' 改前 (5行)
                'If strA(fd - 1) = strB(j - 1) Then
                '    distance(fd, j) = distance(fd - 1, j - 1)
                'Else
                '    distance(fd, j) = Math.Min(Math.Min(distance(fd - 1, j) + 1,
                '                                       distance(fd, j - 1) + 1), distance(fd - 1, j - 1) + 1)
                'End If

                ' 改後 (1行)
                distance(i, j) = If(strA(i - 1) = strB(j - 1),
                    distance(i - 1, j - 1), Math.Min(Math.Min(distance(i - 1, j) + 1,
                                                              distance(i, j - 1) + 1), distance(i - 1, j - 1) + 1))
            Next
        Next
        Return distance(lenA, lenB)

    End Function
#End Region

#Region "■ 09 Tab6: Debug & 設定"
    Private Sub CheckDebug_CheckedChanged(sender As Object, e As EventArgs) Handles CheckDebug.CheckedChanged
        _isDebugMode = CheckDebug.Checked
        _dbg("開始", _isDebugMode.ToString)
        Dim offset As Integer = If(CheckDebug.Checked, -240, 240)
        Me.Left += offset
        System.Windows.Forms.Cursor.Position = New Point(
            System.Windows.Forms.Cursor.Position.X + offset,
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
    Private Sub DebugButton_Click(sender As Object, e As EventArgs) Handles DebugButton.Click

        ' 測試 DASL 是否能在 GetTable 直接濾出含有特定附檔名的信件
        Dim folder As Outlook.Folder = TryCast(SimTree3.SelectedNode.Tag, Outlook.Folder)
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
    Private Async Sub CacheSnifferAsync(cToken As System.Threading.CancellationToken)
        ' === CacheSniffer — 背景快取預讀系統 (B4) ===
        ' ===============================================================================
        ' 職責: 程式啟動後在背景靜默預讀 Tab1 / Tab2 / Tab3 ，快取後讓使用者點選時直接從記憶體讀取，不再等待 COM 查詢。
        '
        ' 設計原則:
        '   - 廣度優先 (BFS) : 淺層資料夾優先預讀，使用者最常點選的位置最先就緒
        '   - 固定 1 秒間隔: 每完成一個資料夾的三項快取，固定等 1 秒再繼續，讓 Outlook 有充足空閒時間回應使用者互動
        '   - COM 全在 UI 執行緒 (STA) : 所有 Await 都不切執行緒，不需要 Task.Run
        '   - CancellationToken: FormClosing 時呼叫 _cacheSnifferCts.Cancel()，確保程式關閉後不留殘餘 COM 呼叫
        '   - 快取命中就跳過: 若使用者已先點選觸發過快取，CacheSniffer 直接略過不重做
        '   - 停用方式: 把 Form1_Load 末尾的 CacheSnifferAsync(...) 那行加上 ' 即可，其餘程式碼完全不受影響
        '
        ' 預讀順序 (每個資料夾) :
        '   1. Tab1: mailCountCache + folderCountCache (GetMailCountAll / GetTotalFolderCountAsync)
        '   2. Tab2: yearCountsCache (GetYearCountsForFolderAsync)
        '   3. Tab3: _cacheAttachMailList (CheckTab3CacheOrRescan)
        '
        ' 2026-03-16 B4 新增，由 PrewarmAllCachesAsync 重構整合，改名為 CacheSniffer
        '       只要偵測到正在進行 AfterSelect 或是正在跑複雜統計，就自動閉嘴等閒下來再繼續
        ' ===============================================================================

        'If _pstStoreList Is Nothing OrElse _pstStoreList.Count = 0 Then Return
        'Await Task.Delay(10000, cToken)      ' 等待 10 秒: 確保 Form1_Load 完全結束、UI 呈現完畢，再開始佔用 Outlook COM

        'Try
        '    _dbg("開始", "預讀快取")
        '    ' ── BFS 初始化: 把所有 PST 的第一層子資料夾加進佇列 ─────────
        '    ' 不從 root 本身開始，因為 root ("個人資料夾") 通常不含郵件，
        '    ' 直接從第一層子資料夾 (收件匣、寄件匣…) 開始
        '    Dim queue As New Queue(Of Outlook.Folder)
        '    For Each store As Outlook.Store In _pstStoreList
        '        If cToken.IsCancellationRequested Then Return
        '        For Each subFolder As Outlook.Folder In GetSortedSubFolders(store.GetRootFolder())
        '            queue.Enqueue(subFolder)
        '        Next
        '    Next
        '    ' ── BFS 主迴圈 ───────────────────────────────────────────────
        '    ' 每次取出一個資料夾，依序預讀 Tab1 / Tab2 / Tab3 的快取，
        '    ' 完成後把它的直屬子資料夾再放入佇列 (廣度優先，淺層先完成)
        '    Dim processed As Integer = 0
        '    While queue.Count > 0
        '        If cToken.IsCancellationRequested Then Return
        '        Dim folder As Outlook.Folder = queue.Dequeue()
        '        processed += 1
        '        ' ── Tab1: mailCountCache + folderCountCache ───────────────
        '        ' GetMailCountAll 和 GetTotalFolderCountAsync 內部各自寫入自己的快取
        '        ' 已命中的快取直接跳過，不重複呼叫 COM
        '        Try
        '            Await GetCachedMailCountAllAsync(folder)
        '            Await GetCachedFolderCountAllAsync(folder)
        '        Catch ex As System.Exception
        '            _dbg("CacheSniffer Tab1 Error: ", folder.Name & " - " & ex.Message)
        '        End Try
        '        If cToken.IsCancellationRequested Then Return
        '        ' ── Tab2: yearCountsCache ─────────────────────────────────
        '        ' GetYearCountsForFolderAsync 內部有快取命中判斷，已快取直接回傳
        '        Try
        '            Dim key As String = folder.FolderPath
        '            If Not _cacheYearCounts.ContainsKey(key) Then Await GetYearCountsForFolder(folder)
        '        Catch ex As System.Exception
        '            _dbg("CacheSniffer Tab2 Error: ", folder.Name & " - " & ex.Message)
        '        End Try
        '        If cToken.IsCancellationRequested Then Return
        '        ' ── Tab3: _cacheAttachMailList ────────────────────────────────
        '        ' CheckTab3CacheOrRescan 內部有 Items.Count 失效判斷
        '        Try
        '            Await CheckTab3CacheOrRescan(folder, Nothing)
        '        Catch ex As System.Exception
        '            _dbg("CacheSniffer Tab3 Error: ", folder.Name & " - " & ex.Message)
        '        End Try
        '        If cToken.IsCancellationRequested Then Return
        '        ' ── 固定 1 秒間隔: 讓 Outlook 保持回應能力 ───────────────
        '        _dbg($"CacheSniffer: [{processed}] {folder.Name} 完成，等 1 秒")
        '        Await Task.Delay(1000, cToken)
        '        Await Task.Yield()
        '        ' ── 把直屬子資料夾加入佇列 (廣度優先) ────────────────────
        '        ' GetSortedSubFolders 有 folderTreeCache，不重打 COM
        '        Try
        '            For Each subFolder As Outlook.Folder In GetSortedSubFolders(folder)
        '                queue.Enqueue(subFolder)
        '            Next
        '        Catch ex As System.Exception
        '            _dbg("錯誤", folder.Name & " - " & ex.Message)
        '        End Try
        '    End While
        '    _dbg("結束", $"預讀完成 | 總計: {processed} 個資料夾")
        'Catch ex As System.Threading.Tasks.TaskCanceledException
        '    _dbg("CacheSniffer: 已取消 (FormClosing) ")
        'Catch ex As System.Exception
        '    _dbg("錯誤", ex.Message)
        'Finally
        '    _dbg("結束")
        'End Try

    End Sub

    ' todo: 讀入OST檔案的功能
    Private Sub OST_Click(sender As Object, e As EventArgs)
        'Dim outlookApp As Outlook.Application = Nothing
        'Dim ns As Outlook.NameSpace = Nothing
        'Dim inbox As Outlook.Folder = Nothing
        'Try
        '    ReadEmailsFromOST("D:\Users\Simon\Documents\Outlook 檔案\Work\Inbox_2011_GLI.ost")
        'Finally
        '    If inbox IsNot Nothing Then Marshal.ReleaseComObject(inbox)
        '    If ns IsNot Nothing Then Marshal.ReleaseComObject(ns)
        '    If outlookApp IsNot Nothing Then Marshal.ReleaseComObject(outlookApp)
        'End Try

    End Sub
    Private Sub ReadEmailsFromOST(path As String)
        _dbg("開始")
        ' 創建 Outlook 應用程序對象
        Dim outlookApp As New Outlook.Application()
        ' 獲取 Outlook 命名空間
        Dim ns As Outlook.NameSpace = outlookApp.GetNamespace("MAPI")
        ' 添加本地 OST 文件
        ns.AddStore(path)
        ' 獲取默認的收件箱文件夾
        Dim inbox As Outlook.Folder = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox)
        ' 讀取郵件
        ReadFolderEmails(inbox)
        ' 釋放 COM 對象
        TryMarshalRelease(inbox)
        TryMarshalRelease(ns)
        TryMarshalRelease(outlookApp)

    End Sub
    Private Sub ReadFolderEmails(folder As Outlook.Folder)
        ' 迭代郵件項
        For Each item As Object In folder.Items
            If TypeOf item Is Outlook.MailItem Then
                Dim mail As Outlook.MailItem = CType(item, Outlook.MailItem)
                ' 在這裡處理郵件，比如顯示主題和內容等
                MessageBox.Show($"Subject: {mail.Subject}, Received: {mail.ReceivedTime}")
                TryMarshalRelease(mail)
            End If
        Next
        _dbg("結束")

    End Sub


#End Region


End Class
