Imports System.Collections.Concurrent
Imports System.Threading
Imports System.Windows.Forms.DataVisualization.Charting
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Outlook

Partial Class Form1

#Region "■ 01 全域宣告"
    Private ctxMenuLv1 As ContextMenuStrip

    Private _tab1SelectSeq As Integer = 0                       ' Tab1 快速點選防護序號
    Private _tab2SelectSeq As Integer = 0                       ' Tab2 快速點選防護序號
    'Private _tab2TotalMailCount As Long = 0                    ' 快取 _tv2FolderList 的總郵件數，省去切換月份時呼叫 GetMailCount() 算進度分母
    Private _tv2FolderPaths As List(Of String) = Nothing        ' 快取 _tv2FolderList 對應的 FolderPath，省去切換月份時的 COM 讀取
    Private _lv2IsMonthView As Boolean = False                  ' 目前 ListView2 顯示的是月份視圖還是年度視圖
    Private _lv2MonthViewYear As Integer = 0                    ' 目前月份視圖顯示的是哪一年

    Private _tv2FolderList As List(Of (eid As String, sid As String, fPath As String)) = Nothing    ' 記住目前 Tab2 的資料夾清單，供月份展開使用 (2026/06/28 Stage2: 改帶 eid/sid 不帶 COM)
    Private _lv2DataYear As ConcurrentDictionary(Of Integer, Integer) = Nothing         ' Tab2 年度視圖 session 快取 (已合併多資料夾)，月份進出時直接 render 不重算
    Private _lv2DataMonth As ConcurrentDictionary(Of Integer, Integer) = Nothing        ' Tab2 月份視圖 session 快取 (已合併多資料夾)；_lv2MonthViewYear 記錄對應年份 (方案A：單一變數) 

    Private Structure ProgressReport            ' by Gemini, 2026/04/02: 統一進度回報結構，用於 IProgress(Of T)
        Dim CurrentCount As Integer             ' 目前完成數 (郵件數、資料夾數或位元組)
        Dim TotalCount As Integer               ' 總數 (分母)
        Dim Message As String                   ' 顯示在狀態列的文字
        Dim IsIndeterminate As Boolean          ' 是否為不確定的進度 (跑馬燈模式)
    End Structure
    Private Class FolderBfsEntry                ' 候選待掃瞄剪枝的資料夾結構
        Public Folder As Outlook.Folder         ' 2026/06/29 by Simon/Claude [Option A1]: 走樹階段保持 Nothing(零物化)，GetBfsResult 才對 root+直屬子夾物化
        Public Eid As String                    ' 2026/06/29 by Simon/Claude [Option A1]: 身分證 — id-tuple BFS 走 DB 不物化 COM
        Public Sid As String                    ' 2026/06/29 by Simon/Claude [Option A1]: 身分證
        Public ParentIndex As Integer           ' -1 = rootFolder；>= 0 = 父節點在 allEntries 的索引
        Public DirectMailCount As Long          ' 本層郵件數 (不含子孫)，由 Layer3 填入
        Public TotalMailCount As Long           ' 含子孫郵件總數，Layer2 底部向上彙總後填入
        Public TotalSubCount As Long            ' 含子孫資料夾總數，Layer2 底部向上彙總後填入
        Public IsFromCache As Boolean           ' True = TotalMailCount/TotalSubCount 從快取取得，子樹已剪枝
        Public FolderPath As String             ' ✅ 新增：快取 FolderPath 避免後續重複呼叫 COM
    End Class
#End Region

#Region "■ 04 Tab1: 資料夾統計 — 重構後程式碼 v5 (最終版) ==="
    ' ==============================================================
    '   Layer1  TreeView1_AfterSelect         UI 事件層：序號防護 + 批次寫 ListView
    '   Layer2  CollectFolderInfoByBFS       流程協調層，拆成五個子函數：
    '             - BuildBfsFolderTree        GetSubtree骨架 + 記憶體剪枝 (2026-04-08 加入 DB lazy 不驗 snapshot；2026/07/02 骨架整合)
    '             - FetchDirectMailCountsAsync呼叫 GetMailCount (有記憶體+DB lazy+COM 三層) 
    '             - SumUpSubTreeBottomUp  純記憶體底部向上加總
    '             - UpdateFolderInfoCache    寫入 L2.5 快取字典
    '             - GetBfsResult              提取 root + 直屬子資料夾
    '   Layer3  GetMailCountOOM / GetFolderCountOOM 等 COM 底層
    '
    ' ── 版本演進摘要 ──────────────────────────────────────────────
    '
    '   原始版 循序 Await GetInfoForListview × N，各自等遞迴完成後才輪下一個
    '               A. 用 Task.Run 包 COM (STA 違規) + B. s4Task.Result 潛在 deadlock (cache: 0.10~0.19s)
    '
    '   v1  BFS 一次展開整棵子樹：
    '           GetMailCountOOM 循序讀 PR_CONTENT_COUNT，底部向上彙總後一次寫快取，
    '           之後點選子資料夾直接命中，架構最乾淨，但有 bug: root 快取命中時不展開子資料夾 → 第二次點選 ListView 只顯示 root 自身
    '           cache: 0.01s (最快，因為完全不碰 thread pool)
    '
    '   v2  Task.WhenAll 同時發起 N 個子資料夾的計算 (並行的並行)：
    '           修掉 s4Task.Result deadlock, 1st read 明顯變快；但 cache 仍有 40 次 Task.Run dispatch overhead (cache: 0.04~0.09s, 因 Task.Run overhead 限制)
    '
    '   v3  BFS + Task.WhenAll 試圖合併 v1 + v2 優點：
    '           但 ComputeFolderDisplayList 在 UI 執行緒循序走整棵子樹 → 更慢
    '   v3fix修正 v3 過深遍歷問題，ComputeFolderDisplayList 只收 depth=0/1：
    '           效能介於 v1 和 v2 之間，但仍有 Task.Run overhead (cache: 0.05~0.08s)
    '
    '   v4  v1 的 BFS 架構 + 一行 bug fix: root 永遠展開直屬子資料夾：
    '           保留 v1 的所有效能優勢，同時修正第二次點選只顯示 root 的問題
    '           不引入 Task.WhenAll (實測 sequential BFS 比 parallel of parallel 快) (cache: 0.01s，應當與 v1 相同)
    '
    '   v5  大幅重構 CollectFolderInfoByBFS，依「單一職責原則」拆分為五個子函數，確保各步驟隔離互不干擾。(2026/04/04 by Gemini)
    '
    ' ── 為什麼 v4 不用 Task.WhenAll？─────────────────────────────
    '   v2/v3fix 的「並行的並行」看起來應該更快，但實測反而輸給 v1，
    '   原因: PST 的 PR_CONTENT_COUNT 讀取是 COM overhead 主導 (不是 I/O bottleneck)
    '   v1 的 BFS sequential: N 個資料夾 × 1 PR_CONTENT_COUNT call = O(N)，無其他 overhead
    '   v2/v3fix 的 Task.WhenAll: 20 子資料夾 × 2 Task.Run = 40 次 thread pool dispatch，每次 dispatch ~1~2ms，40 次 = 40~80ms → 這就是 cache 0.05s 的來源
    '   PST 是單一檔案，並行讀取可能造成 I/O 競爭，在慢速 HDD 上優勢也有限 → v1 的 sequential BFS 在此場景下已是最優，不需要 Task.WhenAll  
    '
    ' ── 分層架構 ──────────────────────────────────────────────────
    '   Layer1  TreeView1_AfterSelect   UI 事件層
    '       取得選中資料夾 → 呼叫 Layer2 → 批次更新 ListView1
    '       規則: 不做計算，不直接操作 COM，只傳達意圖與呈現結果
    '
    '   Layer2  CollectFolderInfoByBFS 流程協調層 (核心)
    '       BFS 展開整棵子樹 (root 永遠展開直屬子，其餘節點依快取決定)
    '       → 呼叫 Layer3 讀每個節點的直接郵件數
    '       → 底部向上彙總 (O(N)，無遞迴 stack overflow 風險)
    '       → 一次性寫快取 (整棵子樹預讀)
    '       → 回傳 root + 直屬子資料夾清單供 Layer1 顯示
    '       回呼 onProgress 讓 Layer1 更新進度，Layer2 自身不碰任何 UI 控制項
    '
    '   Layer3  GetMailCountOOM            COM 資料層
    '       只讀單一資料夾的 PR_CONTENT_COUNT (本層郵件數，不含子孫)
    '       不遞迴，不展開子資料夾，最小化 COM 呼叫量
    '
    ' ── 快取策略 ──────────────────────────────────────────────────
    '   mailCountCache      → TotalMailCount (含子孫郵件總數)   Layer2 底部向上彙總後寫入，TryAdd 不覆蓋既有值
    '   folderCountCache    → TotalSubCount (含子孫資料夾總數)  Layer2 底部向上彙總後寫入，TryAdd 不覆蓋既有值
    '   folderSizeCache     → 資料夾大小 (Lazy，由 ColumnClick / 右鍵觸發計算)
    '   folderTreeCache     → 子資料夾排序清單 (GetSortedSubFolders 負責維護)
    ' ─────────────────────────────────────────────────────────────
    '   FolderBfsEntry: BFS 過程中每個資料夾節點的容器
    '   貫穿 Layer2 的所有步驟 (BFS 展開 → Layer3 讀取 → 底部向上彙總 → 快取寫入 → 回傳清單)
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
    ' delete: 2026/04/01 by simon, 直接從設計工具建立 SimTree 控制項到 Form1
    Private Async Sub SimTree1_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles SimTree1.AfterSelect
        ' ==============================================================
        ' === Layer 1 (UI 事件層) — SimTree1 多選版 ===
        ' 2026/04/13 by Simon/Claude: Tab1 由原生 TreeView1 升級為 SimTree1 (多選控制項) 
        '
        ' 職責: 讀取 SimTree1.SelectedNodes (多選清單) → 對每個選中的資料夾呼叫 Layer2
        '        → 組裝「群組標題行 + 直屬子資料夾行」清單 → 批次寫入 ListView1
        '
        ' 顯示格式 (統一，不論單選或多選) :
        '    ▸ 資料夾名稱 (群組標題行，粗體淡藍底，含該資料夾的完整統計數字) 
        '    - 子資料夾A
        '    - 子資料夾B  ...
        '    [多選時底部追加合計列]
        '
        ' Tag 結構 (ValueTuple) :
        '    群組標題行 & 合計列 → Tag = Nothing (EnterFolder / ComputeSize 看到 Nothing 直接跳過) 
        '    一般子資料夾行      → Tag = (SubFolder:=Outlook.Folder, ParentNode:=TreeNode)
        '                           ComputeSize 從 .SubFolder 取資料夾；EnterSelectedFolder 從 .ParentNode 找節點
        ' ==============================================================
        ' 重構抽離, 2026/5/14 by simon
        '   - SimTree.GetDedupedSelection : 父子去重，確保同時選中父資料夾與子資料夾時不重複計算, 可重複用於F5 Refresh
        '                                   (2026/07/10 由本檔 GetDeDupedNodes 內建進 SimTree，舊函數移至 Module_Win32API.vb 備用待刪區)
        '   - CollectTab1FolderInfo : 非同步計算統計數字，支援一般模式與 F5 強制刷新模式
        '   - RenderLv1             : 將計算好的 ListViewItem 清單更新至 ListView1，包含雙緩衝優化與過期狀態清理
        ' ==============================================================
        _dbg("開始") : Dim sw As Stopwatch = Stopwatch.StartNew()  ' by Claude Sonnet 4.6, 2026/06/07

        Dim selectedNodes As List(Of TreeNode) = SimTree1.SelectedNodes
        If selectedNodes.Count = 0 Then Return

        ' ── 父子去重 (2026/04/13 by Simon/Claude) ──────────────────────────
        ' 問題: 使用者同時選父資料夾與其子孫節點時，父的 TotalMailCount 已含子孫，若再重複計算則數字不準。
        ' 解法: 只保留「沒有任何祖先被選中」的節點。
        Dim deDupedNodes As List(Of TreeNode) = SimTree1.GetDedupedSelection()   ' 2026/07/10 by Simon/Claude: 改用 SimTree 內建版
        Dim skippedCount As Integer = selectedNodes.Count - deDupedNodes.Count
        If skippedCount > 0 Then _dbg(" ├ 父子去重", $"移除 {skippedCount:N0} 個子孫節點，實際處理 {deDupedNodes.Count:N0} 個")

        ' 非同步序列與 Token 管理
        Dim mySeq As Integer = System.Threading.Interlocked.Increment(_tab1SelectSeq)
        Dim cToken As CancellationToken = OkayNowYouHaveToken()

        _isUserBusy = True : Cursor = Cursors.WaitCursor
        PgrsBar1.Text = "" : PgrsBar2.Text = ""

        Try
            ' ── 執行運算 (Layer 1.5) ──
            ' 傳入去重後的節點，取得組裝好的 ListViewItem 清單
            Dim allItems As List(Of ListViewItem) = Await CollectTab1FolderInfo(deDupedNodes, cToken)

            If _tab1SelectSeq <> mySeq Then Return  ' 序號機制配對：在 Await 回來後，若使用者已點擊其他節點，則放棄渲染
            RenderLv1(allItems)                     ' ── 執行渲染 (UI 呈現) ──

            ' 2026/07/06 by Simon/Claude: 這是啟動後第一次完整顯示 Tab1 (資料夾樹 + ListView1 統計數字都已呈現)，
            '   在此停下從 Form1_Load 開始跑的啟動計時，取代舊版「Me.Show() 後就停錶」的過早計時。
            If _startupStopwatch IsNot Nothing AndAlso _startupStopwatch.IsRunning Then
                _startupStopwatch.Stop()
                _startupElapsedMsg = "啟動花費 " & _startupStopwatch.Elapsed.TotalSeconds.ToString("0.00") & " 秒。"
            End If

            ' 2026/07/03 by Simon/Claude: 與 EnterSelectedFolder 統一行為 — 單選時自動算子資料夾大小。
            '   GetSubtree 骨架整合後 ComputeFolderSize 明顯變快，單層子資料夾即使亂點也能承受；
            '   多選(deDupedNodes.Count > 1) 則不觸發，避免多個大資料夾同時展開造成瀏覽卡頓。
            If deDupedNodes.Count = 1 Then ComputeFolderSize(Nothing, Nothing)

        Catch ex As OperationCanceledException
            _dbg("結束", "ESC 中斷") : PgrsBar1.Text = "由使用者中斷。" : Return
        Catch ex As System.Exception
            _dbg("錯誤", ex.Message)
        Finally
            OkeyNowByeByeToken(cToken)                      ' 2026/07/07 by Simon/Claude: 歸還 token — 運算中判定 token 化(見 OkayNowYouHaveToken/OkeyNowByeByeToken)
            Cursor = Cursors.Default : _isUserBusy = False  ' 2026/07/07: 移入 Finally，ESC 中斷路徑(上方 Return)也要復原
        End Try

        sw.Stop()
        PgrsBar1.Text = "統計花費 " & sw.Elapsed.TotalSeconds.ToString("0.00") & " 秒。"
        If Me.ActiveControl Is SimTree1 Then SimTree1.Focus()
        _dbg("結束")
    End Sub
    Private Async Sub Lv1_KeyDown(sender As Object, e As KeyEventArgs) Handles ListView1.KeyDown
        ''' <summary>
        ''' ListView1: 資料夾導覽 (2026/04/16 by Gemini 3.1 Pro: 從 HandleListViewKeyPress 拆分回歸)
        ''' </summary>
        ' 2026/07/07 by Simon/Claude: 原本在此處每次按鍵都 OkayNowYouHaveToken()，但只有「退上一層」分支真正用到 —
        '   token 化的運算中判定上線後，按鍵搶 token 會讓 _cts 殘留(沒人歸還)。改移入實際用到的分支內取用+歸還。
        Dim lv As ListView = DirectCast(sender, ListView)

        _dbg("開始", $"鍵值: {e.KeyCode}")
        If e.KeyCode = Keys.Enter Then
            If lv.SelectedItems.Count = 0 Then Return

            ' by Gemini 3 Flash, 2026/04/13: 選取多個項目時，改用 MessageBox 顯示數量加總
            If lv.SelectedItems.Count > 1 Then
                Dim sumDirect As Long = 0 : Dim sumTotal As Long = 0
                For Each item As ListViewItem In lv.SelectedItems
                    ' 2026/04/13 v2: 移除「所屬父資料夾」欄後，索引回歸原位
                    ' SubItems(1): 郵件數量(直屬)；SubItems(3): 郵件總計(含子孫)
                    Dim strDirect As String = item.SubItems(1).Text.Replace(",", "").Trim()
                    Dim strTotal As String = item.SubItems(3).Text.Replace(",", "").Trim()
                    Dim valDirect As Long = 0 : Dim valTotal As Long = 0
                    Long.TryParse(strDirect, valDirect) : Long.TryParse(strTotal, valTotal)
                    sumDirect += valDirect : sumTotal += valTotal
                Next
                MessageBox.Show($"已選取 {lv.SelectedItems.Count:N0} 個資料夾統計結果：" & vbCrLf & vbCrLf &
                                $"【本層郵件】加總：{sumDirect:N0} 封" & vbCrLf &
                                $"【包含子樹】加總：{sumTotal:N0} 封", "複選數量加總", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim selectedItem As ListViewItem = lv.SelectedItems(0)          ' 獲取點選的資料夾並進入 (單選時維持原邏輯)
            If selectedItem IsNot Nothing Then EnterSelectedFolder(selectedItem)
            e.Handled = True : e.SuppressKeyPress = True

        ElseIf e.KeyCode = Keys.Back Or e.KeyCode = Keys.Escape Then
            ' 2026/05/26 by Simon/Claude: ESC 先退回上一層資料夾，退無可退才把焦點移回左側樹
            '   原始意圖：Lv1 → ESC → 退回上層 → Lv1（游標落在剛才那個資料夾）
            '   原因：Gemini 3.1 Pro 2026/04/22 把退層邏輯整段 comment 掉，改成無條件 SimTree1.Focus()，行為不符預期。
            '   修正：有父節點 → 收攏當前節點、選取父節點、重算 Lv1、高亮剛才的資料夾、保持焦點在 Lv1

            Dim currentNode As TreeNode = SimTree1.SelectedNode     ' 2026/04/13 by Simon/Claude: Tab1 改用 SimTree1
            If currentNode IsNot Nothing AndAlso currentNode.Parent IsNot Nothing Then
                Dim cToken As CancellationToken = OkayNowYouHaveToken()   ' 2026/07/07 by Simon/Claude: 從函式開頭移入本分支(唯一用到 token 的地方)
                Try
                    ' 記下當前 Folder 物件，用於回到上層後在 ListView1 定位游標
                    Dim currentFolder As Folder = TryCast(currentNode.Tag, Folder)
                    Dim parentNode As TreeNode = currentNode.Parent

                    ' 用 SimTree1 正確選取父節點 (不呼叫 FireAfterSelect，避免與下方手動計算重複觸發)
                    SimTree1.ClearSelectedNodes()
                    SimTree1.AddSelectedNode(parentNode)

                    ' 手動計算統計並渲染 (等同 SimTree1_AfterSelect 的流程)
                    Dim dedupedNodes As List(Of TreeNode) = SimTree1.GetDedupedSelection()   ' 2026/07/10 by Simon/Claude: 改用 SimTree 內建版
                    Dim items As List(Of ListViewItem) = Await CollectTab1FolderInfo(dedupedNodes, cToken)
                    RenderLv1(items)

                    ' 在 ListView1 中找到代表「剛才那個資料夾」的列並移去高亮
                    ' by Gemini 3.5 Flash, 2026/05/27: 將此巢狀尋找高亮邏輯重構抽離至獨立的輔助子程序，簡化事件代碼並強化型別轉型保護
                    SelectFolderInListView(lv, currentFolder)
                    lv.Focus()
                Finally
                    OkeyNowByeByeToken(cToken)              ' 2026/07/07 by Simon/Claude: 歸還 token
                End Try
            ElseIf e.KeyCode = Keys.Escape Then             ' 2026/06/27 新增 by simon：Backspace 只會退到最頂層，要ESC才會退回左側資料夾樹
                SimTree1.Focus()                            ' 2026/05/30 新增：若已在最頂層，則退回左側資料夾樹
            End If
            e.Handled = True : e.SuppressKeyPress = True

        ElseIf e.KeyCode = Keys.F5 Then                     ' F5 強制刷新 ListView1：繞過快取直接走 L3 OOM (2026/05/13 by Claude Sonnet 4.6)
            Await ForceLv1Refresh()
        ElseIf e.Control AndAlso e.KeyCode = Keys.C Then    ' Ctrl-C 複製選取列到剪貼簿 (by Claude Sonnet 4.6, 2026/04/27)
            LviCopyToClipboard(lv, e)
        ElseIf e.Control AndAlso e.KeyCode = Keys.A Then    ' Ctrl-A 全選 listview1 所有項目
            LviSelectAll(lv, e)
        End If
        If _iLikeNoisy Then _dbg("結束")

    End Sub
    Private Sub Lv1_MouseClick(sender As Object, e As MouseEventArgs) Handles ListView1.MouseClick
        ' ✅ 直接顯示已初始化好的選單，不重複建立和 AddHandler
        If e.Button = MouseButtons.Right Then ctxMenuLv1.Show(System.Windows.Forms.Cursor.Position)
        ' 2026/3/6: 原有程式碼每次都會新建一個ContextMenuStrip, 每次都新建一個都要重新AddHandler會造成memory leak
        ' 現在改成只在initial的時候建立一次, 之後每次右鍵點擊的時候直接Show()就好, 不用再重複建立
    End Sub
    Private Sub Lv1_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles ListView1.MouseDoubleClick
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
    Private Async Function CollectTab1FolderInfo(dedupedNodes As List(Of TreeNode), cToken As CancellationToken, Optional skipCache As Boolean = False) As Task(Of List(Of ListViewItem))
        ''' <summary>
        ''' [Layer 1.5 業務邏輯層] 針對去重後的節點進行非同步統計運算。
        ''' 支援一般 BFS 運算與 F5 強制刷新模式。
        ''' skipCache=True   時委派給 CollectFolderInfoForceL3 (直打 L3)；
        ''' skipCache=False  走 CollectFolderInfoByBFS (BFS + 快取剪枝)。
        ''' </summary>
        ''' <param name="dedupedNodes">已去重過的選中節點清單</param>
        ''' <param name="cToken">取消標記</param>
        ''' <param name="skipCache">是否繞過快取強制重新讀取 L3 資料</param>
        ''' <returns>組裝完畢的 ListViewItem 清單</returns>

        ' 預分配容量為 128，優化 UI 項目渲染效能 (by Gemini 3 Flash, 2026/05/04)
        Dim allItems As New List(Of ListViewItem)(128)
        Dim subTotalMail As Long = 0 : Dim subTotalFolders As Integer = 0
        Dim multiMode As Boolean = dedupedNodes.Count > 1
        Dim mySeq As Integer = _tab1SelectSeq

        For Each node As TreeNode In dedupedNodes
            Dim folder As Folder = TryCast(node.Tag, Folder)
            If folder Is Nothing Then Continue For

            Dim rows As List(Of FolderBfsEntry)
            If skipCache = False Then
                ' === 標準模式：呼叫 Layer2 BFS 展開，含快取剪枝機制 ===
                rows = Await CollectFolderInfoByBFS(folder, New Progress(Of ProgressReport)(Sub(p) PgrsBar2.Text = p.Message), cToken:=cToken)
            Else
                ' === F5 強制刷新策略：直接呼叫 L3 接口，不走 BFS 隊列 ===
                rows = Await CollectFolderInfoForceL3(folder, cToken)
            End If

            ' 序列安全性檢查：若在 Await 期間使用者切換了選擇，立即終止運算
            If _tab1SelectSeq <> mySeq Then Return New List(Of ListViewItem)

            ' 將統計結果轉換為 UI 行 (rows(0)為群組標題, rows(1..)為子資料夾)
            If rows.Count > 0 Then
                allItems.Add(BuildLv1GroupHeader(rows(0), node))
                For i As Integer = 1 To rows.Count - 1
                    allItems.Add(BuildLv1Item(rows(i), node))
                Next
                subTotalMail += rows(0).TotalMailCount
                subTotalFolders += CInt(rows(0).TotalSubCount)
            End If
        Next

        ' 多選模式：在清單最底部追加合計列
        If multiMode AndAlso allItems.Count > 0 Then
            allItems.Add(BuildLv1SumRow(dedupedNodes.Count, subTotalFolders, subTotalMail))
            PgrsBar2.Text = $"統計完成: 共選取 {dedupedNodes.Count:N0} 個資料夾，合計 {subTotalMail:N0} 封郵件。"
        End If

        Return allItems
    End Function
    Private Async Function CollectFolderInfoByBFS(rootFolder As Folder, progress As IProgress(Of ProgressReport), cToken As CancellationToken) As Task(Of List(Of FolderBfsEntry))
        ' ==============================================================
        ' === Layer 2 (流程協調層) ===
        ' 職責: BFS 廣度優先搜索，展開整棵子樹，管理快取剪枝，驅動 Layer3，底部向上彙總，回傳顯示清單
        ' 
        ' 2026/04/04 by Gemini 重構紀錄:
        ' v5: 原有的百行巨型函數已被依「單一職責原則」拆分為五個子函數，確保各步驟隔離互不干擾。
        '
        ' 拆分後的五個步驟 (Steps):
        '   Step 1. BuildBfsFolderTree      : GetSubtree 骨架一次展開 + 記憶體剪枝(2026/07/02 骨架整合)；若快取命中(非root)則剪枝。
        '   Step 2. FetchDirectMailCounts   : 對未快取節點逐一呼叫 GetMailCount() 取本層郵件數。(處理 progress 報告並支援 _cancelRequested 中斷。)
        '   Step 3. SumUpSubTreeBottomUp    : 利用 BFS「父索引 < 子索引」特性，從陣列尾端往前掃一次完成加總。
        '   Step 4. UpdateFolderInfoCache   : 將最新結果寫入 Layer2.5 的 _cacheMailCountAll 等字典。
        '   Step 5. GetBfsResult            : 從陣列中挑出 root 與直屬子資料夾 (ParentIndex=0) 並補讀快取。
        '
        ' 架構與效能考量:
        '   - allEntries 是 Reference Type，在此作為狀態載體在各子函數間傳遞，避免不必要的陣列複製。
        '   - 為防 BFS 索引錯亂，以 IReadOnlyList 宣告參數，確保子函數不可改變 allEntries 長度或顛倒內部順序。
        '   - v4 bug fix: BFS 剪枝規則為「root 永遠展開直屬子資料夾，不論快取」。
        ' ==============================================================
        Dim rName As String = rootFolder?.Name
        _dbg(" ├ 開始", rName)

        ' ── Step 1: 負責展開樹狀結構與初步快取剪枝 (by Gemini, 2026/04/05 改為非同步以提升響應)
        Dim allEntries As List(Of FolderBfsEntry) = Await BuildBfsFolderTree(rootFolder, cToken:=cToken, progress:=progress)

        ' ── Step 2: 負責與 COM 溝通，取得基本數據
        Await FetchDirectMailCounts(allEntries, progress, cToken:=cToken)

        ' ── Step 3 & 4: 純記憶體運算與快取更新
        SumUpSubTreeBottomUp(allEntries)
        UpdateFolderInfoCache(allEntries)

        ' ── Step 5: 提取 UI 所需的結果並回報最終進度
        Dim res = GetBfsResult(allEntries, progress)
        _dbg(" ├ 結束", rName)
        Return res

    End Function
    Private Async Function CollectFolderInfoForceL3(folder As Folder, cToken As CancellationToken) As Task(Of List(Of FolderBfsEntry))
        ' 2026/05/27 by Simon/Claude: 從 CollectTab1FolderInfo skipCache 分支抽出
        ' F5 強制刷新策略：直打 L3，只計算 root + 直屬子資料夾，不走 BFS 整棵子樹
        Dim rootPath As String = SafeGetPath(folder)
        PgrsBar2.Text = $"F5: 讀取 {ExtractFolderName(rootPath)}..."

        ' 2026/06/13 by Simon/Claude Opus 4.8: skipCache:=True 一路穿透到 L2.5/L3，F5 跳過記憶體+DB 快取直打 L3 完整重掃
        ' 2026/06/23 by Simon/Claude: F5 改走 proxy skipCache(RDO 派工),仍繞過快取讀寫
        Dim rootMc As Long = GetMailCount(folder, rootPath, skipCache:=True)
        Dim rootFc As Long = GetFolderCount(folder, rootPath, skipCache:=True)
        Dim rootMca As Long = Await GetMailCountAllOOM(folder, skipCache:=True, cToken:=cToken)
        Dim rootFca As Long = Await GetFolderCountAllOOM(folder, skipCache:=True, cToken:=cToken)

        ' 更新快取
        _cacheMailCount(rootPath) = rootMc : _cacheMailCountAll(rootPath) = rootMca
        _cacheFolderCount(rootPath) = rootFc : _cacheFolderCountAll(rootPath) = rootFca

        Dim rows As New List(Of FolderBfsEntry) From {New FolderBfsEntry With {.Folder = folder, .FolderPath = rootPath,
                                                                               .DirectMailCount = rootMc, .TotalMailCount = rootMca, .TotalSubCount = rootFca}}

        ' ✅ 2026/5/31 by Gemini/Simon: 加入 skipCache 引數判斷是否要強制讀取COM
        ' 處理直屬子資料夾 (ps. 這裡是不是多餘重複了?? 上面不是已經用CountAll也都更新cache了, 為何還要再逐一讀一次直屬子資料夾?? 這裡的邏輯是什麼??)
        ' 解答：這並非多餘。by Gemini 3.5 Flash, 2026/06/27
        '   (1) 根資料夾呼叫的 GetMailCountAllOOM(root) 只會回傳整棵子樹的總郵件數，在計算過程中「完全不會」寫入或更新任何直屬子資料夾的個別快取字典。
        '   (2) F5 強制重刷的 UI 畫面需要同時呈現 Root 與其所有「直屬子資料夾」的獨立統計數據。
        '       為了取得各直屬子資料夾個別的 DirectMailCount 與 TotalMailCount/TotalSubCount 數據，必須逐一尋訪直屬子資料夾 (child)，
        '       對其單獨呼叫 GetMailCountAllOOM(child) 以取得其子樹總數，寫入其 childPath 對應的快取字典中，並將個別的 FolderBfsEntry 包入 rows 回傳。
        For Each sf In GetSortedSubFolders(folder, rootPath, skipCache:=True)  ' 2026/07/10: tuple 版, childPath 直接取 sf.fPath 免 COM 讀 Name
            cToken.ThrowIfCancellationRequested()
            Dim child As Folder = sf.f
            Dim childPath As String = sf.fPath
            ' 2026/06/13 by Simon/Claude Opus 4.8: 同上，子資料夾亦以 skipCache:=True 強制重掃
            ' 2026/06/23 by Simon/Claude: 同上,F5 子夾改 proxy skipCache
            _cacheMailCount(childPath) = GetMailCount(child, childPath, skipCache:=True)
            _cacheFolderCount(childPath) = GetFolderCount(child, childPath, skipCache:=True)
            _cacheMailCountAll(childPath) = Await GetMailCountAllOOM(child, skipCache:=True, cToken:=cToken)
            _cacheFolderCountAll(childPath) = Await GetFolderCountAllOOM(child, skipCache:=True, cToken:=cToken)

            rows.Add(New FolderBfsEntry With {.Folder = child, .FolderPath = childPath,
                                              .DirectMailCount = _cacheMailCount(childPath),
                                              .TotalMailCount = _cacheMailCountAll(childPath),
                                              .TotalSubCount = _cacheFolderCountAll(childPath)})
        Next
        Return rows
    End Function

    ' 以下為 CollectFolderInfoByBFS 專用的拆分子函數 (Steps 1~5)
    Private Async Function BuildBfsFolderTree(rootFolder As Folder, cToken As CancellationToken, Optional progress As IProgress(Of ProgressReport) = Nothing) As Task(Of List(Of FolderBfsEntry))
        ' 負責: 展開整棵子樹 + 依 Layer2.5 快取字典剪枝，產出「父索引 < 子索引」的 FolderBfsEntry 陣列。
        ' 2026/07/02 by Simon/Claude [骨架整合]: 不再自己逐節點走樹(原: 每節點 LazyGetOrderedSubFolderIDs 查一次 DB + 未知節點退 COM 物化)。
        '   改為一次呼叫 GetSubtree(L2.5) 取完整骨架: ①記憶體 _cacheSubTreeList → ②DB 一條 LIKE → ③RDO 批次(GetSubtreeRdo) → ④OOM BFS。
        '   Tab1 從此與 Tab2-5 共用同一份骨架快取；冷啟動由 RDO 批次/OOM 全樹掃接手，
        '   原 selfKnownToDb 冷啟動特判(2026/07/01)與 GetSortedSubFolderIDs 的 ③ COM 物化退路從此不再需要。
        '   保留的既有語意 (只是搬到記憶體樹上執行):
        '     - 剪枝規則: 非 root 節點 mca/fca 雙快取命中才剪枝；root 永遠展開直屬子夾(v4 bug fix)。
        '     - DB lazy: 未命中節點 LazyGetFolderInfo + FillCacheFromDbRow(skipAggregates:=True) 只填本層欄位，不驗 snapshot(原樣)。
        '     - 模式過濾: 骨架為完整全集，套 FilterSubtreeByMode 對齊原 LazyGetOrderedSubFolderIDs 的 is_mail 過濾。
        '     - 顯示排序: 比照 LazyGetOrderedSubFolderIDs 的 ORDER BY has_chinese ASC, folder_path ASC，在記憶體對每層子夾排序，
        '                 確保 GetBfsResult 取 ParentIndex=0 的顯示順序不變。
        If _iLikeNoisy Then _dbg("    ├ 開始", rootFolder.Name)
        Dim swP As Stopwatch = Stopwatch.StartNew()     ' PROBE_TIMING

        ' ── ① 一次取得完整骨架 (首元素為 root 自身，含非郵件夾) ──
        Dim rootPath As String = SafeGetPath(rootFolder)
        Dim skeleton As List(Of (eid As String, sid As String, fPath As String)) = Await GetSubtree(rootFolder, includeSubF:=True, progress:=progress, cToken:=cToken)
        Dim tSkel As Double = swP.Elapsed.TotalMilliseconds : swP.Restart()     ' PROBE_TIMING

        ' ── ② 模式過濾 + 建 parent → children 記憶體樹 (字典寫法比照 FilterSubtreeByMode) ──
        Dim filtered As List(Of (eid As String, sid As String, fPath As String)) = FilterSubtreeByMode(skeleton, rootPath)
        Dim byPath As New Dictionary(Of String, (eid As String, sid As String, fPath As String))(filtered.Count)
        For Each t In filtered : byPath(t.fPath) = t : Next
        Dim childrenOf As New Dictionary(Of String, List(Of (eid As String, sid As String, fPath As String)))(filtered.Count)
        For Each t In filtered
            Dim sepIdx As Integer = t.fPath.LastIndexOf("\"c)
            If sepIdx > 0 Then
                Dim parentPath As String = t.fPath.Substring(0, sepIdx)
                If byPath.ContainsKey(parentPath) Then
                    Dim lst As List(Of (eid As String, sid As String, fPath As String)) = Nothing
                    If Not childrenOf.TryGetValue(parentPath, lst) Then lst = New List(Of (eid As String, sid As String, fPath As String))() : childrenOf(parentPath) = lst
                    lst.Add(t)
                End If
            End If
        Next
        For Each kv In childrenOf   ' 英文優先排序，對齊 LazyGetOrderedSubFolderIDs 的 SQL ORDER BY (has_chinese ASC, folder_path ASC)
            kv.Value.Sort(Function(a, b)
                              Dim ha As Integer = If(TextHasChineseChar(ExtractFolderName(a.fPath)), 1, 0)
                              Dim hb As Integer = If(TextHasChineseChar(ExtractFolderName(b.fPath)), 1, 0)
                              If ha <> hb Then Return ha.CompareTo(hb)
                              Return String.CompareOrdinal(a.fPath, b.fPath)
                          End Function)
        Next

        ' ── ③ 記憶體樹上執行原有 BFS + 快取剪枝 (剪枝/DB lazy 邏輯原樣，資料來源從 DB/COM 換成 childrenOf 字典) ──
        Dim allEntries As New List(Of FolderBfsEntry)(Math.Max(filtered.Count, 16))
        ' 2026/06/29 by Simon/Claude [Option A1]: queue 持純 id tuple(eid/sid/path/parentIdx)，走樹零 COM 物化 (原樣保留)
        Dim queue As New Queue(Of (eid As String, sid As String, path As String, parentIdx As Integer))(512)
        Dim rootT As (eid As String, sid As String, fPath As String) = Nothing
        If Not byPath.TryGetValue(rootPath, rootT) Then     ' 理論上骨架必含 root；保底直接讀一次 COM 身分證
            Try : rootT = (rootFolder.EntryID, rootFolder.StoreID, rootPath) : Catch : rootT = ("", "", rootPath) : End Try
        End If
        queue.Enqueue((rootT.eid, rootT.sid, rootPath, -1))

        ' by Gemini, 2026/04/05: 每 100ms 主動讓出執行緒並檢查中斷，兼顧效能與靈敏度
        Dim swThrottle As Stopwatch = Stopwatch.StartNew()  ' by Claude Sonnet 4.6, 2026/06/07
        Try
            Do While queue.Count > 0
                Dim curr = queue.Dequeue()
                Dim fPath As String = curr.path
                ' 2026/06/29 by Simon/Claude [Option A1]: .Folder 走樹保持 Nothing，UI 階段(GetBfsResult)才物化；.Eid/.Sid 帶身分證
                Dim entry As New FolderBfsEntry With {.Folder = Nothing, .Eid = curr.eid, .Sid = curr.sid, .ParentIndex = curr.parentIdx, .IsFromCache = False, .FolderPath = fPath}
                Dim myIdx As Integer = allEntries.Count
                allEntries.Add(entry)

                ' 快取命中判斷: 兩個快取都有才算完整命中 (任一失效都重新計算，確保一致性)
                Dim cachedMail As Integer, cachedSub As Integer
                Dim isHit As Boolean = False
                If _cacheMailCountAll.TryGetValue(fPath, cachedMail) AndAlso _cacheFolderCountAll.TryGetValue(fPath, cachedSub) Then
                    isHit = True            ' ① 記憶體命中
                Else
                    ' ② DB lazy load：只填本層欄位（mc/fc/fs/身分標識），不以 mca/fca 做剪枝
                    ' by Claude Sonnet 4.6, 2026/04/25: 選項 A 修正 — DB 的 mca/fca/fsa 帶有模式語意，無法確認是在哪個 _showAllFolders 模式下計算並儲存的。
                    '   若直接使用 DB 值做剪枝，切換模式後或重啟後第一次統計會顯示舊模式的錯誤加總。
                    '   skipAggregates:=True → FillCacheFromDbRow 只填 mc/fc/fs 等本層無模式語意的欄位，
                    '   isHit 保持 False → BFS 繼續展開子資料夾，自行重算 mca/fca，重算結果透過 UpdateFolderInfoCache 寫入記憶體，下次同模式點選從記憶體命中（①）。
                    '   效能代價：每次切換或重啟後第一次統計需完整展開（不能 DB 剪枝），可接受。
                    '   (2026/07/02 骨架整合註記: 此處每節點一次 LazyGetFolderInfo 點查詢原本被刻意保留，目的是預熱 _cacheMailCount，
                    '    讓 Step2 的 GetMailCount 走 ① 記憶體命中而不觸發「DB lazy + snapshot 驗證」的逐夾 COM 讀取。
                    '   2026/07/03 by Simon/Claude [PROBE_HIERCNT 通過後]: mc/fc 現在已由 Step①的 GetSubtree(RDO 批次)免費回填,
                    '    只剩 fs 仍無免費來源。三者都已在記憶體時這次點查詢已無新資訊可拿，跳過(FillCacheFromDbRow 是 TryAdd 語意本就不覆寫，
                    '    此處只是省掉白做工的 SQL 往返，行為不變)。任一者仍缺才照原樣查 DB。
                    If Not _cacheMailCount.ContainsKey(fPath) OrElse Not _cacheFolderCount.ContainsKey(fPath) OrElse Not _cacheFolderSize.ContainsKey(fPath) Then
                        Dim row = LazyGetFolderInfo(fPath)
                        If row IsNot Nothing Then FillCacheFromDbRow(fPath, row, skipAggregates:=True)   ' 只填本層欄位，不填 mca/fca/fsa
                    End If
                End If

                If isHit Then
                    entry.TotalMailCount = cachedMail
                    entry.TotalSubCount = cachedSub
                    entry.IsFromCache = True

                    ' ★ v4 bug fix: root (parentIdx=-1) 即使快取命中，也要繼續展開直屬子資料夾，只有非 root 節點才允許剪枝
                    If curr.parentIdx <> -1 Then Continue Do
                End If

                ' 未命中，或是 root (不論有無快取) → 展開直屬子資料夾 (純記憶體查表，零 DB/COM)
                Dim kids As List(Of (eid As String, sid As String, fPath As String)) = Nothing
                If childrenOf.TryGetValue(fPath, kids) Then
                    For Each k In kids : queue.Enqueue((k.eid, k.sid, k.fPath, myIdx)) : Next
                End If
                Await SmartThrottle(swThrottle, cToken:=cToken, ThrottleFreq.Hii)  ' 2026/04/16 by Simon/Claude: 改用 SmartThrottle，省去 If/Restart/Task.Delay 三行套路
            Loop
        Catch ex As OperationCanceledException
            ' 2026/04/11 by Claude: 改為 re-throw，確保不完整的 BFS 樹不會繼續傳入
            ' UpdateFolderInfoCache，避免錯誤的中途統計結果汙染快取。
            ' (原本 catch 後繼續 Return allEntries，導致上層 CollectFolderInfoByBFS 看不到中斷，仍執行 SumUpSubTreeBottomUp + UpdateFolderInfoCache)
            _dbg("    ├ 中斷", $"BuildBfsFolderTree 已由使用者中斷")
            Throw
        End Try

        ' PROBE_TIMING: 骨架取得 vs 記憶體剪枝分段計時，供與舊版 S1 (整段自走樹) 的歷史 log 對比
        _dbg("    ├ 結束", $"骨架={skeleton.Count} 過濾後={filtered.Count} 節點={allEntries.Count}(含剪枝) | 骨架 {tSkel:F0}ms + 剪枝 {swP.Elapsed.TotalMilliseconds:F0}ms")
        Return allEntries

    End Function
    Private Async Function FetchDirectMailCounts(allEntries As IReadOnlyList(Of FolderBfsEntry), progress As IProgress(Of ProgressReport), cToken As CancellationToken) As Task
        ' 負責: 對未快取節點打 COM (呼叫 GetMailCount)，並負責 UI 節流 (Task.Yield) 與 ESC 中斷檢查。
        ' 2026/04/11 by Claude: 回傳值從 Task(Of Boolean) 改為 Task，原本的 Return True/False 均改為 re-throw。
        '   理由: 呼叫端 Await FetchDirectMailCounts(...) 完全丟棄了 Task(Of Boolean) 的回傳值，
        '         等同 Return True 無效，ESC 後上層照樣執行 UpdateFolderInfoCache 污染快取。
        '         改為 Throw 後，OperationCanceledException 直接傳到 TreeView1_AfterSelect 的 catch 攔截。

        If _iLikeNoisy Then _dbg("    ├ 開始") ' by Gemini, 2026/04/10: 調整縮排層級為 Level 2
        Dim total As Integer = allEntries.Count, processed As Integer = 0
        Dim swThrottle As Stopwatch = Stopwatch.StartNew()  ' by Claude Sonnet 4.6, 2026/06/07

        Try
            For fd As Integer = 0 To total - 1
                Dim entry As FolderBfsEntry = allEntries(fd)
                If Not entry.IsFromCache Then
                    ' 2026/06/29 by Simon/Claude [Option A1]: 改走 folder-free GetMailCount(fPath,eid,sid)，走樹階段 .Folder 仍 Nothing
                    entry.DirectMailCount = GetMailCount(entry.FolderPath, entry.Eid, entry.Sid)
                    entry.TotalMailCount = entry.DirectMailCount             ' 初始值 = 本層，後面底部向上累加子孫
                    entry.TotalSubCount = 0                                  ' 初始為 0，後面累加子孫資料夾數
                End If
                processed += 1

                ' 2026/04/16 by Simon/Claude: 改用 ThrottleFreq.Hii + SmartThrottle 與 onThrottled 委派
                Await SmartThrottle(swThrottle, cToken:=cToken, ThrottleFreq.Hii,
                                          Sub() progress?.Report(New ProgressReport With {.CurrentCount = processed, .TotalCount = total,
                                                                                          .Message = $"正在統計郵件數: {processed:N0} / {total:N0} 個資料夾..."}))
            Next
        Catch ex As OperationCanceledException
            ' 2026/04/11 by Claude: 改為 re-throw，不再 Return True (上層丟棄了回傳值，Return True 等於無效) 
            _dbg("    ├ 中斷", "FetchDirectMailCountsAsync 已由使用者中斷")
            Throw
        End Try
        _dbg("    ├ 結束", $"共讀取 {total:N0} 個節點 (非快取) ") ' by Gemini, 2026/04/04: Issue 2 補上結束

    End Function
    Private Sub SumUpSubTreeBottomUp(allEntries As IReadOnlyList(Of FolderBfsEntry))
        ' 負責: 底部向上彙總。利用 BFS 父節點索引必小於子節點的特性，反向遍歷即可。
        ' 備註: 純記憶體 O(N) 運算且無 COM 呼叫，無 StackOverflow 風險。

        If _iLikeNoisy Then _dbg("    ├ 開始") ' by Gemini, 2026/04/10: 調整縮排層級為 Level 2
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
        _dbg("    ├ 結束", $"共讀取 {allEntries.Count:N0} 個節點 (非快取) ") ' by Gemini, 2026/04/10: 調整縮排層級為 Level 2

    End Sub
    Private Sub UpdateFolderInfoCache(allEntries As IReadOnlyList(Of FolderBfsEntry))
        ' 負責: 將新計算的彙總結果寫入 Layer2.5 快取 
        ' 備註: TryAdd 不覆蓋既有值，避免污染快取
        If _iLikeNoisy Then _dbg("    ├ 開始") ' by Gemini, 2026/04/10: 調整縮排層級為 Level 2
        For Each entry As FolderBfsEntry In allEntries
            If Not entry.IsFromCache Then
                Dim fPath As String = entry.FolderPath
                _cacheMailCountAll.TryAdd(fPath, entry.TotalMailCount)
                _cacheFolderCountAll.TryAdd(fPath, entry.TotalSubCount)
            End If
        Next
        _dbg("    ├ 結束", $"共更新 {allEntries.Count:N0} 個節點的快取") ' by Gemini, 2026/04/10: 調整縮排層級為 Level 2

    End Sub
    Private Function GetBfsResult(allEntries As IReadOnlyList(Of FolderBfsEntry), progress As IProgress(Of ProgressReport)) As List(Of FolderBfsEntry)
        ' 負責: 找出 root 與直屬子資料夾 (ParentIndex=0) 組裝成新的 UI 呈現清單，並在最後發布結束進度。
        If _iLikeNoisy Then _dbg("    ├ 開始")

        ' 預分配容量為 512，對應 BFS 展開後的資料夾清單 (by Gemini 3 Flash, 2026/05/04)
        Dim result As New List(Of FolderBfsEntry)(512)
        result.Add(allEntries(0))   ' index 0 = rootFolder 本身

        ' 2026/06/29 by Simon/Claude [Option A1]: UI 物化交界 — 只對 root + 直屬子夾物化 COM Folder(供 .Name/IsMailFolder/Tag)，
        '   走樹的全樹其餘節點維持零物化。root 物化僅 1 次 COM(可忽略)。
        ' 2026/07/04 by Simon/Claude Fable 5 [PROBE_TAB1E2E 拆帳後]: 徹底去物化 — 拆帳實測 26 store 全選第二輪:
        '   GetFolderById 物化 178 次 ≈205ms + folder 版 GetMailCount(內含 SafeGetPath COM 讀 .FolderPath) ≈37ms,佔 0.55 秒的大宗。
        '   root+直屬子夾也不再物化(.Folder 維持 Nothing);本層郵件數改走免-folder GetMailCount(fPath,eid,sid)(①記憶體命中 0 COM);
        '   UI Tag 改帶 id-tuple(見 BuildLv1Item),Enter/ComputeSize 消費端用到那一刻才物化單一資料夾。

        For i As Integer = 1 To allEntries.Count - 1
            Dim entry As FolderBfsEntry = allEntries(i)
            If entry.ParentIndex = 0 Then
                ' 若直屬子資料夾快取命中，補讀一下其本層郵件 (DirectMailCount) — 免-folder 版,0 COM
                If entry.IsFromCache Then entry.DirectMailCount = GetMailCount(entry.FolderPath, entry.Eid, entry.Sid)
                result.Add(entry)
            End If
        Next

        ' 若 root 自身快取命中，也補讀其本層郵件數 — 免-folder 版
        If allEntries(0).IsFromCache Then allEntries(0).DirectMailCount = GetMailCount(allEntries(0).FolderPath, allEntries(0).Eid, allEntries(0).Sid)

        ' 2026/5/6 by Claude Sonnet 4.6, GetBfsResult 內，建立 result 清單後，預載 fsa
        For Each e In allEntries
            If Not _cacheFolderSizeAll.ContainsKey(e.FolderPath) Then
                Dim row = LazyGetFolderInfo(e.FolderPath)
                If row IsNot Nothing AndAlso row.fsa >= 0 Then _cacheFolderSizeAll.TryAdd(e.FolderPath, row.fsa)
            End If
        Next

        Dim totalMail As Long = allEntries(0).TotalMailCount
        Dim totalFolder As Integer = allEntries(0).TotalSubCount
        progress?.Report(New ProgressReport With {.CurrentCount = allEntries.Count, .TotalCount = allEntries.Count,
                                                  .Message = $"統計完成: 共 {totalFolder:N0} 個子資料夾，{totalMail:N0} 封郵件。"})

        _dbg("    ├ 結束", $"回傳 {result.Count:N0} 列 (1 root + {result.Count - 1:N0} 直屬子資料夾)") ' by Gemini, 2026/04/10
        Return result

    End Function
#End Region
#Region "  ├ UI 渲染"
    Private Sub RenderLv1(items As List(Of ListViewItem))
        ''' <summary>
        ''' [UI 渲染層] 負責將計算好的項目更新至 ListView1。
        ''' 包含雙緩衝優化 (BeginUpdate) 與清理過期的滑鼠懸停狀態。
        ''' </summary>
        ''' <param name="items">要顯示的 ListViewItem 清單</param>
        _dbg("開始")
        ListView1.BeginUpdate()         ' BeginUpdate 可防止大規模更新時的畫面閃爍
        ListView1.Items.Clear()
        _lastHoveredLvItem = Nothing    ' 2026/04/14 fix: 重建清單前清掉 stale 參照，避免第一次 hover 閃動

        If items IsNot Nothing AndAlso items.Count > 0 Then ListView1.Items.AddRange(items.ToArray())
        ListView1.EndUpdate()

        ' 2026/07/04 by Simon/Claude: 列數變化會讓捲軸出現/消失，進而改變 ClientSize.Width，
        ' 但控制項本身 Width 不變 → 不會觸發 Resize 事件 → 欄寬不會自動跟著重算。
        ' CalculateLvColumnSize 內部已有「ClientSize.Width 未變就直接返回」的防線，這裡呼叫不會造成額外負擔。
        CalculateLvColumnSize(ListView1)
        _dbg("結束")
    End Sub
    Private Function BuildLv1GroupHeader(rootEntry As FolderBfsEntry, parentNode As TreeNode) As ListViewItem
        ' 群組標題行：取代舊版的 isRoot=True 第一列
        ' 顯示選中資料夾本身的完整統計 (TotalMailCount / TotalSubCount) 
        ' 欄位: ▸ 資料夾名稱 / 郵件數量 / 資料夾數量 / 郵件總計 / 大小 (5欄回归) 
        ' Tag = Nothing：EnterSelectedFolder 與 ComputeFolderSize 看到 Nothing 直接跳過
        ' 2026/04/13 by Simon/Claude: B方案 — 統一格式，單選與多選皆顯示群組標題行
        ' 2026/04/13 v2: 移除「所屬父資料夾」欄 (該欄內容永遠等於標題行本身，元余) 
        ' 2026/05/27 by Simon/Claude: 改用 FormatFolderSizeStr 取代重複邏輯，fPath 直接用 .FolderPath
        _dbg("    ├ 開始")

        Dim sizeStr As String = FormatFolderSizeStr(rootEntry.FolderPath)
        Dim directMailStr As String = rootEntry.DirectMailCount.ToString("N0") & " "
        Dim totalSubStr As String = rootEntry.TotalSubCount.ToString("N0") & " "
        Dim totalMailStr As String = rootEntry.TotalMailCount.ToString("N0") & " "

        ' 欄位順序: 名稱 / 郵件數量 / 資料夾數量 / 郵件總計 / 大小
        ' 2026/07/04 by Simon/Claude Fable 5 [去物化配套]: 名稱改取自 FolderPath 尾段(純字串),不再讀 .Folder.Name(COM;且 BFS 路徑 .Folder 現為 Nothing)
        Dim lvi As New ListViewItem({"▸ " & ExtractFolderName(rootEntry.FolderPath), directMailStr, totalSubStr, totalMailStr, sizeStr})
        lvi.Font = New Font(ListView1.Font, _fontBold)
        lvi.BackColor = SystemColors.GradientInactiveCaption
        lvi.Tag = Nothing

        Return lvi
        _dbg("    ├ 結束") ' by Gemini, 2026/04/10

    End Function
    Private Function BuildLv1Item(entry As FolderBfsEntry, parentNode As TreeNode) As ListViewItem
        ' 組裝 ListView1 的單一資料列 (直屬子資料夾) 
        ' 欄位: 資料夾名稱 / 郵件數量 / 資料夾數量 / 郵件總計 / 大小(Lazy)
        ' 2026/04/13 by Simon/Claude: B方案升級
        '   - 移除 isRoot 參數 (群組標題行已獨立為 BuildLv1GroupHeader) 
        '   - 移除 parentFolderName 參數 (該欄元余，已移除) 
        '   - parentNode 存入 Tag ValueTuple，供 EnterSelectedFolder 使用
        '   - Tag 改為 ValueTuple(SubFolder, ParentNode)，ComputeFolderSize 同步更新
        ' 舊版註記: by Gemini, 2026/03/31: 視覺優化重構
        '   1. 資料夾名稱：还原開頭縮排空白 (" - ")，保持整齊。
        '   2. 防止切邊：斜體時字串結尾補一格空白。
        ' 2026/05/27 by Simon/Claude: 改用 FormatFolderSizeStr 取代重複邏輯，fPath 直接用 .FolderPath
        ' 2026/07/04 by Simon/Claude Fable 5 [PROBE_TAB1E2E 拆帳後,去物化配套]:
        '   1. isMail 改查 _cacheFolderIDs(骨架展開時已回填;原 IsMailFolder(folder) 無 fPath → SafeGetPath 每列 1 次 COM)。
        '   2. 查無時: BFS 路徑依「查無預設視為 mail」慣例(不斜體);F5 路徑 .Folder 在,照舊走 IsMailFolder(帶 fPath 免 COM)。
        '   3. 名稱改取自 FolderPath 尾段(原 .Folder.Name 每列 1 次 COM)。兩者合計原本 ≈112ms/輪(26 store 全選)。
        Dim isItalicFolder As Boolean = False
        Dim fInfo As (eid As String, sid As String, isMail As Boolean, hasCh As Boolean) = Nothing
        If _cacheFolderIDs.TryGetValue(entry.FolderPath, fInfo) Then
            isItalicFolder = Not fInfo.isMail
        ElseIf entry.Folder IsNot Nothing Then
            isItalicFolder = Not IsMailFolder(entry.Folder, fPath:=entry.FolderPath)
        End If
        Dim displayName As String = " - " & ExtractFolderName(entry.FolderPath)
        If isItalicFolder Then displayName &= " "

        ' 大小: Lazy，從快取讀；未計算過則留空，等 ColumnClick 或右鍵選單觸發計算
        Dim sizeStr As String = FormatFolderSizeStr(entry.FolderPath)
        Dim directMailStr As String = entry.DirectMailCount.ToString("N0") & " "
        Dim totalSubStr As String = entry.TotalSubCount.ToString("N0") & " "
        Dim totalMailStr As String = entry.TotalMailCount.ToString("N0") & " "

        ' 統計數字字串化 (字串結尾一律補一格空白，確保斜體與正常字體對齊且不切邊, by Gemini, 2026/03/31)
        If sizeStr <> "- " Then sizeStr &= " "

        ' 欄位順序: 名稱 / 郵件數量 / 資料夾數量 / 郵件總計 / 大小
        Dim lvi As New ListViewItem({displayName, directMailStr, totalSubStr, totalMailStr, sizeStr})
        If isItalicFolder Then
            ' by Gemini, 2026/03/29: 特殊顯示非郵件資料夾 (斜體 + 灰色)
            lvi.ForeColor = Color.DarkGray : lvi.Font = New Font(ListView1.Font, _fontItalic)
        End If

        ' 2026/04/13 by Simon/Claude: Tag 改為 ValueTuple，ComputeFolderSize 與 EnterSelectedFolder 同步更新
        ' 2026/07/04 by Simon/Claude Fable 5 [去物化配套]: Tag 擴充為 5 欄 id-tuple —
        '   BFS 路徑 SubFolder=Nothing(帶 Eid/Sid/FPath 身分證);F5 路徑 SubFolder 照舊有值(Eid/Sid 可能為空)。
        '   消費端(Enter/ComputeSize/F5/SelectFolderInListView)以 ResolveLv1TagFolder 統一「用到那一刻才物化」。
        lvi.Tag = (SubFolder:=entry.Folder, Eid:=entry.Eid, Sid:=entry.Sid, FPath:=entry.FolderPath, ParentNode:=parentNode)
        Return lvi

    End Function
    Private Function ResolveLv1TagFolder(t As (SubFolder As Folder, Eid As String, Sid As String, FPath As String, ParentNode As TreeNode)) As Folder
        ' 2026/07/04 by Simon/Claude Fable 5 [去物化配套]: Lv1 Tag → Folder 的統一物化交界。
        '   F5 路徑 SubFolder 已在 → 直接用;BFS 路徑 SubFolder=Nothing → 以身分證 GetFolderById 物化單一資料夾(1 次 COM,只付給真的被點到的那列)。
        If t.SubFolder IsNot Nothing Then Return t.SubFolder
        If String.IsNullOrEmpty(t.Eid) Then Return Nothing
        Return GetFolderById(t.Eid, t.Sid)
    End Function
    Private Function BuildLv1SumRow(selectedCount As Integer, totalSub As Integer, totalMail As Long) As ListViewItem
        ' 合計列：多選模式才插入，顯示跨資料夾加總。Tag = Nothing
        ' Tag = Nothing：同群組標題行，不可進入
        ' 2026/04/13 by Simon/Claude: B方案新增，5欄格式
        Dim totalMailStr As String = totalMail.ToString("N0") & " "
        Dim totalSubStr As String = totalSub.ToString("N0") & " "
        Dim lvi As New ListViewItem({"▶ 合計 (" & selectedCount.ToString("N0") & " 個 PST 檔案) ", "", totalSubStr, totalMailStr, ""})
        lvi.Font = New Font(ListView1.Font, _fontBold)
        lvi.BackColor = Color.FromArgb(220, 235, 252)
        lvi.ForeColor = Color.FromArgb(0, 70, 140)
        lvi.Tag = Nothing
        Return lvi

    End Function
    Private Async Function ForceLv1Refresh() As Task
        ' ── F5 強制刷新 ListView1 ──────────────────────────────────────────────
        ' 職責: 完全繞過記憶體快取與 DB，直接呼叫 GetMailCountAllOOM / GetFolderCountAllOOM 取得真實值
        '       讀完後同時寫入記憶體快取（_cacheXXX）並更新 DB。Size 重算僅針對目前 column 4 != "- " 的 ListViewItem。
        ' 效能原理：有 RDO → GetMailCountAllOOM 內部呼叫 _rdo.TotalItemCount（單次 MAPI 屬性讀取）
        '           整體複雜度 O(M)，M = 直屬子資料夾數；相較 BFS O(N) 大幅節省
        '           無 RDO → 內部 BFS，與原架構同等 O(N)
        ' 2026/05/13 by Claude Sonnet 4.6
        ' ─────────────────────────────────────────────────────────────────────
        _dbg("開始")
        Dim sw As Stopwatch = Stopwatch.StartNew()   ' 2026/07/01 by Claude: F5 計時
        Dim selectedNodes As List(Of TreeNode) = SimTree1.SelectedNodes
        If selectedNodes.Count = 0 Then Return

        _isUserBusy = True : Cursor = Cursors.WaitCursor
        PgrsBar1.Text = "F5 強制更新中..." : PgrsBar2.Text = ""

        ' ① 去重
        Dim dedupedNodes As List(Of TreeNode) = SimTree1.GetDedupedSelection()   ' 2026/07/10 by Simon/Claude: 改用 SimTree 內建版
        Dim cToken As CancellationToken = OkayNowYouHaveToken()

        ' 收集目前 ListView1 有 size 顯示的 fPath (重算用)
        Dim sizeItemPaths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        For Each lvi As ListViewItem In ListView1.Items
            If lvi.Tag IsNot Nothing AndAlso lvi.SubItems.Count > 4 AndAlso lvi.SubItems(4).Text <> "- " Then
                ' 2026/07/04 by Simon/Claude Fable 5 [去物化配套]: Tag 改 5 欄 id-tuple,fPath 直接取 Tag.FPath(0 COM),不再 SafeGetPath(t.SubFolder)
                Dim t = DirectCast(lvi.Tag, (SubFolder As Folder, Eid As String, Sid As String, FPath As String, ParentNode As TreeNode))
                Dim fp As String = If(String.IsNullOrEmpty(t.FPath), SafeGetPath(t.SubFolder), t.FPath)
                If Not String.IsNullOrEmpty(fp) Then sizeItemPaths.Add(fp)
            End If
        Next

        Try
            ' ② 清除快取 (將意圖委託給 Layer 2.5 專責函數, 2026/5/31)
            For Each node In dedupedNodes
                Dim targetPath As String = SafeGetPath(TryCast(node.Tag, Folder))
                If Not String.IsNullOrEmpty(targetPath) Then InvalidateFolderTreeCache(targetPath) ' ✅ 一行解決，隱藏實作細節 (by Gemini/Simon, 2026/5/31)
            Next

            ' ③ 核心統計 (skipCache:=True)
            Dim items As List(Of ListViewItem) = Await CollectTab1FolderInfo(dedupedNodes, cToken, skipCache:=True)
            RenderLv1(items)

            ' ④ Size 重算
            If sizeItemPaths.Count > 0 Then
                PgrsBar2.Text = "正在重算資料夾大小..."
                For Each lvi As ListViewItem In ListView1.Items
                    If lvi.Tag Is Nothing Then Continue For

                    ' 2026/07/04 by Simon/Claude Fable 5 [去物化配套]: fPath 取 Tag.FPath;F5 強制重算必打 COM,此處才 ResolveLv1TagFolder 物化
                    Dim t = DirectCast(lvi.Tag, (SubFolder As Folder, Eid As String, Sid As String, FPath As String, ParentNode As TreeNode))
                    Dim fp As String = If(String.IsNullOrEmpty(t.FPath), SafeGetPath(t.SubFolder), t.FPath)
                    If Not sizeItemPaths.Contains(fp) Then Continue For
                    Dim tagFolder As Folder = ResolveLv1TagFolder(t)
                    If tagFolder Is Nothing Then Continue For

                    lvi.SubItems(4).Text = "計算中..."
                    Dim dl As Long : _cacheFolderSize.TryRemove(fp, dl) : _cacheFolderSizeAll.TryRemove(fp, dl)
                    Dim sz As Long = Await GetFolderSizeAll(tagFolder, fp, skipCache:=True, cToken:=cToken)       ' 2026/06/27 by Simon/Claude: 改走 skipCache 跳過 ①記憶體+②DB,保證強制重算(原 TryRemove 只清記憶體擋不住 DB lazy)
                    If sz < 0 Then : lvi.SubItems(4).Text = "計算失敗"
                    Else : lvi.SubItems(4).Text = FormatFolderSizeStr(fp) ' 2026/07/04 by Simon/Claude: 改用 FormatFolderSizeStr 統一格式(0 才會顯示 "- ",避免與 BuildLv1Item 顯示不一致)
                    End If
                    If sz >= 0 Then _cacheFolderSizeAll(fp) = sz
                Next
            End If

            ' ⑤ 持久化
            Await SaveCachesToDB()
            PgrsBar1.Text = "F5 強制更新完成，花費 " & sw.Elapsed.TotalSeconds.ToString("0.00") & " 秒。"

        Catch ex As OperationCanceledException
            PgrsBar1.Text = "由使用者中斷。"
        Catch ex As System.Exception
            _dbg("錯誤", ex.Message)
        Finally
            OkeyNowByeByeToken(cToken)   ' 2026/07/07 by Simon/Claude: 歸還 token — 運算中判定 token 化
            Cursor = Cursors.Default : _isUserBusy = False : _dbg("結束")
        End Try
    End Function

    ' ── ListView1 OwnerDraw handlers (2026/04/13 by Simon/Claude) ──────────────────
    ' 問題根因: Windows 在 ListView 的 hover/select 狀態下會覆蓋自訂 BackColor，導致群組標題行的淡藍底在滑鼠移上去或點擊後消失。
    ' 解法: ListView.OwnerDraw = True，只對 Tag=Nothing 的行自訂繪製，其餘一律 DrawDefault=True。該方式不影響一般資料列的外觀和排序等功能。
    Private Sub Lv1_DrawColumnHeader(sender As Object, e As DrawListViewColumnHeaderEventArgs) Handles ListView1.DrawColumnHeader
        e.DrawDefault = True   ' 欄位標頭用預設繪製
    End Sub
    Private Sub Lv1_DrawItem(sender As Object, e As DrawListViewItemEventArgs) Handles ListView1.DrawItem
        ' Tag=Nothing (群組標題行 / 合計列) ：自己填背景色，阻止 OS 在此階段畫 hover/select 高亮 → DrawSubItem 再負責各欄文字
        ' 一般列：DrawDefault=True，OS 正常繪製 (選取藍、hover 灰等維持原樣) 
        '
        ' 2026/04/14 fix by Gemini 3.1 Pro: 原本 DrawDefault=True，OS 仍會在 DrawItem 階段塗高亮背景，
        '   之後 DrawSubItem 雖覆蓋，但 hover 重繪有時只觸發 DrawItem 不重觸發 DrawSubItem，導致背景色被清掉。
        '   之前嘗試在此處填色，但這會導致沒有觸發 DrawSubItem 時，被背景色覆蓋掉文字而產生閃爍。
        '   最好的解法是什麼都不做 (且不設 DrawDefault=True)，讓畫面保留原本 DrawSubItem 畫好的狀態。
        If e.Item.Tag Is Nothing Then
            ' 不做任何事，保持原有的畫面像素
        ElseIf e.Item Is _lastHoveredLvItem AndAlso Not e.Item.Selected Then
            ' 2026/04/14: 自己處理 Hover，不要將 DrawDefault 設為 True，讓它進入 DrawSubItem 畫淡灰底
        Else
            e.DrawDefault = True
        End If
    End Sub
    Private Sub Lv1_DrawSubItem(sender As Object, e As DrawListViewSubItemEventArgs) Handles ListView1.DrawSubItem
        ' 2026/05/09 by Gemini 3 Flash: Resize 期間暫停繪製
        ' If _isResizingLv Then Return

        ' ==========================================
        ' 1. 狀態判定：決定是否接管繪製與使用什麼背景色
        '   Tag = Nothing (群組標題行 / 合計列)：自訂繪製，防止 OS hover/select 顏色覆蓋我們設定的 BackColor
        ' ==========================================
        Dim needCustomDraw As Boolean = False
        Dim bgColor As Color
        If e.Item.Tag Is Nothing Then
            needCustomDraw = True : bgColor = e.Item.BackColor          ' 群組標題行 / 合計列
        ElseIf e.Item Is _lastHoveredLvItem AndAlso Not e.Item.Selected Then
            needCustomDraw = True : bgColor = ThemeColors.MercuryGray   ' Hover 項目且未選取
        End If

        ' ==========================================
        ' 2. 執行繪製
        ' ==========================================
        ' 將不變的排版設定提取為 Const，避免每次宣告重複計算
        ' by Claude Sonnet 4.6, 2026/05/22: 加入 NoPrefix，防止資料夾名稱含 & 時在 hover 狀態下被當作快捷鍵前綴吃掉而消失
        Const baseFlags As TextFormatFlags = TextFormatFlags.VerticalCenter Or TextFormatFlags.SingleLine Or
                                             TextFormatFlags.EndEllipsis Or TextFormatFlags.NoPrefix
        If needCustomDraw Then
            ' --- A. 繪製背景 ---
            ' 2026/04/14 by Gemini 3.1 Pro: 為了避免修改 BackColor 觸發版面重算效能異常，我們手動為 Hover 項目自訂繪製底色
            Using bgBrush As New SolidBrush(bgColor) : e.Graphics.FillRectangle(bgBrush, e.Bounds) : End Using

            ' --- B. 決定文字位置與對齊參數 ---
            Dim textRect As Rectangle = e.Bounds
            Dim alignFlags As TextFormatFlags
            If e.ColumnIndex = 0 Then
                textRect.X += 2 : textRect.Width -= 2
                alignFlags = TextFormatFlags.Left       ' 第一欄：微調 Padding 消除系統預設位移感，並靠左
            Else
                alignFlags = TextFormatFlags.Right      ' 其餘欄位：靠右 (trailing space 本身就是視覺間距，不額外 Inflate)
            End If

            ' --- C. 繪製文字 ---
            ' 2026/04/14 fix by Gemini 3.1 Pro: 捨棄 GDI+ (e.Graphics.DrawString) 造成的測量位移與空白吃斷，
            ' 全面回歸使用與原生系統 (DrawDefault) 一致的 Win32 GDI 引擎 (TextRenderer.DrawText)。
            TextRenderer.DrawText(e.Graphics, e.SubItem.Text, e.Item.Font, textRect, e.Item.ForeColor, baseFlags Or alignFlags)
        Else
            ' 其餘一般列：交由作業系統原生處理
            e.DrawDefault = True
        End If
    End Sub
    Private Sub Lv1_ItemSelectionChanged(sender As Object, e As ListViewItemSelectionChangedEventArgs) Handles ListView1.ItemSelectionChanged
        ' 群組標題行 / 合計列不可被選取，選中即強制取消。
        ' 進一步防止 OS 在取消選取時再覆蓋一次我們的 BackColor。
        If e.Item.Tag Is Nothing AndAlso e.IsSelected Then e.Item.Selected = False
    End Sub
#End Region
#Region "  └ 輔助函數"

    Private Sub SelectFolderInListView(lv As ListView, targetFolder As Folder)
        ''' <summary>
        ''' 在指定的 ListView 中尋找並選取對應的 Folder 項目
        ''' </summary>
        ''' <param name="lv">目標 ListView</param>
        ''' <param name="targetFolder">要尋找並選取的目標 Folder 物件</param>
        ''' <remarks>by Gemini 3.5 Flash, 2026/05/27: 將原 ESC 鍵退回上層時高亮原資料夾的巢狀尋找邏輯重構抽離為獨立子程序，以提供更好的型別轉換安全性</remarks>
        If lv Is Nothing OrElse targetFolder Is Nothing Then Return

        Dim targetEid As String = targetFolder.EntryID
        For Each item As ListViewItem In lv.Items
            If item.Tag IsNot Nothing Then
                Try
                    ' 2026/07/04 by Simon/Claude Fable 5 [去物化配套]: 優先用 Tag.Eid 比對(0 COM);F5 路徑 Eid 空才退 SubFolder.EntryID
                    Dim t = DirectCast(item.Tag, (SubFolder As Folder, Eid As String, Sid As String, FPath As String, ParentNode As TreeNode))
                    Dim itemEid As String = t.Eid
                    If String.IsNullOrEmpty(itemEid) AndAlso t.SubFolder IsNot Nothing Then itemEid = t.SubFolder.EntryID
                    If Not String.IsNullOrEmpty(itemEid) AndAlso String.Equals(itemEid, targetEid, StringComparison.OrdinalIgnoreCase) Then
                        item.Selected = True
                        item.Focused = True
                        item.EnsureVisible()
                        Exit For
                    End If
                Catch ex As System.InvalidCastException
                    ' 忽略不支援轉換為該 Tuple 結構的 Tag 項目 (例如標題或合計列)
                End Try
            End If
        Next
    End Sub
    Private Async Sub ComputeFolderSize(sender As Object, e As EventArgs)
        _isUserBusy = True
        _dbg(" ├ 開始", $"選取項目數: {ListView1.SelectedItems.Count}")

        Dim cToken As CancellationToken = Nothing   ' 2026/07/07 by Simon/Claude: 宣告提到 Try 外供 Finally 歸還(未取用時為 None，OkeyNowByeByeToken 對 None 不動作)
        Try
            Dim stopwatch As Stopwatch = Stopwatch.StartNew()

            ' by Gemini 3.5 Flash, 2026/06/27: 若有選取項目則僅計算選取者；若無選取項目，則預設計算 ListView1 內的所有項目
            Dim targetItems As New List(Of ListViewItem)()
            If ListView1.SelectedItems.Count > 0 Then
                For Each item As ListViewItem In ListView1.SelectedItems : targetItems.Add(item) : Next
            Else
                For Each item As ListViewItem In ListView1.Items : targetItems.Add(item) : Next
            End If

            If targetItems.Count > 0 Then
                cToken = OkayNowYouHaveToken()      ' 2026/07/07 by Simon/Claude: 宣告移至 Try 外，取用時機不變
                For Each s As ListViewItem In targetItems
                    If s.Tag Is Nothing Then Continue For ' 排除標題列或合計列
                    If s.SubItems.Count > 4 Then s.SubItems(4).Text = "計算中..." Else s.SubItems.Add("計算中...")
                    ' 提高反應速度, 先占位 (如果已經有FolderSize的子項目就先把它改成「計算中...」, 如果還沒有就先加一個占位用的子項目)
                Next

                Dim swThrottle As Stopwatch = Stopwatch.StartNew()
                ' 僅統計有 Tag（有效資料夾）的項目數量
                Dim totalCount As Integer = 0
                For Each s As ListViewItem In targetItems
                    If s.Tag IsNot Nothing Then totalCount += 1
                Next
                Dim processedCount As Integer = 0

                For Each s As ListViewItem In targetItems
                    'If s.Index = 0 Then Continue For ' 一樣, 若選中本體目錄則跳過 (之前統計速度還很慢的時候, 怕計算量太大跑太久)
                    ' 2026/04/13 by Simon/Claude: Tag 升級為 ValueTuple (SubFolder, ParentNode)；群組標題行 / 合計列 Tag=Nothing，直接跳過
                    If s.Tag Is Nothing Then Continue For

                    ' 2026/07/04 by Simon/Claude Fable 5 [去物化配套]: ① fsa 記憶體命中 → 0 COM(GetBfsResult 的 fsa DB 補讀已預熱);
                    '   未命中才 ResolveLv1TagFolder 物化單一資料夾走原 GetFolderSizeAll 路徑(該路徑本來就要打 COM 算大小)。
                    Dim t = DirectCast(s.Tag, (SubFolder As Folder, Eid As String, Sid As String, FPath As String, ParentNode As TreeNode))
                    Dim fp As String = If(String.IsNullOrEmpty(t.FPath), SafeGetPath(t.SubFolder), t.FPath)
                    Dim folderSize As Long
                    If Not _cacheFolderSizeAll.TryGetValue(fp, folderSize) Then
                        Dim folder As Folder = ResolveLv1TagFolder(t)
                        If folder Is Nothing Then Continue For
                        folderSize = Await GetFolderSizeAll(folder, fp, cToken:=cToken)
                    End If

                    Dim strFolderSize As String
                    ' 2026/07/04 by Simon/Claude: 改用 FormatFolderSizeStr 統一格式(0 才會顯示 "- ",避免與 BuildLv1Item 顯示不一致)
                    If folderSize < 0 Then : strFolderSize = "計算失敗"
                    Else : strFolderSize = FormatFolderSizeStr(fp)
                    End If
                    If s.SubItems.Count > 4 Then s.SubItems(4).Text = strFolderSize Else s.SubItems.Add(strFolderSize)

                    processedCount += 1
                    Await SmartThrottle(swThrottle, cToken:=cToken, ThrottleFreq.Hii,
                                              Sub() PgrsBar2.Text = $"正在計算資料夾大小: {processedCount:N0} / {totalCount:N0} ({ExtractFolderName(fp)})")
                Next
            End If

            PgrsBar2.Text = "統計資料夾大小花費了 " & stopwatch.Elapsed.TotalSeconds.ToString("0.00") & " 秒。"
        Catch ex As OperationCanceledException
            PgrsBar2.Text = "計算已由使用者中斷。"
            _dbg(" ├ 中斷", "ComputeFolderSize 已中斷")
        Catch ex As System.Exception
            PgrsBar2.Text = "發生錯誤: " & ex.Message
            _dbg(" ├ 錯誤", ex.Message)
        Finally
            OkeyNowByeByeToken(cToken)   ' 2026/07/07 by Simon/Claude: 歸還 token — 運算中判定 token 化
            _isUserBusy = False
            _dbg("結束")
        End Try

    End Sub
    Private Function FormatFolderSizeStr(fPath As String) As String
        ' 2026/05/27 by Simon/Claude: 抽出 BuildLv1GroupHeader / BuildLv1Item 重複的大小字串格式化
        Dim sizeVal As Long
        If _cacheFolderSizeAll.TryGetValue(fPath, sizeVal) AndAlso sizeVal > 0 Then
            Return (sizeVal / 1024 ^ 2).ToString(If(sizeVal >= 1024 ^ 2, "N0", "N2")) & " MB" ' 2026/6/27 by simon: 根據 mbSize 是否大於等於 1，動態決定格式是要 "N0" 還是 "N2"
        End If
        Return "- "
    End Function
    Private Async Sub EnterSelectedFolder(selectedItem As ListViewItem)
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
        '   ① parentNode.Expand()            → 只展開父節點，載入真實子節點 (一次 BeforeExpand，正確)
        '   ② 在真實子節點裡找 foundNode      → 不遞迴 (FindNodeByName 每個節點都 Expand()，已知錯誤)
        '   ③ foundNode.Tag判斷folder.count  → 確認目標資料夾有子資料夾才進入
        '   ④ SimTree1 選取方法              → 更新 _selectedNodes + 觸發 AfterSelect
        '
        ' ★ 2026/04/13 by Simon/Claude: SimTree1 升級後，Tag 改為 ValueTuple:
        '   群組標題行 & 合計列 → Tag = Nothing → 直接 Return，不進入
        '   一般子資料夾行     → Tag = (SubFolder, ParentNode) → 從 ParentNode 直接取得父節點，不再依賴 TreeView1.SelectedNode
        '
        ' ★ 2026/04/30 by Simon/Claude: 修正④步 SendMessage TVM_SELECTITEM 的根本缺陷
        '   原本用 Win32 TVM_SELECTITEM 觸發 TVN_SELCHANGED → AfterSelect，但此路徑
        '   完全繞過 SimTree 內部的 _selectedNodes 清單，導致 AfterSelect 讀到的
        '   SimTree1.SelectedNodes 仍是舊選取（父資料夾），不是 foundNode，統計算錯資料夾。
        '   修正：改呼叫 SimTree 自己的選取方法，與 GotoDefaultInbox 的路徑一致。
        _dbg(" ├ 開始", selectedItem.SubItems(0).Text)

        ' 群組標題行 / 合計列的 Tag 是 Nothing，不可進入
        If selectedItem.Tag Is Nothing Then Return

        ' 從 ValueTuple Tag 取得子資料夾與其父 TreeNode
        ' 2026/07/04 by Simon/Claude Fable 5 [去物化配套]: Tag 改 5 欄 id-tuple;EntryID 直接取 Tag.Eid(0 COM),BFS 路徑不再需要 SubFolder
        Dim t = DirectCast(selectedItem.Tag, (SubFolder As Folder, Eid As String, Sid As String, FPath As String, ParentNode As TreeNode))
        Dim parentNode As TreeNode = t.ParentNode
        If parentNode Is Nothing Then Return

        ' ① 確保父節點已展開 (若只有 ":::" 則展開一次，載入真實子節點)
        parentNode.Expand()

        ' ② 在直屬子節點裡找目標 (不遞迴，不呼叫任何 Expand)
        ' 2026/04/22 by Gemini 3 Flash: UI 顯示字串有格式化前綴與防切邊空白，改用 EntryID 進行精確匹配，避免搜尋失敗。
        Dim targetEntryID As String = t.Eid
        If String.IsNullOrEmpty(targetEntryID) AndAlso t.SubFolder IsNot Nothing Then targetEntryID = t.SubFolder.EntryID   ' F5 路徑保底
        Dim foundNode As TreeNode = Nothing
        _dbg("    ├ 搜尋節點", $"目標 EntryID: '{targetEntryID}', 父節點: '{parentNode.Text}', 子節點數: {parentNode.Nodes.Count}")

        For Each node As TreeNode In parentNode.Nodes
            Dim nodeFolder As Folder = TryCast(node.Tag, Folder)
            ' 2026/07/04 by Simon/Claude Fable 5: EntryID 比對改不分大小寫 — Tag.Eid 可能來自 RDO 骨架 hex 字串,與 OOM EntryID 大小寫保險對齊
            If nodeFolder IsNot Nothing AndAlso String.Equals(nodeFolder.EntryID, targetEntryID, StringComparison.OrdinalIgnoreCase) Then
                foundNode = node : Exit For
            End If
        Next
        If foundNode Is Nothing Then Return     'dbg("    ├ 錯誤", "找不到對應的子節點")

        ' ③ 確認目標資料夾有子資料夾才進入
        Dim targetFolder As Folder = TryCast(foundNode.Tag, Folder)
        Dim fc As Integer = If(targetFolder IsNot Nothing, GetFolderCount(targetFolder), -1)
        _dbg("    ├ 檢查", $"找到節點: '{foundNode.Text}', TargetFolder: IsNot Nothing = {targetFolder IsNot Nothing}, 子資料夾數: {fc}")
        If targetFolder Is Nothing OrElse fc <= 0 Then
            _dbg("    ├ 放棄進入", "目標無子資料夾或 Tag 型別錯誤")
            Return
        End If
        foundNode.EnsureVisible()

        ' ④ 2026/05/13 by Gemini 3 Flash: 改用 Await 同步統計與渲染，解決導覽後的焦點競爭
        SimTree1.ClearSelectedNodes()
        SimTree1.AddSelectedNode(foundNode)

        Dim deduped As List(Of TreeNode) = SimTree1.GetDedupedSelection()   ' 2026/07/10 by Simon/Claude: 改用 SimTree 內建版
        Dim cToken As CancellationToken = OkayNowYouHaveToken()             ' 2026/07/07 by Simon/Claude: 行內取用改具名，才能在 Finally 歸還(運算中判定 token 化)
        Try
            Dim items As List(Of ListViewItem) = Await CollectTab1FolderInfo(deduped, cToken)
            RenderLv1(items)
        Finally
            OkeyNowByeByeToken(cToken)
        End Try

        ' by Gemini 3.5 Flash, 2026/06/27: 進入資料夾後，自動呼叫 ComputeFolderSize 計算該層各個子資料夾的大小
        ComputeFolderSize(Nothing, Nothing)

        ListView1.Focus()
        If ListView1.Items.Count > 0 Then
            ListView1.Items(1).Selected = True : ListView1.Items(1).Focused = True
        End If
        _dbg("結束")

    End Sub
#End Region
#End Region

#Region "■ 05 Tab2: 依日期統計"
    ' ==============================================================
    ' 重構目標: COM/UI/流程邏輯與業務分離清晰分層，去除全域狀態，優化快取機制
    ' 1. 分層架構: 將原本混在一起的程式碼重構成三個明確的層次
    '    - Layer 1 (UI 事件層)  : 回應使用者操作，組裝參數後交給 Layer2 執行，最後把結果交給顯示函數
    '    - Layer 2 (流程協調層) : BFS 遍歷 fList，管理快取，驅動 Layer3 計算，合併結果，回報進度
    '    - Layer3 (COM 資料層) : 對 Outlook 發出 COM 呼叫，回傳單一資料夾的年份郵件分佈
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
    '     - GetYearCountsAsync_CL()                 (已由 CollectYearCount 取代)
    '     - CountMailByYearAsync_CLayer2()          (已由 GetYearCountsForFolderAsync 取代)
    '     - UpdateCounterProgress()                 (已改成 callback 機制，函數可刪除)
    '     - ShowProgressTab2()                      (簽章已更改，請替換)(2026/4/12 重構 v2 已刪除)
    '
    ' ==============================================================
    ' 2026/04/12 重構 v2 (render層拆分+導覽函數整合):
    '   刪除: ShowYearView, ShowMonthView, ShowResultTab2, ShowProgressTab2
    '         UpdateChart2forYearView, UpdateChart2forMonthView
    '   新增: CollectMonthCount            ← 月份資料收集 Layer2 (純計算，不碰UI) 
    '         GoToLv2YearView()                ← 共用導覽：返回年度視圖 (純 render from cache，DoubleClick/KeyPress 共用) 
    '         GoToLv2MonthView(year, cToken)   ← 共用導覽：進入月份視圖 (方案A cache，DoubleClick/KeyPress 共用) 
    '         RenderLv2YearView              ← 年度 ListView render (純UI，不計算) 
    '         RenderLv2MonthView             ← 月份 ListView render (純UI，不計算) 
    '         RenderCt2YearView              ← 年度 Chart render (純UI，不計算) 
    '         RenderCt2MonthView             ← 月份 Chart render (純UI，不計算) 
    '   更改: Form1.vb HandleListViewKeyPress 的 ShowYearView/ShowMonthView 改呼叫 GoToLv2YearView/GoToLv2MonthView
    ' ==============================================================
    ' 分層架構 (更新後):
    '   Layer 1 (UI 事件層)       : SimTree2_AfterSelect, CheckSubFolder2_CheckStateChanged
    '                                Lv2_MouseDoubleClick (只兩行，委派 GoToLv2YearView/GoToLv2MonthView) 
    '   Layer 2 (流程協調層)      : CollectYearCount, CollectMonthCount
    '                                RenderLv2YearView, RenderCt2YearView
    '                                RenderLv2MonthView, RenderCt2MonthView
    '                                GoToLv2YearView, GoToLv2MonthView
    '   Layer3 (COM 資料層)      : GetYearCountRdo/OOM, GetMonthCountRdo/OOM (Module_Outlook.vb)
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
        _dbg("開始") : Dim stopwatch As Stopwatch = Stopwatch.StartNew()    ' 開始計時，初始化畫面狀態; by Claude Sonnet 4.6, 2026/06/07
        Cursor = Cursors.WaitCursor : PgrsBar1.Text = "" : PgrsBar2.Text = ""

        ' 序號機制: 每次點選遞增；計算完成後若序號已變，代表有更新的點選，丟棄本次結果
        Dim mySeq As Integer = System.Threading.Interlocked.Increment(_tab2SelectSeq)

        ' 取得 SimTree2 多選清單 (SelectedNodes 是 SimTree 提供的 List(Of TreeNode))
        Dim selectedNodes As List(Of TreeNode) = SimTree2.SelectedNodes
        If selectedNodes Is Nothing OrElse selectedNodes.Count = 0 Then
            _dbg("結束", "無節點被選取")
            Cursor = Cursors.Default : Return           ' 選擇節點為空，直接結束
        End If

        Dim targetFolderList =                          ' 把所有已選 TreeNode 的 Tag 轉換成 Outlook.Folder，過濾掉無效節點
            selectedNodes.Select(Function(n) TryCast(n.Tag, Folder)).Where(Function(f) f IsNot Nothing).ToList()
        If targetFolderList.Count = 0 Then
            _dbg("結束", "所有選定節點均無效資料夾")
            Cursor = Cursors.Default : Return           ' 如果沒有任何有效的資料夾 (List.Count=0) 就直接結束
        End If

        Dim cToken As CancellationToken = OkayNowYouHaveToken()  ' ✅ 取得新 Token，同時取消上一次未完成的操作
        ' 2026/07/07 by Simon/Claude: 取用時機由函式開頭下移到早退檢查之後 — token 化的運算中判定上線後，
        '   早退路徑不再有「取了沒還」的殘留 token；取用/歸還(下方 Finally)自此成對。
        Try ' by Claude Opus, 2026/04/11: Try 上移，包住 GetSubtree 的 Await，否則 ESC 時拋出的 OperationCanceledException 沒有被捕捉
            Dim progressTree = New Progress(Of ProgressReport)(Sub(p) PgrsBar2.Text = p.Message)
            Dim folderList = Await GetUniqueFolderList(selectedNodes, _includeSubTab2, progress:=progressTree, cToken:=cToken)
            _lv2IsMonthView = False        ' 切換選取時，重置視圖狀態為年度視圖
            _tv2FolderList = folderList    ' ✅ 記住本次統計的資料夾清單，供 GoToLv2MonthView (CollectMonthCount) 使用
            ' 2026/04/16 by Gemini: 這裡的 f.fPath 已經是 Tuple 屬性，完全無 COM 開銷
            _tv2FolderPaths = folderList.Select(Function(f) f.fPath).ToList() ' ★ 記住對應路徑 (by Gemini 3.1 Pro, 2026/04/15)

            ''Dim totalMailCount As Integer =                                                   ' 計算所有選定根資料夾的郵件總數作為進度分母
            ''    If(CheckSub2.Checked, rootFolders.Sum(Function(f) GetMailCountRecursive(f)),  ' CheckSubFolder2.Checked = True  → 含子資料夾: 各自完整子樹的總和
            ''                          rootFolders.Sum(Function(f) GetMailCountOOM(f)))         ' CheckSubFolder2.Checked = False → 只算選定的那一層
            ''' 2026/3/20, 重寫了底層GetMailCountAll() 效能還是比不過現在上面的遞迴版本
            '' 原因: 原版遞迴只走一遍 COM 資料夾樹，新版走了兩遍COM:
            '' 第一遍: GetSubtree()  → BFS 遍歷，存取每個 folder.Folders
            '' 第二遍: For Each allFolders → GetMailCountOOM() 再讀每個資料夾一次

            ' --- 計算所有選定根資料夾的郵件總數，作為 CollectYearCount 進度條的分母
            ' 2026/04/16 by Gemini: 這裡優化為直接對 folderList (已展開的子資料夾) 進行一圈快速統計
            Dim totalMailCount As Long = 0
            Dim processedCountLocal As Integer = 0
            Dim totalFoldersLocal As Integer = folderList.Count
            Dim swThrottleCount As Stopwatch = Stopwatch.StartNew()  ' by Claude Sonnet 4.6, 2026/06/07

            For i As Integer = 0 To folderList.Count - 1
                ' ✅ 使用 Tuple 內的 .Folder 與 .FolderPath，效能從 400ms 降至近乎 0ms
                Dim c As Integer = GetMailCount(_tv2FolderPaths(i), folderList(i).eid, folderList(i).sid)   ' 2026/06/28 Stage2: 免-folder 多載
                If c > 0 Then totalMailCount += c
                processedCountLocal += 1
                ' 2026/04/16 by Gemini: 每 100 毫秒更新一次預計計數進度
                Await SmartThrottle(swThrottleCount, cToken:=cToken, ThrottleFreq.Hii,
                                          Sub() PgrsBar2.Text = $"正在計算郵件分母: {processedCountLocal:N0}/{totalFoldersLocal:N0} 個資料夾 (累計 {totalMailCount:N0} 封)...")
            Next

            ' 呼叫 Layer2 流程協調層執行統計；結果存入 _lv2DataYear session 快取，GoToLv2MonthView/GoToLv2YearView 直接 render 不重算
            ListView2.Tag = totalMailCount    ' ★ 把總計數量快取起來，供 CollectMonthCount 回報進度分母使用 (by Gemini 3.1 Pro, 2026/04/15)
            Dim progressYear = New Progress(Of ProgressReport)(Sub(p) PgrsBar2.Text = p.Message)
            _lv2DataYear = Await CollectYearCount(folderList, totalMailCount, progressYear, cToken:=cToken, _tv2FolderPaths)

            ' --- 序號校驗點 2 (核心運算完成後) ---
            If _tab2SelectSeq <> mySeq Then Return  ' _dbg("結束", "序號已不匹配，丟棄本次結果 (運算完畢中斷) ")
            stopwatch.Stop()                        ' ✅ 統計完成後才停錶

            ' 2026/04/12: ShowResultTab2 + ShowProgressTab2 拆分為 Render 函數 + inline progress
            RenderLv2YearView(_lv2DataYear)
            RenderCt2YearView(_lv2DataYear)

            Dim _yTotal As Integer = _lv2DataYear.Values.Sum   ' Values.Sum 是最可靠的實際計數 (含/不含子資料夾皆正確) 
            Dim _ySpd As Double = If(stopwatch.Elapsed.TotalSeconds > 0, _yTotal / stopwatch.Elapsed.TotalSeconds, 0)
            PgrsBar1.Text = $"共 {_yTotal:N0} 封 / {stopwatch.Elapsed.TotalSeconds:0.00} 秒"
            PgrsBar2.Text = $"(年度統計完成 - 處理速度為 {_ySpd:N0}/sec)"
            sender.Enabled = True : sender.Focus() : Cursor = Cursors.Default
            _dbg("結束")
        Catch ex As OperationCanceledException
            _dbg("結束", "ESC 中斷")
            PgrsBar1.Text = "由使用者中斷。" : PgrsBar2.Text = "" : Cursor = Cursors.Default
        Catch ex As System.Exception
            _dbg("錯誤", ex.Message) : Cursor = Cursors.Default
        Finally
            OkeyNowByeByeToken(cToken)   ' 2026/07/07 by Simon/Claude: 歸還 token — 運算中判定 token 化
        End Try
    End Sub
    Private Async Sub Lv2_KeyDown(sender As Object, e As KeyEventArgs) Handles ListView2.KeyDown
        ''' <summary>
        ''' ListView2: 年度 / 月份視圖導覽 (2026/04/16 by Gemini 3.1 Pro: 從 HandleListViewKeyPress 拆分回歸)
        ''' </summary>
        _dbg("開始", $"鍵值: {e.KeyCode}")
        ' 2026/07/07 by Simon/Claude: 原本在此每次按鍵都 OkayNowYouHaveToken()，但只有「進入月份視圖」分支用到 —
        '   token 化的運算中判定上線後改移入該分支內取用+歸還，避免按鍵殘留 _cts。
        Dim lv As ListView = DirectCast(sender, ListView)

        If e.KeyCode = Keys.Enter Then                ' Enter = 等同雙擊目前選定的項目
            If lv.SelectedItems.Count = 0 Then Return

            ' by Gemini 3 Flash, 2026/04/13: 選取多個項目時，改用 MessageBox 顯示數量加總
            If lv.SelectedItems.Count > 1 Then
                Dim sumYearMonth As Long = 0
                For Each item As ListViewItem In lv.SelectedItems
                    ' 跳過特殊控制列或標題列 (例如 "BACK" 或含有 "──")
                    If item.Tag?.ToString() = "BACK" OrElse item.Text.Contains("──") Then Continue For
                    ' SubItems(1): 郵件個數
                    Dim strCount As String = item.SubItems(1).Text.Replace(",", "").Trim()
                    Dim valCount As Long = 0
                    Long.TryParse(strCount, valCount)
                    sumYearMonth += valCount
                Next
                MessageBox.Show($"已選取 {lv.SelectedItems.Count:N0} 個統計項目：" & vbCrLf & vbCrLf &
                                $"郵件總計：{sumYearMonth:N0} 封", "複選數量加總", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim selectedItem As ListViewItem = lv.SelectedItems(0)
            If _lv2IsMonthView AndAlso                     ' 在月份視圖按 Enter 於返回列 → 回到年度視圖
                selectedItem.Tag IsNot Nothing AndAlso
                selectedItem.Tag.ToString() = "BACK" Then
                Try
                    Await GoToLv2YearView()
                Catch ex As OperationCanceledException
                    _dbg("中斷", "GoToYearView 中斷")
                End Try

            ElseIf Not _lv2IsMonthView Then                ' 在年度視圖按 Enter → 進入月份視圖
                Dim selectedYear As Integer = 0
                If Integer.TryParse(selectedItem.Text.Trim(), selectedYear) AndAlso
                    _tv2FolderList IsNot Nothing AndAlso _tv2FolderList.Count > 0 Then
                    Dim cToken As CancellationToken = OkayNowYouHaveToken()   ' 2026/07/07 by Simon/Claude: 從函式開頭移入本分支(唯一用到 token 的地方)
                    Try
                        Await GoToLv2MonthView(selectedYear, cToken:=cToken)
                    Catch ex As OperationCanceledException
                        _dbg("結束", "ESC 中斷")
                        PgrsBar1.Text = "由使用者中斷。" : PgrsBar2.Text = "" : Cursor = Cursors.Default
                    Finally
                        OkeyNowByeByeToken(cToken)   ' 2026/07/07 by Simon/Claude: 歸還 token
                    End Try
                End If
            End If
            e.Handled = True
            e.SuppressKeyPress = True

        ElseIf e.KeyCode = Keys.Escape Or e.KeyCode = Keys.Back Then    ' 2026/04/22 by Gemini 3.1 Pro: 補上 ESC 退出邏輯 ' 2026/6/27 by simon: 加上Backspace
            If _lv2IsMonthView Then
                ' 從月份視圖退回年度視圖
                Try
                    Await GoToLv2YearView()
                Catch ex As OperationCanceledException
                    _dbg("中斷", "GoToYearView 中斷")
                End Try
            Else
                ' 從年度視圖退回左側資料夾樹
                SimTree2.Focus()
            End If
            e.Handled = True
            e.SuppressKeyPress = True
        ElseIf e.Control AndAlso e.KeyCode = Keys.A Then                ' Ctrl-A 全選 listview2 所有項目
            LviSelectAll(lv, e)

        ElseIf e.Control AndAlso e.KeyCode = Keys.C Then                ' Ctrl-C 複製選取列到剪貼簿 (by Claude Sonnet 4.6, 2026/04/27)
            LviCopyToClipboard(lv, e)

        End If
        If _iLikeNoisy Then _dbg("結束")

    End Sub
    Private Async Sub Lv2_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles ListView2.MouseDoubleClick
        ' ---------------------------------------------------------------
        ' ListView2 雙擊事件 (2026/04/12 重構：委派 GoToLv2YearView / GoToLv2MonthView，消除與 HandleListViewKeyPress 的重複) 
        ' 年度視圖: 雙擊某一年 → GoToLv2MonthView(selectedYear, cToken:=cToken)
        ' 月份視圖: 雙擊「← 返回」→ GoToLv2YearView()
        ' ✅ cToken 重構: 每次進入都 OkayNowYouHaveToken() 取得全新 cToken (同 2026/04/11 原設計) 
        ' ---------------------------------------------------------------
        _dbg("開始")
        Dim clickedItem As ListViewItem = ListView2.GetItemAt(e.X, e.Y)
        If clickedItem Is Nothing Then Return

        Dim cToken As CancellationToken = OkayNowYouHaveToken()
        Try
            If _lv2IsMonthView AndAlso clickedItem.Tag?.ToString() = "BACK" Then
                Await GoToLv2YearView() : Return
            End If
            Dim selectedYear As Integer = ParseYearFromText(clickedItem.Text)
            If selectedYear = 0 OrElse _tv2FolderList Is Nothing OrElse _tv2FolderList.Count = 0 Then Return
            Await GoToLv2MonthView(selectedYear, cToken:=cToken)
            _dbg("結束", $"{selectedYear} 年")
        Catch ex As OperationCanceledException
            _dbg("結束", "ESC 中斷")
            PgrsBar1.Text = "由使用者中斷。" : PgrsBar2.Text = "" : Cursor = Cursors.Default
        Catch ex As System.Exception
            _dbg("錯誤", ex.Message)
        Finally
            OkeyNowByeByeToken(cToken)   ' 2026/07/07 by Simon/Claude: 歸還 token — 運算中判定 token 化
        End Try

    End Sub
    Private Sub Lv2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListView2.SelectedIndexChanged
        ' ---------------------------------------------------------------
        ' ListView2 選取變更 ↔ Chart2 對應長條同步高亮
        ' 年度視圖: 選取某年 → 高亮 Chart2 中對應年份的長條
        ' 月份視圖: 選取某月 → 高亮 Chart2 中對應月份的長條
        ' 與 Ct2_MouseMove 共用 _lastHoveredPointIndex，確保兩者高亮互斥不累積
        ' 注意: Ct2_MouseLeave 會清掉 _lastHoveredPointIndex，
        '       滑鼠離開圖表後 ListView 的選取高亮也會消失
        ' 2026-03-18, by Claude.ai / 2026-04-04 by Gemini: 共用函數重構
        ' ---------------------------------------------------------------
        If ListView2.SelectedItems.Count = 0 Then Return
        _dbg("開始")
        If Chart2.Series.Count = 0 OrElse Chart2.Series(0).Points.Count = 0 Then Return  ' Chart 尚未載入資料，直接結束

        Dim selectedItem As ListViewItem = ListView2.SelectedItems(0)
        Dim selectedText As String = selectedItem.Text.Trim()

        ' ── 防護特殊控制列 ──
        If selectedItem.Tag?.ToString() = "BACK" Then Return
        If selectedText.Contains("──") Then Return

        ' ── 找出目標 DataPoint index ──
        Dim targetIndex As Integer = -1
        If Not _lv2IsMonthView Then
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
            If selectedMonth > 0 Then targetIndex = selectedMonth - 1 ' RenderCt2MonthView 依 1~12 月順序加入 DataPoints，月份N = index N-1
        End If

        If targetIndex < 0 OrElse targetIndex >= Chart2.Series(0).Points.Count Then Return

        ' ── 呼叫共用的圖表渲染函數套用高亮 ──
        BrushCt2HoverState(Chart2, targetIndex)

    End Sub
    Private Sub Ct2_MouseClick(sender As Object, e As MouseEventArgs) Handles Chart2.MouseClick
        ' ---------------------------------------------------------------
        ' Chart2 點擊長條 → 同步高亮 ListView2 對應的年份或月份列
        ' 反向對應: Lv2_SelectedIndexChanged 負責 ListView → Chart2
        ' 設定 item.Selected = True 會觸發 Lv2_SelectedIndexChanged，
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
        If Not _lv2IsMonthView Then
            targetItem = FindLv2ItemByYear(CInt(pt.XValue))             ' 年度視圖: pt.XValue = 年份
        Else
            Dim monthNum As Integer = ParseMonthFromText(pt.AxisLabel)  ' 月份視圖: 呼叫共用函數解析
            If monthNum > 0 Then targetItem = FindLv2ItemByMonth(monthNum)
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
    Private Sub Ct2_MouseMove(sender As Object, e As MouseEventArgs) Handles Chart2.MouseMove
        ' ✅ 用 MouseMove，滑鼠移動時持續觸發，才能追蹤到每個長條
        Dim chart As Chart = CType(sender, Chart)
        If chart.Series.Count = 0 OrElse chart.Series(0).Points.Count = 0 Then Return

        Dim hit As HitTestResult = chart.HitTest(e.X, e.Y)
        If hit.ChartElementType = ChartElementType.DataPoint Then
            BrushCt2HoverState(chart, hit.PointIndex)
        Else
            ClearCt2HoverState(chart)   ' 滑鼠離開所有長條，還原上一個點與標題
        End If

    End Sub
    Private Sub Ct2_MouseLeave(sender As Object, e As EventArgs) Handles Chart2.MouseLeave
        ' 滑鼠離開 Chart2，還原上一個高亮點與標題
        Dim chart As Chart = CType(sender, Chart)
        ClearCt2HoverState(chart)
    End Sub
    Private Sub CheckSubFolder2_CheckStateChanged(sender As Object, e As EventArgs) Handles CheckSubFolder2.CheckStateChanged
        _dbg("開始", _includeSubTab2.ToString)

        ' by Gemini, 2026/03/29: 合併為 SimTree2 單一操作路徑
        Dim selectedNodes As List(Of TreeNode) = SimTree2.SelectedNodes
        If selectedNodes IsNot Nothing AndAlso selectedNodes.Count > 0 Then SimTree2_AfterSelect(SimTree2, New TreeViewEventArgs(selectedNodes(0)))
        _dbg("結束")
    End Sub
#End Region
#Region "  ├ Layer2 流程協調層"
    Private Async Function CollectYearCount(fList As List(Of (eid As String, sid As String, fPath As String)), totalMailCount As Long, progress As IProgress(Of ProgressReport), cToken As CancellationToken, Optional fPaths As List(Of String) = Nothing) As Task(Of ConcurrentDictionary(Of Integer, Integer))   ' 2026/06/28 Stage2: 合約改 (eid,sid,fPath)
        ' ---------------------------------------------------------------
        ' === Layer 2: 流程協調層 ===
        ' 職責: BFS 遍歷 fList，管理快取，驅動 Layer3 計算，合併結果，回報進度
        '       逐資料夾計算年份統計並合併，是 Tab2 所有統計流程的唯一入口
        ' 規則: 不直接碰 UI 控制項 (ProgressBar1 等)，進度透過 onProgress callback 傳出, 自己不會知道上一層是單選還是多選，只知道接受傳入的 fList 清單
        '
        ' 參數:
        '   fList           : 由 Layer1 組裝好的目標資料夾清單 (已包含 BFS 展開結果，2026/06/28 Stage2 純資料 Tuple)
        '   totalMailCount  : 總郵件數，用來計算進度百分比的分母
        '   onProgress      : 進度 callback，每處理完一個資料夾呼叫一次，回傳 (已處理, 總數)
        '   cToken          : CancellationToken，由 Layer1 透過 OkayNowYouHaveToken() 取得，ESC 時拋 OperationCanceledException
        ' ---------------------------------------------------------------
        _dbg(" ├ 開始", $"目標資料夾數: {fList.Count:N0}")

        Dim swThrottle As Stopwatch = Stopwatch.StartNew()  ' by Claude Sonnet 4.6, 2026/06/07
        Dim processedCount As Integer = 0
        Dim processedFolders As Integer = 0
        Dim totalFolders As Integer = fList.Count

        Dim merged As New ConcurrentDictionary(Of Integer, Integer)
        Try
            For i As Integer = 0 To totalFolders - 1
                ' ✅ 2026/06/28 Stage2: 直接從純資料 Tuple 取得 eid, sid, fPath
                Dim eid As String = fList(i).eid
                Dim sid As String = fList(i).sid
                Dim fPath As String = fList(i).fPath

                ' ✅ 2026/04/10: 提前過濾沒有信件的資料夾 (by Gemini) 既然根本沒有信，就不必去查 DB 或打 COM，直接跳過
                ' 改用免-folder 多載，完全避免觸碰 COM 物件
                If GetMailCount(fPath, eid, sid) <= 0 Then ' 放個空快取避免下次又查 (<= 0 也包含 -1 的情況視同沒信防護)
                    _cacheYearCount(fPath) = New ConcurrentDictionary(Of Integer, Integer)()
                    MarkMailFolderDirty(fPath)   ' 2026/07/03 by Simon/Claude: dirty 追蹤 (空表本身不會產生 DB 列，標記只是保持行為一致)
                    processedFolders += 1 : Continue For
                End If

                ' 2026/04/17 by Claude: 改呼叫 GetYearCount (L2.5)，移除原本內嵌的①②③快取邏輯
                ' ①記憶體命中 ②DB lazy ③Layer3 COM(RDO優先,失敗fallback OOM) 全部封裝在 GetYearCount 內，與其他 L2.5 cache proxy layer 一致
                ' OCE re-throw 由 GetYearCount → GetYearCountOOM 往上冒泡，被本層 Catch OCE 接住(RDO 路徑不拋 OCE)
                ' 2026/06/29 by Simon/Claude [Stage2]: 改傳 id-tuple,眼物化移除——folder 由免-folder 多載延後到 ③ 才建
                Dim folderResult As ConcurrentDictionary(Of Integer, Integer) = Await GetYearCount(fPath, eid, sid, cToken:=cToken)

                merged = MergeDictionaries(merged, folderResult)    ' 把這個資料夾的結果合併到總計 (純 .NET 運算，不碰 COM)
                processedCount += folderResult.Values.Sum()         ' 累加已處理郵件數，透過 callback 通知 Layer1 更新進度顯示
                processedFolders += 1

                ' 2026/04/16 by Gemini 3.0 flash: 改用 ThrottleFreq.Hii + SmartThrottle 與 onThrottled 委派
                Await SmartThrottle(swThrottle, cToken:=cToken, ThrottleFreq.Hii,
                                          Sub() progress?.Report(New ProgressReport With {.CurrentCount = processedCount, .TotalCount = totalMailCount,
                                                                                .Message = $"正在統計年度分佈: ({processedFolders:N0}/{totalFolders:N0})個資料夾 (已統計 {processedCount:N0} / {totalMailCount:N0} 封信)..."}))
            Next
        Catch ex As OperationCanceledException
            ' by Gemini, 2026/04/11: 捕捉 ESC 中斷，回傳已計算的部分結果而不拋出異常
            _dbg(" ├ 中斷", "CollectYearCount 已中斷")
        End Try
        _dbg(" ├ 結束", $"共 {merged.Count:N0} 個年份 | 郵件總計: {merged.Values.Sum():N0}") ' by Gemini, 2026/04/10
        Return merged

    End Function
    Private Async Function CollectMonthCount(selectedYear As Integer, cToken As CancellationToken) As Task(Of ConcurrentDictionary(Of Integer, Integer))
        ' ---------------------------------------------------------------
        ' 月份資料收集 Layer2 (2026/04/12 由 ShowMonthView 計算部分拆出) 
        ' 職責: 遍歷 _tv2FolderList，對每個資料夾呼叫 GetMonthCount(RDO優先,失敗fallback OOM)，合併結果，回報進度
        '       不碰 UI render (render 由 GoToLv2MonthView 的 RenderLv2MonthView / RenderCt2MonthView 負責) 
        '       cToken 與 CollectYearCount 同理，都需要傳入以支援 ESC 中斷
        '       OperationCanceledException 由 caller (GoToLv2MonthView → DoubleClick / HandleListViewKeyPress) 的 Catch 攔截
        ' ---------------------------------------------------------------
        _dbg(" ├ 開始", selectedYear.ToString())

        Dim monthCount As New ConcurrentDictionary(Of Integer, Integer)
        Dim totalFolders As Integer = _tv2FolderList.Count
        Dim processedFolders As Integer = 0
        Dim totalMailCount As Long = ListView2.Tag    ' ★ 直接取用快取好的分母，省掉整個 For Each GetMailCount 迴圈

        ' 逐資料夾取月份分布並合併
        Dim swThrottle As Stopwatch = Stopwatch.StartNew()  ' by Claude Sonnet 4.6, 2026/06/07
        For i As Integer = 0 To totalFolders - 1
            Dim eid As String = _tv2FolderList(i).eid
            Dim sid As String = _tv2FolderList(i).sid
            Dim fPath As String = _tv2FolderList(i).fPath
            processedFolders += 1
            ' 2026/04/15 by Gemini 3.1 Pro: 傳入快取好的 fPath，消除 GetMonthCount 內的 COM 開銷
            ' 2026/04/17 by Claude: 改呼叫 GetMonthCount (L2.5)，提前過濾/快取/DB lazy 全封裝於內
            ' 2026/06/29 by Simon/Claude [Stage2]: 改傳 id-tuple,眼物化移除——folder 由免-folder 多載延後到 ③ 才建
            Dim folderMonthCount As ConcurrentDictionary(Of Integer, Integer) = Await GetMonthCount(fPath, eid, sid, selectedYear, cToken:=cToken)
            monthCount = MergeDictionaries(monthCount, folderMonthCount)

            ' 2026/04/16 by Gemini 3.0 flash: 改用 ThrottleFreq.Hii + SmartThrottle 與 onThrottled 委派，移除 OrElse processedFolders=totalFolders 特判
            Await SmartThrottle(swThrottle, cToken:=cToken, ThrottleFreq.Hii,
                                      Sub()
                                          PgrsBar1.Text = "正在讀取..."
                                          PgrsBar2.Text = $"正在統計 {selectedYear} 年月份分佈: ({processedFolders:N0}/{totalFolders:N0})個資料夾 (相依包含共計 {totalMailCount:N0} 封信)。"
                                      End Sub)
        Next
        _dbg(" ├ 結束", $"{selectedYear} 年 | 月份數: {monthCount.Count:N0}")
        Return monthCount

    End Function
    Private Async Function GoToLv2YearView() As Task
        ' ---------------------------------------------------------------
        ' 共用導覽：返回年度視圖 (2026/04/12 取代 ShowYearView，供 DoubleClick 與 KeyPress 共用) 
        ' 職責: 純 render from _lv2DataYear session 快取，完全不碰 COM / Layer2 計算層
        '       _lv2MonthViewYear 刻意不 reset，讓 _lv2DataMonth 方案A快取跨 back-and-forth 繼續有效
        ' ---------------------------------------------------------------
        _dbg(" ├ 開始")
        Dim yearToRestore As Integer = _lv2MonthViewYear   ' 先記住要還原游標的年份
        _lv2IsMonthView = False
        Await Task.Yield()  ' 讓 UI 喘口氣，確保畫面流暢切換

        If _lv2DataYear IsNot Nothing AndAlso _tv2FolderList IsNot Nothing AndAlso _tv2FolderList.Count > 0 Then
            Cursor = Cursors.WaitCursor
            RenderLv2YearView(_lv2DataYear)
            RenderCt2YearView(_lv2DataYear)

            Dim _rTotal As Integer = _lv2DataYear.Values.Sum
            PgrsBar1.Text = $"共 {_rTotal:N0} 封"
            PgrsBar2.Text = "(返回年度統計)"
            Cursor = Cursors.Default
        End If

        ' 還原游標到進入月份前的那一年，讓使用者感覺回到剛才看的地方
        ' ✅ 2026-03-16 Bug fix: 移除此處多餘的 item.Selected = True (造成兩個 highlighted item 的根源) 
        If yearToRestore > 0 AndAlso ListView2.Items.Count > 0 Then
            Dim tgt = FindLv2ItemByYear(yearToRestore)
            If tgt IsNot Nothing Then
                tgt.Selected = True : tgt.Focused = True : tgt.EnsureVisible()
                ListView2.Focus()
            End If
        End If
        _dbg(" ├ 結束")

    End Function
    Private Async Function GoToLv2MonthView(selectedYear As Integer, cToken As CancellationToken) As Task
        ' ---------------------------------------------------------------
        ' 共用導覽：進入月份視圖 (2026/04/12 取代 ShowMonthView，供 DoubleClick 與 KeyPress 共用) 
        ' 職責: 方案A _lv2DataMonth 快取判斷 → 命中時純 render；未命中時 CollectMonthCount → render
        '       _lv2MonthViewYear 同時作為「目前顯示年份」與「方案A快取 tag」
        ' OperationCanceledException 由 caller (DoubleClick / HandleListViewKeyPress) 的 Catch 攔截
        ' ---------------------------------------------------------------
        _dbg(" ├ 開始", selectedYear.ToString())
        Dim swM As Stopwatch = Stopwatch.StartNew()  ' by Claude Sonnet 4.6, 2026/06/07
        PgrsBar1.Text = "" : PgrsBar2.Text = "" : Cursor = Cursors.WaitCursor

        If _lv2DataMonth IsNot Nothing AndAlso _lv2MonthViewYear = selectedYear Then
            ' ★ 快取命中：直接 render，完全不碰計算層 (方案A：同一年份才命中) 
            _dbg(" ├ _lv2DataMonth 快取命中", selectedYear.ToString())
            _lv2IsMonthView = True
            RenderLv2MonthView(selectedYear, _lv2DataMonth)
            RenderCt2MonthView(_lv2DataMonth, selectedYear)

            Dim _mHit As Integer = _lv2DataMonth.Values.Sum
            PgrsBar1.Text = $"共 {_mHit:N0} 封"
            PgrsBar2.Text = $"({selectedYear} 年月份分佈 - 按 ESC 或雙擊標題橫列可返回視圖) "
        Else
            ' ★ 快取未命中：CollectMonthCount → _cacheMonthCount 一定命中 → merge → render
            _dbg(" ├ _lv2DataMonth 快取未命中，開始計算", selectedYear.ToString())
            Dim mc As ConcurrentDictionary(Of Integer, Integer) = Await CollectMonthCount(selectedYear, cToken:=cToken)
            _lv2DataMonth = mc : _lv2MonthViewYear = selectedYear : _lv2IsMonthView = True
            swM.Stop()
            RenderLv2MonthView(selectedYear, mc)
            RenderCt2MonthView(mc, selectedYear)

            Dim _mMiss As Integer = mc.Values.Sum
            Dim _mSpd As Double = If(swM.Elapsed.TotalSeconds > 0, _mMiss / swM.Elapsed.TotalSeconds, 0)
            PgrsBar1.Text = $"共 {_mMiss:N0} 封 / {swM.Elapsed.TotalSeconds:0.00} 秒"
            PgrsBar2.Text = $"({selectedYear} 年月份分佈讀取完成 - 按 ESC 或雙擊標題橫列可返回視圖) "
        End If

        Cursor = Cursors.Default
        _dbg(" ├ 結束", selectedYear.ToString())

    End Function
#End Region
#Region "  ├ UI 渲染"
    Private Sub RenderLv2YearView(yearCount As ConcurrentDictionary(Of Integer, Integer))
        ' ---------------------------------------------------------------
        ' 年度視圖 ListView2 渲染 (2026/04/12 由 ShowResultTab2 拆出) 
        ' 職責: 純 UI render，不做計算，不查快取，不碰 COM
        ' 對稱: RenderCt2YearView 負責同一視圖的 Chart2 部分
        ' ---------------------------------------------------------------
        _dbg(" ├ 開始", yearCount?.Count)

        ListView2.Items.Clear()
        If yearCount Is Nothing OrElse yearCount.IsEmpty Then
            ClearCt2Series()     ' ★ 空資料夾時也要清除 Chart2，否則前一個資料夾的圖表會殘留
            ListView2.Items.Add(New ListViewItem("找不到郵件"))
        Else
            ' 預分配容量為 64，優化 Tab2 年度/月份視圖的 UI 項目渲染 (by Gemini 3 Flash, 2026/05/04)
            Dim items As New List(Of ListViewItem)(64)
            Dim sortedYearCount = yearCount.OrderBy(Function(pair) pair.Key).ToList()
            For Each pair In sortedYearCount
                items.Add(New ListViewItem({pair.Key, pair.Value.ToString("N0") & " "}))
            Next
            ListView2.Items.AddRange(items.ToArray())
        End If
        _dbg(" ├ 結束")

    End Sub
    Private Sub RenderLv2MonthView(selectedYear As Integer, monthCount As ConcurrentDictionary(Of Integer, Integer))
        ' ---------------------------------------------------------------
        ' 月份視圖 ListView2 渲染 (2026/04/12 由 ShowMonthView render 部分拆出) 
        ' 職責: 純 UI render，不做計算，不查快取，不碰 COM
        ' 對稱: RenderCt2MonthView 負責同一視圖的 Chart2 部分
        ' ---------------------------------------------------------------
        _dbg(" ├ 開始", selectedYear.ToString())
        ListView2.Items.Clear()

        ' 預分配容量為 16，減少暫存所有資料夾統計項目的 Resize 次數 (by Gemini 3 Flash, 2026/05/04)
        Dim itemsList As New List(Of ListViewItem)(16)    ' 建立一個 List 來暫存所有的 ListViewItem

        ' 第一行: 返回按鈕
        Dim backItem As New ListViewItem("← 返回年度統計")
        backItem.SubItems.Add("") : backItem.Tag = "BACK"
        backItem.ForeColor = Color.Gray
        backItem.Font = New Font(_fontDefault, _fontItalic)
        itemsList.Add(backItem)

        ' 第二行: 年份標題
        Dim titleItem As New ListViewItem($"── {selectedYear} 年月份分佈 ──")
        titleItem.SubItems.Add($"共 {monthCount.Values.Sum:N0}  封")  ' 字串結尾補上空白防止選取時切邊，與下方對齊
        titleItem.ForeColor = Color.DimGray
        titleItem.Font = New Font(_fontDefault, _fontBold)
        itemsList.Add(titleItem)

        ' 逐月顯示 (只顯示有郵件的月份) 
        For month As Integer = 1 To 12
            Dim count As Integer = 0
            If monthCount.TryGetValue(month, count) AndAlso count > 0 Then ' 稍微優化 TryGetValue 判斷式
                Dim monthItem As New ListViewItem($"{selectedYear} /  {month:D2}月")
                monthItem.SubItems.Add(count.ToString("N0") & " ")  ' 字串結尾一律補一格空白
                itemsList.Add(monthItem)
            End If
        Next
        ListView2.Items.AddRange(itemsList.ToArray())   ' 將收集好的 List 轉為 Array，一次性加入 ListView
        _dbg(" ├ 結束", selectedYear.ToString())

    End Sub
    Private Sub RenderCt2YearView(yearCount As ConcurrentDictionary(Of Integer, Integer))
        ' ---------------------------------------------------------------
        ' 年度視圖 Chart2 渲染 (2026/04/12 由 UpdateChart2forYearView 改名重構) 
        ' 職責: 純 UI render；接受 ConcurrentDictionary，內部自行排序 (原版由 caller 排序後傳 List，現改為自己排序讓介面更乾淨) 
        ' 對稱: RenderLv2YearView 負責同一視圖的 ListView2 部分
        ' ---------------------------------------------------------------
        _dbg(" ├ 開始")
        ClearCt2Series()     ' 清除之前的統計結果，包括 Series Points、平均線 Series、平均值標籤 Annotation
        If yearCount Is Nothing OrElse yearCount.IsEmpty Then Return
        Dim sortedYearCount = yearCount.OrderBy(Function(p) p.Key).ToList()

        ' 添加數據到 Series, 在 Chart2 中顯示統計結果
        Dim series As Series = Chart2.Series(0)
        For Each pair In sortedYearCount
            series.Points.AddXY(pair.Key, pair.Value)
        Next

        ' 依內容大小來設置 Chart2 的 X 軸上下限
        With Chart2.ChartAreas(0).AxisX
            .Minimum = sortedYearCount.Min(Function(p) p.Key) - 0.5
            .Maximum = sortedYearCount.Max(Function(p) p.Key) + 0.5
            .Interval = 1
            .IntervalOffset = 0                 ' ✅ 還原年度視圖的長條置中偏移
            .LabelStyle.Format = "####"         ' ✅ 還原年份格式
            .LabelStyle.Interval = 1
            .LabelStyle.IntervalOffset = 0.5    ' ✅ 校正還原上面 max/min 的 0.5 偏移
            .MajorTickMark.IntervalOffset = 0   ' ✅ 還原刻度偏移
        End With

        ' 添加一條代表平均值的線 (獨立 Series 才能控制線型，StripLine 不支援虛線) 
        ' 2026/3/6 by Claude Code；2026/04/12 移入 RenderCt2YearView
        Dim average As Double = sortedYearCount.Average(Function(pair) pair.Value)
        Dim xMin As Double = sortedYearCount.Min(Function(pair) pair.Key)
        Dim xMax As Double = sortedYearCount.Max(Function(pair) pair.Key)

        Dim avgSeries As New Series("平均線") With {.ChartType = SeriesChartType.Line,
                                                    .Color = ThemeColors.avgLineColor,
                                                    .BorderWidth = 2,
                                                    .BorderDashStyle = ChartDashStyle.Dash,  ' ✅ 虛線
                                                    .ChartArea = Chart2.ChartAreas(0).Name,
                                                    .IsVisibleInLegend = False}
        avgSeries.Points.AddXY(xMin - 1, average)  ' 0: 從 X 軸最小值往左延伸
        avgSeries.Points.AddXY(xMax, average)       ' 1: 圖表最右邊長條的確切 X 座標 (錨定用) 
        avgSeries.Points.AddXY(xMax + 1, average)  ' 2: 到 X 軸最大值往右延伸

        ' 用 TextAnnotation 顯示平均值標籤 (by Gemini, 2026/04/04 改用 DeepAmber 提升辨識度)
        Dim avgLabel As New TextAnnotation With {.Name = "平均值標籤",
                                                 .Text = "AVG: " & average.ToString("N0"),
                                                 .ForeColor = ThemeColors.avgLineColor,
                                                 .Font = New Font("Tahoma", 10.0F, FontStyle.Bold),
                                                 .AnchorDataPoint = avgSeries.Points(1),            ' 錨定在最右側長條的中間點 X 座標
                                                 .AnchorAlignment = ContentAlignment.BottomCenter,  ' ★ 強制對齊點的正上方 (避免 MS Chart 自動亂飄移) 
                                                 .AnchorOffsetX = 0,    ' 保持置中
                                                 .AnchorOffsetY = -1,   ' 產生 1% 的空隙，確保不在線上
                                                 .BackColor = Color.Transparent,
                                                 .LineColor = Color.Transparent}
        Chart2.Series.Add(avgSeries)
        Chart2.Annotations.Add(avgLabel)
        Chart2.Invalidate()     ' 強制重新繪製圖表
        _dbg(" ├ 結束")

    End Sub
    Private Sub RenderCt2MonthView(monthCount As ConcurrentDictionary(Of Integer, Integer), year As Integer)
        ' ---------------------------------------------------------------
        ' 月份視圖 Chart2 渲染 (2026/04/12 由 UpdateChart2forMonthView 改名) 
        ' 月份長條圖：只畫 1~12 月，X 軸標籤顯示「M月」，不畫平均線
        ' 完整替換 Chart2 的內容，與 RenderCt2YearView 平行存在
        ' 對稱: RenderLv2MonthView 負責同一視圖的 ListView2 部分
        ' ---------------------------------------------------------------
        _dbg(" ├ 開始", year)
        ClearCt2Series()     ' 清除之前的所有圖表內容 (同 RenderCt2YearView 的清除邏輯) 

        ' 把 1~12 月的資料全部加入 (沒有郵件的月份補 0，讓 X 軸保持完整 12 格) 
        Dim series As Series = Chart2.Series(0)
        For month As Integer = 1 To 12
            Dim count As Integer = 0
            monthCount.TryGetValue(month, count)
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
        _dbg(" ├ 結束", year)

    End Sub
#End Region
#Region "  └ 輔助函數"
    Private Function BuildFilterDateRangeTab2(year As Integer, Optional mon1 As Integer = 1, Optional mon2 As Integer = 12) As String
        If year < 1974 Then Return Nothing
        'Const DATE_FORMAT As String = "yyyy/MM/dd HH:mm:ss"
        'Dim startDate As Date = Date.ParseExact($"{year}/01/01 00:00:00", DATE_FORMAT, Nothing) ' 建立當年的起始日期和結束日期
        'Dim endDate As Date = Date.ParseExact($"{year}/12/31 23:59:59", DATE_FORMAT, Nothing)   ' 設置結束日期的時間為23:59:59
        'Return $"[RcvTime] >= '{startDate}' AND [RcvTime] <= '{endDate}'"
        ' 2026/3/11, by Claude, 重構BuildFilterDateRangeTab2 函數: 增加了月份參數，並且直接用 Date 物件來建立日期範圍，避免字串格式問題
        Dim startDate As New Date(year, mon1, 1, 0, 0, 0)                                   ' ✅ 用 mon1/mon2 決定起訖月份，預設 1~12 代表整年
        Dim endDate As New Date(year, mon2, Date.DaysInMonth(year, mon2), 23, 59, 59)       ' mon2 的結束日用該月最後一天，避免硬寫 31 日造成 2 月等短月份抓不準
        Return $"[ReceivedTime] >= '{startDate}' AND [ReceivedTime] <= '{endDate}'"

    End Function
    Private Function Find1stYear(selectedFolder As Folder) As Integer
        Dim sName As String = selectedFolder?.Name
        _dbg("開始", sName)
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
            ' 資料夾裡可能混有 MeetingRequest / ContactItem / Note 等, 這些物件沒有 RcvTime
            ' 透過 COM late binding 存取會拋 COMException 或 AccessViolationException (.NET 4+ 的 corrupted state exception)，bare Catch 接不住
            ' ✅ 先 Restrict 過濾掉 null/零值 RcvTime 的壞項目，再升冪排序取最舊年份
            allItems = selectedFolder.Items : If allItems Is Nothing OrElse allItems.Count = 0 Then Return 1974
            validItems = allItems.Restrict("[ReceivedTime] > '1974/01/01'") : If validItems.Count = 0 Then Return 1974
            validItems.Sort("[ReceivedTime]", OlSortOrder.olDescending)
            Dim firstItem As Object = validItems.GetFirst() : If firstItem Is Nothing Then Return 1974
            mail = TryCast(firstItem, Outlook.MailItem) : If mail Is Nothing Then Return 1974
            Dim year As Integer = mail.ReceivedTime.Year : Return If(year <= 0 OrElse year > Date.Today.Year, 1974, year)
        Catch ex As System.Exception
            _dbg("錯誤", sName & " - " & ex.Message)
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
        Dim yearIdx As Integer = cleanText.IndexOf("年"c)
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
        Dim moonIdx As Integer = text.IndexOf("月"c)
        If moonIdx < 0 Then Return 0

        Dim numStr As String = ""
        For k As Integer = moonIdx - 1 To 0 Step -1
            If Char.IsDigit(text(k)) Then numStr = text(k) & numStr Else Exit For
        Next

        Dim selectedMonth As Integer = 0
        If Integer.TryParse(numStr, selectedMonth) AndAlso selectedMonth >= 1 AndAlso selectedMonth <= 12 Then Return selectedMonth
        Return 0

    End Function
    Private Function FindLv2ItemByYear(targetYear As Integer) As ListViewItem
        ' by Gemini, 2026/04/04: 根據年份尋找對應的 ListViewItem
        For Each item As ListViewItem In ListView2.Items
            If ParseYearFromText(item.Text) = targetYear Then Return item
        Next
        Return Nothing
    End Function
    Private Function FindLv2ItemByMonth(targetMonth As Integer) As ListViewItem
        ' by Gemini, 2026/04/04: 根據月份數字尋找對應的 ListViewItem
        Dim monthStr As String = targetMonth.ToString("D2") & "月"  ' e.g. "03月"
        For Each item As ListViewItem In ListView2.Items
            If item.Tag IsNot Nothing AndAlso item.Tag.ToString() = "BACK" Then Continue For
            If item.Text.Contains(monthStr) Then Return item
        Next
        Return Nothing
    End Function

    Private Sub ClearCt2Series()
        ' 清除之前的統計結果, 包括 Series Points 和 平均線 Series 以及平均值標籤 Annotation (避免重複加入)
        Chart2.Series(0).Points.Clear()

        Dim existingAvg As Series = Chart2.Series.FindByName("平均線")       ' 清除舊的平均線 Series (避免重複加入)
        If existingAvg IsNot Nothing Then Chart2.Series.Remove(existingAvg)

        Dim existingAnnotation = Chart2.Annotations.FindByName("平均值標籤") ' 先清除舊的 Annotation (避免重複加入)
        If existingAnnotation IsNot Nothing Then Chart2.Annotations.Remove(existingAnnotation)

    End Sub
    Private Sub BrushCt2HoverState(chart As Chart, pointIndex As Integer)
        ' by Gemini, 2026/04/04: 抽取共用的圖表高亮渲染邏輯
        If pointIndex = _lastHoveredPointIndex Then Return ' 如果跟上次是同一個點就不重複處理，避免閃爍

        ' ✅ 先把上一個點的顏色跟狀態清除，但在這裡不要馬上 Refresh 畫面，等新的屬性上完再一起重繪避免畫面閃爍
        ClearCt2HoverState(chart, refreshChart:=False)

        ' ✅ 把目前這個點變成設定的高亮色
        chart.Series(0).Points(pointIndex).Color = ThemeColors.barHighlight
        _lastHoveredPointIndex = pointIndex

        ' ✅ 取得資料點並計算顯示名稱 (修正年度檢視 AxisLabel 為空的問題)
        Dim dP As DataPoint = chart.Series(0).Points(pointIndex)
        Dim xLabel As String = If(Not String.IsNullOrEmpty(dP.AxisLabel), dP.AxisLabel, dP.XValue.ToString("0000"))
        Dim headerText As String = If(xLabel.Contains("月"c), "月份", "年份")

        ' ✅ 動態數據標籤
        Dim formattedX As String = If(xLabel.Contains("月"c), xLabel, xLabel & "年"c)
        dP.Label = $"{formattedX}:{dP.YValues(0):##0}"
        dP.IsValueShownAsLabel = True
        chart.Series(0).ToolTip = $"{headerText}: {xLabel}, 數量: {dP.YValues(0):##0}"

        chart.Refresh()         ' ✅ 確保標籤與顏色即時更新

    End Sub
    Private Sub ClearCt2HoverState(chart As Chart, Optional refreshChart As Boolean = True)
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

End Class