Imports System.Collections.Concurrent
Imports System.Net
Imports System.Runtime.InteropServices
Imports System.Threading
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Outlook
Imports Redemption

' === 從頭重新設計 Layer3 / Layer2.5 底層計數函數 ===
' 目的: 提供一個純粹的 COM 資料層函數，專注於讀取資料，不做任何流程控制或快取邏輯
'       取代目前散落在各處的 GetMailCountByMAPINew、GetFolderSizeLegacy 等函數，統一為一個簡單的 GetXxxLayer3 函數
' 架構: Layer3 純資料層，Layer2 流程協調層，Layer2.5 快取代理層，Layer1 UI 事件層
'       Layer3 只負責讀取資料夾的本層郵件數 (GetDirectMailCountLayer3) ，不遞迴、不展開子資料夾，最小化 COM 呼叫量
'       上層流程 (如 ComputeFolderStatsAsync) 負責決定何時呼叫、如何使用結果、快取管理等
' ==============================================================
' === Layer3 底層 COM 資料層函數群 ===
' 設計原則:
'   1. 每個函數只負責一件事: 讀取單一資料夾或單封郵件的一種屬性
'   2. 不做快取、不做遞迴、不做 BFS 展開——這些全部交給 Layer2 流程協調層
'   3. Fallback 鏈: RDO → MAPI GetArray → OOM最後手段
'                   parallel.foreach → BFS → Recursive，每層不論成功失敗都丟 Debug message
'   4. 失敗統一回傳 -1 (不回 0) ，讓 Layer2 能區分「真的是 0」或「讀取失敗」
'   5. 在 Finally 中使用 TryMarshalRelease() 統一釋放所有 COM 物件，確保 RCW 不殘留
' ==============================================================

Partial Class Form1

#Region "■ 01 全域宣告"
    Private WithEvents _olApp As Outlook.Application = Nothing
    Private _olNS As Outlook.NameSpace = Nothing
    Private _rdo As Redemption.RDOSession = Nothing ' _rdoSession 就等同是outlook.namespace 的意思, 就是Redemption的MAPI session
    Private _pstStoreList As List(Of Outlook.Store) = Nothing
    ' 2026-03-22 新增: 用於測試 Redemption.dll 整合 (注意: session.MAPIOBJECT 必須在 Outlook MAPI 連線建立後才能設定 (Form1_Load 尾端)
    '------------------------------------------------------------------------------------------------
    ' Outlook 物件(OOM)	    Redemption 物件 (RDO)     說明
    '------------------------------------------------------------------------------------------------
    ' Outlook.Application	Redemption 本體	        Redemption 是底層 MAPI 封裝，它不負責 UI 或視窗管理。
    ' Outlook.NameSpace	    Redemption.RDOSession	最接近。 負責管理登入、StoreID、PST 檔案庫與全域設定。
    ' Outlook.Folder	    Redemption.RDOFolder	對應資料夾層級。
    ' Outlook.MailItem	    Redemption.RDOMail	    對應單封郵件層級。
    ' Outlook.Store	        Redemption.RDOStore	    對應 PST 或 Exchange 帳戶。

    Private Shared _cacheMailCount As New ConcurrentDictionary(Of String, Integer)          ' 自身資料夾的郵件個數
    Private Shared _cacheMailCountAll As New ConcurrentDictionary(Of String, Integer)       ' 整支子樹的所有郵件總數
    Private Shared _cacheFolderCount As New ConcurrentDictionary(Of String, Integer)        ' 自身資料夾的子目錄個數
    Private Shared _cacheFolderCountAll As New ConcurrentDictionary(Of String, Integer)     ' 整支子樹的所有子目錄總數
    Private Shared _cacheFolderSize As New ConcurrentDictionary(Of String, Long)            ' 自身資料夾的郵件大小加總
    Private Shared _cacheFolderSizeAll As New ConcurrentDictionary(Of String, Long)         ' 整支子樹的所有子目錄郵件大小加總

    Private Shared _cacheIsMailFolder As New ConcurrentDictionary(Of String, Boolean)                   ' 資料夾是否為郵件類型
    Private Shared _cacheFolderTree As New ConcurrentDictionary(Of String, List(Of Outlook.Folder))     ' GetSortedSubFolders() 已排序的子資料夾清單
    Private Shared _cacheSubFolderList As New ConcurrentDictionary(Of String, List(Of Outlook.Folder))  ' GetSubtreeToList() 的樹狀展開平坦化清單
    Private Shared _cacheAttachMailList As New ConcurrentDictionary(Of String, FolderCacheTab3)         ' 包含附件的郵件預掃描結果 (速度很快, 不用存入SSD?)
    Private Shared _cacheAttachFilename As New ConcurrentDictionary(Of String, List(Of String))         ' 所有附件檔名清單

    ' by Gemini, 2026/04/10: 專門儲存資料夾的身分標識與屬性標籤，用以橋接 Folder 物件與 SQLite 持久化
    Private Shared _cacheFolderIDs As New ConcurrentDictionary(Of String, (eid As String, sid As String, isMail As Boolean, hasCh As Boolean))

    Private Shared _cacheYearCounts As New ConcurrentDictionary(Of String, ConcurrentDictionary(Of Integer, Integer))
    Private Shared _cacheMonthCounts As New ConcurrentDictionary(Of String, ConcurrentDictionary(Of Integer, Integer))

    Private Structure FolderSortInfo
        ' by Gemini, 2026/03/29: 用於 GetSortedSubFolders 排序優化，減少 COM 屬性讀取次數 (O(N) vs O(N log N))
        Dim FolderObj As Outlook.Folder
        Dim Name As String
        Dim HasChinese As Boolean
    End Structure
    Friend Structure MailItemInfo
        ' 候選郵件的純資料結構 (不帶 COM 物件，不受 GC 影響)
        Dim EntryID As String
        Dim Subject As String
        Dim Size As Long
        Dim ReceivedTime As DateTime
        Dim SenderName As String
        Dim AttachCount As Integer
    End Structure
    Private Structure FolderCacheTab3
        Dim AttachMailList As List(Of MailItemInfo) ' 所有 hasAttachment 候選 (無大小篩選)
        Dim ItemCountSnap As Integer                ' 快取當下的 PR_CONTENT_COUNT，失效偵測用
    End Structure
#End Region

#Region "■ 10 底層 COM 函數群 (新設計，現役主力) "
#Region "  ├ 初始化載入函數"
    Private Sub InitOutlookNamespace()
        _dbg(" ├ 開始") ' by Gemini, 2026/04/10: 調整縮排層級為 Level 1
        ' by Gemini, 2026/04/01: 從 Form1_Load 抽離出 Outlook 初始化邏輯，優化結構並加入 TryMarshalRelease 以防內存洩漏
        ' 1. 檢查系統中是否已經啟動 Outlook
        Dim processes() As Process = Process.GetProcessesByName("OUTLOOK")
        If processes.Length = 0 Then    ' 如果 Outlook 尚未啟動，顯示訊息並關閉應用程式
            MessageBox.Show("請先啟動 Outlook", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information) : End
        End If
        ' 2. 初始化 Outlook 核心物件模型與 MAPI Namespace
        Try
            _olApp = New Outlook.Application
            _olNS = _olApp.GetNamespace("MAPI")
            If _olApp IsNot Nothing AndAlso _olNS IsNot Nothing Then
                _pstStoreList = GetSortedStores(_olNS)  ' 取得所有 PST 檔
            Else
                Throw New System.Exception("無法取得 Outlook Application 或 MAPI Namespace。")
            End If
        Catch ex As System.Exception
            _dbg(" ├ Outlook App OR NameSpace init FAIL", ex.Message) ' by Gemini, 2026/04/10
            TryMarshalRelease(_olNS)
            TryMarshalRelease(_olApp)
            _olApp = Nothing : _olNS = Nothing
            MessageBox.Show("Outlook Object 連接失敗!" & vbCrLf & ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error) : End
        End Try
        _dbg(" ├ 結束") ' by Gemini, 2026/04/10

    End Sub
    Private Sub InitRdoSession()
        ' 3. 初始化 Redemption Session (目前停用，保留開發記錄)
        Try
            ' ── Redemption Session 初始化, 2026-03-22 測試用:
            ' _rdo = New Redemption.RDOSession()  ' _rdoSession 就等同是outlook.namespace 的意思, 就是Redemption的MAPI session
            ' _rdo.MAPIOBJECT = _olNS.MAPIOBJECT  ' 直接 attach 到現有 Outlook MAPI session, 就不會另開視窗 (必須在 objNameSpace 已建立之後才呼叫)
            ' _dbg("Redemption init OK", $"Version={_rdo.Version}") ' 關鍵: 不建新連線，這樣就不會彈出第二個 Outlook 視窗，也不需要登入
            ' 2026/3/27 總算全部寫好RDO的導入,
            ' 但過程中優化了很多東西之後發現, 好像對效能沒有幫助到太多, 反而是演算法的改進才快更多
            ' RDO 的部份好像反而增加了程式碼複雜度跟拖慢啟動速度而已, 所以先關閉不使用
            Dim unused = InitRedemptionSessionWithoutDeclaration()
        Catch ex As System.Exception
            _dbg("Redemption init FAIL", ex.Message)
            TryMarshalRelease(_rdo)
            _rdo = Nothing
        End Try

    End Sub
    Private Async Function InitRedemptionSessionWithoutDeclaration() As Task
        _dbg(" ├ 開始") ' by Gemini, 2026/04/10: 調整縮排層級為 Level 1
        ' 2026-03-23 v3:
        '   Task.Run 包裝保留 (讓 UI 執行緒繼續跑 LoadStoreToTreeView，平行初始化)
        '   第一次執行競爭條件改用 Thread.Sleep(1) 在 Set() 前解決，
        '   確保 AutoDismiss 輪詢 loop 已執行第一次再放行 New RDOSession()
        Try
            If _rdo IsNot Nothing Then Return
            Dim threadStarted As New System.Threading.ManualResetEventSlim(False)
            AutoDismissRedemptionDialog(threadStarted)
            ' 等 AutoDismiss thread 確認輪詢已開始，最多等 500ms
            threadStarted.Wait(500)
            _dbg(" ├ InitRedemption", "AutoDismiss thread 已就緒，開始 New RDOSession") ' by Gemini, 2026/04/10

            ' ✅ Task.Run: UI 執行緒 不阻塞，LoadStoreToTreeView 可以同時跑
            Dim session As Redemption.RDOSession = Nothing
            Await Task.Run(Sub() session = New Redemption.RDOSession())

            ' MAPIOBJECT 必須回 UI 執行緒賦值 (_olNS 是 STA COM 物件)
            session.MAPIOBJECT = _olNS.MAPIOBJECT
            _rdo = session
            _dbg(" ├ Redemption init OK", $"Version={_rdo.Version}") ' by Gemini, 2026/04/10
        Catch ex As System.Exception
            _rdo = Nothing
            _dbg("Redemption init FAIL", ex.Message)
        End Try

    End Function
    Private Sub AutoDismissRedemptionDialog(threadStarted As System.Threading.ManualResetEventSlim)
        ' 自動點掉 Redemption EULA dialog
        ' 使用 WinSpy++ 確認的視窗結構 (2026-03-23) :
        '   視窗 class    = TEULAForm (Delphi VCL 表單) ，title = "Outlook Redemption"
        '   "I agree"     = TRadioButton，text = "I agree"
        '   "I do NOT..." = TRadioButton，text = "I do NOT agree"
        '   "Ok"          = TButton，    text = "Ok"
        '   "Cancel"      = TButton，    text = "Cancel"
        '
        ' v1 (2026-03-23): PostMessage 取代 SendMessage
        '   SendMessage 在 modal dialog 阻塞 UI 執行緒時死結，
        '   PostMessage 非同步丟進佇列，由 dialog 自己的訊息泵處理
        '
        ' v2 (2026-03-23): ShowWindow(SW_HIDE) 立刻隱藏視窗
        '   找到 TEULAForm 後立刻隱藏，使用者看不到閃爍
        '
        ' v3 (2026-03-23): 輪詢間隔 100ms → 5ms，移除固定 Thread.Sleep
        '   改成輪詢等子控制項出現，控制項一建立就立刻動作
        '   Thread.Priority = AboveNormal 確保首次啟動也能及時執行
        '
        ' v4 (2026-03-23): 加入 ManualResetEventSlim 同步點
        '   threadStarted.Set() 通知呼叫端「輪詢已開始」，
        '   呼叫端等到 Set 後才呼叫 New RDOSession()，
        '   解決 thread pool 競爭導致首次執行抓不到視窗的問題
        Dim t As New System.Threading.Thread(
            Sub()
                ' ✅ 先讓輪詢 loop 跑第一次，再通知呼叫端
                '   避免 Set() 後呼叫端立刻 New RDOSession()，
                '   但此 thread 還沒執行到 FindWindow 的競爭條件
                System.Threading.Thread.Sleep(1)
                threadStarted.Set()
                Dim hWnd As IntPtr = IntPtr.Zero
                Dim timeout As Integer = 0

                ' 輪詢找 TEULAForm，最多等 30 秒 (3000 × 10ms)
                Do While hWnd = IntPtr.Zero AndAlso timeout < 3000
                    hWnd = FindWindow("TEULAForm", Nothing)
                    If hWnd = IntPtr.Zero Then
                        System.Threading.Thread.Sleep(5)
                        timeout += 1
                    End If
                Loop

                If hWnd = IntPtr.Zero Then
                    _dbg("    ├ AutoDismissRedemption", "逾時: 找不到 TEULAForm") : Return ' by Gemini, 2026/04/10: 調整縮排層級為 Level 2 (輔助執行緒)
                End If

                ' ✅ 立刻隱藏，使用者不會看到 EULA dialog 閃出來
                ShowWindow(hWnd, SW_HIDE)
                _dbg("    ├ AutoDismissRedemption", $"TEULAForm 隱藏 hWnd=0x{hWnd:X}") ' by Gemini, 2026/04/10

                ' ── Step 1: "I agree" TRadioButton ──────────────────────
                ' 輪詢等子控制項建立完成 (視窗已隱藏，等待時間使用者無感)
                Dim hAgree As IntPtr = IntPtr.Zero
                Dim childTimeout As Integer = 0
                Do While hAgree = IntPtr.Zero AndAlso childTimeout < 50
                    hAgree = FindWindowEx(hWnd, IntPtr.Zero, "TRadioButton", "I agree")
                    If hAgree = IntPtr.Zero Then
                        System.Threading.Thread.Sleep(5) : childTimeout += 1
                    End If
                Loop

                If hAgree <> IntPtr.Zero Then
                    PostMessage(hAgree, WM_LBUTTONDOWN, New IntPtr(1), IntPtr.Zero)
                    PostMessage(hAgree, WM_LBUTTONUP, New IntPtr(1), IntPtr.Zero)
                    _dbg("    ├ AutoDismissRedemption", "'I agree' PostMessage 送出") ' by Gemini, 2026/04/10
                Else
                    _dbg("    ├ AutoDismissRedemption", "找不到 'I agree' (已逾時) ") ' by Gemini, 2026/04/10
                End If

                ' ── Step 2: "Ok" TButton ────────────────────────────────
                Dim hOk As IntPtr = IntPtr.Zero
                Dim okTimeout As Integer = 0
                Do While hOk = IntPtr.Zero AndAlso okTimeout < 50
                    hOk = FindWindowEx(hWnd, IntPtr.Zero, "TButton", "Ok")
                    If hOk = IntPtr.Zero Then
                        System.Threading.Thread.Sleep(5) : okTimeout += 1
                    End If
                Loop

                If hOk <> IntPtr.Zero Then
                    PostMessage(hOk, WM_LBUTTONDOWN, New IntPtr(1), IntPtr.Zero)
                    PostMessage(hOk, WM_LBUTTONUP, New IntPtr(1), IntPtr.Zero)
                    _dbg("    ├ AutoDismissRedemption", "'Ok' PostMessage 送出") ' by Gemini, 2026/04/10
                Else
                    _dbg("    ├ AutoDismissRedemption", "找不到 'Ok' (已逾時) ") ' by Gemini, 2026/04/10
                End If
            End Sub)

        t.Priority = System.Threading.ThreadPriority.AboveNormal
        t.IsBackground = True
        t.Start()

    End Sub
#End Region
#Region "  ├ 資料樹展開 & BFS"
    Private Function GetSortedStores(space As Outlook.NameSpace) As List(Of Outlook.Store)
        ' ==========================================
        ' 取得排序後的 NameSpace 下所有Outlook.Store
        ' 包含目前config內的所有帳號和所有開啟的PST檔
        '
        ' ⚠️ 注意: 不在這裡 ReleaseComObject(space)
        '           space 就是外層的 objNameSpace，釋放後其他地方 (Tab2/Tab3 等) 再用 objNameSpace 會觸發 RCW 已釋放的例外
        '           objNameSpace 的生命週期就只由 Form1_FormClosing 統一管理
        ' ==========================================
        _dbg(" ├ 開始", space.CurrentProfileName) ' by Gemini, 2026/04/10: 調整縮排層級為 Level 2 (由 InitOutlookNamespace 呼叫)

        ' 遍歷所有Outlook.Store並添加到列表中, 使用LINQ擴充方法就夠快了, 不再使用非同步或Parallel.Foreach了
        Dim stores As List(Of Outlook.Store) = space.Stores.Cast(Of Outlook.Store)().ToList()
        stores = stores.OrderBy(Function(st) If(TextHasChineseChar(st.DisplayName), 1, 0)).ThenBy(Function(st) st.DisplayName).ToList() ' 使用 LINQ 排序Outlook.Store

        _dbg(" ├ 結束", $"Profile={space.CurrentProfileName} | 庫數量: {stores.Count}") ' by Gemini, 2026/04/10
        Return stores

    End Function
    Private Function GetSortedSubFolders(folder As Outlook.Folder) As List(Of Outlook.Folder)
        ' ==========================================
        ' 取得引數folder下的所有subFolders並排序後傳回
        ' 優化紀錄: 2026/03/29 by Gemini 3.1 Pro
        ' 1. 加入 Layer3 過濾: 只保留郵件目錄 (olMailItem)，排除行事曆/聯絡人等
        ' 2. 單次屬性讀取: 先快取 Name 後排序，避開 LINQ 重複打 COM 的 N log N 效能陷阱
        ' ==========================================
        If _iLikeNoisy Then _dbg(" ├ 開始", folder.Name)
        Dim fPath As String = folder.FolderPath, value As List(Of Outlook.Folder) = Nothing
        If _cacheFolderTree.TryGetValue(fPath, value) Then Return value

        ' ② DB lazy load: 優先從 SSD 載入 ID 並點對點取回物件 (by Gemini, 2026/04/10)
        ' ----------------------------------------------------------------------
        Dim dbIDs = DbGetOrderedSubFolderIDs(fPath, _showAllFolders) ' by Gemini, 2026/04/10: 改用全域變數以提升效能，避免頻繁讀取 UI 狀態
        If dbIDs IsNot Nothing Then
            Dim dbFolders As New List(Of Outlook.Folder)
            For Each row In dbIDs
                Try ' GetFolderFromID 是點對點查詢，不需遞迴，極快
                    Dim f = TryCast(_olNS.GetFolderFromID(row.eid, row.sid), Outlook.Folder)
                    If f IsNot Nothing Then dbFolders.Add(f)
                Catch : End Try ' 若 EntryID 變更或路徑已不存在，默默跳過就好
            Next
            If dbFolders.Count > 0 Then
                _cacheFolderTree(fPath) = dbFolders
                If _iLikeNoisy Then _dbg("    ├ SSD Hit", $"{folder.Name}: 已從資料庫載入 {dbFolders.Count} 個子目錄") ' by Gemini, 2026/04/10: 調整路徑判斷
                Return dbFolders
            End If
        End If

        ' ③ 傳統 COM 掃描: 快取未命中或 SSD 無紀錄時走原路徑 (OOM 遍歷)
        ' ----------------------------------------------------------------------
        Dim infoList As New List(Of FolderSortInfo)
        Try
            ''' 2024/5/13記錄: 已經試過很多種優化, 好像很難再比現在下面這二行LINQ還快了??
            ''' Dim subFolders As List(Of Outlook.Folder) = folder.Folders.Cast(Of Outlook.Folder)().ToList()
            ''' subFolders = subFolders.OrderBy(Function(subF) If(TextHasChineseChar(subF.Name), 1, 0)).
            '''                         ThenBy(Function(subF) subF.Name).ToList()
            ''' [上面是舊版紀錄] 原本使用 LINQ 直接 Cast().ToList() 後排序，缺點是 OrderBy 會重複觸發 COM 讀取屬性
            ''' 
            ' [下面是新版優化] by Gemini, 2026/03/29:
            ' 1. 動態過濾: 根據 checkShowAllFolders.Checked 決定是否顯示非郵件目錄
            ' 2. 單次屬性讀取: 先快取 Name 後排序，避開 LINQ 重複打 COM 的 N log N 效能陷阱

            For Each subF As Outlook.Folder In folder.Folders
                Dim isMail As Boolean = IsMailFolder(subF) ' 這裡已具備記憶體快取
                ' 🔥 核心過濾: 正常若「沒勾選顯示全部」且「不是郵件資料夾」時就排除
                If Not _showAllFolders AndAlso Not isMail Then Continue For ' by Gemini, 2026/04/10: 合併為 AndAlso 邏輯以簡化結構

                Dim fName As String = subF.Name
                Dim hasCh As Boolean = TextHasChineseChar(fName)
                infoList.Add(New FolderSortInfo With {.FolderObj = subF, .Name = fName, .HasChinese = hasCh})

                ' by Gemini, 2026/04/10: 登記身分標識與屬性標記，供 Save Cache 時持久化到 SSD 以繞過 COM OOM 遍歷
                _cacheFolderIDs.TryAdd(subF.FolderPath, (subF.EntryID, subF.StoreID, isMail, hasCh))
            Next
        Catch ex As System.Exception
            _dbg(" ├ GetSortedSubFolders 遍歷失敗", ex.Message) ' by Gemini, 2026/04/10
        End Try

        ' ② 純記憶體排序: 完全不觸發 COM 呼叫 (快速且不卡 UI)
        Dim sortedFolders = infoList.OrderBy(Function(i) If(i.HasChinese, 1, 0)).
                                     ThenBy(Function(i) i.Name, StringComparer.OrdinalIgnoreCase).
                                     Select(Function(i) i.FolderObj).ToList()
        ' 2026/4/7 進一步優化 by Gemini: 加入 StringComparer.OrdinalIgnoreCase 略過語系分析，爆發性提速

        _cacheFolderTree(fPath) = sortedFolders
        If _iLikeNoisy Then _dbg(" ├ 結束", $"{folder.Name} | 子資料夾數: {sortedFolders.Count}") ' by Gemini, 2026/04/10
        Return sortedFolders

    End Function
    Private Sub LoadStoreToTreeView(storeList As List(Of Outlook.Store), tv As TreeView)
        _dbg(" ├ 開始", tv.Name) ' by Gemini, 2026/04/10: 調整縮排層級為 Level 1

        ''' 2024/5/17全部重寫, 只先動態載入一層的rootFolder, 不花時間遍歷所有的subFolders
        ''' 2024/5/19試過Task.Run(), Parallel.Foreach跟LINQ擴充方法了, 都沒有比較快, 別再試了, 就算virtual mode也沒有比我現在的lazy load還快
        ''tv.BeginUpdate()
        ''For Each st In storeList
        ''    Dim root As Outlook.Folder = st.GetRootFolder
        ''    'Dim node As TreeNode = Await Task.Run(Function() Me.Invoke(Function() tv.Nodes.Add(root.Name)))
        ''    Dim node As TreeNode = tv.Nodes.Add(root.Name)
        ''    node.Tag = root
        ''    If root.Folders.Count > 0 Then node.Nodes.Add(":::") '若發現底下還有subFolders也不讀取, 只先填入一個假的":::"暫代, 才能出現"+"號
        ''Next
        ''tv.EndUpdate()

        ' 2024/5/20昨天才說不會更快了, 今天改用Nodes.AddRange(), 又更快了一點, 連BeginUpdate/EndUpdate都不需要了
        ' 遍歷 storeList 並創建節點, 加進List而不是直接加到Treeview.Nodes
        Dim nodeList As New List(Of TreeNode) ' 創建一個 TreeNode 的 List 來暫存所有要添加的節點
        For Each store In storeList
            Dim root As Outlook.Folder = store.GetRootFolder
            Dim node As New TreeNode(root.Name) With {.Tag = root}
            node.Nodes.Add(":::")  ' ✅ 無條件加佔位節點，省掉判斷 root.Folders.Count 這一次多餘的 COM 往返
            nodeList.Add(node)
            ' PST root folder 幾乎 100% 都有子資料夾，這個假設安全；
            ' 就算 PST 真的空了，展開時 LoadSubFolderToTreeView 清除 ":::" 後不加任何子節點，節點就會自動收起 "+" 號，行為正確
        Next

        tv.Nodes.AddRange(nodeList.ToArray()) ' 將所有組裝好的節點一次性添加到 tv.Nodes
        _dbg(" ├ 結束", $"{tv.Name} 共 {nodeList.Count} 個 Store") ' by Gemini, 2026/04/10

    End Sub
    Private Sub LoadSubFolderToTreeView(sender As Object, e As TreeViewCancelEventArgs)
        _dbg(" ├ 開始", sender.Name) ' by Gemini, 2026/04/10: 調整縮排層級為 Level 1
        ' 2024/5/17全部重寫, 把現在要點開的資料夾, 讀出其子資料夾並加載進treeview
        ' 5/19試過Task.Run(), Parallel.Foreach跟LINQ擴充方法了, 都沒有比較快, 別再試了, 就算virtual mode也沒有比我現在的lazy load還快
        Dim selectedNode As TreeNode = e.Node                   ' 取得點選的node
        Dim selectedFolder As Outlook.Folder = selectedNode.Tag ' 取得點選的資料夾
        Dim sortedFolders = GetSortedSubFolders(selectedFolder) ' 取得所有子資料夾並排序
        If selectedNode.Nodes.Count = 1 AndAlso selectedNode.FirstNode.Text = ":::" Then
            selectedNode.Nodes.Clear()  '清除原本暫代的假node ":::"
            ' 5/20昨天才說不會更快了, 今天改用Nodes.AddRange(), 又更快了一點, 連BeginUpdate/EndUpdate都不需要了
            ' 遍歷 storeList 並創建節點, 先加進List而不是直接加到Treeview.Nodes
            Dim nodeList As New List(Of TreeNode) ' 創建一個 TreeNode 的 List 來暫存所有要添加的節點
            For Each folder As Outlook.Folder In sortedFolders
                Dim node As New TreeNode(folder.Name) With {.Tag = folder}
                Try
                    'If GetCachedFolderCount(folder) > 0 Then node.Nodes.Add(":::")
                    If HasSubFoldersFast(folder) Then node.Nodes.Add(":::") ' 2026/4/7 by Gemini, 光速版子資料夾加號預測 (專為 TreeView 展開設計)
                Catch ex As System.Exception : End Try
                nodeList.Add(node) ' 先加進List在記憶體中快速操作, 而不是直接加到Treeview.Nodes
            Next
            selectedNode.Nodes.AddRange(nodeList.ToArray()) ' 將所有節點一次性添加到 selectedNode.Nodes
        End If
        _dbg(" ├ 結束", $"{selectedFolder.Name} 展開 {sortedFolders.Count} 個子資料夾") ' by Gemini, 2026/04/10

    End Sub

    Private Async Function GetSubtreeToList(rootFolder As Outlook.Folder, includeSubF As Boolean, Optional progress As IProgress(Of ProgressReport) = Nothing, Optional cToken As CancellationToken = Nothing) As Task(Of List(Of Outlook.Folder))    ' by Claude, 2026/04/11: 改為 Async，加 cToken 參數以支援 ESC 中斷
        ' --------------------------------------------------------------
        ' GetSubtreeToList: 取得目標資料夾下, 整個資料夾子樹清單 (BFS，含子資料夾)
        ' ① OOM BFS: 目前唯一的路徑，使用 Outlook Object Model 廣度優先搜尋
        ' by Gemini, 2026/04/02: 導入 IProgress 支援
        ' --------------------------------------------------------------
        If _iLikeNoisy Then _dbg(" ├ 開始", rootFolder.Name) ' by Gemini, 2026/04/10: 調整縮排層級為 Level 1
        Dim sw As New Stopwatch() : sw.Start()
        Dim swThrottle As New Stopwatch() : swThrottle.Start()

        Dim result As New List(Of Outlook.Folder)
        If includeSubF Then
            Dim cachedList As List(Of Outlook.Folder) = Nothing
            If _cacheSubFolderList.TryGetValue(rootFolder.FolderPath, cachedList) Then
                sw.Stop()
                _dbg(" ├ 結束", $"{rootFolder.Name} (Cache Hit) | 資料夾總計: {cachedList.Count} | {sw.ElapsedMilliseconds}ms") ' by Gemini, 2026/04/10
                Return cachedList
            End If

            ' ② DB lazy load: 利用 LIKE 一取回整棵樹的 ID 並重建物件 (by Gemini, 2026/04/10)
            Dim dbIDs = DbGetSubFolderIDList(rootFolder.FolderPath, _showAllFolders)
            If dbIDs IsNot Nothing Then
                Dim dbFolders As New List(Of Outlook.Folder)
                For Each row In dbIDs
                    Try
                        ' 更正命名空間變數: _mapiNameSpace -> _olNS
                        Dim f = TryCast(_olNS.GetFolderFromID(row.eid, row.sid), Outlook.Folder)
                        If f IsNot Nothing Then dbFolders.Add(f)
                    Catch
                    End Try
                Next
                If dbFolders.Count > 0 Then
                    _cacheSubFolderList(rootFolder.FolderPath) = dbFolders
                    sw.Stop()
                    If _iLikeNoisy Then _dbg("    ├ SSD Hit (Tree)", $"{rootFolder.Name}: 已從資料庫載入 {dbFolders.Count} 個子目錄") ' by Gemini, 2026/04/10: 調整縮排層級為 Level 2
                    Return dbFolders
                End If
            End If
        End If

        ' ③ 傳統 OOM BFS 掃描: 快取未命中或 SSD 無紀錄時走原路徑
        result.Add(rootFolder)
        If Not includeSubF Then
            sw.Stop()
            _dbg(" ├ 結束", $"{rootFolder.Name} (Single) | {sw.ElapsedMilliseconds}ms") ' by Gemini, 2026/04/10
            Return result     ' 若不包含子資料夾，直接回傳只有 rootFolder 的清單
        End If

        ' 取得目標資料夾清單 (BFS，含子資料夾)
        Dim queue As New Queue(Of Outlook.Folder)
        queue.Enqueue(rootFolder)
        Try
            While queue.Count > 0
                Dim current As Outlook.Folder = queue.Dequeue()
                Try
                    For Each subF As Outlook.Folder In current.Folders
                        Dim isMail As Boolean = IsMailFolder(subF) ' 這裡已具備記憶體快取
                        ' 🔥 核心過濾: 正常若「沒勾選顯示全部」且「不是郵件資料夾」時就排除
                        If Not _showAllFolders AndAlso Not isMail Then Continue For ' by Gemini, 2026/04/10: 合併為 AndAlso 邏輯以簡化結構

                        ' by Gemini, 2026/04/10: 登記子資料夾 ID 進 Layer 2.5 ID 快取，供 SSD 持久化使用
                        Dim fPath As String = subF.FolderPath
                        Dim eID As String = subF.EntryID
                        Dim sID As String = subF.StoreID
                        Dim fName As String = subF.Name
                        ' ✅ 優化 by Gemini, 2026/04/10: 先將常用 COM 屬性取成區域變數，避免同一行內重複讀取 (simon: 但這個好像跟速度優化沒什麼關係?)
                        _cacheFolderIDs.TryAdd(fPath, (eID, sID, isMail, TextHasChineseChar(fName)))
                        result.Add(subF)       ' 把子資料夾加入結果清單
                        queue.Enqueue(subF)    ' 把子資料夾加入佇列，繼續往下搜尋
                    Next

                Catch ex As System.Exception
                    _dbg("    ├ ① OOM 失敗", current.Name & " - " & ex.Message) ' by Gemini, 2026/04/10: 調整縮排層級為 Level 2 (迴圈內)
                End Try

                ' by Gemini, 2026/04/02: 100ms 節流回報已發現的資料夾數
                ' by Claude, 2026/04/11: 函數已改為 Async，移除舊 _cancelRequested 旗標，改用 Task.Delay(1,cToken) 統一處理 ESC 中斷
                If swThrottle.ElapsedMilliseconds >= 100 Then
                    progress?.Report(New ProgressReport With {.CurrentCount = result.Count, .Message = $"正在展開資料夾結構: 已發現 {result.Count} 個資料夾..."})
                    swThrottle.Restart()
                    Await Task.Delay(1, cToken)  ' by Claude, 2026/04/11: Async 化後可用 Await 讓消息泵有機會處理 ESC 中斷
                End If
            End While
        Catch ex As OperationCanceledException
            ' 2026/04/11 by Claude: 改為 re-throw，確保 folderList 不完整時上層的 for loop 能即時中止，
            ' 不繼續用殘缺的清單計算 totalMailCount 或 CollectYearCounts。
            ' (原本 catch 後直接 End Try 再 Return result，回傳殘缺清單繼續被使用)
            _dbg(" ├ 中斷", $"GetSubFolderList 已由使用者中斷，已發現 {result.Count} 個")
            Throw
        Catch ex As System.Exception
            _dbg(" ├ ② OOM BFS 失敗", ex.Message) ' by Gemini, 2026/04/10: 調整縮排層級為 Level 1 (外層 Try)
        End Try
        sw.Stop()

        If includeSubF AndAlso Not cToken.IsCancellationRequested AndAlso result.Count > 0 Then  ' by Claude, 2026/04/11: 改用 cToken 判斷
            _cacheSubFolderList.TryAdd(rootFolder.FolderPath, result)
        End If

        _dbg(" ├ 結束", $"{rootFolder.Name} (BFS) | 資料夾總計: {result.Count} | {sw.ElapsedMilliseconds}ms") ' by Gemini, 2026/04/10
        Return result

    End Function
    Private Async Function GetUniqueFolderList(selectedNodes As List(Of TreeNode), includeSub As Boolean, cToken As CancellationToken) As Task(Of List(Of Outlook.Folder))
        ''' <summary>
        ''' 共用邏輯 (Tab 2, 3, 4, 5 都有用到)：將多個 TreeNode 轉換為無重複的實體資料夾清單
        '''   對每個選定的根資料夾執行 BFS，合併成一個完整的目標資料夾清單  
        '''   用 HashSet(Of String) 以 FolderPath 去重，避免使用者選到父子資料夾時重複計算
        '''   若Add() 回傳 False 代表已存在，自動去重
        ''' </summary>
        _dbg(" ├ 開始")
        Dim folderList As New List(Of Outlook.Folder)
        Dim addedPaths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        For Each node As TreeNode In selectedNodes
            cToken.ThrowIfCancellationRequested()

            Dim rootF = TryCast(node.Tag, Outlook.Folder)
            If rootF Is Nothing Then Continue For

            For Each subF In Await GetSubtreeToList(rootF, includeSub, cToken:=cToken)  ' 呼叫 Layer 3 的 BFS 取得該節點下的所有子資料夾
                If addedPaths.Add(subF.FolderPath) Then folderList.Add(subF)                ' 利用 HashSet 高效去重（跨節點選取到父子層級時不會重複統計）
            Next
            Await Task.Yield()
        Next
        Return folderList

    End Function
    Private Function GetSubtreeToList_RDO(rootFolder As Redemption.RDOFolder, includeSubF As Boolean) As List(Of Redemption.RDOFolder)
        ' --------------------------------------------------------------
        ' 2026/3/24 by Gemini: GetSubtreeToList_RDO
        ' 目的: 專門提供給 RDO 平行路徑使用，回傳 List(Of Redemption.RDOFolder)
        ' 說明: 因為 Redemption 是 free-threaded，可以用 Parallel.ForEach 安全平行展開子樹
        ' --------------------------------------------------------------
        _dbg("    ├ 開始", rootFolder.Name)
        Dim sw As New Stopwatch() : sw.Start()

        Dim resultBag As New ConcurrentBag(Of Redemption.RDOFolder)
        resultBag.Add(rootFolder)
        If Not includeSubF Then
            sw.Stop()
            _dbg("    ├ 結束", $"{rootFolder.Name} (RDO-Single) | {sw.ElapsedMilliseconds}ms") ' by Gemini, 2026/04/10
            Return resultBag.ToList()
        End If

        ' 使用兩層佇列作層級遍歷，每層用 Parallel.ForEach 探索
        Dim currentLayer As New ConcurrentQueue(Of Redemption.RDOFolder)
        currentLayer.Enqueue(rootFolder)
        Do
            Dim layerList = currentLayer.ToList()
            If layerList.Count = 0 Then Exit Do

            ' 清空 queue 準備裝下一層的資料夾
            Do While currentLayer.TryDequeue(Nothing) : Loop

            ' 平行處理當前層的資料夾，將它們的子資料夾加進 queue 與結果中
            Parallel.ForEach(layerList,
                Sub(current)
                    Try
                        For Each subFolder As Redemption.RDOFolder In current.Folders
                            resultBag.Add(subFolder)
                            currentLayer.Enqueue(subFolder)
                        Next
                    Catch ex As System.Exception
                        _dbg("    ├ 錯誤", current.Name & " - " & ex.Message) ' by Gemini, 2026/04/10
                    End Try
                End Sub)
        Loop

        sw.Stop()
        _dbg("    ├ 結束", $"{rootFolder.Name} (RDO-Parallel BFS) | 資料夾總計: {resultBag.Count} | {sw.ElapsedMilliseconds}ms") ' by Gemini, 2026/04/10
        Return resultBag.ToList()

    End Function
#End Region
#Region "  ├ Layer2.5 快取存取點 (Cache Proxy Layer)"
    ' 2026/03/27 by Gemini: 新增 Layer2.5 快取存取點 (Cache Proxy Layer)，保護 Layer3 不被頻繁呼叫
    ' ---------------------------------------------------------------
    '   - GetCachedMailCount(folder)            ' 單一資料夾郵件數，有 DB lazy + snapshot 驗證
    '   - GetCachedFolderCount(folder)          ' 單一資料夾子資料夾數，有 DB lazy + snapshot 驗證
    '   - GetCachedMailCountAllAsync(folder)    ' 整棵子樹郵件總數，有 DB lazy
    '   - GetCachedFolderCountAllAsync(folder)  ' 整棵子樹資料夾總數，有 DB lazy
    '   - GetCachedFolderSizeAsync(folder)      ' 單一資料夾大小，有 DB lazy
    '   - GetCachedFolderSizeAllAsync(folder)   ' 整棵子樹大小，有 DB lazy
    '   - GetCachedAttachMailList(folder)       ' Tab3 Phase1，有 DB lazy (attach_maillist)
    '   - GetCachedAttachFilename(mail)         ' Tab3 Phase2，有 DB lazy (attach_filenames)
    ' ---------------------------------------------------------------
    ' 2026/04/07: Phase 2 — 在記憶體 miss 時加入 SQLite lazy SELECT，命中後一次填滿所有欄位
    '             寫入仍由 SaveCachesToSQLiteAsync (SaveCache 按鈕) 批次處理，本層不做即時寫入
    '
    ' 2026/4/7 by Claude, 加入「持久化快取」存入SSD:
    '             讀資料時若快取不存在就先去撈SSD, 有就放進快取, 沒有才算cache miss去COM讀取
    ' ---------------------------------------------------------------
    ' 呼叫順序 (每個 GetCachedXxx 函數):
    '   ① 記憶體命中 → 直接回傳（最快，0 COM call）
    '   ② DB 命中 + snapshot 驗證通過 → 填滿記憶體快取 → 回傳（快，0 COM call）
    '   ③ DB miss 或 snapshot 不符 → 呼叫 Layer3 → 填入記憶體快取 → 回傳（慢，有 COM call）
    '
    ' snapshot 驗證: DB 儲存的 content_count_snapshot = save 時的 PR_CONTENT_COUNT 值
    '   用 GetLiveFolderSnap (單次 PropertyAccessor call) 與 snapshot 比對
    '   相同 → 快取仍有效；不同 → 資料夾內容已變，跳過 DB，呼叫 Layer3
    ' ---------------------------------------------------------------
    Private Function GetCachedMailCount(folder As Outlook.Folder) As Integer
        ' ---------------------------------------------------------------
        ' GetCachedMailCount — 單一資料夾本層郵件數 (PR_CONTENT_COUNT)
        ' 2026/4/7 by Claude, 加入Database lazy load, 讀資料順序:
        '   ① 記憶體命中
        '   ② DB lazy load：命中且 mc 有效且 snapshot 吻合 → 一次填滿所有欄位
        '   ③ fallback: Layer3 呼叫
        ' ---------------------------------------------------------------
        Dim count As Integer, fPath As String = folder.FolderPath
        If _cacheMailCount.TryGetValue(fPath, count) Then Return count  ' ① 記憶體命中

        ' ② DB lazy load：命中且 mc 有效且 snapshot 吻合 → 一次填滿所有欄位
        Dim row = DbGetFolderStats(fPath)
        If row IsNot Nothing AndAlso row.mc >= 0 AndAlso GetLiveFolderSnap(folder) = row.snap Then
            FillFolderCacheFromDbRow(fPath, row) : Return row.mc
        End If

        ' ③ fallback: Layer3 呼叫
        count = GetMailCount(folder)
        _cacheMailCount.TryAdd(fPath, count)
        Return count

    End Function
    Private Async Function GetCachedMailCountAllAsync(folder As Outlook.Folder, Optional progress As IProgress(Of ProgressReport) = Nothing) As Task(Of Long)
        ' ---------------------------------------------------------------
        ' GetCachedMailCountAllAsync — 整棵子樹的郵件總數
        ' 2026/4/7 by Claude, 加入Database lazy load, 讀資料順序:
        '   ① 記憶體命中
        '   ② DB lazy load：命中且 mca 有效且 snapshot 吻合 → 一次填滿所有欄位
        '   ③ fallback: Layer3 呼叫
        ' ---------------------------------------------------------------
        _dbg(" ├ 開始", folder.Name)
        Dim count As Integer, fPath As String = folder.FolderPath
        If _cacheMailCountAll.TryGetValue(fPath, count) Then Return count  ' ①

        ' ② DB lazy load（mca 欄位）
        Dim row = DbGetFolderStats(fPath)
        If row IsNot Nothing AndAlso row.mca >= 0 AndAlso GetLiveFolderSnap(folder) = row.snap Then
            FillFolderCacheFromDbRow(fPath, row) : Return row.mca
        End If

        ' ③ fallback: Layer3 呼叫
        Dim total As Long = Await GetMailCountAll(folder, progress)
        If total >= 0 AndAlso Not _cancelRequested Then _cacheMailCountAll.TryAdd(fPath, CInt(total))
        Return total

    End Function
    Private Function GetCachedFolderCount(folder As Outlook.Folder) As Integer
        ' ---------------------------------------------------------------
        ' GetCachedFolderCount — 單一資料夾直屬子資料夾數
        ' 2026/4/7 by Claude, 加入Database lazy load, 讀資料順序:
        '   ① 記憶體命中
        '   ② DB lazy load：命中且 fc 有效且 snapshot 吻合 → 一次填滿所有欄位
        '   ③ fallback: Layer3 呼叫
        ' ---------------------------------------------------------------
        Dim count As Integer, fPath As String = folder.FolderPath
        If _cacheFolderCount.TryGetValue(fPath, count) Then Return count  ' ①

        ' ② DB lazy load（fc 欄位）
        Dim row = DbGetFolderStats(fPath)
        If row IsNot Nothing AndAlso row.fc >= 0 AndAlso GetLiveFolderSnap(folder) = row.snap Then
            FillFolderCacheFromDbRow(fPath, row) : Return row.fc
        End If

        ' ③ fallback: Layer3 呼叫
        count = GetFolderCount(folder)
        _cacheFolderCount.TryAdd(fPath, count)
        Return count

    End Function
    Private Async Function GetCachedFolderCountAllAsync(folder As Outlook.Folder) As Task(Of Integer)
        ' ---------------------------------------------------------------
        ' GetCachedFolderCountAllAsync — 整棵子樹的資料夾總數
        ' 2026/4/7 by Claude, 加入Database lazy load, 讀資料順序:
        '   ① 記憶體命中
        '   ② DB lazy load：命中且 fca 有效且 snapshot 吻合 → 一次填滿所有欄位
        '   ③ fallback: Layer3 呼叫
        ' ---------------------------------------------------------------
        _dbg(" ├ 開始", folder.Name)
        Dim count As Integer, fPath As String = folder.FolderPath
        If _cacheFolderCountAll.TryGetValue(fPath, count) Then Return count  ' ①

        ' ② DB lazy load（fca 欄位）
        Dim row = DbGetFolderStats(fPath)
        If row IsNot Nothing AndAlso row.fca >= 0 AndAlso GetLiveFolderSnap(folder) = row.snap Then
            FillFolderCacheFromDbRow(fPath, row) : Return row.fca
        End If

        ' ③ fallback: Layer3 呼叫
        count = Await GetFolderCountAll(folder)
        If count >= 0 AndAlso Not _cancelRequested Then _cacheFolderCountAll.TryAdd(fPath, count)
        Return count
    End Function
    Private Async Function GetCachedFolderSizeAsync(folder As Outlook.Folder) As Task(Of Long)
        ' ---------------------------------------------------------------
        ' GetCachedFolderSizeAsync — 單一資料夾本層大小 (GetTable 加總)
        ' 2026/3/29 by Gemini: Layer2.5 快取代理層 - 取得單一資料夾本層的大小 (含快取機制)
        ' 2026/4/7 by Claude, 加入Database lazy load, 讀資料順序:
        '   ① 記憶體命中
        '   ② DB lazy load：命中且 fs 有效且 snapshot 吻合 → 一次填滿所有欄位
        '   ③ fallback: Layer3 呼叫
        ' ---------------------------------------------------------------
        Dim size As Long, fPath As String = folder.FolderPath
        If _cacheFolderSize.TryGetValue(fPath, size) Then Return size  ' ①

        ' ② DB lazy load（fs 欄位）
        Dim row = DbGetFolderStats(fPath)
        If row IsNot Nothing AndAlso row.fs >= 0 AndAlso GetLiveFolderSnap(folder) = row.snap Then
            FillFolderCacheFromDbRow(fPath, row) : Return row.fs
        End If

        ' ③ Layer3
        size = Await GetFolderSize(folder)
        If size >= 0 Then _cacheFolderSize.TryAdd(fPath, size)
        Return size

    End Function
    Private Async Function GetCachedFolderSizeAllAsync(folder As Outlook.Folder) As Task(Of Long)
        ' ---------------------------------------------------------------
        ' GetCachedFolderSizeAllAsync — 整棵子樹大小總計
        ' 2026/3/29 by Gemini: Layer2.5 快取代理層 - 取得整棵子樹的大小總計 (含快取機制)
        ' 2026/4/7 by Claude, 加入Database lazy load, 讀資料順序:
        '   ① 記憶體命中
        '   ② DB lazy load：命中且 fsa 有效且 snapshot 吻合 → 一次填滿所有欄位
        '   ③ fallback: Layer3 呼叫
        ' ---------------------------------------------------------------
        _dbg(" ├ 開始", folder.Name)
        Dim size As Long, fPath As String = folder.FolderPath
        If _cacheFolderSizeAll.TryGetValue(fPath, size) Then Return size  ' ①

        ' ② DB lazy load（fsa 欄位）
        Dim row = DbGetFolderStats(fPath)
        If row IsNot Nothing AndAlso row.fsa >= 0 AndAlso GetLiveFolderSnap(folder) = row.snap Then
            FillFolderCacheFromDbRow(fPath, row) : Return row.fsa
        End If

        ' ③ fallback: Layer3 呼叫
        size = Await GetFolderSizeAll(folder)
        If size >= 0 AndAlso Not _cancelRequested Then _cacheFolderSizeAll.TryAdd(fPath, size)
        Return size

    End Function
    Private Async Function GetCachedAttachMailList(folder As Outlook.Folder, progress As IProgress(Of ProgressReport)) As Task(Of List(Of MailItemInfo))
        ' ---------------------------------------------------------------
        ' GetCachedAttachMailList — Tab3 Phase1：含附件的候選郵件清單
        ' by Gemini, 2026/04/05: Layer2.5 快取代理層 - Tab3 Phase 1 快取 - 取得單一資料夾本層含附件的郵件清單
        ' 2026/4/7 by Claude, 加入Database lazy load, 讀資料順序:
        '   ① 記憶體命中
        '   ② DB lazy load：命中且 mca 有效且 snapshot 吻合 → 一次填滿所有欄位
        '   ③ fallback: Layer3 呼叫
        ' ---------------------------------------------------------------
        If _iLikeNoisy Then _dbg(" ├ 開始", folder.Name)
        Dim key As String = folder.FolderPath
        Dim currentCount As Integer = GetCachedMailCount(folder)  ' 依賴同層快取（本身已有 DB lazy load）

        ' ① 記憶體命中
        Dim entry As FolderCacheTab3 = Nothing ' 補上初始化以消除 BC42108 警告
        If _cacheAttachMailList.TryGetValue(key, entry) AndAlso entry.ItemCountSnap = currentCount Then Return entry.AttachMailList

        ' ② DB lazy load (attach_maillist)：item_count_snap == currentCount → 快取仍有效
        Dim dbResult = DbGetAttachMailList(key)
        If dbResult IsNot Nothing AndAlso dbResult.Snap = currentCount Then
            Dim cached As New FolderCacheTab3 With {.AttachMailList = dbResult.Mails, .ItemCountSnap = currentCount}
            _cacheAttachMailList(key) = cached   ' 覆蓋式寫入，確保 ItemCountSnap 對應正確
            If _iLikeNoisy Then _dbg(" ├ DB 命中", $"{folder.Name} ({dbResult.Mails.Count} 封)")
            Return dbResult.Mails
        End If

        ' ③ fallback: Layer3 呼叫
        Dim targetMailList As List(Of MailItemInfo) = Await GetAttachMailList(folder, progress)
        _cacheAttachMailList(key) = New FolderCacheTab3 With {.AttachMailList = targetMailList, .ItemCountSnap = currentCount}
        ' 2026/04/05: 不使用 TryAdd/TryUpdate，確保最後的 cache entry 是正確的 (ItemCountSnap 與 mail list 對應)
        If _iLikeNoisy Then _dbg(" ├ 結束", folder.Name)
        Return targetMailList

    End Function
    Private Function GetCachedAttachFilename(mail As MailItemInfo) As List(Of String)
        ' ---------------------------------------------------------------
        ' GetCachedAttachFilename — Tab3 Phase2：附件檔名清單 (by EntryID)
        ' by Gemini, 2026/04/04: Layer2.5 快取代理層 - 取得附件檔名清單 (含 _cacheAttachFilename 機制)
        ' 2026/4/7 by Claude, 加入Database lazy load, 讀資料順序:
        '   ① 記憶體命中
        '   ② DB lazy load：命中且 fc 有效且 snapshot 吻合 → 一次填滿所有欄位
        '   ③ fallback: Layer3 呼叫
        ' ---------------------------------------------------------------
        Dim result As List(Of String) = Nothing
        If _cacheAttachFilename.TryGetValue(mail.EntryID, result) Then Return result  ' ①

        ' ② DB lazy load (attach_filenames)
        result = DbGetAttachFilenames(mail.EntryID)
        If result IsNot Nothing Then
            _cacheAttachFilename.TryAdd(mail.EntryID, result)
            Return result
        End If

        ' ③ fallback: Layer3 呼叫
        result = GetAttachFilename(mail)
        If result IsNot Nothing Then _cacheAttachFilename.TryAdd(mail.EntryID, result)
        Return result
    End Function
    Private Async Function PreloadAttachByRDOAsync1(sourceList As List(Of MailItemInfo), progress As IProgress(Of ProgressReport), cToken As CancellationToken) As Task
        ' by Gemini, 2026/04/05: Layer2.5 快取代理層 - 批次預熱附件檔名快取
        ' 利用 Redemption (RDO) Free-Threaded 安全的特性，在進入 Layer2 迴圈前平行提早把附件檔名讀進 _cacheAttachFilename。
        ' 完全不動原來的迴圈運作邏輯，以防呆的姿態大幅壓縮等待時間。
        If _rdo Is Nothing OrElse sourceList.Count = 0 Then Return

        _dbg("開始", $"RDO平行預載 {sourceList.Count} 筆")
        Dim swThrottle As New Stopwatch() : swThrottle.Start()
        Dim swTotal As New Stopwatch() : swTotal.Start()
        Dim processed As Integer = 0
        Dim total As Integer = sourceList.Count

        ' 設定並發數：嘗試設為 CPU 核心數的 4 倍，壓榨 SSD 的 Queue Depth
        Dim maxConcurrency As Integer = Environment.ProcessorCount * 4

        Await Task.Run(Sub()
                           ' ✅ 2026/04/11 cToken 重構: CancellationToken 傳入 ParallelOptions，取消時 Parallel.ForEach 會拋 OperationCanceledException
                           Dim parallelOptions As New ParallelOptions With {.MaxDegreeOfParallelism = maxConcurrency, .CancellationToken = cToken}
                           Try
                               Parallel.ForEach(sourceList, parallelOptions,
                                                Sub(mail)
                                                    If Not _cacheAttachFilename.ContainsKey(mail.EntryID) Then
                                                        Dim rdoMsg As Redemption.RDOMail = Nothing
                                                        Try
                                                            rdoMsg = TryCast(_rdo.GetMessageFromID(mail.EntryID), Redemption.RDOMail)
                                                            If rdoMsg IsNot Nothing Then
                                                                Dim list As New List(Of String)()
                                                                ' COM 的 index 從 1 開始
                                                                For i As Integer = 1 To rdoMsg.Attachments.Count
                                                                    list.Add(rdoMsg.Attachments.Item(i).FileName)
                                                                Next
                                                                _cacheAttachFilename.TryAdd(mail.EntryID, list)
                                                            End If
                                                        Catch
                                                        Finally
                                                            If rdoMsg IsNot Nothing Then TryMarshalRelease(rdoMsg)
                                                        End Try
                                                    End If

                                                    Dim currentProcessed As Integer = Interlocked.Increment(processed)
                                                    If swThrottle.ElapsedMilliseconds >= 100 OrElse currentProcessed = total Then
                                                        Dim elapsedSec As Double = Math.Max(swTotal.Elapsed.TotalSeconds, 0.001)
                                                        Dim speed As Double = currentProcessed / elapsedSec
                                                        Dim etaString As String = ""
                                                        If total > 1000 AndAlso speed > 0 Then
                                                            Dim remainingSec As Integer = CInt(Math.Max(0, (total - currentProcessed) / speed))
                                                            If remainingSec > 3 Then etaString = $"，預估剩餘 {remainingSec \ 60:D2}:{remainingSec Mod 60:D2}"
                                                        End If
                                                        progress?.Report(New ProgressReport With {.CurrentCount = currentProcessed,
                                                                                                  .TotalCount = total,
                                                                                                  .Message = $"Phase 2 (RDO 預載快取): {currentProcessed} / {total} ({speed:F0} 封/秒{etaString})"})
                                                        swThrottle.Restart()
                                                    End If
                                                End Sub)
                           Catch ex As OperationCanceledException
                               ' cToken 取消時 Parallel.ForEach 拋出，正常中斷，不需處理
                           End Try
                       End Sub, cToken)
        _dbg(" ├ 結束", $"RDO 預載完成，處理共 {processed} 筆") ' by Gemini, 2026/04/10: 調整縮排層級為 Level 1
    End Function
    Private Async Function PreloadAttachByRDOAsync2(sourceList As List(Of MailItemInfo), progress As IProgress(Of ProgressReport), cToken As CancellationToken) As Task
        ' by AntiGravity, 2026/04/07: 實驗性質 - 使用 Task.WhenAll + SemaphoreSlim，試圖推高 SSD I/O 並發度
        If _rdo Is Nothing OrElse sourceList.Count = 0 Then Return

        _dbg(" ├ 開始", $"WhenAll平行預載 {sourceList.Count} 筆") ' by Gemini, 2026/04/10: 調整縮排層級為 Level 1
        Dim swThrottle As New Stopwatch() : swThrottle.Start()
        Dim swTotal As New Stopwatch() : swTotal.Start()
        Dim processed As Integer = 0
        Dim total As Integer = sourceList.Count

        ' 設定並發數：嘗試設為 CPU 核心數的 4 倍，壓榨 SSD 的 Queue Depth
        Dim maxConcurrency As Integer = Environment.ProcessorCount * 4
        Dim throttler As New SemaphoreSlim(maxConcurrency)
        Dim tasks As New List(Of Task)()

        For Each m As MailItemInfo In sourceList
            If cToken.IsCancellationRequested Then Exit For  ' ✅ 2026/04/11 cToken 重構（原 _cancelRequested 版）
            Dim mail = m ' 在 lambda 中避免變數捕獲問題

            tasks.Add(Task.Run(Async Function()
                                   Await throttler.WaitAsync(cToken)   ' ✅ cToken 取消時直接拋 OperationCanceledException
                                   Try
                                       If Not _cacheAttachFilename.ContainsKey(mail.EntryID) Then
                                           Dim rdoMsg As Redemption.RDOMail = Nothing
                                           Try
                                               rdoMsg = TryCast(_rdo.GetMessageFromID(mail.EntryID), Redemption.RDOMail)
                                               If rdoMsg IsNot Nothing Then
                                                   Dim list As New List(Of String)()
                                                   For i As Integer = 1 To rdoMsg.Attachments.Count
                                                       list.Add(rdoMsg.Attachments.Item(i).FileName)
                                                   Next
                                                   _cacheAttachFilename.TryAdd(mail.EntryID, list)
                                               End If
                                           Catch
                                           Finally
                                               If rdoMsg IsNot Nothing Then TryMarshalRelease(rdoMsg)
                                           End Try
                                       End If

                                       Dim currentProcessed As Integer = Interlocked.Increment(processed)
                                       If swThrottle.ElapsedMilliseconds >= 100 OrElse currentProcessed = total Then
                                           Dim elapsedSec As Double = Math.Max(swTotal.Elapsed.TotalSeconds, 0.001)
                                           Dim speed As Double = currentProcessed / elapsedSec
                                           Dim etaString As String = ""
                                           If total > 1000 AndAlso speed > 0 Then
                                               Dim remainingSec As Integer = CInt(Math.Max(0, (total - currentProcessed) / speed))
                                               If remainingSec > 3 Then etaString = $"，預估剩餘 {remainingSec \ 60:D2}:{remainingSec Mod 60:D2}"
                                           End If
                                           progress?.Report(New ProgressReport With {.CurrentCount = currentProcessed,
                                                                                     .TotalCount = total,
                                                                                     .Message = $"Phase 2 (WhenAll 預載): {currentProcessed} / {total} ({speed:F0} 封/秒{etaString})"})
                                           swThrottle.Restart()
                                           If cToken.IsCancellationRequested Then Return ' 進度回報後立即再檢查一次，盡快停止處理
                                       End If
                                   Finally
                                       throttler.Release()
                                   End Try
                               End Function, cToken))
        Next

        If tasks.Count > 0 Then Await Task.WhenAll(tasks)
        _dbg(" ├ 結束", $"WhenAll 預載完成，處理共 {processed} 筆") ' by Gemini, 2026/04/10
    End Function

#End Region
#Region "  ├ Layer3 直接存取底層計數函數"
    Private Function GetMailSize(item As Object) As Long
        ' --------------------------------------------------------------
        ' GetMailSize: 讀取單封郵件的大小 (bytes)，供 GetFolderSize fallback 路徑呼叫
        '
        ' Fallback 鏈:
        '   ⓪ Redemption : RDOMail.Size
        '                  free-threaded 安全，可在 Task.Run 內呼叫
        '                  繞過 Outlook Security Guard，不會彈出安全性警告
        '                  _rdo 未就緒時自動跳過此層
        '   ① MAPI : PR_MESSAGE_SIZE_EXTENDED (0x0E080014, PT_I8, 64-bit Long)
        '            避免 PR_MESSAGE_SIZE (PT_LONG, 32-bit) 在超大郵件時溢位
        '   ② MAPI : PR_MESSAGE_SIZE (0x0E080003, PT_LONG, 32-bit Integer)
        '            Fallback 到 32-bit 版本，CInt → CLng 安全轉型
        '   ③ OOM  : mail.Size
        '            最後手段，OOM 的 Size 屬性單位是 bytes，回傳 Integer，
        '            大郵件 (>2GB) 理論上會溢位，但實務上 Outlook 的 PST 限制在 50GB 總量，
        '            單封郵件超過 2GB 極不可能，此層可視為安全
        '
        ' 注意: 此函數接受 Object 型別參數，是因為 GetFolderSize 的 fallback 路徑
        '       用 Items.GetFirst/GetNext 取回的是 Object，省去呼叫端的 TryCast 成本
        '       若是 MailItem 就正常讀取，若是其他型別 (Contact、Appointment 等) 就回 0
        '
        ' 取代: GetFolderSizeOld 內的 mailItem.Size 直接呼叫 行3385 的同名 stub (完整替換)
        ' --------------------------------------------------------------

        ' 非 MailItem 的項目 (Calendar、Contact 等) 直接略過，回 0
        Dim mail As Outlook.MailItem = TryCast(item, Outlook.MailItem)
        If mail Is Nothing Then Return 0

        ' ⓪ Redemption: RDOMail.Size
        '   GetMessageFromID 的 StoreID 從 mail.Parent 取得，多一次 COM call 但避免跨 PST 找錯 item
        '   2026-03-22 新增
        If _rdo IsNot Nothing Then
            Dim rdoMail As Redemption.RDOMail = Nothing
            Try
                Dim parentFolder As Outlook.Folder = TryCast(mail.Parent, Outlook.Folder)
                Dim storeId As String = If(parentFolder?.StoreID, "")
                rdoMail = TryCast(_rdo.GetMessageFromID(mail.EntryID, storeId), Redemption.RDOMail)
                If rdoMail IsNot Nothing Then
                    Dim sz As Long = CLng(rdoMail.Size)
                    ' _dbg("GetMailSize ⓪ RDO 成功", $"size={sz}") ' 高頻率項目平時不輸出 Log
                    Return sz
                End If
            Catch ex As System.Exception
                _dbg("    ├ GetMailSize ⓪ RDO 失敗，走MAPI fallback", ex.Message) ' by Gemini, 2026/04/10: 調整縮排層級為 Level 2 (內部失敗路徑)
            Finally
                TryMarshalRelease(rdoMail)
            End Try
        End If

        ' ① MAPI: PR_MESSAGE_SIZE_EXTENDED (0x0E080014, PT_I8) — 64-bit，無溢位風險
        Try
            Const PR_SIZE_EXTENDED As String = "http://schemas.microsoft.com/mapi/proptag/0x0E080014"
            ' by Gemini, 2026/03/29: 移除 TypeOf 判斷，CLng() 可自動處理 Long/Integer 轉型，若屬性不存在或回傳 Nothing/DBNull，CLng 會拋例外進入 Catch
            Return CLng(mail.PropertyAccessor.GetProperty(PR_SIZE_EXTENDED))
        Catch ex As System.Exception
            _dbg("    ├ GetMailSize ① PR_MESSAGE_SIZE_EXTENDED失敗", ex.Message) ' by Gemini, 2026/04/10
        End Try

        ' ② MAPI: PR_MESSAGE_SIZE (0x0E080003, PT_LONG) — 32-bit，超大郵件理論上溢位
        Try
            Const PR_SIZE As String = "http://schemas.microsoft.com/mapi/proptag/0x0E080003"
            ' by Gemini, 2026/03/29: 同上，移除 TypeOf 判斷
            Return CLng(mail.PropertyAccessor.GetProperty(PR_SIZE))
        Catch ex As System.Exception
            _dbg("    ├ GetMailSize ② PR_MESSAGE_SIZE失敗", ex.Message) ' by Gemini, 2026/04/10
        End Try

        ' ③ OOM: mail.Size (Integer，超大郵件理論上不準，但實務上 PST 內不會發生)
        Try
            Return CLng(mail.Size)
        Catch ex As System.Exception
            _dbg("    ├ GetMailSize ③ OOM mail.Size也失敗", ex.Message) ' by Gemini, 2026/04/10
        End Try
        Return -1

    End Function
    Private Function GetMailCount(folder As Outlook.Folder) As Integer
        ' --------------------------------------------------------------
        ' GetMailCount: 只讀單一資料夾的本層郵件數 (不含子孫)
        ' Fallback 鏈:
        '   ⓪ Redemption : RDOFolder.Items.Count (可在非 STA 執行緒呼叫)
        '   ① MAPI : PR_CONTENT_COUNT (0x36020003) (最快快取屬性)
        '   ② OOM  : folder.Items.Count (會建立 Items 集合)
        '   ③ fail : Return -1
        ' --------------------------------------------------------------
        If _iLikeNoisy Then _dbg("開始", folder.Name)
        Dim sw As New Stopwatch() : sw.Start()

        ' ⓪ Redemption: RDOFolder.Items.Count
        If _rdo IsNot Nothing Then
            Dim rdoFolder As Redemption.RDOFolder = Nothing
            Try
                rdoFolder = _rdo.GetFolderFromID(folder.EntryID, folder.StoreID)
                Dim count As Integer = rdoFolder.Items.Count : Return count
            Catch ex As System.Exception
                _dbg("    ├ 錯誤路徑", $"GetMailCount ⓪ RDO: {folder.Name} | {ex.Message}") ' by Gemini, 2026/04/10
            Finally
                TryMarshalRelease(rdoFolder)
            End Try
        End If

        ' ① MAPI: PR_CONTENT_COUNT (0x36020003)
        Try
            Const PR_CONTENT_COUNT As String = "http://schemas.microsoft.com/mapi/proptag/0x36020003"
            Dim count As Integer = CInt(folder.PropertyAccessor.GetProperty(PR_CONTENT_COUNT))
            Return count
        Catch ex As System.Exception
            _dbg("    ├ 錯誤路徑", $"GetMailCount ① MAPI: {folder.Name} | {ex.Message}") ' by Gemini, 2026/04/10
        End Try

        ' ② OOM: folder.Items.Count
        Try
            Dim items As Outlook.Items = Nothing
            Try
                items = folder.Items
                Dim count As Integer = items.Count : Return count
            Finally
                TryMarshalRelease(items)
            End Try
        Catch ex As System.Exception
            _dbg("    ├ 錯誤路徑", $"GetMailCount ② OOM: {folder.Name} | {ex.Message}") ' by Gemini, 2026/04/10
        End Try

        sw.Stop()
        _dbg("結束", $"FAIL: {folder.Name} | -1 | {sw.ElapsedMilliseconds}ms")
        Return -1

    End Function
    Private Async Function GetMailCountAll(rootFolder As Outlook.Folder, Optional progress As IProgress(Of ProgressReport) = Nothing) As Task(Of Long)
        ' --------------------------------------------------------------
        ' GetMailCountAll v3.5: 讀取某資料夾及其整棵子樹的郵件總數
        ' by Gemini, 2026/04/02: 升級為 IProgress(Of ProgressReport) 並加入 100ms 節流回報
        '
        ' v3.0 變更說明 (2026-03-22):
        '   合併原 GetMailCountAll + GetMailCountAllParallel 為單一函數，
        '   統一 fallback 鏈，呼叫端不再需要選擇要用哪個版本。
        '   GetMailCountAllParallel 可標記廢棄或直接刪除。
        '
        ' 設計說明:
        '   為何呼叫 GetMailCount() 而非直接用 GetTable():
        '     PR_CONTENT_COUNT 是 Folder 物件上的已儲存屬性，Outlook 自動維護，讀取等於讀一個整數，一次 COM call 結束。
        '     GetTable() 會把資料夾內所有郵件 row 逐一回傳，只為了計數代價太高。GetTable 適合讀郵件內容 (大小、日期)，不適合純計數。
        '
        '   回傳型別 Long 而非 Integer:
        '     單一資料夾用 Integer 夠 (PR_CONTENT_COUNT 是 PT_LONG 32-bit)，
        '     但整棵子樹加總若有多個大資料夾，理論上可能超過 Integer.MaxValue (2,147,483,647)，用 Long 安全。
        '
        ' Fallback 鏈 (依速度由快到慢) :
        '   ⓪ Redemption : rdoFolder.TotalItemCount
        '                   MAPI 快取的彙總屬性，一次 COM call 直接取得整棵子樹總數，完全不需 BFS 遍歷
        '                   Redemption 可正確讀取 PST 上此屬性 (原生 OOM 的 PR_MESSAGE_SIZE_EXTENDED 在 PST 上無效)
        '                   _rdoSession 未就緒時自動跳過此層
        '                   注意: 走此路徑時 onProgress callback 不會被觸發 (無中間進度可回報)
        '   ① Task.WhenAll 平行 BFS:
        '                   BFS 展開後每個資料夾各建一個 Task.Run，全部 WhenAll 等待
        '                   Task.Run 內的 GetMailCount(f) 走 Redemption ⓪ 時是 free-threaded 安全的
        '                   若 GetMailCount fallback 到 MAPI PropertyAccessor，仍有 STA 違規風險，需留意
        '   ② BFS 循序累加:
        '                   GetSubtreeToList BFS 展開 + GetMailCount(Layer3) 逐一加總
        '                   支援取消檢查和 onProgress 進度回報
        '                   平行路徑失敗時的安全 fallback
        '   ③ 遞迴 fallback:
        '                   GetSubtreeToList 本身失敗時 (極少見) 的最後保險
        '                   無法精確回報進度，但確保加總結果正確
        '   ④ Return -1: 四層都失敗，由 Layer2 決定如何處理
        '
        ' cancelRequested:
        '   檢查 _cancelRequested 旗標，取消時回傳 -1，由 Layer1 判斷是否需要清空 UI
        '   ⓪ Redemption 路徑不插入取消檢查 (單次 call，幾乎瞬間完成)
        '
        ' onProgress 參數 (可選):
        '   傳入 Action(Of Integer, Integer) callback
        '   Layer2 每處理一個資料夾回報 (已完成數, 總數)，讓 Layer1 更新狀態列
        '   不需要進度回報時傳 Nothing
        '   ⓪ 和 ① 路徑不觸發 onProgress，② 路徑才會逐一回報
        '
        ' 取代:
        '   GetMailCountByMAPINew 的整棵子樹加總用途
        '   GetMailCountAllParallel (v3.0 已合併，舊版可廢棄)
        ' --------------------------------------------------------------
        _dbg(" ├ 開始", rootFolder.Name) ' by Gemini, 2026/04/10: Level 1

        ' ⓪ Redemption: TotalItemCount 直接回傳整棵子樹郵件總數
        '   一次 COM call 結束，不需要任何 BFS 遍歷或平行處理
        '   2026-03-22 新增
        If _rdo IsNot Nothing Then
            Dim rdoFolder As Redemption.RDOFolder = Nothing
            Try
                rdoFolder = _rdo.GetFolderFromID(rootFolder.EntryID, rootFolder.StoreID)
                Dim total As Long = CLng(rdoFolder.TotalItemCount)
                _dbg(" ├ 結束", $"⓪ RDO 成功: {rootFolder.Name} | TotalItemCount={total}")
                Return total
            Catch ex As System.Exception
                _dbg(" ├ GetMailCountAll ⓪ RDO 失敗，走平行BFS fallback", $"{rootFolder.Name} | {ex.Message}") ' by Gemini, 2026/04/10: 調整縮排層級為 Level 1
            Finally
                TryMarshalRelease(rdoFolder)
            End Try
        End If

        ' 2026/3/24 by Gemini: ① 平行 BFS (RDO)
        '   使用 GetSubtreeToList_RDO 取得清單，以 Parallel.ForEach 搭配 Interlocked.Add 快速加總
        '   Redemption (RDO) 是 free-threaded，在背景平行執行安全且極為高效
        If _rdo IsNot Nothing Then
            Dim rdoRoot As Redemption.RDOFolder = Nothing
            Try
                rdoRoot = _rdo.GetFolderFromID(rootFolder.EntryID, rootFolder.StoreID)
                Dim rdoFolderList As List(Of Redemption.RDOFolder) = GetSubtreeToList_RDO(rdoRoot, includeSubF:=True)
                Dim targetFolderCount As Integer = rdoFolderList.Count
                Dim totalCount As Long = 0
                Dim processedCount As Integer = 0
                Parallel.ForEach(rdoFolderList,
                    Sub(rdoF As Redemption.RDOFolder)
                        If _cancelRequested Then Return
                        Try
                            Dim count As Integer = rdoF.Items.Count
                            Interlocked.Add(totalCount, CLng(count))
                        Catch ex As System.Exception
                            _dbg("    ├ GetMailCountAll ① 略過失敗資料夾", rdoF.Name) ' by Gemini, 2026/04/10: 調整縮排層級為 Level 2 (並行處理內部)
                        End Try
                        Dim done As Integer = Interlocked.Increment(processedCount)

                        ' by Gemini, 2026/04/02: 更新為 IProgress 且加上簡易模數節流避免平行洗板
                        If done Mod 10 = 0 Then
                            progress?.Report(New ProgressReport With {.CurrentCount = done, .TotalCount = targetFolderCount,
                                                                      .Message = $"正在平行統計: {done} / {targetFolderCount} 個資料夾..."})
                        End If
                    End Sub)
                If _cancelRequested Then
                    _dbg(" ├ GetMailCountAll ① 已取消", $"總資料夾數: {targetFolderCount}") : Return -1
                End If
                _dbg(" ├ 結束", $"① 平行BFS成功 (RDO): {rootFolder.Name} | total={totalCount} | folders={targetFolderCount}")
                Return totalCount
            Catch ex As System.Exception
                _dbg(" ├ GetMailCountAll ① 平行BFS失敗，走循序BFS fallback", $"{rootFolder.Name} | {ex.Message}") ' by Gemini, 2026/04/10
            Finally
                TryMarshalRelease(rdoRoot)
            End Try
        End If

        ' ② BFS 循序累加: GetSubtreeToList 展開 + GetMailCount(Layer3) 逐一加總
        '   支援取消檢查和 progress 進度回報，比平行版保守但穩定
        Try
            Dim targetFolderList As List(Of Outlook.Folder) = Await GetSubtreeToList(rootFolder, includeSubF:=True)
            Dim grandTotal As Long = 0
            Dim swThrottle As New Stopwatch() : swThrottle.Start() ' by Gemini, 2026/04/02: 100ms 節流閥

            For i As Integer = 0 To targetFolderList.Count - 1
                If _cancelRequested Then
                    _dbg(" ├ GetMailCountAll ② 被取消", $"已處理 {i}/{targetFolderList.Count}") : Return -1
                End If
                Dim f As Outlook.Folder = targetFolderList(i)
                Dim count As Integer = GetMailCount(f)
                ' GetMailCount 的所有 fallback 都失敗才會到這個 else，記錄但不中止整體加總
                If count >= 0 Then grandTotal += CLng(count) Else _dbg("    ├ GetMailCountAll ② 略過失敗資料夾", f.Name) ' by Gemini, 2026/04/10

                ' by Gemini, 2026/04/02: 100ms 節流回報進度，且在此區塊不輸出 _dbg()
                If swThrottle.ElapsedMilliseconds >= 100 Then
                    progress?.Report(New ProgressReport With {.CurrentCount = i + 1, .TotalCount = targetFolderList.Count,
                                                              .Message = $"正在統計郵件數: {i + 1} / {targetFolderList.Count} 個資料夾..."})
                    swThrottle.Restart()
                End If

                If i Mod 10 = 0 Then Await Task.Yield()
            Next
            _dbg(" ├ GetMailCountAll ② 循序BFS成功", $"{rootFolder.Name} | total={grandTotal}")
            Return grandTotal
        Catch ex As System.Exception
            _dbg(" ├ GetMailCountAll ② 循序BFS失敗，走遞迴fallback", $"{rootFolder.Name} | {ex.Message}") ' by Gemini, 2026/04/10
        End Try

        ' ③ 遞迴 fallback: GetSubtreeToList 本身失敗時的最後保險
        '   無法精確回報進度，但確保加總結果正確
        '   注意: 遞迴呼叫會重新進入本函數，⓪ Redemption 已失敗所以 _rdoSession 仍 Nothing 或故障
        '         ① ② 也已失敗，只會走到 ③ 再次遞迴——理論上 ③ 不會無限展開，因為每層只遞迴直屬子資料夾
        '        若 ③ 常被觸發，需回頭檢查 GetSubtreeToList 失敗的根本原因 ' pending:
        Try
            Dim totalCount As Long = 0
            Dim count As Integer = GetMailCount(rootFolder)     ' 本層 mailcount
            If count >= 0 Then totalCount += count
            Await Task.Yield()
            For Each f As Outlook.Folder In rootFolder.Folders
                Dim subCount As Long = Await GetMailCountAll(f) ' 遞迴，每個直屬子資料夾各自展開
                If subCount >= 0 Then totalCount += subCount
            Next
            _dbg(" ├ 結束", $"③ 遞迴fallback成功: {rootFolder.Name} | total={totalCount}")
            Return totalCount
        Catch ex As System.Exception
            _dbg(" ├ GetMailCountAll ③ 遞迴fallback也失敗", $"{rootFolder.Name} | {ex.Message}")
            Return -1   ' ④ 四層都失敗，回傳 -1 讓 Layer2 知道這是「讀取失敗」而非「真的是 0 封」
        End Try

    End Function
    Private Function GetFolderCount(folder As Outlook.Folder) As Integer
        ' --------------------------------------------------------------
        ' GetFolderCount: 讀取單一資料夾的本層直屬子資料夾數
        '
        ' Fallback 鏈:
        '   ⓪ Redemption : RDOFolder.Folders.Count
        '            可從非 STA 執行緒呼叫，繞過 Outlook Security Guard
        '            _rdoSession 未就緒時自動跳過此層
        '   ① MAPI : PR_FOLDER_CHILD_COUNT (0x66380003, PT_LONG) 一次 PropertyAccessor call，在大多數情況下準確
        '            注意: PST 上此屬性在剛移動資料夾後可能短暫不同步，但 Outlook 關閉再開就會修正，日常使用可接受
        '            2026/3/20 實測: PR_FOLDER_CHILD_COUNT 沒有一次成功過，已暫時 comment 出
        '   ② OOM  : folder.Folders.Count
        '            Folders 集合比 Items 輕量，載入速度可接受，且永遠準確
        '   ③ fail : Return -1
        '
        ' 關於「先讀 PR_SUBFOLDERS (0x360A000B) 再讀個數」的設計討論:
        '   PR_SUBFOLDERS 是 PT_BOOLEAN，只告訴你有沒有子資料夾 (不告訴你幾個)
        '   先讀它再讀 PR_FOLDER_CHILD_COUNT 等於多一次 COM call，只有「大多數資料夾都沒有子資料夾」時才划算，
        '   實際 PST 不符合此條件，因此直接讀 PR_FOLDER_CHILD_COUNT，不做 PR_SUBFOLDERS 前置判斷
        '
        ' 取代: 散落各處的 folder.Folders.Count 直接呼叫 (建議逐一替換)
        ' --------------------------------------------------------------
        _dbg(" ├ 開始", folder.Name) ' by Gemini, 2026/04/10: Level 1

        ' ⓪ Redemption: RDOFolder.Folders.Count
        '   與 OOM folder.Folders.Count 等價，但可在任意執行緒呼叫, 2026-03-22 新增
        If _rdo IsNot Nothing Then
            Dim rdoFolder As Redemption.RDOFolder = Nothing
            Try
                rdoFolder = _rdo.GetFolderFromID(folder.EntryID, folder.StoreID)
                Dim count As Integer = rdoFolder.Folders.Count : Return count
            Catch ex As System.Exception
                _dbg("       ├ 錯誤路徑", $"GetFolderCount ⓪ RDO: {folder.Name} | {ex.Message}") ' by Gemini, 2026/04/11: Level 3
            Finally
                TryMarshalRelease(rdoFolder)
            End Try
        End If

        ' ① MAPI: PR_FOLDER_CHILD_COUNT (0x66380003)
        ' 2026/3/20, 奇怪PR_FOLDER_CHILD_COUNT 沒有一次成功過??? 乾脆先拿掉這個try, 省得一直fallback也是浪費開銷
        Try
            Dim count As Integer = folder.Folders.Count : Return count
        Catch ex As System.Exception
            _dbg("       ├ 錯誤路徑", $"GetFolderCount ① OOM: {folder.Name} | {ex.Message}") ' by Gemini, 2026/04/11: Level 3
        End Try

        _dbg(" ├ 結束", $"FAIL: {folder.Name}") ' by Gemini, 2026/04/10
        Return -1

    End Function
    Private Async Function GetFolderCountAll(rootFolder As Outlook.Folder, Optional progress As IProgress(Of ProgressReport) = Nothing) As Task(Of Integer)
        ' --------------------------------------------------------------
        ' GetFolderCountAll v1.5: 讀取某資料夾整棵子樹的資料夾總數 (不含 rootFolder 自身)
        ' by Gemini, 2026/04/02: 增加 IProgress 支援與 100ms 節流回報
        '
        ' 2026/3/24 by Gemini: Fallback 鏈 (由快到慢):
        '   ⓪ Redemption + Parallel.ForEach (最快): RDO 是 free-threaded，平行展開子樹
        '   ① Redemption + BFS 循序累加: RDO 循序，平行失敗時的安全路徑
        '   ② OOM + BFS 循序: 無 Redemption 時，走 OOM COM 循序處理
        '   ③ Return -1: 全部失敗
        '
        ' 取代: GetTotalFolderCountAsync (快取邏輯移至 Layer2 呼叫端)
        '
        ' [Redemption說明] 2026-03-22
        '   此函數計算的是整棵子樹的遞迴總數，Redemption 沒有單一 API 可直接取得遞迴資料夾總數
        '    (rdoFolder.Folders.Count 只回傳直屬子資料夾數，與 OOM 相同) 。
        '   因此此函數本身不需要直接加 Redemption 呼叫。
        '   ① BFS 路徑: GetSubtreeToList 內部走 OOM folder.Folders 展開，展開後直接 .Count，不需 Layer3 讀取。
        '   ② 遞迴 fallback: 內部的 rootFolder.Folders.Count 和 ForEach 走 OOM，
        '      若日後改為呼叫 GetFolderCount(Layer3)，即可自動走 Redemption ⓪ 路徑。
        ' --------------------------------------------------------------
        _dbg("開始", rootFolder.Name)

        ' by Gemini, 2026/04/02: 預跑一次顯示準備中
        progress?.Report(New ProgressReport With {.Message = "正在展開資料夾結構...", .IsIndeterminate = True})

        ' 2026/3/24 by Gemini: ⓪ Redemption + 平行處理 (最快路徑)
        '   使用 GetSubtreeToList_RDO 取得清單，以 Parallel.ForEach 搭配 Interlocked.Add(rdoF.Folders.Count) 快速加總
        If _rdo IsNot Nothing Then
            Dim rdoRoot As Redemption.RDOFolder = Nothing
            Try
                rdoRoot = _rdo.GetFolderFromID(rootFolder.EntryID, rootFolder.StoreID)
                Dim rdoFolderList As List(Of Redemption.RDOFolder) = GetSubtreeToList_RDO(rdoRoot, includeSubF:=True)
                Dim targetFolderCount As Integer = rdoFolderList.Count
                Dim totalCount As Integer = 0
                Dim processedCount As Integer = 0
                Dim swThrottle As New Stopwatch() : swThrottle.Start() ' by Gemini, 2026/04/02
                Parallel.ForEach(rdoFolderList,
                    Sub(rdoF As Redemption.RDOFolder)
                        If _cancelRequested Then Return
                        Try
                            Dim count As Integer = rdoF.Folders.Count
                            Interlocked.Add(totalCount, count)
                        Catch ex As System.Exception
                            _dbg("       ├ GetFolderCountAll ⓪ RDO 略過失敗資料夾", rdoF.Name) ' by Gemini, 2026/04/11: Level 3
                        End Try

                        Dim done As Integer = Interlocked.Increment(processedCount)
                        ' by Gemini, 2026/04/02: 更新為 IProgress 且加上 100ms 節流，取代原有的 Mod 10
                        If swThrottle.ElapsedMilliseconds >= 100 Then
                            progress?.Report(New ProgressReport With {.CurrentCount = done, .TotalCount = targetFolderCount,
                                                                      .Message = $"正在統計資料夾樹: {done} / {targetFolderCount}..."})
                            swThrottle.Restart()
                        End If
                    End Sub)
                If _cancelRequested Then
                    _dbg("GetFolderCountAll ⓪ 已取消", "") : Return -1
                End If
                _dbg(" ├ 結束", $"⓪ RDO平行成功: {rootFolder.Name} | total={totalCount}") ' by Gemini, 2026/04/10
                Return totalCount
            Catch ex As System.Exception
                _dbg(" ├ GetFolderCountAll ⓪ RDO平行失敗，走OOM循序fallback", $"{rootFolder.Name} | {ex.Message}") ' by Gemini, 2026/04/10
            Finally
                TryMarshalRelease(rdoRoot)
            End Try
        End If

        ' 2026/3/24 by Gemini: ② OOM + BFS 循序 (無 Redemption 時的最後手段)
        '   必須循序處理 OOM COM 物件以避免 STA 違規
        Try
            Dim allFolders As List(Of Outlook.Folder) = Await GetSubtreeToList(rootFolder, includeSubF:=True, progress:=progress)
            ' by Gemini, 2026/04/02: BFS 展開後回傳數量
            Dim total = allFolders.Count - 1
            progress?.Report(New ProgressReport With {.CurrentCount = total,
                                                      .TotalCount = total,
                                                      .Message = $"資料夾結構已展開: 共 {total} 個資料夾。"})
            Await Task.Yield()
            _dbg(" ├ 結束", $"② OOM BFS成功: {rootFolder.Name} | total={total}") ' by Gemini, 2026/04/10
            Return total
        Catch ex As System.Exception
            _dbg(" ├ GetFolderCountAll ② OOM BFS失敗", $"{rootFolder.Name} | {ex.Message}") ' by Gemini, 2026/04/10
        End Try
        ' ③ 全部失敗
        Return -1

    End Function
    Private Async Function GetFolderSize(folder As Outlook.Folder, Optional progress As IProgress(Of ProgressReport) = Nothing) As Task(Of Long)
        ' --------------------------------------------------------------
        ' GetFolderSize v1.5: 讀取單一資料夾本層大小 (bytes)
        ' by Gemini, 2026/04/02: 加入 IProgress 支援以回報分批讀取進度 (100ms 節流)
        ' 2026/3/24 by Gemini: Fallback 鏈重構
        '   ⓪ Redemption : rdoFolder.Fields(PR_MESSAGE_SIZE_EXTENDED) (部分 Exchange 支援，極快)
        '   ① OOM  : folder.GetTable(PR_MESSAGE_SIZE_EXTENDED) + GetArray(1000) (最快安全招式)
        '   ② OOM  : folder.GetTable(PR_MESSAGE_SIZE_EXTENDED) + GetNextRow() (備案)
        '   ③ fail : Return -1
        ' --------------------------------------------------------------
        If _iLikeNoisy Then _dbg("開始", folder.Name)
        Dim sw As New Stopwatch() : sw.Start()

        ' ⓪ Redemption 層 (嘗試讀取資料夾本身的總量屬性)
        ' RDO 沒有 GetTable().GetArray()，故若屬性讀不到直接 fallback
        If _rdo IsNot Nothing Then
            Dim rdoFolder As Redemption.RDOFolder = Nothing
            Try
                rdoFolder = _rdo.GetFolderFromID(folder.EntryID, folder.StoreID)
                ' PR_MESSAGE_SIZE_EXTENDED (0x0E080014)
                Const PR_SIZE_EX As Integer = &HE080014
                Dim val As Object = rdoFolder.Fields(PR_SIZE_EX)
                If val IsNot Nothing AndAlso Not IsDBNull(val) Then
                    Dim totalSize As Long = CLng(val) : sw.Stop()
                    _dbg(" ├ 結束", $"⓪ RDO Fields 成功: {folder.Name} | size={totalSize} | {sw.ElapsedMilliseconds}ms") ' by Gemini, 2026/04/10
                    Return totalSize
                End If
            Catch ex As System.Exception
                _dbg("       ├ 錯誤: ⓪ RDO 失敗，走 OOM GetArray fallback", $"{folder.Name} | {ex.Message}") ' by Gemini, 2026/04/11: Level 3
            Finally
                TryMarshalRelease(rdoFolder)
            End Try
        End If

        ' ① OOM GetTable + GetArray(1000) (目前最穩、最快的批次讀取)
        Const PR_SIZE_EX_STR As String = "http://schemas.microsoft.com/mapi/proptag/0x0E080014"
        Dim table As Outlook.Table = Nothing
        Try
            table = folder.GetTable()
            table.Columns.RemoveAll()
            table.Columns.Add(PR_SIZE_EX_STR)
            Dim totalSize As Long = 0
            Dim swThrottle As New Stopwatch() : swThrottle.Start() ' by Gemini, 2026/04/02

            Do While Not table.EndOfTable
                Dim arr As Object = table.GetArray(1000)
                If arr Is Nothing Then Exit Do
                Dim data(,) As Object = DirectCast(arr, Object(,))
                For r As Integer = 0 To data.GetUpperBound(0)
                    Dim sz = data(r, 0)
                    If sz IsNot Nothing AndAlso Not IsDBNull(sz) Then totalSize += CLng(sz)
                Next

                ' by Gemini, 2026/04/02: 單一資料夾內部進度回報 (100ms 節流)
                If swThrottle.ElapsedMilliseconds >= 100 Then
                    progress?.Report(New ProgressReport With {.Message = $"正在計算 {folder.Name} 大小: {totalSize / 1024 / 1024:0.0} MB..."})
                    swThrottle.Restart()
                End If
                Await Task.Yield() ' 讓出 UI 避免卡死
            Loop
            sw.Stop()
            If _iLikeNoisy Then _dbg(" ├ 結束", $"① OOM GetArray 成功: {folder.Name} | size={totalSize} | {sw.ElapsedMilliseconds}ms") ' by Gemini, 2026/04/10
            Return totalSize
        Catch ex As System.Exception
            _dbg("       ├ 錯誤: ① OOM GetArray 失敗，走 GetNextRow fallback", $"{folder.Name} | {ex.Message}") ' by Gemini, 2026/04/11: Level 3
        Finally
            TryMarshalRelease(table)
        End Try

        ' ② OOM GetTable + GetNextRow() (不依賴二維陣列的最後保險)
        Dim table2 As Outlook.Table = Nothing
        Try
            table2 = folder.GetTable()
            table2.Columns.RemoveAll()
            table2.Columns.Add(PR_SIZE_EX_STR)
            Dim totalSize As Long = 0
            Dim loopCount As Integer = 0
            Do While Not table2.EndOfTable
                Dim row As Outlook.Row = table2.GetNextRow()
                If row IsNot Nothing Then
                    totalSize += SafeGet(Of Long)(row, PR_SIZE_EX_STR, 0L)
                    TryMarshalRelease(row)
                End If
                loopCount += 1
                If loopCount Mod 500 = 0 Then Await Task.Yield()
            Loop
            sw.Stop()
            _dbg(" ├ 結束", $"② OOM GetNextRow 成功: {folder.Name} | size={totalSize} | {sw.ElapsedMilliseconds}ms") ' by Gemini, 2026/04/10
            Return totalSize
        Catch ex As System.Exception
            _dbg("       ├ 錯誤: ② OOM GetNextRow 失敗", $"{folder.Name} | {ex.Message}") ' by Gemini, 2026/04/11: Level 3
        Finally
            TryMarshalRelease(table2)
        End Try

        sw.Stop()
        _dbg("結束", $"FAIL: {folder.Name} | -1 | {sw.ElapsedMilliseconds}ms")
        Return -1
    End Function
    Private Async Function GetFolderSizeAll(rootFolder As Outlook.Folder, Optional progress As IProgress(Of ProgressReport) = Nothing) As Task(Of Long)
        ' --------------------------------------------------------------
        ' GetFolderSizeAll v1.5: 讀取某資料夾及整棵子樹的大小總計 (bytes)
        ' by Gemini, 2026/04/02: 增加 IProgress 支援與 100ms 節流回報
        '
        ' 2026/3/24 by Gemini: 落實新的 Fallback 鏈設計，並修正平行處理的 STA 問題
        '   ⓪ Redemption 平行路徑 (最快):
        '      利用 GetSubtreeToList_RDO 一次把該子樹下所有 RDOFolder 拿出來，
        '      放到 Parallel.ForEach 中，各別讀取 MAPI 屬性 PR_MESSAGE_SIZE_EXTENDED。
        '      (RDOFolder 不支援 GetTable().GetArray()，故依賴屬性直讀)
        '
        '   ① OOM 循序路徑 (最安全):
        '      當 RDO 平行路徑失敗 (或是未匯入 Redemption) ，退回使用 OOM。
        '      OOM 絕對不可以在 Task.Run / WhenAll 等背景執行緒內呼叫 COM，否則會觸發 STA 錯誤。
        '      故改為嚴格的 For 迴圈，逐一 Await GetFolderSize()。
        '      而內部的 GetFolderSize 會走到它專屬的 GetTable().GetArray(1000) OOM 極速路徑。
        '
        '   ② 兩層都失敗: 回傳 -1，交給上一層流程處理。
        ' --------------------------------------------------------------
        _dbg("開始", rootFolder.Name)
        ' 2026/3/24 by Gemini: ⓪ Redemption 平行累加 PR_MESSAGE_SIZE_EXTENDED
        If _rdo IsNot Nothing Then
            Dim rdoRoot As Redemption.RDOFolder = Nothing
            Try
                rdoRoot = _rdo.GetFolderFromID(rootFolder.EntryID, rootFolder.StoreID)
                Dim rdoFolderList As List(Of Redemption.RDOFolder) = GetSubtreeToList_RDO(rdoRoot, includeSubF:=True)
                Dim grandTotal As Long = 0
                Const PR_SIZE_EX As Integer = &HE080014

                ' 利用 Parallel.ForEach 與 Interlocked.Add 達到極致的多核並發加總
                Dim validCount As Integer = 0
                Parallel.ForEach(rdoFolderList,
                    Sub(rdoF As Redemption.RDOFolder)
                        If _cancelRequested Then Return
                        Try
                            Dim val As Object = rdoF.Fields(PR_SIZE_EX)
                            If val IsNot Nothing AndAlso Not IsDBNull(val) Then
                                Interlocked.Add(grandTotal, CLng(val))
                                Interlocked.Increment(validCount)
                            End If
                        Catch ex As System.Exception
                            _dbg("錯誤: ⓪ RDO 略過讀取失敗的資料夾", rdoF.Name)
                        End Try
                    End Sub)

                If _cancelRequested Then
                    _dbg("錯誤: ⓪ 已取消", $"總資料夾數: {rdoFolderList.Count}") : Return -1
                End If
                If validCount = 0 AndAlso rdoFolderList.Count > 0 Then
                    _dbg("錯誤: ⓪ RDO 讀取失敗 (無支援的屬性) ", "退回 OOM")
                    Throw New System.Exception("RDO PR_SIZE_EX returned empty for all folders")
                End If
                _dbg("結束", $"⓪ RDO平行成功: {rootFolder.Name} | totalSize={grandTotal} | folders={rdoFolderList.Count}")
                Return grandTotal
            Catch ex As System.Exception
                _dbg("錯誤: ⓪ RDO平行失敗，走 OOM 循序 fallback", $"{rootFolder.Name} | {ex.Message}")
            Finally
                TryMarshalRelease(rdoRoot)
            End Try
        End If

        ' 2026/3/24 by Gemini: ① OOM 循序 BFS 累加 (避免 STA 錯誤的保險路徑)
        ' 因為 OOM 的 GetTable() 必須在 UI Thread，我們必須循序 Await 每一層
        Try
            Dim targetFolderList As List(Of Outlook.Folder) = Await GetSubtreeToList(rootFolder, includeSubF:=True)
            Dim grandTotal As Long = 0
            Dim swThrottle As New Stopwatch() : swThrottle.Start() ' by Gemini, 2026/04/02

            For i As Integer = 0 To targetFolderList.Count - 1
                If _cancelRequested Then
                    _dbg("錯誤: ① 被取消", $"已處理 {i}/{targetFolderList.Count}") : Return -1
                End If
                Dim f As Outlook.Folder = targetFolderList(i)
                ' by Gemini, 2026/04/02: 傳遞 progress 進去以獲得更細緻的(郵件級別)進度回報
                Dim sz As Long = Await GetFolderSize(f, progress)

                If sz >= 0 Then
                    grandTotal += sz
                Else
                    _dbg("錯誤: ① 略過了大小計算失敗的資料夾", f.Name)
                End If

                ' by Gemini, 2026/04/02: 100ms 節流回報資料夾級別進度
                If swThrottle.ElapsedMilliseconds >= 100 Then
                    progress?.Report(New ProgressReport With {.CurrentCount = i + 1, .TotalCount = targetFolderList.Count,
                                                              .Message = $"正在計算大小: {i + 1} / {targetFolderList.Count} ({f.Name})..."})
                    swThrottle.Restart()
                End If

                ' 避免卡死 UI
                If i Mod 5 = 0 Then Await Task.Yield()
            Next
            _dbg("結束", $"① 循序BFS成功: {rootFolder.Name} | totalSize={grandTotal}")
            Return grandTotal
        Catch ex As System.Exception
            _dbg("錯誤: ① 循序BFS失敗，放棄計算", $"{rootFolder.Name} | {ex.Message}")
        End Try

        ' ② 兩層都失敗，回傳 -1 讓呼叫端知道失敗了
        Return -1
    End Function
    Private Async Function GetYearCountsForFolder(folder As Outlook.Folder) As Task(Of ConcurrentDictionary(Of Integer, Integer))
        ' ---------------------------------------------------------------
        ' === Layer 3: COM 資料層 ===
        ' 職責: 對 Outlook 發出 COM 呼叫，回傳單一資料夾的年份郵件分佈
        ' 規則: 不遞迴、不碰 UI、不修改任何全域狀態，
        '       只做一件事: 詢問 Outlook 某資料夾每年有幾封郵件，回傳結果
        '       不遞迴、不知道上層的進度計數、不碰 UI，完全純粹的資料查詢函數
        ' 2026/3/24 by Gemini: 從逐年 Restrict 改為 GetTable + GetArray 一次讀完再記憶體分組
        '   原本每年一次 Restrict + Items.Count = ~30 次 COM call
        '   現在 1 次 GetTable + ceil(N/1000) 次 GetArray，大幅減少 COM 跨程序呼叫
        ' todo: 目前最耗時間的function(), 占整體時間60~65%
        ' ---------------------------------------------------------------
        If _iLikeNoisy Then _dbg("開始", folder.Name)

        ' 2026/3/11再次重構: 優化 COM 呼叫，減少 RCW 物件積累，提升效能和穩定性
        'Dim folderItems As Outlook.Items = Nothing
        Dim yearCounts As New ConcurrentDictionary(Of Integer, Integer)
        Const BATCH_SIZE As Integer = 500  ' 2026/3/24 by Gemini: 每次批量讀取的筆數
        Dim table As Outlook.Table = Nothing
        Try
            ' 2026/3/24 by Gemini: 改用 GetTable + GetArray 取代逐年 Restrict
            ' 只讀 ReceivedTime 一欄，最小化每 row 的傳輸量
            table = folder.GetTable()
            table.Columns.RemoveAll()
            table.Columns.Add("ReceivedTime")   ' 欄位索引 0

            ' by Gemini, 2026/04/05: 每批次讀取後，若超過 100ms 則釋放執行緒並檢查中斷
            Dim swThrottle As New Stopwatch() : swThrottle.Start()
            Do While Not table.EndOfTable
                If _cancelRequested Then Exit Do
                Dim arr As Object = table.GetArray(BATCH_SIZE)
                If arr Is Nothing Then Exit Do

                Dim data(,) As Object = DirectCast(arr, Object(,))
                Dim rows As Integer = data.GetUpperBound(0) + 1
                For r As Integer = 0 To rows - 1
                    Dim receivedTime As DateTime = SafeGet(Of DateTime)(data, r, 0, DateTime.MinValue)
                    If receivedTime > DateTime.MinValue Then
                        Dim year As Integer = receivedTime.Year
                        If year > 0 AndAlso year <= Date.Today.Year Then
                            yearCounts.AddOrUpdate(year, 1, Function(k, v) v + 1)
                        End If
                    End If
                Next

                ' by Gemini, 2026/04/05: 每 100ms 節流讓出執行緒
                If swThrottle.ElapsedMilliseconds >= 100 Then
                    swThrottle.Restart()
                    Await Task.Delay(1) ' 這裡一定要保留至少 .delay(1) 才能讓 ESC 中斷生效 (simon, 2026/04/05)
                    If _cancelRequested Then Exit Do
                End If
            Loop
        Catch ex As System.Exception
            _dbg("錯誤", $"{folder.Name}: {ex.Message}") ' by Gemini, 2026/04/04: Issue 4 格式標準化
        Finally
            TryMarshalRelease(table)
        End Try
        Await Task.Yield()   ' ✅ 函數結束前再讓出一次，確保畫面有機會更新

        If _iLikeNoisy Then _dbg("結束", $"{folder.Name} | 年份分佈: {yearCounts.Count}")
        Return yearCounts

    End Function
    Private Async Function GetMonthCountsForYear(folder As Outlook.Folder, year As Integer) As Task(Of ConcurrentDictionary(Of Integer, Integer))
        ' ---------------------------------------------------------------
        ' GetMonthCountsForYear (完整替換舊版，加入快取和進度支援)
        ' Layer3 COM 資料層: 計算單一資料夾在指定年份中每個月的郵件數量
        ' 快取 key = FolderPath + "_" + year，與 yearCountsCache 的命名慣例一致
        ' 2026/3/24 by Gemini: 從逐月 Restrict 改為 GetTable + GetArray 一次讀完再記憶體分組
        '   原本 12 次 Restrict + 12 次 Items.Count = 24 次 COM call
        '   現在 1 次 GetTable (含日期範圍 filter) + ceil(N/1000) 次 GetArray
        ' ---------------------------------------------------------------
        If _iLikeNoisy Then _dbg("開始", $"{folder.Name} ({year} 年)")

        ' ✅ 2026/04/10: 提前過濾 1: 該資料夾總郵件數為 0，直接略過 (by Gemini)
        If GetCachedMailCount(folder) = 0 Then Return New ConcurrentDictionary(Of Integer, Integer)()

        ' ✅ 2026/04/10: 提前過濾 2: 由先前的年度統計直接判定該資料夾今年無信件 (by Gemini)
        '    解決 DB 沒有存 0 封信的紀錄，導致 lazy_load 回傳 Nothing 而被迫打 COM 的問題
        Dim yCache As ConcurrentDictionary(Of Integer, Integer) = Nothing
        If _cacheYearCounts.TryGetValue(folder.FolderPath, yCache) Then
            Dim countInYear As Integer = 0
            yCache.TryGetValue(year, countInYear)
            If countInYear = 0 Then Return New ConcurrentDictionary(Of Integer, Integer)()
        End If

        ' ① 記憶體快取命中: 直接回傳，不打任何 COM
        Dim cacheKey As String = folder.FolderPath & "_" & year.ToString()
        Dim value As ConcurrentDictionary(Of Integer, Integer) = Nothing
        If _cacheMonthCounts.TryGetValue(cacheKey, value) Then Return value

        ' ② DB lazy load（2026/04/09 by Claude）: 記憶體 miss 時先查 SSD，有就填回記憶體快取
        ' 2026/04/09 修正：改用 (folderPath, year) 兩個參數，對應新 PK 設計
        Dim dbResult = DbGetMonthCountsForFolder(folder.FolderPath, year)
        If dbResult IsNot Nothing Then
            _cacheMonthCounts.TryAdd(cacheKey, dbResult)
            If _iLikeNoisy Then _dbg("DB 命中", $"{folder.Name} {year} 年 ({dbResult.Count} 個月)")
            Return dbResult
        End If

        ' ③ L3 COM 呼叫: DB miss 才真正打 COM
        Dim monthCounts As New ConcurrentDictionary(Of Integer, Integer)
        Const BATCH_SIZE As Integer = 500  ' 2026/3/24 by Gemini
        Dim table As Outlook.Table = Nothing
        Try
            ' 2026/3/24 by Gemini: 改用 GetTable + 日期範圍 DASL filter + GetArray
            ' 用整年的日期範圍一次篩選，不再逐月 Restrict
            Dim startDate As New Date(year, 1, 1, 0, 0, 0)
            Dim endDate As New Date(year, 12, 31, 23, 59, 59)
            Dim dateFilter As String = $"[ReceivedTime] >= '{startDate}' AND [ReceivedTime] <= '{endDate}'"
            table = folder.GetTable(dateFilter)
            table.Columns.RemoveAll()
            table.Columns.Add("ReceivedTime")   ' 欄位索引 0

            ' by simon, 2026/04/08: 每批次讀取後，若超過 100ms 則釋放執行緒並檢查中斷
            Dim swThrottle As New Stopwatch() : swThrottle.Start()
            Do While Not table.EndOfTable
                Dim arr As Object = table.GetArray(BATCH_SIZE)
                If arr Is Nothing Then Exit Do
                Dim data(,) As Object = DirectCast(arr, Object(,))
                Dim rows As Integer = data.GetUpperBound(0) + 1
                For r As Integer = 0 To rows - 1
                    Dim receivedTime As DateTime = SafeGet(Of DateTime)(data, r, 0, DateTime.MinValue)
                    If receivedTime > DateTime.MinValue Then
                        Dim month As Integer = receivedTime.Month
                        If _iLikeNoisy Then _dbg("DB miss, 從 L3 COM: ", $"{folder.Name} {year} 年")
                        monthCounts.AddOrUpdate(month, 1, Function(k, v) v + 1)
                    End If
                Next

                ' by simon, 2026/04/08: 每 100ms 節流讓出執行緒
                If swThrottle.ElapsedMilliseconds >= 100 Then
                    swThrottle.Restart()
                    Await Task.Delay(1) ' 這裡一定要保留至少 .delay(1) 才能讓 ESC 中斷生效 (simon, 2026/04/05)
                    If _cancelRequested Then Exit Do ' by Gemini, 2026/04/09: 移出 100ms 節流，確保每一批次 (或 Task.Delay 醒來後) 都能立刻偵測到 ESC
                End If

            Loop
        Catch ex As System.Exception
            _dbg("錯誤", $"{folder.Name}, year={year}: {ex.Message}") ' by Gemini, 2026/04/04: Issue 4 格式標準化
        Finally
            TryMarshalRelease(table)
        End Try
        _cacheMonthCounts(cacheKey) = monthCounts    ' ✅ 第一次統計完, 一律存入快取，下次進入同一年份直接命中

        '' ✅ 2026/04/09 新增：L3 計算完後立刻增量寫入 DB，不等 SaveCache 按鈕批次。
        ''    根本原因：若使用者沒手動 SaveCache，下次重開程式 DB lazy 仍回 Nothing，每次都打 COM。
        ''    修正後：每次 L3 計算完月份就持久化，下次 DB lazy 直接命中，首次以外均免 COM。
        'If _iLikeNoisy Then DbSaveMonthCountsSingle(folder.FolderPath, year, monthCounts)
        If _iLikeNoisy Then _dbg("結束", $"{folder.Name} ({year} 年)")
        Return monthCounts

    End Function
    Private Async Function GetAttachMailList(folder As Outlook.Folder, progress As IProgress(Of ProgressReport)) As Task(Of List(Of MailItemInfo))
        ' Phase 1 / Layer3 純資料層: GetTable + GetArray 批次掃描單一資料夾
        ' 設計: 這裡只專注於透過 MAPI 取回資料，不會做快取判定，也無關大小設定過濾
        If _iLikeNoisy Then _dbg(" ├ 開始", folder.Name)

        Const PR_MESSAGE_SIZE As String = "http://schemas.microsoft.com/mapi/proptag/0x0E080003"
        Const BATCH_SIZE As Integer = 500
        Dim table As Outlook.Table = Nothing

        Dim strFilterHasAttachment As String = "@SQL=" & Chr(34) & "urn:schemas:httpmail:hasattachment" & Chr(34) & " = True"
        Dim result As New List(Of MailItemInfo)
        Try
            table = folder.GetTable(strFilterHasAttachment)
            table.Columns.RemoveAll()
            table.Columns.Add("EntryID")
            table.Columns.Add("Subject")
            table.Columns.Add(PR_MESSAGE_SIZE)
            table.Columns.Add("ReceivedTime")
            table.Columns.Add("SenderName")

            Dim swThrottle As New Stopwatch() : swThrottle.Start()
            Dim rowCount As Integer = 0
            Do While Not table.EndOfTable
                If _cancelRequested Then Exit Do

                Dim arr As Object = table.GetArray(BATCH_SIZE)
                If arr Is Nothing Then Exit Do

                If swThrottle.ElapsedMilliseconds >= 100 Then
                    progress?.Report(New ProgressReport With {.Message = $"Phase 1 掃描: {folder.Name} (已找 {result.Count} 封)"})
                    swThrottle.Restart()
                    Await Task.Delay(1) ' ✅ 只有在更新進度時才讓出執行緒，效能與響應兼具
                End If

                Dim data(,) As Object = DirectCast(arr, Object(,))
                Dim rows As Integer = data.GetUpperBound(0) + 1
                For r As Integer = 0 To rows - 1
                    Dim entryID As String = SafeGet(Of String)(data, r, 0, "")
                    If entryID = "" Then Continue For
                    Dim info As New MailItemInfo With {.EntryID = entryID,
                                                       .Subject = SafeGet(Of String)(data, r, 1, ""),
                                                       .Size = SafeGet(Of Long)(data, r, 2, 0L),
                                                       .ReceivedTime = SafeGet(Of DateTime)(data, r, 3, DateTime.MinValue),
                                                       .SenderName = SafeGet(Of String)(data, r, 4, "")}
                    result.Add(info)
                    rowCount += 1
                Next
                If _cancelRequested Then Exit Do
            Loop
        Catch ex As System.Exception
            _dbg(" ├ 錯誤: ", folder.Name & " — " & ex.Message)
        Finally
            TryMarshalRelease(table)
        End Try
        If _iLikeNoisy Then _dbg(" ├ 結束", $"找到 {result.Count} 封有附件郵件")
        Return result
    End Function
    Private Async Function GetFolderBasicMailInfosAsync(folder As Outlook.Folder, needTopic As Boolean, ct As CancellationToken) As Task(Of List(Of (Mail As MailItemInfo, Topic As String)))
        Dim resultList As New List(Of (MailItemInfo, String))
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
            If needTopic Then table.Columns.Add(PR_CONVERSATION_TOPIC)

            Dim swYield As New Stopwatch() : swYield.Start()

            Do While Not table.EndOfTable
                Dim arr As Object = table.GetArray(1000)
                If arr Is Nothing Then Exit Do
                Dim data(,) As Object = DirectCast(arr, Object(,))

                For r As Integer = 0 To data.GetUpperBound(0)
                    Dim entryID As String = SafeGet(Of String)(data, r, 0, "")
                    If entryID = "" Then Continue For

                    Dim info As New MailItemInfo With {
                        .EntryID = entryID,
                        .Subject = SafeGet(Of String)(data, r, 1, ""),
                        .Size = SafeGet(Of Long)(data, r, 2, 0L),
                        .ReceivedTime = SafeGet(Of DateTime)(data, r, 3, DateTime.MinValue),
                        .SenderName = SafeGet(Of String)(data, r, 4, "")
                    }

                    Dim topic As String = If(needTopic, SafeGet(Of String)(data, r, 5, ""), "")
                    resultList.Add((info, topic))
                Next

                ' ✅ 內建標準 200ms 節流讓位與中斷檢查
                If swYield.ElapsedMilliseconds >= 200 Then
                    swYield.Restart()
                    Await Task.Yield()
                    ct.ThrowIfCancellationRequested()
                End If
            Loop
        Catch ex As System.Exception
            _dbg("錯誤", $"{folder.Name}: {ex.Message}")
        Finally
            TryMarshalRelease(table)
        End Try

        Return resultList
    End Function
    Private Function GetAttachFilename(mail As MailItemInfo) As List(Of String)
        ' by Gemini, 2026/04/04: 取得郵件的附件檔名清單 (純 Layer3 邏輯，不做快取)
        If _iLikeNoisy Then _dbg(" ├ 開始", mail.Subject)
        Dim result As New List(Of String)()

        ' ⓪ Redemption 優先: 繞過 OOM 開信的記憶體開銷，直接透過 MAPI Table 抓取檔名
        If _rdo IsNot Nothing Then
            Dim rdoMsg As Redemption.RDOMail = Nothing
            Try
                rdoMsg = TryCast(_rdo.GetMessageFromID(mail.EntryID), Redemption.RDOMail)
                If rdoMsg IsNot Nothing Then
                    For i As Integer = 1 To rdoMsg.Attachments.Count
                        Dim att As Redemption.RDOAttachment = rdoMsg.Attachments.Item(i)
                        Try : If att.Type = 1 Then result.Add(att.FileName)   ' 2026/04/09 by Gemini: 僅處理 olByValue (1)
                        Finally : TryMarshalRelease(att)
                        End Try
                    Next
                End If
                Return result
            Catch ex As System.Exception
                _dbg("⓪ RDO 失敗，走OOM fallback", ex.Message)
            Finally
                TryMarshalRelease(rdoMsg)
            End Try
        End If

        ' ① Fallback: 使用 Outlook OOM (極為昂貴的物件實例化)
        Dim tempMail As Outlook.MailItem = Nothing
        Dim attachments As Outlook.Attachments = Nothing
        Try
            tempMail = TryCast(_olNS.GetItemFromID(mail.EntryID), Outlook.MailItem)
            If tempMail IsNot Nothing Then
                attachments = tempMail.Attachments
                For i As Integer = 1 To attachments.Count ' COM 是 1-based index
                    Dim att As Outlook.Attachment = attachments.Item(i)
                    Try : If att.Type = Outlook.OlAttachmentType.olByValue Then result.Add(att.FileName) ' 2026/04/09 by Gemini: 僅處理 olByValue (1) 類型的附件
                    Finally : TryMarshalRelease(att)
                    End Try
                Next
            End If
        Catch ex As System.Exception
            _dbg("① OOM 失敗", ex.Message)
        Finally
            If _iLikeNoisy Then _dbg(" ├ 結束")
            TryMarshalRelease(attachments)
            TryMarshalRelease(tempMail)
        End Try

        Return result
    End Function
#End Region
#Region "  └ 輔助函數"
    Private Function TextHasChineseChar(name As String) As Boolean
        Return name.Any(Function(c) c >= ChrW(&H4E00) AndAlso c <= ChrW(&H9FFF))
    End Function
    Private Function IsMailFolder(folder As Outlook.Folder) As Boolean
        If _iLikeNoisy Then _dbg(" ├ 開始", folder.Name)
        Dim fPath As String = folder.FolderPath
        Dim isMail As Boolean
        If _cacheIsMailFolder.TryGetValue(fPath, isMail) Then Return isMail

        Static allowedTypes As Outlook.OlItemType() = {Outlook.OlItemType.olMailItem, Outlook.OlItemType.olPostItem}
        Try
            Dim itemType As Outlook.OlItemType = folder.DefaultItemType
            isMail = allowedTypes.Contains(itemType)
            _cacheIsMailFolder.TryAdd(fPath, isMail)
            If Not isMail Then _dbg("過濾非郵件資料夾", $"{folder.Name} ({itemType})") ' 只有非郵件時才記錄
            Return isMail
        Catch
            Return False
        End Try
    End Function

    Private Function objFolder2odoFolder(objFolder As Outlook.Folder) As Redemption.RDOFolder
        If _rdo Is Nothing OrElse objFolder Is Nothing Then Return Nothing
        Return _rdo.GetFolderFromID(objFolder.EntryID, objFolder.StoreID)
    End Function
    Private Function rdoFolder2objFolder(rdoFolder As Redemption.RDOFolder) As Outlook.Folder
        If rdoFolder Is Nothing Then Return Nothing
        Return _olNS.GetFolderFromID(rdoFolder.EntryID, rdoFolder.StoreID)
    End Function
    Private Sub TryMarshalRelease(ByRef obj As Object)
        Try
            If obj IsNot Nothing AndAlso Marshal.IsComObject(obj) Then Marshal.ReleaseComObject(obj)
        Catch ex As System.Exception
            _dbg("TryMarshalRelease 異常: ", ex.Message)
        Finally
            obj = Nothing
        End Try
    End Sub

    ' Layer2.5 輔助函數
    ' ---------------------------------------------------------------
    '   - HasSubFoldersFast(folder)             ：快速預測子資料夾有無，TreeView 展開用（記憶體→DB→COM 三層）
    '   - GetLiveFolderSnap(folder)             ：單次 PropertyAccessor 讀 PR_CONTENT_COUNT，專門用於 snapshot 驗證（< 1ms）；失敗回 -999（永遠不等於合法 snap 值）
    '   - FillFolderCacheFromDbRow(fPath, row)  ：DB 命中時一次填滿六個記憶體快取（只填 Not NULL 欄位）
    ' ---------------------------------------------------------------
    Private Function HasSubFoldersFast(folder As Outlook.Folder) As Boolean
        ' ---------------------------------------------------------------
        ' HasSubFoldersFast — 光速版子資料夾加號預測 (專為 TreeView 展開設計)
        ' 2026/4/7 by Gemini, 解決 SSD 讀出後無法出現假節點 + 號，以及嚴重卡頓問題
        ' ---------------------------------------------------------------
        '   呼叫順序：① _cacheFolderCount 記憶體 → ② DbGetFolderStats(fPath).fc → ③ folder.Folders.Count COM
        '   已在 LoadSubFolderToTreeView 第 489 行啟用， 解決 DB 載入後 TreeView 不顯示 "+" 的問題
        '   比直接 folder.Folders.Count 快： 記憶體命中~0μs，DB命中~0.1ms，COM才~1-5ms
        ' ---------------------------------------------------------------
        Dim fPath As String
        Try : fPath = folder.FolderPath
        Catch : Return False : End Try

        Dim fc As Integer
        If _cacheFolderCount.TryGetValue(fPath, fc) Then Return fc > 0

        Dim row = DbGetFolderStats(fPath)
        If row IsNot Nothing AndAlso row.fc >= 0 Then
            _cacheFolderCount.TryAdd(fPath, row.fc) ' 把確認的值送回記憶體快取
            Return row.fc > 0
        End If

        ' 萬一都沒有，直接保底呼叫一次 COM (比 PR_CONTENT_COUNT 驗證還快)
        Try : Return folder.Folders.Count > 0 : Catch : Return False : End Try
    End Function
    Private Function GetLiveFolderSnap(folder As Outlook.Folder) As Integer
        ' 快速讀取 PR_CONTENT_COUNT，專門只用於 SQLite snapshot 驗證
        ' 故意不走完整 Layer3 fallback 的GetMailCount，只走最快的 PropertyAccessor 路徑
        ' 失敗時回傳 -999（不可能等於任何正常 snapshot 值，確保快取失效）
        If _iLikeNoisy Then _dbg(" ├ 開始", folder.Name)
        Try
            Const PR_CC As String = "http://schemas.microsoft.com/mapi/proptag/0x36020003"
            Return CInt(folder.PropertyAccessor.GetProperty(PR_CC))
        Catch
            Try : Return folder.Items.Count : Catch : Return -999 : End Try
        End Try
    End Function
    Private Sub FillFolderCacheFromDbRow(fPath As String, row As FolderStatsDbRow)
        ' DB 命中且 snapshot 驗證通過時，一次填滿所有欄位
        ' 使用 TryAdd：記憶體已有值（例如另一個 Layer2.5 函數剛填入）時不覆蓋
        ' -1 代表 DB 中該欄位尚未存入（例如 mca 還沒算過），跳過，不污染記憶體快取
        If row.mc >= 0 Then _cacheMailCount.TryAdd(fPath, row.mc)
        If row.mca >= 0 Then _cacheMailCountAll.TryAdd(fPath, row.mca)
        If row.fc >= 0 Then _cacheFolderCount.TryAdd(fPath, row.fc)
        If row.fca >= 0 Then _cacheFolderCountAll.TryAdd(fPath, row.fca)
        If row.fs >= 0 Then _cacheFolderSize.TryAdd(fPath, row.fs)
        If row.fsa >= 0 Then _cacheFolderSizeAll.TryAdd(fPath, row.fsa)

        ' by Gemini, 2026/04/10: 填充身分標識與標籤快取
        If Not String.IsNullOrEmpty(row.eid) Then _cacheFolderIDs.TryAdd(fPath, (row.eid, row.sid, row.isMail = 1, row.hasCh = 1))
        If row.isMail >= 0 Then _cacheIsMailFolder.TryAdd(fPath, row.isMail = 1)
    End Sub

#End Region
#End Region

End Class
