Imports System.Collections.Concurrent
Imports System.Runtime.InteropServices
Imports System.Threading
Imports Microsoft.Office.Interop

' === 從頭重新設計 L3 / L2.5 底層計數函數 ===
' 目的: 提供一個純粹的 COM 資料層函數，專注於讀取資料，不做任何流程控制或快取邏輯
'       取代目前散落在各處的 GetMailCountByMAPINew、GetFolderSizeLegacy 等函數，統一為一個簡單的 GetXxxL3 函數
' 架構: L3 純資料層，L2 流程協調層，L2.5 快取代理層，L1 UI 事件層
'       L3 只負責讀取資料夾的本層郵件數 (GetDirectMailCountL3) ，不遞迴、不展開子資料夾，最小化 COM 呼叫量
'       上層流程 (如 ComputeFolderStatsAsync) 負責決定何時呼叫、如何使用結果、快取管理等
' ==============================================================
' === L3 底層 COM 資料層函數群 ===
' 設計原則:
'   1. 每個函數只負責一件事: 讀取單一資料夾或單封郵件的一種屬性
'   2. 不做快取、不做遞迴、不做 BFS 展開——這些全部交給 L2 流程協調層
'   3. Fallback 鏈: RDO → MAPI GetArray → OOM最後手段
'                   parallel.foreach → BFS → Recursive，每層不論成功失敗都丟 Debug message
'   4. 失敗統一回傳 -1 (不回 0) ，讓 L2 能區分「真的是 0」或「讀取失敗」
'   5. 所有 COM 物件在 Finally 中釋放，確保 RCW 不殘留
' ==============================================================

Partial Class Form1

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
    Private Shared _cacheFolderTree As New ConcurrentDictionary(Of String, List(Of Outlook.Folder))     ' GetSortedSubFolders 已排序的子資料夾清單
    Private Shared _cacheSubFolderList As New ConcurrentDictionary(Of String, List(Of Outlook.Folder))  ' GetSubFolderList 的樹狀展開平坦化清單
    Private Shared _cacheAttachPreScan As New ConcurrentDictionary(Of String, FolderCacheTab3)          ' 包含附件的郵件預掃描結果 (速度很快, 不用存入SSD?)
    Private Shared _cacheAttachFilename As New ConcurrentDictionary(Of String, List(Of String))         ' 所有附件檔名清單

    Public Structure MailItemInfo
        ' 候選郵件的純資料結構 (不帶 COM 物件，不受 GC 影響)
        Dim EntryID As String
        Dim Subject As String
        Dim Size As Long
        Dim ReceivedTime As DateTime
        Dim SenderName As String
        Dim AttachCount As Integer
    End Structure
    Public Structure L3ProgressReport
        ' by Gemini, 2026/04/02: 統一進度回報結構，用於 IProgress(Of T)
        Dim CurrentCount As Integer   ' 目前完成數 (郵件數、資料夾數或位元組)
        Dim TotalCount As Integer     ' 總數 (分母)
        Dim Message As String         ' 顯示在狀態列的文字
        Dim IsIndeterminate As Boolean ' 是否為不確定的進度 (跑馬燈模式)
    End Structure
    Private Structure FolderSortInfo
        ' by Gemini, 2026/03/29: 用於 GetSortedSubFolders 排序優化，減少 COM 屬性讀取次數 (O(N) vs O(N log N))
        Dim FolderObj As Outlook.Folder
        Dim Name As String
        Dim HasChinese As Boolean
    End Structure
    Private Class FolderBfsEntry
        ' 候選待掃瞄剪枝的資料夾結構
        Public Folder As Outlook.Folder
        Public ParentIndex As Integer       ' -1 = rootFolder；>= 0 = 父節點在 allEntries 的索引
        Public DirectMailCount As Integer   ' 本層郵件數 (不含子孫) ，由 L3 填入
        Public TotalMailCount As Integer    ' 含子孫郵件總數，L2 底部向上彙總後填入
        Public TotalSubCount As Integer     ' 含子孫資料夾總數，L2 底部向上彙總後填入
        Public IsFromCache As Boolean       ' True = TotalMailCount/TotalSubCount 從快取取得，子樹已剪枝
    End Class

#Region "■ 10 底層 COM 函數群 (新設計，現役主力) "
#Region "  ├ 各種載入函數"
    Private Sub InitOutlookNamespace()
        Dbg("開始")
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
            Dbg("Outlook App OR NameSpace init FAIL", ex.Message)
            TryMarshalRelease(_olNS)
            TryMarshalRelease(_olApp)
            _olApp = Nothing : _olNS = Nothing
            MessageBox.Show("Outlook Object 連接失敗!" & vbCrLf & ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error) : End
        End Try
        Dbg("結束")

    End Sub
    Private Sub InitRdoSession()
        ' 3. 初始化 Redemption Session (目前停用，保留開發記錄)
        Try
            ' ── Redemption Session 初始化, 2026-03-22 測試用:
            ' _rdo = New Redemption.RDOSession()  ' _rdoSession 就等同是outlook.namespace 的意思, 就是Redemption的MAPI session
            ' _rdo.MAPIOBJECT = _olNS.MAPIOBJECT  ' 直接 attach 到現有 Outlook MAPI session, 就不會另開視窗 (必須在 objNameSpace 已建立之後才呼叫)
            ' Dbg("Redemption init OK", $"Version={_rdo.Version}") ' 關鍵: 不建新連線，這樣就不會彈出第二個 Outlook 視窗，也不需要登入
            ' 2026/3/27 總算全部寫好RDO的導入,
            ' 但過程中優化了很多東西之後發現, 好像對效能沒有幫助到太多, 反而是演算法的改進才快更多
            ' RDO 的部份好像反而增加了程式碼複雜度跟拖慢啟動速度而已, 所以先關閉不使用
            Dim unused = InitRedemptionSessionWithoutDeclaration()
        Catch ex As System.Exception
            Dbg("Redemption init FAIL", ex.Message)
            TryMarshalRelease(_rdo)
            _rdo = Nothing
        End Try

    End Sub
    Private Async Function InitRedemptionSessionWithoutDeclaration() As Task
        Dbg("開始")
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
            Dbg("InitRedemption", "AutoDismiss thread 已就緒，開始 New RDOSession")

            ' ✅ Task.Run: UI 執行緒 不阻塞，LoadStoreToTreeView 可以同時跑
            Dim session As Redemption.RDOSession = Nothing
            Await Task.Run(Sub() session = New Redemption.RDOSession())

            ' MAPIOBJECT 必須回 UI 執行緒賦值 (_olNS 是 STA COM 物件)
            session.MAPIOBJECT = _olNS.MAPIOBJECT
            _rdo = session
            Dbg("Redemption init OK", $"Version={_rdo.Version}")
        Catch ex As System.Exception
            _rdo = Nothing
            Dbg("Redemption init FAIL", ex.Message)
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
                    Dbg("AutoDismissRedemption", "逾時: 找不到 TEULAForm") : Return
                End If

                ' ✅ 立刻隱藏，使用者不會看到 EULA dialog 閃出來
                ShowWindow(hWnd, SW_HIDE)
                Dbg("AutoDismissRedemption", $"TEULAForm 隱藏 hWnd=0x{hWnd:X}")

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
                    Dbg("AutoDismissRedemption", "'I agree' PostMessage 送出")
                Else
                    Dbg("AutoDismissRedemption", "找不到 'I agree' (已逾時) ")
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
                    Dbg("AutoDismissRedemption", "'Ok' PostMessage 送出")
                Else
                    Dbg("AutoDismissRedemption", "找不到 'Ok' (已逾時) ")
                End If
            End Sub)

        t.Priority = System.Threading.ThreadPriority.AboveNormal
        t.IsBackground = True
        t.Start()

    End Sub

    Private Function GetSortedStores(space As Outlook.NameSpace) As List(Of Outlook.Store)
        ' ==========================================
        ' 取得排序後的 NameSpace 下所有Outlook.Store
        ' 包含目前config內的所有帳號和所有開啟的PST檔
        ' ==========================================
        ' todo: 這系列基礎函數的效能不夠好, 急需優化
        ' GetSortedStores, GetSortedSubFolders, GetSubFolderList, LoadStoreToTreeView, LoadSubFolderToTreeView()
        Dbg("開始", space.CurrentProfileName)
        ' 使用 Task.Run 將同步操作包裝在獨立的工作線程中 (違反STA安全, 不再使用)
        'Await Task.Run(Sub() stores = space.Stores.Cast(Of Outlook.Store)().ToList())
        ' 遍歷所有Outlook.Store並添加到列表中, 使用LINQ擴充方法就夠快了, 不再使用非同步或Parallel.Foreach了
        Dim stores As List(Of Outlook.Store) = space.Stores.Cast(Of Outlook.Store)().ToList()
        ' 使用 LINQ 排序Outlook.Store
        stores = stores.OrderBy(Function(st) If(TextHasChineseChar(st.DisplayName), 1, 0)).ThenBy(Function(st) st.DisplayName).ToList()
        Dbg("結束", $"Profile={space.CurrentProfileName} | 庫數量: {stores.Count}") ' 2026/03/31 by Gemini
        Return stores
        ' ⚠️ 注意: 不在這裡 ReleaseComObject(space)
        ' space 就是外層的 objNameSpace，釋放後其他地方 (Tab2/Tab3 等) 再用 objNameSpace 會觸發 RCW 已釋放的例外
        ' objNameSpace 的生命週期就只由 Form1_FormClosing 統一管理

    End Function
    Private Function GetSortedSubFolders(folder As Outlook.Folder) As List(Of Outlook.Folder)
        ' ==========================================
        ' 取得引數folder下的所有subFolders並排序後傳回
        ' 優化紀錄: 2026/03/29 by Gemini (Gemini 3.1 Pro)
        ' 1. 加入 L3 過濾: 只保留郵件目錄 (olMailItem)，排除行事曆/聯絡人等
        ' 2. 單次屬性讀取: 先快取 Name 後排序，避開 LINQ 重複打 COM 的 N log N 效能陷阱
        ' ==========================================
        If _iLikeNoisy Then Dbg("開始", folder.Name)
        Dim fPath As String = folder.FolderPath, value As List(Of Outlook.Folder) = Nothing
        If _cacheFolderTree.TryGetValue(fPath, value) Then Return value

        ' ① 單次遍歷: 抓取實體物件與名稱屬性 (O(N) COM 呼叫)
        Dim infoList As New List(Of FolderSortInfo)
        Try
            ''' 2024/5/13記錄: 已經試過很多種優化, 好像很難再比現在下面這二行LINQ還快了??
            ''' Dim subFolders As List(Of Outlook.Folder) = folder.Folders.Cast(Of Outlook.Folder)().ToList()
            ''' subFolders = subFolders.OrderBy(Function(subFolder) If(TextHasChineseChar(subFolder.Name), 1, 0)).
            '''                         ThenBy(Function(subFolder) subFolder.Name).ToList()
            ''' [上面是舊版紀錄] 原本使用 LINQ 直接 Cast().ToList() 後排序，缺點是 OrderBy 會重複觸發 COM 讀取屬性
            ' [下面是新版優化] by Gemini, 2026/03/29:
            ' 1. 動態過濾: 根據 checkIncludeAllFolders.Checked 決定是否顯示非郵件目錄
            ' 2. 單次屬性讀取: 先快取 Name 後排序，避開 LINQ 重複打 COM 的 N log N 效能陷阱

            ' 2026/4/7 進一步優化 by Gemini: 把 checkIncludeAllFolders.Checked 提前到迴圈外面讀取，避免每次迴圈都去問一次 UI 狀態
            Dim isIncludeAll As Boolean = checkIncludeAllFolders.Checked
            For Each subFolder As Outlook.Folder In folder.Folders
                ' 🔥 核心過濾: 正常若「沒勾選顯示全部」且「不是郵件資料夾」時就排除
                ' by Gemini, 2026/04/01: 明確短路邏輯
                ' 如果「沒勾選顯示全部」且「不是郵件資料夾」時就排除
                If Not isIncludeAll Then
                    ' 只有在非全部顯示模式下，才去調用 IsMailFolder (現在已具備快取)
                    If Not IsMailFolder(subFolder) Then Continue For
                End If
                Dim folderName As String = subFolder.Name
                infoList.Add(New FolderSortInfo With {.FolderObj = subFolder, .Name = folderName, .HasChinese = TextHasChineseChar(folderName)})
            Next
        Catch ex As System.Exception
            Dbg("GetSortedSubFolders 遍歷失敗", ex.Message)
        End Try

        ' ② 純記憶體排序: 完全不觸發 COM 呼叫 (快速且不卡 UI)
        'Dim sortedFolders = infoList.OrderBy(Function(i) If(i.HasChinese, 1, 0)).
        '                             ThenBy(Function(i) i.Name).
        '                             Select(Function(i) i.FolderObj).ToList()
        ' 2026/4/7 進一步優化 by Gemini: 加入 StringComparer.OrdinalIgnoreCase 略過語系分析，爆發性提速
        Dim sortedFolders = infoList.OrderBy(Function(i) If(i.HasChinese, 1, 0)).
                                     ThenBy(Function(i) i.Name, StringComparer.OrdinalIgnoreCase).
                                     Select(Function(i) i.FolderObj).ToList()

        _cacheFolderTree(fPath) = sortedFolders
        If _iLikeNoisy Then Dbg("結束", $"{folder.Name} | 子資料夾數: {sortedFolders.Count}")
        Return sortedFolders

    End Function
    Private Function GetSubFolderList(rootFolder As Outlook.Folder, includeSubF As Boolean, Optional progress As IProgress(Of L3ProgressReport) = Nothing) As List(Of Outlook.Folder)
        ' --------------------------------------------------------------
        ' GetSubFolderList: 取得目標資料夾下, 整個資料夾子樹清單 (BFS，含子資料夾)
        ' ① OOM BFS: 目前唯一的路徑，使用 Outlook Object Model 廣度優先搜尋
        ' by Gemini, 2026/04/02: 導入 IProgress 支援
        ' --------------------------------------------------------------
        If _iLikeNoisy Then Dbg("開始", rootFolder.Name)
        Dim sw As New Stopwatch() : sw.Start()

        If includeSubF Then
            Dim cachedList As List(Of Outlook.Folder) = Nothing
            If _cacheSubFolderList.TryGetValue(rootFolder.FolderPath, cachedList) Then
                sw.Stop()
                Dbg("結束", $"{rootFolder.Name} (Cache Hit) | 資料夾總計: {cachedList.Count} | {sw.ElapsedMilliseconds}ms")
                Return cachedList
            End If
        End If

        Dim result As New List(Of Outlook.Folder)
        result.Add(rootFolder)
        If Not includeSubF Then
            sw.Stop()
            Dbg("結束", $"{rootFolder.Name} (Single) | {sw.ElapsedMilliseconds}ms")
            Return result     ' 若不包含子資料夾，直接回傳只有 rootFolder 的清單
        End If

        ' 取得目標資料夾清單 (BFS，含子資料夾)
        Static swThrottle As New Stopwatch()
        If Not swThrottle.IsRunning Then swThrottle.Start()

        Dim isIncludeAll As Boolean = checkIncludeAllFolders.Checked
        Dim queue As New Queue(Of Outlook.Folder)
        queue.Enqueue(rootFolder)
        While queue.Count > 0
            Dim current As Outlook.Folder = queue.Dequeue()
            Try
                For Each subFolder As Outlook.Folder In current.Folders
                    ' 🔥 核心過濾: 正常若「沒勾選顯示全部」且「不是郵件資料夾」時就排除
                    If Not isIncludeAll Then  ' by Gemini, 2026/04/08: 提前至迴圈外讀取，避免高頻 UI 存取
                        If Not IsMailFolder(subFolder) Then Continue For
                    End If
                    result.Add(subFolder)       ' 把子資料夾加入結果清單
                    queue.Enqueue(subFolder)    ' 把子資料夾加入佇列，繼續往下搜尋
                Next

            Catch ex As System.Exception
                Dbg("GetSubFolderList ① OOM 失敗", current.Name & " - " & ex.Message)
            End Try

            ' by Gemini, 2026/04/05: 在大迴圈中增加中斷檢查，減少長時間阻塞
            ' 但其實這裡沒有放await task.delay(1), 應該永遠偵測不到下面這行 :-(
            If _cancelRequested Then Exit While

            ' by Gemini, 2026/04/02: 100ms 節流回報已發現的資料夾數
            If progress IsNot Nothing AndAlso swThrottle.ElapsedMilliseconds >= 100 Then
                progress.Report(New L3ProgressReport With {.CurrentCount = result.Count, .Message = $"正在展開資料夾結構: 已發現 {result.Count} 個資料夾..."})
                swThrottle.Restart()
                ' BFS 展開通常很快，但在超大 PST 中可能卡頓，Yield 讓 UI 有機會處理中斷
                'Await Task.Yield() ' 此函數不是 Async，故不使用 Await
            End If
        End While
        sw.Stop()

        If includeSubF AndAlso Not _cancelRequested AndAlso result.Count > 0 Then
            _cacheSubFolderList.TryAdd(rootFolder.FolderPath, result)
        End If

        Dbg("結束", $"{rootFolder.Name} (BFS) | 資料夾總計: {result.Count} | {sw.ElapsedMilliseconds}ms")
        Return result

    End Function
    Private Function GetSubFolderList_RDO(rootFolder As Redemption.RDOFolder, includeSubF As Boolean) As List(Of Redemption.RDOFolder)
        ' --------------------------------------------------------------
        ' 2026/3/24 by Gemini: GetSubFolderList_RDO
        ' 目的: 專門提供給 RDO 平行路徑使用，回傳 List(Of Redemption.RDOFolder)
        ' 說明: 因為 Redemption 是 free-threaded，可以用 Parallel.ForEach 安全平行展開子樹
        ' --------------------------------------------------------------
        Dbg("開始", rootFolder.Name)
        Dim sw As New Stopwatch() : sw.Start()
        Dim resultBag As New ConcurrentBag(Of Redemption.RDOFolder)
        resultBag.Add(rootFolder)
        If Not includeSubF Then
            sw.Stop()
            Dbg("結束", $"{rootFolder.Name} (RDO-Single) | {sw.ElapsedMilliseconds}ms")
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
                        Dbg("GetSubFolderList_RDO Error: ", current.Name & " - " & ex.Message)
                    End Try
                End Sub)
        Loop
        sw.Stop()
        Dbg("結束", $"{rootFolder.Name} (RDO-Parallel BFS) | 資料夾總計: {resultBag.Count} | {sw.ElapsedMilliseconds}ms")
        Return resultBag.ToList()

    End Function

    Private Sub LoadStoreToTreeView(storeList As List(Of Outlook.Store), tv As TreeView)
        Dbg("開始", tv.Name)
        ' 2024/5/17全部重寫, 只先動態載入一層的rootFolder, 不花時間遍歷所有的subFolders
        ' 2024/5/19試過Task.Run(), Parallel.Foreach跟LINQ擴充方法了, 都沒有比較快, 別再試了, 就算virtual mode也沒有比我現在的lazy load還快
        'tv.BeginUpdate()
        'For Each st In storeList
        '    Dim root As Outlook.Folder = st.GetRootFolder
        '    'Dim node As TreeNode = Await Task.Run(Function() Me.Invoke(Function() tv.Nodes.Add(root.Name)))
        '    Dim node As TreeNode = tv.Nodes.Add(root.Name)
        '    node.Tag = root
        '    If root.Folders.Count > 0 Then node.Nodes.Add(":::") '若發現底下還有subFolders也不讀取, 只先填入一個假的":::"暫代, 才能出現"+"號
        'Next
        'tv.EndUpdate()
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
            ' by Gemini, 2026/04/04: 移除每筆 Dbg("", root.Name)，改在結束時一次輸出節點總計（Issue 1 高頻去噪）
        Next
        tv.Nodes.AddRange(nodeList.ToArray()) ' 將所有組裝好的節點一次性添加到 tv.Nodes
        Dbg("結束", $"{tv.Name} 共 {nodeList.Count} 個 Store")

    End Sub
    Private Sub LoadSubFolderToTreeView(sender As Object, e As TreeViewCancelEventArgs)
        Dbg("開始", sender.Name)
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
                    If GetCachedFolderCount(folder) > 0 Then node.Nodes.Add(":::")
                    'If HasCachedSubFoldersFast(folder) > 0 Then node.Nodes.Add(":::") ' 2026/4/7 by Gemini, 光速版子資料夾加號預測 (專為 TreeView 展開設計)
                Catch ex As System.Exception : End Try
                nodeList.Add(node) ' 先加進List在記憶體中快速操作, 而不是直接加到Treeview.Nodes
                ' by Gemini, 2026/04/04: 移除每筆 Dbg("", ...)，改在結束時一次輸出（Issue 1 高頻去噪）
            Next
            selectedNode.Nodes.AddRange(nodeList.ToArray()) ' 將所有節點一次性添加到 selectedNode.Nodes
        End If
        Dbg("結束", $"{selectedFolder.Name} 展開 {sortedFolders.Count} 個子資料夾")

    End Sub
#End Region
#Region "  ├ L2.5 快取存取點 (Cache Proxy Layer)"
    ' 2026/03/27 by Gemini: 新增 L2.5 快取存取點 (Cache Proxy Layer)，保護 L3 不被頻繁呼叫
    ' 2026/04/07: Phase 2 — 在記憶體 miss 時加入 SQLite lazy SELECT，命中後一次填滿所有欄位
    '             寫入仍由 SaveCachesToSQLiteAsync (SaveCache 按鈕) 批次處理，本層不做即時寫入

    ' 2026/4/7 by Claude, 加入「持久化快取」存入SSD:
    '             讀資料時若快取不存在就先去撈SSD, 有就放進快取, 沒有才算cache miss去COM讀取
    '
    ' 呼叫順序 (每個 GetCachedXxx 函數):
    '   ① 記憶體命中 → 直接回傳（最快，0 COM call）
    '   ② DB 命中 + snapshot 驗證通過 → 填滿記憶體快取 → 回傳（快，0 COM call）
    '   ③ DB miss 或 snapshot 不符 → 呼叫 L3 → 填入記憶體快取 → 回傳（慢，有 COM call）
    '
    ' snapshot 驗證: DB 儲存的 content_count_snapshot = save 時的 PR_CONTENT_COUNT 值
    '   用 GetLivePrContentCount (單次 PropertyAccessor call) 與 snapshot 比對
    '   相同 → 快取仍有效；不同 → 資料夾內容已變，跳過 DB，呼叫 L3

    ' ---------------------------------------------------------------
    ' L2.5 輔助函數
    ' ---------------------------------------------------------------
    Private Function GetLivePrContentCount(folder As Outlook.Folder) As Integer
        ' 快速讀取 PR_CONTENT_COUNT，用於 SQLite snapshot 驗證
        ' 不走完整 L3 fallback chain，只走最快的 PropertyAccessor 路徑
        ' 失敗時回傳 -999（不可能等於任何正常 snapshot 值，確保快取失效）
        ' 2026/4/7 by Claude, 加入「持久化快取」存入SSD:
        ' 讀資料時若快取不存在就先去撈SSD, 有就放進快取, 沒有才算cache miss去COM讀取
        Try
            Const PR_CC As String = "http://schemas.microsoft.com/mapi/proptag/0x36020003"
            Return CInt(folder.PropertyAccessor.GetProperty(PR_CC))
        Catch
            Try : Return folder.Items.Count : Catch : Return -999 : End Try
        End Try
    End Function

    Private Sub PopulateFolderCacheFromDbRow(fPath As String, row As FolderStatsDbRow)
        ' DB 命中且 snapshot 驗證通過時，一次填滿所有欄位
        ' 使用 TryAdd：記憶體已有值（例如另一個 L2.5 函數剛填入）時不覆蓋
        ' -1 代表 DB 中該欄位尚未存入（例如 mca 還沒算過），跳過，不污染記憶體快取
        If row.mc >= 0 Then _cacheMailCount.TryAdd(fPath, row.mc)
        If row.mca >= 0 Then _cacheMailCountAll.TryAdd(fPath, row.mca)
        If row.fc >= 0 Then _cacheFolderCount.TryAdd(fPath, row.fc)
        If row.fca >= 0 Then _cacheFolderCountAll.TryAdd(fPath, row.fca)
        If row.fs >= 0 Then _cacheFolderSize.TryAdd(fPath, row.fs)
        If row.fsa >= 0 Then _cacheFolderSizeAll.TryAdd(fPath, row.fsa)
    End Sub

    ' ---------------------------------------------------------------
    ' GetCachedMailCount — 單一資料夾本層郵件數 (PR_CONTENT_COUNT)
    ' 2026/4/7 by Claude, 加入「持久化快取」存入SSD:
    ' 讀資料時若快取不存在就先去撈SSD, 有就放進快取, 沒有才算cache miss去COM讀取
    ' ---------------------------------------------------------------
    Private Function GetCachedMailCount(folder As Outlook.Folder) As Integer
        Dim count As Integer, fPath As String = folder.FolderPath
        If _cacheMailCount.TryGetValue(fPath, count) Then Return count  ' ① 記憶體命中

        ' ② DB lazy load：命中且 mc 有效且 snapshot 吻合 → 一次填滿所有欄位
        Dim row = DbGetFolderStats(fPath)
        If row IsNot Nothing AndAlso row.mc >= 0 AndAlso GetLivePrContentCount(folder) = row.snap Then
            PopulateFolderCacheFromDbRow(fPath, row) : Return row.mc
        End If

        ' ③ L3 呼叫
        count = GetMailCount(folder)
        _cacheMailCount.TryAdd(fPath, count)
        Return count
    End Function

    ' ---------------------------------------------------------------
    ' HasCachedSubFoldersFast — 光速版子資料夾加號預測 (專為 TreeView 展開設計)
    ' 2026/4/7 by Gemini, 解決 SSD 讀出後無法出現假節點 + 號，以及嚴重卡頓問題
    ' todo: 好像沒有作用???
    ' ---------------------------------------------------------------
    Private Function HasCachedSubFoldersFast(folder As Outlook.Folder) As Boolean
        Dim fPath As String
        Try
            fPath = folder.FolderPath
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

    ' ---------------------------------------------------------------
    ' GetCachedFolderCount — 單一資料夾直屬子資料夾數
    ' 2026/4/7 by Claude, 加入「持久化快取」存入SSD:
    ' 讀資料時若快取不存在就先去撈SSD, 有就放進快取, 沒有才算cache miss去COM讀取
    ' ---------------------------------------------------------------
    Private Function GetCachedFolderCount(folder As Outlook.Folder) As Integer
        Dim count As Integer, fPath As String = folder.FolderPath
        If _cacheFolderCount.TryGetValue(fPath, count) Then Return count  ' ①

        ' ② DB lazy load（fc 欄位）
        Dim row = DbGetFolderStats(fPath)
        If row IsNot Nothing AndAlso row.fc >= 0 AndAlso GetLivePrContentCount(folder) = row.snap Then
            PopulateFolderCacheFromDbRow(fPath, row) : Return row.fc
        End If

        ' ③ L3
        count = GetFolderCount(folder)
        _cacheFolderCount.TryAdd(fPath, count)
        Return count
    End Function

    ' ---------------------------------------------------------------
    ' GetCachedMailCountAllAsync — 整棵子樹的郵件總數
    ' 2026/4/7 by Claude, 加入「持久化快取」存入SSD:
    ' 讀資料時若快取不存在就先去撈SSD, 有就放進快取, 沒有才算cache miss去COM讀取
    ' ---------------------------------------------------------------
    Private Async Function GetCachedMailCountAllAsync(folder As Outlook.Folder, Optional progress As IProgress(Of L3ProgressReport) = Nothing) As Task(Of Long)
        Dim count As Integer, fPath As String = folder.FolderPath
        If _cacheMailCountAll.TryGetValue(fPath, count) Then Return count  ' ①

        ' ② DB lazy load（mca 欄位）
        Dim row = DbGetFolderStats(fPath)
        If row IsNot Nothing AndAlso row.mca >= 0 AndAlso GetLivePrContentCount(folder) = row.snap Then
            PopulateFolderCacheFromDbRow(fPath, row) : Return row.mca
        End If

        ' ③ L3
        Dim total As Long = Await GetMailCountAll(folder, progress)
        If total >= 0 AndAlso Not _cancelRequested Then _cacheMailCountAll.TryAdd(fPath, CInt(total))
        Return total
    End Function

    ' ---------------------------------------------------------------
    ' GetCachedFolderCountAllAsync — 整棵子樹的資料夾總數
    ' 2026/4/7 by Claude, 加入「持久化快取」存入SSD:
    ' 讀資料時若快取不存在就先去撈SSD, 有就放進快取, 沒有才算cache miss去COM讀取
    ' ---------------------------------------------------------------
    Private Async Function GetCachedFolderCountAllAsync(folder As Outlook.Folder) As Task(Of Integer)
        Dim count As Integer, fPath As String = folder.FolderPath
        If _cacheFolderCountAll.TryGetValue(fPath, count) Then Return count  ' ①

        ' ② DB lazy load（fca 欄位）
        Dim row = DbGetFolderStats(fPath)
        If row IsNot Nothing AndAlso row.fca >= 0 AndAlso GetLivePrContentCount(folder) = row.snap Then
            PopulateFolderCacheFromDbRow(fPath, row) : Return row.fca
        End If

        ' ③ L3
        count = Await GetFolderCountAll(folder)
        If count >= 0 AndAlso Not _cancelRequested Then _cacheFolderCountAll.TryAdd(fPath, count)
        Return count
    End Function

    ' ---------------------------------------------------------------
    ' GetCachedFolderSizeAsync — 單一資料夾本層大小 (GetTable 加總)
    ' 2026/4/7 by Claude, 加入「持久化快取」存入SSD:
    ' 讀資料時若快取不存在就先去撈SSD, 有就放進快取, 沒有才算cache miss去COM讀取
    ' ---------------------------------------------------------------
    Private Async Function GetCachedFolderSizeAsync(folder As Outlook.Folder) As Task(Of Long)
        ' 2026/3/29 by Gemini: L2.5 快取代理層 - 取得單一資料夾本層的大小 (含快取機制)
        Dim size As Long, fPath As String = folder.FolderPath
        If _cacheFolderSize.TryGetValue(fPath, size) Then Return size  ' ①

        ' ② DB lazy load（fs 欄位）
        Dim row = DbGetFolderStats(fPath)
        If row IsNot Nothing AndAlso row.fs >= 0 AndAlso GetLivePrContentCount(folder) = row.snap Then
            PopulateFolderCacheFromDbRow(fPath, row) : Return row.fs
        End If

        ' ③ L3
        size = Await GetFolderSize(folder)
        If size >= 0 Then _cacheFolderSize.TryAdd(fPath, size)
        Return size
    End Function

    ' ---------------------------------------------------------------
    ' GetCachedFolderSizeAllAsync — 整棵子樹大小總計
    ' 2026/4/7 by Claude, 加入「持久化快取」存入SSD:
    ' 讀資料時若快取不存在就先去撈SSD, 有就放進快取, 沒有才算cache miss去COM讀取
    ' ---------------------------------------------------------------
    Private Async Function GetCachedFolderSizeAllAsync(folder As Outlook.Folder) As Task(Of Long)
        ' 2026/3/29 by Gemini: L2.5 快取代理層 - 取得整棵子樹的大小總計 (含快取機制)
        Dim size As Long, fPath As String = folder.FolderPath
        If _cacheFolderSizeAll.TryGetValue(fPath, size) Then Return size  ' ①

        ' ② DB lazy load（fsa 欄位）
        Dim row = DbGetFolderStats(fPath)
        If row IsNot Nothing AndAlso row.fsa >= 0 AndAlso GetLivePrContentCount(folder) = row.snap Then
            PopulateFolderCacheFromDbRow(fPath, row) : Return row.fsa
        End If

        ' ③ L3
        size = Await GetFolderSizeAll(folder)
        If size >= 0 AndAlso Not _cancelRequested Then _cacheFolderSizeAll.TryAdd(fPath, size)
        Return size
    End Function

    ' ---------------------------------------------------------------
    ' GetCachedMailWithAttachment — Tab3 Phase1：含附件的候選郵件清單
    ' 2026/4/7 by Claude, 加入「持久化快取」存入SSD:
    ' 讀資料時若快取不存在就先去撈SSD, 有就放進快取, 沒有才算cache miss去COM讀取
    ' ---------------------------------------------------------------
    Private Async Function GetCachedMailWithAttachment(folder As Outlook.Folder, progress As IProgress(Of L3ProgressReport)) As Task(Of List(Of MailItemInfo))
        ' by Gemini, 2026/04/05: L2.5 快取代理層 - Tab3 Phase 1 快取 - 取得單一資料夾本層含附件的郵件清單
        ' 2026/04/07: Phase 2 — 加入 DB lazy load (mail_basic)
        Dbg("開始", folder.Name)
        Dim key As String = folder.FolderPath
        Dim currentCount As Integer = GetCachedMailCount(folder)  ' 依賴同層快取（本身已有 DB lazy load）

        ' ① 記憶體命中
        Dim entry As FolderCacheTab3
        If _cacheAttachPreScan.TryGetValue(key, entry) AndAlso entry.ItemCountWhenCached = currentCount Then Return entry.mailWithAttachment

        ' ② DB lazy load (mail_basic)：item_count_snap == currentCount → 快取仍有效
        Dim dbResult = DbGetMailBasic(key)
        If dbResult IsNot Nothing AndAlso dbResult.Snap = currentCount Then
            Dim cached As New FolderCacheTab3 With {.mailWithAttachment = dbResult.Mails, .ItemCountWhenCached = currentCount}
            _cacheAttachPreScan(key) = cached   ' 覆蓋式寫入，確保 ItemCountWhenCached 對應正確
            Dbg("DB 命中", $"{folder.Name} ({dbResult.Mails.Count} 封)")
            Return dbResult.Mails
        End If

        ' ③ L3 呼叫
        Dim targetMailList As List(Of MailItemInfo) = Await GetMailWithAttachment(folder, progress)
        _cacheAttachPreScan(key) = New FolderCacheTab3 With {.mailWithAttachment = targetMailList, .ItemCountWhenCached = currentCount}
        ' 2026/04/05: 不使用 TryAdd/TryUpdate，確保最後的 cache entry 是正確的 (ItemCountWhenCached 與 mail list 對應)
        Return targetMailList
    End Function

    ' ---------------------------------------------------------------
    ' GetCachedAttachFilename — Tab3 Phase2：附件檔名清單 (by EntryID)
    ' 2026/4/7 by Claude, 加入「持久化快取」存入SSD:
    ' 讀資料時若快取不存在就先去撈SSD, 有就放進快取, 沒有才算cache miss去COM讀取
    ' ---------------------------------------------------------------
    Private Function GetCachedAttachFilename(mail As MailItemInfo) As List(Of String)
        ' by Gemini, 2026/04/04: L2.5 快取代理層 - 取得附件檔名清單 (含 _cacheAttachFilename 機制)
        ' 2026/04/07: Phase 2 — 加入 DB lazy load (mail_attachments)「持久化快取」存入SSD
        ' 讀資料時若快取不存在就先去撈SSD, 有就放進快取, 沒有才算cache miss去COM讀取
        Dim result As List(Of String) = Nothing
        If _cacheAttachFilename.TryGetValue(mail.EntryID, result) Then Return result  ' ①

        ' ② DB lazy load (mail_attachments)
        result = DbGetAttachFilenames(mail.EntryID)
        If result IsNot Nothing Then
            _cacheAttachFilename.TryAdd(mail.EntryID, result)
            Return result
        End If

        ' ③ L3 呼叫
        result = GetAttachFilename(mail)
        If result IsNot Nothing Then _cacheAttachFilename.TryAdd(mail.EntryID, result)
        Return result
    End Function

    Friend Async Function PreloadAttachByRDOAsync1(sourceList As List(Of MailItemInfo), progress As IProgress(Of L3ProgressReport)) As Task
        ' by Gemini, 2026/04/05: L2.5 快取代理層 - 批次預熱附件檔名快取
        ' 利用 Redemption (RDO) Free-Threaded 安全的特性，在進入 L2 迴圈前平行提早把附件檔名讀進 _cacheAttachFilename。
        ' 完全不動原來的迴圈運作邏輯，以防呆的姿態大幅壓縮等待時間。
        If _rdo Is Nothing OrElse sourceList.Count = 0 Then Return

        Dbg("開始", $"RDO平行預載 {sourceList.Count} 筆")
        Dim swThrottle As New Stopwatch() : swThrottle.Start()
        Dim swTotal As New Stopwatch() : swTotal.Start()
        Dim processed As Integer = 0
        Dim total As Integer = sourceList.Count

        Await Task.Run(Sub()
                           Dim parallelOptions As New ParallelOptions With {.MaxDegreeOfParallelism = Environment.ProcessorCount}
                           Parallel.ForEach(sourceList, parallelOptions,
                                            Sub(mail)
                                                If _cancelRequested Then Return ' 迴圈內也要檢查ESC中斷，盡快停止處理

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
                                                If progress IsNot Nothing AndAlso (swThrottle.ElapsedMilliseconds >= 100 OrElse currentProcessed = total) Then
                                                    Dim elapsedSec As Double = Math.Max(swTotal.Elapsed.TotalSeconds, 0.001)
                                                    Dim speed As Double = currentProcessed / elapsedSec
                                                    progress.Report(New L3ProgressReport With {.CurrentCount = currentProcessed,
                                                                                               .TotalCount = total,
                                                                                               .Message = $"Phase 2 (RDO 預載快取): {currentProcessed} / {total} ({speed:F0} 封/秒)"})
                                                    swThrottle.Restart()
                                                End If
                                            End Sub)
                       End Sub)
        Dbg("結束", $"RDO 預載完成，處理共 {processed} 筆")
    End Function
    Friend Async Function PreloadAttachByRDOAsync2(sourceList As List(Of MailItemInfo), progress As IProgress(Of L3ProgressReport)) As Task
        ' by AntiGravity, 2026/04/07: 實驗性質 - 使用 Task.WhenAll + SemaphoreSlim，試圖推高 SSD I/O 並發度
        If _rdo Is Nothing OrElse sourceList.Count = 0 Then Return

        Dbg("開始", $"WhenAll平行預載 {sourceList.Count} 筆")
        Dim swThrottle As New Stopwatch() : swThrottle.Start()
        Dim swTotal As New Stopwatch() : swTotal.Start()
        Dim processed As Integer = 0
        Dim total As Integer = sourceList.Count

        ' 設定並發數：嘗試設為 CPU 核心數的 4 倍，壓榨 SSD 的 Queue Depth
        Dim maxConcurrency As Integer = Environment.ProcessorCount * 4
        Dim throttler As New SemaphoreSlim(maxConcurrency)
        Dim tasks As New List(Of Task)()

        For Each m As MailItemInfo In sourceList
            If _cancelRequested Then Exit For
            Dim mail = m ' 在 lambda 中避免變數捕獲問題

            tasks.Add(Task.Run(Async Function()
                                   Await throttler.WaitAsync()
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
                                       If progress IsNot Nothing AndAlso (swThrottle.ElapsedMilliseconds >= 100 OrElse currentProcessed = total) Then
                                           Dim elapsedSec As Double = Math.Max(swTotal.Elapsed.TotalSeconds, 0.001)
                                           Dim speed As Double = currentProcessed / elapsedSec
                                           progress.Report(New L3ProgressReport With {.CurrentCount = currentProcessed,
                                                                   .TotalCount = total,
                                                                   .Message = $"Phase 2 (WhenAll 預載): {currentProcessed} / {total} ({speed:F0} 封/秒)"})
                                           swThrottle.Restart()
                                       End If
                                   Finally
                                       throttler.Release()
                                   End Try
                               End Function))
        Next

        If tasks.Count > 0 Then Await Task.WhenAll(tasks)
        Dbg("結束", $"WhenAll 預載完成，處理共 {processed} 筆")
    End Function
#End Region
#Region "  ├ L3 直接存取底層計數函數"
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
                Dim storeId As String = If(parentFolder IsNot Nothing, parentFolder.StoreID, "")
                rdoMail = TryCast(_rdo.GetMessageFromID(mail.EntryID, storeId), Redemption.RDOMail)
                If rdoMail IsNot Nothing Then
                    Dim sz As Long = CLng(rdoMail.Size)
                    ' Dbg("GetMailSize ⓪ RDO 成功", $"size={sz}") ' 高頻率項目平時不輸出 Log
                    Return sz
                End If
            Catch ex As System.Exception
                Dbg("GetMailSize ⓪ RDO 失敗，走MAPI fallback", ex.Message)
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
            Dbg("GetMailSize ① PR_MESSAGE_SIZE_EXTENDED失敗", ex.Message)
        End Try

        ' ② MAPI: PR_MESSAGE_SIZE (0x0E080003, PT_LONG) — 32-bit，超大郵件理論上溢位
        Try
            Const PR_SIZE As String = "http://schemas.microsoft.com/mapi/proptag/0x0E080003"
            ' by Gemini, 2026/03/29: 同上，移除 TypeOf 判斷
            Return CLng(mail.PropertyAccessor.GetProperty(PR_SIZE))
        Catch ex As System.Exception
            Dbg("GetMailSize ② PR_MESSAGE_SIZE失敗", ex.Message)
        End Try

        ' ③ OOM: mail.Size (Integer，超大郵件理論上不準，但實務上 PST 內不會發生)
        Try
            Return CLng(mail.Size)
        Catch ex As System.Exception
            Dbg("GetMailSize ③ OOM mail.Size也失敗", ex.Message)
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
        If _iLikeNoisy Then Dbg("開始", folder.Name)
        Dim sw As New Stopwatch() : sw.Start()

        ' ⓪ Redemption: RDOFolder.Items.Count
        If _rdo IsNot Nothing Then
            Dim rdoFolder As Redemption.RDOFolder = Nothing
            Try
                rdoFolder = _rdo.GetFolderFromID(folder.EntryID, folder.StoreID)
                Dim count As Integer = rdoFolder.Items.Count
                ' by Gemini, 2026/04/04: 成功路徑靜默，不輸出 Dbg（Issue 1 高頻去噪）
                Return count
            Catch ex As System.Exception
                Dbg("錯誤路徑", $"GetMailCount ⓪ RDO: {folder.Name} | {ex.Message}")
            Finally
                TryMarshalRelease(rdoFolder)
            End Try
        End If
        ' ① MAPI: PR_CONTENT_COUNT (0x36020003)
        Try
            Const PR_CONTENT_COUNT As String = "http://schemas.microsoft.com/mapi/proptag/0x36020003"
            Dim count As Integer = CInt(folder.PropertyAccessor.GetProperty(PR_CONTENT_COUNT))
            ' by Gemini, 2026/04/04: 成功路徑靜默，不輸出 Dbg
            Return count
        Catch ex As System.Exception
            Dbg("錯誤路徑", $"GetMailCount ① MAPI: {folder.Name} | {ex.Message}")
        End Try
        ' ② OOM: folder.Items.Count
        Try
            Dim items As Outlook.Items = Nothing
            Try
                items = folder.Items
                Dim count As Integer = items.Count
                ' by Gemini, 2026/04/04: 成功路徑靜默，不輸出 Dbg
                Return count
            Finally
                TryMarshalRelease(items)
            End Try
        Catch ex As System.Exception
            Dbg("錯誤路徑", $"GetMailCount ② OOM: {folder.Name} | {ex.Message}")
        End Try
        sw.Stop()
        Dbg("結束", $"FAIL: {folder.Name} | -1 | {sw.ElapsedMilliseconds}ms")
        Return -1

    End Function
    Private Async Function GetMailCountAll(rootFolder As Outlook.Folder, Optional progress As IProgress(Of L3ProgressReport) = Nothing) As Task(Of Long)
        ' --------------------------------------------------------------
        ' GetMailCountAll v3.5: 讀取某資料夾及其整棵子樹的郵件總數
        ' by Gemini, 2026/04/02: 升級為 IProgress(Of L3ProgressReport) 並加入 100ms 節流回報
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
        '                   GetSubFolderList BFS 展開 + GetMailCount(L3) 逐一加總
        '                   支援取消檢查和 onProgress 進度回報
        '                   平行路徑失敗時的安全 fallback
        '   ③ 遞迴 fallback:
        '                   GetSubFolderList 本身失敗時 (極少見) 的最後保險
        '                   無法精確回報進度，但確保加總結果正確
        '                   todo: 這裡遞迴會重複呼叫 GetSubFolderList，若 ③ 常被觸發需檢查根本原因
        '   ④ Return -1: 四層都失敗，由 L2 決定如何處理
        '
        ' cancelRequested:
        '   檢查 _cancelRequested 旗標，取消時回傳 -1，由 L1 判斷是否需要清空 UI
        '   ⓪ Redemption 路徑不插入取消檢查 (單次 call，幾乎瞬間完成)
        '
        ' onProgress 參數 (可選):
        '   傳入 Action(Of Integer, Integer) callback
        '   L2 每處理一個資料夾回報 (已完成數, 總數)，讓 L1 更新狀態列
        '   不需要進度回報時傳 Nothing
        '   ⓪ 和 ① 路徑不觸發 onProgress，② 路徑才會逐一回報
        '
        ' 取代:
        '   GetMailCountByMAPINew 的整棵子樹加總用途
        '   GetMailCountAllParallel (v3.0 已合併，舊版可廢棄)
        ' --------------------------------------------------------------
        Dbg("開始", rootFolder.Name)

        ' ⓪ Redemption: TotalItemCount 直接回傳整棵子樹郵件總數
        '   一次 COM call 結束，不需要任何 BFS 遍歷或平行處理
        '   2026-03-22 新增
        If _rdo IsNot Nothing Then
            Dim rdoFolder As Redemption.RDOFolder = Nothing
            Try
                rdoFolder = _rdo.GetFolderFromID(rootFolder.EntryID, rootFolder.StoreID)
                Dim total As Long = CLng(rdoFolder.TotalItemCount)
                Dbg("結束", $"⓪ RDO 成功: {rootFolder.Name} | TotalItemCount={total}")
                Return total
            Catch ex As System.Exception
                Dbg("GetMailCountAll ⓪ RDO 失敗，走平行BFS fallback", $"{rootFolder.Name} | {ex.Message}")
            Finally
                TryMarshalRelease(rdoFolder)
            End Try
        End If

        ' 2026/3/24 by Gemini: ① 平行 BFS (RDO)
        '   使用 GetSubFolderList_RDO 取得清單，以 Parallel.ForEach 搭配 Interlocked.Add 快速加總
        '   Redemption (RDO) 是 free-threaded，在背景平行執行安全且極為高效
        If _rdo IsNot Nothing Then
            Dim rdoRoot As Redemption.RDOFolder = Nothing
            Try
                rdoRoot = _rdo.GetFolderFromID(rootFolder.EntryID, rootFolder.StoreID)
                Dim rdoFolderList As List(Of Redemption.RDOFolder) = GetSubFolderList_RDO(rdoRoot, includeSubF:=True)
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
                            Dbg("GetMailCountAll ① 略過失敗資料夾", rdoF.Name)
                        End Try
                        Dim done As Integer = Interlocked.Increment(processedCount)

                        ' by Gemini, 2026/04/02: 更新為 IProgress 且加上簡易模數節流避免平行洗板
                        If progress IsNot Nothing AndAlso done Mod 10 = 0 Then
                            progress.Report(New L3ProgressReport With {.CurrentCount = done, .TotalCount = targetFolderCount,
                                                                       .Message = $"正在平行統計: {done} / {targetFolderCount} 個資料夾..."})
                        End If
                    End Sub)
                If _cancelRequested Then
                    Dbg("GetMailCountAll ① 已取消", $"總資料夾數: {targetFolderCount}") : Return -1
                End If
                Dbg("結束", $"① 平行BFS成功 (RDO): {rootFolder.Name} | total={totalCount} | folders={targetFolderCount}")
                Return totalCount
            Catch ex As System.Exception
                Dbg("GetMailCountAll ① 平行BFS失敗，走循序BFS fallback", $"{rootFolder.Name} | {ex.Message}")
            Finally
                TryMarshalRelease(rdoRoot)
            End Try
        End If

        ' ② BFS 循序累加: GetSubFolderList 展開 + GetMailCount(L3) 逐一加總
        '   支援取消檢查和 progress 進度回報，比平行版保守但穩定
        Try
            Dim targetFolderList As List(Of Outlook.Folder) = GetSubFolderList(rootFolder, includeSubF:=True)
            Dim grandTotal As Long = 0
            Dim swThrottle As New Stopwatch() : swThrottle.Start() ' by Gemini, 2026/04/02: 100ms 節流閥

            For i As Integer = 0 To targetFolderList.Count - 1
                If _cancelRequested Then
                    Dbg("GetMailCountAll ② 被取消", $"已處理 {i}/{targetFolderList.Count}") : Return -1
                End If
                Dim f As Outlook.Folder = targetFolderList(i)
                Dim count As Integer = GetMailCount(f)
                ' GetMailCount 的所有 fallback 都失敗才會到這個 else，記錄但不中止整體加總
                If count >= 0 Then grandTotal += CLng(count) Else Dbg("GetMailCountAll ② 略過失敗資料夾", f.Name)

                ' by Gemini, 2026/04/02: 100ms 節流回報進度，且在此區塊不輸出 Dbg()
                If progress IsNot Nothing AndAlso swThrottle.ElapsedMilliseconds >= 100 Then
                    progress.Report(New L3ProgressReport With {.CurrentCount = i + 1, .TotalCount = targetFolderList.Count,
                                                               .Message = $"正在統計郵件數: {i + 1} / {targetFolderList.Count} 個資料夾..."})
                    swThrottle.Restart()
                End If

                If i Mod 10 = 0 Then Await Task.Yield()
            Next
            Dbg("GetMailCountAll 結束", $"② 循序BFS成功: {rootFolder.Name} | total={grandTotal}")
            Return grandTotal
        Catch ex As System.Exception
            Dbg("GetMailCountAll ② 循序BFS失敗，走遞迴fallback", $"{rootFolder.Name} | {ex.Message}")
        End Try

        ' ③ 遞迴 fallback: GetSubFolderList 本身失敗時的最後保險
        '   無法精確回報進度，但確保加總結果正確
        '   注意: 遞迴呼叫會重新進入本函數，⓪ Redemption 已失敗所以 _rdoSession 仍 Nothing 或故障
        '         ① ② 也已失敗，只會走到 ③ 再次遞迴——理論上 ③ 不會無限展開，因為每層只遞迴直屬子資料夾
        '         todo: 若 ③ 常被觸發，需回頭檢查 GetSubFolderList 失敗的根本原因
        Try
            Dim totalCount As Long = 0
            Dim count As Integer = GetMailCount(rootFolder)     ' 本層 mailcount
            If count >= 0 Then totalCount += count
            Await Task.Yield()
            For Each f As Outlook.Folder In rootFolder.Folders
                Dim subCount As Long = Await GetMailCountAll(f) ' 遞迴，每個直屬子資料夾各自展開
                If subCount >= 0 Then totalCount += subCount
            Next
            Dbg("結束", $"③ 遞迴fallback成功: {rootFolder.Name} | total={totalCount}")
            Return totalCount
        Catch ex As System.Exception
            Dbg("GetMailCountAll ③ 遞迴fallback也失敗", $"{rootFolder.Name} | {ex.Message}")
            Return -1   ' ④ 四層都失敗，回傳 -1 讓 L2 知道這是「讀取失敗」而非「真的是 0 封」
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
        Dbg("開始", folder.Name)
        ' ⓪ Redemption: RDOFolder.Folders.Count
        '   與 OOM folder.Folders.Count 等價，但可在任意執行緒呼叫
        '   2026-03-22 新增
        If _rdo IsNot Nothing Then
            Dim rdoFolder As Redemption.RDOFolder = Nothing
            Try
                rdoFolder = _rdo.GetFolderFromID(folder.EntryID, folder.StoreID)
                Dim count As Integer = rdoFolder.Folders.Count
                ' by Gemini, 2026/04/04: 成功路徑靜默，不輸出 Dbg（Issue 1 高頻去噪）
                Return count
            Catch ex As System.Exception
                Dbg("錯誤路徑", $"GetFolderCount ⓪ RDO: {folder.Name} | {ex.Message}")
            Finally
                TryMarshalRelease(rdoFolder)
            End Try
        End If
        ' ① MAPI: PR_FOLDER_CHILD_COUNT (0x66380003)
        ' 2026/3/20, 奇怪PR_FOLDER_CHILD_COUNT 沒有一次成功過??? 乾脆先拿掉這個try, 省得一直fallback也是浪費開銷
        Try
            Dim count As Integer = folder.Folders.Count
            ' by Gemini, 2026/04/04: 成功路徑靜默，不輸出 Dbg
            Return count
        Catch ex As System.Exception
            Dbg("錯誤路徑", $"GetFolderCount ① OOM: {folder.Name} | {ex.Message}")
        End Try
        Dbg("結束", $"FAIL: {folder.Name}")
        Return -1

    End Function
    Private Async Function GetFolderCountAll(rootFolder As Outlook.Folder, Optional progress As IProgress(Of L3ProgressReport) = Nothing) As Task(Of Integer)
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
        ' 取代: GetTotalFolderCountAsync (快取邏輯移至 L2 呼叫端)
        '
        ' [Redemption說明] 2026-03-22
        '   此函數計算的是整棵子樹的遞迴總數，Redemption 沒有單一 API 可直接取得遞迴資料夾總數
        '    (rdoFolder.Folders.Count 只回傳直屬子資料夾數，與 OOM 相同) 。
        '   因此此函數本身不需要直接加 Redemption 呼叫。
        '   ① BFS 路徑: GetSubFolderList 內部走 OOM folder.Folders 展開，展開後直接 .Count，不需 L3 讀取。
        '   ② 遞迴 fallback: 內部的 rootFolder.Folders.Count 和 ForEach 走 OOM，
        '      若日後改為呼叫 GetFolderCount(L3)，即可自動走 Redemption ⓪ 路徑。
        ' --------------------------------------------------------------
        Dbg("開始", rootFolder.Name)

        ' by Gemini, 2026/04/02: 預跑一次顯示準備中
        progress?.Report(New L3ProgressReport With {.Message = "正在展開資料夾結構...", .IsIndeterminate = True})

        ' 2026/3/24 by Gemini: ⓪ Redemption + 平行處理 (最快路徑)
        '   使用 GetSubFolderList_RDO 取得清單，以 Parallel.ForEach 搭配 Interlocked.Add(rdoF.Folders.Count) 快速加總
        If _rdo IsNot Nothing Then
            Dim rdoRoot As Redemption.RDOFolder = Nothing
            Try
                rdoRoot = _rdo.GetFolderFromID(rootFolder.EntryID, rootFolder.StoreID)
                Dim rdoFolderList As List(Of Redemption.RDOFolder) = GetSubFolderList_RDO(rdoRoot, includeSubF:=True)
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
                            Dbg("GetFolderCountAll ⓪ RDO 略過失敗資料夾", rdoF.Name)
                        End Try

                        Dim done As Integer = Interlocked.Increment(processedCount)
                        ' by Gemini, 2026/04/02: 更新為 IProgress 且加上 100ms 節流，取代原有的 Mod 10
                        If progress IsNot Nothing AndAlso swThrottle.ElapsedMilliseconds >= 100 Then
                            progress.Report(New L3ProgressReport With {.CurrentCount = done, .TotalCount = targetFolderCount,
                                                                       .Message = $"正在統計資料夾樹: {done} / {targetFolderCount}..."})
                            swThrottle.Restart()
                        End If
                    End Sub)
                If _cancelRequested Then
                    Dbg("GetFolderCountAll ⓪ 已取消", "") : Return -1
                End If
                Dbg("結束", $"⓪ RDO平行成功: {rootFolder.Name} | total={totalCount}")
                Return totalCount
            Catch ex As System.Exception
                Dbg("GetFolderCountAll ⓪ RDO平行失敗，走OOM循序fallback", $"{rootFolder.Name} | {ex.Message}")
            Finally
                TryMarshalRelease(rdoRoot)
            End Try
        End If

        ' 2026/3/24 by Gemini: ② OOM + BFS 循序 (無 Redemption 時的最後手段)
        '   必須循序處理 OOM COM 物件以避免 STA 違規
        Try
            Dim allFolders As List(Of Outlook.Folder) = GetSubFolderList(rootFolder, includeSubF:=True, progress:=progress)
            ' by Gemini, 2026/04/02: BFS 展開後回傳數量
            Dim total = allFolders.Count - 1
            progress?.Report(New L3ProgressReport With {
                .CurrentCount = total,
                .TotalCount = total,
                .Message = $"資料夾結構已展開: 共 {total} 個資料夾。"
            })
            Await Task.Yield()
            Dbg("結束", $"② OOM BFS成功: {rootFolder.Name} | total={total}")
            Return total
        Catch ex As System.Exception
            Dbg("GetFolderCountAll ② OOM BFS失敗", $"{rootFolder.Name} | {ex.Message}")
        End Try
        ' ③ 全部失敗
        Return -1

    End Function
    Private Async Function GetFolderSize(folder As Outlook.Folder, Optional progress As IProgress(Of L3ProgressReport) = Nothing) As Task(Of Long)
        ' --------------------------------------------------------------
        ' GetFolderSize v1.5: 讀取單一資料夾本層大小 (bytes)
        ' by Gemini, 2026/04/02: 加入 IProgress 支援以回報分批讀取進度 (100ms 節流)
        ' 2026/3/24 by Gemini: Fallback 鏈重構
        '   ⓪ Redemption : rdoFolder.Fields(PR_MESSAGE_SIZE_EXTENDED) (部分 Exchange 支援，極快)
        '   ① OOM  : folder.GetTable(PR_MESSAGE_SIZE_EXTENDED) + GetArray(1000) (最快安全招式)
        '   ② OOM  : folder.GetTable(PR_MESSAGE_SIZE_EXTENDED) + GetNextRow() (備案)
        '   ③ fail : Return -1
        ' --------------------------------------------------------------
        If _iLikeNoisy Then Dbg("開始", folder.Name)
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
                    Dbg("結束", $"⓪ RDO Fields 成功: {folder.Name} | size={totalSize} | {sw.ElapsedMilliseconds}ms")
                    Return totalSize
                End If
            Catch ex As System.Exception
                Dbg("GetFolderSize ⓪ RDO 失敗，走 OOM GetArray fallback", $"{folder.Name} | {ex.Message}")
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
                If progress IsNot Nothing AndAlso swThrottle.ElapsedMilliseconds >= 100 Then
                    progress.Report(New L3ProgressReport With {.Message = $"正在計算 {folder.Name} 大小: {totalSize / 1024 / 1024:0.0} MB..."})
                    swThrottle.Restart()
                End If
                Await Task.Yield() ' 讓出 UI 避免卡死
            Loop
            sw.Stop()
            Dbg("結束", $"① OOM GetTable.GetArray 成功: {folder.Name} | size={totalSize} | {sw.ElapsedMilliseconds}ms")
            Return totalSize
        Catch ex As System.Exception
            Dbg("GetFolderSize ① OOM GetArray 失敗，走 GetNextRow fallback", $"{folder.Name} | {ex.Message}")
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
            Dbg("結束", $"② OOM GetNextRow 成功: {folder.Name} | size={totalSize} | {sw.ElapsedMilliseconds}ms")
            Return totalSize
        Catch ex As System.Exception
            Dbg("GetFolderSize ② OOM GetNextRow 失敗", $"{folder.Name} | {ex.Message}")
        Finally
            TryMarshalRelease(table2)
        End Try

        sw.Stop()
        Dbg("結束", $"FAIL: {folder.Name} | -1 | {sw.ElapsedMilliseconds}ms")
        Return -1
    End Function
    Private Async Function GetFolderSizeAll(rootFolder As Outlook.Folder, Optional progress As IProgress(Of L3ProgressReport) = Nothing) As Task(Of Long)
        ' --------------------------------------------------------------
        ' GetFolderSizeAll v1.5: 讀取某資料夾及整棵子樹的大小總計 (bytes)
        ' by Gemini, 2026/04/02: 增加 IProgress 支援與 100ms 節流回報
        '
        ' 2026/3/24 by Gemini: 落實新的 Fallback 鏈設計，並修正平行處理的 STA 問題
        '   ⓪ Redemption 平行路徑 (最快):
        '      利用 GetSubFolderList_RDO 一次把該子樹下所有 RDOFolder 拿出來，
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
        Dbg("開始", rootFolder.Name)
        ' 2026/3/24 by Gemini: ⓪ Redemption 平行累加 PR_MESSAGE_SIZE_EXTENDED
        If _rdo IsNot Nothing Then
            Dim rdoRoot As Redemption.RDOFolder = Nothing
            Try
                rdoRoot = _rdo.GetFolderFromID(rootFolder.EntryID, rootFolder.StoreID)
                Dim rdoFolderList As List(Of Redemption.RDOFolder) = GetSubFolderList_RDO(rdoRoot, includeSubF:=True)
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
                            Dbg("GetFolderSizeAll ⓪ RDO 略過讀取失敗的資料夾", rdoF.Name)
                        End Try
                    End Sub)

                If _cancelRequested Then
                    Dbg("GetFolderSizeAll ⓪ 已取消", $"總資料夾數: {rdoFolderList.Count}") : Return -1
                End If
                If validCount = 0 AndAlso rdoFolderList.Count > 0 Then
                    Dbg("GetFolderSizeAll ⓪ RDO 讀取失敗 (無支援的屬性) ", "退回 OOM")
                    Throw New System.Exception("RDO PR_SIZE_EX returned empty for all folders")
                End If
                Dbg("結束", $"⓪ RDO平行成功: {rootFolder.Name} | totalSize={grandTotal} | folders={rdoFolderList.Count}")
                Return grandTotal
            Catch ex As System.Exception
                Dbg("GetFolderSizeAll ⓪ RDO平行失敗，走 OOM 循序 fallback", $"{rootFolder.Name} | {ex.Message}")
            Finally
                TryMarshalRelease(rdoRoot)
            End Try
        End If

        ' 2026/3/24 by Gemini: ① OOM 循序 BFS 累加 (避免 STA 錯誤的保險路徑)
        ' 因為 OOM 的 GetTable() 必須在 UI Thread，我們必須循序 Await 每一層
        Try
            Dim targetFolderList As List(Of Outlook.Folder) = GetSubFolderList(rootFolder, includeSubF:=True)
            Dim grandTotal As Long = 0
            Dim swThrottle As New Stopwatch() : swThrottle.Start() ' by Gemini, 2026/04/02

            For i As Integer = 0 To targetFolderList.Count - 1
                If _cancelRequested Then
                    Dbg("GetFolderSizeAll ① 被取消", $"已處理 {i}/{targetFolderList.Count}") : Return -1
                End If
                Dim f As Outlook.Folder = targetFolderList(i)
                ' by Gemini, 2026/04/02: 傳遞 progress 進去以獲得更細緻的(郵件級別)進度回報
                Dim sz As Long = Await GetFolderSize(f, progress)

                If sz >= 0 Then
                    grandTotal += sz
                Else
                    Dbg("GetFolderSizeAll ① 略過了大小計算失敗的資料夾", f.Name)
                End If

                ' by Gemini, 2026/04/02: 100ms 節流回報資料夾級別進度
                If progress IsNot Nothing AndAlso swThrottle.ElapsedMilliseconds >= 100 Then
                    progress.Report(New L3ProgressReport With {.CurrentCount = i + 1, .TotalCount = targetFolderList.Count,
                                                               .Message = $"正在計算大小: {i + 1} / {targetFolderList.Count} ({f.Name})..."})
                    swThrottle.Restart()
                End If

                ' 避免卡死 UI
                If i Mod 5 = 0 Then Await Task.Yield()
            Next
            Dbg("結束", $"① 循序BFS成功: {rootFolder.Name} | totalSize={grandTotal}")
            Return grandTotal
        Catch ex As System.Exception
            Dbg("GetFolderSizeAll ① 循序BFS失敗，放棄計算", $"{rootFolder.Name} | {ex.Message}")
        End Try

        ' ② 兩層都失敗，回傳 -1 讓呼叫端知道失敗了
        Return -1
    End Function
    Private Async Function GetMailWithAttachment(folder As Outlook.Folder, progress As IProgress(Of L3ProgressReport)) As Task(Of List(Of MailItemInfo))
        ' Phase 1 / L3 純資料層: GetTable + GetArray 批次掃描單一資料夾
        ' 設計: 這裡只專注於透過 MAPI 取回資料，不會做快取判定，也無關大小設定過濾
        Dbg("開始", folder.Name)

        Const PR_MESSAGE_SIZE As String = "http://schemas.microsoft.com/mapi/proptag/0x0E080003"
        Const BATCH_SIZE As Integer = 1000
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

                If progress IsNot Nothing AndAlso swThrottle.ElapsedMilliseconds >= 100 Then
                    progress.Report(New L3ProgressReport With {.Message = $"Phase 1 掃描: {folder.Name} (已找 {result.Count} 封)"})
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
            Dbg("GetMailWithAttachmentL3 Error: ", folder.Name & " — " & ex.Message)
        Finally
            TryMarshalRelease(table)
        End Try
        Dbg("結束", $"找到 {result.Count} 封有附件郵件")
        Return result
    End Function
    Private Function GetAttachFilename(mail As MailItemInfo) As List(Of String)
        ' by Gemini, 2026/04/04: 取得郵件的附件檔名清單 (純 L3 邏輯，不做快取)
        Dim result As New List(Of String)()

        ' ⓪ Redemption 優先: 繞過 OOM 開信的記憶體開銷，直接透過 MAPI Table 抓取檔名
        If _rdo IsNot Nothing Then
            Dim rdoMsg As Redemption.RDOMail = Nothing
            Try
                rdoMsg = TryCast(_rdo.GetMessageFromID(mail.EntryID), Redemption.RDOMail)
                If rdoMsg IsNot Nothing Then
                    For Each att In rdoMsg.Attachments
                        result.Add(att.FileName)
                    Next
                End If
                Return result
            Catch ex As System.Exception
                Dbg("GetAttachmentNamesL3 ⓪ RDO 失敗，走OOM fallback", ex.Message)
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
                    result.Add(att.FileName)
                    TryMarshalRelease(att)
                Next
            End If
        Catch ex As System.Exception
            Dbg("GetAttachmentNamesL3 ① OOM 失敗", ex.Message)
        Finally
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
        Dim fPath As String = folder.FolderPath
        Dim isMail As Boolean
        If _cacheIsMailFolder.TryGetValue(fPath, isMail) Then Return isMail
        ' by Gemini, 2026/04/04: 移除 Dbg("開始")，改為只在非郵件資料夾時記錄（Issue 1 高頻去噪）
        Static allowedTypes As Outlook.OlItemType() = {Outlook.OlItemType.olMailItem, Outlook.OlItemType.olPostItem}
        Try
            Dim itemType As Outlook.OlItemType = folder.DefaultItemType
            isMail = allowedTypes.Contains(itemType)
            _cacheIsMailFolder.TryAdd(fPath, isMail)
            If Not isMail Then Dbg("過濾非郵件資料夾", $"{folder.Name} ({itemType})") ' 只有非郵件時才記錄
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
            Dbg("TryMarshalRelease 異常: ", ex.Message)
        Finally
            obj = Nothing
        End Try
    End Sub
#End Region
#End Region

End Class
