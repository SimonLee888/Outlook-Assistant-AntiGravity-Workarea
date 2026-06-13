Imports System.Collections.Concurrent
Imports System.IO.Hashing
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports Microsoft.Office.Interop.Outlook
' 2026/3/22 正式導入Redemption, 測試logon成功, 傳回數值成功
Imports Redemption
Imports Outlook = Microsoft.Office.Interop.Outlook
'Imports MailKit        ' MailKit is a cross-platform mail client library built on top of MimeKit.

' === 從頭重新設計 Layer3 / Layer2.5 底層計數函數 ===
' 目的: 提供一個純粹的 COM 資料層函數，專注於讀取資料，不做任何流程控制或快取邏輯
'       取代目前散落在各處的 GetMailCountByMAPINew、GetFolderSizeLegacy 等函數，統一為一個簡單的 GetXxxLayer3 函數
' 架構: Layer3 純資料層，Layer2 流程協調層，Layer2.5 快取代理層，Layer1 UI 事件層
'       Layer3 只負責讀取資料夾的本層郵件數 (GetDirectMailCountLayer3)，不遞迴、不展開子資料夾，最小化 COM 呼叫量
'       上層流程 (如 CollectFolderStatsByBFS) 負責決定何時呼叫、如何使用結果、快取管理等
' ==============================================================
' === Layer3 底層 COM 資料層函數群 ===
' 設計原則:
'   1. 每個函數只負責一件事: 讀取單一資料夾或單封郵件的一種屬性
'   2. 不做快取、不做遞迴、不做 BFS 展開——這些全部交給 Layer2 流程協調層
'   3. Fallback 鏈: RDO → MAPI GetArray → OOM最後手段
'                   parallel.foreach → BFS → Recursive，每層不論成功失敗都丟 Debug message
'   4. 失敗統一回傳 -1 (不回 0)，讓 Layer2 能區分「真的是 0」或「讀取失敗」
'   5. 在 Finally 中使用 TryMarshalRelease() 統一釋放所有 COM 物件，確保 RCW 不殘留
' ==============================================================
Partial Class Form1

#Region "■ 01 全域宣告"
    Private WithEvents _olApp As Outlook.Application = Nothing
    Private _olNS As Outlook.NameSpace = Nothing
    Private _pstStoreList As List(Of Outlook.Store) = Nothing
    Private _rdo As Redemption.RDOSession = Nothing ' _rdoSession 就等同是outlook.namespace 的意思, 就是Redemption的MAPI session
    ' 2026-03-22 新增: 用於測試 Redemption.dll 整合 (注意: session.MAPIOBJECT 必須在 Outlook MAPI 連線建立後才能設定 (Form1_Load 尾端)
    '------------------------------------------------------------------------------------------------
    ' Outlook 物件(OOM)	    Redemption 物件 (RDO)     說明
    '------------------------------------------------------------------------------------------------
    ' Outlook.Application	Redemption 本體	        Redemption 是底層 MAPI 封裝，它不負責 UI 或視窗管理。
    ' Outlook.NameSpace	    Redemption.RDOSession	最接近。 負責管理登入、StoreID、PST 檔案庫與全域設定。
    ' Outlook.Folder	    Redemption.RDOFolder	對應資料夾層級。
    ' Outlook.MailItem	    Redemption.RDOMail	    對應單封郵件層級。
    ' Outlook.Store	        Redemption.RDOStore	    對應 PST 或 Exchange 帳戶。

    Private Shared _cacheIsMailFolder As New ConcurrentDictionary(Of String, Boolean)       ' 資料夾是否為郵件類型
    Private Shared _cacheMailCount As New ConcurrentDictionary(Of String, Long)             ' 自身資料夾的郵件個數
    Private Shared _cacheMailCountAll As New ConcurrentDictionary(Of String, Long)          ' 整支子樹的所有郵件總數
    Private Shared _cacheFolderCount As New ConcurrentDictionary(Of String, Long)           ' 自身資料夾的子目錄個數
    Private Shared _cacheFolderCountAll As New ConcurrentDictionary(Of String, Long)        ' 整支子樹的所有子目錄總數
    Private Shared _cacheFolderSize As New ConcurrentDictionary(Of String, Long)            ' 自身資料夾的郵件大小加總
    Private Shared _cacheFolderSizeAll As New ConcurrentDictionary(Of String, Long)         ' 整支子樹的所有子目錄郵件大小加總

    Private Shared _cacheFolderTree As New ConcurrentDictionary(Of String, List(Of Folder))     ' GetSortedSubFolders() 已排序的子資料夾清單
    Private Shared _cacheAttachMailList As New ConcurrentDictionary(Of String, FolderCacheTab3) ' 包含附件的郵件預掃描結果 (速度很快, 不用存入SSD?)
    Private Shared _cacheAttachFilename As New ConcurrentDictionary(Of String, List(Of String)) ' 所有附件檔名清單
    Private Shared _cacheYearCounts As New ConcurrentDictionary(Of String, ConcurrentDictionary(Of Integer, Integer))
    Private Shared _cacheMonthCounts As New ConcurrentDictionary(Of String, ConcurrentDictionary(Of Integer, Integer))

    Private Shared _cacheSubTreeList As New ConcurrentDictionary(Of String, List(Of (folder As Outlook.Folder, fPath As String)))               ' GetSubtreeToList() 的樹狀展開平坦化清單 (by Gemini, 2026/04/10: 帶路徑優化)
    Private Shared _cacheFolderIDs As New ConcurrentDictionary(Of String, (eid As String, sid As String, isMail As Boolean, hasCh As Boolean))  ' by Gemini, 2026/04/10: 專門儲存資料夾的身分標識與屬性標籤，用以橋接 Folder 物件與 SQLite 持久化
    Private Shared _cacheBasicMailInfo As New ConcurrentDictionary(Of String, (Mails As List(Of (Mail As MailItemInfo, Topic As String)), Snap As Long)) ' by Gemini, 2026/04/20: 專用於 Tab4 的郵件預掃描快取，Key 是資料夾路徑，Value 是該資料夾下所有郵件的基本資訊列表 (不帶 COM 物件) 與當下的 PR_CONTENT_COUNT 快照，用於快速顯示搜尋結果與驗證快取有效性

    Private _lv4BodyCache As New ConcurrentDictionary(Of String, String)    ' by Gemini 3 Flash, 2026/04/26: Tab4 相似度計算用的 Body 快取 (session 級，避免重複讀取 Outlook mailitem.Body)
    Private _isForceRefreshing As Boolean = False                           ' ✅ 2026/05/31 新增：F5 強制更新旗標，指示底層完全繞過 SSD 快取

    Private Structure FolderSortInfo
        ' by Gemini, 2026/03/29: 用於 GetSortedSubFolders 排序優化，減少 COM 屬性讀取次數 (O(N) vs O(N log N))
        Dim FolderObj As Folder
        Dim Name As String
        Dim HasChinese As Boolean
    End Structure
    Friend Structure MailItemInfo
        ' 候選郵件的純資料結構 (不帶 COM 物件，不受 GC 影響)
        Dim EntryID As String
        Dim Subject As String
        Dim Size As Long
        Dim RcvTime As DateTime
        Dim SenderName As String
        Dim AttachCount As Integer
        Dim FolderPath As String    ' by Gemini 3 Flash, 2026/04/19: 加入路徑資訊，用於 ListView ToolTip 顯示
        Dim MsgIDhash As String     ' 2026/05/06 by Claude: Tab5 去重 (PR_INTERNET_MESSAGE_ID)
        Dim SenderEmail As String   ' 2026/05/06 by Claude: Tab5 去重 (PR_SENDER_EMAIL_ADDRESS)
    End Structure
    Private Structure FolderCacheTab3
        Dim AttachMailList As List(Of MailItemInfo) ' 所有 hasAttachment 候選 (無大小篩選)
        Dim ItemCountSnap As Long                   ' 快取當下的 PR_CONTENT_COUNT，失效偵測用
    End Structure

    ' 2026/06/12 by Simon/Claude: Compiled Regex，程式啟動時編譯一次，後續呼叫零額外開銷
    ' Pattern 說明：^ 錨定開頭；[：:] 同吃半形/全形冒號；外層 + 一次處理所有巢狀前綴
    Private Shared ReadOnly _subjectPrefixRe As New Regex(
        "^(?:(?:RE|FW|FWD|AW|WG|VS|Rép|TR|回覆|回信|轉寄|轉發|回复|答复|转发|返信|転送|답장|회신|전달)\s*[：:]\s*)+",
        RegexOptions.IgnoreCase Or RegexOptions.Compiled)
#End Region

#Region "■ 10 底層 COM 函數群 (新設計，現役主力) "
#Region "  ├ 全域初始化 & 載入釋放函數"
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
        '   視窗 class    = TEULAForm (Delphi VCL 表單)，title = "Outlook Redemption"
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
    Private Sub TryMarshalRelease(ByRef obj As Object)
        Try
            If obj IsNot Nothing AndAlso Marshal.IsComObject(obj) Then Marshal.ReleaseComObject(obj)
        Catch ex As System.Exception
            _dbg("TryMarshalRelease 異常: ", ex.Message)
        Finally
            obj = Nothing
        End Try
    End Sub
#End Region
#Region "  ├ Layer2 UI 流程輔助"
    Private Function GetSortedStores(space As Outlook.NameSpace) As List(Of Outlook.Store)
        ' ==========================================
        ' 取得排序後的 NameSpace 下所有Outlook.Store
        ' 包含目前config內的所有帳號和所有開啟的PST檔
        '
        ' ⚠️ 注意: 不在這裡 ReleaseComObject(space)
        '     space 就是外層的 objNameSpace，釋放後其他地方 (Tab2/Tab3 等) 再用 objNameSpace 會觸發 RCW 已釋放的例外
        '     objNameSpace 的生命週期就只由 Form1_FormClosing 統一管理
        ' ==========================================
        _dbg(" ├ 開始", space.CurrentProfileName) ' by Gemini, 2026/04/10: 調整縮排層級為 Level 2 (由 InitOutlookNamespace 呼叫)

        ' 遍歷所有Outlook.Store並添加到列表中, 使用LINQ擴充方法就夠快了, 不再使用非同步或Parallel.Foreach了
        Dim stores As List(Of Outlook.Store) = space.Stores.Cast(Of Outlook.Store)().ToList()
        stores = stores.OrderBy(Function(st) If(TextHasChineseChar(st.DisplayName), 1, 0)).ThenBy(Function(st) st.DisplayName).ToList() ' 使用 LINQ 排序Outlook.Store

        _dbg(" ├ 結束", $"Profile={space.CurrentProfileName} | 庫數量: {stores.Count}") ' by Gemini, 2026/04/10
        Return stores

    End Function
    Private Function GetSortedSubFolders(pFolder As Folder, Optional fPath As String = "", Optional forceRefresh As Boolean = False) As List(Of Folder)
        ' ==========================================
        ' 取得引數pFolder下的所有subFolders並排序後傳回
        ' 優化紀錄: 2026/03/29 by Gemini 3.1 Pro
        ' 1. 加入 Layer3 過濾: 只保留郵件目錄 (olMailItem)，排除行事曆/聯絡人等
        ' 2. 單次屬性讀取: 先快取 Name 後排序，避開 LINQ 重複打 COM 的 N log N 效能陷阱
        ' 2026/04/15: 支援傳入 fPath，並透過字串拼接子路徑，減少內部 COM 呼叫
        ' 2026/04/16: 整合記憶體與 SSD 快取機制 (修復損補遺失) by Gemini 3.0 flash
        ' ==========================================
        fPath = SafeGetPath(pFolder, fPath)
        Dim fName As String = ExtractFolderName(fPath)
        If _iLikeNoisy Then _dbg(" ├ 開始", fName)

        ' ① 記憶體快取檢查: 命中則直接回傳，0ms
        ' by Gemini 3.0 flash, 2026/04/17: 引入鍵值分支，避免全顯/過濾模式交錯污染
        Dim cacheKey As String = fPath & "|" & _showAllFolders
        Dim cachedFolders As List(Of Folder) = Nothing
        If _cacheFolderTree.TryGetValue(cacheKey, cachedFolders) Then Return cachedFolders

        ' ② SSD / DB 讀取分支 (Lazy Load): TreeView 展開時的主要加速點
        ' ✅ 2026/5/31 by Gemini/Simon: 加入 forceRefresh 引數判斷是否要強制讀取COM，避免在需要最新資料的情況下誤用過期快取
        If _db IsNot Nothing AndAlso Not forceRefresh Then
            Dim dbIDs = DbGetOrderedSubFolderIDs(fPath, _showAllFolders)
            If dbIDs IsNot Nothing Then
                ' 預分配容量為 512，足以涵蓋多數資料夾搜尋結果，減少陣列頻繁 Resize 開銷 (by Gemini 3 Flash, 2026/05/04)
                Dim dbResults As New List(Of Folder)(512)
                For Each row In dbIDs
                    Try
                        ' DbGetSubFolderIDList 回傳的是 (eid, sid, path) 的具名 Tuple 列表 by Gemini 3.0 flash, 2026/04/16
                        Dim f = TryCast(_olNS.GetFolderFromID(row.eid, row.sid), Folder)
                        If f IsNot Nothing Then dbResults.Add(f)
                    Catch
                    End Try
                Next
                If dbResults.Count > 0 Then
                    _cacheFolderTree(cacheKey) = dbResults
                    If _iLikeNoisy Then _dbg("    ├ SSD Hit", $"{fName}: 已從資料庫載入 {dbResults.Count} 個子目錄")
                    Return dbResults
                End If
            End If
        End If

        ' ③ 傳統 OOM 分支 (Fallback): 快取未命中時才打 COM 掃描
        ' 收集子資料夾並快取屬性 (減少 COM 屬性重複呼叫)
        ' 預分配容量為 512，優化資料夾排序時的暫存資訊處理 (by Gemini 3 Flash, 2026/05/04)
        Dim infoList As New List(Of FolderSortInfo)(512)
        Dim subs As Folders = pFolder.Folders
        Try
            For Each subF As Folder In subs
                ' 修復過濾：若未勾選「顯示全部」，且該資料夾非郵件類，則跳過 (by Gemini 3.0 flash, 2026/04/17)
                ' 利用現有 fPath 拼接路徑傳入 IsMailFolder，實現零額外 COM 屬性讀取
                Dim sName As String = subF.Name
                Dim childPath As String = fPath & "\" & sName
                If Not _showAllFolders AndAlso Not IsMailFolder(subF, childPath) Then Continue For

                infoList.Add(New FolderSortInfo With {.FolderObj = subF, .Name = sName, .HasChinese = TextHasChineseChar(sName)})
                ' 這裡 subF 被加入 infoList 成為物件清單，所以不能在這裡 TryRelease 它

                ' 2026/6/2: 再次修正F5 強制刷新的總數讀取不正確
                ' 🔽🔽🔽 【修復點 2】順手把展開的資料夾也註冊身分證 🔽🔽🔽
                'Try
                '    _cacheFolderIDs.TryAdd(childPath, (subF.EntryID, subF.StoreID, IsMailFolder(subF, childPath), TextHasChineseChar(sName)))
                'Catch : End Try
                ' 🔼🔼🔼 🔼🔼🔼
            Next
        Finally
            TryMarshalRelease(subs) ' 存到變數後可以TryRelease subs，避免後續 COM 呼叫時 RCW 已釋放的例外
        End Try

        ' 純記憶體排序: 完全不觸發 COM 呼叫
        Dim sortedFolders = infoList.OrderBy(Function(i) If(i.HasChinese, 1, 0)).
                                     ThenBy(Function(i) i.Name, StringComparer.OrdinalIgnoreCase).
                                     Select(Function(i) i.FolderObj).ToList()
        ' 2026/4/7 進一步優化 by Gemini: 加入 StringComparer.OrdinalIgnoreCase 略過語系分析，爆發性提速

        _cacheFolderTree(cacheKey) = sortedFolders
        If _iLikeNoisy Then _dbg(" ├ 結束", $"{fName} (BFS) | 子資料夾數: {sortedFolders.Count}")
        Return sortedFolders

    End Function
    Private Sub LoadStoreToTreeView(storeList As List(Of Outlook.Store), tv As SimTree)
        ' ===========================================================
        ' 將所有 Outlook.Store 的根資料夾載入指定的 TreeView 控制項
        ' ===========================================================
        ' 設計原則:
        ' 1. 只在初始化階段載入 Store 的根資料夾，使用佔位節點 ":::" 來騙過 TreeView, 讓它顯示展開 "+" 號，實現真正的 Lazy Load
        ' 2. 遍歷 storeList 並創建節點, 用一個 TreeNode 的 List 來暫存所有要添加的節點
        ' 3. 優化節點添加邏輯，新增時加進List而不是直接加到Treeview.Nodes, 最後再一次addRange() 到 TreeView，減少 UI 更新次數提升效能
        '
        ' 2026/04/10 by Gemini: 調整 Debug 縮排層級，並加入開始/結束訊息，方便追蹤載入事件的觸發與完成情況
        ' 2024/5/20: 昨天才說不會更快了, 今天改用Nodes.AddRange(), 又更快了一點, 連BeginUpdate/EndUpdate都不需要了
        ' 2026/4/?? by Claude: PST root pFolder 幾乎 100% 都有子資料夾，這個假設安全；
        '   就算 PST 真的空了，展開時 LoadSubFolderToTreeView 清除 ":::" 後不加任何子節點，節點就會自動收起 "+" 號，行為正確
        '
        ' 2026/05/04 by Gemini 3 Flash, 減少 TreeView 節點批次添加時 List() 的 Resize 次數 (預分配容量先訂16, 通常不會在一個資料夾內還有超過16個子資料夾)
        ' 2026/5/20 by simon: 在node.Name屬性加上值, 以便後續可以使用TreeNode.Find()
        ' ===========================================================
        _dbg(" ├ 開始", tv.Name)

        Dim nodeList As New List(Of TreeNode)(16)
        For Each store In storeList
            Dim root As Folder = store.GetRootFolder
            Dim node As New TreeNode(root.Name) With {.Tag = root, .Name = root.FullFolderPath}
            node.Nodes.Add(":::")   ' ✅ PST root內必定有資料夾, 所以無條件加佔位節點，省掉判斷 root.Folders.Count 這一次多餘的 COM 往返
            nodeList.Add(node)      ' 先加進List在記憶體中快速操作, 而不是直接加到Treeview.Nodes
        Next

        tv.Nodes.AddRange(nodeList.ToArray()) ' 將所有組裝好的節點一次性添加到 tv.Nodes
        _dbg(" ├ 結束", $"{tv.Name} 共 {nodeList.Count} 個 Store")

    End Sub
    Private Sub LoadSubFolderToTreeView(sender As Object, e As TreeViewCancelEventArgs)
        ' ===========================================================
        ' TreeView 的 BeforeExpand 事件處理器，負責在使用者展開節點時動態載入子資料夾
        ' ===========================================================
        ' 設計原則:
        ' 1. 只在節點第一次展開時載入子資料夾，使用佔位節點 ":::" 判斷是否為第一次展開，載入後立即清除佔位節點
        ' 2. 使用 GetSortedSubFolders 取得已排序的子資料夾清單，並根據是否有子資料夾決定是否添加佔位節點
        ' 3. 遍歷 nodeList 並創建節點, 用一個 TreeNode 的 List 來暫存所有要添加的節點
        ' 4. 新增時加進List而不是直接加到Treeview.Nodes, 最後再一次addRange() 到 TreeView，減少 UI 更新次數提升效能
        '
        ' 2026/04/10 by Gemini: 調整 Debug 縮排層級，並加入開始/結束訊息，方便追蹤展開事件的觸發與完成情況
        ' ✅ 2026/04/20 by Gemini 2.0 Flash: 若 Tab4 處於搜尋結果顯示模式，則不執行自動資料夾載入
        ' 2024/5/17重寫，優化資料夾載入邏輯，加入多層級展開的快取機制，並修復 Tab4 搜尋結果模式下的誤觸發問題
        ' 2024/5/19試過Task.Run(), Parallel.Foreach跟LINQ擴充方法了, 都沒有比較快, 別再試了, 就算virtual mode也沒有比我現在的lazy load還快
        ' 2024/5/20昨天才說不會更快了, 今天改用Nodes.AddRange(), 又更快了一點, 連BeginUpdate/EndUpdate都不需要了
        '
        ' 2026/4/7 by Gemini, 光速版子資料夾加號預測 HasSubFoldersFast() (專為 TreeView 展開設計)
        ' 2026/05/04 by Gemini 3 Flash, 減少 TreeView 節點批次添加時 List() 的 Resize 次數 (預分配容量先訂16, 通常不會在一個資料夾內還有超過16個子資料夾)
        ' 2026/5/20 by simon: 在node.Name屬性加上值, 以便後續可以使用TreeNode.Find()
        ' ✅ 2026/5/31 by Gemini/Simon: 加入 forceRefresh 引數判斷是否要強制讀取COM
        ' ===========================================================

        _dbg(" ├ 開始", sender.Name)

        Dim selectedNode As TreeNode = e.Node                   ' 取得點選的node
        Dim selectedFolder As Folder = selectedNode.Tag         ' 取得點選的資料夾
        Dim sortedFolders = GetSortedSubFolders(
            selectedFolder, forceRefresh:=_isForceRefreshing)   ' 取得所有子資料夾並排序 

        If selectedNode.Nodes.Count = 1 AndAlso selectedNode.FirstNode.Text = ":::" Then
            selectedNode.Nodes.Clear()  '清除原本暫代的假node ":::"

            Dim nodeList As New List(Of TreeNode)(16)
            For Each folder As Folder In sortedFolders
                Dim node As New TreeNode(folder.Name) With {.Tag = folder, .Name = folder.FullFolderPath}
                Try
                    'If GetFolderCount(pFolder) > 0 Then node.Nodes.Add(":::")
                    If HasSubFoldersFast(folder) Then node.Nodes.Add(":::") ' 2026/4/7 by Gemini, 光速版子資料夾加號預測 (專為 TreeView 展開設計)
                Catch ex As System.Exception : End Try
                nodeList.Add(node)  ' 先加進List在記憶體中快速操作, 而不是直接加到Treeview.Nodes
            Next
            selectedNode.Nodes.AddRange(nodeList.ToArray()) ' 將所有節點一次性添加到 selectedNode.Nodes
        End If
        _dbg(" ├ 結束", $"{selectedFolder.Name} 展開 {sortedFolders.Count} 個子資料夾")

    End Sub
    Private Async Function GetUniqueFolderList(selectedNodes As List(Of TreeNode), includeSub As Boolean, cToken As CancellationToken, Optional progress As IProgress(Of ProgressReport) = Nothing) As Task(Of List(Of (folder As Folder, fPath As String)))
        ''' <summary>
        ''' 共用邏輯：將多個 TreeNode 轉換為無重複的實體資料夾清單
        ''' 2026/04/16 by Gemini: 升級回傳 Tuple (Folder, fPath)，消除呼叫端對 COM .FolderPath 的一次性集體讀取
        ''' </summary>
        _dbg(" ├ 開始")
        ' 預分配容量為 512，優化多選資料夾後的路徑合併清單處理 (by Gemini 3 Flash, 2026/05/04)
        Dim fList As New List(Of (folder As Folder, fPath As String))(512)
        Dim addedPaths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        For Each node As TreeNode In selectedNodes
            cToken.ThrowIfCancellationRequested()

            Dim rootF = TryCast(node.Tag, Folder)
            If rootF Is Nothing Then Continue For

            ' GetSubtreeToList 回傳的 subF 現在是 (Folder, fPath) Tuple
            ' 2026/04/17 by Claude: 改呼叫 GetSubtreeToList (L2.5)，原 GetSubtreeToList 已改名為 L3
            For Each subF In Await GetSubtreeToList(rootF, includeSub, progress:=progress, cToken:=cToken)
                ' ✅ 直接讀取 subF.fPath (Tuple 屬性)，再也不用打 COM!
                If addedPaths.Add(subF.fPath) Then fList.Add(subF)
            Next
            Await Task.Yield()
        Next
        Return fList
    End Function
    Private Async Function GetSubtreeToList(rootFolder As Folder, includeSubF As Boolean, Optional progress As IProgress(Of ProgressReport) = Nothing, Optional cToken As CancellationToken = Nothing) As Task(Of List(Of (folder As Folder, fPath As String)))
        ' ---------------------------------------------------------------
        ' GetSubtreeToList — 整棵子資料夾清單 (Layer2.5 快取代理)
        ' 2026/04/17 by Claude: 從 GetSubtreeToList 拆出快取邏輯
        '   原來的快取邏輯混在 BFS 函數裡，現在統一到此 L2.5 層
        '   GetSubtreeToListL3 (原 GetSubtreeToList) 只剩純 BFS COM 掃描
        ' 呼叫順序: ① 記憶體命中 → ② DB lazy load → ③ Layer3 GetSubtreeToListL3
        ' includeSubF=False 時無需快取，直接呼叫 L3 回傳單節點清單
        ' ---------------------------------------------------------------
        Dim rootPath As String = SafeGetPath(rootFolder)

        If Not includeSubF Then Return Await GetSubtreeToListL3(rootFolder, False, progress, cToken:=cToken) ' 單節點不快取

        ' 2026/04/17 by Gemini 3.0 flash: 引入鍵值分支，避免全顯/過濾模式交錯污染快取
        Dim cacheKey As String = rootPath & "|" & _showAllFolders

        Dim cachedList As List(Of (folder As Folder, fPath As String)) = Nothing
        If _cacheSubTreeList.TryGetValue(cacheKey, cachedList) Then              ' ① 記憶體命中
            _dbg(" ├ 結束", $"{ExtractFolderName(rootPath)} (Cache Hit) | 資料夾總計: {cachedList.Count}")
            Return cachedList
        End If

        ' ② DB lazy load: 利用 LIKE 一取回整棵樹的 ID 並重建物件
        ' 注意: DB 存放的是 (EntryID, StoreID, FolderPath)，我們在這裡重建 Tuple
        Dim dbIDs = DbGetSubFolderIDList(rootPath, _showAllFolders)                ' ② DB lazy load
        If dbIDs IsNot Nothing Then
            ' 預分配容量為 512，優化從 DB 載入資料夾子樹時的處理速度 (by Gemini 3 Flash, 2026/05/04)
            Dim dbResults As New List(Of (folder As Folder, fPath As String))(512)
            For Each row In dbIDs
                Try
                    ' DbGetSubFolderIDList 回傳的是 (eid, sid, path) 的具名 Tuple 列表 by Gemini 3.0 flash, 2026/04/16
                    Dim f = TryCast(_olNS.GetFolderFromID(row.eid, row.sid), Folder)
                    If f IsNot Nothing Then dbResults.Add((Folder:=f, fPath:=row.path))
                Catch
                End Try
            Next
            If dbResults.Count > 0 Then
                _cacheSubTreeList(cacheKey) = dbResults
                If _iLikeNoisy Then _dbg("    ├ SSD Hit (Tree)", $"{ExtractFolderName(rootPath)}: 已從資料庫載入 {dbResults.Count} 個子目錄")
                Return dbResults
            End If
        End If

        ' ③ Layer3 BFS COM 掃描；OCE re-throw，快取寫入在 L3 完成後由 L3 自行負責
        Return Await GetSubtreeToListL3(rootFolder, True, progress, cToken:=cToken)

    End Function
#End Region
#Region "  ├ Layer2.5 快取存取點 (Cache Proxy Layer)"
    ' 2026/03/27 by Gemini: 新增 Layer2.5 快取存取點 (Cache Proxy Layer)，保護 Layer3 不被頻繁呼叫
    ' ---------------------------------------------------------------
    '   - GetMailCount(pFolder)             ' 單一資料夾郵件數，有 DB lazy + snapshot 驗證
    '   - GetFolderCount(pFolder)           ' 單一資料夾子資料夾數，有 DB lazy + snapshot 驗證
    '   - GetMailCountAllAsync(pFolder)     ' 整棵子樹郵件總數，有 DB lazy
    '   - GetFolderCountAllAsync(pFolder)   ' 整棵子樹資料夾總數，有 DB lazy
    '   - GetFolderSizeAsync(pFolder)       ' 單一資料夾大小，有 DB lazy
    '   - GetFolderSizeAllAsync(pFolder)    ' 整棵子樹大小，有 DB lazy
    '   - GetAttachMailList(pFolder)        ' Tab3 Phase1，有 DB lazy (attach_maillist)
    '   - GetAttachFilename(mail)           ' Tab3 Phase2，有 DB lazy (attach_filenames)
    '   - GetSubtreeToList(rootFolder)      ' 整棵子資料夾清單，有 DB lazy (2026/04/17)
    '   - GetYearCountsForFolder(sFolder, fPath:=fPath) ' 單一資料夾年份分佈，有 DB lazy (2026/04/17)
    '   - GetMonthCountsForYear(sFolder, year)   ' 單一資料夾月份分佈，有 DB lazy + 提前過濾 (2026/04/17)
    ' ---------------------------------------------------------------
    ' 2026/04/07: Phase 2 — 在記憶體 miss 時加入 SQLite lazy SELECT，命中後一次填滿所有欄位
    '             寫入仍由 SaveCachesToDB (SaveCache 按鈕) 批次處理，本層不做即時寫入
    '
    ' 2026/4/7 by Claude, 加入「持久化快取」存入SSD:
    '             讀資料時若快取不存在就先去撈SSD, 有就放進快取, 沒有才算cache miss去COM讀取
    ' ---------------------------------------------------------------
    ' 呼叫順序 (每個 Layer2.5 函數):
    '   ① 記憶體命中 → 直接回傳 (最快，0 COM call) 
    '   ② DB 命中 + snapshot 驗證通過 → 填滿記憶體快取 → 回傳 (快，0 COM call) 
    '   ③ DB miss 或 snapshot 不符 → 呼叫 Layer3 → 填入記憶體快取 → 回傳 (慢，有 COM call) 
    '
    ' ---------------------------------------------------------------
    Private Function GetMailCount(folder As Folder, Optional fPath As String = "") As Long
        ' ---------------------------------------------------------------
        ' GetMailCount — 單一資料夾本層郵件數 (PR_CONTENT_COUNT)
        ' 2026/4/7 by Claude, 加入Database lazy load, 讀資料順序:
        '   ① 記憶體命中
        '   ② DB lazy load：命中且 mc 有效且 snapshot 吻合 → 一次填滿所有欄位
        '   ③ fallback: Layer3 呼叫
        ' 2026/04/15 by Gemini 3.1 Pro, 加入 optional fPath 參數，若有傳入則可省去 pFolder.FolderPath 1ms 耗時
        ' 2026/05/09 by Gemini 3.1 Pro: 重構成使用 SafeGetDbRow，回傳統一改成 Long 共用函式
        ' ---------------------------------------------------------------
        fPath = SafeGetPath(folder, fPath)
        Dim count As Long : If _cacheMailCount.TryGetValue(fPath, count) Then Return count  ' ① 記憶體命中

        ' ② DB lazy load：命中且 mc 有效且 snapshot 吻合 → 一次填滿所有欄位
        Dim row = SafeGetDbRow(folder, fPath)
        If row IsNot Nothing AndAlso row.mc >= 0 Then Return row.mc

        ' ③ fallback: Layer3 呼叫
        count = GetMailCountL3(folder, fPath:=fPath)
        _cacheMailCount.TryAdd(fPath, count)
        Return count

    End Function
    Private Function GetFolderCount(folder As Folder, Optional fPath As String = "") As Long
        ' ---------------------------------------------------------------
        ' GetFolderCount — 單一資料夾直屬子資料夾數
        ' 2026/4/7 by Claude, 加入Database lazy load, 讀資料順序:
        '   ① 記憶體命中
        '   ② DB lazy load：命中且 fc 有效且 snapshot 吻合 → 一次填滿所有欄位
        '   ③ fallback: Layer3 呼叫
        ' 2026/05/09 by Gemini 3.1 Pro: 重構成使用 SafeGetDbRow，回傳統一改成 Long 共用函式
        ' ---------------------------------------------------------------
        fPath = SafeGetPath(folder, fPath)
        Dim count As Long : If _cacheFolderCount.TryGetValue(fPath, count) Then Return count  ' ① 記憶體命中

        ' ② DB lazy load (fc 欄位) 
        Dim row = SafeGetDbRow(folder, fPath)
        If row IsNot Nothing AndAlso row.fc >= 0 Then Return row.fc

        ' ③ fallback: Layer3 呼叫
        count = GetFolderCountL3(folder, fPath:=fPath)
        _cacheFolderCount.TryAdd(fPath, count)
        Return count

    End Function
    Private Async Function GetFolderSizeAsync(folder As Folder, Optional fPath As String = "", Optional cToken As CancellationToken = Nothing) As Task(Of Long)
        ' ---------------------------------------------------------------
        ' GetFolderSizeAsync — 單一資料夾本層大小 (GetTable 加總)
        ' 2026/3/29 by Gemini: Layer2.5 快取代理層 - 取得單一資料夾本層的大小 (含快取機制)
        ' 2026/4/7 by Claude, 加入Database lazy load, 讀資料順序:
        '   ① 記憶體命中
        '   ② DB lazy load：命中且 fs 有效且 snapshot 吻合 → 一次填滿所有欄位
        '   ③ fallback: Layer3 呼叫
        ' 2026/05/09 by Gemini 3.1 Pro: 重構成使用 SafeGetDbRow，回傳統一改成 Long 共用函式
        ' ---------------------------------------------------------------
        fPath = SafeGetPath(folder, fPath)
        Dim size As Long : If _cacheFolderSize.TryGetValue(fPath, size) Then Return size  ' ① 記憶體命中

        ' ② DB lazy load (fs 欄位) 
        Dim row = SafeGetDbRow(folder, fPath)
        If row IsNot Nothing AndAlso row.fs >= 0 Then Return row.fs

        ' ③ Layer3
        size = Await GetFolderSizeL3(folder, fPath:=fPath, cToken:=cToken)
        If size >= 0 Then _cacheFolderSize.TryAdd(fPath, size)
        Return size

    End Function
    Private Async Function GetFolderSizeAllAsync(folder As Folder, Optional fPath As String = "", Optional cToken As CancellationToken = Nothing) As Task(Of Long)
        ' ---------------------------------------------------------------
        ' GetFolderSizeAllAsync — 整棵子樹大小總計
        ' 2026/3/29 by Gemini: Layer2.5 快取代理層 - 取得整棵子樹的大小總計 (含快取機制)
        ' 2026/4/7 by Claude, 加入Database lazy load, 讀資料順序:
        '   ① 記憶體命中
        '   ② DB lazy load：命中且 fsa 有效且 snapshot 吻合 → 一次填滿所有欄位
        '   ③ fallback: Layer3 呼叫
        ' 2026/04/15 by Claude: 加入 cToken 參數，取代 _cancelRequested 旗標
        ' 2026/05/09 by Gemini 3.1 Pro: 重構成使用 SafeGetDbRow，回傳統一改成 Long 共用函式
        ' ---------------------------------------------------------------
        fPath = SafeGetPath(folder, fPath)
        If _iLikeNoisy Then _dbg(" ├ 開始", ExtractFolderName(fPath))
        Dim size As Long : If _cacheFolderSizeAll.TryGetValue(fPath, size) Then Return size  ' ① 記憶體命中

        ' ② DB lazy load (fsa 欄位) 
        Dim row = SafeGetDbRow(folder, fPath)
        If row IsNot Nothing AndAlso row.fsa >= 0 Then Return row.fsa

        ' ③ fallback: Layer3 呼叫
        size = Await GetFolderSizeAllL3(folder, cToken:=cToken)
        If size >= 0 Then _cacheFolderSizeAll.TryAdd(fPath, size)  ' 2026/04/15: 改用 cToken 判斷
        Return size

    End Function
    Private Async Function GetYearCountsForFolder(folder As Folder, fPath As String, cToken As CancellationToken) As Task(Of ConcurrentDictionary(Of Integer, Integer))
        ' ---------------------------------------------------------------
        ' GetYearCountsForFolder — 單一資料夾年份郵件分佈 (Layer2.5 快取代理)
        ' 2026/04/17 by Claude: 從 CollectYearCounts (L2) 拆出，對齊其他 Layer2.5 快取函數架構
        ' 呼叫順序: ① 記憶體命中 → ② DB lazy load → ③ Layer3 GetYearCountsForFolderL3
        ' OCE 不在此攔截，直接 re-throw 讓 CollectYearCounts (L2) 的 Catch OCE 接住
        ' ---------------------------------------------------------------
        Dim value As ConcurrentDictionary(Of Integer, Integer) = Nothing
        If _cacheYearCounts.TryGetValue(fPath, value) Then     ' ① 記憶體命中
            If _iLikeNoisy Then _dbg("    ├ Cache Hit: ", ExtractFolderName(fPath))
            Return value
        End If

        Dim dbResult = DbGetYearCountsForFolder(fPath)          ' ② DB lazy load
        If dbResult IsNot Nothing Then
            _cacheYearCounts(fPath) = dbResult
            If _iLikeNoisy Then _dbg("    ├ DB Hit: ", ExtractFolderName(fPath))
            Return dbResult
        End If

        If _iLikeNoisy Then _dbg("    ├ Cache miss: ", ExtractFolderName(fPath))
        Dim folderResult = Await GetYearCountsForFolderL3(folder, fPath:=fPath, cToken:=cToken) ' ③ Layer3 COM；OCE re-throw 至 L2
        _cacheYearCounts(fPath) = folderResult                  ' ✅ OCE 時走不到此行，快取僅在完整計算後寫入
        Return folderResult

    End Function
    Private Async Function GetMonthCountsForYear(folder As Folder, year As Integer, fPath As String, cToken As CancellationToken) As Task(Of ConcurrentDictionary(Of Integer, Integer))
        ' ---------------------------------------------------------------
        ' GetMonthCountsForYear — 單一資料夾指定年份月份分佈 (Layer2.5 快取代理)
        ' 2026/04/17 by Claude: 從 GetMonthCountsForYearL3 拆出快取與提前過濾邏輯
        '   原來的快取/過濾邏輯混在 L3 裡，現在統一到此 L2.5 層，L3 只剩純 COM
        ' 呼叫順序:
        '   提前過濾 1 — GetMailCount=0   → 直接回傳空，不打 COM
        '   提前過濾 2 — _cacheYearCounts 已知該年無信 → 直接回傳空，不打 COM
        '   ① 記憶體命中 → ② DB lazy load → ③ Layer3 GetMonthCountsForYearL3
        ' OCE 不在此攔截，直接 re-throw 讓 CollectMonthCounts (L2) 的 Catch OCE 接住
        ' ---------------------------------------------------------------

        ' 提前過濾 1: 該資料夾完全無郵件，不必查快取或打 COM
        ' 2026/04/10 by Gemini: 解決 DB 沒存 0 封信記錄，lazy_load 回 Nothing 被迫打 COM 的問題
        If GetMailCount(folder, fPath:=fPath) = 0 Then Return New ConcurrentDictionary(Of Integer, Integer)()

        ' 提前過濾 2: 年度快取已知此年份信件數為 0，不必打月份 COM
        ' 2026/04/10 by Gemini: 省掉「某資料夾在 2001 年確定無信」的多餘 COM 呼叫
        Dim yCache As ConcurrentDictionary(Of Integer, Integer) = Nothing
        If _cacheYearCounts.TryGetValue(fPath, yCache) Then
            Dim countInYear As Integer = 0
            yCache.TryGetValue(year, countInYear)
            If countInYear = 0 Then Return New ConcurrentDictionary(Of Integer, Integer)()
        End If

        Dim cacheKey As String = fPath & "_" & year.ToString()
        Dim value As ConcurrentDictionary(Of Integer, Integer) = Nothing
        If _cacheMonthCounts.TryGetValue(cacheKey, value) Then Return value  ' ① 記憶體命中

        Dim dbResult = DbGetMonthCountsForFolder(fPath, year)                 ' ② DB lazy load
        If dbResult IsNot Nothing Then
            _cacheMonthCounts.TryAdd(cacheKey, dbResult)
            If _iLikeNoisy Then _dbg("DB 命中", $"{ExtractFolderName(fPath)} {year} 年 ({dbResult.Count} 個月)")
            Return dbResult
        End If

        ' ③ Layer3 COM 呼叫；OCE re-throw，不在此攔截 (寫入快取在 COM 完成後，OCE 天然繞過)
        Dim monthCounts = Await GetMonthCountsForYearL3(folder, year, fPath:=fPath, cToken:=cToken)
        _cacheMonthCounts(cacheKey) = monthCounts           ' ✅ 完整計算後存入快取
        ' DbSaveMonthCountsSingle(fPath, year, monthCounts) ' ✅ 2026/04/09 設計: 增量寫入 DB (待啟用)
        Return monthCounts

    End Function
    Private Async Function GetAttachMailList(folder As Folder, progress As IProgress(Of ProgressReport), Optional fPath As String = "", Optional cToken As CancellationToken = Nothing) As Task(Of List(Of MailItemInfo))
        ' ---------------------------------------------------------------
        ' GetAttachMailList — Tab3 Phase1：含附件的候選郵件清單
        ' by Gemini, 2026/04/05: Layer2.5 快取代理層 - Tab3 Phase 1 快取 - 取得單一資料夾本層含附件的郵件清單
        ' 2026/4/7 by Claude, 加入Database lazy load, 讀資料順序:
        '   ① 記憶體命中
        '   ② DB lazy load：命中且 mca 有效且 snapshot 吻合 → 一次填滿所有欄位
        '   ③ fallback: Layer3 呼叫
        ' 2026/04/15 by Claude: 加入 cToken 參數，傳遞至 GetAttachMailListL3 (Layer3)
        ' ---------------------------------------------------------------
        fPath = SafeGetPath(folder, fPath)
        Dim fName As String = ExtractFolderName(fPath)
        If _iLikeNoisy Then _dbg(" ├ 開始", fName)
        Dim key As String = fPath
        Dim currentCount As Long = GetMailCount(folder, fPath:=fPath)  ' 依賴同層快取 (本身已有 DB lazy load) 

        ' ① 記憶體命中
        Dim entry As FolderCacheTab3 = Nothing ' 補上初始化以消除 BC42108 警告
        If _cacheAttachMailList.TryGetValue(key, entry) AndAlso entry.ItemCountSnap = currentCount Then Return entry.AttachMailList

        ' ② DB lazy load (attach_maillist)：item_count_snap == currentCount → 快取仍有效
        Dim dbResult = DbGetAttachMailList(key)
        If dbResult IsNot Nothing AndAlso dbResult.Snap = currentCount Then
            Dim cached As New FolderCacheTab3 With {.AttachMailList = dbResult.Mails, .ItemCountSnap = currentCount}
            _cacheAttachMailList(key) = cached   ' 覆蓋式寫入，確保 ItemCountSnap 對應正確
            If _iLikeNoisy Then _dbg(" ├ DB 命中", $"{fName} ({dbResult.Mails.Count} 封)")
            Return dbResult.Mails
        End If

        ' ③ fallback: Layer3 呼叫
        ' 2026/04/15: 加入 cToken 傳遞，取消時 GetAttachMailListL3 回傳空 List，不寫入快取 (見下方 If 判斷) 
        Dim targetMailList As List(Of MailItemInfo) = Await GetAttachMailListL3(folder, progress, cToken:=cToken)
        _cacheAttachMailList(key) = New FolderCacheTab3 With {.AttachMailList = targetMailList, .ItemCountSnap = currentCount}
        ' 2026/04/05: 不使用 TryAdd/TryUpdate，確保最後的 cache entry 是正確的 (ItemCountSnap 與 mail list 對應)
        If _iLikeNoisy Then _dbg(" ├ 結束", fName)
        Return targetMailList

    End Function
    Private Function GetAttachFilename(ByRef mail As MailItemInfo) As List(Of String)
        ' ---------------------------------------------------------------
        ' GetAttachFilename — Tab3 Phase2：附件檔名清單 (by EntryID)
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
        result = GetAttachFilenameL3(mail)
        If result IsNot Nothing Then _cacheAttachFilename.TryAdd(mail.EntryID, result)
        Return result
    End Function
    Private Function GetMailBody(entryID As String) As String
        ' ── L2.5 快取代理層 ──────────────────────────────────────────────────────────
        ' ---------------------------------------------------------------
        ' GetMailBody — Layer2.5 快取代理：Body 快取存取點
        ' 2026/04/28 by Simon/Claude: 依照 L2.5 架構抽出快取邏輯，L3 只剩純 COM
        '   ① 快取命中（_lv4BodyCache）→ 直接回傳，0 COM call
        '   ② 快取未命中 → 呼叫 L3 GetMailBodyL3 讀取並正規化
        '   ③ 無論成功或失敗都存快取（失敗存 ""），避免同一封信重複嘗試 COM
        ' ---------------------------------------------------------------
        Dim cached As String = Nothing
        If _lv4BodyCache.TryGetValue(entryID, cached) Then Return cached

        Dim body As String = GetMailBodyL3(entryID)
        _lv4BodyCache(entryID) = body   ' 無論成功失敗都存入，避免重複打 COM
        Return body

    End Function

    Private Async Function GetBasicMailInfo(folder As Folder, needTopic As Boolean, cToken As CancellationToken, Optional fPath As String = "") As Task(Of List(Of (Mail As MailItemInfo, Topic As String)))
        ' ---------------------------------------------------------------
        ' GetBasicMailInfo — Layer2.5 快取存取點 (Tab4/Tab5/Tab7)
        ' 2026/05/06 by Claude: cache key 改為純 fPath（移除 |needTopic 後綴）
        ' 2026/05/11 by Claude Sonnet 4.6: 改用 L2.5 快取，記憶體命中時 0 COM；配合刪除後主動 invalidate _cacheMailCount 確保不污染
        ' 2026/05/12 by Simon/Claude: ① 記憶體命中邏輯重構
        '   - _cacheMailCount 有值 → 用來驗 snap，避免 COM → Return entry.Mails
        '   - _cacheMailCount 無值 → 直接信任 _cacheBasicMailInfo，不打 COM → Return entry.Mails
        '   - 信任依賴：刪除後 InvalidateBasicMailCache 會同時清兩個快取，確保不回傳鬼魂資料
        ' ---------------------------------------------------------------
        fPath = SafeGetPath(folder, fPath)
        'Dim cacheKey As String = fPath                                 ' 2026/05/06 by Claude: 純路徑，Tab4/Tab5/Tab7 共用
        'Dim fName As String = ExtractFolderName(fPath)                 ' 2026/05/11 by simon: 這個fName好像沒用到? 先保留未來可能用於除錯輸出
        'Dim currentSnap As Long = GetLiveFolderSnapL3(sFolder, fPath)   ' 2026/05/11 by Claude Sonnet 4.6: 改用 L2.5 快取，記憶體命中時 0 COM；配合刪除後主動 invalidate _cacheMailCount 確保不污染

        ' ① 記憶體命中檢查
        Dim entry As (Mails As List(Of (Mail As MailItemInfo, Topic As String)), Snap As Long) = Nothing
        If _cacheBasicMailInfo.TryGetValue(fPath, entry) Then Return entry.Mails

        ' ② DB lazy load (basic_maillist 存在的話)
        ' 這裡不管 needTopic 是 True 還是 False，只要 DB 有最新資料，我們都拿來用
        Dim currentSnap As Long = GetMailCount(folder, fPath)  ' L2.5，memory > DB > COM
        Dim dbResult = DbGetBasicMailInfo(fPath)
        If dbResult.HasValue AndAlso dbResult.Value.Snap = currentSnap Then
            Dim mails = dbResult.Value.Mails
            _cacheBasicMailInfo(fPath) = (mails, currentSnap)
            Return mails
        End If

        ' ③ Fallback: Layer3 COM 掃描
        Dim resultList = Await GetBasicMailInfoL3(folder, needTopic, cToken, fPath)

        ' 掃描完畢，存入記憶體快取 (SaveCache 時會持久化到 SSD)
        If resultList IsNot Nothing Then _cacheBasicMailInfo(fPath) = (resultList, currentSnap)

        Return resultList
    End Function
    Private Async Function RdoPreloadAttach_1(sourceList As List(Of MailItemInfo), progress As IProgress(Of ProgressReport), cToken As CancellationToken) As Task

        ' =================================================================
        ' by Gemini, 2026/04/05: Layer2.5 快取代理層 - 批次預熱附件檔名快取
        '   利用 Redemption (RDO) Free-Threaded 安全的特性，
        '   在進入 Layer2 迴圈前平行提早把附件檔名讀進 _cacheAttachFilename。
        '   完全不更改原有的迴圈運作邏輯，以預讀取的型態塞資料進快取來大幅壓縮等待時間。
        ' =================================================================
        If _rdo Is Nothing OrElse sourceList.Count = 0 Then Return

        _dbg("開始", $"RDO平行預載 {sourceList.Count} 筆")
        Dim swTotal As Stopwatch = Stopwatch.StartNew()     ' by Claude Sonnet 4.6, 2026/06/07
        Dim swThrottle As Stopwatch = Stopwatch.StartNew()  ' by Claude Sonnet 4.6, 2026/06/07
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
                                                                Dim list As New List(Of String)(512)
                                                                For i As Integer = 1 To rdoMsg.Attachments.Count    ' COM 的 index 從 1 開始而不是0
                                                                    list.Add(rdoMsg.Attachments.Item(i).FileName)
                                                                Next
                                                                _cacheAttachFilename.TryAdd(mail.EntryID, list)
                                                            End If
                                                        Catch
                                                        Finally
                                                            If rdoMsg IsNot Nothing Then TryMarshalRelease(rdoMsg)
                                                        End Try
                                                    End If

                                                    Dim curProcessed As Integer = Interlocked.Increment(processed)
                                                    If swThrottle.ElapsedMilliseconds >= ThrottleFreq.Hii OrElse curProcessed = total Then
                                                        Dim eta = CalculateSpeedAndETA(total, curProcessed, swTotal.Elapsed.TotalSeconds)
                                                        progress?.Report(New ProgressReport With {.CurrentCount = curProcessed, .TotalCount = total,
                                                                                                  .Message = $"Phase 2 (RDO 預載快取): {curProcessed} / {total} ({eta.Speed:F0} 封/秒{eta.EtaString})"})
                                                        swThrottle.Restart()
                                                    End If
                                                End Sub)
                           Catch ex As OperationCanceledException
                               ' cToken 取消時 Parallel.ForEach 拋出，正常中斷，不需處理
                           End Try
                       End Sub, cToken)
        _dbg(" ├ 結束", $"RDO 預載完成，處理共 {processed} 筆") ' by Gemini, 2026/04/10: 調整縮排層級為 Level 1
    End Function
    Private Async Function RdoPreloadAttach_2(sourceList As List(Of MailItemInfo), progress As IProgress(Of ProgressReport), cToken As CancellationToken) As Task

        ' ==============================================================
        ' by AntiGravity, 2026/04/07: 實驗性質
        ' - 使用 Task.WhenAll + SemaphoreSlim，試圖推高 SSD I/O 並發度
        ' ==============================================================
        If _rdo Is Nothing OrElse sourceList.Count = 0 Then Return

        _dbg(" ├ 開始", $"WhenAll平行預載 {sourceList.Count} 筆") ' by Gemini, 2026/04/10: 調整縮排層級為 Level 1
        Dim swThrottle As Stopwatch = Stopwatch.StartNew()  ' by Claude Sonnet 4.6, 2026/06/07
        Dim swTotal As Stopwatch = Stopwatch.StartNew()     ' by Claude Sonnet 4.6, 2026/06/07
        Dim processed As Integer = 0
        Dim total As Integer = sourceList.Count

        ' 設定並發數：嘗試設為 CPU 核心數的 4 倍，壓榨 SSD 的 Queue Depth
        Dim maxConcurrency As Integer = Environment.ProcessorCount * 4
        Dim throttler As New SemaphoreSlim(maxConcurrency)
        Dim tasks As New List(Of Task)(32)

        For Each m As MailItemInfo In sourceList
            Dim mail = m ' 在 lambda 中避免變數捕獲問題

            tasks.Add(Task.Run(Async Function()
                                   Await throttler.WaitAsync(cToken)   ' ✅ cToken 取消時直接拋 OperationCanceledException
                                   Try
                                       If Not _cacheAttachFilename.ContainsKey(mail.EntryID) Then
                                           Dim rdoMsg As Redemption.RDOMail = Nothing
                                           Try
                                               rdoMsg = TryCast(_rdo.GetMessageFromID(mail.EntryID), Redemption.RDOMail)
                                               If rdoMsg IsNot Nothing Then
                                                   Dim list As New List(Of String)(512)
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

                                       Dim curProcessed As Integer = Interlocked.Increment(processed)
                                       If swThrottle.ElapsedMilliseconds >= ThrottleFreq.Hii OrElse curProcessed = total Then
                                           Dim eta = CalculateSpeedAndETA(total, curProcessed, swTotal.Elapsed.TotalSeconds)
                                           progress?.Report(New ProgressReport With {.CurrentCount = curProcessed, .TotalCount = total,
                                                                                     .Message = $"Phase 2 (WhenAll 預載): {curProcessed} / {total} ({eta.Speed:F0} 封/秒{eta.EtaString})"})
                                           swThrottle.Restart()
                                       End If
                                   Finally
                                       throttler.Release()
                                   End Try
                               End Function, cToken))
        Next

        If tasks.Count > 0 Then Await Task.WhenAll(tasks)
        _dbg(" ├ 結束", $"WhenAll 預載完成，處理共 {processed} 筆") ' by Gemini, 2026/04/10
    End Function
    Private Async Function PreLoadBasicMailCacheAsync(folderList As List(Of (Folder As Folder, fPath As String)), cToken As CancellationToken) As Task
        ' ---------------------------------------------------------------
        ' PreLoadBasicMailCacheAsync — SSD 批次預熱（優化B）
        ' 對尚未在記憶體的路徑發出一次 SQL IN 批次查詢，填入 _cacheBasicMailInfo。
        ' 之後主迴圈的 GetBasicMailInfo 全部命中 memory，不再逐個打 DB。
        ' 不做 snap 驗證：信任快取，失效由刪除後的 InvalidateBasicMailCacheForPaths 負責。
        ' 2026/05/11 by Simon/Claude: 優化B
        ' ---------------------------------------------------------------
        Dim missedPaths = folderList.Where(Function(f) Not _cacheBasicMailInfo.ContainsKey(f.fPath)) _
                                    .Select(Function(f) f.fPath).ToList()
        If missedPaths.Count = 0 Then Return

        _dbg(" ├ 開始", $"DB 批次查詢 {missedPaths.Count} 個未命中路徑")
        Dim dbBatch = Await Task.Run(Function() DbGetBasicMailInfoBatch(missedPaths), cToken)

        For Each kvp In dbBatch
            _cacheBasicMailInfo.TryAdd(kvp.Key, (kvp.Value.Mails, kvp.Value.Snap))
        Next

        _dbg(" ├ 結束", $"預熱完成，填入 {dbBatch.Count} 個資料夾")
        Await Task.Yield()
    End Function
    Private Sub InvalidateBasicMailCache(fPath As String)
        ' ---------------------------------------------------------------
        ' InvalidateBasicMailCache — 刪除郵件後，主動清除指定 fPath 的記憶體快取
        ' 只清 _cacheBasicMailInfo 和 _cacheMailCount 兩個 key，不影響其他資料夾
        ' 配合 DbDeleteBasicMailInfoByPath 一起呼叫，確保記憶體與 DB 兩層同步失效
        ' 2026/05/11 by Claude Sonnet 4.6
        ' 2026/05/12 by Simon/Claude: 擴充清除範圍至所有受影響的快取
        ' ---------------------------------------------------------------
        If String.IsNullOrEmpty(fPath) Then Return

        ' ── 層次一：該資料夾本身 ──────────────────────────────────────
        Dim dummy1 As (Mails As List(Of (Mail As MailItemInfo, Topic As String)), Snap As Long) = Nothing
        _cacheBasicMailInfo.TryRemove(fPath, dummy1)

        Dim dummy2 As Long
        _cacheMailCount.TryRemove(fPath, dummy2)
        _cacheMailCountAll.TryRemove(fPath, dummy2)
        _cacheMailCountAll.TryRemove(fPath & "|True", dummy2)
        _cacheMailCountAll.TryRemove(fPath & "|False", dummy2)

        _cacheFolderSize.TryRemove(fPath, dummy2)
        _cacheFolderSizeAll.TryRemove(fPath, dummy2)

        _cacheYearCounts.TryRemove(fPath, Nothing)

        ' month_counts key 格式為 "fPath_YYYY"，不知道是哪年，清所有匹配的
        For Each mk In _cacheMonthCounts.Keys.Where(Function(k) k.StartsWith(fPath & "_")).ToList()
            _cacheMonthCounts.TryRemove(mk, Nothing)
        Next

        _cacheAttachMailList.TryRemove(fPath, Nothing)

        ' ── 層次二：所有祖先路徑的聚合快取 ──────────────────────────
        For Each ancestor In GetAncestors(fPath)
            _cacheMailCountAll.TryRemove(ancestor, dummy2)
            _cacheMailCountAll.TryRemove(ancestor & "|True", dummy2)
            _cacheMailCountAll.TryRemove(ancestor & "|False", dummy2)
            _cacheFolderSizeAll.TryRemove(ancestor, dummy2)
        Next

        _dbg("結束", ExtractFolderName(fPath))
    End Sub
    Friend Sub InvalidateFolderTreeCache(fPath As String)
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

        ' 【修復關鍵 2】補上身分證與屬性字典的清理, 2026/6/1 by Simon/Gemini 3.1 Pro
        Dim dummyId As (eid As String, sid As String, isMail As Boolean, hasCh As Boolean) = Nothing
        For Each key In _cacheFolderIDs.Keys.Where(isInSub).ToList() : _cacheFolderIDs.TryRemove(key, dummyId) : Next
        For Each key In _cacheIsMailFolder.Keys.Where(isInSub).ToList() : _cacheIsMailFolder.TryRemove(key, False) : Next

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
#End Region
#Region "  ├ Layer3 直接存取底層計數函數"
    Private Function GetMailCountL3(folder As Folder, Optional fPath As String = "") As Long
        ' --------------------------------------------------------------
        ' GetMailCountL3: 只讀單一資料夾的本層郵件數 (不含子孫)
        ' Fallback 鏈:
        '   ⓪ Redemption : RDOFolder.Items.Count (可在非 STA 執行緒呼叫)
        '   ① MAPI : PR_CONTENT_COUNT (0x36020003) (最快快取屬性)
        '   ② OOM  : pFolder.Items.Count (會建立 Items 集合)
        '   ③ fail : Return -1
        ' --------------------------------------------------------------
        fPath = SafeGetPath(folder, fPath)
        Dim fName As String = ExtractFolderName(fPath)
        If _iLikeNoisy Then _dbg("開始", fName)
        Dim sw As Stopwatch = Stopwatch.StartNew()  ' by Claude Sonnet 4.6, 2026/06/07

        ' ⓪ Redemption: RDOFolder.Items.Count
        If _rdo IsNot Nothing Then
            Dim rdoFolder As Redemption.RDOFolder = Nothing
            Try
                rdoFolder = _rdo.GetFolderFromID(folder.EntryID, folder.StoreID)
                Dim count As Long = CLng(rdoFolder.Items.Count) : Return count
            Catch ex As System.Exception
                If _iLikeNoisy Then _dbg("    ├ 錯誤路徑", $"GetMailCount ⓪ RDO: {fName} | {ex.Message}") ' by Gemini, 2026/04/10
            Finally
                TryMarshalRelease(rdoFolder)
            End Try
        End If

        ' ① MAPI: PR_CONTENT_COUNT (0x36020003)
        Try
            Const PR_CONTENT_COUNT As String = "http://schemas.microsoft.com/mapi/proptag/0x36020003"
            Dim count As Long = CLng(folder.PropertyAccessor.GetProperty(PR_CONTENT_COUNT))
            Return count
        Catch ex As System.Exception
            If _iLikeNoisy Then _dbg("    ├ 錯誤路徑", $"GetMailCount ① MAPI: {fName} | {ex.Message}") ' by Gemini, 2026/04/10
        End Try

        ' ② OOM: pFolder.Items.Count
        Try
            Dim items As Outlook.Items = Nothing
            Try
                items = folder.Items
                Dim count As Long = CLng(items.Count) : Return count
            Finally
                TryMarshalRelease(items)
            End Try
        Catch ex As System.Exception
            If _iLikeNoisy Then _dbg("    ├ 錯誤路徑", $"GetMailCount ② OOM: {fName} | {ex.Message}") ' by Gemini, 2026/04/10
        End Try

        sw.Stop()
        If _iLikeNoisy Then _dbg("    ├ 結束", $"FAIL: {fName} | -1 | {sw.ElapsedMilliseconds}ms")
        Return -1

    End Function
    Private Function GetFolderCountL3(sFolder As Folder, Optional fPath As String = "") As Long
        ' --------------------------------------------------------------
        ' GetFolderCountL3: 讀取單一資料夾的本層直屬子資料夾數
        '
        ' Fallback 鏈:
        '   ⓪ Redemption : RDOFolder.Folders.Count
        '            可從非 STA 執行緒呼叫，繞過 Outlook Security Guard
        '            _rdoSession 未就緒時自動跳過此層
        '   ① MAPI : PR_FOLDER_CHILD_COUNT (0x66380003, PT_LONG) 一次 PropertyAccessor call，在大多數情況下準確
        '            注意: PST 上此屬性在剛移動資料夾後可能短暫不同步，但 Outlook 關閉再開就會修正，日常使用可接受
        '            2026/3/20 實測: PR_FOLDER_CHILD_COUNT 沒有一次成功過，已暫時 comment 出
        '   ② OOM  : pFolder.Folders.Count
        '            Folders 集合比 Items 輕量，載入速度可接受，且永遠準確
        '   ③ fail : Return -1
        '
        ' 關於「先讀 PR_SUBFOLDERS (0x360A000B) 再讀個數」的設計討論:
        '   PR_SUBFOLDERS 是 PT_BOOLEAN，只告訴你有沒有子資料夾 (不告訴你幾個)
        '   先讀它再讀 PR_FOLDER_CHILD_COUNT 等於多一次 COM call，只有「大多數資料夾都沒有子資料夾」時才划算，
        '   實際 PST 不符合此條件，因此直接讀 PR_FOLDER_CHILD_COUNT，不做 PR_SUBFOLDERS 前置判斷
        '
        ' 取代: 散落各處的 pFolder.Folders.Count 直接呼叫 (建議逐一替換)
        ' --------------------------------------------------------------
        fPath = SafeGetPath(sFolder, fPath)
        Dim fName As String = ExtractFolderName(fPath)
        If _iLikeNoisy Then _dbg("    ├ 開始", fName) ' by Gemini, 2026/04/10: Level 1

        ' ⓪ Redemption: RDOFolder.Folders.Count
        '   與 OOM pFolder.Folders.Count 等價，但可在任意執行緒呼叫, 2026-03-22 新增
        If _rdo IsNot Nothing Then
            Dim rdoFolder As Redemption.RDOFolder = Nothing
            Try
                rdoFolder = _rdo.GetFolderFromID(sFolder.EntryID, sFolder.StoreID)
                Dim count As Long = CLng(rdoFolder.Folders.Count) : Return count
            Catch ex As System.Exception
                _dbg("    ├ 錯誤路徑", $"GetFolderCount ⓪ RDO: {fName} | {ex.Message}") ' by Gemini, 2026/04/11: Level 3
            Finally
                TryMarshalRelease(rdoFolder)
            End Try
        End If

        ' ① MAPI: PR_FOLDER_CHILD_COUNT (0x66380003)
        ' 2026/3/20, 奇怪PR_FOLDER_CHILD_COUNT 沒有一次成功過??? 乾脆先拿掉這個try, 省得一直fallback也是浪費開銷


        '   ② OOM  : Folder.Folders.Count
        Try
            'Dim count As Long = CLng(sFolder.Folders.Count) : Return count
            ' 2026/6/6 by simon, Folders 集合直接串接了 .Count，物件可能無法被 TryMarshalRelease 捕捉，
            '   改為先取得 Folders 物件再讀 Count，確保釋放 COM 物件
            Dim flds As Outlook.Folders = sFolder.Folders
            Dim count As Long = CLng(flds.Count)
            TryMarshalRelease(flds)
            Return count
        Catch ex As System.Exception
            _dbg("    ├ 錯誤路徑", $"GetFolderCount ① OOM: {fName} | {ex.Message}") ' by Gemini, 2026/04/11: Level 3
        End Try

        If _iLikeNoisy Then _dbg("    ├ 結束", $"FAIL: {fName}") ' by Gemini, 2026/04/10
        Return -1

    End Function
    Private Async Function GetMailCountAllL3(rootFolder As Folder, Optional progress As IProgress(Of ProgressReport) = Nothing, Optional cToken As CancellationToken = Nothing) As Task(Of Long)
        ' --------------------------------------------------------------
        ' GetMailCountAllL3 v3.6: 讀取某資料夾及其整棵子樹的郵件總數
        ' by Gemini, 2026/04/02: 升級為 IProgress(Of ProgressReport) 並加入 100ms 節流回報
        ' 2026/04/15 by Claude: 加入 cToken 參數，取代 _cancelRequested 旗標 (見函數內說明) 
        '
        ' v3.0 變更說明 (2026-03-22):
        '   合併原 GetMailCountAllL3 + GetMailCountAllParallel 為單一函數，
        '   統一 fallback 鏈，呼叫端不再需要選擇要用哪個版本。
        '   GetMailCountAllParallel 可標記廢棄或直接刪除。
        '
        ' 設計說明:
        '   為何呼叫 GetMailCountL3() 而非直接用 GetTable():
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
        '                   Task.Run 內的 GetMailCountL3(cFolder) 走 Redemption ⓪ 時是 free-threaded 安全的
        '                   若 GetMailCountL3 fallback 到 MAPI PropertyAccessor，仍有 STA 違規風險，需留意
        '   ② BFS 循序累加:
        '                   GetSubtreeToList BFS 展開 + GetMailCountL3(Layer3) 逐一加總
        '                   支援取消檢查和 onProgress 進度回報
        '                   平行路徑失敗時的安全 fallback
        '   ③ 遞迴 fallback:
        '                   GetSubtreeToList 本身失敗時 (極少見) 的最後保險
        '                   無法精確回報進度，但確保加總結果正確
        '   ④ Return -1: 四層都失敗，由 Layer2 決定如何處理
        '
        ' cToken 取消偵測 (2026/04/15 by Claude):
        '   ⓪ Redemption 路徑不插入取消檢查 (單次 call，幾乎瞬間完成)
        '   ① RDO Parallel.ForEach 路徑: 透過 ParallelOptions.CancellationToken 傳入 cToken；
        '      取消時 Parallel.ForEach 拋 OperationCanceledException，由 Catch OCE 接住並回傳 -1
        '   ② BFS 循序路徑: SmartThrottle 每 100ms 讓出一次；
        '      cToken 取消時 Task.Delay 拋 OCE，由 Catch OCE 接住並回傳 -1
        '   ③ 遞迴 fallback: 遞迴呼叫時透過 cToken 傳遞，子層 OCE 向上冒泡
        '   取消後統一回傳 -1 (與讀取失敗同語義) ，呼叫端 Layer2.5 檢查 IsCancellationRequested 不寫入快取
        '
        ' onProgress 參數 (可選):
        '   傳入 IProgress(Of ProgressReport) callback
        '   Layer2 每處理一個資料夾回報 (已完成數, 總數)，讓 Layer1 更新狀態列
        '   不需要進度回報時傳 Nothing
        '   ⓪ 和 ① 路徑不觸發 onProgress，② 路徑才會逐一回報
        '
        ' 取代:
        '   GetMailCountByMAPINew 的整棵子樹加總用途
        '   GetMailCountAllParallel (v3.0 已合併，舊版可廢棄)
        ' ---------------------------------------------------------------
        ' 2026/4/28 by simon: 目前GetxxxCountAll系列函數已成死碼, 沒有任何呼叫端與進入點
        ' 原始設計意圖:
        ' 	呼叫端 → GetMailCountAllAsync (L2.5) → GetMailCountAllL3 (L3)
        ' 	呼叫端 → GetFolderCountAllAsync (L2.5) → GetFolderCountAllL3 (L3)
        '
        ' 後來使用了BFS剪枝速度更快:
        ' 	Compute → BFS → SummarizeSubTreeBottomUp → UpdateFolderStatsCache
        ' 	(自己計算並直接寫入 _cacheMailCountAll / _cacheFolderCountAll，完全繞過 AllAsync系列函數)
        ' ---------------------------------------------------------------
        Dim rName As String = rootFolder.Name ' by Gemini 3.1 Pro 2026/04/16: 避免重複 COM 呼叫
        If _iLikeNoisy Then _dbg("    ├ 開始", rName) ' by Gemini, 2026/04/10: Level 1

        ' ⓪ Redemption: TotalItemCount 直接回傳整棵子樹郵件總數
        '   一次 COM call 結束，不需要任何 BFS 遍歷或平行處理
        '   2026-03-22 新增
        If _rdo IsNot Nothing Then
            Dim rdoFolder As Redemption.RDOFolder = Nothing
            Try
                rdoFolder = _rdo.GetFolderFromID(rootFolder.EntryID, rootFolder.StoreID)
                Dim total As Long = CLng(rdoFolder.TotalItemCount)
                If _iLikeNoisy Then _dbg("    ├ 結束", $"⓪ RDO 成功: {rName} | TotalItemCount={total}")
                Return total
            Catch ex As System.Exception
                If _iLikeNoisy Then _dbg("    ├ ⓪ RDO 失敗，走平行BFS fallback", $"{rName} | {ex.Message}") ' by Gemini, 2026/04/10: 調整縮排層級為 Level 1
            Finally
                TryMarshalRelease(rdoFolder)
            End Try
        End If

        ' 2026/3/24 by Gemini: ① 平行 BFS (RDO)
        '   使用 GetSubtreeToListL3_Rdo 取得清單，以 Parallel.ForEach 搭配 Interlocked.Add 快速加總
        '   Redemption (RDO) 是 free-threaded，在背景平行執行安全且極為高效
        ' 2026/04/15 by Claude: 改用 ParallelOptions.CancellationToken 取代 _cancelRequested 旗標
        If _rdo IsNot Nothing Then
            Dim rdoRoot As Redemption.RDOFolder = Nothing
            Try
                rdoRoot = _rdo.GetFolderFromID(rootFolder.EntryID, rootFolder.StoreID)
                Dim rdoFolderList As List(Of Redemption.RDOFolder) = GetSubtreeToListL3_Rdo(rdoRoot, includeSubF:=True)
                Dim targetFolderCount As Integer = rdoFolderList.Count
                Dim totalCount As Long = 0
                Dim processedCount As Integer = 0
                Dim swThrottle As Stopwatch = Stopwatch.StartNew()  ' by Gemini 3.0 flash, 2026/04/16: 補回計時器宣告; refactored by Claude Sonnet 4.6, 2026/06/07
                Dim parallelOptions As New ParallelOptions With {.CancellationToken = cToken}  ' 2026/04/15: cToken 傳入，取消時 ForEach 拋 OCE
                Parallel.ForEach(rdoFolderList, parallelOptions,
                    Sub(rdoF As Redemption.RDOFolder)
                        Try
                            Dim count As Integer = rdoF.Items.Count
                            Interlocked.Add(totalCount, CLng(count))
                        Catch ex As System.Exception
                            If _iLikeNoisy Then _dbg("    ├ ① 略過失敗資料夾", rdoF.Name) ' by Gemini, 2026/04/10: 調整縮排層級為 Level 2 (並行處理內部)
                        End Try
                        Dim done As Integer = Interlocked.Increment(processedCount)

                        ' by Gemini, 2026/04/02: 更新為 IProgress 且加上節流
                        ' 2026/04/16 by Gemini 3.0 flash: [註：此處位於 Parallel.ForEach 同步區塊內，不可 await] 故維持手動計時器節流。
                        If swThrottle.ElapsedMilliseconds >= ThrottleFreq.Hii OrElse done = targetFolderCount Then
                            progress?.Report(New ProgressReport With {.CurrentCount = done, .TotalCount = targetFolderCount,
                                                                      .Message = $"正在平行統計: {done} / {targetFolderCount} 個資料夾..."})
                            swThrottle.Restart()
                        End If
                    End Sub)
                If _iLikeNoisy Then _dbg("    ├ 結束", $"① 平行BFS成功 (RDO): {rName} | total={totalCount} | folders={targetFolderCount}")
                Return totalCount
            Catch ex As OperationCanceledException
                ' 2026/04/15: ParallelOptions.CancellationToken 取消時拋 OCE，正常中斷，回傳 -1 表示取消/失敗
                If _iLikeNoisy Then _dbg("    ├ ① 已取消", $"{rName}") : Return -1
            Catch ex As System.Exception
                If _iLikeNoisy Then _dbg("    ├ ① 平行BFS失敗，走循序BFS fallback", $"{rName} | {ex.Message}") ' by Gemini, 2026/04/10
            Finally
                TryMarshalRelease(rdoRoot)
            End Try
        End If

        ' ② BFS 循序累加: GetSubtreeToList 展開 + GetMailCountL3(Layer3) 逐一加總
        '   支援取消檢查和 progress 進度回報，比平行版保守但穩定
        ' 2026/04/15 by Claude: _cancelRequested 取代為 SmartThrottle(swThrottle, cToken)
        '   cToken 取消時 Task.Delay(1,cToken) 拋 OCE → Catch OCE → Return -1
        '   同時移除舊的 i Mod 10 Await Task.Yield()，統一由 SmartThrottle 每 100ms 讓出一次
        Try
            ' 2026/04/17 by Claude: 改呼叫 GetSubtreeToList (L2.5)，享有快取加速
            Dim targetFolderList As List(Of (folder As Folder, fPath As String)) = Await GetSubtreeToList(rootFolder, includeSubF:=True, cToken:=cToken)
            Dim grandTotal As Long = 0
            Dim swThrottle As Stopwatch = Stopwatch.StartNew()  ' by Gemini, 2026/04/02: 100ms 節流閥; refactored by Claude Sonnet 4.6, 2026/06/07

            For i As Integer = 0 To targetFolderList.Count - 1
                Dim cFolder As Folder = targetFolderList(i).folder
                Dim count As Integer = GetMailCountL3(cFolder)
                ' GetMailCountL3 的所有 fallback 都失敗才會到這個 else，記錄但不中止整體加總
                If count >= 0 Then grandTotal += CLng(count) Else If _iLikeNoisy Then _dbg("    ├ Get MailCountAll ② 略過失敗資料夾", cFolder.Name) ' by Gemini, 2026/04/10

                ' by Gemini, 2026/04/02: 100ms 節流回報進度
                ' 2026/04/15 by Claude: 節流區塊整合讓出與取消偵測，SmartThrottle 在 cToken 取消時拋 OCE
                ' 2026/04/16 by Gemini 3.0 flash: 節流區塊整合讓出與取消偵測，改用 onThrottled 委派
                Await SmartThrottle(swThrottle, cToken:=cToken, ThrottleFreq.Hii,
                                          Sub() progress?.Report(New ProgressReport With {.CurrentCount = i + 1, .TotalCount = targetFolderList.Count,
                                                                                          .Message = $"正在統計郵件數: {i + 1} / {targetFolderList.Count} 個資料夾..."}))
            Next
            If _iLikeNoisy Then _dbg("    ├ ② 循序BFS成功", $"{rName} | total={grandTotal}")
            Return grandTotal
        Catch ex As OperationCanceledException
            ' 2026/04/15: SmartThrottle 或 GetSubtreeToList 取消時拋 OCE，正常中斷
            If _iLikeNoisy Then _dbg("    ├ ② 已取消", $"{rName}") : Return -1
        Catch ex As System.Exception
            If _iLikeNoisy Then _dbg("    ├ ② 循序BFS失敗，走遞迴fallback", $"{rName} | {ex.Message}") ' by Gemini, 2026/04/10
        End Try

        ' ③ 遞迴 fallback: GetSubtreeToList 本身失敗時的最後保險
        '   無法精確回報進度，但確保加總結果正確
        '   注意: 遞迴呼叫會重新進入本函數，⓪ Redemption 已失敗所以 _rdoSession 仍 Nothing 或故障
        '         ① ② 也已失敗，只會走到 ③ 再次遞迴——理論上 ③ 不會無限展開，因為每層只遞迴直屬子資料夾
        '        若 ③ 常被觸發，需回頭檢查 GetSubtreeToList 失敗的根本原因 ' pending:
        Try
            Dim totalCount As Long = 0
            Dim count As Integer = GetMailCountL3(rootFolder)     ' 本層 mailcount
            If count >= 0 Then totalCount += count
            Await Task.Yield()
            ' 優化第六點：提取 Folders 集合並在 Finally 顯式釋放，防止遞迴過程中的 RCW 洩漏 (by Gemini 3 Flash, 2026/05/05)
            Dim subFolders As Folders = rootFolder.Folders
            Try
                For Each f As Folder In subFolders
                    Dim subCount As Long = Await GetMailCountAllL3(f, cToken:=cToken) ' 遞迴，傳遞 cToken
                    If subCount >= 0 Then totalCount += subCount
                Next
            Finally
                TryMarshalRelease(subFolders)
            End Try
            If _iLikeNoisy Then _dbg("    ├ 結束", $"③ 遞迴fallback成功: {rName} | total={totalCount}")
            Return totalCount
        Catch ex As System.Exception
            If _iLikeNoisy Then _dbg("    ├ ③ 遞迴fallback也失敗", $"{rName} | {ex.Message}")
            Return -1   ' ④ 四層都失敗，回傳 -1 讓 Layer2 知道這是「讀取失敗」而非「真的是 0 封」
        End Try

    End Function
    Private Async Function GetFolderCountAllL3(rootFolder As Folder, Optional progress As IProgress(Of ProgressReport) = Nothing, Optional cToken As CancellationToken = Nothing) As Task(Of Long)
        ' --------------------------------------------------------------
        ' GetFolderCountAllL3 v1.5: 讀取某資料夾整棵子樹的資料夾總數 (不含 rootFolder 自身)
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
        '   ① BFS 路徑: GetSubtreeToList 內部走 OOM pFolder.Folders 展開，展開後直接 .Count，不需 Layer3 讀取。
        '   ② 遞迴 fallback: 內部的 rootFolder.Folders.Count 和 ForEach 走 OOM，
        '      若日後改為呼叫 GetFolderCountL3(Layer3)，即可自動走 Redemption ⓪ 路徑。
        ' 2026/04/15 by Claude: 加入 cToken 參數，取代 _cancelRequested 旗標
        ' ---------------------------------------------------------------
        ' 2026/4/28 by simon: 目前GetxxxCountAll系列函數已成死碼, 沒有任何呼叫端與進入點
        ' 原始設計意圖:
        ' 	呼叫端 → GetMailCountAllAsync (L2.5) → GetMailCountAllL3 (L3)
        ' 	呼叫端 → GetFolderCountAllAsync (L2.5) → GetFolderCountAllL3 (L3)
        '
        ' 後來使用了BFS剪枝速度更快:
        ' 	Compute → BFS → SummarizeSubTreeBottomUp → UpdateFolderStatsCache
        ' 	(自己計算並直接寫入 _cacheMailCountAll / _cacheFolderCountAll，完全繞過 AllAsync系列函數)
        ' ---------------------------------------------------------------
        Dim rName As String = rootFolder.Name ' by Gemini 3.1 Pro 2026/04/16: 避免重複 COM 呼叫
        If _iLikeNoisy Then _dbg("    ├ 開始", rName)

        ' by Gemini, 2026/04/02: 預跑一次顯示準備中
        progress?.Report(New ProgressReport With {.Message = "正在展開資料夾結構...", .IsIndeterminate = True})

        ' 2026/3/24 by Gemini: ⓪ Redemption + 平行處理 (最快路徑)
        '   使用 GetSubtreeToListL3_Rdo 取得清單，以 Parallel.ForEach 搭配 Interlocked.Add(rdoF.Folders.Count) 快速加總
        ' 2026/04/15 by Claude: 改用 ParallelOptions.CancellationToken 取代 _cancelRequested 旗標
        If _rdo IsNot Nothing Then
            Dim rdoRoot As Redemption.RDOFolder = Nothing
            Try
                rdoRoot = _rdo.GetFolderFromID(rootFolder.EntryID, rootFolder.StoreID)
                Dim rdoFolderList As List(Of Redemption.RDOFolder) = GetSubtreeToListL3_Rdo(rdoRoot, includeSubF:=True)
                Dim targetFolderCount As Integer = rdoFolderList.Count
                Dim totalCount As Long = 0
                Dim processedCount As Integer = 0
                Dim swThrottle As Stopwatch = Stopwatch.StartNew()  ' by Gemini, 2026/04/02; refactored by Claude Sonnet 4.6, 2026/06/07
                Dim parallelOptions As New ParallelOptions With {.CancellationToken = cToken}  ' 2026/04/15
                Parallel.ForEach(rdoFolderList, parallelOptions,
                    Sub(rdoF As Redemption.RDOFolder)
                        Try
                            Dim count As Long = CLng(rdoF.Folders.Count)
                            Interlocked.Add(totalCount, count)
                        Catch ex As System.Exception
                            If _iLikeNoisy Then _dbg("    ├ ⓪ RDO 略過失敗資料夾", rdoF.Name) ' by Gemini, 2026/04/11: Level 3
                        End Try

                        Dim done As Integer = Interlocked.Increment(processedCount)
                        ' by Gemini, 2026/04/02: 更新為 IProgress 且加上 100ms 節流，取代原有的 Mod 10
                        ' 
                        If swThrottle.ElapsedMilliseconds >= ThrottleFreq.Hii Then
                            progress?.Report(New ProgressReport With {.CurrentCount = done, .TotalCount = targetFolderCount,
                                                                      .Message = $"正在統計資料夾樹: {done} / {targetFolderCount}..."})
                            swThrottle.Restart()
                        End If
                    End Sub)
                If _iLikeNoisy Then _dbg("    ├ 結束", $"⓪ RDO平行成功: {rName} | total={totalCount}") ' by Gemini, 2026/04/10
                Return totalCount
            Catch ex As OperationCanceledException
                ' 2026/04/15: ParallelOptions.CancellationToken 取消時拋 OCE，正常中斷
                If _iLikeNoisy Then _dbg("    ├ ⓪ 已取消", $"{rName}") : Return -1
            Catch ex As System.Exception
                If _iLikeNoisy Then _dbg("    ├ ⓪ RDO平行失敗，走OOM循序fallback", $"{rName} | {ex.Message}") ' by Gemini, 2026/04/10
            Finally
                TryMarshalRelease(rdoRoot)
            End Try
        End If

        ' 2026/3/24 by Gemini: ② OOM + BFS 循序 (無 Redemption 時的最後手段)
        '   必須循序處理 OOM COM 物件以避免 STA 違規
        ' 2026/04/15 by Claude: 傳入 cToken，GetSubtreeToList 本身支援取消，OCE 向上冒泡
        Try
            ' 2026/04/16 by Gemini: GetSubtreeToList 現在回傳 Tuple，解開它以維持後續邏輯
            ' 2026/04/17 by Claude: 改呼叫 GetSubtreeToList (L2.5)，享有快取加速
            Dim targetTupleList = Await GetSubtreeToList(rootFolder, includeSubF:=True, progress:=progress, cToken:=cToken)
            Dim allFolders = targetTupleList.Select(Function(x) x.folder).ToList()
            ' by Gemini, 2026/04/02: BFS 展開後回傳數量
            Dim total As Long = CLng(targetTupleList.Count - 1)
            progress?.Report(New ProgressReport With {.CurrentCount = CInt(total), .TotalCount = CInt(total), .Message = $"資料夾結構已展開: 共 {total} 個資料夾。"})
            Await Task.Yield()
            If _iLikeNoisy Then _dbg("    ├ 結束", $"② OOM BFS成功: {rName} | total={total}") ' by Gemini, 2026/04/10
            Return total
        Catch ex As OperationCanceledException
            If _iLikeNoisy Then _dbg("    ├ ② 已取消", $"{rName}") : Return -1
        Catch ex As System.Exception
            If _iLikeNoisy Then _dbg("    ├ ② OOM BFS失敗", $"{rName} | {ex.Message}") ' by Gemini, 2026/04/10
        End Try
        ' ③ 全部失敗
        Return -1

    End Function
    Private Async Function GetFolderSizeL3(folder As Folder, Optional progress As IProgress(Of ProgressReport) = Nothing, Optional fPath As String = "", Optional cToken As CancellationToken = Nothing) As Task(Of Long)
        ' --------------------------------------------------------------
        ' GetFolderSizeL3 v1.6: 讀取單一資料夾本層大小 (bytes)
        ' by Gemini, 2026/04/02: 加入 IProgress 支援以回報分批讀取進度 (100ms 節流)
        ' 2026/3/24 by Gemini: Fallback 鏈重構
        '   ⓪ Redemption : rdoFolder.Fields(PR_MESSAGE_SIZE_EXTENDED) (部分 Exchange 支援，極快)
        '   ① OOM  : pFolder.GetTable(PR_MESSAGE_SIZE_EXTENDED) + GetArray(500) (最快安全招式)
        '   ② OOM  : pFolder.GetTable(PR_MESSAGE_SIZE_EXTENDED) + GetNextRow() (備案)
        '   ③ fail : Return -1
        ' 2026/04/15 by Claude: 加入 cToken 參數
        '   ① ② 迴圈中改用 SmartThrottle(swThrottle, cToken) 取代 Task.Yield()
        '   cToken 取消時 Task.Delay 拋 OCE，由 Catch OCE 接住後 re-throw (讓 GetFolderSizeAllL3 感知)
        ' --------------------------------------------------------------
        fPath = SafeGetPath(folder, fPath)
        Dim fName As String = ExtractFolderName(fPath)
        If _iLikeNoisy Then _dbg("    ├ 開始", fName)
        Dim sw As Stopwatch = Stopwatch.StartNew()  ' by Claude Sonnet 4.6, 2026/06/07

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
                    If _iLikeNoisy Then _dbg("    ├ 結束", $"⓪ RDO Fields 成功: {fName} | size={totalSize} | {sw.ElapsedMilliseconds}ms") ' by Gemini, 2026/04/10
                    Return totalSize
                End If
            Catch ex As System.Exception
                If _iLikeNoisy Then _dbg("    ├ 錯誤: ⓪ RDO 失敗，走 OOM GetArray fallback", $"{fName} | {ex.Message}") ' by Gemini, 2026/04/11: Level 3
            Finally
                TryMarshalRelease(rdoFolder)
            End Try
        End If

        ' ① OOM GetTable + GetArray(500) (目前最穩、最快的批次讀取)
        Const PR_SIZE_EX_STR As String = "http://schemas.microsoft.com/mapi/proptag/0x0E080014"
        Dim table As Outlook.Table = Nothing
        Try
            table = SafeGetTable(folder, "", PR_SIZE_EX_STR)
            Dim totalSize As Long = 0
            Dim swThrottle As Stopwatch = Stopwatch.StartNew()  ' by Gemini, 2026/04/02; refactored by Claude Sonnet 4.6, 2026/06/07

            Do
                Dim data = SafeGetArray(table)
                If data Is Nothing Then Exit Do
                For r As Integer = 0 To data.GetUpperBound(0)
                    Dim sz = data(r, 0)
                    If sz IsNot Nothing AndAlso Not IsDBNull(sz) Then totalSize += CLng(sz)
                Next

                ' by Gemini, 2026/04/02: 單一資料夾內部進度回報 (100ms 節流)
                ' 2026/04/15 by Claude: 改用 SmartThrottle 整合讓出與取消偵測
                '   進度回報移入節流區塊內，由 SmartThrottle 統一控制頻率
                ' 2026/04/16 by Gemini 3.0 flash: 改用 SmartThrottle 整合進度回報與讓出點
                Await SmartThrottle(swThrottle, cToken:=cToken, ThrottleFreq.Hii,
                                          Sub() progress?.Report(New ProgressReport With {.Message = $"正在計算 {fName} 大小: {totalSize / 1024 / 1024:0.0} MB..."}))
            Loop
            sw.Stop()
            If _iLikeNoisy Then _dbg(" ├ 結束", $"① OOM GetArray 成功: {fName} | size={totalSize} | {sw.ElapsedMilliseconds}ms") ' by Gemini, 2026/04/10
            Return totalSize
        Catch ex As OperationCanceledException
            ' 2026/04/15: cToken 取消時 re-throw，讓 GetFolderSizeAllL3 ① 的 For 迴圈感知並中止
            If _iLikeNoisy Then _dbg("    ├ ① OOM GetArray 已取消", fName) : Throw
        Catch ex As System.Exception
            If _iLikeNoisy Then _dbg("    ├ 錯誤: ① OOM GetArray 失敗，走 GetNextRow fallback", $"{fName} | {ex.Message}") ' by Gemini, 2026/04/11: Level 3
        Finally
            TryMarshalRelease(table)
        End Try

        ' ② OOM GetTable + GetNextRow() (不依賴二維陣列的最後保險)
        Dim table2 As Outlook.Table = Nothing
        Try
            table2 = SafeGetTable(folder, "", PR_SIZE_EX_STR)
            Dim totalSize As Long = 0
            Dim swThrottle2 As Stopwatch = Stopwatch.StartNew()  ' 2026/04/15: 獨立命名避免與①的 swThrottle 衝突; refactored by Claude Sonnet 4.6, 2026/06/07
            Do While Not table2.EndOfTable
                Dim row As Outlook.Row = table2.GetNextRow()
                If row IsNot Nothing Then
                    totalSize += SafeGet(Of Long)(row, PR_SIZE_EX_STR, 0L)
                    TryMarshalRelease(row)
                End If
                Await SmartThrottle(swThrottle2, cToken:=cToken)  ' 2026/04/15: 取代 loopCount Mod 500 Yield
            Loop
            sw.Stop()
            If _iLikeNoisy Then _dbg("    ├ 結束", $"② OOM GetNextRow 成功: {fName} | size={totalSize} | {sw.ElapsedMilliseconds}ms") ' by Gemini, 2026/04/10
            Return totalSize
        Catch ex As OperationCanceledException
            ' 2026/04/15: re-throw 讓上層感知
            If _iLikeNoisy Then _dbg("    ├ ② OOM GetNextRow 已取消", fName) : Throw
        Catch ex As System.Exception
            If _iLikeNoisy Then _dbg("    ├ 錯誤: ② OOM GetNextRow 失敗", $"{fName} | {ex.Message}") ' by Gemini, 2026/04/11: Level 3
        Finally
            TryMarshalRelease(table2)
        End Try

        sw.Stop()
        If _iLikeNoisy Then _dbg("    ├ 結束", $"FAIL: {fName} | -1 | {sw.ElapsedMilliseconds}ms")
        Return -1
    End Function
    Private Async Function GetFolderSizeAllL3(rootFolder As Folder, Optional progress As IProgress(Of ProgressReport) = Nothing, Optional cToken As CancellationToken = Nothing) As Task(Of Long)
        ' --------------------------------------------------------------
        ' GetFolderSizeAllL3 v1.6: 讀取某資料夾及整棵子樹的大小總計 (bytes)
        ' by Gemini, 2026/04/02: 增加 IProgress 支援與 100ms 節流回報
        '
        ' 2026/3/24 by Gemini: 落實新的 Fallback 鏈設計，並修正平行處理的 STA 問題
        '   ⓪ Redemption 平行路徑 (最快):
        '      利用 GetSubtreeToListL3_Rdo 一次把該子樹下所有 RDOFolder 拿出來，
        '      放到 Parallel.ForEach 中，各別讀取 MAPI 屬性 PR_MESSAGE_SIZE_EXTENDED。
        '      (RDOFolder 不支援 GetTable().GetArray()，故依賴屬性直讀)
        '
        '   ① OOM 循序路徑 (最安全):
        '      當 RDO 平行路徑失敗 (或是未匯入 Redemption)，退回使用 OOM。
        '      OOM 絕對不可以在 Task.Run / WhenAll 等背景執行緒內呼叫 COM，否則會觸發 STA 錯誤。
        '      故改為嚴格的 For 迴圈，逐一 Await GetFolderSizeL3()。
        '      而內部的 GetFolderSizeL3 會走到它專屬的 GetTable().GetArray(500) OOM 極速路徑。
        '
        '   ② 兩層都失敗: 回傳 -1，交給上一層流程處理。
        ' 2026/04/15 by Claude: 加入 cToken 參數，取代 _cancelRequested 旗標
        '   ⓪ RDO Parallel.ForEach 路徑: 透過 ParallelOptions.CancellationToken 傳入 cToken
        '   ① OOM 循序路徑: SmartThrottle 每 100ms 讓出一次，cToken 取消時 OCE 冒泡
        '      GetFolderSizeL3 內部 OCE 會 re-throw，For 迴圈 Catch OCE → Return -1
        ' --------------------------------------------------------------
        Dim rName As String = rootFolder.Name ' by Gemini 3.1 Pro 2026/04/16: 避免重複 COM 呼叫
        If _iLikeNoisy Then _dbg("    ├ 開始", rName)

        ' 2026/3/24 by Gemini: ⓪ Redemption 平行累加 PR_MESSAGE_SIZE_EXTENDED
        ' 2026/04/15 by Claude: 改用 ParallelOptions.CancellationToken 取代 _cancelRequested 旗標
        If _rdo IsNot Nothing Then
            Dim rdoRoot As Redemption.RDOFolder = Nothing
            Try
                rdoRoot = _rdo.GetFolderFromID(rootFolder.EntryID, rootFolder.StoreID)
                Dim rdoFolderList As List(Of Redemption.RDOFolder) = GetSubtreeToListL3_Rdo(rdoRoot, includeSubF:=True)
                Dim grandTotal As Long = 0
                Const PR_SIZE_EX As Integer = &HE080014

                ' 利用 Parallel.ForEach 與 Interlocked.Add 達到極致的多核並發加總
                Dim validCount As Integer = 0
                Dim parallelOptions As New ParallelOptions With {.CancellationToken = cToken}  ' 2026/04/15
                Parallel.ForEach(rdoFolderList, parallelOptions,
                    Sub(rdoF As Redemption.RDOFolder)
                        Try
                            Dim val As Object = rdoF.Fields(PR_SIZE_EX)
                            If val IsNot Nothing AndAlso Not IsDBNull(val) Then
                                Interlocked.Add(grandTotal, CLng(val))
                                Interlocked.Increment(validCount)
                            End If
                        Catch ex As System.Exception
                            If _iLikeNoisy Then _dbg("    ├ 錯誤: ⓪ RDO 略過讀取失敗的資料夾", rdoF.Name)
                        End Try
                    End Sub)

                If validCount = 0 AndAlso rdoFolderList.Count > 0 Then
                    If _iLikeNoisy Then _dbg("    ├ 錯誤: ⓪ RDO 讀取失敗 (無支援的屬性) ", "退回 OOM")
                    Throw New System.Exception("RDO PR_SIZE_EX returned empty for all folders")
                End If
                If _iLikeNoisy Then _dbg("    ├ 結束", $"⓪ RDO平行成功: {rName} | totalSize={grandTotal} | folders={rdoFolderList.Count}")
                Return grandTotal
            Catch ex As OperationCanceledException
                ' 2026/04/15: ParallelOptions.CancellationToken 取消時正常中斷
                If _iLikeNoisy Then _dbg("    ├ 錯誤: ⓪ 已取消", $"{rName}") : Return -1
            Catch ex As System.Exception
                If _iLikeNoisy Then _dbg("    ├ 錯誤: ⓪ RDO平行失敗，走 OOM 循序 fallback", $"{rName} | {ex.Message}")
            Finally
                TryMarshalRelease(rdoRoot)
            End Try
        End If

        ' 2026/3/24 by Gemini: ① OOM 循序 BFS 累加 (避免 STA 錯誤的保險路徑)
        ' 因為 OOM 的 GetTable() 必須在 UI Thread，我們必須循序 Await 每一層
        ' 2026/04/15 by Claude: _cancelRequested 取代為 SmartThrottle(swThrottle, cToken)
        '   GetFolderSizeL3 內部 OCE re-throw → For 迴圈 Catch OCE → Return -1
        '   同時移除舊的 i Mod 5 Await Task.Yield()，統一由 SmartThrottle 每 100ms 讓出
        Try
            ' 2026/04/17 by Claude: 改呼叫 GetSubtreeToList (L2.5)，享有快取加速
            Dim targetFolderList As List(Of (Folder As Folder, fPath As String)) = Await GetSubtreeToList(rootFolder, includeSubF:=True, cToken:=cToken)
            Dim grandTotal As Long = 0
            Dim swThrottle As Stopwatch = Stopwatch.StartNew()  ' by Gemini, 2026/04/02; refactored by Claude Sonnet 4.6, 2026/06/07

            For i As Integer = 0 To targetFolderList.Count - 1
                Dim cFolder As Folder = targetFolderList(i).Folder
                Dim cName As String = ExtractFolderName(targetFolderList(i).fPath)

                ' by Gemini, 2026/04/02: 傳遞 progress 進去以獲得更細緻的(郵件級別)進度回報
                ' 2026/04/15: 同時傳入 cToken，GetFolderSizeL3 取消時 OCE re-throw 冒泡至此
                ' by Gemini, 2026/04/18: 替換 OOM fallback 路徑，從 GetFolderSizeL3() 變更為 GetFolderSizeAsync() (Layer 2.5) 以利用快取
                Dim sz As Long = Await GetFolderSizeAsync(cFolder, fPath:=targetFolderList(i).fPath, cToken:=cToken)
                If sz >= 0 Then grandTotal += sz Else If _iLikeNoisy Then _dbg("    ├ 錯誤: ① 略過了大小計算失敗的資料夾", cName)

                ' by Gemini, 2026/04/02: 100ms 節流回報資料夾級別進度
                ' 2026/04/15 by Claude: SmartThrottle 整合讓出與取消偵測，取代舊的 i Mod 5 Yield
                ' 2026/04/16 by Gemini 3.0 flash: 改用 SmartThrottle 整合進度回報與讓出點
                Await SmartThrottle(swThrottle, cToken:=cToken, ThrottleFreq.Hii,
                                          Sub() progress?.Report(New ProgressReport With {.CurrentCount = i + 1, .TotalCount = targetFolderList.Count,
                                                                                          .Message = $"正在計算大小: {i + 1} / {targetFolderList.Count} ({cName})..."}))
            Next
            If _iLikeNoisy Then _dbg("    ├ 結束", $"① 循序BFS成功: {rName} | totalSize={grandTotal}")
            Return grandTotal
        Catch ex As OperationCanceledException
            ' 2026/04/15: GetSubtreeToList 或 SmartThrottle 或 GetFolderSizeAsync 取消時冒泡至此
            If _iLikeNoisy Then _dbg("    ├ 錯誤: ① 已取消", $"{rName}") : Return -1
        Catch ex As System.Exception
            If _iLikeNoisy Then _dbg("    ├ 錯誤: ① 循序BFS失敗，放棄計算", $"{rName} | {ex.Message}")
        End Try

        ' ② 兩層都失敗，回傳 -1 讓呼叫端知道失敗了
        Return -1
    End Function
    Private Async Function GetYearCountsForFolderL3(folder As Folder, Optional fPath As String = "", Optional cToken As CancellationToken = Nothing) As Task(Of ConcurrentDictionary(Of Integer, Integer))
        ' ---------------------------------------------------------------
        ' === Layer3: COM 資料層 ===
        ' 職責: 對 Outlook 發出 COM 呼叫，回傳單一資料夾的年份郵件分佈
        ' 規則: 不遞迴、不碰 UI、不修改任何全域狀態，
        '       只做一件事: 詢問 Outlook 某資料夾每年有幾封郵件，回傳結果
        '       不遞迴、不知道上層的進度計數、不碰 UI，完全純粹的資料查詢函數
        ' 2026/3/24 by Gemini: 從逐年 Restrict 改為 GetTable + GetArray 一次讀完再記憶體分組
        '   原本每年一次 Restrict + Items.Count = ~30 次 COM call
        '   現在 1 次 GetTable + ceil(N/1000) 次 GetArray，大幅減少 COM 跨程序呼叫
        ' todo: 目前最耗時間的function(), 占整體時間60~65%
        ' 2026/04/15 by Claude: 加入 cToken 參數
        '   取代 _cancelRequested 旗標，改用 SmartThrottle(swThrottle, cToken) 節流讓出
        '   cToken 取消時 Task.Delay 拋 OCE，此函數不攔截 (讓 OCE 冒泡至 CollectYearCounts) 
        '   原因: 攔住後回傳半截 yearCounts，L2 會誤以為該資料夾已統計完畢，導致計數偏低
        ' ---------------------------------------------------------------
        fPath = SafeGetPath(folder, fPath)
        Dim fName As String = ExtractFolderName(fPath)
        If _iLikeNoisy Then _dbg("    ├ 開始", fName)

        ' 2026/3/11再次重構: 優化 COM 呼叫，減少 RCW 物件積累，提升效能和穩定性
        'Dim folderItems As Outlook.Items = Nothing
        Dim yearCounts As New ConcurrentDictionary(Of Integer, Integer)
        Dim table As Outlook.Table = Nothing
        Try
            ' 2026/3/24 by Gemini: 改用 GetTable + GetArray 取代逐年 Restrict
            table = SafeGetTable(folder, "", "ReceivedTime") ' 只讀 RcvTime 一欄，最小化每 row 的傳輸量

            ' by Gemini, 2026/04/05: 每批次讀取後，若超過 100ms 則釋放執行緒並檢查中斷
            ' 2026/04/15 by Claude: _cancelRequested 取代為 SmartThrottle(swThrottle, cToken)
            Dim swThrottle As Stopwatch = Stopwatch.StartNew()  ' by Claude Sonnet 4.6, 2026/06/07
            Do
                Dim data = SafeGetArray(table)
                If data Is Nothing Then Exit Do
                For r As Integer = 0 To data.GetUpperBound(0)
                    Dim receivedTime As DateTime = SafeGet(Of DateTime)(data, r, 0, DateTime.MinValue)
                    If receivedTime > DateTime.MinValue Then
                        Dim year As Integer = receivedTime.Year
                        If year > 0 AndAlso year <= Date.Today.Year Then yearCounts.AddOrUpdate(year, 1, Function(k, v) v + 1)
                    End If
                Next

                Await SmartThrottle(swThrottle, cToken:=cToken)
                ' by Gemini, 2026/04/05: 每 100ms 節流讓出執行緒
                ' 2026/04/15: SmartThrottle 整合讓出與取消偵測，OCE 冒泡至呼叫端 CollectYearCounts
            Loop
        Catch ex As OperationCanceledException
            ' 2026/04/15: 不攔截 OCE，直接 re-throw 讓 CollectYearCounts 感知取消
            If _iLikeNoisy Then _dbg("    ├ 已取消", fName) : Throw
        Catch ex As System.Exception
            If _iLikeNoisy Then _dbg("    ├ 錯誤", $"{fName}: {ex.Message}") ' by Gemini, 2026/04/04: Issue 4 格式標準化
        Finally
            TryMarshalRelease(table)
        End Try
        Await Task.Yield()   ' ✅ 函數結束前再讓出一次，確保畫面有機會更新

        If _iLikeNoisy Then _dbg("    ├ 結束", $"{fName} | 年份分佈: {yearCounts.Count}")
        Return yearCounts

    End Function
    Private Async Function GetMonthCountsForYearL3(folder As Folder, year As Integer, Optional fPath As String = "", Optional cToken As CancellationToken = Nothing) As Task(Of ConcurrentDictionary(Of Integer, Integer))
        ' ---------------------------------------------------------------
        ' GetMonthCountsForYearL3 — Layer3 COM 資料層: 單一資料夾月份郵件分佈
        ' 職責: 對 Outlook 發出 COM 呼叫，回傳單一資料夾在指定年份中每個月的郵件數量
        ' 規則: 不做快取、不做提前過濾、不遞迴 (這些全部交給 GetMonthCountsForYear L2.5 負責)
        '       OCE 不在此攔截，直接 re-throw 讓呼叫端感知取消
        ' 原始設計: 2026/3/24 by Gemini — 從逐月 Restrict 改為 GetTable + GetArray 一次讀完
        '   原本 12 次 Restrict + 12 次 Items.Count = 24 次 COM call
        '   現在 1 次 GetTable (含日期範圍 filter) + ceil(N/1000) 次 GetArray
        ' 2026/04/15 by Claude/Gemini: 加入 cToken 參數與 fPath 參數
        '   由 L2.5 直接傳入 fPath，完全消除 pFolder.FolderPath 的 COM 開銷
        ' 2026/04/17 by Claude: 拆出快取/提前過濾邏輯至 GetMonthCountsForYear (L2.5)，此函數僅剩純 COM
        ' ---------------------------------------------------------------
        fPath = SafeGetPath(folder, fPath)
        Dim fName As String = ExtractFolderName(fPath)
        If _iLikeNoisy Then _dbg("    ├ 開始", $"{fName} ({year} 年)")

        Dim monthCounts As New ConcurrentDictionary(Of Integer, Integer)
        Dim table As Outlook.Table = Nothing
        Try
            ' 2026/3/24 by Gemini: 改用 GetTable + 日期範圍 DASL filter + GetArray
            ' 用整年的日期範圍一次篩選，不再逐月 Restrict
            Dim startDate As New Date(year, 1, 1, 0, 0, 0)
            Dim endDate As New Date(year, 12, 31, 23, 59, 59)
            Dim dateFilter As String = $"[ReceivedTime] >= '{startDate}' AND [ReceivedTime] <= '{endDate}'"
            table = SafeGetTable(folder, dateFilter, "ReceivedTime")

            Dim swThrottle As Stopwatch = Stopwatch.StartNew()  ' by Claude Sonnet 4.6, 2026/06/07
            Do
                Dim data = SafeGetArray(table)
                If data Is Nothing Then Exit Do
                For r As Integer = 0 To data.GetUpperBound(0)
                    Dim receivedTime As DateTime = SafeGet(Of DateTime)(data, r, 0, DateTime.MinValue)
                    If receivedTime > DateTime.MinValue Then monthCounts.AddOrUpdate(receivedTime.Month, 1, Function(k, v) v + 1)
                Next

                Await SmartThrottle(swThrottle, cToken:=cToken)
                ' by simon, 2026/04/08: 每批次讀取後，若超過 100ms 則釋放執行緒並檢查中斷
                ' 2026/04/15 by Claude: SmartThrottle 取代舊的 swThrottle + Task.Delay(1) + _cancelRequested
                ' 2026/04/15: OCE 向上冒泡，不在此攔截 (快取寫入在呼叫端 L2.5，OCE 自然繞過)
            Loop
        Catch ex As OperationCanceledException
            ' 2026/04/15: re-throw 讓 GetMonthCountsForYear (L2.5) 感知，不寫入快取
            If _iLikeNoisy Then _dbg("    ├ 已取消", $"{fName} ({year} 年)") : Throw
        Catch ex As System.Exception
            If _iLikeNoisy Then _dbg("    ├ 錯誤", $"{fName}, year={year}: {ex.Message}") ' by Gemini, 2026/04/04: Issue 4 格式標準化
        Finally
            TryMarshalRelease(table)
        End Try

        If _iLikeNoisy Then _dbg("    ├ 結束", $"{fName} ({year} 年)")
        Return monthCounts
    End Function
    Private Async Function GetAttachMailListL3(folder As Folder, progress As IProgress(Of ProgressReport), Optional cToken As CancellationToken = Nothing) As Task(Of List(Of MailItemInfo))
        ' ----------------------------------------------------------------------------------------
        ' Phase 1 / Layer3 純資料層: GetTable + GetArray 批次掃描單一資料夾
        ' 設計: 這裡只專注於透過 MAPI 取回資料，不會做快取判定，也無關大小設定過濾
        ' 2026/04/15 by Claude: 加入 cToken 參數，取代 _cancelRequested 旗標
        '   SmartThrottle(swThrottle, cToken) 每 100ms 讓出一次，cToken 取消時拋 OCE
        '   取消時捕捉 OCE → 回傳空 List (不回傳已掃到的半截清單) 
        '   原因: 呼叫端 GetAttachMailList 取消時不寫入快取 (見其 cToken.IsCancellationRequested 判斷) 
        ' ----------------------------------------------------------------------------------------
        Dim fName As String = folder.Name ' by Gemini 3.1 Pro 2026/04/16: 避免重複 COM 呼叫
        If _iLikeNoisy Then _dbg(" ├ 開始", fName)

        Const PR_MESSAGE_SIZE As String = "http://schemas.microsoft.com/mapi/proptag/0x0E080003"
        Dim table As Outlook.Table = Nothing

        Dim strFilterHasAttachment As String = "@SQL=" & Chr(34) & "urn:schemas:httpmail:hasattachment" & Chr(34) & " = True"
        ' 預分配容量為 4096，顯著降低掃描大量附件郵件時的記憶體配置開銷 (by Gemini 3 Flash, 2026/05/04)
        Dim result As New List(Of MailItemInfo)(4096)

        ' 2026/04/22 by Gemini 3.1 Pro: 提前取得路徑，讓此資料夾內的所有郵件都能獲得歸屬路徑，且只需 1 次 COM 存取
        Dim fPath As String = ""
        fPath = SafeGetPath(folder)

        Try
            table = SafeGetTable(folder, strFilterHasAttachment, "EntryID", "Subject", PR_MESSAGE_SIZE, "ReceivedTime", "SenderName")

            Dim swThrottle As Stopwatch = Stopwatch.StartNew()  ' by Claude Sonnet 4.6, 2026/06/07
            Dim rowCount As Integer = 0
            Do
                Dim data = SafeGetArray(table)
                If data Is Nothing Then Exit Do
                For r As Integer = 0 To data.GetUpperBound(0)
                    Dim entryID As String = SafeGet(Of String)(data, r, 0, "")
                    If entryID = "" Then Continue For
                    Dim info As New MailItemInfo With {.EntryID = entryID,
                                                       .Subject = SafeGet(Of String)(data, r, 1, ""),
                                                       .Size = SafeGet(Of Long)(data, r, 2, 0L),
                                                       .RcvTime = SafeGet(Of DateTime)(data, r, 3, DateTime.MinValue),
                                                       .SenderName = SafeGet(Of String)(data, r, 4, ""),
                                                       .FolderPath = fPath}
                    result.Add(info)
                    rowCount += 1
                Next

                ' 2026/04/15: 整合 SmartThrottle，取代舊的 _cancelRequested 雙重檢查 + Task.Delay(1)
                ' 2026/04/16 by Gemini 3.0 flash: 改用 SmartThrottle 整合進度回報與讓出點
                Await SmartThrottle(swThrottle, cToken:=cToken, ThrottleFreq.Hii,
                                          Sub() progress?.Report(New ProgressReport With {.Message = $"Phase 1 掃描: {fName} (已找 {result.Count} 封)"}))
            Loop
        Catch ex As OperationCanceledException
            ' 2026/04/15: 取消時回傳空 List，不回傳半截清單，呼叫端不會將不完整結果寫入快取 (是嗎???)
            If _iLikeNoisy Then _dbg("    ├ 已取消", fName)
            Return New List(Of MailItemInfo)()
        Catch ex As System.Exception
            If _iLikeNoisy Then _dbg("    ├ 錯誤: ", fName & " — " & ex.Message)
        Finally
            TryMarshalRelease(table)
        End Try
        If _iLikeNoisy Then _dbg("    ├ 結束", $"找到 {result.Count} 封有附件郵件")
        Return result
    End Function
    Private Function GetAttachFilenameL3(ByRef mail As MailItemInfo) As List(Of String)
        ' by Gemini, 2026/04/04: 取得郵件的附件檔名清單 (純 Layer3 邏輯，不做快取)
        If _iLikeNoisy Then _dbg("    ├ 開始", mail.Subject)
        Dim result As New List(Of String)(4096)

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
                If _iLikeNoisy Then _dbg("    ├ ⓪ RDO 失敗，走OOM fallback", ex.Message)
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
                Dim attCount As Integer = attachments.Count                     ' by simon 2026/04/19: 存成變數避免 COM 呼叫重複
                Dim olbyValue As Integer = Outlook.OlAttachmentType.olByValue   ' by simon 2026/04/19: 存成變數避免 COM 呼叫重複
                For i As Integer = 1 To attCount ' COM 是 1-based index
                    Dim att As Outlook.Attachment = attachments.Item(i)
                    Try : If att.Type = olbyValue Then result.Add(att.FileName) ' 2026/04/09 by Gemini: 僅處理 olByValue (1) 類型的附件
                    Finally : TryMarshalRelease(att)
                    End Try
                Next
            End If
        Catch ex As System.Exception
            If _iLikeNoisy Then _dbg("    ├ ① OOM 失敗", ex.Message)
        Finally
            If _iLikeNoisy Then _dbg(" ├ 結束")
            TryMarshalRelease(attachments)
            TryMarshalRelease(tempMail)
        End Try

        Return result
    End Function
    Private Function GetMailBodyL3(entryID As String) As String
        ' ── L3 COM 資料層 ────────────────────────────────────────────────────────────
        ' ---------------------------------------------------------------
        ' GetMailBodyL3 — Layer3 COM 資料層：讀取郵件 Body 並正規化
        ' 2026/04/28 by Simon/Claude: 以 Simon 的 GetMailBodyByEntryID 為基礎
        '   + 加入 NormalizeMailBody 正規化（去除 HTML 標籤、空白換行）
        '   + Await Task.Yield() 確保每封讀完後讓 UI 執行緒喘氣
        '   支援 MailItem 與 PostItem 兩種型別
        '   使用獨立的 ns（Simon 的設計），確保 COM namespace 不跨執行緒共用
        ' ---------------------------------------------------------------
        If String.IsNullOrEmpty(entryID) Then Return ""

        Dim ns As Outlook.NameSpace = Nothing
        Dim item As Object = Nothing
        Dim body As String = ""
        Try
            ns = _olApp.GetNamespace("MAPI")    ' 2026/04/26 by Gemini, 使用自己內部的 NameSpace 以更好封裝, 並自行TryMarshalRelease以減少GCW洩漏
            item = ns.GetItemFromID(entryID)
            'item = _olNS.GetItemFromID(entryID) ' 2026/04/28 by simon, 使用共用的 NameSpace 以減少多建一次namespace的 COM 開銷

            If item IsNot Nothing Then
                If TypeOf item Is Outlook.MailItem Then
                    body = NormalizeMailBody(DirectCast(item, Outlook.MailItem).Body)
                ElseIf TypeOf item Is Outlook.PostItem Then
                    body = NormalizeMailBody(DirectCast(item, Outlook.PostItem).Body)
                End If
            End If
        Catch ex As System.Exception
            _dbg("GetMailBodyL3 失敗", $"{entryID}: {ex.Message}")
        Finally
            TryMarshalRelease(item)
            TryMarshalRelease(ns)   ' 2026/04/26 by Gemini, 使用自己內部的 NameSpace 以更好封裝, 並自行TryMarshalRelease以減少GCW洩漏
        End Try
        ' 2026/05/09 by Gemini 3.0 flash: 移除內部的 Yield。改由調用端依批次執行呼吸，減少微切換開銷提升讀取性能
        Return body

    End Function
    Private Async Function GetBasicMailInfoL3(folder As Folder, needTopic As Boolean, cToken As CancellationToken, Optional fPath As String = "") As Task(Of List(Of (Mail As MailItemInfo, Topic As String)))
        ' ---------------------------------------------------------------
        ' 2026/05/06 by Claude: 永遠讀取全部 8 欄（含 topic/msgId/senderEmail）
        '   needTopic 參數保留供 API 相容，但 L3 層已不區分，統一讀取
        '   欄位索引: 0=EntryID, 1=Subject, 2=Size, 3=RcvTime,
        '             4=SenderName, 5=Topic, 6=MsgIDhash, 7=SenderEmail
        '
        ' 2026/06/12 by Simon/Claude: 移除 PR_CONVERSATION_TOPIC (欄位 5) topic 改由 GetCleanSubject(subject) 動態計算，與 DB 讀取路徑保持一致
        '   欄位索引: 0=EntryID, 1=Subject, 2=Size, 3=RcvTime, 4=SenderName, 5=MsgIDhash, 6=SenderEmail
        ' ---------------------------------------------------------------

        fPath = SafeGetPath(folder, fPath)
        Dim fName As String = ExtractFolderName(fPath)
        If _iLikeNoisy Then _dbg("    ├ 開始 (掃描)", fName)

        Dim resultList As New List(Of (MailItemInfo, String))(4096) ' 預分配容量為 4096，優化批次讀取郵件基本資訊時的清單填充 (by Gemini 3 Flash, 2026/05/04)
        Dim table As Outlook.Table = Nothing
        Try
            Const PR_SENDER_EMAIL As String = "http://schemas.microsoft.com/mapi/proptag/0x0C1F001E"
            Const PR_MESSAGE_SIZE As String = "http://schemas.microsoft.com/mapi/proptag/0x0E080003"
            Const PR_CONVERSATION_TOPIC As String = "http://schemas.microsoft.com/mapi/proptag/0x0070001E"
            Const PR_INTERNET_MESSAGE_ID As String = "http://schemas.microsoft.com/mapi/proptag/0x1035001E"

            table = SafeGetTable(folder, "",
                                 "EntryID", "Subject", PR_MESSAGE_SIZE,     ' 0~2
                                 "ReceivedTime", "SenderName",              ' 3~4
                                 PR_INTERNET_MESSAGE_ID, PR_SENDER_EMAIL)   ' 5~6 (2026/05/06 by Claude: Tab5 去重)

            Dim swYield As Stopwatch = Stopwatch.StartNew()  ' by Claude Sonnet 4.6, 2026/06/07
            Do
                Dim data = SafeGetArray(table)
                If data Is Nothing Then Exit Do
                For r As Integer = 0 To data.GetUpperBound(0)
                    Dim entryID As String = SafeGet(Of String)(data, r, 0, "")
                    If entryID = "" Then Continue For

                    Dim subj As String = SafeGet(Of String)(data, r, 1, "")
                    Dim mail As New MailItemInfo With {.EntryID = entryID,
                                                       .Subject = subj,
                                                       .Size = SafeGet(Of Long)(data, r, 2, 0L),
                                                       .RcvTime = SafeGet(Of DateTime)(data, r, 3, DateTime.MinValue),
                                                       .SenderName = SafeGet(Of String)(data, r, 4, ""),
                                                       .FolderPath = fPath,
                                                       .MsgIDhash = StringToXxHash64Hex(SafeGet(Of String)(data, r, 5, "")),
                                                       .SenderEmail = SafeGet(Of String)(data, r, 6, "")}
                    ' 2026/06/12 by Simon/Claude: topic 從 GetCleanSubject(subject) 動態計算
                    resultList.Add((mail, GetCleanSubject(subj)))
                Next
                Await SmartThrottle(swYield, cToken, ThrottleFreq.Mid)  ' ✅ 使用統一節流讓位，內部 Task.Delay(1) 確保 ESC 中斷靈敏度 (by Antigravity, 2026/04/19)
            Loop

        Catch ex As System.Exception
            If _iLikeNoisy Then _dbg("    ├ 錯誤", $"{fName}: {ex.Message}")
        Finally
            If _iLikeNoisy Then _dbg("    ├ 結束 (掃描)")
            TryMarshalRelease(table)
        End Try
        Return resultList
    End Function
    Private Async Function GetSubtreeToListL3(rootFolder As Folder, includeSubF As Boolean, Optional progress As IProgress(Of ProgressReport) = Nothing, Optional cToken As CancellationToken = Nothing) As Task(Of List(Of (folder As Folder, fPath As String)))
        ' --------------------------------------------------------------
        ' GetSubtreeToListL3 — Layer3 COM 資料層: BFS 純掃描
        ' 原名 GetSubtreeToList，2026/04/17 by Claude: 拆出快取邏輯至 GetSubtreeToList (L2.5)
        ' 職責: BFS 廣度優先掃描 rootFolder 下整棵子樹，回傳 (Folder, fPath) Tuple 清單
        ' 規則: 不做快取讀取；BFS 完成後若未中斷，由本層負責寫入 _cacheSubTreeList
        '       OCE 中斷時 re-throw，不寫快取 (確保不存入不完整的樹)
        ' 2026/04/16 by Gemini: 升級回傳 Tuple (Folder, fPath)，消除呼叫端對 COM .FolderPath 的依賴
        ' --------------------------------------------------------------
        ' 2026/04/24 by Gemini 3.0 flash: 使用 SafeGetPath 並增加 root 狀態檢查
        Dim rootPath As String = SafeGetPath(rootFolder)
        If String.IsNullOrEmpty(rootPath) Then
            _dbg(" ├ 錯誤", "無法取得 rootFolder 路徑，中斷掃描")
            Return New List(Of (folder As Folder, fPath As String))
        End If

        Dim rootName As String = ExtractFolderName(rootPath)
        Dim cacheKey As String = rootPath & "|" & _showAllFolders  ' 2026/04/17: 鍵值含 _showAllFolders 分支

        ' 2026/6/2: 再次修正F5 強制刷新的總數讀取不正確
        ' 🔽🔽🔽 【修復點 1】補上 rootFolder 自己的身分證註冊！ 🔽🔽🔽
        Try
            Dim isRootMail As Boolean = IsMailFolder(rootFolder, rootPath)
            _cacheFolderIDs.TryAdd(rootPath, (rootFolder.EntryID, rootFolder.StoreID, isRootMail, TextHasChineseChar(rootName)))
        Catch : End Try
        ' 🔼🔼🔼 🔼🔼🔼

        If _iLikeNoisy Then _dbg(" ├ 開始", rootName)
        Dim sw As Stopwatch = Stopwatch.StartNew()          ' by Claude Sonnet 4.6, 2026/06/07
        Dim swThrottle As Stopwatch = Stopwatch.StartNew()  ' by Claude Sonnet 4.6, 2026/06/07

        ' 預分配容量為 512，足以涵蓋 90% 以上用戶的資料夾數量，避免 BFS 過程中的陣列頻繁 Resize 開銷 (by Gemini 3 Flash, 2026/05/04)
        Dim result As New List(Of (folder As Folder, fPath As String))(512)
        result.Add((rootFolder, rootPath))

        If Not includeSubF Then
            sw.Stop()
            If _iLikeNoisy Then _dbg("    ├ 結束", $"{rootName} (Single) | {sw.ElapsedMilliseconds}ms")
            Return result
        End If

        ' BFS COM 掃描 (快取 miss 時走此路徑)
        Dim queue As New Queue(Of (folder As Folder, Path As String))(512)
        queue.Enqueue((rootFolder, rootPath))
        Try
            While queue.Count > 0
                Dim current = queue.Dequeue()
                Try
                    ' 優化第六點：提取 Folders 集合並在 Finally 顯式釋放，防止 BFS 過程中的 RCW 洩漏 (by Gemini 3 Flash, 2026/05/05)
                    Dim subFolders As Folders = current.folder.Folders
                    Try
                        For Each subF As Folder In subFolders
                            ' 2026/04/24 by Gemini 3.0 flash: 將可能發生 COM 崩潰的讀取點全部加上安全保護
                            Dim fName As String = ""
                            Try : fName = subF.Name : Catch : Continue For : End Try

                            ' ✅ 通過參數傳遞 childPath，IsMailFolder 內部不再重複呼叫 COM
                            Dim childPath As String = current.Path & "\" & fName
                            Dim isMail As Boolean = IsMailFolder(subF, childPath)
                            If Not _showAllFolders AndAlso Not isMail Then Continue For

                            ' ✅ 加強 EntryID/StoreID 讀取的安全性
                            Try
                                _cacheFolderIDs.TryAdd(childPath, (subF.EntryID, subF.StoreID, isMail, TextHasChineseChar(fName)))
                            Catch : End Try

                            result.Add((subF, childPath))   ' ✅ 同步存入預計好的路徑，不再打 COM
                            queue.Enqueue((subF, childPath))
                        Next
                    Finally
                        TryMarshalRelease(subFolders)
                    End Try
                Catch ex As System.Exception
                    If _iLikeNoisy Then _dbg("    ├ ① OOM 失敗", ExtractFolderName(current.Path) & " - " & ex.Message)
                End Try

                Await SmartThrottle(swThrottle, cToken:=cToken, ThrottleFreq.Low,
                                          Sub() progress?.Report(New ProgressReport With {.CurrentCount = result.Count, .Message = $"正在展開資料夾結構: 已發現 {result.Count} 個資料夾..."}))
            End While
        Catch ex As OperationCanceledException
            If _iLikeNoisy Then _dbg("    ├ 中斷", $"GetSubtreeToListL3 已由使用者中斷，已發現 {result.Count} 個")
            Throw   ' re-throw，確保不完整的樹不寫入快取
        Catch ex As System.Exception
            If _iLikeNoisy Then _dbg("    ├ ② OOM BFS 失敗", ex.Message)
        End Try
        sw.Stop()

        ' BFS 完成且未中斷時寫入快取 (由 L3 自行負責，L2.5 不重複寫)
        If Not cToken.IsCancellationRequested AndAlso result.Count > 0 Then _cacheSubTreeList.TryAdd(cacheKey, result)
        If _iLikeNoisy Then _dbg("    ├ 結束", $"{rootName} (BFS) | 資料夾總計: {result.Count} | {sw.ElapsedMilliseconds}ms")
        Return result

    End Function
    Private Function GetSubtreeToListL3_Rdo(rootFolder As Redemption.RDOFolder, includeSubF As Boolean) As List(Of Redemption.RDOFolder)
        ' --------------------------------------------------------------
        ' 2026/3/24 by Gemini: GetSubtreeToListL3_Rdo
        ' 目的: 專門提供給 RDO 平行路徑使用，回傳 List(Of Redemption.RDOFolder)
        ' 說明: 因為 Redemption 是 free-threaded，可以用 Parallel.ForEach 安全平行展開子樹
        ' --------------------------------------------------------------
        Dim rootName As String = rootFolder.Name
        _dbg("    ├ 開始", rootName)
        Dim sw As Stopwatch = Stopwatch.StartNew()  ' by Claude Sonnet 4.6, 2026/06/07

        Dim resultBag As New ConcurrentBag(Of Redemption.RDOFolder)
        resultBag.Add(rootFolder)
        If Not includeSubF Then
            sw.Stop()
            _dbg("    ├ 結束", $"{rootName} (RDO-Single) | {sw.ElapsedMilliseconds}ms") ' by Gemini, 2026/04/10
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
            Parallel.ForEach(layerList, Sub(current)
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
        _dbg("    ├ 結束", $"{rootName} (RDO-Parallel BFS) | 資料夾總計: {resultBag.Count} | {sw.ElapsedMilliseconds}ms") ' by Gemini, 2026/04/10
        Return resultBag.ToList()

    End Function
    Private Function GetLiveFolderSnapL3(folder As Folder, Optional fPath As String = "") As Integer
        ' ---------------------------------------------------------------
        ' 快速讀取 PR_CONTENT_COUNT，專門只用於 SQLite snapshot 驗證
        ' 故意不走完整 Layer3 fallback 的GetMailCount，只走最快的 PropertyAccessor 路徑
        ' 失敗時回傳 -999 (不可能等於任何正常 snapshot 值，確保快取失效) 
        ' 2026/4/7 by Gemini, 解決 SSD 讀出後 snapshot 驗證失敗導致的重複統計問題
        ' ---------------------------------------------------------------
        Dim fName As String = ExtractFolderName(fPath)
        If _iLikeNoisy Then _dbg(" ├ 開始", fName)
        Try
            Const PR_CONTENT_COUNT As String = "http://schemas.microsoft.com/mapi/proptag/0x36020003"
            Return CLng(folder.PropertyAccessor.GetProperty(PR_CONTENT_COUNT))
        Catch
            Try : Return folder.Items.Count : Catch : Return -999 : End Try
        End Try
    End Function
#End Region
#Region "  ├ Legacy 保留（暫無呼叫端）"
    Private Function GetMailSizeL3(item As Object) As Long
        ' --------------------------------------------------------------
        ' GetMailSizeL3: 讀取單封郵件的大小 (bytes)，供 GetFolderSizeL3 fallback 路徑呼叫
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
        ' 注意: 此函數接受 Object 型別參數，是因為 GetFolderSizeL3 的 fallback 路徑
        '       用 Items.GetFirst/GetNext 取回的是 Object，省去呼叫端的 TryCast 成本
        '       若是 MailItem 就正常讀取，若是其他型別 (Contact、Appointment 等) 就回 0
        '
        ' 取代: GetFolderSizeOld 內的 mailItem.Size 直接呼叫 行3385 的同名 stub (完整替換)
        ' ---------------------------------------------------------------
        ' 2026/4/28 by simon: 目前此函數已成死碼, 沒有任何呼叫端與進入點
        ' 原始設計意圖:
        ' 	呼叫端 → GetFolderSizeAll → GetFolderSize → GetMailSizeL3 (L3)
        ' 後來使用GetTable.GetArray() 直接整個目錄的table一起讀出來在記憶體內運算
        ' 	(自己計算並直接寫入 _cacheFolderSize，完全繞過此L3層級函數)
        ' ---------------------------------------------------------------

        ' 非 MailItem 的項目 (Calendar、Contact 等) 直接略過，回 0
        If _iLikeNoisy Then _dbg("    ├ 開始")
        Dim mail As Outlook.MailItem = TryCast(item, Outlook.MailItem)
        If mail Is Nothing Then Return 0

        ' ⓪ Redemption: RDOMail.Size
        '   GetMessageFromID 的 StoreID 從 mail.Parent 取得，多一次 COM call 但避免跨 PST 找錯 item
        '   2026-03-22 新增
        If _rdo IsNot Nothing Then
            Dim rdoMail As Redemption.RDOMail = Nothing
            Try
                Dim parentFolder As Folder = TryCast(mail.Parent, Folder)
                Dim storeId As String = If(parentFolder?.StoreID, "")
                rdoMail = TryCast(_rdo.GetMessageFromID(mail.EntryID, storeId), Redemption.RDOMail)
                If rdoMail IsNot Nothing Then
                    Dim sz As Long = CLng(rdoMail.Size)
                    If _iLikeNoisy Then _dbg("    ├ ⓪ RDO 成功", $"size={sz}") ' 高頻率項目平時不輸出 Log
                    Return sz
                End If
            Catch ex As System.Exception
                If _iLikeNoisy Then _dbg("    ├ ⓪ RDO 失敗，走MAPI fallback", ex.Message) ' by Gemini, 2026/04/10: 調整縮排層級為 Level 2 (內部失敗路徑)
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
            If _iLikeNoisy Then _dbg("    ├ ① PR_MESSAGE_SIZE_EXTENDED失敗", ex.Message) ' by Gemini, 2026/04/10
        End Try

        ' ② MAPI: PR_MESSAGE_SIZE (0x0E080003, PT_LONG) — 32-bit，超大郵件理論上溢位
        Try
            Const PR_SIZE As String = "http://schemas.microsoft.com/mapi/proptag/0x0E080003"
            Return CLng(mail.PropertyAccessor.GetProperty(PR_SIZE))             ' by Gemini, 2026/03/29: 同上，移除 TypeOf 判斷
        Catch ex As System.Exception
            If _iLikeNoisy Then _dbg("    ├ ② PR_MESSAGE_SIZE失敗", ex.Message) ' by Gemini, 2026/04/10
        End Try

        ' ③ OOM: mail.Size (Integer，超大郵件理論上不準，但實務上 PST 內不會發生)
        Try
            Return CLng(mail.Size)
        Catch ex As System.Exception
            If _iLikeNoisy Then _dbg("    ├ ③ OOM mail.Size也失敗", ex.Message) ' by Gemini, 2026/04/10
        End Try
        Return -1

    End Function
    Private Async Function GetMailCountAllAsync(folder As Folder, Optional progress As IProgress(Of ProgressReport) = Nothing, Optional fPath As String = "", Optional cToken As CancellationToken = Nothing) As Task(Of Long)
        ' ---------------------------------------------------------------
        ' GetMailCountAllAsync — 整棵子樹的郵件總數
        ' 2026/4/7 by Claude, 加入Database lazy load, 讀資料順序:
        '   ① 記憶體命中
        '   ② DB lazy load：命中且 mca 有效且 snapshot 吻合 → 一次填滿所有欄位
        '   ③ fallback: Layer3 呼叫
        ' 2026/04/15 by Claude: 加入 cToken 參數，取代 _cancelRequested 旗標
        ' 2026/04/15 by Gemini: 加入 optional fPath
        ' ---------------------------------------------------------------
        ' 2026/4/28 by simon: 目前GetxxxCountAll系列函數已成死碼, 沒有任何呼叫端與進入點
        ' 原始設計意圖:
        ' 	呼叫端 → GetMailCountAllAsync (L2.5) → GetMailCountAllL3 (L3)
        ' 	呼叫端 → GetFolderCountAllAsync (L2.5) → GetFolderCountAllL3 (L3)
        '
        ' 後來使用了BFS剪枝速度更快:
        ' 	Compute → BFS → SummarizeSubTreeBottomUp → UpdateFolderStatsCache
        ' 	(自己計算並直接寫入 _cacheMailCountAll / _cacheFolderCountAll，完全繞過 AllAsync系列函數)
        ' 2026/05/09 by Gemini 3.1 Pro: 重構成使用 SafeGetDbRow，回傳統一改成 Long 共用函式
        ' ---------------------------------------------------------------
        fPath = SafeGetPath(folder, fPath)
        If _iLikeNoisy Then _dbg(" ├ 開始", ExtractFolderName(fPath))
        Dim count As Long : If _cacheMailCountAll.TryGetValue(fPath, count) Then Return count  ' ① 記憶體命中

        ' ② DB lazy load (mca 欄位) 
        Dim row = SafeGetDbRow(folder, fPath)
        If row IsNot Nothing AndAlso row.mca >= 0 Then Return row.mca

        ' ③ fallback: Layer3 呼叫
        Dim total As Long = Await GetMailCountAllL3(folder, progress, cToken:=cToken)
        If total >= 0 Then _cacheMailCountAll.TryAdd(fPath, total)  ' 2026/04/15: 改用 cToken 判斷，取代 _cancelRequested
        Return total

    End Function
    Private Async Function GetFolderCountAllAsync(folder As Folder, Optional fPath As String = "", Optional cToken As CancellationToken = Nothing) As Task(Of Long)
        ' ---------------------------------------------------------------
        ' GetFolderCountAllAsync — 整棵子樹的資料夾總數
        ' 2026/4/7 by Claude, 加入Database lazy load, 讀資料順序:
        '   ① 記憶體命中
        '   ② DB lazy load：命中且 fca 有效且 snapshot 吻合 → 一次填滿所有欄位
        '   ③ fallback: Layer3 呼叫
        ' 2026/04/15 by Claude: 加入 cToken 參數，取代 _cancelRequested 旗標
        ' ---------------------------------------------------------------
        ' 2026/4/28 by simon: 目前GetxxxCountAll系列函數已成死碼, 沒有任何呼叫端與進入點
        ' 原始設計意圖:
        ' 	呼叫端 → GetMailCountAllAsync (L2.5) → GetMailCountAllL3 (L3)
        ' 	呼叫端 → GetFolderCountAllAsync (L2.5) → GetFolderCountAllL3 (L3)
        '
        ' 後來使用了BFS剪枝速度更快:
        ' 	Compute → BFS → SummarizeSubTreeBottomUp → UpdateFolderStatsCache
        ' 	(自己計算並直接寫入 _cacheMailCountAll / _cacheFolderCountAll，完全繞過 AllAsync系列函數)
        ' 2026/05/09 by Gemini 3.1 Pro: 重構成使用 SafeGetDbRow，回傳統一改成 Long 共用函式
        ' ---------------------------------------------------------------
        fPath = SafeGetPath(folder, fPath)
        If _iLikeNoisy Then _dbg(" ├ 開始", ExtractFolderName(fPath))
        Dim count As Long : If _cacheFolderCountAll.TryGetValue(fPath, count) Then Return count  ' ① 記憶體命中

        ' ② DB lazy load (fca 欄位) 
        Dim row = SafeGetDbRow(folder, fPath)
        If row IsNot Nothing AndAlso row.fca >= 0 Then Return row.fca

        ' ③ fallback: Layer3 呼call
        count = Await GetFolderCountAllL3(folder, cToken:=cToken)
        If count >= 0 Then _cacheFolderCountAll.TryAdd(fPath, count)  ' 2026/04/15: 改用 cToken 判斷
        Return count
    End Function
#End Region
#Region "  └ 其他輔助函數"
    Private Shared Function SafeGetPath(folder As Folder, Optional existingPath As String = "") As String
        ''' <summary>
        ''' 安全取得資料夾路徑 (FolderPath)，防止 COM 物件失效或 Nothing 引發的例外。
        ''' </summary>
        ''' <param name="folder">Outlook 資料夾物件</param>
        ''' <param name="existingPath">選擇性：若已存在路徑則直接回傳，減少 COM 呼叫開銷</param>
        ' by Gemini 3.0 Flash, 2026/04/23
        ' 邏輯：優先使用傳入的路徑，若無則嘗試從物件讀取並捕捉所有潛在例外
        If Not String.IsNullOrEmpty(existingPath) Then Return existingPath
        If folder Is Nothing Then Return ""
        Try
            Return folder.FolderPath
        Catch
            ' 捕捉 RCW 已釋放或物件失效的例外
            Return ""
        End Try
    End Function
    Private Function SafeGetDbRow(folder As Folder, fPath As String) As FolderStatsDbRow
        ''' <summary>
        ''' [Helper] 嘗試從資料庫取得有效的資料夾統計列，並自動處理快取回填。
        ''' </summary>
        ''' <param name="folder">Outlook 資料夾物件</param>
        ''' <param name="fPath">資料夾路徑</param>
        ''' <returns>有效的 FolderStatsDbRow，若無命中或 Snapshot 不符則回傳 Nothing</returns>
        ' 2026/05/09 by Gemini 3.1 Pro: 統一處理 DB lazy load 與 Snapshot 驗證邏輯
        ' snapshot 驗證: DB 儲存的 content_count_snap = save 時的 PR_CONTENT_COUNT 值
        '   用 GetLiveFolderSnapL3 (單次 PropertyAccessor call) 與 snapshot 比對
        '   相同 → 快取仍有效；不同 → 資料夾內容已變，跳過 DB，呼叫 Layer3
        Dim row = DbGetFolderStats(fPath)
        If row IsNot Nothing AndAlso GetLiveFolderSnapL3(folder, fPath) = row.snap Then
            FillCacheFromDbRow(fPath, row) : Return row
        End If
        Return Nothing
    End Function
    Private Function SafeGetTable(folder As Folder, filter As String, ParamArray cols As String()) As Outlook.Table
        ' 2026/05/09 by Simon/Claude: 封裝 GetTable + Columns.RemoveAll + Columns.Add 的固定初始化模板
        ' filter 傳空字串時不套用篩選；cols 用 ParamArray 接受任意數量的欄位名稱
        Dim t As Outlook.Table = If(String.IsNullOrEmpty(filter), folder.GetTable(), folder.GetTable(filter))
        t.Columns.RemoveAll()
        For Each col In cols : t.Columns.Add(col) : Next
        Return t
    End Function
    Private Function SafeGetArray(table As Outlook.Table) As Object(,)
        ' 2026/05/09 by Simon/Claude: 封裝 GetArray + DirectCast 的固定取批次模板
        ' 回傳 Nothing 代表已到結尾或批次為空，呼叫端以 If data Is Nothing Then Exit Do 結束迴圈
        If table.EndOfTable Then Return Nothing

        Const BATCH_SIZE As Integer = 1000  ' 2026/3/24 by Gemini: GetTable.GetArray() 的時候每次批量讀取的筆數
        Dim arr As Object = table.GetArray(BATCH_SIZE)
        If arr Is Nothing Then Return Nothing
        Return DirectCast(arr, Object(,))
    End Function
    Private Sub FillCacheFromDbRow(fPath As String, row As FolderStatsDbRow, Optional skipAggregates As Boolean = False)
        ' DB 命中且 snapshot 驗證通過時，一次填滿所有欄位
        ' 使用 TryAdd：記憶體已有值 (例如另一個 Layer2.5 函數剛填入) 時不覆蓋
        ' -1 代表 DB 中該欄位尚未存入 (例如 mca 還沒算過)，跳過，不污染記憶體快取
        '
        ' by Claude Sonnet 4.6, 2026/04/25: 加入 skipAggregates 參數
        '   skipAggregates = True 時跳過 mca/fca/fsa 三個帶有模式語意的聚合欄位。
        '   原因：mca/fca/fsa 是「含子孫的加總值」，計算時依賴 _showAllFolders 決定走訪哪些子資料夾。
        '   若 DB 中存的是另一個模式下計算的值，填入記憶體快取後 BFS 就會用它剪枝並直接採用 —
        '   導致切換 _showAllFolders 後或重新啟動後，第一次統計顯示舊模式的加總數字（Bug）。
        '   設 skipAggregates=True 讓 BFS 自行展開重算，結果才寫入記憶體（確保模式正確）。
        With row
            If .mc >= 0 Then _cacheMailCount.TryAdd(fPath, .mc)
            If .fc >= 0 Then _cacheFolderCount.TryAdd(fPath, .fc)
            If .fs >= 0 Then _cacheFolderSize.TryAdd(fPath, .fs)

            If Not skipAggregates Then
                ' 只有呼叫端確定 DB 的 mca/fca/fsa 與當前 _showAllFolders 模式一致時才填入
                ' (例如：此 session 中已用相同模式計算過，記憶體命中後又從記憶體拿到，不是從 DB 拿的)
                If .mca >= 0 Then _cacheMailCountAll.TryAdd(fPath, .mca)
                If .fca >= 0 Then _cacheFolderCountAll.TryAdd(fPath, .fca)
                If .fsa >= 0 Then _cacheFolderSizeAll.TryAdd(fPath, .fsa)
            End If

            ' by Gemini, 2026/04/10: 填充身分標識與標籤快取
            If Not String.IsNullOrEmpty(.eid) Then _cacheFolderIDs.TryAdd(fPath, (.eid, .sid, .isMail = 1, .hasCh = 1))
            If .isMail >= 0 Then _cacheIsMailFolder.TryAdd(fPath, .isMail = 1)
        End With
    End Sub

    Private Function ExtractFolderName(fPath As String) As String
        ' ---------------------------------------------------------------
        ' ExtractFolderName: 從 FolderPath 字串中解析出資料夾名稱
        ' 邏輯: 取出最後一個 "\" 之後的內容
        ' by Gemini 3.0 flash, 2026/04/15
        ' ---------------------------------------------------------------
        If String.IsNullOrEmpty(fPath) Then Return "未知"
        Dim idx As Integer = fPath.LastIndexOf("\"c)
        If idx < 0 Then Return fPath
        Return fPath.Substring(idx + 1)
    End Function
    Private Function GetParentPath(fPath As String) As String
        ' ---------------------------------------------------------------
        ' GetParentPath: 從 FolderPath 字串中取得其父層路徑
        ' 邏輯: 移除最後一個 "\" 及其之後的內容
        ' by Gemini 3.0 flash, 2026/04/24
        ' ---------------------------------------------------------------
        If String.IsNullOrEmpty(fPath) Then Return ""
        Dim idx As Integer = fPath.LastIndexOf("\"c)
        If idx < 0 Then Return "" ' 已經是根層級或無斜線
        Return fPath.Substring(0, idx)
    End Function
    Private Function GetAncestors(fPath As String) As List(Of String)
        ' ---------------------------------------------------------------
        ' GetAncestors: 取得路徑的所有祖先清單 (由近到遠)
        ' 範例: "\\A\B\C" -> {"\\A\B", "\\A"}
        ' by Gemini 3.0 flash, 2026/04/24
        ' ---------------------------------------------------------------
        Dim ancestors As New List(Of String)(16)
        Dim current = GetParentPath(fPath)
        While Not String.IsNullOrEmpty(current)
            ancestors.Add(current)
            current = GetParentPath(current)
        End While
        Return ancestors
    End Function
    Private Function SafeGet(Of T)(row As Outlook.Row, column As String, defaultValue As T) As T
        ''' <summary>
        ''' 安全地從 Outlook.Row 讀取欄位，自動處理 Nothing / DBNull / 例外
        ''' 2026/04/01 by Gemini
        ''' </summary>
        Try
            Dim value = row(column)
            If value Is Nothing OrElse IsDBNull(value) Then Return defaultValue
            Return CType(value, T)
        Catch ex As System.Exception
            _dbg("       ├ SafeGet(Row) 失敗", $"{column} | {ex.Message}") ' by Gemini, 2026/04/11: 底層詳細 Level 3
            Return defaultValue
        End Try

    End Function
    Private Function SafeGet(Of T)(data(,) As Object, row As Integer, col As Integer, defaultValue As T) As T
        ''' <summary>
        ''' SafeGet 的二維陣列 (GetArray) Overload 版
        ''' 2026/04/01 by Gemini
        ''' </summary>
        Try
            Dim value = data(row, col)
            If value Is Nothing OrElse IsDBNull(value) Then Return defaultValue
            ' 使用 Convert.ChangeType 確保數值型態 (如 Long/Int/DateTime) 能正確轉換
            ' 2026/6/6 by Gemini: 加上 Fast Path：如果型別已經一致，跳過 ChangeType 昂貴的反射
            If TypeOf value Is T Then Return DirectCast(value, T)
            Return CType(Convert.ChangeType(value, GetType(T)), T)
        Catch
            Return defaultValue
        End Try

    End Function
    Private Function TextHasChineseChar(name As String) As Boolean
        'Return name.Any(Function(c) c >= ChrW(&H4E00) AndAlso c <= ChrW(&H9FFF))
        For Each c In name : If c >= ChrW(&H4E00) AndAlso c <= ChrW(&H9FFF) Then Return True
        Next : Return False
    End Function
    Private Function HasSubFoldersFast(cFolder As Folder, Optional fPath As String = "") As Boolean
        ' ---------------------------------------------------------------
        ' HasSubFoldersFast — 光速版子資料夾加號預測 (專為 TreeView 展開設計)
        ' 2026/4/7 by Gemini, 解決 SSD 讀出後無法出現假節點 + 號，以及嚴重卡頓問題
        ' ---------------------------------------------------------------
        '   呼叫順序：① _cacheFolderCount 記憶體 → ② DbGetFolderStats(fPath).fc → ③ pFolder.Folders.Count COM
        '   已在 LoadSubFolderToTreeView 第 489 行啟用， 解決 DB 載入後 TreeView 不顯示 "+" 的問題
        '   比直接 pFolder.Folders.Count 快： 記憶體命中~0μs，DB命中~0.1ms，COM才~1-5ms
        ' ---------------------------------------------------------------
        fPath = SafeGetPath(cFolder, fPath)
        If fPath = "" Then Return False     ' 2026/04/23 by Gemini 3.0 flash: 確保路徑有效，抓不到代表資料夾異常

        Dim fc As Long
        If _cacheFolderCount.TryGetValue(fPath, fc) Then Return fc > 0

        Dim row = DbGetFolderStats(fPath)
        If row IsNot Nothing AndAlso row.fc >= 0 Then
            _cacheFolderCount.TryAdd(fPath, row.fc) ' 把確認的值送回記憶體快取
            Return row.fc > 0
        End If

        ' 萬一都沒有，直接保底呼叫一次 COM (比 PR_CONTENT_COUNT 驗證還快)
        Try : Return cFolder.Folders.Count > 0 : Catch : Return False : End Try
    End Function
    Private Function IsMailFolder(folder As Folder, Optional fPath As String = "") As Boolean
        fPath = SafeGetPath(folder, fPath)
        Dim fName As String = ExtractFolderName(fPath)
        If _iLikeNoisy Then _dbg(" ├ 開始", fName)
        Dim isMail As Boolean
        If _cacheIsMailFolder.TryGetValue(fPath, isMail) Then Return isMail

        Static allowedTypes As Outlook.OlItemType() = {Outlook.OlItemType.olMailItem,
                                                       Outlook.OlItemType.olPostItem}
        Try
            Dim itemType As Outlook.OlItemType = folder.DefaultItemType
            isMail = allowedTypes.Contains(itemType)
            _cacheIsMailFolder.TryAdd(fPath, isMail)
            If Not isMail Then _dbg("過濾非郵件資料夾", $"{fName} ({itemType})") ' 只有非郵件時才記錄
            Return isMail
        Catch
            Return False
        End Try
    End Function
    Private Function GetCleanSubject(subject As String) As String
        ' by Gemini 3 Flash, 2026/04/20: 移除常見的主旨前綴，讓分組更精準
        ' 支援包含 Re:, FW:, 回覆:, 轉寄: 等多國語言前綴的重複巢狀清理
        If String.IsNullOrEmpty(subject) Then Return ""

        'Dim clean = subject
        'Dim prefixes As String() = {"RE:", "FW:", "回覆:", "回信：", "轉寄:", "答复:", "转发:", "AW:", "VS:"} ' 加入德文/法文常見前綴
        'Dim found As Boolean = True
        'While found
        '    found = False
        '    For Each p In prefixes
        '        If clean.StartsWith(p, StringComparison.OrdinalIgnoreCase) Then
        '            clean = clean.Substring(p.Length).Trim()
        '            found = True
        '            Exit For
        '        End If
        '    Next
        'End While
        'Return clean

        ' 2026/06/12 by Simon/Claude: 改寫為 Compiled Regex 單次掃描，取代 while+for 迴圈
        '   - 新增語系：日(返信/転送)、韓(답장/회신/전달)、德(AW/WG)、法(Rép/TR)、北歐(VS)
        '   - 半形(:)全形(：)冒號通吃；前後空白(\s*)不限；大小寫不分(IgnoreCase)
        '   - 多層巢狀("Re: Re: FW: ...") 一次 Replace 完成，只產生一次 string allocation
        Return _subjectPrefixRe.Replace(subject, "")

    End Function
    Private Function NormalizeMailBody(body As String) As String
        ' ---------------------------------------------------------------
        ' NormalizeMailBody — 正規化郵件 Body，去除雜訊讓相似度比對更精確
        ' 2026/04/28 by Claude
        ' 處理步驟（依序）:
        '   1. Null/空值防護
        '   2. 去除 HTML 標籤（<...> 包圍的內容）
        '   3. 去除常見 HTML entities（&nbsp; &lt; 等）
        '   4. 去除所有空白字元（空格、Tab、換行、全形空白）
        '   5. 轉小寫（大小寫不影響內容相似度）
        ' 效果：兩封內容相同但格式不同的郵件（HTML vs 純文字）相似度會大幅提升
        ' ---------------------------------------------------------------
        If String.IsNullOrEmpty(body) Then Return ""

        ' 去除 HTML 標籤
        Dim result As String = System.Text.RegularExpressions.Regex.Replace(body, "<[^>]+>", "")

        ' 去除常見 HTML entities
        'result = result.Replace("&nbsp;", "").Replace("&lt;", "").Replace("&gt;", "").
        '                Replace("&amp;", "").Replace("&quot;", "").Replace("&#39;", "")
        result = System.Text.RegularExpressions.Regex.Replace(result, "&(nbsp|lt|gt|amp|quot|#39);", "")    ' 2026/6/6 by Gemini: 改用 Regex 優化多重 Replace效能

        ' 去除所有空白字元（含全形空白 \u3000、Tab、換行）
        result = System.Text.RegularExpressions.Regex.Replace(result, "[\s\u3000]+", "")
        Return result.ToLower()

    End Function
    Private Function objFolder2odoFolder(objFolder As Folder) As Redemption.RDOFolder
        If _rdo Is Nothing OrElse objFolder Is Nothing Then Return Nothing
        Return _rdo.GetFolderFromID(objFolder.EntryID, objFolder.StoreID)
    End Function
    Private Function rdoFolder2objFolder(rdoFolder As Redemption.RDOFolder) As Folder
        If rdoFolder Is Nothing Then Return Nothing
        Return _olNS.GetFolderFromID(rdoFolder.EntryID, rdoFolder.StoreID)
    End Function
#End Region
#End Region

End Class
