Imports System.Collections.Concurrent
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports Microsoft.Office.Interop.Outlook
' 2026/3/22 正式導入Redemption, 測試logon成功, 傳回數值成功
Imports Redemption
Imports Outlook = Microsoft.Office.Interop.Outlook
'Imports MailKit        ' MailKit is a cross-platform mail client library built on top of MimeKit.

' === 從頭重新設計 Layer2.5 / Layer3 底層計數函數 ===
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
    ' 2026-03-22 新增: 用於測試 Redemption.dll 整合 (注意: session.MAPIOBJECT 必須在 Outlook MAPI 連線建立後才能設定 (Form1_Load 尾端)
    '------------------------------------------------------------------------------------------------
    ' Outlook 物件(OOM)	    Redemption 物件 (RDO)     說明
    '------------------------------------------------------------------------------------------------
    ' Outlook.Application	Redemption 本體	        Redemption 是底層 MAPI 封裝，它不負責 UI 或視窗管理。
    ' Outlook.NameSpace	    Redemption.RDOSession	最接近。 負責管理登入、StoreID、PST 檔案庫與全域設定。
    ' Outlook.Folder	    Redemption.RDOFolder	對應資料夾層級。
    ' Outlook.MailItem	    Redemption.RDOMail	    對應單封郵件層級。
    ' Outlook.Store	        Redemption.RDOStore	    對應 PST 或 Exchange 帳戶。
    Private _rdo As Redemption.RDOSession = Nothing     ' _rdoSession 就等同是outlook.namespace 的意思, 就是Redemption的MAPI session
    Private _rdo2 As Redemption.RDOSession = Nothing    ' 2026/6/22 新增 by simon, 測試 Redemption 獨立 Session 資料隔離與效能差異
    ' 2026/06/23 by Simon/Claude: _rdo2 store-scoped resolve 用的對照快取(生命週期綁 _rdo2,於 CheckRDO 取消 / FormClosing 由 ReleaseRdoSession 釋放)
    Private _rdo2StoreByName As Dictionary(Of String, Redemption.RDOStore) = Nothing   ' store 顯示名 → RDOStore(權威,擁有 COM ref)
    Private _rdo2StoreByPath As Dictionary(Of String, Redemption.RDOStore) = Nothing   ' FolderPath → RDOStore(記憶化,免熱路徑重跑解析;值為 byName 參考,不另釋放)
    ' 2026/06/13 by Simon/Claude Opus 4.8: GetMailCountAllL3 / GetFolderCountAllL3 的 RDO 快速路徑(⓪TotalItemCount / ①平行枚舉)開關。
    '   問題: Redemption 走 MAPI 會枚舉到 OOM 看不到的隱藏/非-IPM 夾(Recoverable Items、Conversation Action Settings…)，
    '         導致子樹計數比 OOM 多算(實測資料夾數 27 vs 22)，且 ⓪TotalItemCount 是單一彙總值無法做 is_mail 過濾。
    '   現策略: 暫關此快速路徑，這兩個 All 計數一律走 ② OOM 完整骨架 + 計數層模式過濾，結構性保證與 OOM 一致。
    '           (每夾仍由 GetMailCountL3 內部各自享有 RDO 加速，故僅損失「單次彙總呼叫」這層、成本由夾數而非郵件數決定。)
    '   若日後要恢復 RDO 快速路徑: 須先在 GetSubtreeToListL3_Rdo 比照 OOM 可見性/IsMailFolder 過濾隱藏夾，再把此開關設 True。
    ' Private Shared ReadOnly _rdoFastPath As Boolean = False

    Private Shared _cacheIsMailFolder As New ConcurrentDictionary(Of String, Boolean)   ' 資料夾是否為郵件類型
    Private Shared _cacheMailCount As New ConcurrentDictionary(Of String, Long)         ' 自身資料夾的郵件個數
    Private Shared _cacheMailCountAll As New ConcurrentDictionary(Of String, Long)      ' 整支子樹的所有郵件總數
    Private Shared _cacheFolderCount As New ConcurrentDictionary(Of String, Long)       ' 自身資料夾的子目錄個數
    Private Shared _cacheFolderCountAll As New ConcurrentDictionary(Of String, Long)    ' 整支子樹的所有子目錄總數
    Private Shared _cacheFolderSize As New ConcurrentDictionary(Of String, Long)        ' 自身資料夾的郵件大小加總
    Private Shared _cacheFolderSizeAll As New ConcurrentDictionary(Of String, Long)     ' 整支子樹的所有子目錄郵件大小加總

    Private Shared _cacheFolderTree As New ConcurrentDictionary(Of String, List(Of Folder))     ' GetSortedSubFolders() 已排序的子資料夾清單
    Private Shared _cacheAttachMailList As New ConcurrentDictionary(Of String, FolderCacheTab3) ' 包含附件的郵件預掃描結果 (速度很快, 不用存入SSD?)
    Private Shared _cacheAttachFilename As New ConcurrentDictionary(Of String, List(Of String)) ' 所有附件檔名清單
    Private Shared _cacheMailBody As New ConcurrentDictionary(Of String, String)                ' by Gemini 3 Flash, 2026/04/26: Tab4 相似度計算用的 Body 快取 (session 級，避免重複讀取 Outlook mailitem.Body)

    Private Shared _cacheYearCounts As New ConcurrentDictionary(Of String, ConcurrentDictionary(Of Integer, Integer))
    Private Shared _cacheMonthCounts As New ConcurrentDictionary(Of String, ConcurrentDictionary(Of Integer, Integer))
    Private Shared _cacheSubTreeList As New ConcurrentDictionary(Of String, List(Of (folder As Outlook.Folder, fPath As String)))                           ' GetSubtreeToList() 的樹狀展開平坦化清單 (by Gemini, 2026/04/10: 帶路徑優化)
    Private Shared _cacheFolderIDs As New ConcurrentDictionary(Of String, (eid As String, sid As String, isMail As Boolean, hasCh As Boolean))              ' by Gemini, 2026/04/10: 專門儲存資料夾的身分標識與屬性標籤，用以橋接 Folder 物件與 SQLite 持久化
    Private Shared _cacheBasicMailInfo As New ConcurrentDictionary(Of String, (Mails As List(Of (Mail As MailItemInfo, Topic As String)), Snap As Long))    ' by Gemini, 2026/04/20: 專用於 Tab4 的郵件預掃描快取，Key 是資料夾路徑，Value 是該資料夾下所有郵件的基本資訊列表 (不帶 COM 物件) 與當下的 PR_CONTENT_COUNT 快照，用於快速顯示搜尋結果與驗證快取有效性

    Private Enum RefreshResult
        ' 2026/06/14 by Simon/Claude Opus 4.8: 單封刷新結果，供呼叫端決定失效郵件政策 (NotFound vs 暫時錯誤)
        Updated         ' 成功重讀並寫回 info
        NotFound        ' 依 EntryID 找不到 (信已被移動/刪除)
        TransientError  ' 暫時性 COM 錯誤，保留舊資料不動
    End Enum
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
    Private Structure FolderSortInfo
        ' by Gemini, 2026/03/29: 用於 GetSortedSubFolders 排序優化，減少 COM 屬性讀取次數 (O(N) vs O(N log N))
        Dim FolderObj As Folder
        Dim Name As String
        Dim HasChinese As Boolean
    End Structure
    Private Structure FolderCacheTab3
        Dim AttachMailList As List(Of MailItemInfo) ' 所有 hasAttachment 候選 (無大小篩選)
        Dim ItemCountSnap As Long                   ' 快取當下的 PR_CONTENT_COUNT，失效偵測用
    End Structure

    ' 2026/06/12 by Simon/Claude Opus 4.8: Compiled Regex，程式啟動時編譯一次，後續呼叫零額外開銷
    ' Pattern 說明：^ 錨定開頭；[：:] 同吃半形/全形冒號；外層 + 一次處理所有巢狀前綴
    Private Shared ReadOnly _subjectPrefixRe As New Regex(
        "^(?:(?:RE|FW|FWD|AW|WG|VS|Rép|TR|回覆|回信|轉寄|轉發|回复|答复|转发|返信|転送|답장|회신|전달)\s*[：:]\s*)+",
        RegexOptions.IgnoreCase Or RegexOptions.Compiled)

    ' 2026/06/18 by Simon/Claude Opus 4.8: NormalizeMailBody 三個 Regex 改預編譯(Shared ReadOnly + Compiled)。
    '   build pass 每封 3 次 Replace，預編譯省掉靜態快取查找並以 IL 加速。※ 真瓶頸在 COM 讀 mailBody，此為微優化。
    ' 2026/06/18 by Simon/Claude Opus 4.8: 去轉寄引用前綴。(?m)行首定位；[ \t\u3000] 只吃水平空白(空格/Tab/全形空格)，
    '   (?:>[ \t\u3000]*)+ 容許 > 之間夾雜空白(如 "> > >"、">> >")。刻意不吃 \r\n，避免把多行併成一行。
    Private Shared ReadOnly _reHtmlTag As New Regex("<[^>]+>", RegexOptions.Compiled)
    Private Shared ReadOnly _reHtmlEntity As New Regex("&(nbsp|lt|gt|amp|quot|#39);", RegexOptions.Compiled)
    Private Shared ReadOnly _reQuoteMarker As New Regex("(?m)^[ \t\u3000]*(?:>[ \t\u3000]*)+", RegexOptions.Compiled)
    Private Shared ReadOnly _reWhitespace As New Regex("[\s\u3000]+", RegexOptions.Compiled)
#End Region

#Region "■ 10 底層 COM 函數群 (新設計，現役主力) "
#Region "  ├ 全域初始化 & 載入釋放函數"
    Private Sub InitMapiNamespace()
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

            ' 2026/3/27 總算全部寫好RDO的導入, 但過程中優化了很多東西之後發現, 好像對效能沒有幫助到太多, 反而是演算法的改進才快更多
            '   RDO 的部份好像反而增加了程式碼複雜度跟拖慢啟動速度而已, 所以先關閉不使用
            ' 2026/6/23 更正: _rdo 是 piggyback 在 Outlook 原有的 MAPI session 上，速度根本上不來
            '   後來改用 _rdo2 獨立 session，才開發有數十倍的效能差異 (尤其是大量開郵件、讀附件檔名、讀取內文mailbody時)
            Dim unused = InitRdoSessionWithoutEULA()
        Catch ex As System.Exception
            _dbg("Redemption init FAIL", ex.Message)
            TryMarshalRelease(_rdo)
            _rdo = Nothing
        End Try

    End Sub
    Private Async Function InitRdoSessionWithoutEULA() As Task
        _dbg(" ├ 開始") ' by Gemini, 2026/04/10: 調整縮排層級為 Level 1
        ' 2026-03-23 v3:
        '   Task.Run 包裝保留 (讓 UI 執行緒繼續跑 LoadStoreToTreeView，平行初始化)
        '   第一次執行競爭條件改用 Thread.Sleep(1) 在 Set() 前解決，
        '   確保 AutoDismiss 輪詢 loop 已執行第一次再放行 New RDOSession()
        Try
            If _rdo IsNot Nothing Then Return
            Dim threadStarted As New System.Threading.ManualResetEventSlim(False)
            AutoDismissRdoEULA(threadStarted)
            ' 等 AutoDismiss thread 確認輪詢已開始，最多等 500ms
            threadStarted.Wait(500)
            _dbg(" ├ 進度", "AutoDismiss thread 已就緒，開始 New RDOSession") ' by Gemini, 2026/04/10

            ' ✅ Task.Run: UI 執行緒 不阻塞，LoadStoreToTreeView 可以同時跑
            Dim session As Redemption.RDOSession = Nothing
            Await Task.Run(Sub() session = New Redemption.RDOSession())

            ' MAPIOBJECT 必須回 UI 執行緒賦值 (_olNS 是 STA COM 物件)
            session.MAPIOBJECT = _olNS.MAPIOBJECT
            _rdo = session
            _dbg(" ├ _rdo init OK", $"Version={_rdo.Version}") ' by Gemini, 2026/04/10

            ' 2026/6/23 by Simon/Claude: 測試 Redemption 獨立 Session 資料隔離與效能差異
            Dim session2 As Redemption.RDOSession = Nothing
            session2 = New Redemption.RDOSession()
            session2.Logon(ProfileName:=_olNS.CurrentProfileName, Password:="", ShowDialog:=False, NewSession:=True)    ' 獨立session, 不沿用 Outlook MAPI session
            _rdo2 = session2
            _dbg(" ├ _rdo2 init OK", $"Version={_rdo2.Version}")

        Catch ex As System.Exception
            _rdo = Nothing
            _dbg("Redemption init FAIL", ex.Message)
        End Try

    End Function
    Private Sub AutoDismissRdoEULA(threadStarted As System.Threading.ManualResetEventSlim)
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
        '   threadStarted.Set() 通知呼叫端「輪詢已開始」，呼叫端等到 Set 後才呼叫 New RDOSession()，
        '   解決 thread pool 競爭導致首次執行抓不到視窗的問題
        Dim t As New System.Threading.Thread(
            Sub()
                ' ✅ 先讓輪詢 loop 跑第一次，再通知呼叫端
                '   避免 Set() 後呼叫端立刻 New RDOSession() 但此 thread 還沒執行到 FindWindow 的競爭條件
                System.Threading.Thread.Sleep(1)
                threadStarted.Set()
                Dim hWnd As IntPtr = IntPtr.Zero
                Dim timeout As Integer = 0

                ' 輪詢找 TEULAForm，最多等 30 秒 (3000 × 10ms)
                Do While hWnd = IntPtr.Zero AndAlso timeout < 3000
                    hWnd = FindWindow("TEULAForm", Nothing)
                    If hWnd = IntPtr.Zero Then System.Threading.Thread.Sleep(5) : timeout += 1
                Loop
                If hWnd = IntPtr.Zero Then _dbg("    ├ 錯誤", "逾時: 找不到 TEULAForm") : Return ' by Gemini, 2026/04/10: 調整縮排層級為 Level 2 (輔助執行緒)

                ' ✅ 立刻隱藏，使用者不會看到 EULA dialog 閃出來
                ShowWindow(hWnd, SW_HIDE)
                _dbg("    ├ 成功", $"TEULAForm 隱藏 hWnd=0x{hWnd:X}") ' by Gemini, 2026/04/10

                ' ── Step 1: "I agree" TRadioButton ──────────────────────
                ' 輪詢等子控制項建立完成 (視窗已隱藏，等待時間使用者無感)
                Dim hAgree As IntPtr = IntPtr.Zero
                Dim childTimeout As Integer = 0
                Do While hAgree = IntPtr.Zero AndAlso childTimeout < 500
                    hAgree = FindWindowEx(hWnd, IntPtr.Zero, "TRadioButton", "I agree")
                    If hAgree = IntPtr.Zero Then System.Threading.Thread.Sleep(5) : childTimeout += 1
                Loop

                If hAgree <> IntPtr.Zero Then
                    PostMessage(hAgree, WM_LBUTTONDOWN, New IntPtr(1), IntPtr.Zero)
                    PostMessage(hAgree, WM_LBUTTONUP, New IntPtr(1), IntPtr.Zero)
                    _dbg("    ├ 成功", "'I agree' PostMessage 送出") ' by Gemini, 2026/04/10
                Else
                    _dbg("    ├ 錯誤", "找不到 'I agree' (已逾時) ") ' by Gemini, 2026/04/10
                End If

                ' ── Step 2: "Ok" TButton ────────────────────────────────
                Dim hOk As IntPtr = IntPtr.Zero
                Dim okTimeout As Integer = 0
                Do While hOk = IntPtr.Zero AndAlso okTimeout < 500
                    hOk = FindWindowEx(hWnd, IntPtr.Zero, "TButton", "Ok")
                    If hOk = IntPtr.Zero Then System.Threading.Thread.Sleep(5) : okTimeout += 1
                Loop

                If hOk <> IntPtr.Zero Then
                    PostMessage(hOk, WM_LBUTTONDOWN, New IntPtr(1), IntPtr.Zero)
                    PostMessage(hOk, WM_LBUTTONUP, New IntPtr(1), IntPtr.Zero)
                    _dbg("    ├ 成功", "'Ok' PostMessage 送出") ' by Gemini, 2026/04/10
                Else
                    _dbg("    ├ 錯誤", "找不到 'Ok' (已逾時) ") ' by Gemini, 2026/04/10
                End If
            End Sub)

        t.Priority = System.Threading.ThreadPriority.AboveNormal
        t.IsBackground = True
        t.Start()

    End Sub
    Private Sub ReleaseRdoSession()
        _dbg(" ├ 開始")
        ' 釋放 _rdo2 獨立 session 與 store 對照快取(CheckRDO 取消 / FormClosing 共用)。
        '   獨立 session 須 Logoff 再 release,否則 Outlook 關不乾淨。byPath 值是 byName 的參考,不另釋放。
        If _rdo2StoreByName IsNot Nothing Then
            For Each s In _rdo2StoreByName.Values : Dim o As Object = s : TryMarshalRelease(o) : Next
            _rdo2StoreByName.Clear() : _rdo2StoreByName = Nothing
            _dbg(" ├ _rdo2StoreByName 釋放完成")
        End If

        If _rdo2StoreByPath IsNot Nothing Then _rdo2StoreByPath.Clear() : _rdo2StoreByPath = Nothing
        _dbg(" ├ _rdo2StoreByPath 釋放完成")

        If _rdo2 IsNot Nothing Then
            Try : _rdo2.Logoff() : Catch ex As System.Exception : _dbg("_rdo2.Logoff 異常", ex.Message) : End Try
            _dbg(" ├ _rdo2 Logoff 完成")
            Dim r As Object = _rdo2 : TryMarshalRelease(r) : _rdo2 = Nothing
            _dbg(" ├ _rdo2 釋放完成")
        End If

        ' 2026/6/23 by Simon/AntiGravity:
        ' _rdo 是 piggyback 在 Outlook session 上，不需要 Logoff，但要確保 COM ref 釋放且欄位歸 Nothing
        If _rdo IsNot Nothing Then
            Dim r As Object = _rdo : TryMarshalRelease(r) : _rdo = Nothing
            _dbg(" ├ _rdo 釋放完成")
        End If
        _dbg(" ├ 結束")
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
        _dbg(" ├ 開始", space.CurrentProfileName) ' by Gemini, 2026/04/10: 調整縮排層級為 Level 2 (由 InitMapiNamespace 呼叫)

        ' 遍歷所有Outlook.Store並添加到列表中, 使用LINQ擴充方法就夠快了, 不再使用非同步或Parallel.Foreach了
        Dim stores As List(Of Outlook.Store) = space.Stores.Cast(Of Outlook.Store)().ToList()
        stores = stores.OrderBy(Function(st) If(TextHasChineseChar(st.DisplayName), 1, 0)).ThenBy(Function(st) st.DisplayName).ToList() ' 使用 LINQ 排序Outlook.Store

        _dbg(" ├ 結束", $"Profile={space.CurrentProfileName} | 庫數量: {stores.Count}") ' by Gemini, 2026/04/10
        Return stores

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
        ' ✅ 2026/5/31 by Gemini/Simon: 加入 skipCache 引數判斷是否要強制讀取COM
        ' ===========================================================

        _dbg(" ├ 開始", sender.Name)

        Dim selectedNode As TreeNode = e.Node                   ' 取得點選的node
        Dim selectedFolder As Folder = selectedNode.Tag         ' 取得點選的資料夾
        Dim sortedFolders = GetSortedSubFolders(selectedFolder, skipCache:=_isForceRefreshing)   ' 取得所有子資料夾並排序

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
            ' 2026/06/14 by Simon/Claude Opus 4.8: 去模式化後 GetSubtreeToList 回傳「完整骨架(含非郵件夾)」，模式過濾移到計數層。
            '   Tab2-5 多選掃描共用此函數，必須在此補上 FilterSubtreeByMode，否則過濾模式下會把非郵件夾(行事曆/連絡人…)
            '   也納入掃描範圍 (狀態列出現 333 而非 ~307)。showAll 模式 FilterSubtreeByMode 回傳全集，行為不變。
            Dim subTree = Await GetSubtreeToList(rootF, includeSub, progress:=progress, cToken:=cToken)
            For Each subF In FilterSubtreeByMode(subTree, SafeGetPath(rootF))
                If addedPaths.Add(subF.fPath) Then fList.Add(subF)  ' ✅ 直接讀取 subF.fPath (Tuple 屬性)，再也不用打 COM!
            Next
            Await Task.Yield()
        Next
        Return fList
    End Function
    Private Function IsSubtreeComplete(rows As List(Of FolderStatsDbRow), rootPath As String) As Boolean
        ' ---------------------------------------------------------------
        ' 2026/06/13 by Simon/Claude Opus 4.8: 子樹骨架完整性檢查 (無模式分支 — 因骨架本就完整)
        ' 原理: 限定 rootPath 子樹範圍(避開 LIKE 前綴誤匹配 sibling，例如 Inbox 誤匹配 Inbox2)，
        '       對每個資料夾 F 要求「集合內 F 的直屬子夾數 == fc(F) 未過濾」。
        '       fc < 0(未知) 或對不上 → 判殘缺 → 由呼叫端 fallback L3 完整重掃。
        ' 注意: 依賴 DbGetSubFolderIDList 已 SELECT folder_count 並填入 row.fc (2026/06/13 配套修改)。
        ' ---------------------------------------------------------------
        If rows Is Nothing OrElse rows.Count = 0 Then Return False

        Dim prefix As String = rootPath & "\"        ' 只保留 rootPath 子樹範圍內的列 (path == root 或 startsWith root & "\")
        Dim inScope As New List(Of FolderStatsDbRow)(rows.Count)
        For Each r In rows
            If r.path = rootPath OrElse r.path.StartsWith(prefix, StringComparison.Ordinal) Then inScope.Add(r)
        Next
        If inScope.Count = 0 Then Return False

        Dim childCnt As New Dictionary(Of String, Integer)(inScope.Count)
        For Each r In inScope : childCnt(r.path) = 0 : Next
        For Each r In inScope
            Dim idx As Integer = r.path.LastIndexOf("\"c)
            If idx > 0 Then
                Dim parent As String = r.path.Substring(0, idx)
                If childCnt.ContainsKey(parent) Then childCnt(parent) += 1   ' 只累計集合內的子夾
            End If
        Next
        For Each r In inScope
            If r.fc < 0 Then Return False                       ' fc 未知 → 殘缺
            If childCnt(r.path) <> CInt(r.fc) Then Return False ' 集合內直屬子夾數對不上 fc → 殘缺
        Next
        Return True
    End Function
    Private Function FilterSubtreeByMode(skeleton As List(Of (folder As Folder, fPath As String)), rootPath As String) As List(Of (folder As Folder, fPath As String))
        ' ---------------------------------------------------------------
        ' 2026/06/13 by Simon/Claude Opus 4.8: 計數/顯示層的模式過濾 (剪枝移到這裡，骨架層永遠完整、0 COM)
        ' 依 _showAllFolders 從完整骨架即時派生:
        '   全顯(True) : 全數回傳。
        '   關閉(False): 從 root 沿 is_mail 的夾往下剪枝走訪 (碰非郵件夾不往下數)。root 一律納入(比照原 BFS 行為)。
        ' is_mail 來源: _cacheFolderIDs (L3 BFS 與 L2.5 DB 重建兩條路徑皆已回填)；查無時保守視為 is_mail(納入)，避免少算。
        ' ---------------------------------------------------------------
        If _showAllFolders Then Return skeleton

        Dim byPath As New Dictionary(Of String, (folder As Folder, fPath As String))(skeleton.Count)
        For Each t In skeleton : byPath(t.fPath) = t : Next

        ' 建 parent -> 直屬子路徑清單 (僅限集合內)
        Dim childrenOf As New Dictionary(Of String, List(Of String))(skeleton.Count)
        For Each t In skeleton
            Dim idx As Integer = t.fPath.LastIndexOf("\"c)
            If idx > 0 Then
                Dim parent As String = t.fPath.Substring(0, idx)
                If byPath.ContainsKey(parent) Then
                    Dim lst As List(Of String) = Nothing
                    If Not childrenOf.TryGetValue(parent, lst) Then lst = New List(Of String)() : childrenOf(parent) = lst
                    lst.Add(t.fPath)
                End If
            End If
        Next

        Dim result As New List(Of (folder As Folder, fPath As String))(skeleton.Count)
        Dim q As New Queue(Of String)()
        If byPath.ContainsKey(rootPath) Then result.Add(byPath(rootPath)) : q.Enqueue(rootPath)  ' root 一律納入
        While q.Count > 0
            Dim cur As String = q.Dequeue()
            Dim kids As List(Of String) = Nothing
            If childrenOf.TryGetValue(cur, kids) Then
                For Each kp In kids
                    Dim isMail As Boolean = True
                    Dim info As (eid As String, sid As String, isMail As Boolean, hasCh As Boolean) = Nothing
                    If _cacheFolderIDs.TryGetValue(kp, info) Then isMail = info.isMail
                    If isMail Then result.Add(byPath(kp)) : q.Enqueue(kp)   ' 郵件夾才納入並往下走；非郵件夾剪枝(不納入、不下探)
                Next
            End If
        End While
        Return result
    End Function
#End Region
#Region "  ├ Layer2.5 快取代理層 (Cache Proxy Layer)"
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
    Private Function GetSortedSubFolders(pFolder As Folder, Optional fPath As String = "", Optional skipCache As Boolean = False) As List(Of Folder)
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
        Dim suffix As String = If(_showAllFolders, "|True", "|False")
        Dim cacheKey As String = fPath & suffix         ' 2026/6/13 by Simon/Claude小幅優化: 這樣 "True/False" 是 literal string，CLR 會 intern，減少 GC 壓力
        Dim cachedFolders As List(Of Folder) = Nothing
        If Not skipCache AndAlso _cacheFolderTree.TryGetValue(cacheKey, cachedFolders) Then Return cachedFolders    ' 2026/6/27, Gemini 發現skipCache參數沒有穿透到這裡的記憶體快取

        ' ② SSD / DB 讀取分支 (Lazy Load): TreeView 展開時的主要加速點
        ' ✅ 2026/5/31 by Gemini/Simon: 加入 skipCache 引數判斷是否要強制讀取COM，避免在需要最新資料的情況下誤用過期快取
        If _dbCache IsNot Nothing AndAlso Not skipCache Then
            Dim dbIDs = DbGetOrderedSubFolderIDs(fPath, _showAllFolders)
            If dbIDs IsNot Nothing Then
                ' 預分配容量為 512，足以涵蓋多數資料夾搜尋結果，減少陣列頻繁 Resize 開銷 (by Gemini 3 Flash, 2026/05/04)
                Dim dbResults As New List(Of Folder)(512)
                For Each row In dbIDs
                    Try
                        ' DbGetSubFolderIDList 回傳的是 (eid, sid, path) 的具名 Tuple 列表 by Gemini 3.0 flash, 2026/04/16
                        Dim f = TryCast(_olNS.GetFolderFromID(row.eid, row.sid), Folder)
                        If f IsNot Nothing Then dbResults.Add(f)
                    Catch : End Try
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
                ' 2026/06/14 by Simon/Claude Opus 4.8: isMail 改為先算一次，供過濾與身分證註冊共用 (showAll 模式也需算出 is_mail 以正確登記)
                Dim isMail As Boolean = IsMailFolder(subF, childPath)
                If Not _showAllFolders AndAlso Not isMail Then Continue For

                infoList.Add(New FolderSortInfo With {.FolderObj = subF, .Name = sName, .HasChinese = TextHasChineseChar(sName)})
                ' 這裡 subF 被加入 infoList 成為物件清單，所以不能在這裡 TryRelease 它

                ' 2026/6/2: 再次修正F5 強制刷新的總數讀取不正確
                ' 🔽🔽🔽 【修復點 2】順手把展開的資料夾也註冊身分證 🔽🔽🔽
                ' 2026/06/14 by Simon/Claude Opus 4.8: 還原此修復點 (原被註解)。GetSortedSubFolders 是「樹載入」與「BuildBfsFolderTree 計算 BFS」
                '   共用的子夾枚舉樞紐，兩條路徑原本都不寫 _cacheFolderIDs → 子夾存檔時 entry_id/is_mail 為 NULL，
                '   重啟後被 DbGetOrderedSubFolderIDs (entry_id IS NOT NULL / is_mail=1) 濾掉 → 樹崩 (只剩收件匣)。
                '   在此 TryAdd 身分證即可一次修好兩條路徑；isMail 重用上方已算好的值，不重複打 COM。
                Try : _cacheFolderIDs.TryAdd(childPath, (subF.EntryID, subF.StoreID, isMail, TextHasChineseChar(sName)))
                Catch : End Try
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
    Private Function GetMailCount(folder As Folder, Optional fPath As String = "", Optional skipCache As Boolean = False) As Long
        ' ---------------------------------------------------------------
        ' GetMailCount — 單一資料夾本層郵件數 (PR_CONTENT_COUNT)
        ' 2026/4/7 by Claude, 加入Database lazy load, 讀資料順序:
        '   ① 記憶體命中
        '   ② DB lazy load：命中且 mc 有效且 snapshot 吻合 → 一次填滿所有欄位
        '   ③ 讀取派工: _rdo2 在(且 store 可解) → GetMailCountRdo;否則 → GetMailCountL3(OOM)
        ' 2026/04/15 by Gemini 3.1 Pro, 加入 optional fPath 參數，若有傳入則可省去 pFolder.FolderPath 1ms 耗時
        ' 2026/05/09 by Gemini 3.1 Pro: 重構成使用 SafeGetDbRow，回傳統一改成 Long 共用函式
        ' 2026/06/23 by Simon/Claude Opus 4.8: ③ 改為 RDO 派工(GetMailCountRdo via _rdo2,失敗 fallback OOM L3);
        '   加 skipCache 引數(繞過快取讀寫,給 F5 skipCache / snap 重讀等直呼者用,仍走 RDO 派工)。
        ' ---------------------------------------------------------------
        fPath = SafeGetPath(folder, fPath)
        Dim count As Long
        If Not skipCache Then
            If _cacheMailCount.TryGetValue(fPath, count) Then Return count       ' ① 記憶體命中
            Dim row = SafeGetDbRow(folder, fPath)                                ' ② DB lazy load
            If row IsNot Nothing AndAlso row.mc >= 0 Then Return row.mc
        End If

        ' ③ 讀取派工: RDO 優先,失敗 fallback OOM
        count = GetMailCountRdo(fPath, folder.EntryID, folder.StoreID)
        If count < 0 Then count = GetMailCountL3(folder, fPath:=fPath)

        If Not skipCache Then _cacheMailCount.TryAdd(fPath, count)
        Return count

    End Function ''
    Private Function GetFolderCount(folder As Folder, Optional fPath As String = "", Optional skipCache As Boolean = False) As Long
        ' ---------------------------------------------------------------
        ' GetFolderCount — 單一資料夾直屬子資料夾數
        ' 2026/4/7 by Claude, 加入Database lazy load, 讀資料順序:
        '   ① 記憶體命中
        '   ② DB lazy load：命中且 fc 有效且 snapshot 吻合 → 一次填滿所有欄位
        '   ③ 讀取派工: _rdo2 在(且 store 可解) → GetFolderCountRdo;否則 → GetFolderCountL3(OOM)
        ' 2026/05/09 by Gemini 3.1 Pro: 重構成使用 SafeGetDbRow，回傳統一改成 Long 共用函式
        ' 2026/06/23 by Simon/Claude Opus 4.8: ③ 改為 RDO 派工(GetFolderCountRdo via _rdo2,失敗 fallback OOM L3);
        '   加 skipCache 引數(繞過快取讀寫,給 F5 skipCache / snap 重讀等直呼者用,仍走 RDO 派工)。
        ' ---------------------------------------------------------------
        fPath = SafeGetPath(folder, fPath)
        Dim count As Long
        If Not skipCache Then
            If _cacheFolderCount.TryGetValue(fPath, count) Then Return count     ' ① 記憶體命中
            Dim row = SafeGetDbRow(folder, fPath)                                ' ② DB lazy load (fc 欄位)
            If row IsNot Nothing AndAlso row.fc >= 0 Then Return row.fc
        End If

        ' ③ 讀取派工: RDO 優先,失敗 fallback OOM
        count = GetFolderCountRdo(fPath, folder.EntryID, folder.StoreID)
        If count < 0 Then count = GetFolderCountL3(folder, fPath:=fPath)

        If Not skipCache Then _cacheFolderCount.TryAdd(fPath, count)
        Return count

    End Function ''
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
        ' todo: 優先要新增GetAttachMailListRdo(), 轉換至 _rdo2
        ' ---------------------------------------------------------------
        fPath = SafeGetPath(folder, fPath)
        Dim fName As String = ExtractFolderName(fPath)
        If _iLikeNoisy Then _dbg(" ├ 開始", fName)
        Dim key As String = fPath
        Dim currentCount As Long = GetMailCount(folder, fPath:=fPath)  ' 依賴同層快取 (本身已有 DB lazy load)

        ' ① 記憶體命中
        Dim entry As FolderCacheTab3 = Nothing ' 補上初始化以消除 BC42108 警告
        If _cacheAttachMailList.TryGetValue(key, entry) AndAlso entry.ItemCountSnap = currentCount Then Return entry.AttachMailList

        ' ② DB lazy load (attach_maillist)：pr_count_snap == currentCount → 快取仍有效
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
    Private Function GetAttachFilename(ByRef mail As MailItemInfo, Optional skipCache As Boolean = False) As List(Of String)
        ' ---------------------------------------------------------------
        ' GetAttachFilename — Tab3 Phase2：附件檔名清單 (by EntryID)
        ' by Gemini, 2026/04/04: Layer2.5 快取代理層 - 取得附件檔名清單 (含 _cacheAttachFilename 機制)
        ' 2026/4/7 by Claude, 加入Database lazy load, 讀資料順序:
        '   ① 記憶體命中
        '   ② DB lazy load：命中且 fc 有效且 snapshot 吻合 → 一次填滿所有欄位
        '   ③ fallback: Layer3 呼叫
        ' ---------------------------------------------------------------
        Dim result As List(Of String) = Nothing
        If Not skipCache AndAlso _cacheAttachFilename.TryGetValue(mail.EntryID, result) Then Return result  ' ①

        ' ② DB lazy load (attach_filenames)
        If Not skipCache Then
            result = DbGetAttachFilenames(mail.EntryID)
            If result IsNot Nothing Then
                _cacheAttachFilename.TryAdd(mail.EntryID, result)
                Return result
            End If
        End If

        ' ③ 讀取分派: _rdo2 在 → GetAttachFilenameRdo(store-scoped 高速,store 找不到時內部回 Nothing);否則 → GetAttachFilenameL3(OOM)。RDO 已上移至 L2.5。 2026/06/23 by Simon/Claude
        If _rdo2 IsNot Nothing Then
            result = GetAttachFilenameRdo(mail)
            If result Is Nothing Then result = GetAttachFilenameL3(mail)   ' RDO 解析失敗保底
        Else
            result = GetAttachFilenameL3(mail)
        End If
        If Not skipCache AndAlso result IsNot Nothing Then _cacheAttachFilename.TryAdd(mail.EntryID, result)
        Return result

    End Function ''
    Private Function GetMailBody(entryID As String, Optional folderPath As String = "", Optional skipCache As Boolean = False) As String
        ' ---------------------------------------------------------------
        ' GetMailBody — Layer2.5 快取代理：Body 快取存取點
        ' 2026/04/28 by Simon/Claude: 依照 L2.5 架構抽出快取邏輯，L3 只剩純 COM
        '   ① 快取命中（_cacheMailBody）→ 直接回傳，0 COM call
        '   ② 快取未命中 → 呼叫 L3 GetMailBodyL3 讀取並正規化
        '   ③ 無論成功或失敗都存快取（失敗存 ""），避免同一封信重複嘗試 COM
        ' ---------------------------------------------------------------
        ' 2026/06/23 by Simon/Claude Opus 4.8: 依照 L2.5 架構重構 GetMailBody，讀取分派邏輯:
        '   ① 快取命中（_cacheMailBody）→ 直接回傳，0 COM call
        '   ② (預留 SSD 層: body 目前不落 SQLite;待純文字總容量評估通過後,於此插入 lazy load,形狀對齊 GetAttachFilename)
        '   ③ 讀取分派: _rdo2 在且有 folderPath → GetMailBodyRdo(store-scoped 高速);否則 → GetMailBodyL3(OOM)。RDO 已上移至 L2.5。
        '      skipCache=True: 跳過 ①讀與寫快取(build pass 掃數萬封避免撐爆 _cacheMailBody),仍走 RDO/OOM 分派。
        ' ---------------------------------------------------------------
        Dim cached As String = Nothing
        If Not skipCache AndAlso _cacheMailBody.TryGetValue(entryID, cached) Then Return cached   ' ①

        Dim body As String
        If _rdo2 IsNot Nothing AndAlso folderPath <> "" Then          ' ③ RDO 優先
            body = GetMailBodyRdo(entryID, folderPath)
            If body Is Nothing Then body = GetMailBodyL3(entryID)     ' RDO 解析失敗保底
        Else
            body = GetMailBodyL3(entryID)                             ' OOM
        End If

        If Not skipCache Then _cacheMailBody(entryID) = body          ' 無論成功失敗都存(失敗存""),避免重複打 COM
        Return body

    End Function ''
    Private Async Function GetSubtreeToList(rootFolder As Folder, includeSubF As Boolean, Optional progress As IProgress(Of ProgressReport) = Nothing, Optional skipCache As Boolean = False, Optional cToken As CancellationToken = Nothing) As Task(Of List(Of (folder As Folder, fPath As String)))
        ' ---------------------------------------------------------------
        ' GetSubtreeToList — 整棵子資料夾清單 (Layer2.5 快取代理)
        ' 2026/04/17 by Claude: 從 GetSubtreeToList 拆出快取邏輯
        '   原來的快取邏輯混在 BFS 函數裡，現在統一到此 L2.5 層
        '   GetSubtreeToListL3 (原 GetSubtreeToList) 只剩純 BFS COM 掃描
        ' 呼叫順序: ① 記憶體命中 → ② DB lazy load → ③ Layer3 GetSubtreeToListL3
        ' includeSubF=False 時無需快取，直接呼叫 L3 回傳單節點清單
        ' 2026/06/13 by Simon/Claude Opus 4.8: 子樹計數鏈「去模式化」重構 —
        '   (1) 去模式化快取鍵: 只存一份完整骨架(含非郵件夾)，鍵不再含 _showAllFolders；模式過濾移至計數層(FilterSubtreeByMode)。
        '   (2) 完整性檢查: DB lazy 撈「未過濾全集」後，必須通過 IsSubtreeComplete(骨架完整) 才採用，否則 fallback L3 完整重掃，
        '       消滅原本「folder_stats 殘缺 → LIKE 默默回傳殘缺子樹 → 子樹靜默少算」的未爆彈。
        '   (3) skipCache=True: F5/discover 跳過記憶體+DB 快取，直打 L3 完整重掃 + 回填(補完整性檢查抓不到 Outlook 新增夾的 staleness 缺口)。
        '   (3) skipCache=True: F5/discover 跳過記憶體+DB 快取讀取，重算並覆寫快取(補完整性檢查抓不到 Outlook 新增夾的 staleness 缺口)。(2026/6/25)
        ' ---------------------------------------------------------------
        Dim rootPath As String = SafeGetPath(rootFolder)

        If Not includeSubF Then Return Await GetSubtreeToListL3(rootFolder, False, progress, cToken:=cToken) ' 單節點不快取

        ' 2026/06/13 by Simon/Claude Opus 4.8: 去模式化 — 鍵不再含 _showAllFolders (原: rootPath & "|" & _showAllFolders)
        Dim cacheKey As String = rootPath

        If Not skipCache Then
            Dim cachedList As List(Of (folder As Folder, fPath As String)) = Nothing
            If _cacheSubTreeList.TryGetValue(cacheKey, cachedList) Then              ' ① 記憶體命中 (完整骨架)
                _dbg(" ├ 結束", $"{ExtractFolderName(rootPath)} (Cache Hit) | 資料夾總計: {cachedList.Count}")
                Return cachedList
            End If

            ' ② DB lazy load: 利用 LIKE 一次取回整棵樹的 ID 並重建物件
            ' 注意: DB 存放的是 (EntryID, StoreID, FolderPath)，我們在這裡重建 Tuple
            ' 2026/06/13 by Simon/Claude Opus 4.8: 一律撈未過濾全集(isIncludeAll:=True)，再用 IsSubtreeComplete 驗證骨架完整性
            Dim dbRows = DbGetSubFolderIDList(rootPath, isIncludeAll:=True)            ' ② DB lazy load (完整全集)
            If dbRows IsNot Nothing AndAlso IsSubtreeComplete(dbRows, rootPath) Then
                ' 預分配容量為 512，優化從 DB 載入資料夾子樹時的處理速度 (by Gemini 3 Flash, 2026/05/04)
                Dim dbResults As New List(Of (folder As Folder, fPath As String))(512)
                For Each row In dbRows
                    Try
                        ' DbGetSubFolderIDList 回傳的是 (eid, sid, path, isMail, hasCh, fc) 的具名列表 by Gemini 3.0 flash, 2026/04/16
                        Dim f = TryCast(_olNS.GetFolderFromID(row.eid, row.sid), Folder)
                        If f IsNot Nothing Then
                            dbResults.Add((Folder:=f, fPath:=row.path))
                            ' 2026/06/13 by Simon/Claude Opus 4.8: 回填身分證(is_mail)與 fc，供計數層 FilterSubtreeByMode 的 is_mail 過濾使用
                            _cacheFolderIDs.TryAdd(row.path, (row.eid, row.sid, row.isMail <> 0, row.hasCh <> 0))
                            If row.fc >= 0 Then _cacheFolderCount(row.path) = row.fc
                        End If
                    Catch
                    End Try
                Next
                If dbResults.Count > 0 Then
                    _cacheSubTreeList(cacheKey) = dbResults
                    If _iLikeNoisy Then _dbg("    ├ SSD Hit (Tree)", $"{ExtractFolderName(rootPath)}: 已從資料庫載入完整骨架 {dbResults.Count} 個子目錄")
                    Return dbResults
                End If
            ElseIf dbRows IsNot Nothing Then
                If _iLikeNoisy Then _dbg("    ├ DB 殘缺", $"{ExtractFolderName(rootPath)}: folder_stats 子樹不完整 → fallback L3 完整重掃")
            End If
        End If

        ' 🆕 2026/06/25 by Simon/Claude: RDO 快速探索派工(上移自 L3)。閘 = _rdo2 在(= CheckRDO 勾)。
        '   skipCache=True 時仍走此處重算並覆寫快取;GetSubtreeListRdo 回 Nothing(RDO 不可用/失敗)才掉回 ③ 純 OOM L3。
        If _rdo2 IsNot Nothing Then
            Dim rdoResult = GetSubtreeListRdo(rootFolder, rootPath, progress)
            If rdoResult IsNot Nothing Then
                If Not cToken.IsCancellationRequested AndAlso rdoResult.Count > 0 Then _cacheSubTreeList(cacheKey) = rdoResult
                Return rdoResult
            End If
        End If

        ' ③ Layer3 BFS COM 完整掃描；OCE re-throw，快取寫入在 L3 完成後由 L3 自行負責 (純 OOM fallback)
        Return Await GetSubtreeToListL3(rootFolder, True, progress, cToken:=cToken)

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
    Private Async Function GetFolderBasicByEntryIDL3(fPath As String, ct As CancellationToken) As Task(Of Dictionary(Of String, MailItemInfo))
        ' 2026/06/14 by Simon/Claude Opus 4.8: 方法B底層 — 對單一資料夾一次 GetTable+GetArray，回傳 EntryID→基本欄位 dict
        '   不加 hasattachment 過濾 (要服務 Lv3/4/5 任意郵件)；解析不到資料夾 → 回 Nothing 讓呼叫端退回方法A
        '   路徑→Folder：經 _cacheFolderIDs 取 (eid,sid) 再 GetFolderFromID (與既有 L3 慣例一致)
        ' todo: 這個跟上面的 GetBasicMailInfoL3 是否有點重複應整合???
        Const PR_MESSAGE_SIZE As String = "http://schemas.microsoft.com/mapi/proptag/0x0E080003"
        Dim ids As (eid As String, sid As String, isMail As Boolean, hasCh As Boolean) = Nothing
        If Not _cacheFolderIDs.TryGetValue(fPath, ids) Then Return Nothing
        Dim folder As Folder = TryCast(_olNS.GetFolderFromID(ids.eid, ids.sid), Folder)
        If folder Is Nothing Then Return Nothing

        Dim result As New Dictionary(Of String, MailItemInfo)(StringComparer.Ordinal)
        Dim table As Outlook.Table = Nothing
        Try
            table = SafeGetTable(folder, "", "EntryID", "Subject", PR_MESSAGE_SIZE, "ReceivedTime", "SenderName")
            Dim swThrottle As Stopwatch = Stopwatch.StartNew()
            Do
                ct.ThrowIfCancellationRequested()
                Dim data = SafeGetArray(table)
                If data Is Nothing Then Exit Do
                For r As Integer = 0 To data.GetUpperBound(0)
                    Dim entryID As String = SafeGet(Of String)(data, r, 0, "")
                    If entryID = "" Then Continue For
                    result(entryID) = New MailItemInfo With {.EntryID = entryID,
                                                             .Subject = SafeGet(Of String)(data, r, 1, ""),
                                                             .Size = SafeGet(Of Long)(data, r, 2, 0L),
                                                             .RcvTime = SafeGet(Of DateTime)(data, r, 3, DateTime.MinValue),
                                                             .SenderName = SafeGet(Of String)(data, r, 4, "")}
                Next
                Await SmartThrottle(swThrottle, ct, ThrottleFreq.Hii, Sub() PgrsBar2.Text = $"批次掃描 {folder.Name}: {result.Count} 筆...")
            Loop
        Catch ex As OperationCanceledException
            Throw
        Catch ex As System.Exception
            If _iLikeNoisy Then _dbg("    ├ GetFolderBasicByEntryIDL3 錯誤", $"{fPath} — {ex.Message}")
            Return Nothing
        Finally
            TryMarshalRelease(table)
            TryMarshalRelease(folder)
        End Try
        Return result
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
            If _iLikeNoisy AndAlso Not isMail Then _dbg("過濾非郵件資料夾", $"{fName} ({itemType})") ' 只有非郵件時才記錄
            Return isMail
        Catch
            Return False
        End Try
    End Function
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
#End Region
#Region "  ├ Layer3 RDO 直接存取底層"
    Private Function GetRdoStore(folderPath As String) As Redemption.RDOStore
        ' 由 FolderPath 取得 _rdo2 上對應的 RDOStore(store-scoped resolve 用)
        '   首次呼叫一次掃 _rdo2.Stores 建 byName 表;之後 byPath 記憶化命中為 O(1) 零解析。找不到回 Nothing(呼叫端 fallback OOM)。
        '   ※ phase 1 單執行緒(UI 緒)存取,用一般 Dictionary;日後平行讀取(phase C)再評估執行緒安全。
        If _rdo2 Is Nothing Then Return Nothing

        If _rdo2StoreByName Is Nothing Then
            _rdo2StoreByName = New Dictionary(Of String, Redemption.RDOStore)()
            _rdo2StoreByPath = New Dictionary(Of String, Redemption.RDOStore)()
            Try
                For i As Integer = 1 To _rdo2.Stores.Count
                    Dim s As Redemption.RDOStore = _rdo2.Stores.Item(i)
                    Dim nm As String = s.Name
                    If Not String.IsNullOrEmpty(nm) AndAlso Not _rdo2StoreByName.ContainsKey(nm) Then _rdo2StoreByName(nm) = s Else Dim o As Object = s : TryMarshalRelease(o)
                Next
            Catch ex As System.Exception
                _dbg("GetRdo2Store 建表失敗", ex.Message)
            End Try
        End If

        Dim store As Redemption.RDOStore = Nothing
        If _rdo2StoreByPath.TryGetValue(folderPath, store) Then Return store    ' 記憶化命中(含 Nothing)
        _rdo2StoreByName.TryGetValue(GetStoreNameFromPath(folderPath), store)
        _rdo2StoreByPath(folderPath) = store                                    ' 含 Nothing 亦記,避免重複解析
        Return store

    End Function
    ' 🆕 2026/06/24 by Simon/Claude Opus 4.8: GetSubtreeListRdo — RDO 批次階層走訪(探針 C 實證版)
    '   兩階段拆乾淨、分開計時:
    '     Phase1 探索(純 RDO): Folders.MAPITable.GetRows 整層批次,只對 PR_SUBFOLDERS=true 遞迴(跳葉夾);
    '            同步寫 _cacheFolderCount(每夾 fc = 該層 RowCount,葉夾=0)。探針已證正確 + ~8ms。
    '     Phase2 還原(純 OOM): 逐夾 _olNS.GetFolderFromID(eid,sid) 還原 Outlook.Folder,呼叫既有 IsMailFolder,
    '            註冊 _cacheFolderIDs,組出與 BFS 同合約 (Folder,fPath) 清單。← 正確性對齊保留,Step2 再拆掉。
    '   任一步解析不到 → 回 Nothing,由 GetSubtreeToListL3 掉回 OOM BFS(絕不產錯結果)。
    '   ※ 唯一未驗假設: Phase2 用 _rdo2 table eid 餵 _olNS(跨 session)。production 跑成功即驗證,失敗則安全網接住。
    ' 🆕 2026/06/25 by Simon/Claude: GetSubtreeListRdo 派工殼 — Phase1 批次主支,失敗退 RDO 枚舉(3e);Phase2 共用 OOM 還原。
    '   rdoRoot 由本殼統一持有/釋放;兩支 helper 只釋放自己開的子夾,故批次失敗後枚舉可重用同一 rdoRoot。
    '   兩個 RDO 法都失敗 → 回 Nothing,由 L2.5 掉回純 OOM GetSubtreeToListL3。
    Private Function GetSubtreeListRdo(rootFolder As Folder, rootPath As String, Optional progress As IProgress(Of ProgressReport) = Nothing) As List(Of (folder As Folder, fPath As String))
        Dim store As Redemption.RDOStore = GetRdoStore(rootPath)
        If store Is Nothing Then _dbg("    ├ RDO 略過", "GetRdo2Store=Nothing → OOM BFS") : Return Nothing
        Dim sid As String = "" : Try : sid = rootFolder.StoreID : Catch : Return Nothing : End Try
        Dim rdoRoot As Redemption.RDOFolder = Nothing
        Try : rdoRoot = TryCast(store.GetFolderFromID(rootFolder.EntryID), Redemption.RDOFolder) : Catch : End Try
        If rdoRoot Is Nothing Then _dbg("    ├ RDO 略過", "root 解析失敗 → OOM BFS") : Return Nothing
        _dbg("    ├ RDO 啟動", $"子樹走訪: {ExtractFolderName(rootPath)}")

        Try
            ' ── Phase 1 探索: 批次主支,失敗退 RDO 枚舉 ──
            Dim swD As Stopwatch = Stopwatch.StartNew()
            Dim method As String = "批次"
            Dim nodes As List(Of (eid As String, name As String, path As String)) = GetSubtreeRdoByBatch(store, rdoRoot, rootPath)
            If nodes Is Nothing Then
                _dbg("    ├ RDO 退枚舉", "批次失敗 → 退簡單 RDO 枚舉")
                method = "枚舉"
                nodes = GetSubtreeRdoByEnum(rdoRoot, rootPath)
            End If
            If nodes Is Nothing Then
                Dim m As String = "✗ RDO 批次+枚舉皆失敗 → 改走 OOM BFS"
                _dbg("    ├ RDO 探索", m) : progress?.Report(New ProgressReport With {.Message = m})
                Return Nothing
            End If
            swD.Stop()

            ' ── Phase 2 還原(純 OOM,正確性對齊) ──
            Dim swM As Stopwatch = Stopwatch.StartNew()
            Dim result As New List(Of (folder As Folder, fPath As String))(nodes.Count + 1)
            result.Add((rootFolder, rootPath))
            For Each nd In nodes
                Dim f As Folder = Nothing
                Try : If nd.eid <> "" Then f = TryCast(_olNS.GetFolderFromID(nd.eid, sid), Folder)
                Catch : End Try
                If f Is Nothing Then
                    Dim m As String = $"✗ RDO 還原失敗於 {ExtractFolderName(nd.path)} → 改走 OOM BFS"
                    _dbg("    ├ RDO Phase2", m) : progress?.Report(New ProgressReport With {.Message = m})
                    Return Nothing
                End If
                Dim isMail As Boolean = IsMailFolder(f, nd.path)
                Try : _cacheFolderIDs.TryAdd(nd.path, (f.EntryID, f.StoreID, isMail, TextHasChineseChar(nd.name))) : Catch : End Try
                result.Add((f, nd.path))
            Next
            swM.Stop()

            Dim tMs As Long = swD.ElapsedMilliseconds + swM.ElapsedMilliseconds
            Dim spd As String = If(tMs > 0, $"{result.Count * 1000.0 / tMs:N0} 夾/秒", "極快(<1ms)")
            Dim doneMsg As String = $"✓ RDO 子樹完成({method}): {result.Count} 夾 | 探索 {swD.ElapsedMilliseconds}ms + 還原 {swM.ElapsedMilliseconds}ms = {tMs}ms | {spd}"
            _dbg("    ├ RDO 完成", doneMsg)
            progress?.Report(New ProgressReport With {.CurrentCount = result.Count, .TotalCount = result.Count, .Message = doneMsg})
            Return result
        Finally
            Dim oo As Object = rdoRoot : TryMarshalRelease(oo)   ' rdoRoot 由本殼統一釋放
        End Try
    End Function
    Private Function GetSubtreeRdoByBatch(store As Redemption.RDOStore, rdoRoot As Redemption.RDOFolder, rootPath As String) As List(Of (eid As String, name As String, path As String))
        ' 🆕 探索主支: Folders.MAPITable 批次,只對 PR_SUBFOLDERS=true 遞迴(跳葉夾)。回 nodes 或 Nothing(交給枚舉)。
        '   不釋放 rdoRoot(外層殼負責),只釋放自己 GetFolderFromID 開出的子夾。
        If _iLikeNoisy Then _dbg("", store.ToString & rdoRoot.ToString & rootPath.ToString)
        Dim nodes As New List(Of (eid As String, name As String, path As String))(512)
        Dim toRel As New List(Of Object)()
        Const COLS As String = "Name, EntryID, http://schemas.microsoft.com/mapi/proptag/0x360A000B"  ' PR_SUBFOLDERS
        Try
            Dim q As New Queue(Of (f As Redemption.RDOFolder, p As String))(512) : q.Enqueue((rdoRoot, rootPath))
            While q.Count > 0
                Dim cur = q.Dequeue()
                Dim foldersCol = cur.f.Folders
                Dim tbl = foldersCol.MAPITable
                Dim rc As Integer = CInt(tbl.RowCount)
                _cacheFolderCount(cur.p) = CLng(rc)
                If rc > 0 Then
                    tbl.Columns = COLS : tbl.GoToFirst()
                    Dim rowsArr As Array = DirectCast(tbl.GetRows(rc), Array)
                    For i As Integer = rowsArr.GetLowerBound(0) To rowsArr.GetUpperBound(0)
                        Dim row As Array = DirectCast(rowsArr.GetValue(i), Array)
                        Dim lb As Integer = row.GetLowerBound(0)
                        Dim nm As String = TryCast(row.GetValue(lb), String) : If nm Is Nothing Then nm = ""
                        Dim eidHex As String = RdoTableEidToHex(row.GetValue(lb + 1))
                        Dim vSub As Object = row.GetValue(lb + 2)
                        Dim hasSub As Boolean = If(TypeOf vSub Is Boolean, CBool(vSub), True)  ' 未知→保守遞迴
                        Dim cp As String = cur.p & "\" & nm
                        nodes.Add((eidHex, nm, cp))
                        If hasSub Then
                            Dim child As Redemption.RDOFolder = Nothing
                            If eidHex <> "" Then child = TryCast(store.GetFolderFromID(eidHex), Redemption.RDOFolder)
                            If child Is Nothing Then Return Nothing   ' 宣稱有子夾卻開不了 → 批次放棄,交給枚舉
                            q.Enqueue((child, cp)) : toRel.Add(child)
                        Else
                            _cacheFolderCount(cp) = 0L
                        End If
                    Next
                End If
                TryMarshalRelease(tbl) : TryMarshalRelease(foldersCol)
            End While
        Catch ex As System.Exception
            _dbg("    ├ RDO 批次失敗", ex.GetBaseException().Message)
            Return Nothing
        Finally
            For Each o In toRel : Dim oo As Object = o : TryMarshalRelease(oo) : Next
        End Try
        Return nodes
    End Function
    Private Function GetSubtreeRdoByEnum(rdoRoot As Redemption.RDOFolder, rootPath As String) As List(Of (eid As String, name As String, path As String))
        ' 🆕 探索 fallback: 簡單 RDO 枚舉(For Each Folders),逐夾讀 .Name/.EntryID。回 nodes 或 Nothing。
        '   不釋放 rdoRoot(外層殼負責),只釋放自己枚舉到的子夾。
        If _iLikeNoisy Then _dbg("", rdoRoot.ToString & rootPath.ToString)
        Dim nodes As New List(Of (eid As String, name As String, path As String))(512)
        Dim toRel As New List(Of Object)()
        Try
            Dim q As New Queue(Of (f As Redemption.RDOFolder, p As String))(512) : q.Enqueue((rdoRoot, rootPath))
            While q.Count > 0
                Dim cur = q.Dequeue()
                Dim subs = cur.f.Folders
                Dim childCount As Integer = 0
                For Each sf As Redemption.RDOFolder In subs
                    toRel.Add(sf)
                    Dim nm As String = "" : Try : nm = sf.Name : Catch : Continue For : End Try
                    Dim eid As String = "" : Try : eid = sf.EntryID : Catch : End Try
                    Dim cp As String = cur.p & "\" & nm
                    nodes.Add((eid, nm, cp)) : childCount += 1
                    q.Enqueue((sf, cp))
                Next
                _cacheFolderCount(cur.p) = CLng(childCount)
                TryMarshalRelease(subs)
            End While
        Catch ex As System.Exception
            _dbg("    ├ RDO 枚舉失敗", ex.GetBaseException().Message)
            Return Nothing
        Finally
            For Each o In toRel : Dim oo As Object = o : TryMarshalRelease(oo) : Next
        End Try
        Return nodes
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

    Private Function GetMailCountRdo(folderPath As String, eid As String, sid As String) As Long
        ' ---------------------------------------------------------------
        ' 本層郵件數 RDO 讀取層。store-scoped on _rdo2,與 GetMailCountL3 原 ⓪ tier 同邏輯(rdoFolder.Items.Count)。
        '   解析失敗(store/folder 找不到/例外)回 -1,由 L2.5 proxy 判 <0 fallback 到 GetMailCountL3。
        '   sid 目前未用(store-scoped 單參數即可解,探針 12/12),保留參數對稱備雙參數 fallback。
        ' 2026/06/23 by Simon/Claude Opus 4.8
        ' ---------------------------------------------------------------
        Dim store As Redemption.RDOStore = GetRdoStore(folderPath)
        If store Is Nothing Then Return -1

        Dim rdoFolder As Redemption.RDOFolder = Nothing
        Try
            rdoFolder = TryCast(store.GetFolderFromID(eid), Redemption.RDOFolder)
            If rdoFolder Is Nothing Then Return -1
            Return CLng(rdoFolder.Items.Count)
        Catch ex As System.Exception
            If _iLikeNoisy Then _dbg("GetMailCountRdo 失敗", $"{ExtractFolderName(folderPath)} | {ex.Message}")
            Return -1
        Finally
            Dim o As Object = rdoFolder : TryMarshalRelease(o)
        End Try
    End Function
    Private Function GetFolderCountRdo(folderPath As String, eid As String, sid As String) As Long
        ' ---------------------------------------------------------------
        ' 本層直屬子資料夾數 RDO 讀取層。store-scoped on _rdo2,與 GetFolderCountL3 原 ⓪ tier 同邏輯(rdoFolder.Folders.Count)。
        '   解析失敗回 -1,由 L2.5 proxy 判 <0 fallback 到 GetFolderCountL3。
        '   sid 目前未用(同上),保留參數對稱。
        ' 2026/06/23 by Simon/Claude Opus 4.8
        ' ---------------------------------------------------------------
        Dim store As Redemption.RDOStore = GetRdoStore(folderPath)
        If store Is Nothing Then Return -1

        Dim rdoFolder As Redemption.RDOFolder = Nothing
        Try
            rdoFolder = TryCast(store.GetFolderFromID(eid), Redemption.RDOFolder)
            If rdoFolder Is Nothing Then Return -1
            Return CLng(rdoFolder.Folders.Count)
        Catch ex As System.Exception
            _dbg("GetFolderCountRdo 失敗", $"{ExtractFolderName(folderPath)} | {ex.Message}")
            Return -1
        Finally
            Dim o As Object = rdoFolder : TryMarshalRelease(o)
        End Try
    End Function
    Private Function GetAttachFilenameRdo(ByRef mail As MailItemInfo) As List(Of String)
        ' ---------------------------------------------------------------
        ' 附件檔名 RDO 讀取層。store-scoped on _rdo2,與 GetAttachFilenameL3 原 ⓪ tier 同邏輯(att.Type=1)。
        '   解析失敗(store 找不到/例外)回 Nothing,由 L2.5 fallback 到 GetAttachFilenameL3。
        ' 2026/06/23: L3 原用 New List(4096) 預配置(每封配 ~32KB,熱路徑浪費)此處改以 mail.AttachCount 精準預配置(上界,絕不 realloc,亦不浪費)。
        ' ---------------------------------------------------------------
        Dim store As Redemption.RDOStore = GetRdoStore(mail.FolderPath)
        If store Is Nothing Then Return Nothing

        Dim result As New List(Of String)(mail.AttachCount)
        Dim rdoMsg As Redemption.RDOMail = Nothing
        Try
            rdoMsg = TryCast(store.GetMessageFromID(mail.EntryID), Redemption.RDOMail)
            If rdoMsg Is Nothing Then Return Nothing

            For i As Integer = 1 To rdoMsg.Attachments.Count
                Dim att As Redemption.RDOAttachment = rdoMsg.Attachments.Item(i)
                Try : If att.Type = 1 Then result.Add(att.FileName)   ' 僅 olByValue(1),與 GetAttachFilenameL3 一致
                Finally : Dim o As Object = att : TryMarshalRelease(o)
                End Try
            Next
            Return result
        Catch ex As System.Exception
            If _iLikeNoisy Then _dbg("GetAttachFilenameRdo 失敗", ex.Message)
            Return Nothing
        Finally
            Dim o As Object = rdoMsg : TryMarshalRelease(o)
        End Try
    End Function
    Private Function GetMailBodyRdo(entryID As String, folderPath As String) As String
        ' ---------------------------------------------------------------
        ' mailbody RDO 讀取層。store-scoped on _rdo2,讀 .Body 後套同一支 NormalizeMailBody(與 GetMailBodyL3 一致)。
        '   RDOMail.Body 不分 Mail/Post 型別;解析失敗回 Nothing,由 L2.5 fallback 到 GetMailBodyL3。
        ' ---------------------------------------------------------------
        Dim store As Redemption.RDOStore = GetRdoStore(folderPath)
        If store Is Nothing Then Return Nothing
        Dim rm As Redemption.RDOMail = Nothing
        Try
            rm = TryCast(store.GetMessageFromID(entryID), Redemption.RDOMail)
            If rm Is Nothing Then Return Nothing
            Return NormalizeMailBody(rm.Body)
        Catch ex As System.Exception
            _dbg("GetMailBodyRdo 失敗", $"{entryID}: {ex.Message}")
            Return Nothing
        Finally
            Dim o As Object = rm : TryMarshalRelease(o)
        End Try
    End Function
    Private Function RdoTableEidToHex(v As Object) As String
        ' table 的 PR_ENTRYID 經 GetRows 回 byte array(探針實證 Byte[]),統一轉 hex 字串供 GetFolderFromID
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
#Region "  ├ Layer3 OOM 直接存取底層"
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
        ' ⓪ RDO 路徑已上移至 L2.5 GetMailCountRdo(store-scoped on _rdo2),L3 回歸純 OOM。 2026/06/23 by Simon/Claude

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

    End Function ''
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

        ' ⓪ Redemption: RDOFolder.Folders.Count 與 OOM pFolder.Folders.Count 等價，但可在任意執行緒呼叫, 2026-03-22 新增
        ' ⓪ RDO 路徑已上移至 L2.5 GetFolderCountRdo(store-scoped on _rdo2),L3 回歸純 OOM。 2026/06/23 by Simon/Claude

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

    End Function ''
    Private Async Function GetMailCountAllL3(rootFolder As Folder, Optional progress As IProgress(Of ProgressReport) = Nothing, Optional skipCache As Boolean = False, Optional cToken As CancellationToken = Nothing) As Task(Of Long)
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
        ' 2026/06/13 by Simon/Claude Opus 4.8: 更正 — 上述「死碼」註解已過時。CollectFolderStatsByL3ForceRefresh (F5 強刷入口)
        '   已復用本函數 (Form1_MainTab12.vb)，故本函數現為「F5/skipCache 子樹計數」的活路徑，非死碼。
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
        '   一次 COM call 結束，不需要任何 BFS 遍歷或平行處理, 2026-03-22 新增
        ' 2026/3/24 by Gemini: ① 平行 BFS (RDO):
        '   使用 GetSubtreeToListL3_Rdo 取得清單，以 Parallel.ForEach 搭配 Interlocked.Add 快速加總
        '   Redemption (RDO) 是 free-threaded，在背景平行執行安全且極為高效
        ' 2026/04/15 by Claude: 改用 ParallelOptions.CancellationToken 取代 _cancelRequested 旗標
        ' 2026/06/13 by Simon/Claude Opus 4.8: RDO 快速路徑以 _rdoFastPath 開關控管 (預設關)。
        '   原因見 _rdoFastPath 宣告處: TotalItemCount 含 OOM 看不到的隱藏夾且無法 is_mail 過濾，會與 OOM 不一致。
        ' 2026/06/25 by Claude: 移除 ⓪TotalItemCount / ①平行BFS 兩條 _rdoFastPath 死分支(恆 False,從未執行)。
        '   停用原因(枚舉到隱藏夾、與 OOM 不一致)見 _rdoFastPath 宣告處(L45);該問題已由 GetSubtreeListRdo(IPM 樹根走訪)解決,
        '   RDO 加速統一在 GetSubtreeToList 層發生。此函數的 RDO 加速經由下方 ② 的 GetSubtreeToList 自動取得。

        ' ② BFS 循序累加: GetSubtreeToList 展開 + GetMailCountL3(Layer3) 逐一加總
        '   支援取消檢查和 progress 進度回報，比平行版保守但穩定
        ' 2026/04/15 by Claude: _cancelRequested 取代為 SmartThrottle(swThrottle, cToken)
        '   cToken 取消時 Task.Delay(1,cToken) 拋 OCE → Catch OCE → Return -1
        '   同時移除舊的 i Mod 10 Await Task.Yield()，統一由 SmartThrottle 每 100ms 讓出一次
        Try
            ' 2026/04/17 by Claude: 改呼叫 GetSubtreeToList (L2.5)，享有快取加速
            ' 2026/06/13 by Simon/Claude Opus 4.8: 取得完整骨架後依 _showAllFolders 在計數層過濾 (剪枝移到這裡)；skipCache 一路 thread
            ' 2026/06/25 by Claude Opus 4.8: forceRefresh引數改 skipCache:=skipCache
            Dim skeleton As List(Of (folder As Folder, fPath As String)) = Await GetSubtreeToList(rootFolder, includeSubF:=True, skipCache:=skipCache, cToken:=cToken)
            Dim targetFolderList As List(Of (folder As Folder, fPath As String)) = FilterSubtreeByMode(skeleton, SafeGetPath(rootFolder))
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
        ' 2026/06/13 by Simon/Claude Opus 4.8: ③ 為 ② 拋非取消例外時的極罕見保險，未套用 _showAllFolders 模式過濾 (會含非郵件夾)；
        '   若實務上發現 ③ 被觸發致關閉模式下數字偏大，再回頭補剪枝。
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
    Private Async Function GetFolderCountAllL3(rootFolder As Folder, Optional progress As IProgress(Of ProgressReport) = Nothing, Optional skipCache As Boolean = False, Optional cToken As CancellationToken = Nothing) As Task(Of Long)
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
        ' 2026/06/13 by Simon/Claude Opus 4.8: 更正 — 同 GetMailCountAllL3，本函數已被 CollectFolderStatsByL3ForceRefresh (F5 強刷) 復用，非死碼。
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
        ' 2026/06/13 by Simon/Claude Opus 4.8: 以 _rdoFastPath 開關控管 (預設關)。RDO 枚舉會多算 OOM 看不到的隱藏/非-IPM 夾
        '   (實測 27 vs OOM 22)，故暫關，改走 ② OOM 完整骨架 + 模式過濾以保證與 OOM 一致。
        ' 2026/06/25 by Claude: 移除 ⓪平行BFS _rdoFastPath 死分支(恆 False,從未執行)。原因同 GetMailCountAllL3,已由 GetSubtreeListRdo 解決。

        ' 2026/3/24 by Gemini: ② OOM + BFS 循序 (無 Redemption 時的最後手段)
        '   必須循序處理 OOM COM 物件以避免 STA 違規
        ' 2026/04/15 by Claude: 傳入 cToken，GetSubtreeToList 本身支援取消，OCE 向上冒泡
        Try
            ' 2026/04/16 by Gemini: GetSubtreeToList 現在回傳 Tuple，解開它以維持後續邏輯
            ' 2026/04/17 by Claude: 改呼叫 GetSubtreeToList (L2.5)，享有快取加速
            ' 2026/06/13 by Simon/Claude Opus 4.8: 取得完整骨架後依 _showAllFolders 在計數層過濾 (剪枝移到這裡)；skipCache 一路 thread
            ' 2026/06/25 by Claude Opus 4.8: forceRefresh引數改 skipCache:=skipCache
            Dim skeleton = Await GetSubtreeToList(rootFolder, includeSubF:=True, progress:=progress, skipCache:=skipCache, cToken:=cToken)
            Dim targetTupleList = FilterSubtreeByMode(skeleton, SafeGetPath(rootFolder))
            Dim allFolders = targetTupleList.Select(Function(x) x.folder).ToList()
            ' by Gemini, 2026/04/02: BFS 展開後回傳數量 (扣除 rootFolder 自身)
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
        ' todo: 2026/6/23: 目前RDO這裡看起來都是無效的段落, 根本沒有讀到值
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
                ' 2026/04/15 by Claude: 改用 SmartThrottle 整合讓出與取消偵測，進度回報移入節流區塊內，由 SmartThrottle 統一控制頻率
                ' 2026/04/16 by Gemini 3.0 flash: 改用 SmartThrottle 整合進度回報與讓出點
                Await SmartThrottle(swThrottle, cToken:=cToken, ThrottleFreq.Mid,
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
        ' todo: 2026/6/23: 目前RDO這裡看起來都是無效的段落, 根本沒有讀到值
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
        ' 2026/3/11 再次重構: 優化 COM 呼叫，減少 RCW 物件積累，提升效能和穩定性
        ' 2026/3/24 by Gemini: 從逐年 Restrict 改為 GetTable + GetArray 一次讀完再記憶體分組
        '   原本每年一次 Restrict + Items.Count = ~30 次 COM call
        '   現在 1 次 GetTable + ceil(N/1000) 次 GetArray，大幅減少 COM 跨程序呼叫
        ' todo: 目前最耗時間的function(), 占整體時間60~65%
        ' 2026/04/05 by Gemini: 每 100ms 節流讓出執行緒
        ' 2026/04/15 by Claude: 加入 cToken 參數
        '   取代 _cancelRequested 旗標，改用 SmartThrottle(swThrottle, cToken) 節流讓出
        '   cToken 取消時 Task.Delay 拋 OCE，此函數不攔截 (讓 OCE 冒泡至 CollectYearCounts)
        '   原因: 攔住後回傳半截 yearCounts，L2 會誤以為該資料夾已統計完畢，導致計數偏低
        ' ---------------------------------------------------------------
        fPath = SafeGetPath(folder, fPath)
        Dim fName As String = ExtractFolderName(fPath)
        If _iLikeNoisy Then _dbg("    ├ 開始", fName)

        Dim yearCounts As New ConcurrentDictionary(Of Integer, Integer)
        Dim table As Outlook.Table = Nothing
        Try
            ' 2026/3/24 by Gemini: 改用 GetTable + GetArray 取代逐年 Restrict
            table = SafeGetTable(folder, "", "ReceivedTime") ' 只讀 RcvTime 一欄，最小化每 row 的傳輸量

            ' by Gemini, 2026/04/05: 每批次讀取後，若超過 100ms 則釋放執行緒並檢查中斷
            Dim swThrottle As Stopwatch = Stopwatch.StartNew()  ' by Claude Sonnet 4.6, 2026/06/07
            Do
                Dim data = SafeGetArray(table) : If data Is Nothing Then Exit Do
                For r As Integer = 0 To data.GetUpperBound(0)
                    Dim receivedTime As DateTime = SafeGet(Of DateTime)(data, r, 0, DateTime.MinValue)
                    If receivedTime > DateTime.MinValue Then
                        Dim year As Integer = receivedTime.Year
                        If year > 0 AndAlso year <= Date.Today.Year Then yearCounts.AddOrUpdate(year, 1, Function(k, v) v + 1)
                    End If
                Next
                Await SmartThrottle(swThrottle, cToken:=cToken)
                ' 2026/04/15 by Claude: _cancelRequested 取代為 SmartThrottle(swThrottle, cToken) 整合讓出與取消偵測，OCE 冒泡至呼叫端 CollectYearCounts
            Loop
        Catch ex As OperationCanceledException
            If _iLikeNoisy Then _dbg("    ├ 已取消", fName) : Throw            ' 2026/04/15: 不攔截 OCE，直接 re-throw 讓 CollectYearCounts 感知取消
        Catch ex As System.Exception
            If _iLikeNoisy Then _dbg("    ├ 錯誤", $"{fName}: {ex.Message}")  ' by Gemini, 2026/04/04: Issue 4 格式標準化
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
            ' 2026/3/24 by Gemini: 改用 GetTable + 日期範圍 DASL filter + GetArray，用整年的日期範圍一次篩選，不再逐月 Restrict
            Dim startDate As New Date(year, 1, 1, 0, 0, 0)
            Dim endDate As New Date(year, 12, 31, 23, 59, 59)
            Dim dateFilter As String = $"[ReceivedTime] >= '{startDate}' AND [ReceivedTime] <= '{endDate}'"
            table = SafeGetTable(folder, dateFilter, "ReceivedTime")

            Dim swThrottle As Stopwatch = Stopwatch.StartNew()  ' by Claude Sonnet 4.6, 2026/06/07
            Do
                Dim data = SafeGetArray(table) : If data Is Nothing Then Exit Do
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
            If _iLikeNoisy Then _dbg("    ├ 已取消", $"{fName} ({year} 年)") : Throw      ' 2026/04/15: re-throw 讓 GetMonthCountsForYear (L2.5) 感知，不寫入快取
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
        Dim result As New List(Of String)(mail.AttachCount)

        ' ⓪ Redemption 優先: 繞過 OOM 開信的記憶體開銷，直接透過 MAPI Table 抓取檔名
        ' ⓪ RDO 路徑已上移至 L2.5 GetAttachFilenameRdo(store-scoped on _rdo2),L3 回歸純 OOM。 2026/06/23 by Simon/Claude

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
    End Function ''
    Private Function GetMailBodyL3(entryID As String) As String
        ' ---------------------------------------------------------------
        ' GetMailBodyL3 — Layer3 COM 資料層：讀取郵件 Body 並正規化
        ' 2026/04/28 by Simon/Claude: 以 Simon 的 GetMailBodyByEntryID 為基礎
        '   + 加入 NormalizeMailBody 正規化（去除 HTML 標籤、空白換行）
        '   + Await Task.Yield() 確保每封讀完後讓 UI 執行緒喘氣
        '   支援 MailItem 與 PostItem 兩種型別
        '   使用獨立的 ns（Simon 的設計），確保 COM namespace 不跨執行緒共用
        ' ---------------------------------------------------------------
        If String.IsNullOrEmpty(entryID) Then Return ""

        ' Dim ns As Outlook.NameSpace = Nothing
        Dim item As Object = Nothing
        Dim mailBody As String = ""
        Try
            item = _olNS.GetItemFromID(entryID) ' 2026/04/28 by simon, 使用共用的 NameSpace 以減少多建一次namespace的 COM 開銷
            ' ns = _olApp.GetNamespace("MAPI")  ' 2026/04/26 by Gemini, 使用自己內部的 NameSpace 以更好封裝, 並自行TryMarshalRelease以減少GCW洩漏
            ' item = ns.GetItemFromID(entryID)

            If item IsNot Nothing Then
                If TypeOf item Is Outlook.MailItem Then
                    mailBody = NormalizeMailBody(DirectCast(item, Outlook.MailItem).Body)
                ElseIf TypeOf item Is Outlook.PostItem Then
                    mailBody = NormalizeMailBody(DirectCast(item, Outlook.PostItem).Body)
                End If
            End If
        Catch ex As System.Exception
            _dbg("GetMailBodyL3 失敗", $"{entryID}: {ex.Message}")
        Finally
            TryMarshalRelease(item)
            ' TryMarshalRelease(ns)   ' 2026/04/26 by Gemini, 使用自己內部的 NameSpace 以更好封裝, 並自行TryMarshalRelease以減少GCW洩漏
            ' 2026/05/09 by Gemini 3.0 flash: 移除內部的 Yield。改由調用端依批次執行呼吸，減少微切換開銷提升讀取性能
        End Try
        Return mailBody

    End Function ''
    Private Async Function GetBasicMailInfoL3(folder As Folder, needTopic As Boolean, cToken As CancellationToken, Optional fPath As String = "") As Task(Of List(Of (Mail As MailItemInfo, Topic As String)))
        ' ---------------------------------------------------------------
        ' 2026/05/06 by Claude: 永遠讀取全部 8 欄（含 topic/msgId/senderEmail）
        '   needTopic 參數保留供 API 相容，但 L3 層已不區分，統一讀取
        '   欄位索引: 0=EntryID, 1=Subject, 2=Size, 3=RcvTime,
        '             4=SenderName, 5=Topic, 6=MsgIDhash, 7=SenderEmail
        '
        ' 2026/06/12 by Simon/Claude Opus 4.8: 移除 PR_CONVERSATION_TOPIC (欄位 5) topic 改由 GetCleanSubject(subject) 動態計算，與 DB 讀取路徑保持一致
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
                    ' 2026/06/12 by Simon/Claude Opus 4.8: topic 從 GetCleanSubject(subject) 動態計算
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
        ' 2026/06/13 by Simon/Claude Opus 4.8: 去模式化 — 鍵不再含 _showAllFolders (原: rootPath & "|" & _showAllFolders)，只存一份完整骨架
        Dim cacheKey As String = rootPath

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

        ' 🆕 2026/06/24 by Simon/Claude Opus 4.8: RDO 快速探索 tier(探針 C 實證 ~60-150× 提速,Step1 正確性對齊版)
        '   旗標開且 _rdo2 在 → GetSubtreeListRdo 批次走訪;成功寫 _cacheSubTreeList 回傳,失敗(Nothing)往下掉回 OOM BFS。
        '   合約不變(回 (Folder,fPath)),_cacheFolderIDs/_cacheFolderCount 由 GetSubtreeListRdo 內比照 BFS 註冊。OOM BFS 當 fallback,一行不動。

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
                            ' 2026/06/13 by Simon/Claude Opus 4.8: 移除模式剪枝 — L3 永遠完整列舉整棵骨架(含非郵件夾)，模式(_showAllFolders)
                            '   過濾移到計數/顯示層(FilterSubtreeByMode)。徹底消滅「folder_stats 殘缺 → 子樹靜默少算」未爆彈。
                            '   (原剪枝: If Not _showAllFolders AndAlso Not isMail Then Continue For — 已移除)

                            ' ✅ 加強 EntryID/StoreID 讀取的安全性
                            Try
                                _cacheFolderIDs.TryAdd(childPath, (subF.EntryID, subF.StoreID, isMail, TextHasChineseChar(fName)))
                            Catch : End Try

                            result.Add((subF, childPath))   ' ✅ 同步存入預計好的路徑，不再打 COM
                            queue.Enqueue((subF, childPath))
                        Next
                        ' 2026/06/13 by Simon/Claude Opus 4.8: 回填 current 的未過濾直屬子夾數(fc)，供 IsSubtreeComplete 完整性檢查用。
                        '   必須在 Finally 釋放 subFolders 前讀取 .Count。
                        _cacheFolderCount(current.Path) = CLng(subFolders.Count)
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
        ' 2026/06/24 by Simon/Claude: 改無條件 + 速度 + 報 PgrsBar2,方便與 RDO tier A/B 對比(純報告,無邏輯變動)
        Dim oomMs As Long = sw.ElapsedMilliseconds
        Dim oomSpd As String = If(oomMs > 0, $"{result.Count * 1000.0 / oomMs:N0} 夾/秒", "極快(<1ms)")
        Dim oomMsg As String = $"✓ OOM BFS 子樹完成: {rootFolder.Name}, {result.Count} 夾 | {oomMs}ms | {oomSpd}"
        _dbg("    ├ 結束", oomMsg)
        progress?.Report(New ProgressReport With {.CurrentCount = result.Count, .TotalCount = result.Count, .Message = oomMsg})
        Return result

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
    Private Function RefreshMailInfoL3(ByRef info As MailItemInfo, readAttachCount As Boolean) As RefreshResult
        ' 2026/06/14 by Simon/Claude Opus 4.8: 依 EntryID 從 COM 重讀單封郵件實體資訊，寫回 info (ByRef，MailItemInfo 為 Structure)
        '   跳過所有 cache (_cacheXXX / SSD)，直接打 COM。RDO→OOM fallback。
        '   readAttachCount=False: 只讀 Subject/Size/RcvTime/SenderName (全體 F5 用，省去開信枚舉附件)
        '   readAttachCount=True : 額外枚舉 olByValue(=1) 附件數 (右鍵單筆/複數刷新用)
        '   回傳 Updated/NotFound/TransientError；失效郵件的移除政策由呼叫端決定 (目前一律保留+記錄)
        Const MAPI_E_NOT_FOUND As Integer = &H8004010F   ' 找不到物件的 HRESULT，用以區分 NotFound vs 暫時性錯誤
        If String.IsNullOrEmpty(info.EntryID) Then Return RefreshResult.NotFound

        ' ⓪ Redemption 優先: 繞過 OOM 開信的記憶體開銷 (目前 _rdo 停用，實務上會直接落到 OOM)
        ' todo: 只剩這裡優先要轉換 _rdo2
        If _rdo IsNot Nothing Then
            Dim rdoMsg As Redemption.RDOMail = Nothing
            Try
                rdoMsg = TryCast(_rdo.GetMessageFromID(info.EntryID), Redemption.RDOMail)
                If rdoMsg IsNot Nothing Then
                    info.Subject = rdoMsg.Subject
                    info.Size = rdoMsg.Size
                    info.RcvTime = rdoMsg.ReceivedTime
                    info.SenderName = rdoMsg.SenderName
                    If readAttachCount Then
                        Dim n As Integer = 0
                        For i As Integer = 1 To rdoMsg.Attachments.Count
                            Dim att As Redemption.RDOAttachment = rdoMsg.Attachments.Item(i)
                            Try : If att.Type = 1 Then n += 1   ' 僅算 olByValue(1)，與 GetAttachFilenameL3 一致
                            Finally : TryMarshalRelease(att)
                            End Try
                        Next
                        info.AttachCount = n
                    End If
                    Return RefreshResult.Updated
                End If
            Catch ex As System.Exception   ' RDO 任何失敗都讓 OOM 再試一次，由 OOM 作最終結論
                If _iLikeNoisy Then _dbg("    ├ ⓪ RDO 失敗，走 OOM fallback", ex.Message)
            Finally
                TryMarshalRelease(rdoMsg)
            End Try
        End If

        ' ① Fallback: Outlook OOM
        Dim mail As Outlook.MailItem = Nothing
        Try
            mail = TryCast(_olNS.GetItemFromID(info.EntryID), Outlook.MailItem)
            If mail Is Nothing Then Return RefreshResult.NotFound

            info.Subject = mail.Subject
            info.Size = mail.Size
            info.RcvTime = mail.ReceivedTime
            info.SenderName = mail.SenderName
            info.FolderPath = SafeGetPath(TryCast(mail.Parent, Folder)) ' 2026/6/15 by simon: 新增直接從 mail.Parent 更新路徑，確保即使資料夾被移動後仍能正確更新
            ' todo: mail.Parent 會建立一個 COM object，傳入函數後沒有變數持有，TryMarshalRelease 就無從釋放 → COM memory leak

            If readAttachCount Then
                Dim attachments As Outlook.Attachments = Nothing
                Try
                    attachments = mail.Attachments
                    Dim attCount As Integer = attachments.Count    ' 存成變數避免 COM 重複呼叫
                    Dim n As Integer = 0
                    For i As Integer = 1 To attCount
                        Dim att As Outlook.Attachment = attachments.Item(i)
                        Try : If att.Type = Outlook.OlAttachmentType.olByValue Then n += 1   ' 僅算 olByValue(1)
                        Finally : TryMarshalRelease(att)
                        End Try
                    Next
                    info.AttachCount = n
                Finally
                    TryMarshalRelease(attachments)
                End Try
            End If
            Return RefreshResult.Updated

        Catch ex As System.Runtime.InteropServices.COMException
            ' 區分 NotFound (ID 失效) 與暫時性錯誤，供呼叫端日後移除政策使用
            If ex.ErrorCode = MAPI_E_NOT_FOUND Then
                If _iLikeNoisy Then _dbg("    ├ ① NotFound", info.EntryID)
                Return RefreshResult.NotFound
            End If
            If _iLikeNoisy Then _dbg("    ├ ① COM 暫時性錯誤", ex.Message)
            Return RefreshResult.TransientError
        Catch ex As System.Exception
            If _iLikeNoisy Then _dbg("    ├ ① 例外", ex.Message)
            Return RefreshResult.TransientError
        Finally
            TryMarshalRelease(mail)
        End Try
    End Function
#End Region
#Region "  ├ Legacy 保留（暫無呼叫端）"

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
    Private Shared Function GetStoreNameFromPath(folderPath As String) As String
        ' 2026/06/19 by Simon/Claude: 從 \\StoreName\sub\... 取出最前面的 store 顯示名 (店名含逗號/空格/&/~/@ 皆安全, 因唯一分隔符是反斜線)
        Dim p = folderPath.TrimStart("\"c)
        Dim idx = p.IndexOf("\"c)
        Return If(idx > 0, p.Substring(0, idx), p)
    End Function
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
        ' snapshot 驗證: DB 儲存的 pr_count_snap = save 時的 PR_CONTENT_COUNT 值
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

        ' 2026/06/12 by Simon/Claude Opus 4.8: 改寫為 Compiled Regex 單次掃描，取代 while+for 迴圈
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
        '   4. 去除轉寄引用前綴（行首的 > 符號，含 > 間夾雜的空白）— 保留被引用文字本體
        '   5. 去除所有空白字元（空格、Tab、換行、全形空白）
        '   6. 轉小寫（大小寫不影響內容相似度）
        ' 效果：兩封內容相同但格式不同的郵件（HTML vs 純文字 / 引用版 vs 原文版）相似度會大幅提升
        ' 2026/06/18 by Simon/Claude Opus 4.8: 改用預編譯 Regex；新增引用前綴移除(步驟4)；ToLower→ToLowerInvariant
        ' ---------------------------------------------------------------
        If String.IsNullOrEmpty(body) Then Return ""
        'Dim result As String = result.Replace("&nbsp;", "").Replace("&lt;", "").Replace("&gt;", "").Replace("&amp;", "").Replace("&quot;", "").Replace("&#39;", "")

        Dim result As String
        ' 2026/6/22: .Body 是純文字，HTML 去標籤幾乎是空轉（且可能誤刪）GetMailBodyL3 讀的是.Body（OOM 回傳純文字）， 不是.HTMLBody。所以純文字信幾乎沒有 <...> 標籤可去。
        result = _reHtmlTag.Replace(body, "")       ' 去除 HTML 標籤
        result = _reHtmlEntity.Replace(result, "")  ' 去除常見 HTML entities ' 2026/6/6 by Gemini: 改用 Regex 優化多重 Replace效能
        result = _reQuoteMarker.Replace(result, "") ' 2026/06/18 by Simon/Claude Opus 4.8: 去除行首轉寄引用前綴(> 之間可夾雜空白)，保留引用文字本體。必須在去空白前做(靠行首定位)
        result = _reWhitespace.Replace(result, "")  ' 去除所有空白字元(含 "\u3000"=全形空白、Tab、換行)
        Return result.ToLowerInvariant()
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
