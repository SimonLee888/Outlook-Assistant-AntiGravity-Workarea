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
'       上層流程 (如 CollectFolderInfoByBFS) 負責決定何時呼叫、如何使用結果、快取管理等
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

    ' MAPI Proptag 常數 (集中宣告，供 Module_Outlook.vb / Module_Win32API.vb 共用)
    Private Const PR_SUBJECT As String = "http://schemas.microsoft.com/mapi/proptag/0x0037001F"
    Private Const PR_CONVERSATION_TOPIC As String = "http://schemas.microsoft.com/mapi/proptag/0x0070001E"
    Private Const PR_SENDER_NAME As String = "http://schemas.microsoft.com/mapi/proptag/0x0C1A001F"
    Private Const PR_SENDER_EMAIL_ADDRESS_W As String = "http://schemas.microsoft.com/mapi/proptag/0x0C1F001F"   ' PidTagSenderEmailAddress (Unicode)
    Private Const PR_SENDER_EMAIL_ADDRESS_A As String = "http://schemas.microsoft.com/mapi/proptag/0x0C1F001E"   ' PidTagSenderEmailAddress (ANSI)
    Private Const PR_MESSAGE_DELIVERY_TIME As String = "http://schemas.microsoft.com/mapi/proptag/0x0E060040"
    Private Const PR_HASATTACH As String = "http://schemas.microsoft.com/mapi/proptag/0x0E1B000B"
    Private Const PR_MESSAGE_SIZE As String = "http://schemas.microsoft.com/mapi/proptag/0x0E080003"             ' PT_LONG, 32-bit
    Private Const PR_MESSAGE_SIZE_EXTENDED As String = "http://schemas.microsoft.com/mapi/proptag/0x0E080014"    ' PT_I8, 64-bit
    Private Const PR_INTERNET_MESSAGE_ID_W As String = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"    ' PidTagInternetMessageId (Unicode)
    Private Const PR_INTERNET_MESSAGE_ID_A As String = "http://schemas.microsoft.com/mapi/proptag/0x1035001E"    ' PidTagInternetMessageId (ANSI)
    Private Const PR_SUBFOLDERS As String = "http://schemas.microsoft.com/mapi/proptag/0x360A000B"
    Private Const PR_CONTENT_COUNT As String = "http://schemas.microsoft.com/mapi/proptag/0x36020003"
    Private Const PR_CONTAINER_CLASS As String = "http://schemas.microsoft.com/mapi/proptag/0x3613001E"
    Private Const PR_ATTR_HIDDEN As String = "http://schemas.microsoft.com/mapi/proptag/0x10F4000B"              ' PT_BOOLEAN, 2026/07/08: 交談動作設定/快速步驟設定/提醒 等隱藏系統夾判定用
    Private Const PR_LOCAL_COMMIT_TIME_MAX As String = "http://schemas.microsoft.com/mapi/proptag/0x670A0040"    ' PT_SYSTIME
    Private Const DASL_SMARTNOATTACH As String = "http://schemas.microsoft.com/mapi/id/{00062008-0000-0000-C000-000000000046}/8514000B"   ' PidLidSmartNoAttach (PSETID_Common/0x8514/PT_BOOLEAN)

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
    Private _rdo2 As Redemption.RDOSession = Nothing    ' 2026/6/22 新增 by simon, 測試 Redemption 獨立 Session 資料隔離與效能差異
    ' 2026/06/23 by Simon/Claude: _rdo2 store-scoped resolve 用的對照快取(生命週期綁 _rdo2,於 CheckRDO 取消 / FormClosing 由 ReleaseRdoSession 釋放)
    Private _rdo2StoreByName As Dictionary(Of String, Redemption.RDOStore) = Nothing   ' store 顯示名 → RDOStore(權威,擁有 COM ref)
    Private _rdo2StoreByPath As Dictionary(Of String, Redemption.RDOStore) = Nothing   ' FolderPath → RDOStore(記憶化,免熱路徑重跑解析;值為 byName 參考,不另釋放)
    ' Private Shared ReadOnly _rdoFastPath As Boolean = False   ' 2026/06/13 by Simon/Claude Opus 4.8: RDO 快速路徑(⓪TotalItemCount / ①平行枚舉)開關。
    ' 問題: Redemption 走 MAPI 會枚舉到 OOM 看不到的隱藏/非-IPM 夾(Recoverable Items、Conversation Action Settings…)，

    Private Shared _cacheMailCount As New ConcurrentDictionary(Of String, Long)         ' 自身資料夾的郵件個數
    Private Shared _cacheMailCountAll As New ConcurrentDictionary(Of String, Long)      ' 整支子樹的所有郵件總數
    Private Shared _cacheFolderCount As New ConcurrentDictionary(Of String, Long)       ' 自身資料夾的子目錄個數
    Private Shared _cacheFolderCountAll As New ConcurrentDictionary(Of String, Long)    ' 整支子樹的所有子目錄總數
    Private Shared _cacheFolderSize As New ConcurrentDictionary(Of String, Long)        ' 自身資料夾的郵件大小加總
    Private Shared _cacheFolderSizeAll As New ConcurrentDictionary(Of String, Long)     ' 整支子樹的所有子目錄郵件大小加總

    ' 2026/07/10 by Simon/Claude Fable 5: 值型別由 List(Of Folder) 升級為 (f, name, fPath) tuple —
    '   GetSortedSubFolders 枚舉時本來就讀過 Name/路徑，舊版回傳時丟棄，害 LoadSubFolderToTreeView 建節點又重打 COM 讀一次。
    Private Shared _cacheFolderTree As New ConcurrentDictionary(Of String, List(Of (f As Folder, name As String, fPath As String)))     ' GetSortedSubFolders() 已排序的子資料夾清單
    ' 2026/07/10 by Simon/Claude Fable 5: store-root 清單借放 _cacheFolderTree 的保留鍵 (真實資料夾路徑一律以 "\\" 開頭，不會撞鍵)，
    '   不另開字典 — ClearMemoryCachesCore / F5 清 _cacheFolderTree 時自動一併失效重讀
    Private Const KEY_STORE_ROOTS As String = ":::StoreRoots"
    Private Shared _cacheAttMailList As New ConcurrentDictionary(Of String, FolderCacheTab3)    ' 包含附件的郵件預掃描結果 (速度很快, 不用存入SSD?)
    Private Shared _cacheAttFilename As New ConcurrentDictionary(Of String, List(Of String))    ' 所有附件檔名清單
    Private Shared _cacheMailBody As New ConcurrentDictionary(Of String, String)                ' by Gemini 3 Flash, 2026/04/26: Tab4 相似度計算用的 Body 快取 (session 級，避免重複讀取 Outlook mailitem.Body)

    Private Shared _cacheYearCount As New ConcurrentDictionary(Of String, ConcurrentDictionary(Of Integer, Integer))
    Private Shared _cacheMonthCount As New ConcurrentDictionary(Of String, ConcurrentDictionary(Of Integer, Integer))

    ' GetSubtree() 的樹狀展開平坦化清單 (by Gemini, 2026/04/10: 帶路徑優化) (2026/06/28 Stage2: 改帶 eid/sid 不帶 COM 物件)
    Private Shared _cacheSubTreeList As New ConcurrentDictionary(Of String, List(Of (eid As String, sid As String, fPath As String)))
    Private Shared _cacheFolderIDs As New ConcurrentDictionary(Of String, (eid As String, sid As String, isMail As Boolean, hasCh As Boolean))      ' by Gemini, 2026/04/10: 專門儲存資料夾的身分標識與屬性標籤，用以橋接 Folder 物件與 SQLite 持久化
    Private Shared _cacheMailInfo As New ConcurrentDictionary(Of String, (Mails As List(Of (Mail As MailItemInfo, Topic As String)), Snap As Long)) ' by Gemini, 2026/04/20: 專用於 Tab4 的郵件預掃描快取，Key 是資料夾路徑，Value 是該資料夾下所有郵件的基本資訊列表 (不帶 COM 物件) 與當下的 PR_CONTENT_COUNT 快照，用於快速顯示搜尋結果與驗證快取有效性

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
        Dim AttCount As Integer
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
        Dim AttMailList As List(Of MailItemInfo)    ' 所有 hasAttachment 候選 (無大小篩選)
        Dim ItemCountSnap As Long                   ' 快取當下的 PR_CONTENT_COUNT，失效偵測用
    End Structure

    ' 2026/06/12 by Simon/Claude Opus 4.8: Compiled Regex，程式啟動時編譯一次，後續呼叫零額外開銷
    ' Pattern 說明：^ 錨定開頭；[：:] 同吃半形/全形冒號；外層 + 一次處理所有巢狀前綴
    Private Shared ReadOnly _subjectPrefixRe As New Regex(
        "^(?:(?:RE|FW|FWD|AW|WG|VS|Rép|TR|回覆|回信|轉寄|轉發|回复|答复|转发|返信|転送|답장|회신|전달)\s*[：:]\s*)+",
        RegexOptions.IgnoreCase Or RegexOptions.Compiled)

    ' 2026/07/03 by Simon/Claude Fable 5: GetCleanSubject 的 memoization 快取。
    '   真實信箱裡 Re:/FW: 討論串會讓大量郵件共用同一個乾淨主旨，30 萬封信換算下來重複度很高；
    '   LoadMailInfoBatch/LoadMailInfoCore 等批次載入路徑每列都會呼叫一次，原本逐列重跑 Regex.Replace，
    '   改成「相同 subject 字串只算一次」後，載入 30 萬列時大部分會直接命中這裡，省掉對應次數的 Regex 掃描。
    Private Shared _cacheCleanSubject As New ConcurrentDictionary(Of String, String)

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
    Private Sub TryMarshalRelease(ByRef obj As Object)
        Try
            If obj IsNot Nothing AndAlso Marshal.IsComObject(obj) Then Marshal.ReleaseComObject(obj)
        Catch ex As System.Exception
            _dbg("TryMarshalRelease 異常: ", ex.Message)
        Finally
            obj = Nothing
        End Try
    End Sub
    Private Async Function InitRdoSessionWithoutEULA() As Task
        _dbg(" ├ 開始") ' by Gemini, 2026/04/10: 調整縮排層級為 Level 1

        ' 2026-03-23 v3:
        '   Task.Run 包裝保留 (讓 UI 執行緒繼續跑 LoadStoreToTreeView，平行初始化)
        '   第一次執行競爭條件改用 Thread.Sleep(1) 在 Set() 前解決，
        '   確保 AutoDismiss 輪詢 loop 已執行第一次再放行 New RDOSession()
        Try
            ' 2026/07/03 by Simon/Claude Fable 5: SetWinEventHook 事件驅動隱藏 EULA
            '   輪詢 SW_HIDE 有跨執行緒競速: (1) hide 太晚 → DWM 已合成一幀就閃; (2) hide 太早 → 被 VCL ShowModal 重新顯示
            '
            ' v5.1 (2026/07/03) 修正: v5 把 hook 裝在 UI 執行緒 → OUTOFCONTEXT 事件送進 UI 執行緒佇列，
            '   但 UI 執行緒正被 New RDOSession 卡住無法 pump，事件到 Finally 卸 hook 時直接被丟棄 (callback 從未執行)。
            '   改開「專職 hook 執行緒」: 自己安裝 hook + 跑 GetMessage 訊息迴圈，事件毫秒級送達，不依賴 UI 執行緒。
            '   若實測仍偶發閃爍 (非同步送達跨過 vblank ~16ms)，再升級 WH_CBT 同步方案
            _dbg(" ├ 進度", $"UI thread id={GetCurrentThreadId()}")   ' 診斷: 與 WinEvent 的事件執行緒比對，確認 TEULAForm 活在哪條執行緒
            Dim hookReady As New System.Threading.ManualResetEventSlim(False)
            Dim hookThread As New System.Threading.Thread(
                Sub()
                    _eulaWinEventProc = AddressOf EulaWinEventProc
                    Dim hHook As IntPtr = SetWinEventHook(EVENT_OBJECT_SHOW, EVENT_OBJECT_SHOW,
                        IntPtr.Zero, _eulaWinEventProc,
                        CUInt(Environment.ProcessId), 0, WINEVENT_OUTOFCONTEXT)
                    _eulaHookThreadId = GetCurrentThreadId()
                    _dbg("    ├ WinEvent", $"hook 安裝{If(hHook = IntPtr.Zero, "失敗", "完成")}, hook thread id={_eulaHookThreadId}")
                    hookReady.Set()
                    If hHook = IntPtr.Zero Then Return

                    ' message pump: OUTOFCONTEXT callback 由本迴圈的 GetMessage 觸發執行
                    Dim m As New NativeMsg
                    While GetMessage(m, IntPtr.Zero, 0, 0) > 0
                        TranslateMessage(m) : DispatchMessage(m)
                    End While

                    UnhookWinEvent(hHook)   ' 與安裝同一執行緒卸載
                    _eulaWinEventProc = Nothing
                    _dbg("    ├ WinEvent", "hook 已卸載, hook thread 結束")
                End Sub)
            hookThread.IsBackground = True
            hookThread.Priority = System.Threading.ThreadPriority.AboveNormal
            hookThread.Start()
            hookReady.Wait(500)

            Dim threadStarted As New System.Threading.ManualResetEventSlim(False)
            AutoDismissRdoEULA(threadStarted)
            ' 等 AutoDismiss thread 確認輪詢已開始，最多等 500ms
            threadStarted.Wait(500)
            _dbg(" ├ 進度", "AutoDismiss thread 已就緒，開始 New RDOSession") ' by Gemini, 2026/04/10

            ' 2026/7/1 by simon, 所有RDO都已切換至獨立session的 _rdo2, 不再沿用 Outlook MAPI session, 讓原有的 _rdo 完全退役
            ' 2026/6/23 by Simon/Claude: 測試 Redemption 獨立 Session 資料隔離與效能差異
            Dim session2 As Redemption.RDOSession = Nothing
            session2 = New Redemption.RDOSession()
            session2.Logon(ProfileName:=_olNS.CurrentProfileName, Password:="", ShowDialog:=False, NewSession:=True)    ' 獨立session, 不沿用 Outlook MAPI session
            _rdo2 = session2
            _dbg(" ├ _rdo2 init OK", $"Version={_rdo2.Version}")

        Catch ex As System.Exception
            _rdo2 = Nothing
            _dbg("Redemption init FAIL", ex.Message)
        Finally
            ' 2026/07/03: EULA 只在 RDOSession 建立期間出現，過後即結束 hook 執行緒 (WM_QUIT → pump 迴圈退出 → 自行卸 hook)
            If _eulaHookThreadId <> 0 Then
                PostThreadMessage(_eulaHookThreadId, WM_QUIT, IntPtr.Zero, IntPtr.Zero)
                _eulaHookThreadId = 0
            End If
        End Try

    End Function
    ' 2026/07/03 by Simon/Claude Fable 5: SetWinEventHook 事件驅動隱藏 EULA (v5 主力 hide 機制)
    '   delegate 必須存成欄位保住參考，否則 GC 回收後 native callback 觸發 → CallbackOnCollectedDelegate crash
    Private _eulaHookThreadId As UInteger = 0   ' v5.1: 專職 hook 執行緒的 native id，Finally 用 WM_QUIT 結束它
    Private _eulaWinEventProc As WinEventDelegate
    Private Sub EulaWinEventProc(hHook As IntPtr, eventType As UInteger, hwnd As IntPtr,
                                 idObject As Integer, idChild As Integer,
                                 dwEventThread As UInteger, dwmsEventTime As UInteger)
        ' 只理會視窗本體的 SHOW 事件 (idObject=OBJID_WINDOW, idChild=0)，排除子物件/控制項雜訊
        If idObject <> OBJID_WINDOW OrElse idChild <> 0 Then Return

        Dim sb As New System.Text.StringBuilder(64)
        GetClassName(hwnd, sb, 64)
        Dim cls As String = sb.ToString()

        ' dwmsEventTime = 事件「產生」的 tick；與「送達」時刻相減 = 佇列延遲，評估是否夠即時 (>16ms 就可能閃一幀)
        Dim delayMs As Long = (Environment.TickCount And &HFFFFFFFFL) - dwmsEventTime
        If delayMs < 0 Then delayMs += &H100000000L

        If cls = "TEULAForm" Then
            ShowWindow(hwnd, SW_HIDE)
            If _iLikeNoisy Then _dbg("    ├ WinEvent", $"TEULAForm SHOW → 已隱藏 hWnd=0x{hwnd:X}, 佇列延遲={delayMs}ms, 事件執行緒={dwEventThread}")
        Else
            ' 診斷 (v5.1): hook 只活在 RDO init 這幾秒，全記錄不會洗版；「事件執行緒」可對照 UI thread id
            '   確認 TEULAForm 到底活在哪條執行緒 → 若要升級 WH_CBT 同步方案，這決定 hook 要裝在哪
            If _iLikeNoisy Then _dbg("    ├ WinEvent", $"SHOW: class={cls}, hWnd=0x{hwnd:X}, 延遲={delayMs}ms, 事件執行緒={dwEventThread}")
        End If
    End Sub
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
        '
        ' v5 (2026-07-03): 隱藏視窗的主力改由 SetWinEventHook (EulaWinEventProc) 事件驅動接手，
        '   本執行緒降級為「點擊器 + 備援 hide」：負責 PostMessage 點 'I agree' + 'Ok' 結束 modal loop
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
                ' v5.2 (2026/07/03 by Simon/Claude Fable 5): 隱藏前先搬到螢幕外。
                '   2026/07/03 log 實測: 本輪詢在 VCL Show 之前 ~60ms 就找到視窗，此時 SW_HIDE 是空包彈，
                '   之後 ShowModal 會把視窗重新顯示 (殘餘閃爍的真正來源，由 WinEvent 事後再藏)。
                '   先搬到 -32000 → 就算被重新顯示也在螢幕外，肉眼零閃爍。若 VCL 在 Show 時自行置中則此招失效 → 屆時升級 WH_CBT
                SetWindowPos(hWnd, IntPtr.Zero, -32000, -32000, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE)
                ShowWindow(hWnd, SW_HIDE)
                _dbg("    ├ 成功", $"TEULAForm 移出螢幕+隱藏 hWnd=0x{hWnd:X}") ' by Gemini, 2026/04/10

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
        ' 獨立 session 須 Logoff 再 release,否則 Outlook 關不乾淨。byPath 值是 byName 的參考, 不另釋放。
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
            _dbg(" ├ _rdo2 物件完整釋放")
        End If
        ' 2026/6/23 by Simon/AntiGravity: _rdo 是 piggyback 在 Outlook session 上，不需要 Logoff，但要確保 COM ref 釋放且欄位歸 Nothing
        ' 2026/7/1 by simon, 所有RDO都已切換至獨立session的 _rdo2, 不再沿用 Outlook MAPI session, 讓原有的 _rdo 完全退役
        _dbg(" ├ 結束")
    End Sub
#End Region
#Region "  ├ Layer2 UI 流程輔助"
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
        ' 2026/05/04 by Gemini 3 Flash, 減少 TreeView 節點批次添加時 List() 的 Resize 次數 (預分配容量先訂32, 通常不會在一個資料夾內還有超過32個子資料夾)
        ' 2026/5/20 by simon: 在node.Name屬性加上值, 以便後續可以使用TreeNode.Find()
        '
        ' 2026/07/10 by Simon/Claude Fable 5: store-root 快取 — 啟動時 5 棵 SimTree 拿同一份 _pstStoreList 連呼 5 次，
        '   每個 store 3 次 COM (GetRootFolder / .Name / .FullFolderPath) 有 4/5 是重複讀。
        '   root 資訊借放 _cacheFolderTree 的保留鍵 KEY_STORE_ROOTS，第 2 棵樹起純記憶體組節點 (~0ms)。
        '   失效跟著 _cacheFolderTree 走 (ClearMemoryCachesCore / _showAllFolders 切換的 Clear 都會清到)，與 _pstStoreList 的生命週期一致。
        ' ===========================================================
        _dbg(" ├ 開始", tv.Name)

        Dim roots As List(Of (f As Folder, name As String, fPath As String)) = Nothing
        If Not _cacheFolderTree.TryGetValue(KEY_STORE_ROOTS, roots) Then
            roots = New List(Of (f As Folder, name As String, fPath As String))(storeList.Count)
            For Each store In storeList
                Dim root As Folder = store.GetRootFolder
                roots.Add((root, root.Name, root.FullFolderPath))
            Next
            _cacheFolderTree(KEY_STORE_ROOTS) = roots
        End If

        Dim nodeList As New List(Of TreeNode)(32)
        For Each r In roots
            Dim node As New TreeNode(r.name) With {.Tag = r.f, .Name = r.fPath}
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
        ' 2026/04/20 by Gemini 2.0 Flash: 若 Tab4 處於搜尋結果顯示模式，則不執行自動資料夾載入
        ' 2024/5/17重寫，優化資料夾載入邏輯，加入多層級展開的快取機制，並修復 Tab4 搜尋結果模式下的誤觸發問題
        ' 2024/5/19試過Task.Run(), Parallel.Foreach跟LINQ擴充方法了, 都沒有比較快, 別再試了, 就算virtual mode也沒有比我現在的lazy load還快
        ' 2024/5/20昨天才說不會更快了, 今天改用Nodes.AddRange(), 又更快了一點, 連BeginUpdate/EndUpdate都不需要了
        '
        ' 2026/4/7 by Gemini, 光速版子資料夾加號預測 HasSubFoldersFast() (專為 TreeView 展開設計)
        ' 2026/05/04 by Gemini 3 Flash, 減少 TreeView 節點批次添加時 List() 的 Resize 次數 (預分配容量先訂16, 通常不會在一個資料夾內還有超過16個子資料夾)
        ' 2026/5/20 by simon: 在node.Name屬性加上值, 以便後續可以使用TreeNode.Find()
        ' 2026/5/31 by Gemini/Simon: 加入 skipCache 引數判斷是否要強制讀取COM
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
    Private Async Function GetUniqueFolderList(selectedNodes As List(Of TreeNode), includeSub As Boolean, cToken As CancellationToken, Optional progress As IProgress(Of ProgressReport) = Nothing) As Task(Of List(Of (eid As String, sid As String, fPath As String)))
        ''' <summary>
        ''' 共用邏輯：將多個 TreeNode 轉換為無重複的實體資料夾清單
        ''' 2026/04/16 by Gemini: 升級回傳 Tuple (Folder, fPath)，消除呼叫端對 COM .FolderPath 的一次性集體讀取
        ''' 2026/06/28 by Simon/Claude [Stage2]: 回傳合約改 (eid,sid,fPath) 純資料 tuple(不帶 COM)。本函數內部僅用 .fPath,純型別跟改。
        ''' </summary>
        _dbg(" ├ 開始")
        ' 預分配容量為 512，優化多選資料夾後的路徑合併清單處理 (by Gemini 3 Flash, 2026/05/04)
        Dim fList As New List(Of (eid As String, sid As String, fPath As String))(512)
        Dim addedPaths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        For Each node As TreeNode In selectedNodes
            cToken.ThrowIfCancellationRequested()
            Dim rootF = TryCast(node.Tag, Folder)
            If rootF Is Nothing Then Continue For

            ' GetSubtree 回傳的 subF 現在是 (eid, sid, fPath) Tuple
            ' 2026/04/17 by Claude: 改呼叫 GetSubtree (L2.5)，原 GetSubtree 已改名為 L3
            ' 2026/06/14 by Simon/Claude Opus 4.8: 去模式化後 GetSubtree 回傳「完整骨架(含非郵件夾)」，模式過濾移到計數層。
            '   Tab2-5 多選掃描共用此函數，必須在此補上 FilterSubtreeByMode，否則過濾模式下會把非郵件夾(行事曆/連絡人…)
            '   也納入掃描範圍 (狀態列出現 333 而非 ~307)。showAll 模式 FilterSubtreeByMode 回傳全集，行為不變。
            Dim subTree = Await GetSubtree(rootF, includeSub, progress:=progress, cToken:=cToken)
            For Each subF In FilterSubtreeByMode(subTree, SafeGetPath(rootF))
                If addedPaths.Add(subF.fPath) Then fList.Add(subF)  ' ✅ 直接讀取 subF.fPath (Tuple 屬性)，再也不用打 COM!
            Next
            Await Task.Yield()
        Next
        Return fList
    End Function
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
    Private Function IsSubtreeComplete(rows As List(Of FolderInfoDbRow), rootPath As String) As Boolean
        ' ---------------------------------------------------------------
        ' 2026/06/13 by Simon/Claude Opus 4.8: 子樹骨架完整性檢查 (無模式分支 — 因骨架本就完整)
        ' 原理: 限定 rootPath 子樹範圍(避開 LIKE 前綴誤匹配 sibling，例如 Inbox 誤匹配 Inbox2)，
        '       對每個資料夾 F 要求「集合內 F 的直屬子夾數 == fc(F) 未過濾」。
        '       fc < 0(未知) 或對不上 → 判殘缺 → 由呼叫端 fallback L3 完整重掃。
        ' 注意: 依賴 LazyGetSubFolderIDAsList 已 SELECT folder_count 並填入 row.fc (2026/06/13 配套修改)。
        ' ---------------------------------------------------------------
        If rows Is Nothing OrElse rows.Count = 0 Then Return False

        Dim prefix As String = rootPath & "\"        ' 只保留 rootPath 子樹範圍內的列 (path == root 或 startsWith root & "\")
        Dim inScope As New List(Of FolderInfoDbRow)(rows.Count)
        Dim hasRoot As Boolean = False               ' 2026/07/04 by Simon/Claude Fable 5 [rootless 骨架未爆彈]
        For Each r In rows
            If r.path = rootPath OrElse r.path.StartsWith(prefix, StringComparison.Ordinal) Then inScope.Add(r)
            If r.path = rootPath Then hasRoot = True
        Next
        If inScope.Count = 0 Then Return False
        ' 2026/07/04 by Simon/Claude Fable 5: root 本人不在集合 → 判殘缺。原檢查只驗每列 fc,漏了「root 列缺席」這一格 —
        '   缺 root 的骨架會讓 FilterSubtreeByMode(從 root 起走)回空集合,整條計數/顯示鏈靜默全滅。
        If Not hasRoot Then Return False

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
    Private Function FilterSubtreeByMode(skeleton As List(Of (eid As String, sid As String, fPath As String)), rootPath As String) As List(Of (eid As String, sid As String, fPath As String))
        ' ---------------------------------------------------------------
        ' 2026/06/13 by Simon/Claude Opus 4.8: 計數/顯示層的模式過濾 (剪枝移到這裡，骨架層永遠完整、0 COM)
        ' 依 _showAllFolders 從完整骨架即時派生:
        '   全顯(True) : 全數回傳。
        '   關閉(False): 從 root 沿 is_mail 的夾往下剪枝走訪 (碰非郵件夾不往下數)。root 一律納入(比照原 BFS 行為)。
        ' is_mail 來源: _cacheFolderIDs (L3 BFS 與 L2.5 DB 重建兩條路徑皆已回填)；查無時保守視為 is_mail(納入)，避免少算。
        ' 2026/06/28 by Simon/Claude [Stage2]: 進出型別改 (eid,sid,fPath)。本函數僅用 .fPath + _cacheFolderIDs 的 isMail,從不 deref folder,純型別跟改。
        ' ---------------------------------------------------------------
        If _showAllFolders Then Return skeleton

        Dim byPath As New Dictionary(Of String, (eid As String, sid As String, fPath As String))(skeleton.Count)
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

        Dim result As New List(Of (eid As String, sid As String, fPath As String))(skeleton.Count)
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
    '   - GetFolderSize(pFolder)            ' 單一資料夾大小，有 DB lazy
    '   - GetMailCountAllAsync(pFolder)     ' 整棵子樹郵件總數，有 DB lazy
    '   - GetFolderCountAllAsync(pFolder)   ' 整棵子樹資料夾總數，有 DB lazy
    '   - GetFolderSizeAll(pFolder)         ' 整棵子樹大小，有 DB lazy
    '   - GetAttMailList(pFolder)           ' Tab3 Phase1，有 DB lazy (att_maillist)
    '   - GetAttFilename(mail)              ' Tab3 Phase2，有 DB lazy (att_filenames)
    '   - GetSubtree(rootFolder)            ' 整棵子資料夾清單，有 DB lazy (2026/04/17)
    '   - GetYearCount(sFolder, fPath:=fPath) ' 單一資料夾年份分佈，有 DB lazy (2026/04/17)
    '   - GetMonthCount(sFolder, year)        ' 單一資料夾月份分佈，有 DB lazy + 提前過濾 (2026/04/17)
    ' ---------------------------------------------------------------
    ' 2026/04/07: Phase 2 — 在記憶體 miss 時加入 SQLite lazy SELECT，命中後一次填滿所有欄位
    '             寫入仍由 SaveCachesToDB (SaveCache 按鈕) 批次處理，本層不做即時寫入
    ' 2026/4/7 by Claude, 加入「持久化快取」存入SSD:
    '             讀資料時若快取不存在就先去撈SSD, 有就放進快取, 沒有才算cache miss去COM讀取
    ' ---------------------------------------------------------------
    ' 呼叫順序 (每個 Layer2.5 函數):
    '   ① 記憶體命中 → 直接回傳 (最快，0 COM call)
    '   ② DB 命中 + snapshot 驗證通過 → 填滿記憶體快取 → 回傳 (快，0 COM call)
    '   ③ DB miss 或 snapshot 不符 → 呼叫 Layer3 → 填入記憶體快取 → 回傳 (慢，有 COM call)
    ' ---------------------------------------------------------------
    Private Function GetMailCount(folder As Folder, Optional fPath As String = "", Optional skipCache As Boolean = False) As Long
        ' ---------------------------------------------------------------
        ' GetMailCount — 單一資料夾本層郵件數 (PR_CONTENT_COUNT)
        ' 2026/4/7 by Claude, 加入Database lazy load, 讀資料順序:
        '   ① 記憶體命中
        '   ② DB lazy load：命中且 mc 有效且 snapshot 吻合 → 一次填滿所有欄位
        '   ③ 讀取派工: _rdo2 在(且 store 可解) → GetMailCountRdo;否則 → GetMailCountOOM(OOM)
        ' 2026/04/15 by Gemini 3.1 Pro, 加入 optional fPath 參數，若有傳入則可省去 pFolder.FolderPath 1ms 耗時
        ' 2026/05/09 by Gemini 3.1 Pro: 重構成使用 SafeGetDbRow，回傳統一改成 Long 共用函式
        ' 2026/06/23 by Simon/Claude Opus 4.8: ③ 改為 RDO 派工(GetMailCountRdo via _rdo2,失敗 fallback OOM L3);
        '   加 skipCache 引數(繞過快取讀寫,給 F5 skipCache / snap 重讀等直呼者用,仍走 RDO 派工)。
        ' 2026/6/27 by simon/Claude: 雖然讀取跳過快取, 強制讀COM, 但是讀回來的值還是應該要回填快取, 更新到最正確的內容
        ' ---------------------------------------------------------------
        fPath = SafeGetPath(folder, fPath)
        If _iLikeNoisy Then _dbg(" ├ 開始", fPath) ' by Gemini 3.5 Flash, 2026/07/01
        Dim count As Long
        If Not skipCache Then
            If _cacheMailCount.TryGetValue(fPath, count) Then Return count       ' ① 記憶體命中
            Dim row = SafeGetDbRow(folder, fPath)                                ' ② DB lazy load
            If row IsNot Nothing AndAlso row.mc >= 0 Then Return row.mc
        End If

        ' ③ 讀取派工: RDO 優先,失敗 fallback OOM
        count = GetMailCountRdo(fPath, folder.EntryID, folder.StoreID)
        If count < 0 Then count = GetMailCountOOM(folder, fPath:=fPath)
        If count >= 0 Then _cacheMailCount.TryAdd(fPath, count)             ' 2026/6/27 by simon/Claude: 雖然讀取跳過快取, 強制讀COM, 但是讀回來的值還是應該要回填快取, 更新到最正確的內容
        If _iLikeNoisy Then _dbg(" ├ 結束", fPath & " | 成果: " & count)    ' by Gemini 3.5 Flash, 2026/07/01
        Return count

    End Function
    Private Function GetMailCount(fPath As String, eid As String, sid As String, Optional skipCache As Boolean = False) As Long
        ' 2026/06/28 by Simon/Claude [Stage1+R2a]: 免-folder 多載(D1=a 同名;D2=ii 既有 folder 版不動)。給資料 tuple 消費者用,免去為拿 eid/sid 而物化整棵樹。
        '   ① 記憶體(by fPath) → ② DB lazy(by fPath 直讀 folder_info,信任 DB 不做 snap;暖快取重啟回快路徑) → ③ RDO 優先,失敗才 GetFolderById+OOM(守底線)。
        '   R2a 取捨: ② 不做 snap(snap 需 live folder,免-folder 做不了),暖路徑信任 DB;要最新值走 skipCache=True(F5)直打 ③ RDO。PST 封存讀多寫少,信任 DB 安全。
        If _iLikeNoisy Then _dbg(" ├ 開始", fPath) ' by Gemini 3.5 Flash, 2026/07/01
        Dim count As Long
        If Not skipCache Then
            If _cacheMailCount.TryGetValue(fPath, count) Then Return count      ' ① 記憶體快取命中
            Dim dbRow = LazyGetFolderInfo(fPath)                                  ' ② DB lazy(信任 DB,不做 snap)
            If dbRow IsNot Nothing AndAlso dbRow.mc >= 0 Then Return dbRow.mc
        End If
        count = GetMailCountRdo(fPath, eid, sid)                                ' ③ RDO 優先
        If count < 0 Then                                                       ' 底線: RDO 失敗 → OOM
            Dim f As Folder = GetFolderById(eid, sid)
            If f IsNot Nothing Then count = GetMailCountOOM(f, fPath:=fPath)
        End If
        If count >= 0 Then _cacheMailCount.TryAdd(fPath, count)
        If _iLikeNoisy Then _dbg(" ├ 結束", fPath & " | 成果: " & count) ' by Gemini 3.5 Flash, 2026/07/01
        Return count
    End Function
    Private Function GetFolderCount(folder As Folder, Optional fPath As String = "", Optional skipCache As Boolean = False) As Long
        ' ---------------------------------------------------------------
        ' GetFolderCount — 單一資料夾直屬子資料夾數
        ' 2026/4/7 by Claude, 加入Database lazy load, 讀資料順序:
        '   ① 記憶體命中
        '   ② DB lazy load：命中且 fc 有效且 snapshot 吻合 → 一次填滿所有欄位
        '   ③ 讀取派工: _rdo2 在(且 store 可解) → GetFolderCountRdo;否則 → GetFolderCountOOM(OOM)
        ' 2026/05/09 by Gemini 3.1 Pro: 重構成使用 SafeGetDbRow，回傳統一改成 Long 共用函式
        ' 2026/06/23 by Simon/Claude Opus 4.8: ③ 改為 RDO 派工(GetFolderCountRdo via _rdo2,失敗 fallback OOM L3);
        '   加 skipCache 引數(繞過快取讀寫,給 F5 skipCache / snap 重讀等直呼者用,仍走 RDO 派工)。
        ' 2026/6/27 by simon/Claude: 雖然讀取跳過快取, 強制讀COM, 但是讀回來的值還是應該要回填快取, 更新到最正確的內容
        ' ---------------------------------------------------------------
        fPath = SafeGetPath(folder, fPath)
        If _iLikeNoisy Then _dbg(" ├ 開始", fPath)    ' by Gemini 3.5 Flash, 2026/07/01
        Dim count As Long
        If Not skipCache Then
            If _cacheFolderCount.TryGetValue(fPath, count) Then Return count     ' ① 記憶體命中
            Dim row = SafeGetDbRow(folder, fPath)                                ' ② DB lazy load (fc 欄位)
            If row IsNot Nothing AndAlso row.fc >= 0 Then Return row.fc
        End If

        ' ③ 讀取派工: RDO 優先,失敗 fallback OOM
        count = GetFolderCountRdo(fPath, folder.EntryID, folder.StoreID)
        If count < 0 Then count = GetFolderCountOOM(folder, fPath:=fPath)
        If count >= 0 Then _cacheFolderCount.TryAdd(fPath, count)           ' 2026/6/27 by simon/Claude: 雖然讀取跳過快取, 強制讀COM, 但是讀回來的值還是應該要回填快取, 更新到最正確的內容
        If _iLikeNoisy Then _dbg(" ├ 結束", fPath & " | 成果: " & count)    ' by Gemini 3.5 Flash, 2026/07/01
        Return count

    End Function
    Private Function GetFolderCount(fPath As String, eid As String, sid As String, Optional skipCache As Boolean = False) As Long
        ' 2026/06/28 by Simon/Claude [Stage1+R2a]: 免-folder 多載。① 記憶體 → ② DB lazy(by fPath 直讀 folder_info,信任 DB) → ③ RDO 優先,失敗 GetFolderById+OOM(守底線)。
        If _iLikeNoisy Then _dbg(" ├ 開始", fPath)    ' by Gemini 3.5 Flash, 2026/07/01
        Dim count As Long
        If Not skipCache Then
            If _cacheFolderCount.TryGetValue(fPath, count) Then Return count    ' ① 記憶體快取命中
            Dim dbRow = LazyGetFolderInfo(fPath)                                  ' ② DB lazy
            If dbRow IsNot Nothing AndAlso dbRow.fc >= 0 Then Return dbRow.fc
        End If
        count = GetFolderCountRdo(fPath, eid, sid)                              ' ③ RDO 優先
        If count < 0 Then                                                       ' 底線: RDO 失敗 → OOM
            Dim f As Folder = GetFolderById(eid, sid)
            If f IsNot Nothing Then count = GetFolderCountOOM(f, fPath:=fPath)
        End If
        If count >= 0 Then _cacheFolderCount.TryAdd(fPath, count)
        If _iLikeNoisy Then _dbg(" ├ 結束", fPath & " | 成果: " & count)        ' by Gemini 3.5 Flash, 2026/07/01
        Return count
    End Function
    Private Async Function GetFolderSize(folder As Folder, Optional fPath As String = "", Optional skipCache As Boolean = False, Optional cToken As CancellationToken = Nothing) As Task(Of Long)
        ' ---------------------------------------------------------------
        ' GetFolderSize — 單一資料夾本層大小 (GetTable 加總)
        ' 2026/3/29 by Gemini: Layer2.5 快取代理層 - 取得單一資料夾本層的大小 (含快取機制)
        ' 2026/4/7 by Claude, 加入Database lazy load, 讀資料順序:
        '   ① 記憶體命中
        '   ② DB lazy load：命中且 fs 有效且 snapshot 吻合 → 一次填滿所有欄位
        '   ③ fallback: Layer3 呼叫
        ' 2026/05/09 by Gemini 3.1 Pro: 重構成使用 SafeGetDbRow，回傳統一改成 Long 共用函式
        ' 2026/06/27 by Simon/Claude Opus 4.8: 加 skipCache 引數(繞過快取讀寫,給 snap 重讀/強制重讀者用,仍走 RDO 派工),對齊 GetMailCount/GetFolderCount。
        ' ---------------------------------------------------------------
        fPath = SafeGetPath(folder, fPath)
        If _iLikeNoisy Then _dbg(" ├ 開始", fPath) ' by Gemini 3.5 Flash, 2026/07/01
        Dim size As Long
        If Not skipCache Then
            If _cacheFolderSize.TryGetValue(fPath, size) Then Return size   ' ① 記憶體命中
            Dim row = SafeGetDbRow(folder, fPath)
            If row IsNot Nothing AndAlso row.fs >= 0 Then Return row.fs     ' ② DB lazy load (fs 欄位)
        End If

        ' ③ COM讀取派工: RDO 優先, 失敗 fallback OOM (對齊 GetFolderCountAsync ③)
        ' 2026/06/27 by Simon/Claude Opus 4.8: 補上 _rdo2 size 槽。GetFolderSizeRdo 走 PR_MESSAGE_SIZE(PT_LONG) GetRows 加總(探針實證快 3~10×、parity 全一致);回 <0 才掉 OOM L3。
        size = GetFolderSizeRdo(fPath, folder.EntryID, folder.StoreID)
        If size < 0 Then size = Await GetFolderSizeOOM(folder, fPath:=fPath, cToken:=cToken)
        If size >= 0 Then _cacheFolderSize.TryAdd(fPath, size)
        If _iLikeNoisy Then _dbg(" ├ 結束", fPath & " | 成果: " & size)     ' by Gemini 3.5 Flash, 2026/07/01
        Return size

    End Function
    Private Async Function GetFolderSize(fPath As String, eid As String, sid As String, Optional skipCache As Boolean = False, Optional cToken As CancellationToken = Nothing) As Task(Of Long)
        ' 2026/06/28 by Simon/Claude [Stage1+R2a]: 免-folder 多載。① 記憶體 → ② DB lazy(by fPath 直讀 folder_info,信任 DB) → ③ RDO 優先,失敗 GetFolderById+OOM(守底線)。
        If _iLikeNoisy Then _dbg(" ├ 開始", fPath) ' by Gemini 3.5 Flash, 2026/07/01
        Dim size As Long
        If Not skipCache Then
            If _cacheFolderSize.TryGetValue(fPath, size) Then Return size       ' ① 記憶體快取命中
            Dim dbRow = LazyGetFolderInfo(fPath)                                  ' ② DB lazy
            If dbRow IsNot Nothing AndAlso dbRow.fs >= 0 Then Return dbRow.fs
        End If

        size = GetFolderSizeRdo(fPath, eid, sid)                                ' ③ RDO 優先
        If size < 0 Then                                                        ' 底線: RDO 失敗 → OOM
            Dim f As Folder = GetFolderById(eid, sid)
            If f IsNot Nothing Then size = Await GetFolderSizeOOM(f, fPath:=fPath, cToken:=cToken)
        End If
        If size >= 0 Then _cacheFolderSize.TryAdd(fPath, size)
        If _iLikeNoisy Then _dbg(" ├ 結束", fPath & " | 成果: " & size) ' by Gemini 3.5 Flash, 2026/07/01
        Return size
    End Function
    Private Async Function GetFolderSizeAll(folder As Folder, Optional fPath As String = "", Optional skipCache As Boolean = False, Optional cToken As CancellationToken = Nothing) As Task(Of Long)
        ' ---------------------------------------------------------------
        ' GetFolderSizeAll — 整棵子樹大小總計
        ' 2026/3/29 by Gemini: Layer2.5 快取代理層 - 取得整棵子樹的大小總計 (含快取機制)
        ' 2026/4/7 by Claude, 加入Database lazy load, 讀資料順序:
        '   ① 記憶體命中
        '   ② DB lazy load：命中且 fsa 有效且 snapshot 吻合 → 一次填滿所有欄位
        '   ③ fallback: Layer3 呼叫
        ' 2026/04/15 by Claude: 加入 cToken 參數，取代 _cancelRequested 旗標
        ' 2026/05/09 by Gemini 3.1 Pro: 重構成使用 SafeGetDbRow，回傳統一改成 Long 共用函式
        ' 2026/06/27 by Simon/Claude Opus 4.8: 加 skipCache(對齊 GetFolderSize)。Tab12 F5 用——TryRemove 只清記憶體,② DB lazy 仍會回 fsa,skipCache 才真跳過 ①②。
        ' 2026/6/27 by simon/Claude: 雖然讀取跳過快取, 強制讀COM, 但是讀回來的值還是應該要回填快取, 更新到最正確的內容
        ' ---------------------------------------------------------------
        fPath = SafeGetPath(folder, fPath)
        If _iLikeNoisy Then _dbg(" ├ 開始", fPath) ' by Gemini 3.5 Flash, 2026/07/01
        Dim size As Long
        If Not skipCache Then
            If _cacheFolderSizeAll.TryGetValue(fPath, size) Then Return size    ' ① 記憶體命中
            Dim row = SafeGetDbRow(folder, fPath)                               ' ② DB lazy load (fsa 欄位)
            If row IsNot Nothing AndAlso row.fsa >= 0 Then Return row.fsa
        End If


        size = Await GetFolderSizeAllOOM(folder, skipCache:=skipCache, cToken:=cToken)  ' ③ fallback: Layer3 呼叫
        If size >= 0 Then _cacheFolderSizeAll.TryAdd(fPath, size)       ' 2026/6/27 by simon/Claude: 雖然讀取跳過快取, 強制讀COM, 但是讀回來的值還是應該要回填快取, 更新到最正確的內容
        If _iLikeNoisy Then _dbg(" ├ 結束", fPath & " | 成果: " & size) ' by Gemini 3.5 Flash, 2026/07/01
        Return size

    End Function
    Private Async Function GetYearCount(fPath As String, eid As String, sid As String, cToken As CancellationToken) As Task(Of ConcurrentDictionary(Of Integer, Integer))
        ' ---------------------------------------------------------------
        ' GetYearCount — 單一資料夾年份郵件分佈 (Layer2.5 快取代理)
        ' 2026/04/17 by Claude: 從 CollectYearCount (L2) 拆出，對齊其他 Layer2.5 快取函數架構
        ' 呼叫順序: ① 記憶體命中 → ② DB lazy load → ③ 讀取派工(RDO 優先,失敗 fallback OOM)
        ' OCE 不在此攔截，直接 re-throw 讓 CollectYearCount (L2) 的 Catch OCE 接住
        ' 2026/06/29 by Simon/Claude [Stage2]: 改免-folder 簽章(fPath,eid,sid)。熱路徑①②只靠 fPath,folder 延後到 ③ 才 GetFolderById 物化,消除每夾眼物化的 COM 稅。
        ' 2026/07/02 by Simon/Claude [PROBE_YEARSQL 驗證通過]: ③ 改為 RDO 派工(GetYearCountRdo via ExecSQL,免物化folder,
        '   失敗才 fallback GetFolderById + GetYearCountOOM)。探針驗證: TOP1+ORDER BY 範圍偵測 55/55、6/6 兩子樹全數相符,
        '   加計範圍偵測固定成本後仍比純 OOM 快 1.3x~4x。
        ' 2026/07/03 by Simon 註解: YearCount/MonthCount多線程平行化: 技術上可行但不值得, 在17~30萬封郵件的全庫重建也都不到二秒, 再砍也只快不到一秒反而增加多線程同步複雜度, 先不做。
        ' ---------------------------------------------------------------
        If _iLikeNoisy Then _dbg(" ├ 開始", fPath & " | ID: " & eid)
        Dim value As ConcurrentDictionary(Of Integer, Integer) = Nothing
        If _cacheYearCount.TryGetValue(fPath, value) Then     ' ① 記憶體命中
            If _iLikeNoisy Then _dbg("    ├ Cache Hit: ", ExtractFolderName(fPath))
            Return value
        End If

        Dim dbResult = LazyGetYearCount(fPath)          ' ② DB lazy load
        If dbResult IsNot Nothing Then
            _cacheYearCount(fPath) = dbResult
            If _iLikeNoisy Then _dbg("    ├ DB Hit: ", ExtractFolderName(fPath))
            Return dbResult
        End If

        If _iLikeNoisy Then _dbg("    ├ Cache miss: ", ExtractFolderName(fPath))
        ' ③ 讀取派工: RDO 優先(免物化 folder),失敗才 GetFolderById 物化 + OOM
        Dim folderResult = GetYearCountRdo(fPath, eid, sid)
        If folderResult Is Nothing Then
            Dim folder As Folder = GetFolderById(eid, sid)      ' 只有掉到 OOM 才物化
            folderResult = Await GetYearCountOOM(folder, fPath:=fPath, cToken:=cToken) ' ③b Layer3 COM；OCE re-throw 至 L2
        End If
        _cacheYearCount(fPath) = folderResult                  ' ✅ OCE 時走不到此行，快取僅在完整計算後寫入
        MarkMailFolderDirty(fPath)   ' 2026/07/03 by Simon/Claude: dirty 追蹤
        Return folderResult

    End Function
    Private Async Function GetMonthCount(fPath As String, eid As String, sid As String, year As Integer, cToken As CancellationToken) As Task(Of ConcurrentDictionary(Of Integer, Integer))
        ' ---------------------------------------------------------------
        ' GetMonthCount — 單一資料夾指定年份月份分佈 (Layer2.5 快取代理)
        ' 2026/04/17 by Claude: 從 GetMonthCountOOM 拆出快取與提前過濾邏輯
        '   原來的快取/過濾邏輯混在 L3 裡，現在統一到此 L2.5 層，L3 只剩純 COM
        ' 呼叫順序:
        '   提前過濾 1 — GetMailCount=0   → 直接回傳空，不打 COM
        '   提前過濾 2 — _cacheYearCount 已知該年無信 → 直接回傳空，不打 COM
        '   ① 記憶體命中 → ② DB lazy load → ③ 讀取派工(RDO 優先,失敗 fallback OOM)
        ' OCE 不在此攔截，直接 re-throw 讓 CollectMonthCount (L2) 的 Catch OCE 接住
        ' 2026/06/29 by Simon/Claude [Stage2]: 改免-folder 簽章(fPath,eid,sid 領頭),folder 延後到 ③才物化
        ' 2026/07/02 by Simon/Claude: ③ 比照 GetYearCount 套用 RDO 派工(年份架構已驗證,月份不另外測,直接套用)。
        ' 2026/07/03 by Simon 註解: YearCount/MonthCount多線程平行化: 技術上可行但不值得, 在17~30萬封郵件的全庫重建也都不到二秒, 再砍也只快不到一秒反而增加多線程同步複雜度, 先不做。
        ' ---------------------------------------------------------------
        If _iLikeNoisy Then _dbg(" ├ 開始", fPath & " | Year: " & year) ' by Gemini 3.5 Flash, 2026/07/01

        ' 提前過濾 1: 該資料夾完全無郵件，不必查快取或打 COM
        ' 2026/04/10 by Gemini: 解決 DB 沒存 0 封信記錄，lazy_load 回 Nothing 被迫打 COM 的問題
        If GetMailCount(fPath, eid, sid) = 0 Then Return New ConcurrentDictionary(Of Integer, Integer)()

        ' 提前過濾 2: 年度快取已知此年份信件數為 0，不必打月份 COM
        ' 2026/04/10 by Gemini: 省掉「某資料夾在 2001 年確定無信」的多餘 COM 呼叫
        Dim yCache As ConcurrentDictionary(Of Integer, Integer) = Nothing
        If _cacheYearCount.TryGetValue(fPath, yCache) Then
            Dim countInYear As Integer = 0
            yCache.TryGetValue(year, countInYear)
            If countInYear = 0 Then Return New ConcurrentDictionary(Of Integer, Integer)()
        End If

        Dim cacheKey As String = fPath & "_" & year.ToString()
        Dim value As ConcurrentDictionary(Of Integer, Integer) = Nothing
        If _cacheMonthCount.TryGetValue(cacheKey, value) Then Return value ' ① 記憶體命中

        Dim dbResult = LazyGetMonthCount(fPath, year)               ' ② DB lazy load
        If dbResult IsNot Nothing Then
            _cacheMonthCount.TryAdd(cacheKey, dbResult)
            If _iLikeNoisy Then _dbg("DB 命中", $"{ExtractFolderName(fPath)} {year} 年 ({dbResult.Count} 個月)")
            Return dbResult
        End If

        ' ③ 讀取派工: RDO 優先(免物化 folder),失敗才 GetFolderById 物化 + OOM;OCE re-throw，不在此攔截 (寫入快取在 COM 完成後，OCE 天然繞過)
        Dim monthCount = GetMonthCountRdo(fPath, eid, sid, year)
        If monthCount Is Nothing Then
            Dim folder As Folder = GetFolderById(eid, sid)   ' 只有掉到 OOM 才物化
            monthCount = Await GetMonthCountOOM(folder, year, fPath:=fPath, cToken:=cToken)
        End If
        _cacheMonthCount(cacheKey) = monthCount           ' ✅ 完整計算後存入快取
        MarkMailFolderDirty(fPath)   ' 2026/07/03 by Simon/Claude: dirty 追蹤 (用 fPath，非 cacheKey；cacheKey 含 "_year" 後綴)
        ' DbSaveMonthCountSingle(fPath, year, monthCount) ' ✅ 2026/04/09 設計: 增量寫入 DB (待啟用)
        If _iLikeNoisy Then _dbg(" ├ 結束", fPath & " | Year: " & year & " | 成果: " & If(monthCount IsNot Nothing, monthCount.Count.ToString(), "Nothing")) ' by Gemini 3.5 Flash, 2026/07/01
        Return monthCount

    End Function
    Private Async Function GetAttMailList(fPath As String, eid As String, sid As String, progress As IProgress(Of ProgressReport), cToken As CancellationToken) As Task(Of List(Of MailItemInfo))
        ' ---------------------------------------------------------------
        ' GetAttMailList — Tab3 Phase1：含附件的候選郵件清單 (Layer2.5 快取代理)
        ' by Gemini, 2026/04/05: Tab3 Phase 1 快取 - 取得單一資料夾本層含附件的郵件清單
        ' 2026/4/7 by Claude: ① 記憶體命中 → ② DB lazy (att_maillist) → ③ Layer3
        ' 2026/06/29 by Simon/Claude [Stage2]: 改免-folder 簽章(fPath,eid,sid)。熱路徑①②只靠 fPath，folder 延後到 ③ 才 GetFolderById 物化，消除每夾眼物化的 COM 稅。(Tab3 為唯一呼叫者、OST 不使用，故直接改簽章不留多載)
        ' 
        ' 2026/07/02 by Simon/Claude: (原 todo「優先要新增 GetAttMailListRdo()，轉換至 _rdo2」— 實測後結案: 不做。
        '   hasAttachment 的迴紋針語意(排除 Hidden/olOLE)是 Outlook 查詢引擎的加工,RDO 端 PR_HASATTACH 多算 59%、
        '   逐封精篩慢 18 倍、Restrict 下推=raw PR_HASATTACH 且本身慢 6.7 倍, 三條路全實測封死,維持純 OOM。詳見 memory_20260623_2210 更正記錄。)
        ' 2026/07/04 by Simon/Claude Fable 5: 上面 07/02 的結案被 PROBE_ATTACHBATCH 推翻 — 當時缺口是沒試過 SmartNoAttach(0x8514)
        '   named prop 當第二欄。③ 改 RDO 優先(GetAttMailListRdo,parity 100%/178k列,快 1.8x~13x),失敗 fallback OOM 保底。
        ' ---------------------------------------------------------------
        If _iLikeNoisy Then _dbg(" ├ 開始", fPath)                ' by Gemini 3.5 Flash, 2026/07/01
        Dim currentCount As Long = GetMailCount(fPath, eid, sid)  ' 免-folder，依賴同層快取

        ' ① 記憶體命中
        Dim entry As FolderCacheTab3 = Nothing ' 補上初始化以消除 BC42108 警告
        If _cacheAttMailList.TryGetValue(fPath, entry) AndAlso entry.ItemCountSnap = currentCount Then Return entry.AttMailList

        ' ② DB lazy load (att_maillist)：pr_count_snap == currentCount → 快取仍有效
        Dim dbResult = LazyGetAttMailList(fPath)
        If dbResult IsNot Nothing AndAlso dbResult.Snap = currentCount Then
            Dim cached As New FolderCacheTab3 With {.AttMailList = dbResult.Mails, .ItemCountSnap = currentCount}
            _cacheAttMailList(fPath) = cached   ' 覆蓋式寫入，確保 ItemCountSnap 對應正確
            If _iLikeNoisy Then _dbg(" ├ DB 命中", $"{fPath} ({dbResult.Mails.Count} 封)")
            Return dbResult.Mails
        End If

        ' ③ 讀取派工: RDO 優先(免物化 folder),回 Nothing 才 GetFolderById 物化 → OOM 保底
        ' 2026/04/15: 加入 cToken 傳遞，取消時 GetAttMailListOOM 回傳空 List，不寫入快取
        ' 2026/07/04 by Simon/Claude Fable 5: RDO 優先派工(形狀對齊 GetMailCount/GetMailBody 的 ③);RDO 路徑單夾 ms 級,不吃 cToken
        Dim targetMailList As List(Of MailItemInfo) = Nothing
        If _rdo2 IsNot Nothing Then targetMailList = GetAttMailListRdo(fPath, eid, sid)
        If targetMailList Is Nothing Then
            Dim folder As Folder = GetFolderById(eid, sid)
            targetMailList = Await GetAttMailListOOM(folder, progress, cToken:=cToken)
        End If
        _cacheAttMailList(fPath) = New FolderCacheTab3 With {.AttMailList = targetMailList, .ItemCountSnap = currentCount}
        MarkMailFolderDirty(fPath)   ' 2026/07/03 by Simon/Claude: dirty 追蹤
        If _iLikeNoisy Then _dbg(" ├ 結束", fPath & " | 成果: " & If(targetMailList IsNot Nothing, targetMailList.Count.ToString(), "Nothing")) ' by Gemini 3.5 Flash, 2026/07/01
        Return targetMailList
    End Function
    Private Function GetAttFilename(ByRef mail As MailItemInfo, Optional skipCache As Boolean = False) As List(Of String)
        ' ---------------------------------------------------------------
        ' GetAttFilename — Tab3 Phase2：附件檔名清單 (by EntryID)
        ' by Gemini, 2026/04/04: Layer2.5 快取代理層 - 取得附件檔名清單 (含 _cacheAttFilename 機制)
        ' 2026/4/7 by Claude, 加入Database lazy load, 讀資料順序:
        '   ① 記憶體命中
        '   ② DB lazy load：命中且 fc 有效且 snapshot 吻合 → 一次填滿所有欄位
        '   ③ fallback: Layer3 呼叫
        ' ---------------------------------------------------------------
        If _iLikeNoisy Then _dbg(" ├ 開始", mail.FolderPath & "\" & mail.Subject)
        Dim result As List(Of String) = Nothing
        If Not skipCache AndAlso _cacheAttFilename.TryGetValue(mail.EntryID, result) Then Return result  ' ①

        ' ② DB lazy load (att_filenames)
        If Not skipCache Then
            result = LazyGetAttFilenames(mail.EntryID)
            If result IsNot Nothing Then
                _cacheAttFilename.TryAdd(mail.EntryID, result)
                Return result
            End If
        End If

        ' ③ 讀取分派: _rdo2 在 → GetAttFilenameRdo(store-scoped 高速,store 找不到時內部回 Nothing);否則 → GetAttFilenameOOM(OOM)。RDO 已上移至 L2.5。 2026/06/23 by Simon/Claude
        If _rdo2 IsNot Nothing Then
            result = GetAttFilenameRdo(mail)
            If result Is Nothing Then result = GetAttFilenameOOM(mail)   ' RDO 解析失敗保底
        Else
            result = GetAttFilenameOOM(mail)
        End If
        If result IsNot Nothing Then _cacheAttFilename.TryAdd(mail.EntryID, result)  ' 2026/6/27 by simon/Claude: 雖然讀取跳過快取, 強制讀COM, 但是讀回來的值還是應該要回填快取, 更新到最正確的內容
        If _iLikeNoisy Then _dbg(" ├ 結束", mail.FolderPath & "\" & mail.Subject & " | 成果: " & If(result IsNot Nothing, result.Count.ToString(), "Nothing"))
        Return result

    End Function
    Private Function GetMailBody(entryID As String, Optional folderPath As String = "", Optional skipCache As Boolean = False) As String
        ' ---------------------------------------------------------------
        ' GetMailBody — Layer2.5 快取代理：Body 快取存取點
        ' 2026/04/28 by Simon/Claude: 依照 L2.5 架構抽出快取邏輯，L3 只剩純 COM
        '   ① 快取命中（_cacheMailBody）→ 直接回傳，0 COM call
        '   ② 快取未命中 → 呼叫 L3 GetMailBodyOOM 讀取並正規化
        '   ③ 成功才存快取（真空信存 ""；讀取失敗回 Nothing 不存，避免把失敗誤當真空信永久卡住，見 2026/07/06 異動）
        ' ---------------------------------------------------------------
        ' 2026/06/23 by Simon/Claude Opus 4.8: 依照 L2.5 架構重構 GetMailBody，讀取分派邏輯:
        '   ① 快取命中（_cacheMailBody）→ 直接回傳，0 COM call
        '   ② (預留 SSD 層: body 目前不落 SQLite;待純文字總容量評估通過後,於此插入 lazy load,形狀對齊 GetAttFilename)
        '   ③ 讀取分派: _rdo2 在且有 folderPath → GetMailBodyRdo(store-scoped 高速);否則 → GetMailBodyOOM(OOM)。RDO 已上移至 L2.5。
        '      skipCache=True: 跳過 ①讀與寫快取(build pass 掃數萬封避免撐爆 _cacheMailBody),仍走 RDO/OOM 分派。
        ' ---------------------------------------------------------------
        If _iLikeNoisy Then _dbg(" ├ 開始", folderPath & " | ID: " & entryID) ' by Claude Sonnet 4.6, 2026/07/01
        Dim cached As String = Nothing
        If Not skipCache AndAlso _cacheMailBody.TryGetValue(entryID, cached) Then Return cached   ' ①

        Dim body As String
        If _rdo2 IsNot Nothing AndAlso folderPath <> "" Then        ' ③ RDO 優先
            body = GetMailBodyRdo(entryID, folderPath)
            If body Is Nothing Then body = GetMailBodyOOM(entryID)  ' RDO 解析失敗保底
        Else
            body = GetMailBodyOOM(entryID)                          ' OOM
        End If

        ' 2026/07/06 by Simon/Claude: 原本無條件存(含失敗存"")會把讀取失敗當成真空信永久卡住;
        '   改成失敗(Nothing)不快取,讓使用者下次點開時重試。真空信仍會存"" (NormalizeMailBody 賦值過)
        If Not skipCache AndAlso body IsNot Nothing Then _cacheMailBody(entryID) = body
        If _iLikeNoisy Then _dbg(" ├ 結束", folderPath & " | ID: " & entryID & " | 成果長度: " & If(body IsNot Nothing, body.Length.ToString(), "0")) ' by Claude Sonnet 4.6, 2026/07/01
        Return body

    End Function
    Private Structure MailPreviewDetails
        ' 2026/07/07 by Simon/Claude: ShowMailQuickPreview 一次性讀取的顯示資料 (不進任何快取，僅使用者主動預覽時讀)
        Dim Html As String
        Dim SenderEmail As String
        Dim ToRecipients As String                      ' mail.To 的顯示名清單 (分號分隔)
        Dim Attachments As List(Of (Name As String, Size As Long))  ' 附件檔名+大小(bytes)，供判斷該不該刪信
    End Structure
    Private Function GetMailPreviewDetails(entryID As String, Optional folderPath As String = "") As MailPreviewDetails
        ' ---------------------------------------------------------------
        ' GetMailPreviewDetails — 快速預覽用一次性讀取 .HTMLBody + 寄件者信箱 + 收件者 + 附件(檔名/大小)，
        '   仿 GetMailBody 的 RDO/OOM 分派，但不進 _cacheMailBody(該快取存的是 NormalizeMailBody 過的純文字，格式不同不能共用)
        '   2026/07/07 by Simon/Claude: 供 ShowMailQuickPreview(HTMLBody + WebView2) 使用，只在使用者主動點開預覽時呼叫，非批次掃描路徑，故不做快取；
        '     所有欄位從同一個 RDO/OOM 物件一次讀完，避免分次呼叫重複建物件
        ' ---------------------------------------------------------------
        Dim result As New MailPreviewDetails With {.Attachments = New List(Of (Name As String, Size As Long))}

        If _rdo2 IsNot Nothing AndAlso folderPath <> "" Then
            Dim store = GetRdoStore(folderPath)
            If store IsNot Nothing Then
                Dim rm As Redemption.RDOMail = Nothing
                Try
                    rm = TryCast(store.GetMessageFromID(entryID), Redemption.RDOMail)
                    If rm IsNot Nothing Then
                        result.Html = rm.HTMLBody
                        result.SenderEmail = rm.SenderEmailAddress
                        result.ToRecipients = rm.To
                        For Each att As Redemption.RDOAttachment In rm.Attachments
                            result.Attachments.Add((att.FileName, CLng(att.Size)))
                        Next
                        Return result
                    End If
                Catch ex As System.Exception
                    _dbg("錯誤", $"GetMailPreviewDetails(RDO) 失敗 (ID: {entryID}): {ex.Message}")
                Finally
                    Dim o As Object = rm : TryMarshalRelease(o)
                End Try
            End If
        End If

        ' OOM fallback
        Dim nSpace As Outlook.NameSpace = Nothing
        Dim olItem As Object = Nothing
        Try
            nSpace = _olApp.GetNamespace("MAPI")
            olItem = nSpace.GetItemFromID(entryID)
            Dim mail = TryCast(olItem, Outlook.MailItem)
            If mail IsNot Nothing Then
                result.Html = mail.HTMLBody
                result.SenderEmail = mail.SenderEmailAddress
                result.ToRecipients = mail.To
                For Each att As Outlook.Attachment In mail.Attachments
                    result.Attachments.Add((att.FileName, CLng(att.Size)))
                Next
            End If
        Catch ex As System.Exception
            _dbg("錯誤", $"GetMailPreviewDetails(OOM) 失敗 (ID: {entryID}): {ex.Message}")
        Finally
            TryMarshalRelease(olItem)
            TryMarshalRelease(nSpace)
        End Try
        Return result
    End Function
    Private Async Function GetMailInfo(folder As Folder, needTopic As Boolean, cToken As CancellationToken, Optional fPath As String = "") As Task(Of List(Of (Mail As MailItemInfo, Topic As String)))
        ' ---------------------------------------------------------------
        ' GetMailInfo — Layer2.5 快取存取點 (Tab7)
        ' 2026/05/06 by Claude: cache key 改為純 fPath（移除 |needTopic 後綴）
        ' 2026/05/11 by Claude Sonnet 4.6: 改用 L2.5 快取，記憶體命中時 0 COM；配合刪除後主動 invalidate _cacheMailCount 確保不污染
        ' 2026/05/12 by Simon/Claude: ① 記憶體命中邏輯重構
        '   - _cacheMailCount 有值 → 用來驗 snap，避免 COM → Return entry.Mails
        '   - _cacheMailCount 無值 → 直接信任 _cacheMailInfo，不打 COM → Return entry.Mails
        '   - 信任依賴：刪除後 InvalidateMailCache 會同時清兩個快取，確保不回傳鬼魂資料
        ' 2026/06/29 by Simon/Claude: 新作了GetMailInfo 免-folder 多載版本後, 原本這個版本只剩 Tab7/OST使用(因為難驗證), Tab4/Tab5 都轉換至免-folder 版本
        ' ---------------------------------------------------------------
        fPath = SafeGetPath(folder, fPath)
        'Dim cacheKey As String = fPath                                 ' 2026/05/06 by Claude: 純路徑，Tab4/Tab5/Tab7 共用
        'Dim fName As String = ExtractFolderName(fPath)                 ' 2026/05/11 by simon: 這個fName好像沒用到? 先保留未來可能用於除錯輸出
        'Dim currentSnap As Long = PeekLiveFolderSnapOOM(sFolder, fPath) ' 2026/05/11 by Claude Sonnet 4.6: 改用 L2.5 快取，記憶體命中時 0 COM；配合刪除後主動 invalidate _cacheMailCount 確保不污染

        ' ① 記憶體命中檢查
        Dim entry As (Mails As List(Of (Mail As MailItemInfo, Topic As String)), Snap As Long) = Nothing
        If _cacheMailInfo.TryGetValue(fPath, entry) Then Return entry.Mails

        ' ② DB lazy load (mail_info 存在的話)
        ' 這裡不管 needTopic 是 True 還是 False，只要 DB 有最新資料，我們都拿來用
        Dim currentSnap As Long = GetMailCount(folder, fPath)  ' L2.5，memory > DB > COM
        Dim dbResult = LazyGetMailInfo(fPath)
        If dbResult.HasValue AndAlso dbResult.Value.Snap = currentSnap Then
            Dim mails = dbResult.Value.Mails
            _cacheMailInfo(fPath) = (mails, currentSnap)
            Return mails
        End If

        ' ③ Fallback: Layer3 COM 掃描
        Dim resultList = Await GetMailInfoOOM(folder, needTopic, cToken, fPath)

        ' 掃描完畢，存入記憶體快取 (SaveCache 時會持久化到 SSD)
        If resultList IsNot Nothing Then _cacheMailInfo(fPath) = (resultList, currentSnap) : MarkMailFolderDirty(fPath)   ' 2026/07/03 by Simon/Claude: dirty 追蹤

        Return resultList
    End Function
    Private Async Function GetMailInfo(fPath As String, eid As String, sid As String, needTopic As Boolean, cToken As CancellationToken) As Task(Of List(Of (Mail As MailItemInfo, Topic As String)))
        ' ---------------------------------------------------------------
        ' GetMailInfo(免-folder 多載) — Layer2.5 快取存取點，給純資料 tuple 消費者用 (Tab4/Tab5)
        ' 2026/06/29 by Simon/Claude [Stage2]: ① 記憶體(by fPath) → ② DB lazy(by fPath) → ③ Layer3 COM。
        '   熱路徑①②完全不碰 COM 物化;folder-based 版保留給 Tab7/OST 等手握 folder 的呼叫者。
        ' 2026/07/02 by Simon/Claude [Task 2b]: ③ Layer3 COM 內部再分兩段 —
        '   RDO(GetMailInfoRdo) 是③這一步的主要讀取手段(免物化 folder,parity 已驗證通過);
        '   OOM(GetFolderById 物化 + GetMailInfoOOM) 才是 RDO 讀不到時的 fallback,比照 GetSubtree 既有寫法(_rdo2 IsNot Nothing 為閘)。
        ' ---------------------------------------------------------------
        ' ① 記憶體命中
        Dim entry As (Mails As List(Of (Mail As MailItemInfo, Topic As String)), Snap As Long) = Nothing
        If _cacheMailInfo.TryGetValue(fPath, entry) Then Return entry.Mails

        ' ② DB lazy load (mail_info)
        Dim currentSnap As Long = GetMailCount(fPath, eid, sid)  ' 免-folder，memory > DB > COM
        Dim dbResult = LazyGetMailInfo(fPath)
        If dbResult.HasValue AndAlso dbResult.Value.Snap = currentSnap Then
            Dim mails = dbResult.Value.Mails
            _cacheMailInfo(fPath) = (mails, currentSnap)
            Return mails
        End If

        ' ③ Layer3 COM 讀取 — RDO 為主要手段
        If _rdo2 IsNot Nothing Then
            Dim rdoResult = GetMailInfoRdo(fPath, eid, sid)
            If rdoResult IsNot Nothing Then
                If _iLikeNoisy Then _dbg("GetMailInfo ③", $"RDO 命中 {ExtractFolderName(fPath)}({rdoResult.Count} 封)")   ' 2026/07/02 by Claude [Task2b 驗證用,暫留]: 確認實際走哪條路徑
                Dim converted = rdoResult.Select(Function(m) (m, GetCleanSubject(m.Subject))).ToList()
                _cacheMailInfo(fPath) = (converted, currentSnap) : MarkMailFolderDirty(fPath)   ' 2026/07/03 by Simon/Claude: dirty 追蹤
                Return converted
            End If
        End If

        ' ③ Fallback: RDO 讀不到才 GetFolderById 物化 → OOM 掃描
        If _iLikeNoisy Then _dbg("GetMailInfo ③", $"OOM fallback {ExtractFolderName(fPath)}(_rdo2={If(_rdo2 Is Nothing, "Nothing", "有")})")   ' 2026/07/02 by Claude [Task2b 驗證用,暫留]
        Dim folder As Folder = GetFolderById(eid, sid)
        Dim resultList = Await GetMailInfoOOM(folder, needTopic, cToken, fPath)
        If resultList IsNot Nothing Then _cacheMailInfo(fPath) = (resultList, currentSnap) : MarkMailFolderDirty(fPath)   ' 2026/07/03 by Simon/Claude: dirty 追蹤
        Return resultList
    End Function
    Private Function RefreshMailInfo(ByRef info As MailItemInfo) As RefreshResult
        ' ---------------------------------------------------------------
        ' RefreshMailInfo — Layer2.5 快取代理層: 依 EntryID 從 COM 重讀單封郵件實體資訊，寫回 info (ByRef，MailItemInfo 為 Structure)
        '   跳過所有 cache (_cacheXXX / SSD)，直接打 COM。RDO→OOM fallback。只讀 Subject/Size/RcvTime/SenderName。
        '   回傳 Updated/NotFound/TransientError；失效郵件的移除政策由呼叫端決定 (目前一律保留+記錄)
        ' 2026/06/14 by Simon/Claude Opus 4.8
        ' 2026/07/02 by Simon/Claude: 原名 RefreshMailInfoL3，拆成 RefreshMailInfoRdo(Layer3 RDO)/RefreshMailInfoOOM(Layer3 OOM)
        '   兩段純資料層後，此函式只剩分派，改名去掉 L3 尾綴、移至 Layer2.5 區塊，比照 GetMailInfo(免-folder版)
        '   ③ RDO優先/OOM兜底 的切分方式(此函式沒有 memory/DB 快取層，只有 RDO/OOM 分派這一層)。
        ' 2026/07/04 by Simon/Claude: 移除 readAttachCount 參數 — 附件個數探針證實 MAPI 無訊息層級批次欄位可讀、
        '   逐封枚舉又是 Tab3 篩選路徑用不到的死碼(Tab3 實際讀 GetAttFilename().Count)，全體F5/右鍵刷新統一不再讀附件數。
        ' ---------------------------------------------------------------
        If String.IsNullOrEmpty(info.EntryID) Then Return RefreshResult.NotFound

        ' ⓪ RDO 優先: store-scoped on _rdo2,繞過 OOM 開信的記憶體開銷;RefreshMailInfoRdo 回 False 才落 ① OOM fallback。
        If _rdo2 IsNot Nothing AndAlso info.FolderPath <> "" Then
            If RefreshMailInfoRdo(info) Then Return RefreshResult.Updated
        End If

        ' ① Fallback: RDO 讀不到才走 OOM
        Return RefreshMailInfoOOM(info)
    End Function
    Private Async Function PreLoadMailCacheAsync(folderList As List(Of (eid As String, sid As String, fPath As String)), cToken As CancellationToken) As Task
        ' ---------------------------------------------------------------
        ' PreLoadMailCacheAsync — SSD 批次預熱（優化B）
        ' 對尚未在記憶體的路徑發出一次 SQL IN 批次查詢，填入 _cacheMailInfo。
        ' 之後主迴圈的 GetMailInfo 全部命中 memory，不再逐個打 DB。
        ' 不做 snap 驗證：信任快取，失效由刪除後的 InvalidateMailCacheForPaths 負責。
        ' 2026/05/11 by Simon/Claude: 優化B
        ' 2026/06/28 by Simon/Claude [Stage2]: folderList 型別改 (eid,sid,fPath),內部僅用 .fPath,無其他變動。
        ' ---------------------------------------------------------------
        Dim missedPaths = folderList.Where(Function(f) Not _cacheMailInfo.ContainsKey(f.fPath)) _
                                    .Select(Function(f) f.fPath).ToList()
        If missedPaths.Count = 0 Then Return

        _dbg(" ├ 開始", $"DB 批次查詢 {missedPaths.Count} 個未命中路徑")
        Dim dbBatch = Await Task.Run(Function() LoadMailInfoCore(missedPaths), cToken)   ' 2026/07/07 by Simon/Claude: 原 LazyGetMailInfoBatch 改 Overloads 同名(List 簽名)

        For Each kvp In dbBatch
            _cacheMailInfo.TryAdd(kvp.Key, (kvp.Value.Mails, kvp.Value.Snap))
        Next

        _dbg(" ├ 結束", $"預熱完成，填入 {dbBatch.Count} 個資料夾")
        Await Task.Yield()
    End Function
    Private Async Function GetSubtree(rootFolder As Folder, includeSubF As Boolean, Optional progress As IProgress(Of ProgressReport) = Nothing, Optional skipCache As Boolean = False, Optional cToken As CancellationToken = Nothing) As Task(Of List(Of (eid As String, sid As String, fPath As String)))
        ' ---------------------------------------------------------------
        ' GetSubtree — 整棵子資料夾清單 (Layer2.5 快取代理)
        ' 2026/04/17 by Claude: 從 GetSubtree 拆出快取邏輯
        '   原來的快取邏輯混在 BFS 函數裡，現在統一到此 L2.5 層
        '   GetSubtreeOOM (原 GetSubtree) 只剩純 BFS COM 掃描
        ' 呼叫順序: ① 記憶體命中 → ② DB lazy load → ③ Layer3 GetSubtreeOOM
        ' includeSubF=False 時無需快取，直接呼叫 L3 回傳單節點清單
        ' 2026/06/13 by Simon/Claude Opus 4.8: 子樹計數鏈「去模式化」重構 — 鍵不再含 _showAllFolders (原: rootPath & "|" & _showAllFolders)
        '   (1) 去模式化快取鍵: 只存一份完整骨架(含非郵件夾)，鍵不再含 _showAllFolders；模式過濾移至計數層(FilterSubtreeByMode)。
        '   (2) 完整性檢查: DB lazy 撈「未過濾全集」後，必須通過 IsSubtreeComplete(骨架完整) 才採用，否則 fallback L3 完整重掃，
        '       消滅原本「folder_info 殘缺 → LIKE 默默回傳殘缺子樹 → 子樹靜默少算」的未爆彈。
        '   (3) skipCache=True: F5/discover 跳過記憶體+DB 快取，直打 L3 完整重掃 + 回填(補完整性檢查抓不到 Outlook 新增夾的 staleness 缺口)。
        '   (3) skipCache=True: F5/discover 跳過記憶體+DB 快取讀取，重算並覆寫快取(補完整性檢查抓不到 Outlook 新增夾的 staleness 缺口)。(2026/6/25)
        ' 2026/06/28 by Simon/Claude [Stage2]: 回傳合約改 (eid,sid,fPath) 純資料 tuple。② DB lazy 不再 GetFolderFromID 物化(直接用 row.eid/sid);
        '   靜態快取 _cacheSubTreeList 從此不握 COM 物件。RDO/OOM 兩條 L3 都已回傳同型別資料 tuple。
        ' ---------------------------------------------------------------
        Dim rootPath As String = SafeGetPath(rootFolder)
        If _iLikeNoisy Then _dbg(" ├ 開始", $"{rootPath} | includeSubF: {includeSubF} | skipCache: {skipCache}") ' by Gemini 3.5 Flash, 2026/07/01")
        If Not includeSubF Then Return Await GetSubtreeOOM(rootFolder, False, progress, cToken:=cToken)         ' 單節點不快取

        If Not skipCache Then
            Dim cachedList As List(Of (eid As String, sid As String, fPath As String)) = Nothing
            If _cacheSubTreeList.TryGetValue(rootPath, cachedList) Then                 ' ① 記憶體命中 (完整骨架)
                _dbg(" ├ 結束", $"{ExtractFolderName(rootPath)} (Cache Hit) | 資料夾總計: {cachedList.Count}")
                Return cachedList
            End If

            ' ② DB lazy load: 利用 LIKE 一次取回整棵樹的 ID 並重建物件
            ' 注意: DB 存放的是 (EntryID, StoreID, FolderPath)，我們在這裡重建 Tuple
            ' 2026/06/13 by Simon/Claude Opus 4.8: 一律撈未過濾全集(isIncludeAll:=True)，再用 IsSubtreeComplete 驗證骨架完整性
            Dim dbRows = LazyGetSubFolderIDAsList(rootPath, isIncludeAll:=True)             ' ② DB lazy load (完整全集)
            If dbRows IsNot Nothing AndAlso IsSubtreeComplete(dbRows, rootPath) Then
                ' 預分配容量為 512，優化從 DB 載入資料夾子樹時的處理速度 (by Gemini 3 Flash, 2026/05/04)
                Dim dbResults As New List(Of (eid As String, sid As String, fPath As String))(512)
                For Each row In dbRows
                    ' LazyGetSubFolderIDAsList 回傳的是 (eid, sid, path, isMail, hasCh, fc) 的具名列表 by Gemini 3.0 flash, 2026/04/16
                    ' 2026/06/28 [Stage2]: 不再 GetFolderFromID 物化,直接用 DB 的 eid/sid 組資料 tuple
                    ' 2026/07/04 by Simon/Claude Fable 5 [rootless 骨架未爆彈]: root 列 DB 身分證可能為空(entry_id NULL 豁免放行)
                    '   → 以 live rootFolder 補上,確保骨架必含帶身分證的 root(FilterSubtreeByMode 從 root 起走,root 缺席=整棵樹全滅)。
                    Dim eid As String = row.eid, sid As String = row.sid
                    If row.path = rootPath AndAlso eid = "" Then
                        Try : eid = rootFolder.EntryID : sid = rootFolder.StoreID : Catch : End Try
                    End If
                    dbResults.Add((eid, sid, row.path))
                    ' 2026/06/13 by Simon/Claude Opus 4.8: 回填身分證(is_mail)與 fc，供計數層 FilterSubtreeByMode 的 is_mail 過濾使用
                    If eid <> "" Then _cacheFolderIDs.TryAdd(row.path, (eid, sid, row.isMail <> 0, row.hasCh <> 0))   ' 2026/07/04: 空身分證不入快取,避免污染
                    If row.fc >= 0 Then _cacheFolderCount(row.path) = row.fc
                Next
                If dbResults.Count > 0 Then
                    _cacheSubTreeList(rootPath) = dbResults
                    If _iLikeNoisy Then _dbg("    ├ SSD Hit (Tree)", $"{ExtractFolderName(rootPath)}: 已從資料庫載入完整骨架 {dbResults.Count} 個子目錄")
                    Return dbResults
                End If
            ElseIf dbRows IsNot Nothing Then
                If _iLikeNoisy Then _dbg("    ├ DB 殘缺", $"{ExtractFolderName(rootPath)}: folder_info 子樹不完整 → fallback L3 完整重掃")
            End If
        End If

        ' 🆕 2026/06/25 by Simon/Claude: RDO 快速探索派工(上移自 L3)。閘 = _rdo2 在(= CheckRDO 勾)。
        '   skipCache=True 時仍走此處重算並覆寫快取;GetSubtreeRdo 回 Nothing(RDO 不可用/失敗)才掉回 ③ 純 OOM L3。
        If _rdo2 IsNot Nothing Then
            Dim rdoResult = GetSubtreeRdo(rootFolder, rootPath, progress)
            If rdoResult IsNot Nothing Then
                If Not cToken.IsCancellationRequested AndAlso rdoResult.Count > 0 Then _cacheSubTreeList(rootPath) = rdoResult
                Return rdoResult
            End If
        End If

        ' ③ Layer3 BFS COM 完整掃描；OCE re-throw，快取寫入在 L3 完成後由 L3 自行負責 (純 OOM fallback)
        Return Await GetSubtreeOOM(rootFolder, True, progress, cToken:=cToken)

    End Function

    Private Function GetSortedSubFolders(pFolder As Folder, Optional fPath As String = "", Optional skipCache As Boolean = False) As List(Of (f As Folder, name As String, fPath As String))
        ' ==========================================
        ' 取得引數pFolder下的所有subFolders並排序後傳回
        ' 優化紀錄: 2026/03/29 by Gemini 3.1 Pro
        '   1. 加入 Layer3 過濾: 只保留郵件目錄 (olMailItem)，排除行事曆/聯絡人等
        '   2. 單次屬性讀取: 先快取 Name 後排序，避開 LINQ 重複打 COM 的 N log N 效能陷阱
        ' 2026/04/15: 支援傳入 fPath，並透過字串拼接子路徑，減少內部 COM 呼叫
        ' 2026/04/16: 整合記憶體與 SSD 快取機制 (修復損補遺失) by Gemini 3.0 flash
        ' 2026/07/05 by Simon/Claude: 去模式化 (呼應 06/13 folder_info 子樹計數鏈同一原則，當時列為範圍外)。
        '   快取鍵拿掉 |_showAllFolders，只存一份完整骨架(含非郵件夾)；模式剪枝從骨架層移到回傳前的顯示層，
        '   根除「全顯/過濾模式各存一份快取、互相看不到對方已掃過的結果」的殘缺風險。
        ' 2026/07/10 by Simon/Claude Fable 5: 回傳型別升級為 (f, name, fPath) tuple — 枚舉/DB 兩條路徑本來就握有 Name 與路徑，
        '   舊版回傳 List(Of Folder) 把這兩個值丟掉，呼叫端 (LoadSubFolderToTreeView 等) 建節點又逐一重打 COM 讀回來；
        '   現在一起帶出去，呼叫端零 COM。快取 _cacheFolderTree 同步改存 tuple。
        ' 2026/07/10 by Simon/Claude [效能決策紀錄 — 決定不轉 RDO，別再重新研究]:
        '   ① 實測: 單層列舉頂多 5~20 個子資料夾、最多 15~20ms，且只在節點第一次展開時觸發一次，體感為零。
        '   ② RDO 每次呼叫要先解 store/folder 的固定解析開銷，N=5~20 時吃掉所有收益 (同 HasSubFoldersFast 2026/7/2 的結論)；
        '      RDO 批次優勢要到幾百幾千筆才顯現 (如 GetSubtree 骨架批次、Tab3 178k 列場景)。
        '   ③ 就算轉了也省不掉 OOM: 呼叫端 LoadSubFolderToTreeView 需要 Folder 物件塞 node.Tag，
        '      RDO 枚舉完仍要逐一 GetFolderFromID 物化回 OOM，兩頭成本都要付。結案不做。
        ' ==========================================
        fPath = SafeGetPath(pFolder, fPath)
        Dim fName As String = ExtractFolderName(fPath)
        If _iLikeNoisy Then _dbg(" ├ 開始", fName)

        ' ① 記憶體快取檢查: 命中則回傳完整骨架，模式過濾在回傳前處理，0ms
        Dim cacheKey As String = fPath
        Dim cachedFolders As List(Of (f As Folder, name As String, fPath As String)) = Nothing
        If Not skipCache AndAlso _cacheFolderTree.TryGetValue(cacheKey, cachedFolders) Then Return FilterSubFoldersByMode(cachedFolders)    ' 2026/6/27, Gemini 發現skipCache參數沒有穿透到這裡的記憶體快取

        ' ② SSD / DB 讀取分支 (Lazy Load): TreeView 展開時的主要加速點
        ' ✅ 2026/5/31 by Gemini/Simon: 加入 skipCache 引數判斷是否要強制讀取COM，避免在需要最新資料的情況下誤用過期快取
        If _dbCache IsNot Nothing AndAlso Not skipCache Then
            Dim dbIDs = LazyGetOrderedSubFolderIDs(fPath, isIncludeAll:=True)     ' 2026/07/05: 永遠撈未過濾全集(完整骨架)，模式過濾移到下面回傳前
            If dbIDs IsNot Nothing Then
                ' 預分配容量為 512，足以涵蓋多數資料夾搜尋結果，減少陣列頻繁 Resize 開銷 (by Gemini 3 Flash, 2026/05/04)
                Dim dbResults As New List(Of Folder)(512)
                For Each row In dbIDs
                    Try
                        ' LazyGetSubFolderIDAsList 回傳的是 (eid, sid, path) 的具名 Tuple 列表 by Gemini 3.0 flash, 2026/04/16
                        Dim f = TryCast(_olNS.GetFolderFromID(row.eid, row.sid), Folder)
                        If f IsNot Nothing Then
                            dbResults.Add(f)
                            ' 2026/07/05: row 已帶 isMail，順手回填身分證(零額外 COM)，讓 FilterSubFoldersByMode 走 0 COM 快取命中
                            Try : _cacheFolderIDs.TryAdd(row.path, (row.eid, row.sid, CBool(row.isMail), CBool(row.hasCh))) : Catch : End Try
                        End If
                    Catch : End Try
                Next
                If dbResults.Count > 0 Then
                    _cacheFolderTree(cacheKey) = dbResults
                    If _iLikeNoisy Then _dbg("    ├ SSD Hit", $"{fName}: 已從資料庫載入 {dbResults.Count} 個子目錄")
                    Return FilterSubFoldersByMode(dbResults, fPath)
                End If
            End If
        End If

        ' ③ 傳統 OOM 分支 (Fallback): 快取未命中時才打 COM 掃描 (收集子資料夾並快取屬性 (減少 COM 屬性重複呼叫)
        Dim infoList As New List(Of FolderSortInfo)(512)
        Dim subs As Folders = pFolder.Folders
        Try
            For Each subF As Folder In subs
                ' 2026/07/05: 移除模式剪枝 — 永遠加入完整骨架(含非郵件夾)，剪枝移到回傳前 (見上方 07/05 說明)
                Dim sName As String = subF.Name
                Dim childPath As String = fPath & "\" & sName
                Dim isMail As Boolean = IsMailFolder(subF, childPath)   ' 2026/06/14 by Simon/Claude Opus 4.8: isMail 改為先算一次，供身分證註冊共用 (骨架本就完整，不分模式皆需算出 is_mail 以正確登記)
                infoList.Add(New FolderSortInfo With {.FolderObj = subF, .Name = sName, .HasChinese = TextHasChineseChar(sName)})
                ' 這裡 subF 被加入 infoList 成為物件清單，所以不能在這裡 TryRelease 它

                ' 2026/6/2: 再次修正F5 強制刷新的總數讀取不正確: 🔽🔽🔽 【修復點 2】順手把展開的資料夾也註冊身分證 🔽🔽🔽
                ' 2026/06/14 by Simon/Claude Opus 4.8: 還原此修復點 (原被註解)。GetSortedSubFolders 是「樹載入」與「BuildBfsFolderTree 計算 BFS」
                '   共用的子夾枚舉樞紐，兩條路徑原本都不寫 _cacheFolderIDs → 子夾存檔時 entry_id/is_mail 為 NULL，
                '   重啟後被 LazyGetOrderedSubFolderIDs (entry_id IS NOT NULL / is_mail=1) 濾掉 → 樹崩 (只剩收件匣)。
                '   在此 TryAdd 身分證即可一次修好兩條路徑；isMail 重用上方已算好的值，不重複打 COM。
                Try : _cacheFolderIDs.TryAdd(childPath, (subF.EntryID, subF.StoreID, isMail, TextHasChineseChar(sName))) : Catch : End Try
            Next
        Finally
            TryMarshalRelease(subs) ' 存到變數後可以TryRelease subs，避免後續 COM 呼叫時 RCW 已釋放的例外
        End Try

        ' 純記憶體排序: 完全不觸發 COM 呼叫
        ' 2026/4/7 進一步優化 by Gemini: 加入 StringComparer.OrdinalIgnoreCase 略過語系分析，爆發性提速
        Dim sortedFolders = infoList.OrderBy(Function(i) If(i.HasChinese, 1, 0)).ThenBy(Function(i) i.Name, StringComparer.OrdinalIgnoreCase).Select(Function(i) i.FolderObj).ToList()
        _cacheFolderTree(cacheKey) = sortedFolders
        If _iLikeNoisy Then _dbg(" ├ 結束", $"{fName} (BFS) | 子資料夾數: {sortedFolders.Count}")
        Return FilterSubFoldersByMode(sortedFolders, fPath)

    End Function
    Private Function FilterSubFoldersByMode(folders As List(Of Folder), parentPath As String) As List(Of Folder)
        ' 2026/07/05 by Simon/Claude: GetSortedSubFolders 去模式化後的顯示層剪枝 — 完整骨架依 _showAllFolders 即時派生。
        '   IsMailFolder 優先查 _cacheFolderIDs(骨架建立時已隨手回填 is_mail)，正常情況 0 COM；只有極少數快取遺漏才退回 1 次 COM 讀取。
        If _showAllFolders Then Return folders
        Return folders.Where(Function(f) IsMailFolder(f, parentPath & "\" & f.Name)).ToList()
    End Function
    Private Function HasSubFoldersFast(cFolder As Folder, Optional fPath As String = "") As Boolean
        ' ---------------------------------------------------------------
        ' HasSubFoldersFast — 光速版子資料夾加號預測 (專為 TreeView 展開設計)
        ' 2026/4/7 by Gemini, 解決 SSD 讀出後無法出現假節點 + 號，以及嚴重卡頓問題
        ' 2026/7/2 by simon, 直接保底呼叫一次 cFolder.Folders.Count > 0
        '   (比 PR_CONTENT_COUNT 還快，也比RDO 要先解 store/folder 才能讀屬性的解析開銷快，所以也不需轉到_rdo2)
        ' ---------------------------------------------------------------
        '   呼叫順序：① _cacheFolderCount 記憶體 → ② LazyGetFolderInfo(fPath).fc → ③ pFolder.Folders.Count COM
        '   已在 LoadSubFolderToTreeView 第 489 行啟用， 解決 DB 載入後 TreeView 不顯示 "+" 的問題
        '   比直接 pFolder.Folders.Count 快： 記憶體命中~0μs，DB命中~0.1ms，COM才~1-5ms
        ' ---------------------------------------------------------------
        fPath = SafeGetPath(cFolder, fPath)
        If fPath = "" Then Return False     ' 2026/04/23 by Gemini 3.0 flash: 確保路徑有效，抓不到代表資料夾異常

        Dim fc As Long
        If _cacheFolderCount.TryGetValue(fPath, fc) Then Return fc > 0

        Dim row = LazyGetFolderInfo(fPath)
        If row IsNot Nothing AndAlso row.fc >= 0 Then
            _cacheFolderCount.TryAdd(fPath, row.fc) ' 把確認的值送回記憶體快取
            Return row.fc > 0
        End If

        ' 萬一都沒有，直接保底呼叫一次 COM (比 PR_CONTENT_COUNT 驗證還快)
        Try : Return cFolder.Folders.Count > 0 : Catch : Return False : End Try
    End Function
    Private Function IsMailFolder(folder As Folder, Optional fPath As String = "") As Boolean
        ' 2026/07/03 by Simon/Claude: 併掉獨立的 _cacheIsMailFolder，改查 _cacheFolderIDs.isMail —
        '   RDO/OOM 展開骨架時(GetSubtreeRdoBatch 的 PR_CONTAINER_CLASS / GetSubtreeOOM 的 IsMailFolder)已經回填這裡，
        '   查無時才落 COM DefaultItemType，並把結果一併補進同一份字典供全體共用(不再各自為政)。
        ' 2026/07/08 by Simon/Claude: 補 PR_CONTAINER_CLASS(經 CcIsMail 判定，與 RDO 路徑共用同一份規則) + PR_ATTR_HIDDEN 雙重判定 ——
        '   交談動作設定/快速步驟設定/提醒 等 Outlook 內建隱藏設定夾，DefaultItemType 對它們一律回傳 olMailItem
        '   (Outlook 本身的已知模糊行為)，原邏輯誤判為郵件夾。
        '   實測(PROBE_ISMAILFIX)發現 PR_ATTR_HIDDEN 對「提醒」夾讀取會拋例外(該夾不支援此屬性讀取)，單靠隱藏判定不夠；
        '   但三夾的 PR_CONTAINER_CLASS 分別為 'IPF.Configuration'/'IPF.Configuration'/'Outlook.Reminder'，
        '   CcIsMail 對這些值本就正確回傳 False(不符 IPF.Note/Post/Imap 開頭)，故以 CcIsMail 為主要判定，
        '   PR_ATTR_HIDDEN 為輔助防線(讀不到時不影響判定，維持原保守預設)。
        fPath = SafeGetPath(folder, fPath)
        Dim fName As String = ExtractFolderName(fPath)
        If _iLikeNoisy Then _dbg(" ├ 開始", fName)

        Dim info As (eid As String, sid As String, isMail As Boolean, hasCh As Boolean) = Nothing
        If _cacheFolderIDs.TryGetValue(fPath, info) Then Return info.isMail

        Static allowedTypes As Outlook.OlItemType() = {Outlook.OlItemType.olMailItem,
                                                       Outlook.OlItemType.olPostItem}
        Try
            Dim isHidden As Boolean = False
            Try : isHidden = CBool(folder.PropertyAccessor.GetProperty(PR_ATTR_HIDDEN)) : Catch : End Try  ' 讀不到視為未隱藏，維持原保守預設
            Dim cc As String = ""
            Try : cc = CStr(folder.PropertyAccessor.GetProperty(PR_CONTAINER_CLASS)) : Catch : End Try     ' 讀不到視為空字串，CcIsMail 空→mail 保守預設
            Dim itemType As Outlook.OlItemType = folder.DefaultItemType
            Dim isMail As Boolean = Not isHidden AndAlso CcIsMail(cc) AndAlso allowedTypes.Contains(itemType)
            Try : _cacheFolderIDs.TryAdd(fPath, (folder.EntryID, folder.StoreID, isMail, TextHasChineseChar(fName))) : Catch : End Try
            If _iLikeNoisy AndAlso Not isMail Then _dbg("過濾非郵件資料夾", $"{fName} ({itemType}, cc={cc}, hidden={isHidden})") ' 只有非郵件時才記錄
            Return isMail
        Catch
            Return False
        End Try
    End Function

#End Region
#Region "  ├ Layer3 RDO 直接存取底層"
    Private Function GetMailCountRdo(folderPath As String, eid As String, sid As String) As Long
        ' ---------------------------------------------------------------
        ' 本層郵件數 RDO 讀取層。store-scoped on _rdo2,與 GetMailCountOOM 原 ⓪ tier 同邏輯(rdoFolder.Items.Count)。
        '   解析失敗(store/folder 找不到/例外)回 -1,由 L2.5 proxy 判 <0 fallback 到 GetMailCountOOM。
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
        ' 本層直屬子資料夾數 RDO 讀取層。store-scoped on _rdo2,與 GetFolderCountOOM 原 ⓪ tier 同邏輯(rdoFolder.Folders.Count)。
        '   解析失敗回 -1,由 L2.5 proxy 判 <0 fallback 到 GetFolderCountOOM。
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
    Private Function GetFolderSizeRdo(folderPath As String, eid As String, sid As String) As Long
        ' ---------------------------------------------------------------
        ' 本層郵件大小加總 RDO 讀取層。store-scoped on _rdo2,讀 rdoFolder.Items.MAPITable 單欄 PR_MESSAGE_SIZE(PT_LONG) 批次 GetRows 加總。
        '   ⚠ 不可用 PR_MESSAGE_SIZE_EXTENDED(PT_I8,0x0E080014): 探針 SpikeFolderSizeReadCompare 實證,PT_I8 經 MAPITable.GetRows
        '     每封都回相同垃圾常數 ≈ -2^31(8-byte marshaling 壞)。改讀 PR_MESSAGE_SIZE(PT_LONG,0x0E080003): 單封 size < 2GB 必夠,
        '     加總進 Long(64-bit) 不溢位(實證 2009 夾 6.4GB 總和正確)。與 OOM GetArray 對拍 parity 全一致,速度快 3~10×。
        '   解析失敗(store/folder 找不到/例外)回 -1,由 L2.5 proxy 判 <0 fallback 到 GetFolderSizeOOM;空夾(RowCount=0)回 0。
        '   sid 目前未用(store-scoped 單參數即可解),保留參數對稱。
        ' 2026/06/27 by Simon/Claude Opus 4.8
        ' ---------------------------------------------------------------
        Dim store As Redemption.RDOStore = GetRdoStore(folderPath)
        If store Is Nothing Then Return -1

        Dim rdoFolder As Redemption.RDOFolder = Nothing
        Dim items As Object = Nothing, tbl As Object = Nothing
        Try
            rdoFolder = TryCast(store.GetFolderFromID(eid), Redemption.RDOFolder)
            If rdoFolder Is Nothing Then Return -1
            items = rdoFolder.Items
            tbl = items.MAPITable
            If CInt(tbl.RowCount) = 0 Then Return 0   ' 空夾: 0 bytes (不走 GetRows,防空表邊界)

            tbl.Columns = PR_MESSAGE_SIZE
            tbl.GoToFirst()
            Dim total As Long = 0
            Do
                Dim chunk As Array = TryCast(tbl.GetRows(5000), Array)
                If chunk Is Nothing Then Exit Do
                Dim got As Integer = 0
                For i As Integer = chunk.GetLowerBound(0) To chunk.GetUpperBound(0)
                    got += 1
                    Dim row As Array = TryCast(chunk.GetValue(i), Array)
                    If row Is Nothing Then Continue For
                    Dim value = row.GetValue(row.GetLowerBound(0))
                    If value IsNot Nothing AndAlso Not IsDBNull(value) Then total += CLng(value)
                Next
                If got < 5000 Then Exit Do   ' 最後一批不足 → 到底
            Loop
            Return total
        Catch ex As System.Exception
            If _iLikeNoisy Then _dbg("GetFolderSizeRdo 失敗", $"{ExtractFolderName(folderPath)} | {ex.Message}")
            Return -1
        Finally
            TryMarshalRelease(tbl) : TryMarshalRelease(items)
            Dim o As Object = rdoFolder : TryMarshalRelease(o)
        End Try
    End Function
    Private Function GetAttFilenameRdo(ByRef mail As MailItemInfo) As List(Of String)
        ' ---------------------------------------------------------------
        ' 附件檔名 RDO 讀取層。store-scoped on _rdo2,與 GetAttFilenameOOM 原 ⓪ tier 同邏輯(att.Type=1)。
        '   解析失敗(store 找不到/例外)回 Nothing,由 L2.5 fallback 到 GetAttFilenameOOM。
        ' 2026/06/23: L3 原用 New List(4096) 預配置(每封配 ~32KB,熱路徑浪費)此處改以 mail.AttCount 精準預配置(上界,絕不 realloc,亦不浪費)。
        ' ---------------------------------------------------------------
        Dim store As Redemption.RDOStore = GetRdoStore(mail.FolderPath)
        If store Is Nothing Then Return Nothing

        Dim result As New List(Of String)(mail.AttCount)
        Dim rdoMsg As Redemption.RDOMail = Nothing
        Try
            rdoMsg = TryCast(store.GetMessageFromID(mail.EntryID), Redemption.RDOMail)
            If rdoMsg Is Nothing Then Return Nothing

            For i As Integer = 1 To rdoMsg.Attachments.Count
                Dim att As Redemption.RDOAttachment = rdoMsg.Attachments.Item(i)
                Try : If att.Type = 1 Then result.Add(att.FileName)   ' 僅 olByValue(1),與 GetAttFilenameOOM 一致
                Finally : Dim o As Object = att : TryMarshalRelease(o)
                End Try
            Next
            Return result
        Catch ex As System.Exception
            If _iLikeNoisy Then _dbg("失敗", ex.Message)
            Return Nothing
        Finally
            Dim o As Object = rdoMsg : TryMarshalRelease(o)
        End Try
    End Function
    Private Function GetMailBodyRdo(entryID As String, folderPath As String) As String
        ' ---------------------------------------------------------------
        ' mailbody RDO 讀取層。store-scoped on _rdo2,讀 .Body 後套同一支 NormalizeMailBody(與 GetMailBodyOOM 一致)。
        '   RDOMail.Body 不分 Mail/Post 型別;解析失敗回 Nothing,由 L2.5 fallback 到 GetMailBodyOOM。
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
    Private Function GetMailInfoRdo(fPath As String, eid As String, sid As String) As List(Of MailItemInfo)
        ' ---------------------------------------------------------------
        ' 整夾基本郵件資訊 RDO 讀取層。store-scoped on _rdo2,MAPITable 7 欄批次 GetRows(5000) chunk。
        '   2026/07/02 by Simon/Claude [Task 2b]: 已接上 GetMailInfo(免-folder版) 與 GetMailInfoAsDict 的③RDO優先路徑。
        '   純資料掃描:不算 Topic、不轉容器,轉換責任交給呼叫端(GetMailInfo 轉 List(Mail,Topic)、GetMailInfoAsDict 轉 Dictionary)。
        '   ⚠ EntryID 為 Byte(),須經 RdoTableEidToHex 轉字串(比照 GetSubtreeRdoBatch)。
        '   ⚠ Subject/ReceivedTime/SenderName 改用明確 proptag(非具名屬性,與 OOM 版讀法不同)。
        '   2026/07/02 by Claude [PROBE_BASICINFO_RDO 驗證後修正]: 字串 proptag 字尾原本誤用 001E(PT_STRING8,ANSI codepage),
        '     CJK 字元會變 "?" 或被最佳近似置換成形似字(如全形［］被換成〔〕),改成 001F(PT_UNICODE)才對。
        '   解析失敗或中途例外 → 回 Nothing,丟棄已累積結果(不可回傳掃一半);空夾 → 回空 List。
        ' 2026/07/02 by Simon/Claude
        ' 2026/07/04 by Simon/Claude Fable 5 [PROBE_ATTACHBATCH 驗證後上線]: COLS 多掛 PR_HASATTACH + SmartNoAttach 兩旗標欄,
        '   同一次 GetRows 順手回填 Tab3 _cacheAttMailList(比照 GetSubtreeRdoBatch 順手回填 _cacheMailCount 的先例)。
        '   迴紋針語意 = PR_HASATTACH=True 且 SmartNoAttach≠True(探針 5 store/270夾/178,000列 parity 100%,兩欄邊際成本實測≈0ms)。
        '   全列讀完才寫快取(中途例外回 Nothing 不寫),snap 用本次讀到的 RowCount,並 MarkMailFolderDirty 讓 SaveCache 落 SSD。
        ' ---------------------------------------------------------------
        Const COLS As String = "EntryID, " & PR_SUBJECT & ", " & PR_MESSAGE_SIZE & ", " & PR_MESSAGE_DELIVERY_TIME & ", " & PR_SENDER_NAME & ", " & PR_INTERNET_MESSAGE_ID_W & ", " & PR_SENDER_EMAIL_ADDRESS_W & ", " & PR_HASATTACH & ", " & DASL_SMARTNOATTACH

        Dim store As Redemption.RDOStore = GetRdoStore(fPath)
        If store Is Nothing Then Return Nothing

        Dim rdoFolder As Redemption.RDOFolder = Nothing
        Dim items As Object = Nothing, tbl As Object = Nothing
        Try
            rdoFolder = TryCast(store.GetFolderFromID(eid), Redemption.RDOFolder)
            If rdoFolder Is Nothing Then Return Nothing
            items = rdoFolder.Items
            tbl = items.MAPITable
            Dim rowTotal As Integer = CInt(tbl.RowCount)
            If rowTotal = 0 Then Return New List(Of MailItemInfo)   ' 空夾: 合法空結果,非失敗

            tbl.Columns = COLS : tbl.GoToFirst()
            Dim result As New List(Of MailItemInfo)(rowTotal)
            Dim attList As New List(Of MailItemInfo)(256)        ' 2026/07/04 piggyback: 迴紋針候選,掃完順手餵 _cacheAttMailList
            Do
                Dim chunk As Array = TryCast(tbl.GetRows(5000), Array)
                If chunk Is Nothing Then Exit Do
                Dim got As Integer = 0
                For i As Integer = chunk.GetLowerBound(0) To chunk.GetUpperBound(0)
                    got += 1
                    Dim row As Array = TryCast(chunk.GetValue(i), Array)
                    If row Is Nothing Then Continue For
                    Dim lb As Integer = row.GetLowerBound(0)
                    Dim entryID As String = RdoTableEidToHex(row.GetValue(lb))
                    If entryID = "" Then Continue For
                    Dim msgIdRaw As String = TryCast(row.GetValue(lb + 5), String)
                    ' 2026/07/02 by Claude [PROBE_BASICINFO_RDO 驗證後修正]: MAPITable 的 PT_SYSTIME 回傳 UTC,
                    '   OOM .ReceivedTime 回傳系統本地時區,parity 驗出固定 -8 小時差(台灣 UTC+8),須轉回本地時間才對齊。
                    Dim rcvTime As DateTime = DateTime.MinValue
                    If TypeOf row.GetValue(lb + 3) Is DateTime Then
                        rcvTime = DateTime.SpecifyKind(CDate(row.GetValue(lb + 3)), DateTimeKind.Utc).ToLocalTime()
                    End If
                    Dim info As New MailItemInfo With {.EntryID = entryID,
                                                       .Subject = If(TryCast(row.GetValue(lb + 1), String), ""),
                                                       .Size = If(row.GetValue(lb + 2) IsNot Nothing, Convert.ToInt64(row.GetValue(lb + 2)), 0L),
                                                       .RcvTime = rcvTime,
                                                       .SenderName = If(TryCast(row.GetValue(lb + 4), String), ""),
                                                       .FolderPath = fPath,
                                                       .MsgIDhash = StringToXxHash64Hex(If(msgIdRaw, "")),
                                                       .SenderEmail = If(TryCast(row.GetValue(lb + 6), String), "")}
                    result.Add(info)
                    ' 2026/07/04 piggyback: 迴紋針判定(SmartNoAttach 未設定時回 Int32 錯誤碼,只認 Boolean=True 為剔除)
                    Dim vHas As Object = row.GetValue(lb + 7)
                    Dim vSmart As Object = row.GetValue(lb + 8)
                    If (TypeOf vHas Is Boolean AndAlso CBool(vHas)) AndAlso Not (TypeOf vSmart Is Boolean AndAlso CBool(vSmart)) Then attList.Add(info)
                Next
                If got < 5000 Then Exit Do
            Loop
            ' 2026/07/04 piggyback: 全列讀取成功才順手回填 Tab3 快取(Tab4/5/7 掃過的夾,Tab3 直接 ① 記憶體命中)
            _cacheAttMailList(fPath) = New FolderCacheTab3 With {.AttMailList = attList, .ItemCountSnap = CLng(rowTotal)}
            MarkMailFolderDirty(fPath)
            Return result
        Catch ex As System.Exception
            If _iLikeNoisy Then _dbg("GetMailInfoRdo 失敗", $"{ExtractFolderName(fPath)} | {ex.Message}")
            Return Nothing
        Finally
            TryMarshalRelease(tbl) : TryMarshalRelease(items)
            Dim o As Object = rdoFolder : TryMarshalRelease(o)
        End Try
    End Function
    Private Function GetAttMailListRdo(fPath As String, eid As String, sid As String) As List(Of MailItemInfo)
        ' ---------------------------------------------------------------
        ' GetAttMailListRdo — Tab3 Phase1 附件候選清單 RDO 讀取層。store-scoped on _rdo2,MAPITable 7 欄批次 GetRows(5000) chunk。
        '   迴紋針語意批次判定: PR_HASATTACH(0x0E1B000B)=True 且 PidLidSmartNoAttach≠True。
        '   SmartNoAttach = PSETID_Common {00062008-0000-0000-C000-000000000046}/dispid 0x8514/PT_BOOLEAN named prop,
        '   MS 官方語意「訊息沒有使用者可見附件」;未設定時 GetRows 回 Int32 錯誤碼、有設定時回 Boolean →
        '   只認 Boolean=True 為剔除,對齊文件「may be unset; default FALSE」。
        ' 2026/07/04 by Simon/Claude Fable 5 [PROBE_ATTACHBATCH 驗證後上線]:
        '   個人 profile 5 store/270夾/178,000列/25,829封,與 OOM @SQL hasattachment 對拍 parity 100%(多0漏0),
        '   暖頁面快 1.8x~13.2x(RDO 每列 ~7µs 恆定;OOM 每夾 ~15ms GetTable 固定成本,夾多且小時差距放大)。
        '   此結果推翻 memory_20260623_2210「Tab3 不做 RDO 版本」結論 — 該輪缺口是沒試過 SmartNoAttach 當第二欄。
        '   ⚠ 工作 profile 問題 store(寄件備份2013~2018,olOLE 樣本)尚未補測;本函式回 Nothing 即 fallback OOM,語意保底不破。
        '   解析失敗/例外 → 回 Nothing(呼叫端 fallback OOM);空夾 → 回空 List。
        ' ---------------------------------------------------------------
        Const COLS As String = "EntryID, " & PR_SUBJECT & ", " & PR_MESSAGE_SIZE & ", " & PR_MESSAGE_DELIVERY_TIME & ", " & PR_SENDER_NAME & ", " & PR_HASATTACH & ", " & DASL_SMARTNOATTACH

        Dim store As Redemption.RDOStore = GetRdoStore(fPath)
        If store Is Nothing Then Return Nothing

        Dim rdoFolder As Redemption.RDOFolder = Nothing
        Dim items As Object = Nothing, tbl As Object = Nothing
        Try
            rdoFolder = TryCast(store.GetFolderFromID(eid), Redemption.RDOFolder)
            If rdoFolder Is Nothing Then Return Nothing
            items = rdoFolder.Items
            tbl = items.MAPITable
            If CInt(tbl.RowCount) = 0 Then Return New List(Of MailItemInfo)   ' 空夾: 合法空結果,非失敗

            tbl.Columns = COLS : tbl.GoToFirst()
            Dim result As New List(Of MailItemInfo)(256)
            Do
                Dim chunk As Array = TryCast(tbl.GetRows(5000), Array)
                If chunk Is Nothing Then Exit Do
                Dim got As Integer = 0
                For i As Integer = chunk.GetLowerBound(0) To chunk.GetUpperBound(0)
                    got += 1
                    Dim row As Array = TryCast(chunk.GetValue(i), Array)
                    If row Is Nothing Then Continue For
                    Dim lb As Integer = row.GetLowerBound(0)
                    Dim entryID As String = RdoTableEidToHex(row.GetValue(lb))
                    If entryID = "" Then Continue For
                    ' 迴紋針判定: HasAtt=True 且 SmartNoAttach≠Boolean True,非候選連 MailItemInfo 都不建
                    Dim vHas As Object = row.GetValue(lb + 5)
                    Dim vSmart As Object = row.GetValue(lb + 6)
                    If Not (TypeOf vHas Is Boolean AndAlso CBool(vHas)) Then Continue For
                    If TypeOf vSmart Is Boolean AndAlso CBool(vSmart) Then Continue For
                    Dim rcvTime As DateTime = DateTime.MinValue                    ' PT_SYSTIME 為 UTC → 轉本地(比照 GetMailInfoRdo)
                    If TypeOf row.GetValue(lb + 3) Is DateTime Then
                        rcvTime = DateTime.SpecifyKind(CDate(row.GetValue(lb + 3)), DateTimeKind.Utc).ToLocalTime()
                    End If
                    result.Add(New MailItemInfo With {
                        .EntryID = entryID,
                        .Subject = If(TryCast(row.GetValue(lb + 1), String), ""),
                        .Size = If(row.GetValue(lb + 2) IsNot Nothing, Convert.ToInt64(row.GetValue(lb + 2)), 0L),
                        .RcvTime = rcvTime,
                        .SenderName = If(TryCast(row.GetValue(lb + 4), String), ""),
                        .FolderPath = fPath})
                Next
                If got < 5000 Then Exit Do
            Loop
            Return result
        Catch ex As System.Exception
            If _iLikeNoisy Then _dbg("失敗", $"{ExtractFolderName(fPath)} | {ex.Message}")
            Return Nothing
        Finally
            TryMarshalRelease(tbl) : TryMarshalRelease(items)
            Dim o As Object = rdoFolder : TryMarshalRelease(o)
        End Try
    End Function
    Private Function RdoDateLiteral(d As Date) As String
        ' ExecSQL 只吃 ISO 'yyyy-MM-dd HH:mm:ss' 字面值(InvariantCulture,避開 zh-TW 上午/下午問題);
        '   8 種候選格式已在 PROBE_YEARSQL 探針全部實測過(見 memory_20260702_1409),只有帶時間的 ISO 格式可解析且不會把 23:59:59 邊界截斷成 00:00:00。
        Return "'" & d.ToString("yyyy-MM-dd HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture) & "'"
    End Function
    Private Function GetYearCountRdo(fPath As String, eid As String, sid As String) As ConcurrentDictionary(Of Integer, Integer)
        ' ---------------------------------------------------------------
        ' 單一資料夾年份郵件分佈 RDO 讀取層。store-scoped on _rdo2,ExecSQL COUNT(*) 逐年查詢。
        '   ① 範圍偵測: TOP 1 ReceivedTime ... ORDER BY ASC/DESC 各查一次,找出真實 min/max 年份(不借用其他來源答案)。
        '   ② 逐年 COUNT(*),日期字面值固定用格式 B('yyyy-MM-dd HH:mm:ss'),不在 production 重跑格式測試(探索階段才需要)。
        '   探針驗證(2026/07/02 PROBE_YEARSQL): 範圍偵測 100% 準確(兩子樹 55/55、6/6 全數相符);
        '     加計範圍偵測固定成本後仍比純 OOM 快 1.3x~4x(非早期樂觀估計的 5.5x,那是借用 OOM 答案的簡化版)。
        '   解析失敗(store/folder 找不到/例外)回 Nothing,由 L2.5 proxy fallback 到 GetYearCountOOM;空夾回空字典(非失敗)。
        ' 2026/07/02 by Simon/Claude
        ' 2026/07/03 by Simon 註解: YearCount/MonthCount多線程平行化: 技術上可行但不值得, 在17~30萬封郵件的全庫重建也都不到二秒, 再砍也只快不到一秒反而增加多線程同步複雜度, 先不做。
        ' ---------------------------------------------------------------
        Dim store As Redemption.RDOStore = GetRdoStore(fPath)
        If store Is Nothing Then Return Nothing

        Dim rdoFolder As Redemption.RDOFolder = Nothing
        Dim items As Object = Nothing, tbl As Object = Nothing
        Try
            rdoFolder = TryCast(store.GetFolderFromID(eid), Redemption.RDOFolder)
            If rdoFolder Is Nothing Then Return Nothing
            items = rdoFolder.Items
            tbl = items.MAPITable

            ' ── ① 範圍偵測 ──
            Dim minDate As Date? = Nothing, maxDate As Date? = Nothing
            Dim rsMin As Object = tbl.ExecSQL("SELECT TOP 1 ReceivedTime FROM Folder ORDER BY ReceivedTime ASC")
            If rsMin IsNot Nothing AndAlso Not CBool(rsMin.EOF) Then minDate = CDate(rsMin.Fields(0).Value)
            If minDate Is Nothing Then Return New ConcurrentDictionary(Of Integer, Integer)()   ' 空夾: 無信,合法空結果,非失敗

            Dim rsMax As Object = tbl.ExecSQL("SELECT TOP 1 ReceivedTime FROM Folder ORDER BY ReceivedTime DESC")
            If rsMax IsNot Nothing AndAlso Not CBool(rsMax.EOF) Then maxDate = CDate(rsMax.Fields(0).Value)
            If maxDate Is Nothing Then Return Nothing   ' min 有 max 沒有,資料異常,交給 OOM 保底

            ' ── ② 逐年 COUNT(*) ──
            Dim result As New ConcurrentDictionary(Of Integer, Integer)
            For y As Integer = minDate.Value.Year To maxDate.Value.Year
                Dim lit1 As String = RdoDateLiteral(New Date(y, 1, 1, 0, 0, 0))
                Dim lit2 As String = RdoDateLiteral(New Date(y, 12, 31, 23, 59, 59))
                Dim rs As Object = tbl.ExecSQL($"SELECT COUNT(*) FROM Folder WHERE ReceivedTime >= {lit1} AND ReceivedTime <= {lit2}")
                Dim cnt As Integer = If(rs IsNot Nothing AndAlso Not CBool(rs.EOF), CInt(rs.Fields(0).Value), 0)
                If cnt > 0 Then result(y) = cnt
            Next
            Return result
        Catch ex As System.Exception
            If _iLikeNoisy Then _dbg("GetYearCountRdo 失敗", $"{ExtractFolderName(fPath)} | {ex.Message}")
            Return Nothing
        Finally
            TryMarshalRelease(tbl) : TryMarshalRelease(items)
            Dim o As Object = rdoFolder : TryMarshalRelease(o)
        End Try
    End Function
    Private Function GetMonthCountRdo(fPath As String, eid As String, sid As String, year As Integer) As ConcurrentDictionary(Of Integer, Integer)
        ' ---------------------------------------------------------------
        ' 單一資料夾指定年份月份分佈 RDO 讀取層。store-scoped on _rdo2,逐月 ExecSQL COUNT(*)(固定12次,年份已知不需範圍偵測)。
        '   架構比照 GetYearCountRdo(2026/07/02 由 Simon 拍板: 年份架構已驗證,月份不用另外開探針測,直接套用)。
        '   解析失敗回 Nothing,由 L2.5 proxy fallback 到 GetMonthCountOOM。
        ' 2026/07/02 by Simon/Claude
        ' 2026/07/03 by Simon 註解: YearCount/MonthCount多線程平行化: 技術上可行但不值得, 在17~30萬封郵件的全庫重建也都不到二秒, 再砍也只快不到一秒反而增加多線程同步複雜度, 先不做。
        ' ---------------------------------------------------------------
        Dim store As Redemption.RDOStore = GetRdoStore(fPath)
        If store Is Nothing Then Return Nothing

        Dim rdoFolder As Redemption.RDOFolder = Nothing
        Dim items As Object = Nothing, tbl As Object = Nothing
        Try
            rdoFolder = TryCast(store.GetFolderFromID(eid), Redemption.RDOFolder)
            If rdoFolder Is Nothing Then Return Nothing
            items = rdoFolder.Items
            tbl = items.MAPITable

            Dim result As New ConcurrentDictionary(Of Integer, Integer)
            For m As Integer = 1 To 12
                Dim startDate As New Date(year, m, 1, 0, 0, 0)
                Dim endDate As Date = startDate.AddMonths(1).AddSeconds(-1)
                Dim rs As Object = tbl.ExecSQL($"SELECT COUNT(*) FROM Folder WHERE ReceivedTime >= {RdoDateLiteral(startDate)} AND ReceivedTime <= {RdoDateLiteral(endDate)}")
                Dim cnt As Integer = If(rs IsNot Nothing AndAlso Not CBool(rs.EOF), CInt(rs.Fields(0).Value), 0)
                If cnt > 0 Then result(m) = cnt
            Next
            Return result
        Catch ex As System.Exception
            If _iLikeNoisy Then _dbg("GetMonthCountRdo 失敗", $"{ExtractFolderName(fPath)} | {year}年 | {ex.Message}")
            Return Nothing
        Finally
            TryMarshalRelease(tbl) : TryMarshalRelease(items)
            Dim o As Object = rdoFolder : TryMarshalRelease(o)
        End Try
    End Function
    Private Function RefreshMailInfoRdo(ByRef info As MailItemInfo) As Boolean
        ' ---------------------------------------------------------------
        ' RefreshMailInfoRdo — Layer3 RDO 讀取層: store-scoped on _rdo2,依 EntryID 重讀單封郵件實體資訊寫回 info。
        '   2026/06/28 by Simon/Claude: _rdo → _rdo2 store-scoped。路徑對齊 GetMailBodyRdo/GetAttFilenameRdo:
        '     由 info.FolderPath 經 GetRdoStore 解 store,再 store.GetMessageFromID(EntryID)。
        '     (沿用舊 _rdo 區塊行為: RDO 路徑不更新 info.FolderPath;移動夾的路徑更新仍由 OOM 負責)
        '   成功(找到郵件並寫回 info) → True;找不到 store/郵件或任何例外 → False,由呼叫端落 OOM fallback。
        ' 2026/07/02 by Simon/Claude: 從舊 RefreshMailInfoL3 拆出,比照 GetMailInfoRdo/GetMailInfoOOM 切分方式;
        '   分派層改名 RefreshMailInfo 並移至 Layer2.5 區塊,OOM 對應層 RefreshMailInfoOOM 移至 Layer3 OOM 區塊。
        ' 2026/07/04 by Simon/Claude: 移除 readAttachCount 分支與附件枚舉 — 探針證實批次無法讀、逐封枚舉又是 Tab3 篩選用不到的死碼。
        ' ---------------------------------------------------------------
        Dim store As Redemption.RDOStore = GetRdoStore(info.FolderPath)
        If store Is Nothing Then Return False

        Dim rdoMsg As Redemption.RDOMail = Nothing
        Try
            rdoMsg = TryCast(store.GetMessageFromID(info.EntryID), Redemption.RDOMail)
            If rdoMsg Is Nothing Then Return False

            info.Subject = rdoMsg.Subject
            info.Size = rdoMsg.Size
            info.RcvTime = rdoMsg.ReceivedTime
            info.SenderName = rdoMsg.SenderName
            Return True
        Catch ex As System.Exception   ' RDO 任何失敗都讓 OOM 再試一次，由 OOM 作最終結論
            If _iLikeNoisy Then _dbg("    ├ RDO 失敗，走 OOM fallback", ex.Message)
            Return False
        Finally
            Dim o As Object = rdoMsg : TryMarshalRelease(o)
        End Try
    End Function

    ' 🆕 2026/06/24 by Simon/Claude Opus 4.8: GetSubtreeRdo — RDO 批次階層走訪(探針 C 實證版)
    '   兩階段拆乾淨、分開計時:
    '     Phase1 探索(純 RDO): Folders.MAPITable.GetRows 整層批次,只對 PR_SUBFOLDERS=true 遞迴(跳葉夾);
    '            同步寫 _cacheFolderCount(每夾 fc = 該層 RowCount,葉夾=0)。探針已證正確 + ~8ms。
    '     Phase2 還原(純 OOM): 逐夾 _olNS.GetFolderFromID(eid,sid) 還原 Outlook.Folder,呼叫既有 IsMailFolder,
    '            註冊 _cacheFolderIDs,組出與 BFS 同合約 (Folder,fPath) 清單。← 正確性對齊保留,Step2 再拆掉。
    '   任一步解析不到 → 回 Nothing,由 GetSubtreeOOM 掉回 OOM BFS(絕不產錯結果)。
    '   ※ 唯一未驗假設: Phase2 用 _rdo2 table eid 餵 _olNS(跨 session)。production 跑成功即驗證,失敗則安全網接住。
    Private Function GetRdoStore(folderPath As String) As Redemption.RDOStore
        ' 由 FolderPath 取得 _rdo2 上對應的 RDOStore(store-scoped resolve 用)
        '   首次呼叫一次掃 _rdo2.Stores 建 byName 表;之後 byPath 記憶化命中為 O(1) 零解析。找不到回 Nothing(呼叫端 fallback OOM)。
        '   ※ phase 1 單執行緒(UI 緒)存取,用一般 Dictionary;日後平行讀取(phase C)再評估執行緒安全。
        If _iLikeNoisy Then _dbg(" ├ 開始", folderPath)
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
                _dbg(" ├ 建表失敗", ex.Message)
            End Try
        End If

        Dim store As Redemption.RDOStore = Nothing
        If _rdo2StoreByPath.TryGetValue(folderPath, store) Then Return store    ' 記憶化命中(含 Nothing)
        _rdo2StoreByName.TryGetValue(GetStoreNameFromPath(folderPath), store)
        _rdo2StoreByPath(folderPath) = store                                    ' 含 Nothing 亦記,避免重複解析
        Return store

    End Function
    Private Function BuildRdoStoreByNameDict(session As Redemption.RDOSession) As Dictionary(Of String, Redemption.RDOStore)
        ' 供多執行緒平行讀取使用：每個獨立 RDOSession(非共用 _rdo2)各自建立一份 store 顯示名→RDOStore 對照表。
        '   不可共用 _rdo2StoreByName/_rdo2StoreByPath，那兩張表綁定共用 _rdo2、僅供 UI 緒單一存取(見 GetRdoStore 開頭註解)。
        ' 2026/07/03 by Simon/Claude [平行化SimHash]: 抽出自 PROBE_BODYPAR 探針驗證過的邏輯，供 production 平行版 worker 共用
        Dim d As New Dictionary(Of String, Redemption.RDOStore)()
        For i As Integer = 1 To session.Stores.Count
            Dim st As Redemption.RDOStore = session.Stores.Item(i)
            Dim nm As String = st.Name
            If Not String.IsNullOrEmpty(nm) AndAlso Not d.ContainsKey(nm) Then
                d(nm) = st
            Else
                Dim o As Object = st : TryMarshalRelease(o)
            End If
        Next
        Return d
    End Function
    Private Function GetSubtreeRdo(rootFolder As Folder, rootPath As String, Optional progress As IProgress(Of ProgressReport) = Nothing) As List(Of (eid As String, sid As String, fPath As String))
        ' 🆕 2026/06/25 by Simon/Claude: GetSubtreeRdo 派工殼 — Phase1 批次主支,失敗退 RDO 枚舉(3e);Phase2 共用 OOM 還原。
        '   rdoRoot 由本殼統一持有/釋放;兩支 helper 只釋放自己開的子夾,故批次失敗後枚舉可重用同一 rdoRoot。
        '   兩個 RDO 法都失敗 → 回 Nothing,由 L2.5 掉回純 OOM GetSubtreeOOM。
        ' 2026/06/28 by Simon/Claude [Stage2]: Phase2 去物化 — 不再 GetFolderFromID,直接用 RDO table eid 組 (eid,sid,fPath) 資料 tuple;
        '   isMail 改由 batch 的 CC 推導(nd.isMail,取代原 IsMailFolder(f))。這是「還原 ms」歸零的關鍵。
        If _iLikeNoisy Then _dbg("    ├ 開始", rootPath)
        Dim store As Redemption.RDOStore = GetRdoStore(rootPath)
        If store Is Nothing Then _dbg("    ├ RDO 略過", "GetRdo2Store=Nothing → OOM BFS") : Return Nothing
        Dim sid As String = "" : Try : sid = rootFolder.StoreID : Catch : Return Nothing : End Try
        Dim rdoRoot As Redemption.RDOFolder = Nothing
        Try : rdoRoot = TryCast(store.GetFolderFromID(rootFolder.EntryID), Redemption.RDOFolder) : Catch : End Try
        If rdoRoot Is Nothing Then _dbg("    ├ RDO 略過", "root 解析失敗 → OOM BFS") : Return Nothing

        Try
            ' ── Phase 1 探索: 批次主支,失敗退 RDO 枚舉 ──
            Dim nodes As List(Of (eid As String, name As String, path As String, isMail As Boolean)) = GetSubtreeRdoBatch(store, rdoRoot, rootPath)
            If nodes Is Nothing Then
                _dbg("    ├ RDO 退枚舉", "批次失敗 → 退簡單 RDO 枚舉")
                nodes = GetSubtreeRdoEnum(rdoRoot, rootPath)
            End If
            If nodes Is Nothing Then If _iLikeNoisy Then _dbg("    ├ ✗ RDO 批次+枚舉皆失敗 → 改走 OOM BFS") : Return Nothing

            ' ── Phase 2 組裝資料 tuple (Stage2: 不再物化 OOM;eid 用 RDO table 原生 eid,isMail 用 CC 推導) ──
            Dim result As New List(Of (eid As String, sid As String, fPath As String))(nodes.Count + 1)
            result.Add((rootFolder.EntryID, sid, rootPath))
            ' 2026/07/04 by Simon/Claude Fable 5 [rootless 骨架未爆彈]: root 也註冊 _cacheFolderIDs(原本只註冊子孫),
            '   讓 SaveCache 能把 root 的 entry_id 寫進 folder_info,DB 骨架從此自帶 root 身分證(治本);root 視為 mail 夾。
            _cacheFolderIDs.TryAdd(rootPath, (rootFolder.EntryID, sid, True, TextHasChineseChar(ExtractFolderName(rootPath))))
            For Each nd In nodes
                _cacheFolderIDs.TryAdd(nd.path, (nd.eid, sid, nd.isMail, TextHasChineseChar(nd.name)))
                result.Add((nd.eid, sid, nd.path))
            Next
            Return result
        Finally
            Dim oo As Object = rdoRoot : TryMarshalRelease(oo)   ' rdoRoot 由本殼統一釋放
            If _iLikeNoisy Then _dbg("    ├ 結束")
        End Try
    End Function
    Private Function GetSubtreeRdoBatch(store As Redemption.RDOStore, rdoRoot As Redemption.RDOFolder, rootPath As String) As List(Of (eid As String, name As String, path As String, isMail As Boolean))
        ' 🆕 探索主支: Folders.MAPITable 批次,只對 PR_SUBFOLDERS=true 遞迴(跳葉夾)。回 nodes 或 Nothing(交給枚舉)。
        '   不釋放 rdoRoot(外層殼負責),只釋放自己 GetFolderFromID 開出的子夾。
        ' 2026/06/28 by Simon/Claude [Stage2]: COLS 多撈 PR_CONTAINER_CLASS(0x3613001E),由 CcIsMail 推 isMail 帶進 node,
        '   讓 Phase2 不需物化 folder 也有 isMail。(若夾普遍無 CC 改 0x3613001F;Stage0 實證夾無CC僅 ~4.5% 且全是郵件夾,CcIsMail 空→mail 已涵蓋)
        ' 2026/07/03 by Simon/Claude [PROBE_HIERCNT 通過後上線]: COLS 多撈 PR_CONTENT_COUNT(0x36020003),同一次 GetRows 順手回填 _cacheMailCount。
        '   探針實證(13 個真實子樹、374+ 夾): mc 與 OOM PropertyAccessor 全一致、額外耗時在雜訊等級(單一 +202ms 冷啟動異常值二次重跑不再重現)。
        '   讀不到值(Nothing/轉型失敗)時不寫快取，留給既有 GetMailCount 的 DB lazy/COM fallback 處理，不寫入猜測值。
        ' 2026/07/08 by Simon/Claude: COLS 多撈 PR_ATTR_HIDDEN(0x10F4000B) — 交談動作設定/快速步驟設定/提醒 等隱藏系統夾
        '   PR_CONTAINER_CLASS 常為空、命中 CcIsMail「查無→保守視為 mail」，導致被誤判為郵件夾納入樹。隱藏夾一律排除，不受 CcIsMail 結果影響。
        If _iLikeNoisy Then _dbg(" ├ 開始", store.ToString & rdoRoot.ToString & rootPath.ToString)
        Dim nodes As New List(Of (eid As String, name As String, path As String, isMail As Boolean))(512)
        Dim toRel As New List(Of Object)()
        Const COLS As String = "Name, EntryID, " & PR_SUBFOLDERS & ", " & PR_CONTAINER_CLASS & ", " & PR_CONTENT_COUNT & ", " & PR_ATTR_HIDDEN
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
                        Dim cc As String = TryCast(row.GetValue(lb + 3), String)               ' PR_CONTAINER_CLASS → CcIsMail
                        Dim cp As String = cur.p & "\" & nm
                        Dim vHid As Object = row.GetValue(lb + 5)                              ' PR_ATTR_HIDDEN
                        Dim isHidden As Boolean = If(TypeOf vHid Is Boolean, CBool(vHid), False) ' 讀不到視為未隱藏，維持原保守預設
                        nodes.Add((eidHex, nm, cp, Not isHidden AndAlso CcIsMail(cc)))

                        Dim vCnt As Object = row.GetValue(lb + 4)                              ' PR_CONTENT_COUNT → _cacheMailCount
                        If vCnt IsNot Nothing Then
                            Try : _cacheMailCount(cp) = Convert.ToInt64(vCnt) : Catch : End Try
                        End If

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
    Private Function GetSubtreeRdoEnum(rdoRoot As Redemption.RDOFolder, rootPath As String) As List(Of (eid As String, name As String, path As String, isMail As Boolean))
        ' 🆕 探索 fallback: 簡單 RDO 枚舉(For Each Folders),逐夾讀 .Name/.EntryID。回 nodes 或 Nothing。
        '   不釋放 rdoRoot(外層殼負責),只釋放自己枚舉到的子夾。
        ' 2026/06/28 by Simon/Claude [Stage2]: node 加 isMail 欄;枚舉路徑無 CC,保守視為 mail(True),與 FilterSubtreeByMode「查無預設納入」一致。
        ' 2026/07/08 by Simon/Claude: 已知殘留缺口 — 此枚舉 fallback 未讀 PR_ATTR_HIDDEN,隱藏系統夾(交談動作設定/快速步驟設定/提醒)
        '   若走到這條路徑仍會被誤判為 mail。此路徑僅在 GetSubtreeRdoBatch 失敗時才會觸發(罕見)，故暫不補救；
        '   主線 IsMailFolder(OOM) 與 GetSubtreeRdoBatch(RDO 主批次) 已修正，見同日其他變更。
        If _iLikeNoisy Then _dbg(" ├ 開始", rdoRoot.ToString & rootPath.ToString)
        Dim nodes As New List(Of (eid As String, name As String, path As String, isMail As Boolean))(512)
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
                    nodes.Add((eid, nm, cp, True)) : childCount += 1   ' 枚舉 fallback 無 CC,保守視為 mail
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
    Private Function RdoTableEidToHex(value As Object) As String
        ' table 的 PR_ENTRYID 經 GetRows 回 byte array(探針實證 Byte[]),統一轉 hex 字串供 GetFolderFromID
        If value Is Nothing Then Return ""
        If TypeOf value Is String Then Return CStr(value)
        If TypeOf value Is Byte() Then Return BitConverter.ToString(DirectCast(value, Byte())).Replace("-", "")
        If TypeOf value Is Array Then
            Dim a As Array = DirectCast(value, Array)
            Dim sb As New System.Text.StringBuilder(a.Length * 2)
            For k As Integer = a.GetLowerBound(0) To a.GetUpperBound(0) : sb.Append(Convert.ToByte(a.GetValue(k)).ToString("X2")) : Next
            Return sb.ToString()
        End If
        Return ""
    End Function
#End Region
#Region "  ├ Layer3 OOM 直接存取底層"
    Private Function GetMailCountOOM(folder As Folder, Optional fPath As String = "") As Long
        ' --------------------------------------------------------------
        ' GetMailCountOOM: 只讀單一資料夾的本層郵件數 (不含子孫)
        ' Fallback 鏈:
        '   ⓪ Redemption : RDOFolder.Items.Count (可在非 STA 執行緒呼叫)
        '   ① MAPI : PR_CONTENT_COUNT (0x36020003) (最快快取屬性)
        '   ② OOM  : pFolder.Items.Count (會建立 Items 集合)
        '   ③ fail : Return -1
        ' --------------------------------------------------------------
        fPath = SafeGetPath(folder, fPath)
        Dim fName As String = ExtractFolderName(fPath)
        If _iLikeNoisy Then _dbg("開始", fName)

        ' ⓪ Redemption: RDOFolder.Items.Count
        ' ⓪ RDO 路徑已上移至 L2.5 GetMailCountRdo(store-scoped on _rdo2),L3 回歸純 OOM。 2026/06/23 by Simon/Claude

        ' ① MAPI: PR_CONTENT_COUNT (0x36020003)
        Try
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

        If _iLikeNoisy Then _dbg("    ├ 結束", $"FAIL: {fName}")
        Return -1

    End Function
    Private Function GetFolderCountOOM(sFolder As Folder, Optional fPath As String = "") As Long
        ' --------------------------------------------------------------
        ' GetFolderCountOOM: 讀取單一資料夾的本層直屬子資料夾數
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

    End Function
    Private Async Function GetMailCountAllOOM(rootFolder As Folder, Optional progress As IProgress(Of ProgressReport) = Nothing, Optional skipCache As Boolean = False, Optional cToken As CancellationToken = Nothing) As Task(Of Long)
        ' --------------------------------------------------------------
        ' GetMailCountAllOOM v3.6: 讀取某資料夾及其整棵子樹的郵件總數
        ' by Gemini, 2026/04/02: 升級為 IProgress(Of ProgressReport) 並加入 100ms 節流回報
        ' 2026/04/15 by Claude: 加入 cToken 參數，取代 _cancelRequested 旗標 (見函數內說明)
        '
        ' v3.0 變更說明 (2026-03-22):
        '   合併原 GetMailCountAllOOM + GetMailCountAllParallel 為單一函數，
        '   統一 fallback 鏈，呼叫端不再需要選擇要用哪個版本。
        '   GetMailCountAllParallel 可標記廢棄或直接刪除。
        '
        ' 設計說明:
        '   為何呼叫 GetMailCountOOM() 而非直接用 GetTable():
        '     PR_CONTENT_COUNT 是 Folder 物件上的已儲存屬性，Outlook 自動維護，讀取等於讀一個整數，一次 COM call 結束。
        '     GetTable() 會把資料夾內所有郵件 row 逐一回傳，只為了計數代價太高。GetTable 適合讀郵件內容 (大小、日期)，不適合純計數。
        '
        '   回傳型別 Long 而非 Integer:
        '     單一資料夾用 Integer 夠 (PR_CONTENT_COUNT 是 PT_LONG 32-bit)，
        '     但整棵子樹加總若有多個大資料夾，理論上可能超過 Integer.MaxValue (2,147,483,647)，用 Long 安全。
        '
        ' Fallback 鏈 (依速度由快到慢) :
        '   ⓪ Redemption    : rdoFolder.TotalItemCount
        '                     MAPI hierarchy table 的彙總屬性，一次屬性取得整棵子樹總數，完全不需 BFS 遍歷，_rdoSession 未就緒自動跳過此層
        '                     Redemption 可正確讀取 PST 上此屬性 (原生 OOM 的 PR_MESSAGE_SIZE_EXTENDED 在 PST 上無效)
        '   ① Task.WhenAll  : RDO 走平行 BFS:
        '                     BFS 展開後每個資料夾各建一個 Task.Run，全部 WhenAll 等待
        '                     Task.Run 內的 GetMailCountOOM(cFolder) 走 Redemption ⓪ 時是 free-threaded 安全的
        '                     若 GetMailCountOOM fallback 到 MAPI PropertyAccessor，仍有 STA 違規風險，需留意
        '   ② BFS 循序累加  : GetSubtree BFS 展開 + GetMailCountOOM(Layer3) 逐一加總
        '                     支援取消檢查和 onProgress 進度回報，平行路徑失敗時的安全 fallback
        '   ③ 遞迴fallback  : GetSubtree 本身失敗時 (極少見) 的最後保險
        '                     無法精確回報進度，但確保加總結果正確
        '   ④ Return -1     : 四層都失敗，由 Layer2 決定如何處理
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
        ' 2026/06/13 by Simon/Claude Opus 4.8: 更正 — 上述「死碼」註解已過時。CollectFolderInfoForceL3 (F5 強刷入口)
        '   已復用本函數 (Form1_MainTab12.vb)，故本函數現為「F5/skipCache 子樹計數」的活路徑，非死碼。
        ' 原始設計意圖:
        ' 	呼叫端 → GetMailCountAllAsync (L2.5) → GetMailCountAllOOM (L3)
        ' 	呼叫端 → GetFolderCountAllAsync (L2.5) → GetFolderCountAllOOM (L3)
        '
        ' 後來使用了BFS剪枝速度更快: Compute → BFS → SumUpSubTreeBottomUp → UpdateFolderInfoCache
        ' 	(自己計算並直接寫入 _cacheMailCountAll / _cacheFolderCountAll，完全繞過 AllAsync系列函數)
        ' ---------------------------------------------------------------
        Dim rName As String = rootFolder.Name ' by Gemini 3.1 Pro 2026/04/16: 避免重複 COM 呼叫
        If _iLikeNoisy Then _dbg("    ├ 開始", rName) ' by Gemini, 2026/04/10: Level 1

        ' ⓪ Redemption: TotalItemCount 直接回傳整棵子樹郵件總數
        '   一次 COM call 結束，不需要任何 BFS 遍歷或平行處理, 2026-03-22 新增
        ' 2026/3/24 by Gemini: ① 平行 BFS (RDO): 使用 GetSubtreeToListL3_Rdo 取得清單，以 Parallel.ForEach 搭配 Interlocked.Add 快速加總
        ' 2026/04/15 by Claude: 改用 ParallelOptions.CancellationToken 取代 _cancelRequested 旗標
        ' 2026/06/13 by Simon/Claude Opus 4.8: RDO 快速路徑以 _rdoFastPath 開關控管 (預設關)。
        '   原因見 _rdoFastPath 宣告處: TotalItemCount 含 OOM 看不到的隱藏夾且無法 is_mail 過濾，會與 OOM 不一致。
        ' 2026/06/25 by Claude: 移除 ⓪TotalItemCount / ①平行BFS 兩條 _rdoFastPath 死分支(恆 False,從未執行)。
        '   停用原因(枚舉到隱藏夾、與 OOM 不一致)見 _rdoFastPath 宣告處(L45);該問題已由 GetSubtreeRdo(IPM 樹根走訪)解決,

        ' ② BFS 循序累加: GetSubtree 展開 + GetMailCountOOM(Layer3) 逐一加總
        ' 2026/04/15 by Claude: _cancelRequested 取代為 SmartThrottle(swThrottle, cToken)
        '   cToken 取消時 Task.Delay(1,cToken) 拋 OCE → Catch OCE → Return -1
        '   同時移除舊的 i Mod 10 Await Task.Yield()，統一由 SmartThrottle 每 100ms 讓出一次
        Try
            ' 2026/04/17 by Claude: 改呼叫 GetSubtree (L2.5)，享有快取加速
            ' 2026/06/13 by Simon/Claude Opus 4.8: 取得完整骨架後依 _showAllFolders 在計數層過濾 (剪枝移到這裡)；skipCache 一路 thread
            ' 2026/06/25 by Claude Opus 4.8: forceRefresh引數改 skipCache
            ' 2026/06/28 by Simon/Claude [Stage2]: skeleton/targetFolderList 改 (eid,sid,fPath);逐夾改走免-folder GetMailCount 多載(RDO優先+OOM fallback),不再物化。
            Dim skeleton As List(Of (eid As String, sid As String, fPath As String)) = Await GetSubtree(rootFolder, includeSubF:=True, skipCache:=skipCache, cToken:=cToken)
            Dim targetFolderList As List(Of (eid As String, sid As String, fPath As String)) = FilterSubtreeByMode(skeleton, SafeGetPath(rootFolder))
            Dim grandTotal As Long = 0
            Dim swThrottle As Stopwatch = Stopwatch.StartNew()  ' by Gemini, 2026/04/02: 100ms 節流閥; refactored by Claude Sonnet 4.6, 2026/06/07

            For i As Integer = 0 To targetFolderList.Count - 1
                Dim count As Long = GetMailCount(targetFolderList(i).fPath, targetFolderList(i).eid, targetFolderList(i).sid, skipCache)
                ' GetMailCount 的所有 fallback 都失敗才會到這個 else，記錄但不中止整體加總
                If count >= 0 Then grandTotal += count Else If _iLikeNoisy Then _dbg("    ├ Get MailCountAll ② 略過失敗資料夾", ExtractFolderName(targetFolderList(i).fPath)) ' by Gemini, 2026/04/10

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
            ' 2026/04/15: SmartThrottle 或 GetSubtree 取消時拋 OCE，正常中斷
            If _iLikeNoisy Then _dbg("    ├ ② 已取消", $"{rName}") : Return -1
        Catch ex As System.Exception
            If _iLikeNoisy Then _dbg("    ├ ② 循序BFS失敗，走遞迴fallback", $"{rName} | {ex.Message}") ' by Gemini, 2026/04/10
        End Try

        ' ③ 遞迴 fallback: GetSubtree 本身失敗時的最後保險
        '   無法精確回報進度，但確保加總結果正確
        '   注意: 遞迴呼叫會重新進入本函數，⓪ Redemption 已失敗所以 _rdoSession 仍 Nothing 或故障
        '         ① ② 也已失敗，只會走到 ③ 再次遞迴——理論上 ③ 不會無限展開，因為每層只遞迴直屬子資料夾
        '        若 ③ 常被觸發，需回頭檢查 GetSubtree 失敗的根本原因
        ' 2026/06/13 by Simon/Claude Opus 4.8: ③ 為 ② 拋非取消例外時的極罕見保險，未套用 _showAllFolders 模式過濾 (會含非郵件夾)；
        '   若實務上發現 ③ 被觸發致關閉模式下數字偏大，再回頭補剪枝。
        Try
            Dim totalCount As Long = 0
            Dim count As Integer = GetMailCountOOM(rootFolder)     ' 本層 mailcount
            If count >= 0 Then totalCount += count
            Await Task.Yield()
            ' 優化第六點：提取 Folders 集合並在 Finally 顯式釋放，防止遞迴過程中的 RCW 洩漏 (by Gemini 3 Flash, 2026/05/05)
            Dim subFolders As Folders = rootFolder.Folders
            Try
                For Each f As Folder In subFolders
                    Dim subCount As Long = Await GetMailCountAllOOM(f, cToken:=cToken) ' 遞迴，傳遞 cToken
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
    Private Async Function GetFolderCountAllOOM(rootFolder As Folder, Optional progress As IProgress(Of ProgressReport) = Nothing, Optional skipCache As Boolean = False, Optional cToken As CancellationToken = Nothing) As Task(Of Long)
        ' --------------------------------------------------------------
        ' GetFolderCountAllOOM v1.5: 讀取某資料夾整棵子樹的資料夾總數 (不含 rootFolder 自身)
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
        '   ① BFS 路徑: GetSubtree 內部走 OOM pFolder.Folders 展開，展開後直接 .Count，不需 Layer3 讀取。
        '   ② 遞迴 fallback: 內部的 rootFolder.Folders.Count 和 ForEach 走 OOM，
        '      若日後改為呼叫 GetFolderCountOOM(Layer3)，即可自動走 Redemption ⓪ 路徑。
        ' 2026/04/15 by Claude: 加入 cToken 參數，取代 _cancelRequested 旗標
        ' ---------------------------------------------------------------
        ' 2026/06/13 by Simon/Claude Opus 4.8: 更正 — 同 GetMailCountAllOOM，本函數已被 CollectFolderInfoForceL3 (F5 強刷) 復用，非死碼。
        ' 原始設計意圖:
        ' 	呼叫端 → GetMailCountAllAsync (L2.5) → GetMailCountAllOOM (L3)
        ' 	呼叫端 → GetFolderCountAllAsync (L2.5) → GetFolderCountAllOOM (L3)
        '
        ' 後來使用了BFS剪枝速度更快:
        ' 	Compute → BFS → SumUpSubTreeBottomUp → UpdateFolderInfoCache
        ' 	(自己計算並直接寫入 _cacheMailCountAll / _cacheFolderCountAll，完全繞過 AllAsync系列函數)
        ' ---------------------------------------------------------------
        Dim rName As String = rootFolder.Name ' by Gemini 3.1 Pro 2026/04/16: 避免重複 COM 呼叫
        If _iLikeNoisy Then _dbg("    ├ 開始", rName)

        ' by Gemini, 2026/04/02: 預跑一次顯示準備中
        progress?.Report(New ProgressReport With {.Message = "正在展開資料夾結構...", .IsIndeterminate = True})

        ' 2026/3/24 by Gemini: ⓪ Redemption + 平行處理 (最快路徑)，使用 GetSubtreeToListL3_Rdo 取得清單，以 Parallel.ForEach 搭配 Interlocked.Add(rdoF.Folders.Count) 快速加總
        ' 2026/04/15 by Claude: 改用 ParallelOptions.CancellationToken 取代 _cancelRequested 旗標
        ' 2026/06/13 by Simon/Claude Opus 4.8: 以 _rdoFastPath 開關控管 (預設關)。RDO 枚舉會多算 OOM 看不到的隱藏/非-IPM 夾
        '   (實測 27 vs OOM 22)，故暫關，改走 ② OOM 完整骨架 + 模式過濾以保證與 OOM 一致。
        ' 2026/06/25 by Claude: 移除 ⓪平行BFS _rdoFastPath 死分支(恆 False,從未執行)。原因同 GetMailCountAllOOM,已由 GetSubtreeRdo 解決。

        ' 2026/3/24 by Gemini: ② OOM + BFS 循序 (無 Redemption 時的最後手段) 必須循序處理 OOM COM 物件以避免 STA 違規
        ' 2026/04/15 by Claude: 傳入 cToken，GetSubtree 本身支援取消，OCE 向上冒泡
        Try
            ' 2026/04/16 by Gemini: GetSubtree 現在回傳 Tuple，解開它以維持後續邏輯
            ' 2026/04/17 by Claude: 改呼叫 GetSubtree (L2.5)，享有快取加速
            ' 2026/06/13 by Simon/Claude Opus 4.8: 取得完整骨架後依 _showAllFolders 在計數層過濾 (剪枝移到這裡)；skipCache 一路 thread
            ' 2026/06/25 by Claude Opus 4.8: forceRefresh引數改 skipCache
            ' 2026/06/28 by Simon/Claude [Stage2]: 合約改 (eid,sid,fPath);skeleton/targetTupleList 型別推斷自動跟改。
            '   原 allFolders(.Select(x=>x.folder)) 為死碼(從未使用,真值是清單長度),且新合約無 .folder,故刪除。
            Dim skeleton = Await GetSubtree(rootFolder, includeSubF:=True, progress:=progress, skipCache:=skipCache, cToken:=cToken)
            Dim targetTupleList = FilterSubtreeByMode(skeleton, SafeGetPath(rootFolder))
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
    Private Async Function GetFolderSizeOOM(folder As Folder, Optional progress As IProgress(Of ProgressReport) = Nothing, Optional fPath As String = "", Optional cToken As CancellationToken = Nothing) As Task(Of Long)
        ' --------------------------------------------------------------
        ' GetFolderSizeOOM v1.6: 讀取單一資料夾本層大小 (bytes)
        ' by Gemini, 2026/04/02: 加入 IProgress 支援以回報分批讀取進度 (100ms 節流)
        ' 2026/3/24 by Gemini: Fallback 鏈重構
        '   ⓪ Redemption : rdoFolder.Fields(PR_MESSAGE_SIZE_EXTENDED) (部分 Exchange 支援，極快) <--- 無用，2026/6/27 移除
        '   ① OOM  : pFolder.GetTable(PR_MESSAGE_SIZE_EXTENDED) + GetArray(500) (最快安全招式)
        '   ② OOM  : pFolder.GetTable(PR_MESSAGE_SIZE_EXTENDED) + GetNextRow() (備案)
        '   ③ fail : Return -1
        ' 2026/04/15 by Claude: 加入 cToken 參數
        '   ① ② 迴圈中改用 SmartThrottle(swThrottle, cToken) 取代 Task.Yield()
        '   cToken 取消時 Task.Delay 拋 OCE，由 Catch OCE 接住後 re-throw (讓 GetFolderSizeAllOOM 感知)
        ' --------------------------------------------------------------
        fPath = SafeGetPath(folder, fPath)
        Dim fName As String = ExtractFolderName(fPath)
        If _iLikeNoisy Then _dbg("    ├ 開始", fName)
        Dim sw As Stopwatch = Stopwatch.StartNew()  ' by Claude Sonnet 4.6, 2026/06/07

        ' ⓪ Redemption 層 (嘗試讀取資料夾本身的總量屬性)：RDO 沒有 GetTable().GetArray()，故若屬性讀不到直接 fallback
        ' 2026/06/27 by Simon/Claude Opus 4.8: 移除原 ⓪ 舊 _rdo tier (piggyback session)
        '   原讀 rdoFolder.Fields(PR_MESSAGE_SIZE_EXTENDED) — 本機 PST 無資料夾層級彙總大小屬性,恆回空 → 白做工後落 ①。
        '   故以 ① OOM GetTable+GetArray 逐信加總為單一主路徑。讓這個L3成為單純的OOM函數。

        ' ① OOM GetTable + GetArray(500) (目前最穩、最快的批次讀取)
        Dim table As Outlook.Table = Nothing
        Try
            table = SafeGetTable(folder, "", PR_MESSAGE_SIZE_EXTENDED)
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
            ' 2026/04/15: cToken 取消時 re-throw，讓 GetFolderSizeAllOOM ① 的 For 迴圈感知並中止
            If _iLikeNoisy Then _dbg("    ├ ① OOM GetArray 已取消", fName) : Throw
        Catch ex As System.Exception
            If _iLikeNoisy Then _dbg("    ├ 錯誤: ① OOM GetArray 失敗，走 GetNextRow fallback", $"{fName} | {ex.Message}") ' by Gemini, 2026/04/11: Level 3
        Finally
            TryMarshalRelease(table)
        End Try

        ' ② OOM GetTable + GetNextRow() (不依賴二維陣列的最後保險)
        Dim table2 As Outlook.Table = Nothing
        Try
            table2 = SafeGetTable(folder, "", PR_MESSAGE_SIZE_EXTENDED)
            Dim totalSize As Long = 0
            Dim swThrottle2 As Stopwatch = Stopwatch.StartNew()  ' 2026/04/15: 獨立命名避免與①的 swThrottle 衝突; refactored by Claude Sonnet 4.6, 2026/06/07
            Do While Not table2.EndOfTable
                Dim row As Outlook.Row = table2.GetNextRow()
                If row IsNot Nothing Then
                    totalSize += SafeGet(Of Long)(row, PR_MESSAGE_SIZE_EXTENDED, 0L)
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
    Private Async Function GetFolderSizeAllOOM(rootFolder As Folder, Optional progress As IProgress(Of ProgressReport) = Nothing, Optional skipCache As Boolean = False, Optional cToken As CancellationToken = Nothing) As Task(Of Long)
        ' --------------------------------------------------------------
        ' GetFolderSizeAllOOM v1.6: 讀取某資料夾及整棵子樹的大小總計 (bytes)
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
        '      故改為嚴格的 For 迴圈，逐一 Await GetFolderSizeOOM()。
        '      而內部的 GetFolderSizeOOM 會走到它專屬的 GetTable().GetArray(500) OOM 極速路徑。
        '
        '   ② 兩層都失敗: 回傳 -1，交給上一層流程處理。
        ' 2026/04/15 by Claude: 加入 cToken 參數，取代 _cancelRequested 旗標
        '   ⓪ RDO Parallel.ForEach 路徑: 透過 ParallelOptions.CancellationToken 傳入 cToken
        '   ① OOM 循序路徑: SmartThrottle 每 100ms 讓出一次，cToken 取消時 OCE 冒泡
        '      GetFolderSizeOOM 內部 OCE 會 re-throw，For 迴圈 Catch OCE → Return -1
        ' --------------------------------------------------------------
        Dim rName As String = rootFolder.Name ' by Gemini 3.1 Pro 2026/04/16: 避免重複 COM 呼叫
        If _iLikeNoisy Then _dbg("    ├ 開始", rName)

        ' 2026/3/24 by Gemini: ⓪ Redemption 平行累加 PR_MESSAGE_SIZE_EXTENDED
        ' 2026/04/15 by Claude: 改用 ParallelOptions.CancellationToken 取代 _cancelRequested 旗標
        ' 2026/06/27 by Simon/Claude Opus 4.8: 移除原 ⓪ 舊 _rdo 平行 tier
        '   原用 GetSubtreeToListL3_Rdo 走子樹 + Parallel.ForEach 讀 PR_MESSAGE_SIZE_EXTENDED — 本機 PST 全回空,
        '   validCount=0 拋例外落 ①,等於把整棵子樹白走一遍(① 又用 OOM 走第二遍)。
        '   故以 ① OOM 循序(已走 L2.5 GetSubtree + GetFolderSize 快取)為單一主路徑。讓這個L3成為單純的OOM函數。

        ' 2026/3/24 by Gemini: ① OOM 循序 BFS 累加 (避免 STA 錯誤的保險路徑)
        ' 因為 OOM 的 GetTable() 必須在 UI Thread，我們必須循序 Await 每一層
        ' 2026/04/15 by Claude: _cancelRequested 取代為 SmartThrottle(swThrottle, cToken)
        '   GetFolderSizeOOM 內部 OCE re-throw → For 迴圈 Catch OCE → Return -1
        '   同時移除舊的 i Mod 5 Await Task.Yield()，統一由 SmartThrottle 每 100ms 讓出
        Try
            ' 2026/04/17 by Claude: 改呼叫 GetSubtree (L2.5)，享有快取加速
            ' 2026/06/28 by Simon/Claude [Stage2]: targetFolderList 改 (eid,sid,fPath);逐夾改走免-folder GetFolderSize 多載(RDO優先+OOM fallback),不再物化。
            Dim targetFolderList As List(Of (eid As String, sid As String, fPath As String)) = Await GetSubtree(rootFolder, includeSubF:=True, skipCache:=skipCache, cToken:=cToken)
            Dim grandTotal As Long = 0
            Dim swThrottle As Stopwatch = Stopwatch.StartNew()  ' by Gemini, 2026/04/02; refactored by Claude Sonnet 4.6, 2026/06/07

            For i As Integer = 0 To targetFolderList.Count - 1
                Dim cName As String = ExtractFolderName(targetFolderList(i).fPath)

                ' by Gemini, 2026/04/02: 傳遞 progress 進去以獲得更細緻的(郵件級別)進度回報
                ' 2026/04/15: 同時傳入 cToken，GetFolderSizeOOM 取消時 OCE re-throw 冒泡至此
                ' by Gemini, 2026/04/18: 替換 OOM fallback 路徑，從 GetFolderSizeOOM() 變更為 GetFolderSize() (Layer 2.5) 以利用快取
                Dim sz As Long = Await GetFolderSize(targetFolderList(i).fPath, targetFolderList(i).eid, targetFolderList(i).sid, skipCache:=skipCache, cToken:=cToken)
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
            ' 2026/04/15: GetSubtree 或 SmartThrottle 或 GetFolderSize 取消時冒泡至此
            If _iLikeNoisy Then _dbg("    ├ 錯誤: ① 已取消", $"{rName}") : Return -1
        Catch ex As System.Exception
            If _iLikeNoisy Then _dbg("    ├ 錯誤: ① 循序BFS失敗，放棄計算", $"{rName} | {ex.Message}")
        End Try

        ' ② 兩層都失敗，回傳 -1 讓呼叫端知道失敗了
        Return -1
    End Function
    Private Async Function GetYearCountOOM(folder As Folder, Optional fPath As String = "", Optional cToken As CancellationToken = Nothing) As Task(Of ConcurrentDictionary(Of Integer, Integer))
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
        ' 2026/04/05 by Gemini: 每 100ms 節流讓出執行緒
        ' 2026/04/15 by Claude: 加入 cToken 參數
        '   取代 _cancelRequested 旗標，改用 SmartThrottle(swThrottle, cToken) 節流讓出
        '   cToken 取消時 Task.Delay 拋 OCE，此函數不攔截 (讓 OCE 冒泡至 CollectYearCount)
        '   原因: 攔住後回傳半截 yearCount，L2 會誤以為該資料夾已統計完畢，導致計數偏低
        ' ---------------------------------------------------------------
        fPath = SafeGetPath(folder, fPath)
        Dim fName As String = ExtractFolderName(fPath)
        If _iLikeNoisy Then _dbg("    ├ 開始", fName)

        Dim yearCount As New ConcurrentDictionary(Of Integer, Integer)
        Dim table As Outlook.Table = Nothing
        Try
            ' 2026/3/24 by Gemini: 改用 GetTable + GetArray 取代逐年 Restrict
            table = SafeGetTable(folder, "", "ReceivedTime")    ' 只讀 RcvTime 一欄，最小化每 row 的傳輸量

            ' by Gemini, 2026/04/05: 每批次讀取後，若超過 100ms 則釋放執行緒並檢查中斷
            Dim swThrottle As Stopwatch = Stopwatch.StartNew()  ' by Claude Sonnet 4.6, 2026/06/07
            Do
                Dim data = SafeGetArray(table) : If data Is Nothing Then Exit Do
                For r As Integer = 0 To data.GetUpperBound(0)
                    Dim receivedTime As DateTime = SafeGet(Of DateTime)(data, r, 0, DateTime.MinValue)
                    If receivedTime > DateTime.MinValue Then
                        Dim year As Integer = receivedTime.Year
                        If year > 0 AndAlso year <= Date.Today.Year Then yearCount.AddOrUpdate(year, 1, Function(k, v) v + 1)
                    End If
                Next
                Await SmartThrottle(swThrottle, cToken:=cToken)
                ' 2026/04/15 by Claude: _cancelRequested 取代為 SmartThrottle(swThrottle, cToken) 整合讓出與取消偵測，OCE 冒泡至呼叫端 CollectYearCount
            Loop
        Catch ex As OperationCanceledException
            If _iLikeNoisy Then _dbg("    ├ 已取消", fName) : Throw           ' 2026/04/15: 不攔截 OCE，直接 re-throw 讓 CollectYearCount 感知取消
        Catch ex As System.Exception
            If _iLikeNoisy Then _dbg("    ├ 錯誤", $"{fName}: {ex.Message}")  ' by Gemini, 2026/04/04: Issue 4 格式標準化
        Finally
            TryMarshalRelease(table)
        End Try
        Await Task.Yield()   ' ✅ 函數結束前再讓出一次，確保畫面有機會更新

        If _iLikeNoisy Then _dbg("    ├ 結束", $"{fName} | 年份分佈: {yearCount.Count}")
        Return yearCount

    End Function
    Private Async Function GetMonthCountOOM(folder As Folder, year As Integer, Optional fPath As String = "", Optional cToken As CancellationToken = Nothing) As Task(Of ConcurrentDictionary(Of Integer, Integer))
        ' ---------------------------------------------------------------
        ' GetMonthCountOOM — Layer3 COM 資料層: 單一資料夾月份郵件分佈
        ' 職責: 對 Outlook 發出 COM 呼叫，回傳單一資料夾在指定年份中每個月的郵件數量
        ' 規則: 不做快取、不做提前過濾、不遞迴 (這些全部交給 GetMonthCount L2.5 負責)
        '       OCE 不在此攔截，直接 re-throw 讓呼叫端感知取消
        ' 原始設計: 2026/3/24 by Gemini — 從逐月 Restrict 改為 GetTable + GetArray 一次讀完
        '   原本 12 次 Restrict + 12 次 Items.Count = 24 次 COM call
        '   現在 1 次 GetTable (含日期範圍 filter) + ceil(N/1000) 次 GetArray
        ' 2026/04/15 by Claude/Gemini: 加入 cToken 參數與 fPath 參數
        '   由 L2.5 直接傳入 fPath，完全消除 pFolder.FolderPath 的 COM 開銷
        ' 2026/04/17 by Claude: 拆出快取/提前過濾邏輯至 GetMonthCount (L2.5)，此函數僅剩純 COM
        ' ---------------------------------------------------------------
        fPath = SafeGetPath(folder, fPath)
        Dim fName As String = ExtractFolderName(fPath)
        If _iLikeNoisy Then _dbg("    ├ 開始", $"{fName} ({year} 年)")

        Dim monthCount As New ConcurrentDictionary(Of Integer, Integer)
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
                    If receivedTime > DateTime.MinValue Then monthCount.AddOrUpdate(receivedTime.Month, 1, Function(k, v) v + 1)
                Next

                Await SmartThrottle(swThrottle, cToken:=cToken)
                ' by simon, 2026/04/08: 每批次讀取後，若超過 100ms 則釋放執行緒並檢查中斷
                ' 2026/04/15 by Claude: SmartThrottle 取代舊的 swThrottle + Task.Delay(1) + _cancelRequested
                ' 2026/04/15: OCE 向上冒泡，不在此攔截 (快取寫入在呼叫端 L2.5，OCE 自然繞過)
            Loop
        Catch ex As OperationCanceledException
            If _iLikeNoisy Then _dbg("    ├ 已取消", $"{fName} ({year} 年)") : Throw      ' 2026/04/15: re-throw 讓 GetMonthCount (L2.5) 感知，不寫入快取
        Catch ex As System.Exception
            If _iLikeNoisy Then _dbg("    ├ 錯誤", $"{fName}, year={year}: {ex.Message}") ' by Gemini, 2026/04/04: Issue 4 格式標準化
        Finally
            TryMarshalRelease(table)
        End Try

        If _iLikeNoisy Then _dbg("    ├ 結束", $"{fName} ({year} 年)")
        Return monthCount
    End Function
    Private Async Function GetAttMailListOOM(folder As Folder, progress As IProgress(Of ProgressReport), Optional cToken As CancellationToken = Nothing) As Task(Of List(Of MailItemInfo))
        ' ----------------------------------------------------------------------------------------
        ' Phase 1 / Layer3 純資料層: GetTable + GetArray 批次掃描單一資料夾
        ' 設計: 這裡只專注於透過 MAPI 取回資料，不會做快取判定，也無關大小設定過濾
        ' 2026/04/15 by Claude: 加入 cToken 參數，取代 _cancelRequested 旗標
        '   SmartThrottle(swThrottle, cToken) 每 100ms 讓出一次，cToken 取消時拋 OCE
        '   取消時捕捉 OCE → 回傳空 List (不回傳已掃到的半截清單)
        '   原因: 呼叫端 GetAttMailList 取消時不寫入快取 (見其 cToken.IsCancellationRequested 判斷)
        ' ----------------------------------------------------------------------------------------
        Dim fName As String = folder.Name ' by Gemini 3.1 Pro 2026/04/16: 避免重複 COM 呼叫
        If _iLikeNoisy Then _dbg(" ├ 開始", fName)

        Dim table As Outlook.Table = Nothing

        Dim filterHasAttach As String = "@SQL=" & Chr(34) & "urn:schemas:httpmail:hasattachment" & Chr(34) & " = True"
        ' 預分配容量為 4096，顯著降低掃描大量附件郵件時的記憶體配置開銷 (by Gemini 3 Flash, 2026/05/04)
        Dim result As New List(Of MailItemInfo)(4096)

        ' 2026/04/22 by Gemini 3.1 Pro: 提前取得路徑，讓此資料夾內的所有郵件都能獲得歸屬路徑，且只需 1 次 COM 存取
        Dim fPath As String = ""
        fPath = SafeGetPath(folder)

        Try
            table = SafeGetTable(folder, filterHasAttach, "EntryID", "Subject", PR_MESSAGE_SIZE, "ReceivedTime", "SenderName")

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
    Private Function GetAttFilenameOOM(ByRef mail As MailItemInfo) As List(Of String)
        ' by Gemini, 2026/04/04: 取得郵件的附件檔名清單 (純 Layer3 邏輯，不做快取)
        If _iLikeNoisy Then _dbg("    ├ 開始", mail.Subject)
        Dim result As New List(Of String)(mail.AttCount)

        ' ⓪ Redemption 優先: 繞過 OOM 開信的記憶體開銷，直接透過 MAPI Table 抓取檔名
        ' ⓪ RDO 路徑已上移至 L2.5 GetAttFilenameRdo(store-scoped on _rdo2),L3 回歸純 OOM。 2026/06/23 by Simon/Claude

        ' ① Fallback: 使用 Outlook OOM (極為昂貴的物件實例化)
        Dim tempMail As Outlook.MailItem = Nothing
        Dim objAttach As Outlook.Attachments = Nothing
        Try
            tempMail = TryCast(_olNS.GetItemFromID(mail.EntryID), Outlook.MailItem)
            If tempMail IsNot Nothing Then
                objAttach = tempMail.Attachments
                Dim attCount As Integer = objAttach.Count                       ' by simon 2026/04/19: 存成變數避免 COM 呼叫重複
                Dim olbyValue As Integer = Outlook.OlAttachmentType.olByValue   ' by simon 2026/04/19: 存成變數避免 COM 呼叫重複
                For i As Integer = 1 To attCount ' COM 是 1-based index
                    Dim att As Outlook.Attachment = objAttach.Item(i)
                    Try : If att.Type = olbyValue Then result.Add(att.FileName) ' 2026/04/09 by Gemini: 僅處理 olByValue (1) 類型的附件
                    Finally : TryMarshalRelease(att)
                    End Try
                Next
            End If
        Catch ex As System.Exception
            _dbg("    ├ ① OOM 失敗", $"{mail.EntryID}: {ex.Message}")   ' 2026/07/06 by Simon/Claude: 無條件記錄(原本只在 _iLikeNoisy 才印)，失敗留痕
            Return Nothing   ' 2026/07/06 by Simon/Claude: 失敗回 Nothing 而非空清單，避免跟「真的沒附件」混淆——
            '   L2.5 GetAttFilename 已有「非 Nothing 才快取」判斷，回 Nothing 才不會被 SaveAttFilenameBatch 永久寫進 att_filenames
        Finally
            If _iLikeNoisy Then _dbg(" ├ 結束")
            TryMarshalRelease(objAttach)
            TryMarshalRelease(tempMail)
        End Try

        Return result
    End Function
    Private Function GetMailBodyOOM(entryID As String) As String
        ' ---------------------------------------------------------------
        ' GetMailBodyOOM — Layer3 COM 資料層：讀取郵件 Body 並正規化
        ' 2026/04/28 by Simon/Claude: 以 Simon 的 GetMailBodyByEntryID 為基礎
        '   + 加入 NormalizeMailBody 正規化（去除 HTML 標籤、空白換行）
        '   + Await Task.Yield() 確保每封讀完後讓 UI 執行緒喘氣
        '   支援 MailItem 與 PostItem 兩種型別
        '   使用獨立的 ns（Simon 的設計），確保 COM namespace 不跨執行緒共用
        ' ---------------------------------------------------------------
        If String.IsNullOrEmpty(entryID) Then Return ""

        ' Dim ns As Outlook.NameSpace = Nothing
        Dim item As Object = Nothing
        Dim mailBody As String = Nothing   ' 2026/07/06 by Simon/Claude: 初始值改 Nothing(原 "")，失敗與「真的空內文」才分得開——
        '   成功路徑一定會走到下面的 NormalizeMailBody 賦值，真空信會拿到 ""，失敗維持 Nothing
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

    End Function
    Private Async Function GetMailInfoOOM(folder As Folder, needTopic As Boolean, cToken As CancellationToken, Optional fPath As String = "") As Task(Of List(Of (Mail As MailItemInfo, Topic As String)))
        ' ---------------------------------------------------------------
        ' 2026/05/06 by Claude: 永遠讀取全部 8 欄（含 topic/msgId/senderEmail）
        '   needTopic 參數保留供 API 相容，但 L3 層已不區分，統一讀取
        '   欄位索引: 0=EntryID, 1=Subject, 2=Size, 3=RcvTime,
        '             4=SenderName, 5=Topic, 6=MsgIDhash, 7=SenderEmail
        '
        ' 2026/06/12 by Simon/Claude Opus 4.8: 移除 PR_CONVERSATION_TOPIC (欄位 5) topic 改由 GetCleanSubject(subject) 動態計算，與 DB 讀取路徑保持一致
        '   欄位索引: 0=EntryID, 1=Subject, 2=Size, 3=RcvTime, 4=SenderName, 5=MsgIDhash, 6=SenderEmail
        ' 2026/07/02 by Simon/Claude [Task 2a]: 分頁掃描骨架(開table→GetArray→節流→取消)抽到共用的 ScanFolderTable,
        '   本函式只留欄位解析(row parser)。table/folder 生命週期責任不變:table 由 ScanFolderTable 自己開/釋放,folder 仍由呼叫端持有+不釋放。
        ' ---------------------------------------------------------------
        fPath = SafeGetPath(folder, fPath)
        Dim fName As String = ExtractFolderName(fPath)
        If _iLikeNoisy Then _dbg("    ├ 開始 (掃描)", fName)

        Dim resultList As New List(Of (MailItemInfo, String))(4096) ' 預分配容量為 4096，優化批次讀取郵件基本資訊時的清單填充 (by Gemini 3 Flash, 2026/05/04)
        Try
            Await ScanFolderTable(folder, cToken, ThrottleFreq.Mid, Nothing,
                Sub(data, r)
                    Dim entryID As String = SafeGet(Of String)(data, r, 0, "")
                    If entryID = "" Then Return

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
                End Sub,
                "EntryID", "Subject", PR_MESSAGE_SIZE,     ' 0~2
                "ReceivedTime", "SenderName",              ' 3~4
                PR_INTERNET_MESSAGE_ID_A, PR_SENDER_EMAIL_ADDRESS_A)   ' 5~6 (2026/05/06 by Claude: Tab5 去重)

        Catch ex As System.Exception
            If _iLikeNoisy Then _dbg("    ├ 錯誤", $"{fName}: {ex.Message}")
        Finally
            If _iLikeNoisy Then _dbg("    ├ 結束 (掃描)")
        End Try
        Return resultList
    End Function
    Private Function RefreshMailInfoOOM(ByRef info As MailItemInfo) As RefreshResult
        ' ---------------------------------------------------------------
        ' RefreshMailInfoOOM — Layer3 OOM 讀取層: 依 EntryID 從 Outlook COM 重讀單封郵件實體資訊寫回 info。
        '   回傳 Updated/NotFound/TransientError；失效郵件的移除政策由呼叫端決定 (目前一律保留+記錄)
        ' 2026/07/02 by Simon/Claude: 從舊 RefreshMailInfoL3 拆出,比照 GetMailInfoRdo/GetMailInfoOOM 切分方式;
        '   分派層改名 RefreshMailInfo 並移至 Layer2.5 區塊,RDO 對應層 RefreshMailInfoRdo 留在 Layer3 RDO 區塊。
        ' 2026/07/04 by Simon/Claude: 移除 readAttachCount 分支與附件枚舉 — 探針證實批次無法讀、逐封枚舉又是 Tab3 篩選用不到的死碼。
        ' ---------------------------------------------------------------
        Const MAPI_E_NOT_FOUND As Integer = &H8004010F   ' 找不到物件的 HRESULT，用以區分 NotFound vs 暫時性錯誤
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

            Return RefreshResult.Updated

        Catch ex As System.Runtime.InteropServices.COMException
            ' 區分 NotFound (ID 失效) 與暫時性錯誤，供呼叫端日後移除政策使用
            If ex.ErrorCode = MAPI_E_NOT_FOUND Then
                If _iLikeNoisy Then _dbg("    ├ NotFound", info.EntryID)
                Return RefreshResult.NotFound
            End If
            If _iLikeNoisy Then _dbg("    ├ COM 暫時性錯誤", ex.Message)
            Return RefreshResult.TransientError
        Catch ex As System.Exception
            If _iLikeNoisy Then _dbg("    ├ 例外", ex.Message)
            Return RefreshResult.TransientError
        Finally
            TryMarshalRelease(mail)
        End Try
    End Function
    Private Async Function GetMailInfoAsDict(fPath As String, ct As CancellationToken) As Task(Of Dictionary(Of String, MailItemInfo))
        ' 2026/06/14 by Simon/Claude Opus 4.8: 方法B底層 — 對單一資料夾一次 GetTable+GetArray，回傳 EntryID→基本欄位 dict
        '   不加 hasattachment 過濾 (要服務 Lv3/4/5 任意郵件)；解析不到資料夾 → 回 Nothing 讓呼叫端退回方法A
        '   路徑→Folder：經 _cacheFolderIDs 取 (eid,sid) 再 GetFolderFromID (與既有 L3 慣例一致)
        ' 2026/07/02 by Simon/Claude [Task 2a]: 分頁掃描骨架(開table→GetArray→節流→取消)與 GetMailInfoOOM 共用,抽到 ScanFolderTable。
        '   本函式只留欄位解析(row parser)+folder 物化/釋放責任(維持不變:自己解析+自己釋放)。
        ' 2026/07/02 by Simon/Claude [Task 2a 收尾]: 原名 GetFolderBasicByEntryIDL3,改名 GetMailInfoAsDict —
        '   不叫 GetMailInfo() 是因為它沒有 ①記憶體②DB 快取層(只有③RDO+OOM),不符合這個專案對「L2.5」的定義(=快取代理);
        '   不掛 OOM 字尾是因為它現在 RDO優先/OOM兜底都會走,掛 OOM 會誤導。AsDict 純粹標示回傳容器形狀跟另兩個 List 版本不同。
        '   內部改成 RDO優先(借用 GetMailInfoRdo,免物化OOM folder) → OOM兜底(RDO讀不到才物化,原邏輯不動)。
        Dim ids As (eid As String, sid As String, isMail As Boolean, hasCh As Boolean) = Nothing
        If Not _cacheFolderIDs.TryGetValue(fPath, ids) Then Return Nothing

        ' ③ RDO 優先 — 借用 GetMailInfoRdo,免物化 OOM folder
        If _rdo2 IsNot Nothing Then
            Dim rdoResult = GetMailInfoRdo(fPath, ids.eid, ids.sid)
            If rdoResult IsNot Nothing Then Return rdoResult.ToDictionary(Function(m) m.EntryID, StringComparer.Ordinal)
        End If

        ' ③ Fallback: RDO 讀不到才物化 OOM folder → 掃描
        Dim folder As Folder = TryCast(_olNS.GetFolderFromID(ids.eid, ids.sid), Folder)
        If folder Is Nothing Then Return Nothing

        Dim result As New Dictionary(Of String, MailItemInfo)(StringComparer.Ordinal)
        Try
            Await ScanFolderTable(folder, ct, ThrottleFreq.Hii, Sub() PgrsBar2.Text = $"批次掃描 {folder.Name}: {result.Count} 筆...",
                Sub(data, r)
                    Dim entryID As String = SafeGet(Of String)(data, r, 0, "")
                    If entryID = "" Then Return
                    result(entryID) = New MailItemInfo With {.EntryID = entryID,
                                                             .Subject = SafeGet(Of String)(data, r, 1, ""),
                                                             .Size = SafeGet(Of Long)(data, r, 2, 0L),
                                                             .RcvTime = SafeGet(Of DateTime)(data, r, 3, DateTime.MinValue),
                                                             .SenderName = SafeGet(Of String)(data, r, 4, "")}
                End Sub,
                "EntryID", "Subject", PR_MESSAGE_SIZE, "ReceivedTime", "SenderName")
        Catch ex As OperationCanceledException
            Throw
        Catch ex As System.Exception
            If _iLikeNoisy Then _dbg("    ├ GetMailInfoAsDict 錯誤", $"{fPath} — {ex.Message}")
            Return Nothing
        Finally
            TryMarshalRelease(folder)
        End Try
        Return result
    End Function
    Private Async Function GetSubtreeOOM(rootFolder As Folder, includeSubF As Boolean, Optional progress As IProgress(Of ProgressReport) = Nothing, Optional cToken As CancellationToken = Nothing) As Task(Of List(Of (eid As String, sid As String, fPath As String)))
        ' --------------------------------------------------------------
        ' GetSubtreeOOM — Layer3 COM 資料層: BFS 純掃描
        ' 原名 GetSubtree，2026/04/17 by Claude: 拆出快取邏輯至 GetSubtree (L2.5)
        ' 職責: BFS 廣度優先掃描 rootFolder 下整棵子樹，回傳清單
        ' 規則: 不做快取讀取；BFS 完成後若未中斷，由本層負責寫入 _cacheSubTreeList
        '       OCE 中斷時 re-throw，不寫快取 (確保不存入不完整的樹)
        ' 2026/04/16 by Gemini: 升級回傳 Tuple (Folder, fPath)，消除呼叫端對 COM .FolderPath 的依賴
        ' 2026/06/28 by Simon/Claude [Stage2]: 回傳合約改 (eid,sid,fPath)。BFS queue 仍持 Folder 供走訪,
        '   但 result 改存 eid/sid 字串(BFS 時順手讀,本就要讀來填 _cacheFolderIDs);靜態快取從此不握 COM。OOM 為罕見 fallback,物化無法避免但不外洩。
        ' --------------------------------------------------------------
        ' 2026/04/24 by Gemini 3.0 flash: 使用 SafeGetPath 並增加 root 狀態檢查
        Dim rootPath As String = SafeGetPath(rootFolder)
        If String.IsNullOrEmpty(rootPath) Then
            _dbg(" ├ 錯誤", "無法取得 rootFolder 路徑，中斷掃描")
            Return New List(Of (eid As String, sid As String, fPath As String))
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
        Dim result As New List(Of (eid As String, sid As String, fPath As String))(512)
        result.Add((rootFolder.EntryID, rootFolder.StoreID, rootPath))

        If Not includeSubF Then
            sw.Stop()
            If _iLikeNoisy Then _dbg("    ├ 結束", $"{rootName} (Single) | {sw.ElapsedMilliseconds}ms")
            Return result
        End If

        ' 🆕 2026/06/24 by Simon/Claude Opus 4.8: RDO 快速探索 tier(探針 C 實證 ~60-150× 提速,Step1 正確性對齊版)
        '   旗標開且 _rdo2 在 → GetSubtreeRdo 批次走訪;成功寫 _cacheSubTreeList 回傳,失敗(Nothing)往下掉回 OOM BFS。
        '   _cacheFolderIDs/_cacheFolderCount 由 GetSubtreeRdo 內比照 BFS 註冊。OOM BFS 當 fallback。

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
                            '   過濾移到計數/顯示層(FilterSubtreeByMode)。徹底消滅「folder_info 殘缺 → 子樹靜默少算」未爆彈。
                            '   (原剪枝: If Not _showAllFolders AndAlso Not isMail Then Continue For — 已移除)

                            ' ✅ 加強 EntryID/StoreID 讀取的安全性
                            ' 2026/06/28 [Stage2]: eid/sid 一次讀進區域變數,同時供 _cacheFolderIDs 與 result(資料 tuple),避免重複 COM
                            Dim cEid As String = "" : Dim cSid As String = ""
                            Try : cEid = subF.EntryID : cSid = subF.StoreID : Catch : End Try
                            Try
                                _cacheFolderIDs.TryAdd(childPath, (cEid, cSid, isMail, TextHasChineseChar(fName)))
                            Catch : End Try

                            result.Add((cEid, cSid, childPath))   ' Stage2: 回傳資料 tuple,不帶 COM(queue 仍持 subF 供 BFS 走訪)
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
    Private Function PeekLiveFolderSnapOOM(folder As Folder, Optional fPath As String = "") As Integer
        ' ---------------------------------------------------------------
        ' 快速讀取 PR_CONTENT_COUNT，專門只用於 SQLite snapshot 驗證
        ' 故意不走完整 Layer3 fallback 的GetMailCount，只走最快的 PropertyAccessor 路徑
        ' 失敗時回傳 -999 (不可能等於任何正常 snapshot 值，確保快取失效)
        ' 2026/4/7 by Gemini, 解決 SSD 讀出後 snapshot 驗證失敗導致的重複統計問題
        ' 2026/7/2 by simon, 直接保底呼叫一次 folder.Items.Count, 比RDO 要先解 store/folder 才能讀屬性的解析開銷快，所以也不需轉到_rdo2)
        ' ---------------------------------------------------------------
        Dim fName As String = ExtractFolderName(fPath)
        If _iLikeNoisy Then _dbg(" ├ 開始", fName)
        Try
            Return CLng(folder.PropertyAccessor.GetProperty(PR_CONTENT_COUNT))
        Catch
            Try : Return folder.Items.Count : Catch : Return -999 : End Try
        End Try
    End Function
    Private Function PeekFolderLastUpdateTime(folder As Folder, Optional fPath As String = "") As Long
        ' ---------------------------------------------------------------
        ' 快速讀取 PR_LOCAL_COMMIT_TIME_MAX (0x670A0040, PT_SYSTIME)，回傳 Ticks
        ' RenewCacheToDB 的第二 dirty 訊號：資料夾內任何增/刪/改 (含標旗、已讀狀態) 都會推高此值，
        ' 專抓「郵件數不變但內容已置換」的淨零變動 (copy→修改→放回→刪原始)，純 count 快照對此全盲。
        ' 失敗回 -1 (=未知)：呼叫端遇 -1 一律退回純 count 比對，不誤判 dirty (與 PeekLiveFolderSnapOOM 的 -999 策略同思路)
        ' 注意：純 PST 壓縮不會推高此值，該情境需靠 GetItemFromID 解析失敗時的自癒機制，另案處理
        ' 2026/07/04 by Simon/Claude Fable 5
        ' ---------------------------------------------------------------
        Try
            Return CDate(folder.PropertyAccessor.GetProperty(PR_LOCAL_COMMIT_TIME_MAX)).Ticks
        Catch
            Return -1L
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
        ' 	呼叫端 → GetMailCountAllAsync (L2.5) → GetMailCountAllOOM (L3)
        ' 	呼叫端 → GetFolderCountAllAsync (L2.5) → GetFolderCountAllOOM (L3)
        '
        ' 後來使用了BFS剪枝速度更快:
        ' 	Compute → BFS → SumUpSubTreeBottomUp → UpdateFolderInfoCache
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
        Dim total As Long = Await GetMailCountAllOOM(folder, progress, cToken:=cToken)
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
        ' 	呼叫端 → GetMailCountAllAsync (L2.5) → GetMailCountAllOOM (L3)
        ' 	呼叫端 → GetFolderCountAllAsync (L2.5) → GetFolderCountAllOOM (L3)
        '
        ' 後來使用了BFS剪枝速度更快:
        ' 	Compute → BFS → SumUpSubTreeBottomUp → UpdateFolderInfoCache
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
        count = Await GetFolderCountAllOOM(folder, cToken:=cToken)
        If count >= 0 Then _cacheFolderCountAll.TryAdd(fPath, count)  ' 2026/04/15: 改用 cToken 判斷
        Return count
    End Function
#End Region
#Region "  └ 其他輔助函數"
    Private Function GetStoreNameFromPath(folderPath As String) As String
        ' 2026/06/19 by Simon/Claude: 從 \\StoreName\sub\... 取出最前面的 store 顯示名 (店名含逗號/空格/&/~/@ 皆安全, 因唯一分隔符是反斜線)
        Dim p = folderPath.TrimStart("\"c)
        Dim idx = p.IndexOf("\"c)
        Return If(idx > 0, p.Substring(0, idx), p)
    End Function
    Private Function GetFolderById(eid As String, sid As String) As Folder
        ' 2026/06/28 by Simon/Claude: 集中物化點 — eid+sid → OOM Folder。
        '   給免-folder 多載的 OOM fallback + (Stage3) B群 thunk 共用。
        '   eid 用 RDO table 原生 eid 即可(Stage0 證 796 夾含 IMAP 全可物化)。
        If String.IsNullOrEmpty(eid) Then Return Nothing
        Try : Return TryCast(_olNS.GetFolderFromID(eid, sid), Folder) : Catch : Return Nothing : End Try
    End Function
    Private Async Function ScanFolderTable(folder As Folder, cToken As CancellationToken, throttleFreq As Integer, onThrottled As System.Action, rowHandler As System.Action(Of Object(,), Integer), ParamArray columns() As String) As Task
        ' 2026/07/02 by Simon/Claude [Task 2a]: GetMailInfoAsDict 與 GetMailInfoOOM 共用的分頁掃描骨架 —
        '   開table → GetArray迴圈 → 節流 → 取消處理。欄位解析交給呼叫端傳入的 rowHandler(對自己的容器 closure 累加),
        '   本函式不知道、也不管呼叫端要 List 還是 Dictionary。
        '   table 生命週期由本函式自己管(開+Finally釋放);folder 的物化/釋放責任仍在呼叫端,本函式不介入(對稱既有 GetFolderById 的角色)。
        Dim table As Outlook.Table = Nothing
        Try
            table = SafeGetTable(folder, "", columns)
            Dim sw As Stopwatch = Stopwatch.StartNew()
            Do
                cToken.ThrowIfCancellationRequested()
                Dim data = SafeGetArray(table)
                If data Is Nothing Then Exit Do
                For r As Integer = 0 To data.GetUpperBound(0)
                    rowHandler(data, r)
                Next
                Await SmartThrottle(sw, cToken, throttleFreq, onThrottled)
            Loop
        Finally
            TryMarshalRelease(table)
        End Try
    End Function
    Private Function CcIsMail(cc As String) As Boolean
        ' 2026/06/28 by Simon/Claude [Stage2]: PR_CONTAINER_CLASS → isMail。Stage0 鎖定規則(796 夾實證,含 IMAP)。
        '   IPF.Note/Post/Imap → mail;空 → mail(保守,實證全是 Inbox/Sent/Deleted/_IRM);其餘(如 IPF.Configuration)→ 非 mail。
        '   供 GetSubtreeRdoBatch 去物化後推 isMail 用(取代原 Phase2 的 IsMailFolder(f))。
        If String.IsNullOrEmpty(cc) Then Return True
        Return cc.StartsWith("IPF.Note", StringComparison.OrdinalIgnoreCase) OrElse
               cc.StartsWith("IPF.Post", StringComparison.OrdinalIgnoreCase) OrElse
               cc.StartsWith("IPF.Imap", StringComparison.OrdinalIgnoreCase)
    End Function
    Private Function SafeGetPath(folder As Folder, Optional existingPath As String = "") As String
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
    Private Function SafeGetDbRow(folder As Folder, fPath As String) As FolderInfoDbRow
        ''' <summary>
        ''' [Helper] 嘗試從資料庫取得有效的資料夾統計列，並自動處理快取回填。
        ''' </summary>
        ''' <param name="folder">Outlook 資料夾物件</param>
        ''' <param name="fPath">資料夾路徑</param>
        ''' <returns>有效的 FolderInfoDbRow，若無命中或 Snapshot 不符則回傳 Nothing</returns>
        ' 2026/05/09 by Gemini 3.1 Pro: 統一處理 DB lazy load 與 Snapshot 驗證邏輯
        ' snapshot 驗證: DB 儲存的 pr_count_snap = save 時的 PR_CONTENT_COUNT 值
        '   用 PeekLiveFolderSnapOOM (單次 PropertyAccessor call) 與 snapshot 比對
        '   相同 → 快取仍有效；不同 → 資料夾內容已變，跳過 DB，呼叫 Layer3
        Dim row = LazyGetFolderInfo(fPath)
        If row IsNot Nothing AndAlso PeekLiveFolderSnapOOM(folder, fPath) = row.snap Then
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

        ' 2026/07/03 by Simon/Claude Fable 5: memoization — 相同 subject 字串直接回傳快取結果，不重跑 Regex
        Dim cached As String = Nothing
        If _cacheCleanSubject.TryGetValue(subject, cached) Then Return cached

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
        Dim result = _subjectPrefixRe.Replace(subject, "")
        _cacheCleanSubject(subject) = result   ' 2026/07/03 by Simon/Claude Fable 5: 存回 memoization 快取
        Return result

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
        ' 2026/6/22: .Body 是純文字，HTML 去標籤幾乎是空轉（且可能誤刪）GetMailBodyOOM 讀的是.Body（OOM 回傳純文字）， 不是.HTMLBody。所以純文字信幾乎沒有 <...> 標籤可去。
        result = _reHtmlTag.Replace(body, "")       ' 去除 HTML 標籤
        result = _reHtmlEntity.Replace(result, "")  ' 去除常見 HTML entities ' 2026/6/6 by Gemini: 改用 Regex 優化多重 Replace效能
        result = _reQuoteMarker.Replace(result, "") ' 2026/06/18 by Simon/Claude Opus 4.8: 去除行首轉寄引用前綴(> 之間可夾雜空白)，保留引用文字本體。必須在去空白前做(靠行首定位)
        result = _reWhitespace.Replace(result, "")  ' 去除所有空白字元(含 "\u3000"=全形空白、Tab、換行)
        Return result.ToLowerInvariant()
    End Function
    Private Sub FillCacheFromDbRow(fPath As String, row As FolderInfoDbRow, Optional skipAggregates As Boolean = False)
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

            ' by Gemini, 2026/04/10: 填充身分標識與標籤快取 (2026/07/03: isMail 併入 _cacheFolderIDs，_cacheIsMailFolder 汰除)
            If Not String.IsNullOrEmpty(.eid) Then _cacheFolderIDs.TryAdd(fPath, (.eid, .sid, .isMail = 1, .hasCh = 1))
        End With
    End Sub
    Private Sub ClearMonthCountMemory(fPath As String)
        ' 清除單一資料夾在 _cacheMonthCount 記憶體的所有年份 entry。
        ' _cacheMonthCount key 格式為 "fPath_year"，不知道呼叫當下是哪一年有快取，故用前綴比對清所有匹配的。
        ' 2026/07/09 by Simon/Claude: 消重 — 原本 RenewCacheToDB(孤兒/dirty 兩分支各一次)與 InvalidateMailCache
        '   (Form1_Maintab56.vb) 各自重複同一段 3 行迴圈，抽出共用；三處呼叫端已改呼叫本函式。
        For Each mk In _cacheMonthCount.Keys.Where(Function(k) k.StartsWith(fPath & "_")).ToList()
            _cacheMonthCount.TryRemove(mk, Nothing)
        Next
    End Sub
#End Region
#End Region

    ''' GetFolderCount 兩支同型,只把插槽換成 _cacheFolderCount / Function(r) r.fc / AddressOf GetFolderCountRdo / AddressOf GetFolderCountOOM
    '' 2026/7/5: 同型重構by Claude:
    'Private Function GetMailCount(folder As Folder, Optional fPath As String = "", Optional skipCache As Boolean = False) As Long
    '    Return GetFolderStatCore(fPath, "", "", folder, _cacheMailCount, Function(r) r.mc, AddressOf GetMailCountRdo, AddressOf GetMailCountOOM, skipCache)
    'End Function

    'Private Function GetMailCount(fPath As String, eid As String, sid As String, Optional skipCache As Boolean = False) As Long
    '    Return GetFolderStatCore(fPath, eid, sid, Nothing, _cacheMailCount, Function(r) r.mc, AddressOf GetMailCountRdo, AddressOf GetMailCountOOM, skipCache)
    'End Function

    'Private Function GetFolderStatCore(fPath As String, eid As String, sid As String, folder As Folder, cache As ConcurrentDictionary(Of String, Long), dbField As Func(Of FolderInfoDbRow, Long),
    '                                   rdoFn As Func(Of String, String, String, Long), oomFn As Func(Of Folder, String, Long), skipCache As Boolean) As Long
    '    ' ①②③ 骨架唯一實作(原 GetMailCount/GetFolderCount 四支多載的共同結構)
    '    '   folder IsNot Nothing → ② 走 SafeGetDbRow(帶 snap 檢查);folder Is Nothing → LazyGetFolderInfo(信任 DB)
    '    Dim value As Long
    '    If Not skipCache Then
    '        If cache.TryGetValue(fPath, value) Then Return value                 ' ① 記憶體命中
    '        Dim row = If(folder IsNot Nothing, SafeGetDbRow(folder, fPath), LazyGetFolderInfo(fPath)) ' ② DB lazy
    '        If row IsNot Nothing AndAlso dbField(row) >= 0 Then Return dbField(row)
    '    End If
    '    If folder IsNot Nothing Then eid = folder.EntryID : sid = folder.StoreID ' ⚠坑1: ①②未命中才碰 COM 屬性,熱路徑保持零 COM
    '    value = rdoFn(fPath, eid, sid)                                           ' ③ RDO 優先
    '    If value < 0 Then                                                        ' 底線: RDO 失敗 → OOM
    '        If folder Is Nothing Then folder = GetFolderById(eid, sid)
    '        If folder IsNot Nothing Then value = oomFn(folder, fPath)
    '    End If
    '    If value >= 0 Then cache.TryAdd(fPath, value)                            ' 回填快取
    '    Return value
    'End Function

End Class
