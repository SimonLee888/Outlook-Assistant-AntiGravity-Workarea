Imports System.Collections.Concurrent
Imports System.Threading
Imports Microsoft.Data.Sqlite
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Outlook

Partial Class Form1

#Region "■ 01 Win32 API 宣告"
    ' 統一使用 DllImport (取代舊式 Declare Function)
    ' 2026-03-23 整理: 移除重複的 SendMessage Declare 版本，補齊 FindWindow / FindWindowEx
    <Runtime.InteropServices.DllImport("user32.dll", CharSet:=Runtime.InteropServices.CharSet.Auto)>
    Private Shared Function FindWindow(
        ByVal lpClassName As String,
        ByVal lpWindowName As String) As IntPtr
    End Function
    <Runtime.InteropServices.DllImport("user32.dll", CharSet:=Runtime.InteropServices.CharSet.Auto)>
    Private Shared Function FindWindowEx(
        ByVal hWndParent As IntPtr,
        ByVal hWndChildAfter As IntPtr,
        ByVal lpszClass As String,
        ByVal lpszWindow As String) As IntPtr
    End Function
    <Runtime.InteropServices.DllImport("user32.dll", CharSet:=Runtime.InteropServices.CharSet.Auto)>
    Private Shared Function SendMessage(
        ByVal hWnd As IntPtr,
        ByVal msg As Integer,
        ByVal wParam As IntPtr,
        ByVal lParam As IntPtr) As IntPtr
    End Function

    <Runtime.InteropServices.DllImport("user32.dll")>
    Private Shared Function ShowWindow(
        ByVal hWnd As IntPtr,
        ByVal nCmdShow As Integer) As Boolean
    End Function
    <Runtime.InteropServices.DllImport("user32.dll")>
    Private Shared Function LockWindowUpdate(
        ByVal hWnd As IntPtr) As Boolean
    End Function
    <Runtime.InteropServices.DllImport("user32.dll")>
    Private Shared Function RedrawWindow(
        ByVal hWnd As IntPtr,
        ByVal lprcUpdate As IntPtr,
        ByVal hrgnUpdate As IntPtr,
        ByVal flags As UInteger) As Boolean
    End Function

    <Runtime.InteropServices.DllImport("user32.dll")>
    Private Shared Function PostMessage(
        ByVal hWnd As IntPtr,
        ByVal msg As Integer,
        ByVal wParam As IntPtr,
        ByVal lParam As IntPtr) As Boolean
    End Function
    <Runtime.InteropServices.DllImport("user32.dll")>
    Private Shared Function SetWindowPos(
        ByVal hWnd As IntPtr,
        ByVal hWndInsertAfter As IntPtr,
        ByVal x As Integer,
        ByVal y As Integer,
        ByVal cx As Integer,
        ByVal cy As Integer,
        ByVal uFlags As Integer) As Boolean
    End Function

    ' === 用來強制移除 SplitContainer 焦點框 ===
    <Runtime.InteropServices.DllImport("user32.dll")>
    Private Shared Function GetWindowLong(
        hWnd As IntPtr,
        nIndex As Integer) As Integer
    End Function
    <Runtime.InteropServices.DllImport("user32.dll")>
    Private Shared Function SetWindowLong(
        hWnd As IntPtr,
        nIndex As Integer,
        dwNewLong As Integer) As Integer
    End Function

    ' === 2026/4/19, 用來改變windows計時器精度 ===
    <Runtime.InteropServices.DllImport("winmm.dll", EntryPoint:="timeBeginPeriod", SetLastError:=True)>
    Private Shared Function TimeBeginPeriod(
        ByVal uPeriod As Integer) As Integer
    End Function
    <Runtime.InteropServices.DllImport("winmm.dll", EntryPoint:="timeEndPeriod", SetLastError:=True)>
    Private Shared Function TimeEndPeriod(
        ByVal uPeriod As Integer) As Integer
    End Function

    ' === 2026/07/03 by Simon/Claude Fable 5: EULA 自動關閉改用 SetWinEventHook 事件驅動 (取代純輪詢的 SW_HIDE 競速) ===
    '   OUTOFCONTEXT 不需 DLL injection，事件經由「安裝 hook 那條執行緒」的訊息佇列送達 (該執行緒必須有 message pump)
    Private Delegate Sub WinEventDelegate(hWinEventHook As IntPtr, eventType As UInteger, hwnd As IntPtr, idObject As Integer, idChild As Integer, dwEventThread As UInteger, dwmsEventTime As UInteger)
    <Runtime.InteropServices.DllImport("user32.dll")>
    Private Shared Function SetWinEventHook(
        ByVal eventMin As UInteger,
        ByVal eventMax As UInteger,
        ByVal hmodWinEventProc As IntPtr,
        ByVal lpfnWinEventProc As WinEventDelegate,
        ByVal idProcess As UInteger,
        ByVal idThread As UInteger,
        ByVal dwFlags As UInteger) As IntPtr
    End Function
    <Runtime.InteropServices.DllImport("user32.dll")>
    Private Shared Function UnhookWinEvent(
        ByVal hWinEventHook As IntPtr) As Boolean
    End Function
    <Runtime.InteropServices.DllImport("user32.dll", CharSet:=Runtime.InteropServices.CharSet.Auto)>
    Private Shared Function GetClassName(
        ByVal hWnd As IntPtr,
        ByVal lpClassName As System.Text.StringBuilder,
        ByVal nMaxCount As Integer) As Integer
    End Function
    ' === 2026/07/03 by Simon/Claude Fable 5: 專職 hook 執行緒的 message pump 用 ===
    '   OUTOFCONTEXT 事件送到「安裝 hook 那條執行緒」的訊息佇列；UI 執行緒被 New RDOSession 卡住時
    '   無法 pump，事件永遠送不到 → 必須開專職執行緒自己跑 GetMessage 迴圈
    <Runtime.InteropServices.StructLayout(Runtime.InteropServices.LayoutKind.Sequential)>
    Private Structure NativeMsg
        Public hwnd As IntPtr
        Public message As UInteger
        Public wParam, lParam As IntPtr
        Public time As UInteger
        Public ptX, ptY As Integer
    End Structure
    <Runtime.InteropServices.DllImport("user32.dll")>
    Private Shared Function GetMessage(
        ByRef lpMsg As NativeMsg,
        ByVal hWnd As IntPtr,
        ByVal wMsgFilterMin As UInteger,
        ByVal wMsgFilterMax As UInteger) As Integer
    End Function
    <Runtime.InteropServices.DllImport("user32.dll")>
    Private Shared Function TranslateMessage(
        ByRef lpMsg As NativeMsg) As Boolean
    End Function
    <Runtime.InteropServices.DllImport("user32.dll")>
    Private Shared Function DispatchMessage(
        ByRef lpMsg As NativeMsg) As IntPtr
    End Function
    <Runtime.InteropServices.DllImport("user32.dll")>
    Private Shared Function PostThreadMessage(
        ByVal idThread As UInteger,
        ByVal msg As UInteger,
        ByVal wParam As IntPtr,
        ByVal lParam As IntPtr) As Boolean
    End Function
    <Runtime.InteropServices.DllImport("kernel32.dll")>
    Private Shared Function GetCurrentThreadId() As UInteger
    End Function

    ' ── 常數 ───
    Private Const WM_LBUTTONDOWN As Integer = &H201
    Private Const WM_LBUTTONUP As Integer = &H202
    Private Const SW_HIDE As Integer = 0
    Private Const EVENT_OBJECT_SHOW As UInteger = &H8002    ' 2026/07/03: 視窗變為可見的 WinEvent
    Private Const WINEVENT_OUTOFCONTEXT As UInteger = 0     ' 2026/07/03: 非注入式，事件走訊息佇列非同步送達
    Private Const OBJID_WINDOW As Integer = 0               ' 2026/07/03: 過濾用，只理會視窗本體事件
    Private Const WM_QUIT As UInteger = &H12                ' 2026/07/03: 結束 hook 執行緒的 GetMessage 迴圈用

    ' TreeView 雙緩衝
    Private Const TV_FIRST As Integer = &H1100
    Private Const TVM_SETEXTENDEDSTYLE As Integer = TV_FIRST + 44
    Private Const TVS_EX_DOUBLEBUFFER As Integer = &H4

    ' ListView 雙緩衝
    Private Const LVM_SETEXTENDEDLISTVIEWSTYLE As Integer = &H1036
    Private Const LVS_EX_DOUBLEBUFFER As Integer = &H10000
    Private Const SWP_NOZORDER As Integer = &H4             ' debugForm resize用
    Private Const SWP_NOACTIVATE As Integer = &H10          ' debugForm resize用
    Private Const SWP_NOSIZE As Integer = &H1               ' 2026/06/19 by Simon/Claude: 只搬位置不改尺寸，拖曳跟隨用
    Private Const SWP_NOREDRAW As Integer = &H8             ' debugForm resize 時閃爍, 改手動redraw
    Private Const RDW_ALLCHILDREN As Integer = &H80         ' debugForm resize 時閃爍, 改手動redraw
    Private Const RDW_INVALIDATE As Integer = &H1           ' debugForm resize 時閃爍, 改手動redraw
    Private Const RDW_UPDATENOW As Integer = &H100          ' debugForm resize 時閃爍, 改手動redraw
    Private Const RDW_ERASE As Integer = &H4                ' 2026/03/28 by Gemini: 補上缺失定義

    ' ↓ 新增 (2026-03-20) ListView1 進入資料夾用
    Private Const WM_SETREDRAW As Integer = &HB             ' 2026/3/26 by Gemini
    Private Const WM_SIZE As Integer = &H5                  ' 視窗尺寸變更訊息, 2026/5/7 by Claude
    Private Const WM_ENTERSIZEMOVE As Integer = &H231       ' 2026/06/19 by Simon/Claude: 進入拖曳 size/move modal loop
    Private Const WM_EXITSIZEMOVE As Integer = &H232        ' 2026/06/19 by Simon/Claude: 離開拖曳 size/move modal loop
    Private Const SIZE_MAXIMIZED As Integer = 2             ' WM_SIZE wParam: 最大化
    Private Const SIZE_RESTORED As Integer = 0              ' WM_SIZE wParam: 還原
#End Region

#Region "■ 99 舊版備用 (勿刪)"
    ' 2026/7/1 by simon, 所有RDO都已切換至獨立session的 _rdo2, 不再沿用 Outlook MAPI session, 讓原有的 _rdo 完全退役
    Private _rdo As Redemption.RDOSession = Nothing         ' _rdoSession 就等同是outlook.namespace 的意思, 就是Redemption的MAPI session
    Private Sub InitRdoSession()
        Try
            ' 3. 初始化 Redemption Session (目前停用，保留開發記錄)
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
    Private Function GetSubtreeToListL3_Rdo(rootFolder As Redemption.RDOFolder, includeSubF As Boolean) As List(Of Redemption.RDOFolder)
        ' --------------------------------------------------------------
        ' 2026/3/24 by Gemini: GetSubtreeToListL3_Rdo
        ' 目的: 專門提供給 RDO 平行路徑使用，回傳 List(Of Redemption.RDOFolder)
        ' 說明: 因為 Redemption 是 free-threaded，可以用 Parallel.ForEach 安全平行展開子樹
        ' 2026/6/27 退役，全部轉由整合完成的GetSubtreeListRdo單一入口
        ' --------------------------------------------------------------
        'Dim rootName As String = rootFolder.Name
        '_dbg("    ├ 開始", rootName)
        'Dim sw As Stopwatch = Stopwatch.StartNew()  ' by Claude Sonnet 4.6, 2026/06/07

        'Dim resultBag As New ConcurrentBag(Of Redemption.RDOFolder)
        'resultBag.Add(rootFolder)
        'If Not includeSubF Then
        '    sw.Stop()
        '    _dbg("    ├ 結束", $"{rootName} (RDO-Single) | {sw.ElapsedMilliseconds}ms") ' by Gemini, 2026/04/10
        '    Return resultBag.ToList()
        'End If

        '' 使用兩層佇列作層級遍歷，每層用 Parallel.ForEach 探索
        'Dim currentLayer As New ConcurrentQueue(Of Redemption.RDOFolder)
        'currentLayer.Enqueue(rootFolder)
        'Do
        '    Dim layerList = currentLayer.ToList()
        '    If layerList.Count = 0 Then Exit Do

        '    ' 清空 queue 準備裝下一層的資料夾
        '    Do While currentLayer.TryDequeue(Nothing) : Loop

        '    ' 平行處理當前層的資料夾，將它們的子資料夾加進 queue 與結果中
        '    Parallel.ForEach(layerList, Sub(current)
        '                                    Try
        '                                        For Each subFolder As Redemption.RDOFolder In current.Folders
        '                                            resultBag.Add(subFolder)
        '                                            currentLayer.Enqueue(subFolder)
        '                                        Next
        '                                    Catch ex As System.Exception
        '                                        _dbg("    ├ 錯誤", current.Name & " - " & ex.Message) ' by Gemini, 2026/04/10
        '                                    End Try
        '                                End Sub)
        'Loop

        'sw.Stop()
        '_dbg("    ├ 結束", $"{rootName} (RDO-Parallel BFS) | 資料夾總計: {resultBag.Count} | {sw.ElapsedMilliseconds}ms") ' by Gemini, 2026/04/10
        'Return resultBag.ToList()
    End Function
    Private Async Function RdoPreloadAttach_1(sourceList As List(Of MailItemInfo), progress As IProgress(Of ProgressReport), cToken As CancellationToken) As Task
        ' =================================================================
        ' by Gemini, 2026/04/05: Layer2.5 快取代理層 - 批次預熱附件檔名快取
        '   利用 Redemption (RDO) Free-Threaded 安全的特性，
        '   在進入 Layer2 迴圈前平行提早把附件檔名讀進 _cacheAttFilename。
        '   完全不更改原有的迴圈運作邏輯，以預讀取的型態塞資料進快取來大幅壓縮等待時間。
        ' =================================================================
        If _rdo Is Nothing OrElse sourceList.Count = 0 Then Return

        _dbg("開始", $"RDO預載Parallel.ForEach {sourceList.Count} 筆")
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
                                                    If Not _cacheAttFilename.ContainsKey(mail.EntryID) Then
                                                        Dim rdoMsg As Redemption.RDOMail = Nothing
                                                        Try
                                                            rdoMsg = TryCast(_rdo.GetMessageFromID(mail.EntryID), Redemption.RDOMail)
                                                            If rdoMsg IsNot Nothing Then
                                                                Dim list As New List(Of String)(512)
                                                                For i As Integer = 1 To rdoMsg.Attachments.Count    ' COM 的 index 從 1 開始而不是0
                                                                    list.Add(rdoMsg.Attachments.Item(i).FileName)
                                                                Next
                                                                _cacheAttFilename.TryAdd(mail.EntryID, list)
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
                                                                                                  .Message = $"Phase 2 (RDO預載Parallel.ForEach): {curProcessed} / {total} ({eta.Speed:F0} 封/秒{eta.EtaString})"})
                                                        swThrottle.Restart()
                                                    End If
                                                End Sub)
                           Catch ex As OperationCanceledException
                               ' cToken 取消時 Parallel.ForEach 拋出，正常中斷，不需處理
                           End Try
                       End Sub, cToken)
        _dbg(" ├ 結束", $"RDO預載Parallel.ForEach完成，處理共 {processed} 筆") ' by Gemini, 2026/04/10: 調整縮排層級為 Level 1
    End Function
    Private Async Function RdoPreloadAttach_2(sourceList As List(Of MailItemInfo), progress As IProgress(Of ProgressReport), cToken As CancellationToken) As Task
        ' ==============================================================
        ' by AntiGravity, 2026/04/07: 實驗性質
        ' - 使用 Task.WhenAll + SemaphoreSlim，試圖推高 SSD I/O 並發度
        ' ==============================================================
        If _rdo Is Nothing OrElse sourceList.Count = 0 Then Return

        _dbg(" ├ 開始", $"RDO預載Task.WhenAll {sourceList.Count} 筆") ' by Gemini, 2026/04/10: 調整縮排層級為 Level 1
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
                                       If Not _cacheAttFilename.ContainsKey(mail.EntryID) Then
                                           Dim rdoMsg As Redemption.RDOMail = Nothing
                                           Try
                                               rdoMsg = TryCast(_rdo.GetMessageFromID(mail.EntryID), Redemption.RDOMail)
                                               If rdoMsg IsNot Nothing Then
                                                   Dim list As New List(Of String)(512)
                                                   For i As Integer = 1 To rdoMsg.Attachments.Count
                                                       list.Add(rdoMsg.Attachments.Item(i).FileName)
                                                   Next
                                                   _cacheAttFilename.TryAdd(mail.EntryID, list)
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
                                                                                     .Message = $"Phase 2 (RDO預載Task.WhenAll): {curProcessed} / {total} ({eta.Speed:F0} 封/秒{eta.EtaString})"})
                                           swThrottle.Restart()
                                       End If
                                   Finally
                                       throttler.Release()
                                   End Try
                               End Function, cToken))
        Next

        If tasks.Count > 0 Then Await Task.WhenAll(tasks)
        _dbg(" ├ 結束", $"RDO預載=Task.WhenAll完成，處理共 {processed} 筆") ' by Gemini, 2026/04/10
    End Function
    Private Async Function RdoPreloadAttach_3(sourceList As List(Of MailItemInfo), progress As IProgress(Of ProgressReport), cToken As CancellationToken) As Task
        ' =================================================================
        ' 2026/06/19 by Simon/Claude: Layer2.5 多PST獨立Session平行預載 (實驗版_3)
        '   獨立 RDOSession (自有 MAPI session) 才是 Redemption 真 free-threaded 的前提。
        '   ★ 實測結論: (a) 同一 PST 多 session 無加速(PST provider 對同檔序列化);
        '             (b) 獨立冷 session 每封成本約為 Outlook 熱 session 的 3 倍。
        '   故設計為「每個 PST 一條獨立 session、組內循序、組間平行」,
        '   加速僅來自「跨多個 PST 同時讀取」→ 適用情境是同時選取大量 PST 的整庫掃描,
        '   少數 PST 時請改用 _1/_2 (共用熱 session 反而較快)。
        '   store 物件開啟一次重複使用(store.GetMessageFromID), 避免每封信重新定位 store 的成本。
        ' =================================================================
        If _rdo Is Nothing OrElse sourceList.Count = 0 Then Return   ' _rdo 僅作「RDO 是否可用」的偵測旗標

        _dbg("開始", $"RDO預載Cross.PST {sourceList.Count} 筆")
        Dim swTotal As Stopwatch = Stopwatch.StartNew()
        Dim swThrottle As Stopwatch = Stopwatch.StartNew()
        Dim processed As Integer = 0
        Dim resolveFail As Integer = 0      ' 暫時探針: 解 EntryID 的失敗數, 確認=0 後可移除
        Dim total As Integer = sourceList.Count

        ' 依 PST(store 顯示名) 分組: 每組一條獨立 session, 組內循序, 組間平行才是真正加速來源
        Dim groups = sourceList.GroupBy(Function(m) GetStoreNameFromPath(m.FolderPath)).ToList()
        _dbg(" ├ 分組", $"涵蓋 {groups.Count} 個 PST → 開 {groups.Count} 條平行 session")

        Dim tasks As New List(Of Task)(groups.Count)
        For Each grp In groups
            Dim storeName As String = grp.Key
            Dim items = grp.ToList()
            tasks.Add(Task.Run(Sub()
                                   Dim sess As Redemption.RDOSession = Nothing
                                   Try
                                       sess = New Redemption.RDOSession()
                                       sess.Logon(_rdo.ProfileName, "", False, True)   ' (ProfileName, Pwd, ShowDialog, NewSession): 不沿用 Outlook session
                                       ' 取得該 PST 已開啟的 RDOStore 並重複使用(避免每封信重開 store 的高昂成本)
                                       Dim store As Redemption.RDOStore = Nothing
                                       For i As Integer = 1 To sess.Stores.Count
                                           If sess.Stores.Item(i).Name = storeName Then store = sess.Stores.Item(i) : Exit For
                                       Next
                                       If store Is Nothing Then
                                           _dbg(" ├ 略過", $"獨立 session 找不到 store [{storeName}]，跳過該組 {items.Count} 筆")
                                           Interlocked.Add(resolveFail, items.Count)
                                           Return
                                       End If

                                       For Each mail As MailItemInfo In items
                                           cToken.ThrowIfCancellationRequested()
                                           If Not _cacheAttFilename.ContainsKey(mail.EntryID) Then
                                               Dim rdoMsg As Redemption.RDOMail = Nothing
                                               Try
                                                   rdoMsg = TryCast(store.GetMessageFromID(mail.EntryID), Redemption.RDOMail)   ' ★ 用已開啟的 store, 不每封重開
                                                   If rdoMsg IsNot Nothing Then
                                                       Dim list As New List(Of String)(512)
                                                       For i As Integer = 1 To rdoMsg.Attachments.Count    ' COM 的 index 從 1 開始而不是0
                                                           list.Add(rdoMsg.Attachments.Item(i).FileName)
                                                       Next
                                                       _cacheAttFilename.TryAdd(mail.EntryID, list)
                                                   Else
                                                       Interlocked.Increment(resolveFail)
                                                   End If
                                               Catch
                                                   Interlocked.Increment(resolveFail)
                                               Finally
                                                   If rdoMsg IsNot Nothing Then TryMarshalRelease(rdoMsg)
                                               End Try
                                           End If

                                           Dim curProcessed As Integer = Interlocked.Increment(processed)
                                           If swThrottle.ElapsedMilliseconds >= ThrottleFreq.Hii OrElse curProcessed = total Then
                                               Dim eta = CalculateSpeedAndETA(total, curProcessed, swTotal.Elapsed.TotalSeconds)
                                               progress?.Report(New ProgressReport With {.CurrentCount = curProcessed, .TotalCount = total,
                                                                                         .Message = $"Phase 2 (RDO預載Cross.PST): {curProcessed} / {total} ({eta.Speed:F0} 封/秒{eta.EtaString})"})
                                               swThrottle.Restart()
                                           End If
                                       Next
                                   Catch ex As OperationCanceledException
                                       ' cToken 取消, 正常中斷
                                   Catch ex As System.Exception
                                       _dbg(" ├ 失敗", $"PST [{storeName}] 組例外: {ex.GetBaseException().Message}")
                                   Finally
                                       If sess IsNot Nothing Then
                                           Try : sess.Logoff() : Catch : End Try
                                           TryMarshalRelease(sess)
                                       End If
                                   End Try
                               End Sub, cToken))
        Next

        If tasks.Count > 0 Then Await Task.WhenAll(tasks)
        _dbg(" ├ 結束", $"RDO預載Cross.PST完成，處理共 {processed} 筆，resolve 失敗 {resolveFail} 筆")
    End Function
    Private Function objFolder2odoFolder(objFolder As Folder) As Redemption.RDOFolder
        If _rdo Is Nothing OrElse objFolder Is Nothing Then Return Nothing
        Return _rdo.GetFolderFromID(objFolder.EntryID, objFolder.StoreID)
    End Function
    Private Function rdoFolder2objFolder(rdoFolder As Redemption.RDOFolder) As Folder
        If rdoFolder Is Nothing Then Return Nothing
        Return _olNS.GetFolderFromID(rdoFolder.EntryID, rdoFolder.StoreID)
    End Function

    Private Function GetSortedSubFolderIDs(fPath As String, parentEid As String, parentSid As String, selfKnownToDb As Boolean) As List(Of (path As String, eid As String, sid As String))
        ' ==========================================
        ' 2026/07/02 by Simon/Claude [骨架整合]: 退役 — 唯一呼叫端 BuildBfsFolderTree 已改吃 GetSubtree 骨架(記憶體/DB LIKE/RDO批次/OOM 四層)，
        '   本函數目前無任何呼叫端，暫留作參考；連同 selfKnownToDb 冷啟動特判一併停用，穩定一輪後可刪。
        ' ==========================================
        ' 2026/06/29 by Simon/Claude [Option A1]: GetSortedSubFolders 的「免物化」對稱孿生 —
        '   回直屬子夾的 (path,eid,sid) 純資料 tuple，給 Tab1 暖重啟 id-tuple BFS 用，全程零 COM Folder 物化。
        '   ② DB 優先(LazyGetOrderedSubFolderIDs，已含 entry_id IS NOT NULL [/is_mail=1] 過濾 + 英文優先排序)；
        '   ② 回 Nothing(DB 無此節點子夾) → ③ 退 GetFolderById 物化 parent 走 COM 列舉(保留身分證註冊副作用)。
        '   注意: 本函數不寫 _cacheFolderTree(那是物化版 List(Of Folder) 快取)；BFS 走 ② DB 已足夠快(~120ms/全樹)。
        ' 2026/06/30 by Simon/Claude [A1 修正]: 移除 ③ COM 物化退路。元兇釘死 —
        '   LazyGetOrderedSubFolderIDs 對「葉節點(無子夾)」與「DB缺漏」都回 Nothing,無法區分;
        '   原 ③ 對每個葉節點(佔樹大半)誤觸發 GetFolderById 物化(~2.8ms/夾),正是 S1 沒降的真因。
        '   暖重啟 DB 健康(R1 探針 46 夾全相符),葉節點回 Nothing 屬正常 → 回空清單。
        '   DB 真缺漏的半殘情境按 Q2 既定方針交 F5 強制刷新重建,不在此 graceful。
        ' 2026/07/01 by Simon/Claude: 重新加回 ③ COM 物化退路，改為有條件觸發，修復冷啟動(全新DB/無任何記錄)子樹展不開的 regression —
        '   06/30 拿掉 ③ 的前提是「暖重啟、DB健康」，但完全空白的 DB 下，LazyGetOrderedSubFolderIDs 對「真葉節點」與「DB根本沒掃過」同樣回 Nothing，
        '   導致從 root 往下第一層就展不開，整棵子樹統計全部停在 0。
        '   修法: 呼叫端 BuildBfsFolderTree 多傳入 selfKnownToDb(此節點自己在記憶體/DB 是否已有記錄)；
        '   只有 selfKnownToDb=False(真正未知節點)才觸發 ③；
        '   selfKnownToDb=True 但查無子項，維持 06/30 判斷、信任為真葉節點不物化。效能與正確性兩者兼顧。
        ' ==========================================
        Dim result As New List(Of (path As String, eid As String, sid As String))(512)

        ' ② DB 直讀(免物化優先路徑)
        If _dbCache IsNot Nothing Then
            Dim dbIDs = LazyGetOrderedSubFolderIDs(fPath, _showAllFolders)
            If dbIDs IsNot Nothing Then
                For Each row In dbIDs : result.Add((row.path, row.eid, row.sid)) : Next
                Return result
            End If
        End If

        ' ③ COM 物化退路：僅在此節點本身「DB 也沒有記錄」(真正未知節點)時才觸發，已知的真葉節點(selfKnownToDb=True 但查無子項)不會誤觸發，維持 06/30 的效能優化
        If Not selfKnownToDb Then
            Dim pFolder As Folder = GetFolderById(parentEid, parentSid)
            If pFolder IsNot Nothing Then
                For Each subF In GetSortedSubFolders(pFolder, fPath, skipCache:=True)
                    result.Add((fPath & "\" & subF.Name, subF.EntryID, subF.StoreID))
                Next
            End If
        End If

        Return result
    End Function
    Private Function GetMailCountRecursiveLegacy(folder As Outlook.Folder) As Integer
        _dbg("開始", folder.Name)
        Dim value As Integer
        If _cacheMailCountAll.TryGetValue(folder, value) Then Return value ' 檢查快取中是否已存在值, 若有則直接返回
        ' 改成先用 Parallel.ForEach 遍歷子文件夾並且並行處理
        Dim totalMailCount As Integer = 0
        Dim countingBag As New ConcurrentBag(Of Integer)()
        Try
            ' 5/21記錄: 模仿GetFolderSizeLegacy那一句超快速的LINQ, 但測試結果沒有現在這個快, 所以決定保留這個
            ' 2026/3/20, 重寫了底層GetMailCountAll() 但是不知為何效能還是比不過現在下面這個遞迴版本??
            ' 原因: 原版遞迴只走一遍 COM 資料夾樹，新版走了兩遍COM:
            ' 第一遍: GetSubtree()    → BFS 遍歷，存取每個 folder.Folders
            ' 第二遍: For Each allFolders   → GetMailCountOOM() 再讀每個資料夾一次
            ' 2026/3/22, 導入Redemption, 應該可以刪掉這裡了? 還是讓Redemption 變成on-demand, 需要才啟動?
            'Parallel.ForEach(folder.Folders.Cast(Of Outlook.Folder),' 取得子資料夾的郵件數量並添加到 ConcurrentBag 中
            '                 Sub(subFolder As Outlook.Folder)
            '                     countingBag.Add(GetMailCountRecursive(subFolder))
            '                 End Sub)
            'totalMailCount = countingBag.Sum() ' 累加所有子資料夾的郵件數量
            ''' 最後再獲取選取文件夾自身的郵件數量 (改用MAPI table 的PR_CONTENT_COUNT屬性來getmailcount)
            ''Const PR_CONTENT_COUNT As String = "http://schemas.microsoft.com/mapi/proptag/0x36020003"
            ''totalMailCount += folder.PropertyAccessor.GetProperty(PR_CONTENT_COUNT)
            totalMailCount += GetMailCountOOM(folder)  ' 單一目錄的mail count改成重寫的統一底層函數, 2026/3/20
            _cacheMailCountAll.TryAdd(folder, totalMailCount) ' 第一次計算後就存入快取
        Catch
        End Try
        Return totalMailCount

    End Function
    Private Function GetMailSizeL3(item As Object) As Long
        ' --------------------------------------------------------------
        ' GetMailSizeL3: 讀取單封郵件的大小 (bytes)，供 GetFolderSizeOOM fallback 路徑呼叫
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
        ' 注意: 此函數接受 Object 型別參數，是因為 GetFolderSizeOOM 的 fallback 路徑
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
                Dim parentFolder As Outlook.Folder = TryCast(mail.Parent, Outlook.Folder)
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
            ' by Gemini, 2026/03/29: 移除 TypeOf 判斷，CLng() 可自動處理 Long/Integer 轉型，若屬性不存在或回傳 Nothing/DBNull，CLng 會拋例外進入 Catch
            Return CLng(mail.PropertyAccessor.GetProperty(PR_MESSAGE_SIZE_EXTENDED))
        Catch ex As System.Exception
            If _iLikeNoisy Then _dbg("    ├ ① PR_MESSAGE_SIZE_EXTENDED失敗", ex.Message) ' by Gemini, 2026/04/10
        End Try

        ' ② MAPI: PR_MESSAGE_SIZE (0x0E080003, PT_LONG) — 32-bit，超大郵件理論上溢位
        Try
            Return CLng(mail.PropertyAccessor.GetProperty(PR_MESSAGE_SIZE))             ' by Gemini, 2026/03/29: 同上，移除 TypeOf 判斷
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
    Private Async Function GetFolderSizeLegacy(folder As Outlook.Folder) As Task(Of Long)
        ' ==============================================================
        ' === GetFolderSizeLegacy — 修正版 (移除 Task.Run 包 COM) ===
        ' ==============================================================
        '
        ' 原版問題: Task.Run(Function() folder.Items.Cast(Of Object)().Sum(Function(st) st.Size))
        '          在 thread pool 執行緒上操作 Outlook COM 物件，違反 STA 規定, 在特定情況 (COM interop 敏感時機) 會造成 crash 或傳回錯誤結果
        '
        ' 修正做法: GetTable + PR_MESSAGE_SIZE 在 UI 執行緒循序讀取
        '           GetTable 回傳 MAPI binary table (低層讀取)
        '          一次只讀一個 Row，每個 Row 用後立即 ReleaseComObject，避免 RCW 累積
        '          每 100 筆 Yield 一次讓 UI 保持回應
        '          速度接近原版 LINQ (實測差距在誤差範圍內) ，但 STA 安全
        '
        ' 此函數仍為 Lazy (不主動觸發) :
        '   由 ListView1_ColumnClick 或右鍵選單「Show This Folder Size」觸發
        '   結果存入 folderSizeCache，BuildLv1Item 下次組裝時自動顯示
        ' ==============================================================
        _dbg("開始", folder.Name)
        Dim value As Long   ' 快取命中直接回傳
        If _cacheFolderSize.TryGetValue(folder, value) Then Return value
        '' 已知有問題的資料夾走舊路徑 (不明 COM 例外物件，GetTable 也可能出問題)
        'Dim exceptList As String() = {"Inbox_2000~2018", "Facebook"}
        'If exceptList.Contains(folder.Name) Then Return GetFolderSizeOld(folder)
        Dim table As Outlook.Table = Nothing
        Try
            ' GetTable + PR_MESSAGE_SIZE (0x0E080003) :
            ' PR_MESSAGE_SIZE_EXTENDED (0x0E080014, PT_I8) — PST 本地端的內建彙總屬性
            ' 只讀 Size 欄，不載入其他 MAPI 屬性，減少記憶體與 COM 開銷
            table = folder.GetTable()
            table.Columns.RemoveAll()
            table.Columns.Add(PR_MESSAGE_SIZE_EXTENDED)
            Dim totalSize As Long = 0
            Dim rowCount As Integer = 0
            Do While Not table.EndOfTable
                Dim row As Outlook.Row = table.GetNextRow()
                totalSize += SafeGet(Of Long)(row, PR_MESSAGE_SIZE_EXTENDED, 0L)
                TryMarshalRelease(row)
                rowCount += 1
                If rowCount Mod 100 = 0 Then Await Task.Yield()  ' 每 100 筆統計就讓 UI 回應一次
            Loop
            _cacheFolderSize.TryAdd(folder, totalSize)
            Return totalSize
        Catch ex As OverflowException
            _dbg("Error: GetFolderSizeLegacy overflow", folder.Name)
            Return -1
        Catch ex As System.Exception
            _dbg("Error: GetFolderSizeLegacy", folder.Name & " - " & ex.Message)
            Return -1
        Finally
            TryMarshalRelease(table)
        End Try

    End Function
    Private Async Function GetTotalFolderCountAsync(folder As Outlook.Folder) As Task(Of Integer)
        _dbg("開始", folder.Name)
        Dim value As Integer
        Dim fPath As String = folder.FolderPath
        If _cacheFolderCountAll.TryGetValue(fPath, value) Then Return value     ' 檢查快取中是否已存在值, 若有則直接返回
        Dim totalSubCount As Integer = GetFolderCountOOM(folder, fPath:=fPath)  ' 初始值為點選資料夾的子資料夾數量
        ' 5/21測試記錄: 這裡使用ConcurrentBag跟使用results.sum應該要比較快, 但不知為何實測結果都比GetTotalFolderCount_Old()還慢了5%, 這個函數先保留不清除
        ' 5/21最後決定: 二個函數快慢互有變化, 但GetTotalFolderCountAsync()的穩定性較好, 比New()的標準差來得小, 所以決定使用這個
        ' 使用 Parallel.ForEach 進行多線程遞迴計算subfolder數量
        Dim countingBag As New ConcurrentBag(Of Task(Of Integer))()             ' 使用 ConcurrentBag 來安全地收集每個子資料夾的數量
        Parallel.ForEach(folder.Folders.Cast(Of Outlook.Folder)(),
                         Sub(subFolder As Outlook.Folder)
                             'countingBag.Add(GetTotalFolderCountAsync(subFolder))
                             'countingBag.Add(GetFolderCountAllOOM(subFolder))
                         End Sub)
        Dim results = Await Task.WhenAll(countingBag)   ' 等待所有平行出去收集的數量都確定回來了
        totalSubCount += results.Sum()                  ' 再將回傳的各個子資料夾的數量加總
        _cacheFolderCountAll.TryAdd(fPath, totalSubCount)
        ' ✅ 2026-03-16 移除多餘的 Try/Catch: ConcurrentDictionary.TryAdd 本身不拋例外 (只回傳 True/False)
        ' 原本是從 .Add() 時代留下的防護，改 TryAdd 後應一併移除
        Return totalSubCount

    End Function

    Private Function CalculateSimilarity(strA As String, strB As String) As Double
        ' by Gemini, 2026/04/04: Issue 1 移除 _dbg (Tab5 高頻呼叫，N封×2個函數=2N行輸出) 
        ' 計算編輯距離
        Dim editDistance As Integer = LevenshteinDistance(strA, strB)

        ' 將編輯距離歸一化為範圍在 0 到 1 之間的值
        Dim maxLength As Integer = Math.Max(strA.Length, strB.Length)
        Dim similarity As Double = 1 - CDbl(editDistance) / maxLength
        Return similarity

    End Function
    Private Function LevenshteinDistance(strA As String, strB As String) As Integer
        ' by Gemini, 2026/04/04: Issue 1 移除 _dbg (Tab5 高頻呼叫，同上) 
        ' 計算 Levenshtein 編輯距離的輔助函數
        Dim lenA As Integer = strA.Length
        Dim lenB As Integer = strB.Length
        Dim distance(lenA, lenB) As Integer
        For i As Integer = 0 To lenA : distance(i, 0) = i : Next
        For j As Integer = 0 To lenB : distance(0, j) = j : Next
        For j As Integer = 1 To lenB
            For i As Integer = 1 To lenA
                '' 改前 (5行)
                'If strA(i - 1) = strB(j - 1) Then
                '    distance(i, j) = distance(i - 1, j - 1)
                'Else
                '    distance(i, j) = Math.Min(Math.Min(distance(i - 1, j) + 1,
                '                                       distance(i, j - 1) + 1), distance(i - 1, j - 1) + 1)
                'End If

                ' 改後 (1行)
                distance(i, j) = If(strA(i - 1) = strB(j - 1),
                distance(i - 1, j - 1), Math.Min(Math.Min(distance(i - 1, j) + 1,
                                                          distance(i, j - 1) + 1), distance(i - 1, j - 1) + 1))
            Next
        Next
        Return distance(lenA, lenB)
    End Function
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
        '   1. Tab1: mailCountCache + folderCountCache (GetMailCountAllOOM / GetTotalFolderCountAsync)
        '   2. Tab2: yearCountsCache (GetYearCountsForFolderAsync)
        '   3. Tab3: _cacheAttMailList (CheckTab3CacheOrRescan)
        '
        ' 2026-03-16 B4 新增，由 PrewarmAllCachesAsync 重構整合，改名為 CacheSniffer
        '       只要偵測到正在進行 AfterSelect 或是正在跑複雜統計，就自動閉嘴等閒下來再繼續
        ' ===============================================================================

        'If _pstStoreList Is Nothing OrElse _pstStoreList.Count = 0 Then Return
        Await Task.Delay(10000, cToken)      ' 等待 10 秒: 確保 Form1_Load 完全結束、UI 呈現完畢，再開始佔用 Outlook COM

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
        '        ' ── Tab2: yearCountsCache ─────────────────────────────────
        '        ' ── Tab3: _cacheAttMailList ────────────────────────────────
        '        ' ── 固定 1 秒間隔: 讓 Outlook 保持回應能力 ───────────────
        '        _dbg($"CacheSniffer: [{processed}] {folder.Name} 完成，等 1 秒")
        '        Await Task.Delay(1000, cToken)
        '        Await Task.Yield()
        '        ' ── 把直屬子資料夾加入佇列 (廣度優先) ────────────────────
        '        ' GetSortedSubFolders 有 folderTreeCache，不重打 COM
        '        Try
        '            For Each subFolder As Outlook.Folder In GetSortedSubFolders(folder):queue.Enqueue(subFolder):Next
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
    Private Async Function ForceTvRefresh_old(tv As SimTree) As Task
        ' ── SimTree F5 強制刷新 ──────────────────────────────────────────────
        ' 職責: 不讀任何快取，重新從 Outlook COM 讀取整棵資料夾樹並更新 _cacheFolderTree
        '       ① 記錄目前展開路徑 + 選取路徑
        '       ② 清 _cacheFolderTree (確保 LoadSubFolderToTreeView 重讀 COM)
        '       ③ Nodes.Clear + LoadStoreToTreeView (重建 root 層)
        '       ④ 逐層 node.Expand() 重建已展開路徑 (觸發 LoadSubFolderToTreeView)
        '       ⑤ 還原選取，透過 FireAfterSelect 觸發正常 AfterSelect 流程更新 ListView
        '
        ' 2026/05/13 by Claude Sonnet 4.6
        ' 2026/05/17 by Simon/Claude: ⑤ 改回 FireAfterSelect，解決 ListView 未更新的問題
        '   原本直接呼叫 CollectTab1FolderInfo + RenderLv1 的方式繞過了 SimTree 標準流程，
        '   導致 AfterSelect 沒有被觸發，ListView 顯示內容不對應選取的資料夾。
        ' 2026/05/25 by Simon/Claude: 再度重構使用呼叫simTree內部方法
        ' ─────────────────────────────────────────────────────────────────────
        _dbg("開始", tv.Name)
        If _pstStoreList Is Nothing OrElse _pstStoreList.Count = 0 Then Return

        ' ① 記錄展開路徑與選取路徑
        Dim expandedPaths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        CollectExpandedPaths(tv.Nodes, expandedPaths)

        Dim selectedPaths As New List(Of String)(32)
        For Each node As TreeNode In tv.SelectedNodes
            Dim f As Outlook.Folder = TryCast(node.Tag, Outlook.Folder)
            If f IsNot Nothing Then selectedPaths.Add(SafeGetPath(f))
        Next

        PgrsBar1.Text = $"F5: 重整 {tv.Name}..." : PgrsBar2.Text = ""
        _isUserBusy = True : Cursor = Cursors.WaitCursor

        Try
            ' ② 清 _cacheFolderTree（確保 GetSortedSubFolders 重讀 COM，不用舊快取）
            _cacheFolderTree.Clear()

            ' ③ 重建 root 層
            tv.ClearSelectedNodes()
            tv.Nodes.Clear()
            LoadStoreToTreeView(_pstStoreList, tv)

            ' ④ 逐條路徑重展開（node.Expand() 觸發 BeforeExpand → LoadSubFolderToTreeView → 新鮮 COM）
            For Each path In expandedPaths.OrderBy(Function(p) p.Length)   ' 由淺到深確保父節點先展開
                ReExpandNodeByPath(tv, path)
                Dim unused = TimeBeginPeriod(1)
                Await Task.Delay(1)
                Dim unused1 = TimeEndPeriod(1)
            Next

            ' ⑤ 還原選取
            Dim firstNode As TreeNode = Nothing
            For Each path In selectedPaths
                ' by Gemini 3.5 Flash, 2026/05/21: 改用 tv.GetNodeIn 高效尋路引擎，取代舊有的暴力遞迴 FindNodeByPath
                Dim found As TreeNode = tv.GetNode(path, searchOnlyExpanded:=True)
                If found IsNot Nothing Then
                    tv.AddSelectedNode(found)
                    If firstNode Is Nothing Then firstNode = found
                End If
            Next

            If firstNode IsNot Nothing Then
                firstNode.EnsureVisible()
                ' 2026/05/17 by Simon/Claude:
                tv.FireAfterSelect(firstNode)
                ' 改回 FireAfterSelect，讓 SimTree1_AfterSelect 自行處理統計與 RenderLv1，這才是 SimTree 的標準觸發流程。
                ' 原本直接呼叫 CollectTab1FolderInfo 的方式導致 ListView1 未被正確更新。
            Else
                GotoDefaultInbox(tv)   ' 找不到舊選取時退回預設 Inbox
            End If

            PgrsBar1.Text = $"F5: {tv.Name} 重整完成" : PgrsBar2.Text = ""

        Catch ex As System.Exception
            _dbg("錯誤", ex.Message) : PgrsBar1.Text = $"F5 {tv.Name} 失敗: " & ex.Message
        Finally
            Cursor = Cursors.Default : _isUserBusy = False : _dbg("結束", tv.Name)
        End Try
    End Function
    Private Sub CollectExpandedPaths(nodes As TreeNodeCollection, paths As HashSet(Of String))
        ''' <summary>遞迴收集已展開節點的 FolderPath，供 F5 刷新前記錄狀態用</summary>
        For Each n As TreeNode In nodes
            Dim f As Outlook.Folder = TryCast(n.Tag, Outlook.Folder)
            If f Is Nothing Then Continue For   ' 跳過 ":::" 佔位節點
            If n.IsExpanded Then
                paths.Add(SafeGetPath(f))
                CollectExpandedPaths(n.Nodes, paths)
            End If
        Next
    End Sub
    Private Sub ReExpandNodeByPath(tv As SimTree, fullPath As String)
        ' by Gemini 3.5 Flash, 2026/05/21: 重構以使用底層高效的尋路與展開機制，取代舊的手動逐層循環暴力比對，以防佔用執行緒
        Dim found As TreeNode = Nothing
        If tv.TryGetNode(fullPath, found, searchOnlyExpanded:=False, expandAlongTheWay:=True) Then
            If found IsNot Nothing AndAlso Not found.IsExpanded AndAlso found.Nodes.Count > 0 Then found.Expand()
        End If
    End Sub

    Private Function GetDeDupedNodes(nodes As IEnumerable(Of TreeNode)) As List(Of TreeNode)
        ''' <summary>
        ''' [Layer 1.5 輔助層] 執行父子去重過濾。
        ''' 確保當「父資料夾」及其「子資料夾」同時被選中時，只保留父資料夾以防止重複統計。
        ''' </summary>
        _dbg("開始")
        If nodes Is Nothing Then Return New List(Of TreeNode)

        ' 預分配容量為 64 (by Gemini 3 Flash, 2026/05/04)
        Dim selectedSet As New HashSet(Of TreeNode)(nodes)
        Dim dedupedNodes As New List(Of TreeNode)(64)

        For Each node As TreeNode In nodes
            Dim isDescendantOfSelected As Boolean = False
            Dim ancestor As TreeNode = node.Parent
            While ancestor IsNot Nothing
                ' 若某節點的任一祖先也在選中清單裡，表示該節點已被涵蓋，應跳過
                If selectedSet.Contains(ancestor) Then isDescendantOfSelected = True : Exit While
                ancestor = ancestor.Parent
            End While
            If Not isDescendantOfSelected Then dedupedNodes.Add(node)
        Next
        Return dedupedNodes
    End Function

    ' 2026/06/22 by Simon/Claude Opus 4.8: IRM 保護信隔離夾名稱 (方案 Y: 每顆 PST 各建一個同名夾, 同 store 內搬)
    Private Const QUARANTINE_NAME As String = "_IRM_Protected"
    Private Async Function ScanAndMoveRpmsgRdo() As Task
        '' ============================================================================
        '' 2026/06/22 by Simon/Claude Opus 4.8: 【一次性工具】scan-and-move — 把 message.rpmsg 保護信隔離
        ''   作法: 依 SimTree3 選定節點掃整棵子樹, 命中(任一附件 .rpmsg)就用 RDO 把該信 Move 到
        ''         「同一顆 PST 的 _IRM_Protected 夾」(方案 Y, 同 store 內搬, 避開跨 store 不確定性)。
        ''   為何 scan-and-move 而非餵 EntryID: 搬移後 EntryID 會變, 來回 rebind 脆; 掃描當下手上就有 live
        ''         RDOMail, 就地搬最穩, 且搬前再驗一次 .rpmsg 防呆。全程走 RDO 不會觸發授權 modal。
        ''   ⚠ 破壞性: 信會離開原夾。搬完那些來源夾 + 隔離夾的 SQLite 快照會 stale, 需自行對受影響夾跑 RenewCache。
        ''   ⚠ 完整性: 請先把「所有可能含 rpmsg 的 PST」都選進 SimTree3 再執行, 才能一次搬乾淨。
        '' 2026/6/27 by Simon/Claude Opus 4.8: 原有呼叫GetSubtreeToListL3_RDO()退役, 改成新的GetSubtreeRdoByBatch()
        ''   For Each r In roots
        ''       Try (root)
        ''           rdoRoot → 走子樹 nodes → 釋 rdoRoot → 組 scanEids
        ''           For Each fe In scanEids
        ''               Try (folder)
        ''                   rdoF = Store.GetFolderFromID(fe.eid) → fName → 跳隔離夾
        ''                   [L1464–1520 原 items 掃描/搬移,不動]
        ''               Finally → 釋 rdoF
        ''               End Try
        ''           Next (fe)
        ''       Catch → log
        ''       End Try
        ''   Next (roots)
        '' ============================================================================

        '' ── 0. 確保 RDO 已載入 (改用 _rdo2 獨立 session) ──
        '' 2026/06/27 by Simon/Claude Opus 4.8: _rdo → _rdo2。兩者在 InitRdoSessionWithoutEULA 同生、ReleaseRdoSession 同滅,判 _rdo2 即可。
        'If _rdo2 Is Nothing Then Await InitRdoSessionWithoutEULA()
        'If _rdo2 Is Nothing Then _dbg("RDO隔離", "Redemption (_rdo2) 初始化失敗, 中止") : Return

        '' ── 1. UI 執行緒抽出選定節點 (EntryID, StoreID, 名稱) ──
        'Dim selectedNodes As List(Of TreeNode) = SimTree3.SelectedNodes
        'If selectedNodes Is Nothing OrElse selectedNodes.Count = 0 Then _dbg("RDO隔離", "SimTree3 未選取任何 PST/資料夾") : Return

        '' 2026/06/27 by Simon/Claude: UI 緒先解 _rdo2 store(GetRdoStore 內部用一般 Dictionary,須在 UI 緒呼叫;store 物件本身 free-threaded,可帶進 Task.Run)。
        ''   改抓 path(供走訪當 rootPath + 解 store);sid 不再需要(store-scoped 單參數解夾)。
        'Dim roots As New List(Of (store As Redemption.RDOStore, eid As String, path As String, name As String))(selectedNodes.Count)
        'For Each node As TreeNode In selectedNodes
        '    Dim f As Folder = TryCast(node.Tag, Folder)
        '    If f Is Nothing Then Continue For
        '    Dim p As String = SafeGetPath(f)
        '    Dim st As Redemption.RDOStore = GetRdoStore(p)   ' 記憶化快取 store,不釋放
        '    If st Is Nothing Then _dbg("RDO隔離", $"GetRdoStore 失敗,跳過: {f.Name}") : Continue For
        '    roots.Add((st, f.EntryID, p, f.Name))
        'Next
        'If roots.Count = 0 Then _dbg("RDO隔離", "選取節點皆非有效資料夾") : Return

        '' ── 破壞性動作, 先確認 ──
        'Dim dr As DialogResult = MessageBox.Show(
        '    $"即將掃描 {roots.Count} 個根節點, 把所有 message.rpmsg 保護信搬到各自 PST 的「{QUARANTINE_NAME}」夾。" & vbCrLf & vbCrLf &
        '    "此動作會改變封存結構且不易復原, 確定執行?", "確認隔離搬移", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
        'If dr <> DialogResult.Yes Then _dbg("RDO隔離", "使用者取消") : Return
        '_dbg("RDO隔離 開始", $"掃描 {roots.Count} 個根節點 ...")

        '' ── 2. 背景 scan-and-move (RDO free-threaded) ──
        'Dim moves As New List(Of String)
        'Dim movedCount As Integer = 0, failCount As Integer = 0
        'Dim scanned As Long = 0
        'Dim quarantineCache As New Dictionary(Of String, Redemption.RDOFolder)   ' key: store EntryID → 該 store 的隔離夾

        'Await Task.Run(
        '    Sub()
        '        For Each r In roots
        '            Try
        '                Dim rdoRoot As Redemption.RDOFolder = TryCast(r.store.GetFolderFromID(r.eid), Redemption.RDOFolder)
        '                If rdoRoot Is Nothing Then _dbg("RDO隔離 根節點失敗", $"{r.name} | root 解析失敗") : Continue For

        '                ' 批次走子樹拿 (eid,name,path);失敗退枚舉。rdoRoot 用完即釋。
        '                Dim nodes As List(Of (eid As String, name As String, path As String)) = GetSubtreeRdoBatch(r.store, rdoRoot, r.path)
        '                If nodes Is Nothing Then nodes = GetSubtreeRdoEnum(rdoRoot, r.path)
        '                Dim oRoot As Object = rdoRoot : TryMarshalRelease(oRoot)
        '                If nodes Is Nothing Then _dbg("RDO隔離 根節點失敗", $"{r.name} | 子樹走訪失敗") : Continue For

        '                ' root 自己 + 所有子孫,逐一在 _rdo2 store-scoped 重解夾後掃描
        '                Dim scanEids As New List(Of (eid As String, name As String))(nodes.Count + 1)
        '                scanEids.Add((r.eid, r.name))
        '                For Each nd In nodes : scanEids.Add((nd.eid, nd.name)) : Next

        '                For Each fe In scanEids
        '                    Dim rdoF As Redemption.RDOFolder = Nothing
        '                    Try
        '                        rdoF = TryCast(r.store.GetFolderFromID(fe.eid), Redemption.RDOFolder)
        '                        If rdoF Is Nothing Then Continue For
        '                        Dim fName As String = "" : Try : fName = rdoF.Name : Catch : End Try
        '                        If String.Equals(fName, QUARANTINE_NAME, StringComparison.OrdinalIgnoreCase) Then Continue For   ' 不掃隔離夾自己
        '                        Dim items = Nothing
        '                        Try
        '                            items = rdoF.Items
        '                            Dim cnt As Integer = items.Count
        '                            ' 由後往前 (Move 會把命中信移出本夾, 降序迭代不會影響尚未處理的索引)
        '                            For i As Integer = cnt To 1 Step -1
        '                                Dim m As Redemption.RDOMail = TryCast(items.Item(i), Redemption.RDOMail)
        '                                If m Is Nothing Then Continue For
        '                                Try
        '                                    scanned += 1
        '                                    If scanned Mod 5000 = 0 Then _dbg("RDO隔離 進行中", $"已掃 {scanned}, 已搬 {movedCount} ...")

        '                                    ' 偵測: 任一附件 .rpmsg 即命中 (搬前再驗, 防呆)
        '                                    Dim matched As String = Nothing
        '                                    For k As Integer = 1 To m.Attachments.Count
        '                                        Dim att As Redemption.RDOAttachment = m.Attachments.Item(k)
        '                                        Try
        '                                            Dim afn As String = att.FileName
        '                                            If afn IsNot Nothing AndAlso afn.EndsWith(".rpmsg", StringComparison.OrdinalIgnoreCase) Then matched = afn : Exit For
        '                                        Finally : TryMarshalRelease(att)
        '                                        End Try
        '                                    Next
        '                                    If matched Is Nothing Then Continue For   ' Finally 會釋放 m

        '                                    ' 命中: 先取所屬 store, get-or-create 該 store 的隔離夾
        '                                    Dim st As Redemption.RDOStore = m.Store
        '                                    Dim stKey As String = st.EntryID
        '                                    Dim qf As Redemption.RDOFolder = Nothing
        '                                    If Not quarantineCache.TryGetValue(stKey, qf) Then
        '                                        qf = GetOrCreateQuarantineRdo(st)
        '                                        quarantineCache(stKey) = qf
        '                                    End If

        '                                    ' 搬移前先擷取資訊 (Move 後 m 會失效、EntryID 會變)
        '                                    Dim rcv As String = "" : Try : rcv = m.ReceivedTime.ToString("yyyy/MM/dd HH:mm") : Catch : End Try
        '                                    Dim subj As String = "" : Try : subj = m.Subject : Catch : End Try
        '                                    Dim sndr As String = "" : Try : sndr = m.SenderName : Catch : End Try
        '                                    Dim eidOld As String = "" : Try : eidOld = m.EntryID : Catch : End Try
        '                                    Dim stName As String = "" : Try : stName = st.Name : Catch : End Try

        '                                    m.Move(qf)   ' ← 搬到隔離夾
        '                                    movedCount += 1
        '                                    _dbg($"搬移 #{movedCount}", $"{rcv} | {sndr} | {subj}")
        '                                    moves.Add(String.Join(vbTab, {$"#{movedCount}", rcv, "寄件:" & sndr, "主旨:" & subj, "原夾:" & fName, "PST:" & stName, "舊EntryID:" & eidOld}))
        '                                    TryMarshalRelease(st)
        '                                Catch ex As System.Exception
        '                                    failCount += 1
        '                                    _dbg("RDO隔離 搬移失敗", ex.Message)
        '                                Finally
        '                                    TryMarshalRelease(m)
        '                                End Try
        '                            Next
        '                        Catch ex As System.Exception
        '                            _dbg("RDO隔離 略過夾", $"{fName} | {ex.Message}")
        '                        Finally
        '                            TryMarshalRelease(items)
        '                        End Try
        '                    Finally
        '                        Dim oOF As Object = rdoF : TryMarshalRelease(oOF)   ' 每夾 store-scoped 開出,逐一釋
        '                    End Try
        '                Next   ' For Each fe In scanEids
        '            Catch ex As System.Exception
        '                _dbg("RDO隔離 根節點失敗", $"{r.name} | {ex.Message}")
        '            End Try
        '        Next   ' For Each r In roots
        '    End Sub)

        'For Each kv In quarantineCache : TryMarshalRelease(kv.Value) : Next

        '' ── 3. 寫搬移紀錄檔 (與 OLAcache.db 同目錄) ──
        'Dim logPath As String = ""
        'Try
        '    Dim baseDir As String = If(String.IsNullOrEmpty(_dbCachePath), My.Application.Info.DirectoryPath, System.IO.Path.GetDirectoryName(_dbCachePath))
        '    logPath = System.IO.Path.Combine(baseDir, $"RpmsgMoved_{DateTime.Now:yyyyMMdd_HHmmss}.log")
        '    Dim header As New List(Of String) From {
        '        $"# RDO 保護信隔離搬移   {DateTime.Now:yyyy/MM/dd HH:mm:ss}",
        '        $"# 已掃 {scanned} 封, 搬移 {movedCount} 封, 失敗 {failCount} 封 → 各 PST 的 {QUARANTINE_NAME} 夾",
        '        ""}
        '    System.IO.File.WriteAllLines(logPath, header.Concat(moves), System.Text.Encoding.UTF8)
        'Catch ex As System.Exception
        '    _dbg("RDO隔離 寫檔失敗", ex.Message)
        'End Try

        '_dbg("RDO隔離 完成", $"掃 {scanned} | 搬 {movedCount} | 失敗 {failCount} | log: {logPath}")
    End Function
    Private Function GetOrCreateQuarantineRdo(st As Redemption.RDOStore) As Redemption.RDOFolder
        ' 2026/06/22 by Simon/Claude Opus 4.8: 取得(或建立)指定 store 頂層的 _IRM_Protected 隔離夾
        Dim root As Redemption.RDOFolder = st.IPMRootFolder   ' store 的可見頂層夾 (PST 適用)
        Try
            Dim subs = root.Folders
            For i As Integer = 1 To subs.Count
                Dim f As Redemption.RDOFolder = subs.Item(i)
                If String.Equals(f.Name, QUARANTINE_NAME, StringComparison.OrdinalIgnoreCase) Then Return f   ' 已存在
            Next
            Return subs.Add(QUARANTINE_NAME)   ' ★ 唯一沒在文件逐字確認的 API (鏡像 OOM Folders.Add); 不編譯就是這行
        Finally
            TryMarshalRelease(root)
        End Try
    End Function

    ' 2026/6/19~20 獨立 session 給的 EntryID，能否用 OOM _olNS.GetItemFromID 還原
    ' 2026/06/19 by Simon/Claude: 拋棄式 spike — 驗證 RDO 獨立 session 三件事
    '   (1) Outlook 已掛載 PST 時，獨立 RDOSession 能否 Logon (PST 共享鎖)
    '   (2) 該獨立 session 能否讀到 RdoTest 內信件的附件檔名
    '   (3) 獨立 session 給的 EntryID，能否用 OOM _olNS.GetItemFromID 還原
    ' 測完即可整段刪除。請暫時掛到一個測試按鈕呼叫。
    ' ============================================================
    Private Async Function SpikeParallelReadBenchmark() As Task
        ' 2026/06/22 by Simon/Claude Opus 4.8: 拋棄式 spike P3 — 量測「同 profile 多獨立 session、各讀不同 PST」
        '   的真實平行加速。回答整輪調查唯一未解的問題: K 條 session 跨 PST 並行讀取, wall-clock 是否
        '   勝過序列, 還是被 MSPST provider / 實體磁碟 I/O 序列化。
        '   ★ 兩種 workload 分別計時(附件檔名 vs 內文), 因 Tab3/Tab5 負載特性可能不同。
        '   ★ 公平性: 每個 (workload,K) 各讀「獨立的冷 block」(不同信), 避免暖快取讓後跑的 config 假性變快。
        '   ★ 用與 production 同一支 API: sess.GetMessageFromID(EntryID) + rdoMsg.Attachments/.Body
        Const N As Integer = 2000      ' 每 PST 每個 block 的冷讀信數(想要更穩可調 2000, 時間約翻倍)
        Const M As Integer = 4         ' 取幾個「夠大」的 PST 當標的(K=4 時每 worker 各 1 個)
        Const BLOCKS As Integer = 6    ' 2 workload × 3 K-config; 每 PST 需 >= BLOCKS*N 封冷信
        ' ── 1. 收集階段: 臨時一條 session 走訪, 挑 M 個有 >= BLOCKS*N 封的 PST, 各收 BLOCKS*N 個 EntryID ──
        '    (EntryID 是字串、跨 session 通用, 收一次給所有 worker 重用; RDOStore 物件不可跨 session 持有)
        ' ── 2. 對 2 種 workload × K=1/2/4 量測 ──
        Dim workloads = {"附件檔名", "內文"}
        Dim kConfigs = {1, 2, 4}
        Dim summary As New List(Of String)()
        _dbg("P3", "===== 量測結束, 摘要(看 K=2/4 吞吐相對 K=1 有沒有上去) =====")
        For Each s In summary : _dbg(" │摘要", s) : Next
        _dbg("P3", "===== 請把本段全部貼回 =====")
    End Function   ' 2026/06/22 P3量測「同 profile 多獨立 session、各讀不同 PST」的真實平行加速
    Private Async Function SpikeResolveFormCompare() As Task
        ' 2026/06/22 by Simon/Claude Opus 4.8: 拋棄式 spike B — 釘死「P3 附件 K=1 達 5589 封/s, 但 production_1/_2 只有 200 多」這 25 倍矛盾。空轉假設已被推翻(本批信 ~55% 有附件), 剩三混淆變數:
        '     (a)resolve 形式  (b)session 種類  (c)取樣信 vs sourceList 不同 ←本 spike 用「同批信讀三遍」消掉 c. 單執行緒(純比 per-call 成本, 不平行), 同一批信讀三種形式:
        '     (1)共用_rdo 單參數      = 現行 production
        '     (2)共用_rdo store-scoped → (1)vs(2)= resolve 形式效應(同一 session)
        '     (3)獨立session store-scoped = P3 → (2)vs(3)= session 種類效應
        '   依賴: FindStoreByPath(寫 P4 時放的 class-level 函數)。前提: Outlook 切 Work profile。測完即可整段刪除。
        Const N As Integer = 2000      ' 取樣信數(單執行緒, 同一批讀三遍; 夠大讓 封/s 穩定)
        If _rdo Is Nothing Then Await InitRdoSessionWithoutEULA()
        If _rdo Is Nothing Then _dbg("B", "Redemption 初始化失敗, 中止") : Return
        Dim profileName As String = ""
        Try : profileName = CStr(CallByName(_rdo, "ProfileName", CallType.Get)) : Catch : End Try
        _dbg("B", $"===== resolve 形式對照 (profile=[{profileName}], N={N}, 單執行緒) =====")
        _dbg("B", "===== 對照結束, 請貼回(三個附件數應一致才公平) =====")
    End Function      ' 2026/6/23, 修改P3, 開始比較獨立session 形式對效能的影響倍數, 與平行度效能吞吐量測試
    Private Async Function SpikeBodyResolveCompare() As Task
        ' 2026/06/22 by Simon/Claude Opus 4.8: 拋棄式 spike B-內文版 — 驗證「內文讀取換獨立 session 是否也有 ~10×」。
        '   注意: 內文 production 路徑(GetMailBodyOOM 第2190行)走 OOM, 不是 _rdo, 故基準與附件版不同, 測三條:
        '     (1) OOM _olNS.GetItemFromID + .Body  = 內文現行 production 基準(你說的 70~80 封/s 來源)
        '     (2) 共用 _rdo store-scoped + .Body    → (2)vs(3) 對照「共用 vs 獨立 session」這條槓桿在內文是否成立
        '     (3) 獨立 session store-scoped + .Body = 目標形式
        '   防 IRM: 取樣時用 RDO 預掃 MessageClass, 跳過 IPM.Note.* 受保護(rpmsg)信, 避免 OOM .Body 卡死授權 modal。
        '   ★全程 UI/STA 緒同步跑: OOM COM 不可進 Task.Run; N=1000 單執行緒, UI 短暫凍結可接受。
        '   依賴: FindStoreByPath(P4 放的)。前提: Outlook 切 Work profile。測完即整段刪除。
        Const N As Integer = 1000
        If _rdo Is Nothing Then Await InitRdoSessionWithoutEULA()
        If _rdo Is Nothing Then _dbg("B內文", "Redemption 初始化失敗, 中止") : Return
        Dim profileName As String = ""
        Try : profileName = CStr(CallByName(_rdo, "ProfileName", CallType.Get)) : Catch : End Try
        _dbg("B內文", $"===== 內文 resolve 形式對照 (profile=[{profileName}], N={N}, 單執行緒/UI緒) =====")
    End Function      ' 驗證「內文讀取換獨立 session 效能與平行度效能吞吐量測試」
    ' 2026/06/23 by Simon/Claude: 探針 — 驗證獨立 session _rdo2 的 resolve 形式
    '   目的: 用 OOM 取得的 (EntryID, OOM StoreID, FolderPath) 在 _rdo2 上分別試三種
    '         resolve, 決定 production 該走「雙參數」還是「store-scoped」。
    '   判讀: 看哪種形式 resolve 成功率高、且 Subject 對得上 (= 真解到, 非空 handle)。
    '   ※ 純診斷, 不動 production; 用完即可整段刪除。
    ' =================================================================

    ' 2026/06/24 by Simon/Claude Opus 4.8: 拋棄式探針 — 子樹階層走訪 OOM vs RDO批次 對拍
    '   本輪唯一目的: 先確認 API 讀法寫對 + 取得「暖快取」基準值(供 GetSubtreeRdo 完工後比對是否有額外開銷)。
    '   標的: SimTree3.SelectedNodes 當 root(可多選逐一各跑;Simon 自行換不同深淺節點重跑)。
    '   對手(全單執行緒,全產出「子孫 path 集合」對拍):
    '     A  OOM        : current.Folders 逐夾 BFS(= GetSubtreeOOM 去副作用版,基準)
    '     B  RDO-Enum   : rdoFolder.Folders For Each 逐夾(診斷: 隔離 RDO 層 vs OOM 層)
    '     C  RDO-Batch  : Folders.MAPITable.GetRows 整層批次,只對 PR_SUBFOLDERS=true 遞迴(候選)
    '     C+ RDO-Batch+CC: C 多撈 PR_CONTENT_COUNT(獨立計時,驗免費搭車且不污染 A/B/C)
    '   正確性對拍用 path 集合(最穩);EntryID 經 SpikeEidToHex 統一轉 hex 供遞迴。
    ' ============================================================================

    ' 2026/6/27 開始測試foldersize用的GetRows()和ExecSQL()
    ' 2026/06/27 by Simon/Claude Opus 4.8 (v2): size 讀取法對拍 — 修 v1 三問題
    '   (1)選 PST root 本層 0 封,假性「一致」無意義 → cnt=0 直接標記跳過。
    '   (2)B 讀 PR_MESSAGE_SIZE_EXTENDED(PT_I8)經 GetRows 每封都回相同垃圾常數 ≈ -2^31 → 改讀 PR_MESSAGE_SIZE(PT_LONG,0x0E080003)。
    '   (3)ExecSQL SUM 實測 AV(REDEMP~2.DLL)→ 確認不可用,移除。
    '   對手: A OOM-GetArray(PT_I8,基準) vs B RDO-GetRows(PT_LONG)。parity ✗ 時自動 dump 前3列原始型別/值。測完即刪。
    ' ============================================================================

    Private Class ProbeAttachScanResult
        Public Ok As Boolean = False
        Public ErrMsg As String = ""
        Public RowCount As Integer = 0                                  ' 掃過的全列數
        Public NHasAtt As Integer = 0                                   ' raw PR_HASATTACH=True 封數
        Public NSmartTrue As Integer = 0                                ' 其中 SmartNoAttach=True 封數 (被剔除者)
        Public CandEids As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)   ' 迴紋針候選 = HasAtt 且非 SmartTrue
        Public CandMails As New List(Of MailItemInfo)(256)              ' 候選的完整 MailItemInfo (對齊 production 產出,計時才誠實)
        Public AllMails As List(Of MailItemInfo) = Nothing              ' E 路線用: 全列 MailItemInfo (piggyback 同時餵 _cacheMailInfo 的形狀)
        Public FlagsByEid As New Dictionary(Of String, (subj As String, hasAtt As Boolean, smartRaw As String))(StringComparer.OrdinalIgnoreCase)
        Public SmartTypeTally As String = ""                            ' SmartNoAttach 欄位原始值型別分佈 (診斷 named prop 是否讀得到)
        Public ColsMode As String = ""                                  ' named prop 欄位用裸寫還是引號寫法成功
    End Class
    ' 2026/07/04 by Simon/Claude Fable 5 [PROBE_ATTACHBATCH]: 驗證「迴紋針語意能否純批次判定」。
    '   假說: OOM @SQL hasattachment(迴紋針) ≈ PR_HASATTACH=True 且 PidLidSmartNoAttach≠True。
    '     SmartNoAttach = PSETID_Common {00062008-0000-0000-C000-000000000046} / dispid 0x8514 / PT_BOOLEAN,
    '     MS 官方語意「訊息沒有使用者可見附件」— 正好是 memory_20260623_2210 更正記錄裡
    '     RDO 多算 4373 封(內嵌圖 Hidden / olOLE)缺的那塊拼圖。Redemption MAPITable.Columns 官方文件
    '     明載支援 id/{guid}/dispid 格式 named prop(內部走 GetIDsFromNames)。
    '   若 parity 過 → (1) GetAttMailList ③ 可改 RDO 批次(不再依賴 Outlook 查詢引擎的迴紋針加工);
    '                  (2) GetMailInfoRdo 掃描可 +2 欄「順手」填 Tab3 _cacheAttMailList(Simon 提的 piggyback);
    '                  (3) 可平行化(獨立 RDOSession)做全庫預熱。
    '   對照組(每夾四路):
    '     A = production GetAttMailListOOM            — 語意基準 + 時間基準
    '     B = RDO MAPITable 7欄批次(5欄+HASATTACH+SmartNoAttach) 讀全列,client 過濾 — 語意/速度受測者
    '     C = production GetMailInfoRdo(5欄全列)          — B−C = piggyback 邊際成本
    '     D = ExecSQL WHERE PR_HASATTACH<>0 預過濾小結果集 — Tab3 專用冷路徑候選
    '   觸發: 命令列 /autoprobeattachbatch 或 /autoprobeattachbatch:StoreName|FolderName(或 Store|*),
    '         搭配 /autoclose 測完自動走正常關閉流程。結果寫 %TEMP%\OutlookAssistant_ProbeResult_AttachBatch.txt
    Private Function ProbeAttachBatchRdoScan(fPath As String, eid As String, buildAllMails As Boolean) As ProbeAttachScanResult
        ' B 路線 (buildAllMails=False): 比照 production GetMailInfoRdo 的 MAPITable+Columns+GetRows(5000) 批次形狀,
        '   7 欄(5顯示欄+HASATTACH+SmartNoAttach),讀全列 client 過濾,對候選建完整 MailItemInfo — 即候選函式 GetAttMailListRdo 的原型。
        ' E 路線 (buildAllMails=True): 再多掛 GetMailInfoRdo 的 MSGID/SENDER_EMAIL 兩欄(共9欄),全列建 MailItemInfo(含 XxHash),
        '   同時產出候選清單 — 即「GetMailInfoRdo 掃描順手填 Tab3 快取」的 piggyback 原型,與 production C 相減得邊際成本。
    End Function
    Private Function ProbeAttachExecSqlScan(fPath As String, eid As String) As ProbeAttachScanResult
        ' D 路線: ExecSQL 伺服端預過濾 PR_HASATTACH<>0 → 只回附件候選小結果集,client 再剔 SmartNoAttach。
        '   語法鐵則(memory_20260623_2210 深夜補測): 不吃 = true 字面值要用 <> 0;欄位/條件 DASL 要帶雙引號。
        '   SELECT 帶齊 production 會需要的顯示欄位(Subject/ReceivedTime/SenderName/Size),而且逐列全讀,計時才誠實。
    End Function

    ' 2026/07/04 by Simon/Claude Fable 5 [PROBE_TAB1E2E]: 追查 Simon 回報「Tab1 第二次全選所有 store 仍要 0.55 秒」。
    '   懷疑點: 快取字典(mca/fca/mc)第二輪全命中,但「取用快取之前」有每輪重付的 COM 成本:
    '     (a) GetBfsResult 對 root+直屬子夾逐一 GetFolderById(COM 物化) + 命中夾 GetMailCount(folder版→SafeGetPath 讀 .FolderPath)
    '     (b) BuildLv1Item 的 IsMailFolder(folder)(無 fPath→SafeGetPath COM) + folder.Name(COM)
    '     (c) 骨架記憶體處理(FilterSubtreeByMode+建樹+每層排序) 與 RenderLv1
    '   Pass1=首次全選(冷) / Pass2=第二次全選(Simon 場景) / Pass3=暖態逐步拆帳(重演 GetBfsResult 內部,細分計時計次)。
    '   觸發: /autoprobetab1 (+ /autoclose)。結果寫 %TEMP%\OutlookAssistant_ProbeResult_Tab1E2E.txt

#End Region

End Class
