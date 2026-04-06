Imports System.Collections.Concurrent
Imports Microsoft.Office.Interop

Partial Class Form1

#Region "■ 01 全域宣告"
#Region "  ├ Win32 API 宣告"
    ' ── 函數宣告 ────────────────────────────────────────────────────────────────
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
    Private Shared Function ShowWindow(ByVal hWnd As IntPtr, ByVal nCmdShow As Integer) As Boolean

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
    Private Shared Function GetWindowLong(hWnd As IntPtr, nIndex As Integer) As Integer

    End Function
    <Runtime.InteropServices.DllImport("user32.dll")>
    Private Shared Function SetWindowLong(hWnd As IntPtr, nIndex As Integer, dwNewLong As Integer) As Integer

    End Function

    Private Const GWL_STYLE As Integer = -16
    Private Const WS_TABSTOP As Integer = &H10000
    ' ── 常數 ────────────────────────────────────────────────────────────────────
    Private Const SW_HIDE As Integer = 0
    Private Const WM_COMMAND As Integer = &H111
    Private Const WM_LBUTTONDOWN As Integer = &H201
    Private Const WM_LBUTTONUP As Integer = &H202
    Private Const BM_CLICK As Integer = &HF5
    ' TreeView 雙緩衝
    Private Const TV_FIRST As Integer = &H1100
    Private Const TVM_SETEXTENDEDSTYLE As Integer = TV_FIRST + 44
    Private Const TVS_EX_DOUBLEBUFFER As Integer = &H4
    ' ListView 雙緩衝
    Private Const LVM_SETEXTENDEDLISTVIEWSTYLE As Integer = &H1036
    Private Const LVS_EX_DOUBLEBUFFER As Integer = &H10000
    Private Const SWP_NOZORDER As Integer = &H4                     ' debugForm resize用
    Private Const SWP_NOACTIVATE As Integer = &H10                  ' debugForm resize用
    Private Const SWP_NOREDRAW As Integer = &H8                     ' debugForm resize 時閃爍, 改手動redraw
    Private Const RDW_ALLCHILDREN As Integer = &H80                 ' debugForm resize 時閃爍, 改手動redraw
    Private Const RDW_INVALIDATE As Integer = &H1                   ' debugForm resize 時閃爍, 改手動redraw
    Private Const RDW_UPDATENOW As Integer = &H100                  ' debugForm resize 時閃爍, 改手動redraw
    Private Const RDW_ERASE As Integer = &H4                        ' 2026/03/28 by AntiGravity: 補上缺失定義
    Private Const RDW_FRAME As Integer = &H400                      ' 2026/03/28 by AntiGravity: 補上缺失定義
    Private Const WM_SETREDRAW As Integer = &HB                     ' 2026/3/26 by AntiGravity
    ' ↓ 新增 (2026-03-20) ListView1 進入資料夾用
    Private Const TVM_SELECTITEM As Integer = &H110B                ' = &H1100 + 11
    Private Const TVGN_CARET As Integer = &H9                       ' SendMessage 選取 Treeview 游標節點
    Private Const LVM_SETITEMCOUNT As Integer = &H1000 + 47         ' = &H102F '
#End Region
#End Region

#Region "■ 99 舊版備用 (勿刪)"
    Private Function GetFolderByName(folderName As String) As Outlook.Folder
        Dbg("開始", folderName)
        Dim node As TreeNode = FindNodeByName(TreeView1.SelectedNode, folderName)
        Return If(node Is Nothing, Nothing, node.Tag)
    End Function
    Private Function FindNodeByName(selectedNode As TreeNode, ByVal findName As String) As TreeNode
        Dbg("開始", selectedNode.Name)
        selectedNode.Expand()
        ' 目前的實現是每次比較都把 " - " 去掉再比對，
        ' 建議改成在一開始就把 findName 處理成 cleanName，然後在整個遞迴過程中都使用 cleanName 來比對，
        ' 這樣就不需要每次都呼叫 Replace 了，效能應該會更好... 嗎??
        ' GetFolderByName()跟FindNodeByName()有更好的寫法嗎? 更快的或是更簡潔的?
        ' 2026/4/5 by simon: 目前這二個函數己經沒用到了 (FindNodeByName, FindNodeByName)
        Dim cleanName As String = findName.Replace(" - ", "")
        For Each node As TreeNode In selectedNode.Nodes
            If node.Text = cleanName Then Return node
            Dim foundNode As TreeNode = FindNodeByName(node, findName)  ' 遞迴往下搜尋直到符合才return，找到就不再繼續往下搜尋了
            If foundNode IsNot Nothing Then Return foundNode
        Next
        Return Nothing

    End Function
    Private Function FindNodeOrItemByName(ByVal nodesOrItems As IEnumerable, ByVal itemName As String) As Object
        Dbg("開始", itemName)
        For Each item As Object In nodesOrItems
            Dim text As String = If(TypeOf item Is TreeNode, DirectCast(item, TreeNode).Text, DirectCast(item, ListViewItem).Text)
            If text.Replace(" - ", "") = itemName.Replace(" - ", "") Then
                Dbg("結束", $"找到: {itemName}") ' by AntiGravity, 2026/04/04: 補上找到時的結束（Issue 2）
                Return item
            End If
        Next
        Dbg("結束", $"找不到: {itemName}") ' by AntiGravity, 2026/04/04: 補上找不到時的結束（Issue 2）
        Return Nothing

    End Function

    Private Function GetMailCountRecursiveLegacy(folder As Outlook.Folder) As Integer
        Dbg("開始", folder.Name)
        Dim value As Integer
        If _cacheMailCountAll.TryGetValue(folder, value) Then Return value ' 檢查快取中是否已存在值, 若有則直接返回
        ' 改成先用 Parallel.ForEach 遍歷子文件夾並且並行處理
        Dim totalMailCount As Integer = 0
        Dim countingBag As New ConcurrentBag(Of Integer)()
        Try
            ' 5/21記錄: 模仿GetFolderSizeLegacy那一句超快速的LINQ, 但測試結果沒有現在這個快, 所以決定保留這個
            ' 2026/3/20, 重寫了底層GetMailCountAll() 但是不知為何效能還是比不過現在下面這個遞迴版本?? (todo: 暫時先保留)
            ' 原因: 原版遞迴只走一遍 COM 資料夾樹，新版走了兩遍COM:
            ' 第一遍: GetSubFolderList()    → BFS 遍歷，存取每個 folder.Folders
            ' 第二遍: For Each allFolders   → GetMailCount() 再讀每個資料夾一次
            ' 2026/3/22, 導入Redemption, 應該可以刪掉這裡了? 還是讓Redemption 變成on-demand, 需要才啟動?
            'Parallel.ForEach(folder.Folders.Cast(Of Outlook.Folder),' 取得子資料夾的郵件數量並添加到 ConcurrentBag 中
            '                 Sub(subFolder As Outlook.Folder)
            '                     countingBag.Add(GetMailCountRecursive(subFolder))
            '                 End Sub)
            'totalMailCount = countingBag.Sum() ' 累加所有子資料夾的郵件數量
            ''' 最後再獲取選取文件夾自身的郵件數量 (改用MAPI table 的PR_CONTENT_COUNT屬性來getmailcount)
            ''Const PR_CONTENT_COUNT As String = "http://schemas.microsoft.com/mapi/proptag/0x36020003"
            ''totalMailCount += folder.PropertyAccessor.GetProperty(PR_CONTENT_COUNT)
            totalMailCount += GetMailCount(folder)  ' 單一目錄的mail count改成重寫的統一底層函數, 2026/3/20
            _cacheMailCountAll.TryAdd(folder, totalMailCount) ' 第一次計算後就存入快取
        Catch
        End Try
        Return totalMailCount

    End Function
    Private Async Function GetMailCountAll_1(rootFolder As Outlook.Folder, Optional onProgress As Action(Of Integer, Integer) = Nothing) As Task(Of Long)
        ' --------------------------------------------------------------
        ' GetMailCountAll_1: 讀取某資料夾及其整棵子樹的郵件總數
        ' 先RDO, 再BFS累加, 再遞迴
        '
        ' 設計說明: 不自己遞迴，改用 GetSubFolderList() BFS 展開完整資料夾清單後逐一加總
        '           ① 可以在任意點插入取消檢查 (_cancelRequested) ，遞迴做不到
        '           ② 知道總資料夾數，可以回報準確進度
        '           ③ 多個統計 (mail count + folder size) 可共用同一份清單，不用各跑一次 BFS ' todo: 有沒有地方可以用上這個好處??
        '           ④ 沒有 stack overflow 風險 (BFS 用 Queue，不用 call stack)
        '
        '           為何呼叫 GetMailCount() 而非直接用 GetTable():
        '             PR_CONTENT_COUNT 是 Folder 物件上的已儲存屬性，Outlook 自動維護，讀取等於讀一個整數，一次 COM call 結束。
        '             GetTable() 會把資料夾內所有郵件 row 逐一回傳，只為了計數代價太高。GetTable 適合讀郵件內容 (大小、日期) ，不適合純計數。
        '
        '           回傳型別 Long 而非 Integer:
        '             單一資料夾用 Integer 夠 (PR_CONTENT_COUNT 是 PT_LONG 32-bit) ，
        '             但整棵子樹加總若有多個大資料夾，理論上可能超過 Integer.MaxValue (2,147,483,647) ，用 Long 安全。
        '
        ' Fallback 鏈:
        '   ⓪ Redemption : rdoFolder.TotalItemCount
        '            直接回傳整棵子樹的郵件總數，MAPI 層面的快取彙總值，一次 COM call 結束，完全不需要 BFS 遍歷
        '            Redemption 可正確讀取 PST 上此屬性 (原生 OOM 無法取得)
        '            _rdoSession 未就緒時自動跳過此層
        '   ① GetSubFolderList + GetMailCount(L3) 逐一加總, BFS 展開後逐一呼叫，清單與計算邏輯分離，支援取消和進度回報
        '   ② 遞迴 fallback: GetSubFolderList 本身失敗時 (極少見) 的保險方案, 遞迴版本無法回報精確進度，但確保結果正確
        '   ③ 兩層都失敗就回傳 Return -1 並記錄 DebugForm，不讓單一資料夾的讀取失敗影響整體加總。
        '
        ' cancelRequested 參數: ' todo: 如何使用??
        '   傳入 _cancelRequested 旗標的 ByRef，讓呼叫端可以中途 ESC 取消
        '   取消時回傳 -1，由 L1 判斷是否需要清空 UI
        '
        ' onProgress 參數 (可選) : ' todo: 如何使用??
        '   傳入 Action(Of Integer, Integer) callback，
        '   L2 每處理一個資料夾回報 (已完成數, 總數)，讓 L1 更新狀態列
        '   不需要進度回報時傳 Nothing
        '   注意: ⓪ Redemption 路徑一次取得結果，不會觸發 onProgress callback (無中間進度可回報)
        '
        ' 取代: GetMailCountByMAPINew 的整棵子樹加總用途
        '       (GetMailCountByMAPINew 內的 Parallel.ForEach 遞迴整段, 效能超快, 但不是好的做法)
        '
        ' 2026-03-22 新增 ⓪ Redemption TotalItemCount，_rdoSession 就緒時完全跳過 BFS
        ' --------------------------------------------------------------
        Dbg("開始", rootFolder.Name)
        ' ⓪ Redemption: TotalItemCount 直接回傳整棵子樹郵件總數
        '   MAPI 快取的彙總屬性，一次 call 結束，不需要 BFS 遍歷，也不需要平行處理
        '   原生 OOM 的 PR_MESSAGE_SIZE_EXTENDED 在 PST 上找不到，Redemption 可正確讀取
        '   2026-03-22 新增
        If _rdo IsNot Nothing Then
            Dim rdoFolder As Redemption.RDOFolder = Nothing
            Try
                rdoFolder = _rdo.GetFolderFromID(rootFolder.EntryID, rootFolder.StoreID)
                Dim total As Long = rdoFolder.TotalItemCount
                Dbg("GetMailCountAll ⓪ RDO 成功", $"{rootFolder.Name} | TotalItemCount={total}")
                Return total
            Catch ex As System.Exception
                Dbg("GetMailCountAll ⓪ RDO 失敗，走BFS fallback", $"{rootFolder.Name} | {ex.Message}")
            Finally
                TryMarshalRelease(rdoFolder)
            End Try
        End If
        ' ① 標準路徑: GetSubFolderList BFS 展開 + GetMailCount(L3) 逐一加總
        Try
            ' BFS 展開整棵子樹的資料夾清單 (復用現有函數，不重寫)
            Dim targetFolderList As List(Of Outlook.Folder) = GetSubFolderList(rootFolder, includeSubF:=True)
            Dim grandTotal As Long = 0
            For i As Integer = 0 To targetFolderList.Count - 1
                If _cancelRequested Then    ' ✅ 取消檢查: 任意點都可以乾淨中止，不像遞迴版難以插入
                    Dbg("GetMailCountAll 被取消", $"已處理 {i}/{targetFolderList.Count}") : Return -1
                End If
                Dim f As Outlook.Folder = targetFolderList(i)
                Dim count As Integer = GetMailCount(f)
                ' GetMailCount 的所有 fallback 都失敗才會到這個else，記錄但不中止整體加總
                If count >= 0 Then grandTotal += CLng(count) Else Dbg("GetMailCountAll 略過失敗資料夾", f.Name)
                onProgress?.Invoke(i + 1, targetFolderList.Count) ' 進度回報 (optional callback，呼叫端不需要時傳 Nothing 即可) 因為知道 total，進度條可以準確顯示百分比 'todo: 這個進度回報如何使用?
                If i Mod 10 = 0 Then Await Task.Yield()     ' 每掃瞄10個資料夾處理完就讓出一次，保持 UI 回應 (GetMailCount 本身是同步的，所以這裡的 Yield 是唯一的讓出點)
            Next
            Return grandTotal
        Catch ex As System.Exception
            Dbg("GetMailCountAll ① BFS路徑失敗，走遞迴fallback", $"{rootFolder.Name} | {ex.Message}")
        End Try
        ' ② 遞迴 fallback: GetSubFolderList 本身失敗時使用 (無法精確回報進度，但至少確保加總結果正確)
        '   注意: 遞迴層數受 PST 資料夾巢狀深度限制，實務上 PST 不會太深
        Try
            Dim totalCount As Long = 0
            Dim count As Integer = GetMailCount(rootFolder) ' 本層mailcount
            If count >= 0 Then totalCount += count : Await Task.Yield()
            For Each f As Outlook.Folder In rootFolder.Folders
                Dim subCount As Long = Await GetMailCountAll_1(f) ' 這個有問題!! 這裡遞迴的話, 會一直重複呼叫上面的GetSubFolderList(), 會跑到死....
                If subCount >= 0 Then totalCount += subCount
            Next
            Return totalCount
        Catch ex As System.Exception ' ③ 全部失敗就傳回 -1 讓上層流程去處理
            Dbg("GetMailCountAll ② 遞迴fallback也失敗", $"{rootFolder.Name} | {ex.Message}")
            Return -1   ' ③ 若前兩層都失敗，回傳 -1 讓 L2 知道這是「讀取失敗」而非「真的是 0 封」
        End Try

    End Function
    Private Async Function GetMailCountAll_2(rootFolder As Outlook.Folder, Optional onProgress As Action(Of Integer, Integer) = Nothing) As Task(Of Long)
        ' --------------------------------------------------------------
        ' 就是 GetMailCountAllParallel  v2.0: 讀取某資料夾及其整棵子樹的郵件總數
        ' 先RDO, 再平行, 再BFS累加
        '
        ' 平行策略:
        '   BFS 展開後，對每個資料夾各建一個 Task.Run，全部 Task.WhenAll 等待。不需再依PST StoreID 分組，結構最簡潔。
        '   PR_CONTENT_COUNT 是 Folder 上的已快取屬性，bottleneck 是 cross-process COM overhead，Outlook.exe 端能否真正並發處理需實測確認。
        '
        ' [2026-03-22 重要說明] Redemption 就緒後此函數實質上已被 GetMailCountAll ⓪ 取代
        '   原本設計平行處理是為了加速 BFS 逐一累加的瓶頸，
        '   但 Redemption 的 TotalItemCount 一次 call 就取得整棵子樹總數，
        '   平行處理的必要性消失。此函數保留作為:
        '   (a) _rdoSession 未就緒時的備用高速路徑 (走 Task.WhenAll 平行版)
        '   (b) 將來跨 PST 加總時的協調層 (多個 PST 的 GetMailCountAll 可以 Task.WhenAll)
        '   若確認 Redemption 穩定，日後可考慮廢棄此函數，呼叫端直接改用 GetMailCountAll。
        '
        ' [Redemption說明] 2026-03-22
        '   ⓪ Redemption TotalItemCount 一次取得，走此路徑時整個平行展開邏輯完全跳過
        '   ① Task.WhenAll 平行路徑: _rdoSession 未就緒時的 fallback
        '      Task.Run 內的 GetMailCount(f) 若走 Redemption ⓪，是 free-threaded 安全的
        '      若 fallback 到 MAPI PropertyAccessor，仍有 STA 違規風險，需留意
        ' --------------------------------------------------------------
        Dbg("開始", rootFolder.Name)
        ' ⓪ Redemption: TotalItemCount 直接回傳整棵子樹郵件總數
        '   就緒時完全跳過下方所有平行 BFS 邏輯，等同於 GetMailCountAll ⓪ 的行為
        '   2026-03-22 新增
        If _rdo IsNot Nothing Then
            Dim rdoFolder As Redemption.RDOFolder = Nothing
            Try
                rdoFolder = _rdo.GetFolderFromID(rootFolder.EntryID, rootFolder.StoreID)
                Dim total As Long = rdoFolder.TotalItemCount
                Dbg("GetMailCountAllParallel ⓪ RDO 成功", $"{rootFolder.Name} | TotalItemCount={total}")
                Return total
            Catch ex As System.Exception
                Dbg("GetMailCountAllParallel ⓪ RDO 失敗，走平行BFS fallback", $"{rootFolder.Name} | {ex.Message}")
            Finally
                TryMarshalRelease(rdoFolder)
            End Try
        End If
        ' ① 標準路徑: BFS 展開 → 每個資料夾一個 Task → Task.WhenAll
        Try
            Dim targetFolderList As List(Of Outlook.Folder) = GetSubFolderList(rootFolder, includeSubF:=True)
            Dim targetFolderCount As Integer = targetFolderList.Count
            Dim processedCount As Integer = 0   ' Interlocked 保證多 Task 同時更新的執行緒安全
            Dim folderTasks = targetFolderList.Select(
                Function(f) Task.Run(Function() As Integer
                                         If _cancelRequested Then Return 0
                                         Dim count As Integer = GetMailCount(f)
                                         If count < 0 Then
                                             Dbg("GetMailCountAllParallel 略過失敗資料夾", f.Name)
                                             count = 0
                                         End If
                                         ''Dim done As Integer = Interlocked.Increment(processedCount)
                                         'onProgress?.Invoke(done, targetFolderCount) : Return count
                                     End Function)).ToList()
            Dim results As Integer() = Await Task.WhenAll(folderTasks)
            If _cancelRequested Then
                Dbg("GetMailCountAllParallel 已取消", $"總資料夾數: {targetFolderCount}") : Return -1
            End If
            Return results.Sum(Function(c) CLng(c))
        Catch ex As System.Exception
            Dbg("GetMailCountAllParallel ① 平行路徑失敗，走循序fallback", $"{rootFolder.Name} | {ex.Message}")
        End Try
        ' ② 循序 fallback: 平行路徑失敗時使用，退回單純的逐一加總
        '   不用遞迴 (避免重複呼叫 GetSubFolderList) ，直接重跑 BFS 循序版
        Try
            Dim targetFolderList As List(Of Outlook.Folder) = GetSubFolderList(rootFolder, includeSubF:=True)
            Dim grandTotal As Long = 0
            For i As Integer = 0 To targetFolderList.Count - 1
                If _cancelRequested Then Return -1
                Dim count As Integer = GetMailCount(targetFolderList(i))
                If count >= 0 Then grandTotal += CLng(count)
                If i Mod 10 = 0 Then Await Task.Yield()     ' 每掃瞄10個資料夾處理完就讓出一次，保持 UI 回應
            Next
            Return grandTotal
        Catch ex As System.Exception
            Dbg("GetMailCountAllParallel ② 循序fallback也失敗", $"{rootFolder.Name} | {ex.Message}")
            Return -1       ' ③ 若前兩層都失敗，回傳 -1 讓 L2 知道這是「讀取失敗」而非「真的是 0 封」
        End Try

    End Function

    Private Async Function GetTotalFolderCountAsync(folder As Outlook.Folder) As Task(Of Integer)
        Dbg("開始", folder.Name)
        Dim value As Integer
        Dim fPath As String = folder.FolderPath
        If _cacheFolderCountAll.TryGetValue(fPath, value) Then Return value ' 檢查快取中是否已存在值, 若有則直接返回
        Dim totalSubCount As Integer = GetFolderCount(folder)           ' 初始值為點選資料夾的子資料夾數量
        ' 5/21測試記錄: 這裡使用ConcurrentBag跟使用results.sum應該要比較快, 但不知為何實測結果都比GetTotalFolderCount_Old()還慢了5%, 這個函數先保留不清除
        ' 5/21最後決定: 二個函數快慢互有變化, 但GetTotalFolderCountAsync()的穩定性較好, 比New()的標準差來得小, 所以決定使用這個
        ' 使用 Parallel.ForEach 進行多線程遞迴計算subfolder數量
        Dim countingBag As New ConcurrentBag(Of Task(Of Integer))()     ' 使用 ConcurrentBag 來安全地收集每個子資料夾的數量
        Parallel.ForEach(folder.Folders.Cast(Of Outlook.Folder)(),
                         Sub(subFolder As Outlook.Folder)
                             'countingBag.Add(GetTotalFolderCountAsync(subFolder))
                             countingBag.Add(GetFolderCountAll(subFolder))
                         End Sub)
        Dim results = Await Task.WhenAll(countingBag)   ' 等待所有平行出去收集的數量都確定回來了
        totalSubCount += results.Sum()                  ' 再將回傳的各個子資料夾的數量加總
        _cacheFolderCountAll.TryAdd(folder.FolderPath, totalSubCount)
        ' ✅ 2026-03-16 移除多餘的 Try/Catch: ConcurrentDictionary.TryAdd 本身不拋例外 (只回傳 True/False)
        ' 原本是從 .Add() 時代留下的防護，改 TryAdd 後應一併移除
        Return totalSubCount

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
        '   結果存入 folderSizeCache，BuildListViewItem_Tab1 下次組裝時自動顯示
        ' ==============================================================
        Dbg("開始", folder.Name)
        Dim value As Long   ' 快取命中直接回傳
        If _cacheFolderSize.TryGetValue(folder, value) Then Return value
        '' 已知有問題的資料夾走舊路徑 (不明 COM 例外物件，GetTable 也可能出問題)
        'Dim exceptList As String() = {"Inbox_2000~2018", "Facebook"}
        'If exceptList.Contains(folder.Name) Then Return GetFolderSizeOld(folder)
        Dim table As Outlook.Table = Nothing
        Try
            ' GetTable + PR_MESSAGE_SIZE (0x0E080003) :
            ' PR_MESSAGE_SIZE_EXTENDED (0x0E080014, PT_I8) — PST 本地端的內建彙總屬性
            Const PR_MESSAGE_SIZE As String = "http://schemas.microsoft.com/mapi/proptag/0x0E080003"
            Const PR_SIZE_EXTENDED As String = "http://schemas.microsoft.com/mapi/proptag/0x0E080014"
            ' 只讀 Size 欄，不載入其他 MAPI 屬性，減少記憶體與 COM 開銷
            table = folder.GetTable()
            table.Columns.RemoveAll()
            table.Columns.Add(PR_SIZE_EXTENDED)
            Dim totalSize As Long = 0
            Dim rowCount As Integer = 0
            Do While Not table.EndOfTable
                Dim row As Outlook.Row = table.GetNextRow()
                totalSize += SafeGet(Of Long)(row, PR_SIZE_EXTENDED, 0L)
                TryMarshalRelease(row)
                rowCount += 1
                If rowCount Mod 100 = 0 Then Await Task.Yield()  ' 每 100 筆統計就讓 UI 回應一次
            Loop
            _cacheFolderSize.TryAdd(folder, totalSize)
            Return totalSize
        Catch ex As OverflowException
            Dbg("Error: GetFolderSizeLegacy overflow", folder.Name)
            Return -1
        Catch ex As System.Exception
            Dbg("Error: GetFolderSizeLegacy", folder.Name & " - " & ex.Message)
            Return -1
        Finally
            TryMarshalRelease(table)
        End Try

    End Function
    Private Function GetFolderSizeOld(folder As Outlook.Folder) As Long
        Dbg("開始", folder.Name)
        Dim totalSize As Long = 0
        Dim folderItems As Outlook.Items = Nothing
        Try
            folderItems = folder.Items          ' ✅ 先取出 Items 物件，才能在 Finally 釋放
            For Each item As Object In folderItems
                Try
                    Dim mailItem As Outlook.MailItem = DirectCast(item, Outlook.MailItem)
                    If mailItem IsNot Nothing Then
                        totalSize += mailItem.Size
                        'tasks.Add(Task.Run(Async Function() ' 使用非同步 IO 操作來取得郵件大小
                        '                       'Await mailItem.PropertyAccessor.GetPropertyAsync("http://schemas.microsoft.com/mapi/proptag/0x0E080014")
                        '                       Interlocked.Add(sizeAdder, mailItem.Size)
                        '                   End Function))
                    End If
                Catch
                End Try
            Next
            'Await Task.WhenAll(tasks) ' 等待所有非同步操作完成
        Finally
            TryMarshalRelease(folderItems)  ' ✅ Items 集合釋放
        End Try
        Return totalSize

    End Function
#End Region


End Class
