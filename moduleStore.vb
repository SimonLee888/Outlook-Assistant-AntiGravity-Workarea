Imports System.Collections.Concurrent
Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Outlook
Module moduleStore
    Private folderSizeCache As New Dictionary(Of Outlook.Folder, Long)
    Private ProgressBar1 As Object
    Private ProgressBar2 As Object
    Private sw0 As Object
    Private Async Sub tmrPreCache_Tick(sender As Object, e As EventArgs)
        ' DebugForm.AddMessage("Begin:")
        'tmrPreCache.Enabled = False    ' 避免重覆啟動
        'Dim swa, swb As New Stopwatch
        'swa.Start() : swb.Start()
        'For Each store In PstStoreList
        '    Dim root As Outlook.Folder = store.GetRootFolder
        '    Dim folderQueue As List(Of Outlook.Folder) = Await Task.Run(Function() GetFolderListByTierAsync(root, preCacheTierIndex))
        '    For Each f As Outlook.Folder In folderQueue
        '        Debug.Items.Add("Processing: " & f.FolderPath)
        '        Dim cache1 = Task.Run(Function() GetTotalFolderCountAsync(f))
        '        Dim cache2 = Task.Run(Function() GetMailCountByMAPINew(f))
        '    Next
        '    swa.Stop()
        '    Debug.Items.Add(folderQueue.Item(0).FolderPath & " : " & folderQueue.Count & " (" & swa.Elapsed.TotalSeconds.ToString("##,#0.00)"))
        '    swa.Restart()
        'Next
        'If preCacheTierIndex < preCacheMaxTier Then
        '    preCacheTierIndex += 1 : tmrPreCache.Enabled = True     ' 開始下一輪的預讀
        '    'Else
        '    '    tmrPreCache.Enabled = False
        'End If
        'Debug.Items.Add("total time: " & swb.Elapsed.TotalSeconds.ToString("##,#0.00)"))
        'Debug.Items.Add("")

    End Sub
    Private Sub CacheTotalFolderSize(folder As Outlook.Folder)
        System.Console.WriteLine($"開始cache資料夾: {folder.FolderPath}")
        ' 檢查快取中是否已經存在該資料夾的大小值
        If folderSizeCache.ContainsKey(folder) Then Return
        Dim folderSize As Long = 0
        ' 遞迴計算所有子資料夾的大小
        For Each subFolder As Outlook.Folder In folder.Folders
            CacheTotalFolderSize(subFolder)
            If folderSizeCache.ContainsKey(subFolder) Then
                folderSize += folderSizeCache(subFolder)
            End If
        Next
        ' 計算當前資料夾中所有項目的大小
        For Each item As Object In folder.Items
            If TypeOf item Is Outlook.MailItem Then
                folderSize += DirectCast(item, Outlook.MailItem).Size
            End If
        Next
        ' 將結果存入快取
        folderSizeCache(folder) = folderSize
        System.Console.WriteLine($"完成cache資料夾: {folder.FolderPath }")

    End Sub
    Private Sub LoadSubFolders(folder As Outlook.Folder, nodes As TreeNodeCollection, tvwTarget As TreeView)
        'Dim sortedSubFolders As List(Of Outlook.Folder) = GetSortedSubFoldersAsync(folder).Result
        'For Each subFolder As Outlook.Folder In sortedSubFolders
        '    Dim subNode As TreeNode = Nothing
        '    tvwTarget.Invoke(Sub() subNode = nodes.Add(subFolder.Name))
        '    subNode.Tag = subFolder : TotalFolderCount += 1
        '    LoadSubFolders(subFolder, subNode.Nodes, tvwTarget)
        'Next
        'tvwTarget.BeginInvoke(Sub()
        '                          tvwTarget.EndUpdate()
        'ProgressBar2.Text = "載入所有資料夾花費了 " & sw0.Elapsed.TotalSeconds.ToString("0.00") & " 秒。"
        '                      End Sub)

    End Sub
    Private Function GetSortedSubFoldersAsync(folder As Outlook.Folder) As Object
        Throw New NotImplementedException()
    End Function
    Private Async Function LoadSubFoldersAsync(folder As Outlook.Folder, nodes As TreeNodeCollection, tvwTarget As TreeView) As Task
        Dim sortedSubFolders As List(Of Outlook.Folder) = GetSortedSubFoldersAsync(folder)
        'Dim tasks = sortedSubFolders.Select(Async Function(subFolder)
        '                                        Dim subNode As TreeNode = Nothing
        '                                        tvwTarget.Invoke(Sub() '如何在加入之前先刪除第一個節點就好?
        '                                                             'nodes(0).Remove()
        '                                                             subNode = nodes.Add(subFolder.Name) 'fixme: 是否應該在這裡檢查重覆節點? (ps. 好像沒用)
        '                                                         End Sub)
        '                                        Await Task.Run(Sub()
        '                                                           subNode.Tag = subFolder : TotalFolderCount += 1
        '                                                       End Sub)
        '                                        Await LoadSubFoldersAsync(subFolder, subNode.Nodes, tvwTarget)
        '                                    End Function)
        'Await Task.WhenAll(tasks)

    End Function
    Private Async Sub AddSubFoldersToTreeViewByBgWorker(folder As Outlook.Folder, nodes As TreeNodeCollection, tvwTarget As TreeView)
        '這個版本A還不錯, 可以正確執行, 資料夾正確, UI 沒有閃爍, 在背景讀取很好, 執行緒也未發生衝突, 是好的架構, 但是初次顥示有點慢, 按下後大約一秒才出現form.show,
        '雖然沒有占用UI執行緒, 但是全部的讀取時間超過二秒, 比我原來的慢很多, 這樣是正常的嗎? 是因為backgroundworker的執行緒優先權較低嗎?
        'Dim worker As New BackgroundWorker()
        'AddHandler worker.DoWork, Sub(sender, args) LoadSubFolders(folder, nodes, tvwTarget)
        'worker.RunWorkerAsync()
        'tvwTarget.BeginInvoke(Sub() tvwTarget.BeginUpdate())
        '===============================================================================================================================
        ''這個版本B效能比上面的bg worker更好, 執行緒更流暢, UI更穩定, 但formload延遲顯示沒有解決, 讀取進度不作用了, 所以也無法量測所花費的時間和速度?
        'tvwTarget.Invoke(Sub() tvwTarget.BeginUpdate())
        'LoadSubFoldersAsync(folder, nodes, tvwTarget).Wait()
        'tvwTarget.Invoke(Sub()
        '                     tvwTarget.EndUpdate()
        '                     ProgressBar2.Text = "載入所有資料夾花費了 " & sw0.Elapsed.TotalSeconds.ToString("0.00") & " 秒。"
        '                 End Sub)
        '===============================================================================================================================
        ''''可以正確的版本C
        '''Me.Invoke(Sub() Me.Show()) ' 立即顯示窗體
        '''sw0.Restart()
        '''tvwTarget.Invoke(Sub() tvwTarget.BeginUpdate())
        '''Dim loadTask = LoadSubFoldersAsync(folder, nodes, tvwTarget)
        ''''Await Task.Run(Async Sub() '顯示讀取進度
        ''''                   While Not loadTask.IsCompleted
        ''''                       Await Task.Delay(250) ' 延遲250毫秒
        ''''                       tvwTarget.Invoke(Sub()
        ''''                                            ProgressBar1.Text = $"正在載入: {TotalFolderCount} ({TotalFolderCount / sw0.Elapsed.TotalSeconds:###,##0}/sec) {folder.Name}資料夾"
        ''''                                        End Sub)
        ''''                   End While
        ''''               End Sub)
        '''Await loadTask
        '''tvwTarget.Invoke(Sub()
        '''                     tvwTarget.EndUpdate()
        '''                     ProgressBar2.Text = "載入所有資料夾花費了 " & sw0.Elapsed.TotalSeconds.ToString("0.00") & " 秒。"
        '''                 End Sub)
        '===============================================================================================================================
        ''可以正確的版本D
        'Me.Invoke(Sub() Me.Show()) ' 立即顯示窗體
        'sw0.Restart()
        'tvwTarget.Invoke(Sub() tvwTarget.BeginUpdate())
        'Dim loadTask = LoadSubFoldersAsync(folder, nodes, tvwTarget)
        'Dim progressTask = Task.Run(Async Sub() '顯示讀取進度
        '                                While Not loadTask.IsCompleted
        '                                    Await Task.Delay(250) ' 延遲250毫秒
        '                                    tvwTarget.Invoke(Sub() ProgressBar1.Text = $"正在載入: {TotalFolderCount} ({Math.Round(TotalFolderCount / sw0.Elapsed.TotalSeconds, 2):###,##0}/sec) {folder.Name}資料夾")
        '                                End While
        '                            End Sub)
        'Await loadTask
        'Await progressTask  ' 等待進度任務完成
        'tvwTarget.Invoke(Sub()
        '                     tvwTarget.EndUpdate()
        '                     ProgressBar2.Text = "載入所有資料夾花費了 " & sw0.Elapsed.TotalSeconds.ToString("0.00") & " 秒。"
        '                 End Sub)
        '===============================================================================================================================
        ''這個版本E在一開始就先即時顯示所有根節點, 背景再慢慢讀入其他部份,
        ''但是在背景讀取的時候把整個UI線程都塞住了, 而且根節點重覆加入
        ''不過這版是最快最穩定的, 只要解決重覆加入的問題就好了 (是否可以用hashset來檢查?)
        'Me.Invoke(Sub() Me.Show()) ' 立即顯示窗體
        'sw0.Restart()
        '' 先載入根資料夾
        'Dim rootNode As TreeNode = Nothing
        'tvwTarget.Invoke(Sub()
        '                     rootNode = nodes.Add(folder.Name)
        '                     rootNode.Tag = folder : TotalFolderCount += 1
        '                 End Sub)
        'tvwTarget.Invoke(Sub() tvwTarget.BeginUpdate()) '為什麼這行會執行那麼多次??
        ''fixme: 這裡如果傳入rootNode, 重覆的根節點就變成下一層的根節點, 如果傳入node, 重覆的根節點就變成下一層的第一個子節點,
        'Dim loadTask = LoadSubFoldersAsync(folder, nodes, tvwTarget)
        'Dim progressTask = Task.Run(Async Sub() '顯示讀取進度
        '                                While Not loadTask.IsCompleted
        '                                    Await Task.Delay(250) ' 延遲100毫秒
        '                                    tvwTarget.Invoke(Sub() ProgressBar1.Text = $"正在載入: {TotalFolderCount} ({Math.Round(TotalFolderCount / sw0.Elapsed.TotalSeconds, 2):###,##0}/sec) {folder.Name}資料夾")
        '                                End While
        '                            End Sub)
        'Await loadTask
        'Await progressTask  ' 等待進度任務完成
        'tvwTarget.Invoke(Sub()
        '                     tvwTarget.EndUpdate()
        '                     ProgressBar2.Text = "載入所有資料夾花費了 " & sw0.Elapsed.TotalSeconds.ToString("0.00") & " 秒。"
        '                 End Sub)

    End Sub
    Private Async Function GetTotalFolderCountAsyncDivideAndConquer(folder As Outlook.Folder) As Task(Of Integer)
        Dim totalCount As Integer = folder.Folders.Count
        Dim tasks As New List(Of Task(Of Integer))
        'sw1.Start()
        ' 將 Folders 集合包裝為 ParallelQuery
        Dim folderQuery As ParallelQuery(Of Outlook.Folder) = folder.Folders.AsParallel()
        ' 平行計算每個子資料夾及其子樹
        folderQuery.ForAll(Sub(subFolder) tasks.Add(GetTotalFolderCountAsyncDivideAndConquer(subFolder)))
        For Each task As Task(Of Integer) In tasks
            totalCount += Await task
        Next ' 等待所有子任務完成並累加結果
        'sw1.Stop()
        Return totalCount

    End Function
    Private Function GetTotalFolderCount_Old(folder As Outlook.Folder) As Task(Of Integer)
        '    ''2026/3/5重新檢查, 把GetTotalFolderCount_Old跟GetMailCountByMAPI_Old都註解掉停止使用
        '    ''只使用GetTotalFolderCountAsync跟GetMailCountByMAPINew這二個函數, 穩定性提升比較不會有偶爾會出現的快取例外問題, 但不知為何有時會慢一點點, todo: 有空再來測試看看是為什麼了, 先這樣用著了
        DebugForm.AddMessage3("Begin: ", folder.Name)
        '    '' 5/2記錄: 花了一天把GetTotalFolderCount跟GetTotalMailCount全部計時測試優化完成
        '    'Dim value As Integer
        '    'If folderCountCache.TryGetValue(folder, value) Then Return value ' 檢查快取中是否已存在值, 若有則直接返回
        '    'Dim totalSubfolderCount As Integer = folder.Folders.Count
        '    'Dim tasks As New List(Of Task(Of Integer)) ' 定義用於儲存每個子資料夾遞迴結果的集合
        '    '' 改用多線程平行處理每個子資料夾的遞迴, 遍歷每個子資料夾
        '    ''For Each subFolder As Outlook.Folder In folder.Folders
        '    ''    tasks.Add(Task.Run(Function() GetTotalFolderCount_Old(subFolder)))
        '    ''Next
        '    ''sw2.Start()
        '    'Parallel.ForEach(folder.Folders.Cast(Of Outlook.Folder),
        '    '                 Sub(subFolder)
        '    '                     tasks.Add(GetTotalFolderCount_Old(subFolder))
        '    '                 End Sub)
        '    'For Each subTask As Task(Of Integer) In tasks
        '    '    totalSubfolderCount += Await subTask
        '    'Next ' 等待所有子資料夾的遞迴結果完成並累加總數
        '    'sw2.Stop()
        '    'folderCountCache.Add(folder, totalSubfolderCount) ' 第一次計算後就存入快取
        '    'Return totalSubfolderCount

    End Function
    Private Function GetMailCountLINQ(folder As Outlook.Folder) As Integer
        ' 5/21記錄: 程式很短很簡潔, 速度雖快但很不穩定, 很容易造成執行緒打結慢下來, 先擺著不用
        Dim totalCount As Integer
        Task.Run(Sub()
                     ' 計算本資料夾中的郵件數量
                     Dim folderMailCount = folder.Items.OfType(Of Outlook.MailItem)().Count()
                     ' 遞迴計算子資料夾中的郵件數量並累加
                     Dim subFolderCounts = folder.Folders.Cast(Of Outlook.Folder)().Sum(Function(subFolder) GetMailCountLINQ(subFolder))
                     ' 計算總郵件數量
                     totalCount = folderMailCount + subFolderCounts
                 End Sub)
        Return totalCount

    End Function
    Private Async Function GetMailCountByLINQNew(folder As Outlook.Folder) As Task(Of Integer)
        ' 使用 LINQ 來計算子資料夾的郵件數量
        Dim subFolderTasks = folder.Folders.Cast(Of Outlook.Folder)().
        Select(Function(subFolder) Task.Run(Function() GetMailCountByLINQNew(subFolder)))
        ' 等待所有子資料夾的郵件數量計算完成
        Dim subFolderMailCounts = Await Task.WhenAll(subFolderTasks)
        Dim totalMailCount As Integer = subFolderMailCounts.Sum()
        ' 最後再獲取選取文件夾自身的郵件數量
        Try ' 改用MAPI table 的PR_CONTENT_COUNT屬性來getmailcount
            Const PR_CONTENT_COUNT As String = "http://schemas.microsoft.com/mapi/proptag/0x36020003"
            totalMailCount += folder.PropertyAccessor.GetProperty(PR_CONTENT_COUNT)
        Catch
        End Try
        Return totalMailCount

    End Function
    Private Function GetMailCountByMAPI_Old(targetFolder As Outlook.Folder) As Task(Of Integer)
        '    ''2026/3/5重新檢查, 把GetTotalFolderCount_Old跟GetMailCountByMAPI_Old都註解掉停止使用
        '    ''只使用GetTotalFolderCountAsync跟GetMailCountByMAPINew這二個函數, 穩定性提升比較不會有偶爾會出現的快取例外問題, 但不知為何有時會慢一點點, todo: 有空再來測試看看是為什麼了, 先這樣用著了
        DebugForm.AddMessage3("Begin: ", targetFolder.Name)
        '    'Dim totalMailCount As Integer = 0
        '    'sw3.Start()
        '    'Try ' 使用MAPI table 的PR_CONTENT_COUNT屬性來取得郵件數量
        '    '    Const PR_CONTENT_COUNT As String = "http://schemas.microsoft.com/mapi/proptag/0x36020003"
        '    '    totalMailCount = targetFolder.PropertyAccessor.GetProperty(PR_CONTENT_COUNT)
        '    'Catch
        '    'End Try
        '    '' 使用 Parallel.ForEach 遍歷子文件夾並且並行處理
        '    'Dim countingBag As New ConcurrentBag(Of Task(Of Integer))()
        '    'Parallel.ForEach(targetFolder.Folders.Cast(Of Outlook.Folder)(), Sub(subFolder As Outlook.Folder)
        '    '                                                                     countingBag.Add(GetMailCountByMAPI_Old(subFolder))
        '    '                                                                 End Sub)
        '    'Dim results = Await Task.WhenAll(countingBag)
        '    'totalMailCount += results.Sum()
        '    'sw3.Stop()
        '    'Return totalMailCount

    End Function
    Private Async Function GetFolderSizeLINQ(folder As Outlook.Folder) As Task(Of Long)
        DebugForm.AddMessage3("Begin: ", folder.Name)
        '' 1. 檢查 Cache 中是否已經存在該資料夾的大小
        'Dim value As Long
        'If folderSizeCache.TryGetValue(folder, value) Then Return value
        ''' 2. 若不在cache就嘗試使用 PR_FOLDER_SIZE_B 屬性獲取資料夾大小 (5/18 note: 這個屬性好像只能用在讀取遠端線上??)
        ''Try
        ''    Dim PR_FOLDER_SIZE As String = "&H0E07003A"     ' Extended MAPI: PR_FOLDER_SIZE
        ''    Dim PR_FOLDER_SIZE_B As String = "&H0E08001E"   ' Extended MAPI: PR_FOLDER_SIZE_B
        ''    Dim folderSizeFromProperty As Object = Nothing
        ''    folderSizeFromProperty = folder.PropertyAccessor.GetProperty(PR_FOLDER_SIZE_B)
        ''    If TypeOf folderSizeFromProperty Is Long Then
        ''        folderSize = DirectCast(folderSizeFromProperty, Long)
        ''        folderSizeCache.Add(folder, folderSize)
        ''        Return folderSize
        ''    End If
        ''Catch
        ''End Try
        '' 3. 若無法從 MAPI屬性直接獲取資料夾大小, 則回到使用 PR_MESSAGE_SIZE 或 MailItem.Size 遍歷所有郵件並加總
        ''Dim folderSize As Long = 0
        ''For Each item As Object In folder.Items
        ''    If TypeOf item Is MailItem Then
        ''        Dim intsize As Integer = Await GetMailItemSizeAsync(DirectCast(item, MailItem))
        ''        folderSize += intsize
        ''    End If
        ''Next
        'Dim exceptList As String() = {"Inbox_2000~2018", "Facebook"} '這幾個folder裡面有不知名無法處理的例外物件
        '' FixMe: 在Inbox_2000~2018目錄不知道有什麼奇怪東西, 一直有exception, 還try/catch不到, "Arithmetic operation resulted in an overflow."
        'If Not exceptList.Contains(folder.Name) Then
        '    Try
        '        ' 幹, 下面這行是殺小 為什麼這麼快??!!!
        '        Dim folderSize As Long = Await Task.Run(Function() folder.Items.Cast(Of Object)().Sum(Function(s) s.Size))
        '        ' 4. 將結果存入快取以便下次讀取
        '        folderSizeCache.TryAdd(folder, folderSize)
        '        Return folderSize
        '    Catch ex As OverflowException
        '        ' 在這裡處理溢出異常, 例如記錄日誌或返回一個特定的值，表示在計算資料夾大小時發生了溢出
        '        Console.WriteLine("計算資料夾大小時發生了溢出: " & ex.Message)
        '        Return -1 ' 或者返回其他適當的值
        '    Catch ex As System.Exception
        '        ' 在這裡處理其他類型的異常, 例如記錄日誌或者採取其他適當的措施
        '        Console.WriteLine("計算資料夾大小時發生了異常: " & ex.Message)
        '        Return -1 ' 或者返回其他適當的值
        '    Finally
        '    End Try
        'Else
        '    Dim foldersize As Long, intcount As Integer
        '    Try
        '        ' 使用 LINQ 擴充方法, 平行處理獲取每個郵件的大小
        '        Dim tasks As Task(Of Integer)() = folder.Items.Cast(Of Object)().Select(Function(item) GetMailSizeAsync(TryCast(item, MailItem))).Where(Function(t) t IsNot Nothing).ToArray()
        '        Dim results = Await Task.WhenAll(tasks)
        '        foldersize = results.Sum() ' 累加每個郵件的大小
        '    Catch ex As System.Exception
        '        ' 最後再不行就使用傳統迴圈
        '        'For Each item In folder.Folders
        '        '    If TypeOf item Is MailItem Then
        '        '        foldersize += Await GetMailSizeAsync(item)
        '        '        intcount += 1
        '        '        Invoke(Sub() ProgressBar1.Text = intcount & " / " & GetMailCountByMAPI_Old2(folder)) '更新讀取進度
        '        '    End If
        '        'Next
        '    End Try
        '    ' 更新讀取進度
        '    'Invoke(Sub() ProgressBar1.Text = intcount & " / " & GetMailCountByMAPINew(folder))    'Invoke 包 UI 更新 (Claude說不好)
        '    ProgressBar1.Text = intcount & " / " & GetMailCountByMAPINew(folder)
        '    Return foldersize
        'End If

    End Function
    Private Async Function GetMailSizeAsync(mailItem As MailItem) As Task(Of Integer)
        DebugForm.AddMessage3("Begin: ", mailItem.Name)
        Dim itemSize As Integer = 0
        'If mailSizeCache.TryGetValue(mailItem, itemSize) Then Return itemSize ' 檢查快取中是否已經存在郵件大小
        '' 使用 PR_MESSAGE_SIZE 或 MailItem.Size 遍歷所有郵件並加總
        'Dim PR_MESSAGE_SIZE As String = "http://schemas.microsoft.com/mapi/proptag/0x0E080003"
        'Try
        '    ' 進行非同步操作
        '    Dim sizeObject As Object = Await Task.Run(Function() mailItem.PropertyAccessor.GetProperty(PR_MESSAGE_SIZE))
        '    'MessageBox.Show(mailItem.Subject)
        '    If TypeOf sizeObject Is Integer AndAlso sizeObject > 0 Then
        '        itemSize = DirectCast(sizeObject, Integer)
        '    ElseIf sizeObject = 0 Then
        '        itemSize = mailItem.Size
        '    End If
        'Catch
        'End Try
        'mailSizeCache.TryAdd(mailItem, itemSize) ' 將郵件大小存入快取中
        Return itemSize

    End Function
    Private Sub ExpandTreeToDefaultInbox(treeview As TreeView)
        'DebugForm.AddMessage2("Begin: ", treeview.Name)
        '    If treeview.Nodes.Count > 0 Then
        '        treeview.Nodes(0).Expand()
        '        For i As Integer = 0 To treeview.Nodes.Count - 1
        '            Dim node As TreeNode = treeview.Nodes(0).Nodes(i)
        '            If node.Text.Contains("Inbox") Or node.Text.Contains("收件匣") Then
        '                If TypeOf treeview Is MultiSelectTreeView Then     '检查传入的treeview类型, 如果是MultiSelectTreeView，使用SelectedNodes属性
        '                    Dim multiSelectTreeView As MultiSelectTreeView = CType(treeview, MultiSelectTreeView)
        '                    multiSelectTreeView.ClearSelectedNodes()
        '                    multiSelectTreeView.AddNode(node)
        '                Else                                               '如果是普通的TreeView，使用原有的SelectedNode属性
        '                    treeview.SelectedNode = node : node.EnsureVisible()
        '                    treeview.Focus() : Exit Sub
        '                End If
        '            End If
        '        Next
        '    End If

    End Sub
    Private Async Sub TreeView1_AfterSelect(sender As Object, e As TreeViewEventArgs) 'Handles TreeView1.AfterSelect
        DebugForm.AddMessage3("Begin: ")
        '    Dim stopwatch As New Stopwatch() : stopwatch.Start() ' 開始計時
        '    ProgressBar1.Text = "" : ProgressBar2.Text = ""
        '    Cursor = Cursors.WaitCursor
        '    ' 在 TreeView 中選擇節點時, 更新 ListView 的內容
        '    ' todo: 是否需要: 按下後只顯示資料夾名稱裝一下, 乘機在背景計算, 再逐一補上foldercount/mailcount
        '    ' 5/21否決: 如果還在背景計算的時候, treeview又快速點選其他的folder就會出錯
        '    Dim selectedFolder = TryCast(e.Node.Tag, Outlook.Folder) ' 取得在 TreeView 中選擇的節點, 取得選擇的資料夾
        '    If selectedFolder IsNot Nothing Then
        '        Try
        '            ListView1.Items.Clear()
        '            'ListView1.Items.Add(GetInfoForListview(selectedFolder, False)) ' 顯示目前選中資料夾
        '            Dim item = Await GetInfoForListview(selectedFolder, False)  ' ✅ Await Async 版本
        '            If item IsNot Nothing Then ListView1.Items.Add(item)
        '        Catch
        '        End Try
        '        ListView1.BeginUpdate()
        '        Dim sortedFolders = GetSortedSubFolders(selectedFolder)         ' 取得所有子資料夾
        '        For Each subFolder In sortedFolders
        '            ' 5/21記錄: 改用addRange()但結果沒有比較快
        '            'ListView1.Items.Add(GetInfoForListview(subFolder, True))   ' 顯示所有子資料夾
        '            Dim item = Await GetInfoForListview(subFolder, True)        ' ✅ Await Async 版本
        '            If item IsNot Nothing Then ListView1.Items.Add(item)
        '        Next
        '        ListView1.EndUpdate()
        '        stopwatch.Stop() ' 停止計時, 顯示總共花費的時間
        '        If bln1stInit = True Then : bln1stInit = False
        '        Else : lblStatus2.Text = "更新統計郵件數量花費了 " & stopwatch.Elapsed.TotalSeconds.ToString("0.00") & " 秒。"
        '        End If
        '    End If
        '    lblStatus3.Text = sw1.Elapsed.TotalSeconds.ToString("0.00 / ") & sw2.Elapsed.TotalSeconds.ToString("0.00 / ") & sw3.Elapsed.TotalSeconds.ToString("0.00 / ") & sw4.Elapsed.TotalSeconds.ToString("0.00 / ") & sw5.Elapsed.TotalSeconds.ToString("0.00 / ")
        '    TreeView1.Enabled = True : TreeView1.Focus() : Cursor = Cursors.Default
        'End Sub    Private Async Sub TreeView1_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TreeView1.AfterSelect
        '    DebugForm.AddMessage2("Begin: ")
        '    Dim stopwatch As New Stopwatch() :  stopwatch.Start() ' 開始計時
        '    lblStatus1.Text = "" : lblStatus2.Text = ""
        '    Cursor = Cursors.WaitCursor
        '    ' 在 TreeView 中選擇節點時, 更新 ListView 的內容
        '    ' todo: 是否需要: 按下後只顯示資料夾名稱裝一下, 乘機在背景計算, 再逐一補上foldercount/mailcount
        '    ' 5/21否決: 如果還在背景計算的時候, treeview又快速點選其他的folder就會出錯
        '    Dim selectedFolder = TryCast(e.Node.Tag, Outlook.Folder) ' 取得在 TreeView 中選擇的節點, 取得選擇的資料夾
        'If selectedFolder IsNot Nothing Then
        'Try
        '            ListView1.Items.Clear()
        '            'ListView1.Items.Add(GetInfoForListview(selectedFolder, False)) ' 顯示目前選中資料夾
        '            Dim item = Await GetInfoForListview(selectedFolder, False)  ' ✅ Await Async 版本
        '            If item IsNot Nothing Then ListView1.Items.Add(item)
        '        Catch
        'End Try
        '        ListView1.BeginUpdate()
        '        Dim sortedFolders = GetSortedSubFolders(selectedFolder)         ' 取得所有子資料夾
        'For Each subFolder In sortedFolders
        '' 5/21記錄: 改用addRange()但結果沒有比較快
        ''ListView1.Items.Add(GetInfoForListview(subFolder, True))   ' 顯示所有子資料夾
        'Dim item = Await GetInfoForListview(subFolder, True)        ' ✅ Await Async 版本
        '            If item IsNot Nothing Then ListView1.Items.Add(item)
        '        Next
        '        ListView1.EndUpdate()
        '        stopwatch.Stop() ' 停止計時, 顯示總共花費的時間
        '        If bln1stInit = True Then :  bln1stInit = False
        '        Else :  lblStatus2.Text = "更新統計郵件數量花費了 " & stopwatch.Elapsed.TotalSeconds.ToString("0.00") & " 秒。"
        '        End If
        'End If
        '    lblStatus3.Text = sw1.Elapsed.TotalSeconds.ToString("0.00 / ") & sw2.Elapsed.TotalSeconds.ToString("0.00 / ") & sw3.Elapsed.TotalSeconds.ToString("0.00 / ") & sw4.Elapsed.TotalSeconds.ToString("0.00 / ") & sw5.Elapsed.TotalSeconds.ToString("0.00 / ")
        '    TreeView1.Enabled = True : TreeView1.Focus() : Cursor = Cursors.Default

    End Sub
    Private Async Function GetInfoForListview(folder As Outlook.Folder, Optional iamSub As Boolean = True) As Task(Of ListViewItem)
        DebugForm.AddMessage3("Begin: ", folder.Name)
        'lblStatus1.Text = "正在處理: " & folder.Name
        ''Dim s1Task = Task.Run(Function() GetTotalFolderCountAsync(folder)) 'sw1
        ''Dim s2Task = Task.Run(Function() GetTotalFolderCountAsync(folder))  'sw2 (5/21最後決定: 二個函數快慢互有變化, 但GetTotalFolderCountAsync()的穩定性較好, 比New()的標準差來得小, 所以決定使用這個)
        ''Dim s3Task = Task.Run(Function() GetMailCountByMAPI_Old(folder))  'sw3
        'Dim s4Task = Task.Run(Function() GetMailCountByMAPINew(folder))     'sw4 (5/21最後決定: 這個還是比較快, 不知道為什麼
        ''Dim s5Task = Task.Run(Function() GetMailCountByLINQNew(folder))    'sw5
        ''Await Task.WhenAll(s2Task, s3Task) ' 取消這一行就不會在treeview快速亂點的時候, 舊的還沒算完又顯示到下一個的畫面上了
        '' 2026/3/6: ✅ 改成 Async Function，直接 Await，不再用 Task.Run 包 COM 呼叫，不再用 .Result 阻塞
        'Dim s2Task As Integer = Await GetTotalFolderCountAsync(folder)
        ''Dim s4Task As Integer = Await GetMailCountByMAPINew(folder)
        'Try
        '    Dim FolderName As String = If(iamSub, " - " & folder.Name, folder.Name)
        '    Dim s1 As String = folder.Items.Count.ToString("###,###,##0")
        '    Dim s2 As String = s2Task.ToString("###,###,##0")
        '    Dim s3 As String = s4Task.Result.ToString("###,###,##0")
        '    Dim s4 As String = "", s4value As Integer
        '    If folderSizeCache.TryGetValue(folder, s4value) Then s4 = (s4value / 1024).ToString("###,###,###,##0KB")
        '    Return New ListViewItem({FolderName, s1, s2, s3, s4})
        'Catch
        'End Try
        Return Nothing

    End Function
    Private Async Sub TreeView2_AfterSelect(sender As Object, e As TreeViewEventArgs) 'Handles TreeView2.AfterSelect
        DebugForm.AddMessage3("Begin: TreeView2_AfterSelect()")
        '    ' 開始計時
        '    Dim stopwatch As New Stopwatch() : stopwatch.Start()
        '    lblStatus1.Text = "" : lblStatus2.Text = "" : Cursor = Cursors.WaitCursor
        '    ' 取得選擇的資料夾, 並初始化全域統計變數
        '    Dim selectedFolder As Outlook.Folder = TryCast(e.Node.Tag, Outlook.Folder)
        '    _intTotalMailCount = GetMailCountByMAPINew(selectedFolder)
        '    _intProcessedCount = 0
        '    ' 在背景執行計算, 傳回給yearCounts
        '    Dim yearCounts As ConcurrentDictionary(Of Integer, Integer) = Await GetYearCountsAsync_CL(selectedFolder, CheckSub2.Checked)
        '    ShowTab2Result(yearCounts)
        '    stopwatch.Stop() ' 停止計時, 顯示總共花費的時間
        '    StatusUpdate(sender, yearCounts, stopwatch)
        '    sender.Enabled = True : sender.Focus() : Cursor = Cursors.Default
        DebugForm.AddMessage3("End: TreeView2_AfterSelect()")

    End Sub
    Private Async Sub SimTreeView2_AfterSelect(sender As Object, e As TreeViewEventArgs) 'Handles SimTreeView2.AfterSelect
        DebugForm.AddMessage3("Begin: SimTreeView2_AfterSelect()")
        '    '' 開始計時
        '    'Dim stopwatch As New Stopwatch() : stopwatch.Start()
        '    'lblStatus1.Text = "" : lblStatus2.Text = "" : Cursor = Cursors.WaitCursor
        '    '' 取得選擇的資料夾, 並初始化全域統計變數
        '    'Dim selectedFolder As Outlook.Folder = TryCast(e.Node.Tag, Outlook.Folder)
        '    '_intTotalMailCount = GetMailCountByMAPINew(selectedFolder)
        '    '_intProcessedCount = 0
        '    '''' 取得選擇的資料夾(清單), 開始逐個計算, 再加總到yearCounts
        '    '''Dim selectedNodes As List(Of TreeNode) = sender.SelectedNodes
        '    '''Dim yearCounts As Dictionary(Of Integer, Integer) = Nothing
        '    '''Dim selectedFolder As Outlook.Folder
        '    '''For Each node As TreeNode In selectedNodes
        '    '''    selectedFolder = TryCast(node.Tag, Outlook.Folder)
        '    '''    _intTotalMailCount += GetMailCountByMAPINew(selectedFolder)
        '    '''Next
        '    '''Try
        '    '''    For Each node As TreeNode In selectedNodes
        '    '''        selectedFolder = TryCast(node.Tag, Outlook.Folder)
        '    '''        If selectedFolder Is Nothing Then Continue For
        '    '''        Dim selectedYearCounts As Dictionary(Of Integer, Integer) = Await GetYearCountsAsync_CL(selectedFolder, CheckSub2.Checked)
        '    '''        yearCounts = MergeDictionaries(yearCounts, selectedYearCounts)
        '    '''    Next
        '    '''Catch ' todo: 若在統計期間, 快速改變treeview的點選節點, 經常會發生selectedNodes被改變的例外, 該如何捕捉處理??
        '    '''End Try
        '    '' 在背景執行計算, 傳回給yearCounts
        '    'Dim yearCounts As Dictionary(Of Integer, Integer) = Await GetYearCountsAsync_CL(selectedFolder, CheckSub2.Checked)
        '    'If yearCounts IsNot Nothing Then ShowResult(yearCounts)
        '    'stopwatch.Stop() ' 停止計時, 顯示總共花費的時間
        '    'StatusUpdate(sender, yearCounts, stopwatch)
        '    'sender.Enabled = True : sender.Focus() : Cursor = Cursors.Default
        DebugForm.AddMessage3("End: SimTreeView2_AfterSelect()")

    End Sub
    Private Async Sub MyTreeView2_AfterSelect(sender As Object, e As TreeViewEventArgs) 'Handles MyTreeView2.AfterSelect
        DebugForm.AddMessage3("Begin MyTreeView2_AfterSelect: ")
        '    ' 開始計時
        '    Dim stopwatch As New Stopwatch() : stopwatch.Start()
        '    lblStatus1.Text = "" : lblStatus2.Text = "" : Cursor = Cursors.WaitCursor
        '    ' 取得選擇的資料夾(清單), 開始逐個計算, 再加總到yearCounts
        '    Dim selectedNodes As List(Of TreeNode) = sender.SelectedNodes
        '    Dim yearCounts As Dictionary(Of Integer, Integer) = Nothing
        '    Try
        '        For Each node In selectedNodes
        '            Dim selectedFolder As Outlook.Folder = TryCast(node.Tag, Outlook.Folder)
        '            If selectedFolder Is Nothing Then Continue For
        '            Dim selectedYearCounts As Dictionary(Of Integer, Integer) = Await GetYearCountsAsync_CL(selectedFolder, CheckSub2.Checked)
        '            yearCounts = MergeDictionaries(yearCounts, selectedYearCounts)
        '        Next
        '    Catch ' todo: 若在統計期間, 快速改變treeview的點選節點, 經常會發生selectedNodes被改變的例外, 該如何捕捉處理??
        '    End Try
        '    ShowResult(yearCounts)
        '    stopwatch.Stop() ' 停止計時, 顯示總共花費的時間
        '    StatusUpdate(sender, yearCounts, stopwatch)
        '    sender.Enabled = True : sender.Focus() : Cursor = Cursors.Default
        DebugForm.AddMessage3("End: ")

    End Sub
    Private Sub CheckSub2_CheckedChanged(sender As Object, e As EventArgs) ' Handles CheckSub2.CheckedChanged
        DebugForm.AddMessage3("Begin: ", sender.Name)
        '    ' 如果有選定節點,則手動呼叫 TreeView2_AfterSelect 事件
        '    Dim selectedNode = TreeView2.SelectedNode ' 獲取目前選定的 TreeNode
        '    If TreeView2.Visible = True And selectedNode IsNot Nothing Then TreeView2_AfterSelect(TreeView2, New TreeViewEventArgs(selectedNode))
        '    'If SimTreeView2.Visible = True Then
        '    '    If SimTreeView2.SelectedNodes.Count > 1 Then
        '    '        ' 使用第一個選中的節點作為參數
        '    '        SimTreeView2_AfterSelect(SimTreeView2, New TreeViewEventArgs(SimTreeView2.SelectedNodes(0)))
        '    '    Else
        '    '        SimTreeView2_AfterSelect(SimTreeView2, New TreeViewEventArgs(SimTreeView2.SelectedNode))
        '    '    End If
        '    '    SimTreeView2.Focus()
        '    'End If

    End Sub
    Private Function QueueAllFolderNodes(treeView As TreeView) As List(Of TreeNode)
        ' 獲取 TreeView 中所有資料夾節點的列表（廣度優先）
        Dim nodeList As New List(Of TreeNode)
        Dim queue As New Queue(Of TreeNode)
        'For Each node As TreeNode In treeView.Nodes
        '    queue.Enqueue(node)
        'Next
        While queue.Count > 0
            Dim currentNode As TreeNode = queue.Dequeue()
            nodeList.Add(currentNode)
            For Each childNode As TreeNode In currentNode.Nodes
                queue.Enqueue(childNode)
            Next
        End While
        Return nodeList

    End Function
    Private Function GetAllFolderNodesA(rootFolder As Outlook.Folder) As List(Of Outlook.Folder)
        ' 獲取 rootFolder 以下所有資料夾節點的列表（廣度優先）
        Dim result As New List(Of Outlook.Folder)() ' 初始化一個列表來存儲結果
        Dim queue As New Queue(Of Outlook.Folder)() ' 初始化一個佇列來進行廣度優先搜索
        ' 將根文件夾加入佇列中
        queue.Enqueue(rootFolder)
        ' 進行廣度優先搜索
        While queue.Count > 0
            ' 從佇列中取出當前文件夾
            Dim currentFolder As Outlook.Folder = queue.Dequeue()
            ' 將當前文件夾加入結果列表
            result.Add(currentFolder)
            ' 將所有子文件夾加入佇列中
            For Each subFolder As Outlook.Folder In currentFolder.Folders
                queue.Enqueue(subFolder)
            Next
        End While
        ' 返回結果列表
        Return result

    End Function
    Private Function GetAllFolderNodesB(rootFolder As Outlook.Folder, Optional maxDepth As Integer? = Nothing) As List(Of Outlook.Folder)
        Dim folderList As New List(Of Outlook.Folder)()
        Dim queue As New Queue(Of Tuple(Of Outlook.Folder, Integer))()
        ' 初始化佇列，將根文件夾及其深度（0）加入佇列中
        queue.Enqueue(Tuple.Create(rootFolder, 0))
        ' 進行廣度優先搜索
        While queue.Count > 0
            ' 從佇列中取出當前文件夾及其深度
            Dim current As Tuple(Of Outlook.Folder, Integer) = queue.Dequeue()
            Dim currentFolder As Outlook.Folder = current.Item1
            Dim currentDepth As Integer = current.Item2
            ' 將當前文件夾加入結果列表
            folderList.Add(currentFolder)
            ' 如果達到最大深度，則不再加入子文件夾
            If Not maxDepth.HasValue OrElse currentDepth < maxDepth.Value Then
                ' 將所有子文件夾及其深度（當前深度 + 1）加入佇列中
                For Each subFolder As Outlook.Folder In currentFolder.Folders
                    queue.Enqueue(Tuple.Create(subFolder, currentDepth + 1))
                Next
            End If
        End While
        Return folderList

    End Function
    Private Function CountByYears(selectedFolder As Outlook.Folder) As Dictionary(Of Integer, Integer)
        '' 建立一個字典來存儲每個年份的郵件數量
        'Dim yearCounts As New Dictionary(Of Integer, Integer)()
        'If selectedFolder.Items.Count <= 0 OrElse Not selectedFolder.DefaultItemType = Outlook.OlItemType.olMailItem Then Return yearCounts
        '' ===========================================
        '' 使用 Restric 方法快速獲取郵件項目的日期和數量
        '' ===========================================
        'Dim intCount As Integer
        'For year As Integer = Get1stYear(selectedFolder) To Date.Today.Year
        '    Dim startDate As New Date(year, 1, 1, 0, 0, 0) ' 建立當年的起始日期和結束日期
        '    Dim endDate As New Date(year, 12, 31, 23, 59, 59) ' 設置結束日期的時間為23:59:59
        '    ' 建立篩選字串, 使用 Restrict 方法篩選郵件項目
        '    Dim restrictFilter As String = $"[ReceivedTime] >= '{startDate}' AND [ReceivedTime] <= '{endDate}'"
        '    Dim restrictedItems As Outlook.Items = selectedFolder.Items.Restrict(restrictFilter)
        '    ' 統計郵件數量並存入字典
        '    Dim intCountOfTheYear As Integer = restrictedItems.Count
        '    If intCountOfTheYear > 0 Then
        '        yearCounts.Add(year, intCountOfTheYear) : intCount += intCountOfTheYear
        '        Invoke(Sub() lblStatus1.Text = intCount & " / " & selectedFolder.Items.Count) '更新讀取進度
        '    End If
        'Next
        'Return yearCounts ' 回傳計算結果

    End Function
    Private Function CountByYearsParallel(selectedFolder As Outlook.Folder) As Dictionary(Of Integer, Integer)
        '' 建立一個字典來存儲每個年份的郵件數量
        'Dim yearCounts As New ConcurrentDictionary(Of Integer, Integer)()
        'If selectedFolder.Items.Count <= 0 OrElse Not selectedFolder.DefaultItemType = Outlook.OlItemType.olMailItem Then
        '    Return yearCounts.ToDictionary(Function(pair) pair.Key, Function(pair) pair.Value)
        'End If
        '' ============================================================================
        '' 使用 Restric 方法快速獲取郵件項目的日期和數量, 使用Parallel.For平行處理Restrict
        '' ============================================================================
        'Parallel.For(Get1stYear(selectedFolder), Date.Today.Year + 1,
        '             Sub(year)
        '                 Dim startDate As New Date(year, 1, 1, 0, 0, 0)
        '                 Dim endDate As New Date(year, 12, 31, 23, 59, 59)
        '                 Dim filter As String = $"[ReceivedTime] >= '{startDate}' AND [ReceivedTime] <= '{endDate}'"
        '                 Dim restrictedItems As Outlook.Items = selectedFolder.Items.Restrict(filter)
        '                 ' 統計郵件數量並存入字典
        '                 Dim intCountOfTheYear As Integer = restrictedItems.Count
        '                 If intCountOfTheYear > 0 Then
        '                     yearCounts.TryAdd(year, intCountOfTheYear)
        '                     invoke(Sub() lblStatus1.Text = yearCounts.Values.Sum() & " / " & selectedFolder.Items.Count)
        '                 End If
        '             End Sub)
        'Return yearCounts.ToDictionary(Function(pair) pair.Key, Function(pair) pair.Value)

    End Function
    Private Async Function GetYearCountsAsync_CL(selectedFolder As Outlook.Folder, includeSubFolders As Boolean) As Task(Of ConcurrentDictionary(Of Integer, Integer))
        '    ' ===============================
        '    ' 這個函數是整個統計流程的核心, 包含快取機制和遞迴處理子資料夾的邏輯
        '    ' 用來取得指定資料夾及其子資料夾中每個年份的郵件數量統計, 並快取結果以提升效能
        '    ' ===============================
        '    ' 建立一個唯一的快取鍵值,包含資料夾的 FolderPath 和是否包含子資料夾的選項
        '    Dim cacheKey As String = selectedFolder.FolderPath & "_" & includeSubFolders.ToString()
        '    If yearCountsCache.ContainsKey(cacheKey) Then
        '        Dim value As ConcurrentDictionary(Of Integer, Integer) = yearCountsCache(cacheKey)
        '        _intProcessedCount += value.Values.Sum  ' 若已存在快取, 也要把快取的郵件數量加總至已處理進度
        '        UpdateCounterProgress(_intProcessedCount, selectedFolder.Items.Count, includeSubFolders)                        '✅ 直接呼叫，不要Task.Run，Claude說不好
        '        Return value
        '    End If
        '    ' 如果快取中沒有結果, 才開始進行計算
        '    Dim yearCounts As New ConcurrentDictionary(Of Integer, Integer)   ' 建立一個字典來存儲每個年份的郵件數量
        '    yearCounts = Await CountMailByYearAsync_CL2(selectedFolder, False)
        '    If includeSubFolders Then   ' 遞迴包含所有子資料夾
        '        For Each childFolder As Outlook.Folder In selectedFolder.Folders
        '            Dim childYearCounts As ConcurrentDictionary(Of Integer, Integer) = Await GetYearCountsAsync_CL(childFolder, True)
        '            yearCounts = MergeDictionaries(yearCounts, childYearCounts)
        '            UpdateCounterProgress(_intProcessedCount, selectedFolder.Items.Count, includeSubFolders)                        '✅ 直接呼叫 (Claude說Task.Run 裡呼叫 UI不好)
        '        Next
        '        ' 6/1進度顯示終於正確:
        '        ' 1. 使用_intProcessedCount和_intTotalMailCount, 二個全域變數來追踪統計完整郵件數量
        '        ' 2. 不含子資料夾的時候, 每一年份的restrict就統計一次進度顯示更新
        '        ' 3. 包含子資料夾的時候, 只有每算完一個childFolder才統計更新一次
        '    Else
        '    End If
        '    yearCountsCache(cacheKey) = yearCounts ' 將結果存入快取 (6/3記錄: 有重覆key值時用Dictionary.Add()會發生例外錯誤, 所以改用 "=" 來賦值)
        '    Return yearCounts

    End Function
    Private Async Function CountMailByYearAsync_CL2(folder As Outlook.Folder, includeSubFolders As Boolean) As Task(Of ConcurrentDictionary(Of Integer, Integer))
        '    Dim yearCounts As New ConcurrentDictionary(Of Integer, Integer)()
        '    Dim intTotalMailCount As Integer = folder.Items.Count
        '    If Not includeSubFolders AndAlso intTotalMailCount = 0 Then Return yearCounts   '如果自己的個數為0, 也沒有勾subfolder,就不需計算直接返回
        '    For year As Integer = Find1stYear(folder) To Date.Today.Year  ' ✅ COM 呼叫在 UI 執行緒
        '        Try
        '            ' =========================================================
        '            ' 使用 Restrict方法快速統計郵件項目的日期和數量
        '            ' 修正: 拿掉 Task.Run，COM 呼叫全程在 UI 執行緒 (原本就必須如此)
        '            ' 使用 Await Task.Yield() 在各年份之間讓出控制權，UI 仍可回應
        '            ' Restrict().Count 是 MAPI 索引查詢，速度本來就快，不需要平行化
        '            ' 2026/3/6 by Claude Code
        '            ' =========================================================
        '            Dim restrictFilter As String = BuildFilterString2(year)
        '            Dim restrictedItems As Outlook.Items = folder.Items.Restrict(restrictFilter) ' ✅ COM 在 UI 執行緒
        '            Dim intCountOfTheYear As Integer = restrictedItems.Count                     ' ✅ COM 在 UI 執行緒
        '            If intCountOfTheYear > 0 Then yearCounts.TryAdd(year, intCountOfTheYear)       ' 統計郵件數量並存入字典
        '            _intProcessedCount += intCountOfTheYear
        '            If restrictedItems IsNot Nothing Then Marshal.ReleaseComObject(restrictedItems)  ' ✅ COM 物件使用完畢後釋放，避免記憶體洩漏
        '        Catch ex As System.Exception
        '            DebugForm.AddMessage2(ex.Message & ": " & ex.Source & ": " & ex.Data.ToString)
        '        Finally
        '        End Try
        '        If Not includeSubFolders Then UpdateCounterProgress(_intProcessedCount, intTotalMailCount, includeSubFolders)  ' ✅ UI 操作在 UI 執行緒
        '        If year Mod 10 = 0 Then Await Task.Yield()  ' ✅ 每一小段時間讓出控制權，UI 可回應，不阻塞
        '        ' 6/1進度顯示終於正確:
        '        ' 1. 使用_intProcessedCount和_intTotalMailCount, 二個全域變數來追踪統計完整郵件數量
        '        ' 2. 不含子資料夾的時候, 則只計算此目錄的進度, 每一年份的restrict就統計一次進度顯示更新
        '        ' 3. 包含子資料夾的時候, 只有每算完一個childFolder才統計更新一次
        '    Next
        '    Await Task.Yield()  ' ✅ 每一小段時間讓出控制權，UI 可回應，不阻塞
        '    Return yearCounts

    End Function
    Private Sub UpdateCounterProgress(ByRef processedCount As Integer, selectedFolderItemCount As Integer, includeSubFolders As Boolean)
        '    ' 6/1進度顯示終於正確:
        '    ' 1. 使用_intProcessedCount和_intTotalMailCount, 二個全域變數來追踪統計完整郵件數量
        '    ' 2. 不含子資料夾的時候, 每一年份的restrict就統計一次進度顯示更新
        '    ' 3. 包含子資料夾的時候, 只有每算完一個childFolder才統計更新一次
        '    Dim useGlobalTotal As Boolean = CheckSub2.Checked OrElse includeSubFolders  ' 二個只要一個成立就使用全域
        '    Dim totalCount As Integer = If(useGlobalTotal, _intTotalMailCount, selectedFolderItemCount)
        '    'If CheckSub2.Checked Then
        '    '    totalCount = _intTotalMailCount
        '    'Else
        '    '    totalCount = If(includeSubFolders, _intTotalMailCount, selectedFolderItemCount)
        '    'End If
        '    'Select Case TreeView2.Visible
        '    '    Case 1  ' 單一資料夾
        '    '    Case 2  ' 遞迴資料夾
        '    '    Case 3  ' 多選資料夾
        '    'End Select
        '    lblStatus1.Text = $"{processedCount} / {totalCount}"

    End Sub
    ' 舊的 Button3_Click 事件處理器，保留以供參考，已被改寫為上面新的版本
    Private Async Sub Button8_Click(sender As Object, e As EventArgs) ' Handles Button8.Click
        DebugForm.AddMessage3("Begin: ")
        'Dim stopwatch As New Stopwatch() : stopwatch.Start()
        'ListView3.Items.Clear()
        'TextBox3.Enabled = False : lblStatus1.Text = "" : lblStatus2.Text = ""
        'Button8.Enabled = False : Button3_Stop.BringToFront() : Button3_Stop.Visible = True : blnButton3_Stop = False ' 顯示停止鈕
        'Cursor = Cursors.WaitCursor
        '' 使用 Restrict 方法篩選過濾具有附件的郵件項目
        'Dim filter1 As String = BuildFilterString3() ' 設定restric 篩選的字串
        'If filter1.Length = 0 Then
        '    TextBox3.Enabled = True : Button8.Enabled = True : Exit Sub
        'End If
        'Dim selectedFolder As Outlook.Folder = TryCast(TreeView3.SelectedNode.Tag, Outlook.Folder)
        'Dim filteredBySize As Outlook.Items = selectedFolder.Items.Restrict(filter1)
        '' 在listview3中顯示篩選出來的郵件集合
        '' todo: 1. 測試使用attachments:presentation.pptx語法來restrict filter, 而不是for each
        '' todo: 2. 真的不能用 urn:schemas:httpmail:attachmentfilename" LIKE '%??
        '' todo: 3. for each item + for each attach 的迴圈, 是否改用items.setColumn 來限制欄位加快速度?
        '' todo: 4. 有無檔名+有無大小限制, 各種組合, 是否有更好的邏輯判斷式, 或是判斷順序?
        'If Not CheckAttachName.Checked OrElse TextBox3.Text.Length = 0 Then
        '    Await ShowTab3Result(filteredBySize) '不管附件檔名, 只篩附件和大小
        'Else ' 進一步篩選附件檔名, 及指定關鍵字
        '    Dim filteredByKeyword As List(Of MailItem) = Await FilterAttachByKeywordAsync3(filteredBySize, TextBox3.Text)
        '    Await ShowTab3Result(filteredByKeyword)
        'End If
        '' todo: 5. 試試以下幾個DASL 屬性 (是MAPI Extention??)
        ''DASL Name http: //schemas.microsoft.com/mapi/proptag/0x3707001F
        ''PR_ATTACH_FILENAME         0x3704001E (0x3704001F for Unicode) 8.3 naming
        ''PR_ATTACH_LONG_FILENAME    0x3707001E (0x3707001F for Unicode) Platforms that support long filenames
        ''''Redemption Outlook Attachment Property Reference 這個有免費版可用嗎? (https://www.dimastr.com/redemption/OutlookAttachmentProperties.htm)
        ''''rSession = CreateObject("Redemption.RDOSession")
        ''''rSession.MAPIOBJECT = Application.Session.MAPIOBJECT
        ''''rItem = rSession.GetRDOObjectFromOutlookObject(Item)
        ''''attach = rItem.Attachments(1)
        '''''PR_ATTACH_LONG_FILENAME_W
        ''''attach.Fields("http://schemas.microsoft.com/mapi/proptag/0x3707001F") = "whatever.pdf"
        '''''PR_ATTACH_FILENAME_W
        ''''attach.Fields("http://schemas.microsoft.com/mapi/proptag/0x3704001F") = "whatever.pdf"
        '''''PR_DISPLAY_NAME_W
        ''''attach.Fields("http://schemas.microsoft.com/mapi/proptag/0x3001001F") = "whatever.pdf"
        'stopwatch.Stop() ' 停止計時, 顯示總共花費的時間
        'lblStatus2.Text = "篩選附件檔名花費了 " & stopwatch.Elapsed.TotalSeconds.ToString("0.00") & " 秒。(" & (filteredBySize.Count / stopwatch.Elapsed.TotalSeconds).ToString("###,##0") & "/sec)"
        'TextBox3.Enabled = True : Button8.Enabled = True : Button3.BringToFront() : Button3_Stop.Visible = False
        'Cursor = Cursors.Default
        ''ListView3.Items(0).Selected = True : ListView3.Items(0).Focused = True

    End Sub
    Private Async Function FilterAttachByKeywordAsync3(restrictedItems As Outlook.Items, keyword As String) As Task(Of List(Of MailItem))
        DebugForm.AddMessage3("Begin: ", keyword)
        Dim keywordLower As String = keyword.ToLower        ' ✅ 拉到迴圈外
        Dim filteredItems As New List(Of MailItem)
        Dim intCount As Integer = 0
        Dim totalCount As Integer = restrictedItems.Count   ' ✅ 拉到迴圈外
        'Dim item As Object = restrictedItems.GetFirst()
        'While item IsNot Nothing
        '    If blnButton3_Stop Then Exit While
        '    Try
        '        Dim mail As MailItem = TryCast(item, MailItem)
        '        If mail IsNot Nothing AndAlso mail.Attachments IsNot Nothing Then  ' ✅ 雙重 null 檢查
        '            If mail.Attachments.Cast(Of Outlook.Attachment)().Any(
        '                Function(a) a.FileName.ToLower.Contains(keywordLower)) Then filteredItems.Add(mail)
        '        End If
        '    Catch ex As System.Exception
        '        DebugForm.AddMessage2("FilterAttach Error: ", ex.Message)
        '    End Try
        '    intCount += 1
        '    If intCount Mod 10 = 0 Then
        '        Dim captured As Integer = intCount
        '        Me.Invoke(Sub()
        '                      lblStatus1.Text = captured & " / " & totalCount
        '                  End Sub)
        '        Await Task.Yield()
        '    End If
        '    item = restrictedItems.GetNext()
        'End While
        Return filteredItems

    End Function
    Private Async Function ShowTab3Result(Of T)(data As T, Optional listview As ListView = Nothing) As Task
        DebugForm.AddMessage3("Begin: ", data.ToString)
        'If listview Is Nothing Then listview = ListView3
        'listview.Items.Clear()
        ' 根據 T 的型態進行不同的處理
        'Dim inputItems As Object = Nothing
        'If GetType(T) Is GetType(Outlook.Items) Then ' 處理 Outlook.Items 的情況
        '    inputItems = TryCast(data, Outlook.Items)
        'ElseIf GetType(T) Is GetType(List(Of Outlook.MailItem)) Then ' 處理 List(Of Outlook.MailItem) 的情況
        '    inputItems = TryCast(data, List(Of Outlook.MailItem))
        'End If
        'sw5.Start()
        '' todo: 建立mailitemCache 1st priority (因為用到二次)
        'Dim strKeep As String = lblStatus1.Text
        'Dim countSum As Integer
        'Dim itemList As New List(Of ListViewItem) ' 創建一個 ListViewItem 的 List 來暫存所有要添加的項目
        'For Each item As MailItem In inputItems
        '    If blnButton3_Stop = True Then Exit For
        '    Try
        '        Dim cItem As New ListViewItem({item.Subject, item.Size.ToString("###,###,##0"), item.ReceivedTime.ToShortDateString, item.Sender.Name.ToString, item.Attachments.Count, item.EntryID})
        '        itemList.Add(cItem)
        '    Catch ex As System.Exception
        '        DebugForm.AddMessage2("Exception: ", ex.Message)
        '        DebugForm.AddMessage2("Error: ", item.Subject)
        '    End Try
        '    countSum += 1
        '    If countSum Mod 5 = 0 Then
        '        lblStatus1.Text = strKeep & " (" & countSum & " / " & inputItems.Count & ")" ' 更新讀取進度的計算方式
        '        Await Task.Yield
        '    End If
        'Next item
        'sw5.Stop()
        '' 5/20, 改用AddRange()將所有項目一次性添加到 listview
        'listview.Items.AddRange(itemList.ToArray())
        'lblStatus3.Text = sw5.Elapsed.TotalSeconds.ToString("0.00, ") ' 將所有項目一次性添加到 listview
        'If inputItems.Count = 0 Then listview.Items.Add("找不到符合條件的郵件")
        '''' todo: 嘗試使用 PLINQ 並行處理每個 MailItem
        '''Dim totalCount As Integer = inputItems.Count()
        '''Dim itemList As New List(Of ListViewItem)
        '''Await inputItems.AsParallel().ForAll(
        '''Sub(item)
        '''    If blnButton3_Stop Then Return
        '''    sw5.Start()
        '''    Try
        '''        Dim cItem As New ListViewItem({item.Subject, item.Size.ToString("###,###,##0"), item.ReceivedTime.ToShortDateString, item.Sender.Name.ToString, item.Attachments.Count, item.EntryID})
        '''        SyncLock itemList
        '''            itemList.Add(cItem) : countSum += 1
        '''            If countSum Mod 100 = 0 Then Invoke(Sub() lblStatus1.Text = strKeep & " (" & countSum & " / " & totalCount & ")")
        '''        End SyncLock
        '''    Catch
        '''    End Try
        '''    sw5.Stop()
        '''End Sub)
        '''' 在 UI 線程上添加項目到 listview
        '''listview.Invoke(Sub() listview.Items.AddRange(itemList.ToArray()))
        '''lblStatus3.Text = sw5.Elapsed.TotalSeconds.ToString("0.00, ")
        '''If inputItems.Count = 0 Then listview.Items.Add("找不到符合條件的郵件")

    End Function
    Private Function FilterAttachByKeyword(restrictedItems As Outlook.Items, keyword As String) As List(Of Outlook.MailItem)
        DebugForm.AddMessage3("Begin:", keyword)
        ' 進一步篩選郵件中附件檔名包含指定關鍵字的郵件
        Dim filteredItems As New List(Of Outlook.MailItem)
        Dim intCount As Integer
        '' 遍歷篩選後的郵件項目 (使用for each item)
        'For Each item As MailItem In restrictedItems
        '    Try
        '        Dim mailItem As MailItem = DirectCast(item, MailItem)
        '        For Each attachment As Attachment In mailItem.Attachments
        '            If attachment.FileName.ToLower.Contains(keyword.ToLower) Then
        '                filteredItems.Add(mailItem) : Exit For
        '            End If
        '        Next
        '        intCount += 1: Invoke(Sub() lblStatus1.Text = intCount & " / " & restrictedItems.Count)'更新讀取進度
        '    Catch
        '    End Try
        'Next
        ' 遍歷篩選後的郵件項目 (使用items.getfirst/getnext)
        'Dim mailItem As Outlook.MailItem = restrictedItems.GetFirst()
        'While mailItem IsNot Nothing
        '    Try
        '        For Each attachment As Outlook.Attachment In mailItem.Attachments
        '            If attachment.FileName.ToLower.Contains(keyword.ToLower) Then filteredItems.Add(mailItem) : Exit For
        '        Next
        '        intCount += 1 : invoke(Sub() lblStatus1.Text = intCount & " / " & restrictedItems.Count) ' 需要改善: 更新讀取進度的計算方式(分母用GetTotalMailCount?)
        '    Catch
        '    End Try
        '    mailItem = restrictedItems.GetNext()
        'End While
        Return filteredItems

    End Function
    Private Async Function FilterAttachByKeywordAsync(restrictedItems As Outlook.Items, keyword As String) As Task(Of List(Of Outlook.MailItem))
        DebugForm.AddMessage3("Begin:", keyword)
        Dim filteredItems As New List(Of Outlook.MailItem)()
        ' 使用 Parallel.ForEach 並行處理, 遍歷篩選後的郵件項目
        Await Task.Run(Sub()
                           'Dim intCount As Integer = 0
                           'Parallel.ForEach(restrictedItems.OfType(Of Outlook.MailItem),
                           'Sub(item)
                           '    Try
                           '        ' todo: 建立mailitemCache 2nd priority (因為用到一次)
                           '        ' todo: 使用PR_ATTACH_LONG_FILENAME_W ?? PR_ATTACH_SIZE??
                           '        If TypeOf item Is Outlook.MailItem Then
                           '            Dim mailItem As Outlook.MailItem = DirectCast(item, Outlook.MailItem)
                           '            ' 檢查郵件項目的附件是否包含指定關鍵字
                           '            ' 不使用for each遍歷, 只在比對名稱符合之後才取出mailitem (快了10%)
                           '            If mailItem.Attachments.Cast(Of Outlook.Attachment)().Any(Function(attachment) attachment.FileName.ToLower.Contains(keyword.ToLower)) Then filteredItems.Add(mailItem)
                           '            intCount += 1
                           '            'Interlocked.Increment(intCount) '確保了在每次迭代中，intCount 變數都會以原子方式增加，從而實現正確的計數。這樣做確保了多線程下的安全性，避免多個執行緒可能會同時訪問和修改 intCount 變數
                           '        End If
                           '    Catch ex As System.Exception
                           '        Invoke(Sub() Form1.ListView3.Items.Add("Error Found: " & ex.Message & " / " & ex.Source & " / " & ex.ToString))
                           '        Exit Sub
                           '        'Catch ex2 As InnerException
                           '    End Try
                           '    Invoke(Sub() lblStatus1.Text = intCount & " / " & restrictedItems.Count) ' 需要改善: 更新讀取進度的計算方式(分母用GetTotalMailCount?)
                           'End Sub)
                       End Sub)
        Return filteredItems

    End Function
    Private Async Function FilterAttachByKeywordAsync2(restrictedItems As Outlook.Items, keyword As String) As Task(Of List(Of Outlook.MailItem))
        DebugForm.AddMessage3("Begin:", keyword)
        Dim filteredItems As New List(Of Outlook.MailItem)
        Dim intCount As Integer
        ' 使用集合的非同步方法來獲取第一個郵件物件
        Dim mailItem As Outlook.MailItem = Await Task.Run(Function() restrictedItems.GetFirst())
        'While mailItem IsNot Nothing
        '    Try
        '        For Each attachment As Outlook.Attachment In mailItem.Attachments
        '            If attachment.FileName.ToLower.Contains(keyword.ToLower) Then filteredItems.Add(mailItem) : Exit For
        '        Next
        '        intCount += 1 : Invoke(Sub() lblStatus1.Text = intCount & " / " & restrictedItems.Count)  '更新讀取進度
        '    Catch
        '    End Try
        '    ' 使用集合的非同步方法來獲取下一個郵件物件
        '    mailItem = Await Task.Run(Function() restrictedItems.GetNext())
        'End While
        Return filteredItems

    End Function
    Private Async Function ShowResultToListview3Async2(Of T)(data As T) As Task
        ''=======================================================================
        '' 用了dictionary 跟info structure當作cache, 但速度並沒有提升, 不知道為什麼
        ''=======================================================================
        'ListView3.Invoke(Sub() ListView3.Items.Clear()) ' 在 UI 線程上清空 listview3 中的內容
        '根據 T 的型態進行不同的處理
        Dim inputItems As Object = Nothing
        If GetType(T) Is GetType(Outlook.Items) Then ' 處理 Outlook.Items 的情況
            inputItems = TryCast(data, Outlook.Items)
        ElseIf GetType(T) Is GetType(List(Of Outlook.MailItem)) Then ' 處理 List(Of Outlook.MailItem) 的情況
            inputItems = TryCast(data, List(Of Outlook.MailItem))
        End If
        '' todo: 使用SetColum來限定欄位提高效能
        'Dim strKeep As String = lblStatus1.Text
        'Dim countSum As Integer
        'If inputItems.Count > 10 Then ListView3.BeginUpdate() '避免更新過久
        'Await Task.Run(Sub()
        '                   For Each item As MailItem In inputItems
        '                       Try ' 檢查郵件是否在快取中，如果在，則直接使用快取的值，否則計算並更新快取
        '                           sw5.Start()
        '                           Dim entryKey As String = item.EntryID
        '                           Dim mailItemInfo As CachedMailItemInfo = Nothing
        '                           Dim value As CachedMailItemInfo = Nothing
        '                           If mailItemCache.TryGetValue(entryKey, value) Then : mailItemInfo = value
        '                           Else
        '                               Dim subject As String = item.Subject
        '                               Dim size As Long = item.Size
        '                               Dim receivedTime As Date = item.ReceivedTime
        '                               Dim senderName As String = item.Sender.Name.ToString
        '                               Dim attachmentCount As Integer = item.Attachments.Count
        '                               Dim entryID As String = item.EntryID
        '                               mailItemInfo = New CachedMailItemInfo(subject, size, receivedTime, senderName, attachmentCount, entryID)
        '                               mailItemCache.Add(entryKey, mailItemInfo) '更新快取
        '                           End If
        '                           sw5.Stop() : lblStatus3.Text = sw5.Elapsed.TotalSeconds.ToString("0.00, ")
        '                           ' 將快取的值添加到 listview3 中
        '                           Dim listItem As New ListViewItem({mailItemInfo.Subject, mailItemInfo.Size.ToString("###,###,##0"), mailItemInfo.ReceivedTime.ToShortDateString, mailItemInfo.Sender, mailItemInfo.AttachmentCount, mailItemInfo.EntryID})
        '                           ListView3.Invoke(Sub() ListView3.Items.Add(listItem))
        '                           countSum += 1
        '                           Invoke(Sub() lblStatus1.Text = strKeep & " (" & countSum & " / " & inputItems.Count & ")") ' 需要改善: 更新讀取進度的計算方式(分母用GetTotalMailCount?)
        '                       Catch
        '                       End Try
        '                   Next item
        '               End Sub)
        'If inputItems.Count = 0 Then ListView3.Items.Add("符合條件的郵件項目為 0")
        'ListView3.EndUpdate()

    End Function
    Private Async Function ShowResultToListview3Async3(Of T)(data As T) As Task
        ''==============================================================
        '' 用了dictionary 跟tuple當作cache, 但速度並沒有提升, 不知道為什麼
        ''==============================================================
        'ListView3.Invoke(Sub() ListView3.Items.Clear()) ' 在 UI 線程上清空 listview3 中的內容
        ' 根據 T 的型態進行不同的處理
        Dim inputItems As Object = Nothing
        If GetType(T) Is GetType(Outlook.Items) Then ' 處理 Outlook.Items 的情況
            inputItems = TryCast(data, Outlook.Items)
        ElseIf GetType(T) Is GetType(List(Of Outlook.MailItem)) Then ' 處理 List(Of Outlook.MailItem) 的情況
            inputItems = TryCast(data, List(Of Outlook.MailItem))
        End If
        '' todo: 使用SetColum來限定欄位提高效能
        'Dim strKeep As String = lblStatus1.Text
        'Dim countSum As Integer
        'If inputItems.Count > 10 Then ListView3.BeginUpdate() '避免更新過久
        '' 定義快取字典
        'Dim cache As New Dictionary(Of String, Tuple(Of String, Integer, Date, String, Integer, String))()
        'sw5.Reset()
        'Await Task.Run(Sub()
        '                   For Each item As MailItem In inputItems
        '                       sw5.Start()
        '                       ' todo: 建立mailitemCache 1st priority (因為用到二次)
        '                       Try
        '                           Dim entryKey As String = $"{item.EntryID}_Cache"
        '                           Dim cachedItem As Tuple(Of String, Integer, Date, String, Integer, String) = Nothing
        '                           If Not cache.TryGetValue(entryKey, cachedItem) Then
        '                               ' 如果快取中沒有該郵件的資料，則加入快取
        '                               cachedItem = Tuple.Create(item.Subject, item.Size, item.ReceivedTime, item.Sender.Name, item.Attachments.Count, item.EntryID)
        '                               cache.Add(entryKey, cachedItem)
        '                           End If
        '                           ' 將快取中的資料加入到 ListView 中
        '                           Dim listItem As New ListViewItem({cachedItem.Item1, cachedItem.Item2.ToString("###,###,##0"), cachedItem.Item3.ToShortDateString, cachedItem.Item4, cachedItem.Item5, cachedItem.Item6})
        '                           ListView3.Invoke(Sub() ListView3.Items.Add(listItem))
        '                       Catch
        '                       End Try : countSum += 1
        '                       Invoke(Sub() lblStatus1.Text = strKeep & " (" & countSum & " / " & inputItems.Count & ")") ' 需要改善: 更新讀取進度的計算方式(分母用GetTotalMailCount?)
        '                       sw5.Stop() : lblStatus3.Text = sw5.Elapsed.TotalSeconds.ToString("0.00, ")
        '                   Next item
        '               End Sub)
        'If inputItems.Count = 0 Then ListView3.Items.Add("符合條件的郵件項目為 0")
        'ListView3.EndUpdate()

    End Function
    'Private ReadOnly mailItemCache As New Dictionary(Of String, Tuple(Of String, Integer, Date, String, Integer, String))()
    Private ReadOnly mailItemCache As New Dictionary(Of String, CachedMailItemInfo)
    Public Property TotalFolderCount As Integer
    Private Class CachedMailItemInfo
        '    Public Property Subject As String
        '    Public Property Size As Long
        '    Public Property ReceivedTime As Date
        '    Public Property Sender As String
        '    Public Property AttachmentCount As Integer
        '    Public Property EntryID As String
        '    Public Sub New(subject As String, size As Long, receivedTime As Date, sender As String, attachmentCount As Integer, entryID As String)
        '        Me.Subject = subject
        '        Me.Size = size
        '        Me.ReceivedTime = receivedTime
        '        Me.Sender = sender
        '        Me.AttachmentCount = attachmentCount
        '        Me.EntryID = entryID
        '    End Sub
    End Class

End Module
Module TestCodeFromChatGPT
    Private ReadOnly yearCountsCache As New Dictionary(Of String, Dictionary(Of Integer, Integer))
    Private Async Function GetYearCountsAsync_GPT(selectedFolder As Outlook.Folder, includeSubFolders As Boolean) As Task(Of Dictionary(Of Integer, Integer))
        Dim cacheKey As String = $"{selectedFolder.FolderPath}_{includeSubFolders}"
        If yearCountsCache.ContainsKey(cacheKey) Then Return yearCountsCache(cacheKey)
        Dim yearCounts As Dictionary(Of Integer, Integer) = Await CountMailByYearAsync_GPT(selectedFolder, includeSubFolders)
        yearCountsCache(cacheKey) = yearCounts
        Return yearCounts

    End Function
    Private Async Function CountMailByYearAsync_GPT(folder As Outlook.Folder, includeSubFolders As Boolean) As Task(Of Dictionary(Of Integer, Integer))
        Dim yearCounts As New Dictionary(Of Integer, Integer)()
        If folder.Items.Count = 0 OrElse Not folder.DefaultItemType = Outlook.OlItemType.olMailItem Then Return yearCounts
        'If includeSubFolders Then
        '    For Each childFolder As Outlook.Folder In folder.Folders
        '        Dim childYearCounts As Dictionary(Of Integer, Integer) = Await CountMailByYearAsync_GPT(childFolder, includeSubFolders)
        '        yearCounts = MergeDictionaries(yearCounts, childYearCounts)
        '    Next
        'End If
        '' =========================================================
        '' 使用 Restrict方法快速統計郵件項目的日期和數量, 使用非同步方法
        '' =========================================================
        'Dim intCount As Integer = 0
        'Await Task.Run(
        '    Sub()
        '        For year As Integer = Get1stYear(folder) To Date.Today.Year
        '            Try
        '                Dim restrictFilter As String = GetFilterString(year)
        '                Dim restrictedItems As Outlook.Items = folder.Items.Restrict(restrictFilter)
        '                Dim intCountOfTheYear As Integer = restrictedItems.Count ' 在這裡使用非同步方式獲取 restrictedItems
        '                If intCountOfTheYear > 0 Then
        '                    yearCounts.Add(year, intCountOfTheYear) ' 統計郵件數量並存入字典
        '                    intCount += intCountOfTheYear
        '                    invoke(Sub() updateCounter(intCount))
        '                End If
        '            Catch ex As System.Exception
        '                DebugForm.AddMessage3(ex.Message & ": " & ex.Source & ": " & ex.Data.ToString)
        '            End Try
        '        Next
        '    End Sub)
        Return yearCounts

    End Function
    Private Function GetFilterString(year As Integer) As String
        Throw New NotImplementedException()
    End Function
    Private Sub updateCounter(intCount As Integer)
        Throw New NotImplementedException()
    End Sub
    ''''===== DONE ======
    ''''Private Async Sub TreeView2_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TreeView2.AfterSelect
    ''''    DebugForm.AddMessage(sender.Name)
    ''''    Dim stopwatch As New Stopwatch()
    ''''    stopwatch.Start()
    ''''    ' 清空之前的統計結果
    ''''    ListView2.Items.Clear()
    ''''    lblStatus1.Text = ""
    ''''    lblStatus2.Text = ""
    ''''    ' 取得選擇的資料夾, 如果資料夾為空，直接返回
    ''''    Dim selectedFolder As Outlook.Folder = TryCast(e.Node.Tag, Outlook.Folder)
    ''''    If selectedFolder Is Nothing Then Return
    ''''    ' 在背景執行計算部分
    ''''    Dim yearCounts As Dictionary(Of Integer, Integer) = Await CalculateYearlyStatisticsAsync(selectedFolder, CheckSub2.Checked)
    ''''    ' 顯示統計結果
    ''''    ShowResultToListView2(yearCounts, ListView2)
    ''''    ShowResultToChart2(yearCounts, Chart2)
    ''''    stopwatch.Stop()
    ''''    lblStatus2.Text = $"更新年度統計花費了 {stopwatch.Elapsed.TotalSeconds:0.00} 秒。({selectedFolder.Items.Count / stopwatch.Elapsed.TotalSeconds:###,##0}/sec)"
    ''''    sender.Enabled = True
    ''''    sender.Focus()
    ''''End Sub
    Private Sub invoke(subsub As Object)

    End Sub
    ''''===== DONE ======
    ''''Private Sub ShowResultToListView2(yearCounts As Dictionary(Of Integer, Integer), listview As ListView)
    ''''    DebugForm.AddMessage()
    ''''    If yearCounts Is Nothing OrElse yearCounts.Count = 0 Then
    ''''        listview.Items.Add(New ListViewItem("Nothing Found in Selected Folder"))
    ''''        Return
    ''''    End If
    ''''    Dim sortedYearCounts = yearCounts.OrderBy(Function(pair) pair.Key).ToList()
    ''''    For Each pair As KeyValuePair(Of Integer, Integer) In sortedYearCounts
    ''''        listview.Invoke(Sub() listview.Items.Add(New ListViewItem({pair.Key.ToString(), pair.Value.ToString("###,###,##0")})))
    ''''    Next
    ''''End Sub
    ''''Private Sub ShowResultToChart2(yearCounts As Dictionary(Of Integer, Integer), chart As Chart)
    ''''    DebugForm.AddMessage()
    ''''    chart.Series(0).Points.Clear()
    ''''    chart.ChartAreas(0).AxisY.StripLines.Clear()
    ''''    If yearCounts Is Nothing OrElse yearCounts.Count = 0 Then Return
    ''''    Dim sortedYearCounts = yearCounts.OrderBy(Function(pair) pair.Key).ToList()
    ''''    Dim series As Series = chart.Series(0)
    ''''    For Each pair In sortedYearCounts
    ''''        series.Points.AddXY(pair.Key, pair.Value)
    ''''    Next
    ''''    chart.ChartAreas(0).AxisX.Minimum = sortedYearCounts.Min(Function(pair) pair.Key) - 1
    ''''    chart.ChartAreas(0).AxisX.Maximum = sortedYearCounts.Max(Function(pair) pair.Key) + 1
    ''''    Dim average As Double = yearCounts.Average(Function(pair) pair.Value)
    ''''    Dim stripLine As New StripLine With {
    ''''                                        .Interval = 0,
    ''''                                        .IntervalOffset = average,
    ''''                                        .StripWidth = 0.05,
    ''''                                        .BackColor = Color.Red
    ''''                                         }
    ''''    chart.ChartAreas(0).AxisY.StripLines.Add(stripLine)
    ''''    chart.Invalidate()
    ''''End Sub
End Module
Module TestCodeFromClaude
    Private ReadOnly yearCountsCache As New Dictionary(Of String, Dictionary(Of Integer, Integer))
    Private Async Function CountByYearsAsync_CL(selectedFolder As Outlook.Folder, includeSubFolders As Boolean) As Task(Of Dictionary(Of Integer, Integer))
        'DebugForm.AddMessage("Begin:", selectedFolder.Name)
        '' 建立一個唯一的快取鍵值,包含資料夾的 FolderPath 和是否包含子資料夾的選項
        'Dim cacheKey As String = selectedFolder.FolderPath & "_" & includeSubFolders.ToString()
        'If yearCountsCache.ContainsKey(cacheKey) Then Return yearCountsCache(cacheKey) ' 檢查快取
        '' 如果快取中沒有結果, 才開始進行計算
        'Dim yearCounts As New Dictionary(Of Integer, Integer)   ' 建立一個字典來存儲每個年份的郵件數量
        'If includeSubFolders Then   ' 遞迴包含所有子資料夾
        '    yearCounts = Await CountMailInFolder_CL(selectedFolder, includeSubFolders)
        '    For Each childFolder As Outlook.Folder In selectedFolder.Folders
        '        Dim childYearCounts As Dictionary(Of Integer, Integer) = Await CountByYearsAsync_CL(childFolder, includeSubFolders)
        '        yearCounts = MergeDictionaries(yearCounts, childYearCounts)
        '    Next
        'Else                        ' 只計算目前所點選資料夾
        '    yearCounts = Await CountMailInFolder_CL(selectedFolder, includeSubFolders)
        'End If
        'yearCountsCache.Add(cacheKey, yearCounts)
        'Return yearCounts

    End Function
    Private Async Function CountByYearsAsync_CL_New(selectedFolder As Outlook.Folder, includeSubFolders As Boolean) As Task(Of Dictionary(Of Integer, Integer))
        'DebugForm.AddMessage("Begin:", selectedFolder.Name)
        '' 建立一個唯一的快取鍵值,包含資料夾的 FolderPath 和是否包含子資料夾的選項
        'Dim cacheKey As String = selectedFolder.FolderPath & "_" & includeSubFolders.ToString()
        'If yearCountsCache.ContainsKey(cacheKey) Then Return yearCountsCache(cacheKey) ' 檢查快取
        '' 如果快取中沒有結果, 才開始進行計算
        'Dim yearCounts As New Dictionary(Of Integer, Integer)   ' 建立一個字典來存儲每個年份的郵件數量
        'yearCounts = Await CountMailInFolder_CL(selectedFolder, includeSubFolders)
        'If includeSubFolders Then   ' 遞迴包含所有子資料夾
        '    For Each childFolder As Outlook.Folder In selectedFolder.Folders
        '        Dim childYearCounts As Dictionary(Of Integer, Integer) = Await CountByYearsAsync_CL_New(childFolder, includeSubFolders)
        '        yearCounts = MergeDictionaries(yearCounts, childYearCounts)
        '    Next
        'End If
        'yearCountsCache.Add(cacheKey, yearCounts)
        'Return yearCounts

    End Function
    Private Async Function CountMailInFolder_CL(folder As Outlook.Folder, includeSubFolders As Boolean) As Task(Of Dictionary(Of Integer, Integer))
        Dim yearCounts As New Dictionary(Of Integer, Integer)()
        If Not includeSubFolders AndAlso folder.Items.Count = 0 Then Return yearCounts
        'Dim intCount As Integer = 0
        'Await Task.Run(
        '    Sub()
        '        For year As Integer = Get1stYear(folder) To Date.Today.Year
        '            Try
        '                Dim restrictFilter As String = GetFilterString(year)
        '                Dim restrictedItems As Outlook.Items = folder.Items.Restrict(restrictFilter)
        '                Dim intCountOfTheYear As Integer = restrictedItems.Count ' 在這裡使用非同步方式獲取 restrictedItems
        '                If intCountOfTheYear > 0 Then
        '                    yearCounts.Add(year, intCountOfTheYear) ' 統計郵件數量並存入字典
        '                    intCount += If(includeSubFolders, folder.Items.Count, intCountOfTheYear)
        '                    invoke(Sub() updateCounter(intCount))
        '                End If
        '            Catch ex As System.Exception
        '                DebugForm.AddMessage3(ex.Message & ": " & ex.Source & ": " & ex.Data.ToString)
        '            End Try
        '        Next
        '    End Sub)
        Return yearCounts

    End Function
    ''Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
    ''===================================
    ''測試 DASL 搜尋附件關鍵字的功能,
    ''但一直卡在 AdvancedSearch 的部分,
    ''不確定是 filter 語法錯誤還是其他問題
    ''2026/3/7, by Claude
    ''===================================
    ''    DebugForm.AddMessage2("Begin:")
    ''    Dim selectedFolder As Outlook.Folder = TryCast(TreeView3.SelectedNode.Tag, Outlook.Folder)
    ''    If selectedFolder Is Nothing Then MessageBox.Show("請先選擇資料夾") : Return
    ''    Dim keyword As String = TextBox3.Text.Trim
    ''    If keyword.Length = 0 Then MessageBox.Show("請輸入關鍵字") : Return
    ''    ListView3.Items.Clear()
    ''    lblStatus1.Text = "搜尋中..." : lblStatus2.Text = ""
    ''    Cursor = Cursors.WaitCursor
    ''    Dim sw As New Stopwatch : sw.Start()
    ''    ' ✅ 用 DASL 語法搜尋附件檔名，走 Outlook 索引引擎
    ''    '' ❌ 改寫前（用雙引號）：
    ''    'Dim scope As String = Chr(34) & selectedFolder.FolderPath & Chr(34)
    ''    ' ✅ 改寫後（用單引號，反斜線跳脫）：
    ''    'Dim scope As String = "'" & selectedFolder.FolderPath.Replace("\", "\\") & "'"
    ''    Dim scope As String = "'" & selectedFolder.FolderPath & "'"
    ''    '''' ❌ 改寫前（缺少 @SQL=）：
    ''    '''Dim filter As String =
    ''    '''"urn:schemas:httpmail:hasattachment = True" &
    ''    '''" AND urn:schemas:httpmail:attachmentfilename like '%" & keyword & "%'"
    ''    ''' ✅ 改寫後（加上 @SQL= 和引號包住屬性名）：
    ''    ''Dim filter As String =
    ''    ''                    "@SQL=" &
    ''    ''                    Chr(34) & "urn:schemas:httpmail:hasattachment" & Chr(34) & " = True" &
    ''    ''                    " AND " &
    ''    ''                    Chr(34) & "urn:schemas:httpmail:attachmentfilename" & Chr(34) &
    ''    ''                    " like '%" & keyword & "%'"
    ''    ''' AdvancedSearch 是非同步的，用 WithEvents 等結果回來
    ''    ''_advSearchKeyword = keyword
    ''    ''_advSearchSW = sw
    ''    ''' ✅ 事件掛在 AppOutlook 上，不是 _advSearch 上
    ''    ''AddHandler AppOutlook.AdvancedSearchComplete, AddressOf AdvancedSearch_Complete
    ''    ''DebugForm.AddMessage("scope:", scope)
    ''    ''DebugForm.AddMessage("filter:", filter)
    ''    ''_advSearch = AppOutlook.AdvancedSearch(scope, filter, True, "AttachSearch")
    ''    DebugForm.AddMessage2(selectedFolder.Store.ExchangeStoreType.ToString)
    ''    ' 測試1：最簡單的 filter，只篩有附件，確認 AdvancedSearch 本身能不能跑
    ''    Dim filter1 As String = "@SQL=" & Chr(34) & "urn:schemas:httpmail:hasattachment" & Chr(34) & " = True"
    ''    ' 測試2：加上附件檔名條件
    ''    Dim filter2 As String = "@SQL=" &
    ''        Chr(34) & "urn:schemas:httpmail:hasattachment" & Chr(34) & " = True" &
    ''        " AND " &
    ''        Chr(34) & "urn:schemas:httpmail:attachmentfilename" & Chr(34) & " like '%" & keyword & "%'"
    ''    ' 測試3：改用 PR_ATTACH_LONG_FILENAME
    ''    Dim filter3 As String = "@SQL=" &
    ''        Chr(34) & "urn:schemas:httpmail:hasattachment" & Chr(34) & " = True" &
    ''        " AND " &
    ''        Chr(34) & "http://schemas.microsoft.com/mapi/proptag/0x3707001F" & Chr(34) & " like '%" & keyword & "%'"
    ''    For i As Integer = 1 To 3
    ''        Dim f As String = If(i = 1, filter1, If(i = 2, filter2, filter3))
    ''        Try
    ''            AddHandler AppOutlook.AdvancedSearchComplete, AddressOf AdvancedSearch_Complete
    ''            _advSearch = AppOutlook.AdvancedSearch(Scope, f, True, "AttachSearch")
    ''            DebugForm.AddMessage($"測試{i} 成功送出:", f)
    ''            Exit For  ' 第一個成功的就停下來
    ''        Catch ex As System.Exception
    ''            RemoveHandler AppOutlook.AdvancedSearchComplete, AddressOf AdvancedSearch_Complete
    ''            DebugForm.AddMessage($"測試{i} 失敗 HResult={ex.HResult:X}:", ex.Message)
    ''        End Try
    ''    Next
    ''End Sub
    ''' AdvancedSearch 完成時的回呼
    ''Private _advSearch As Outlook.Search
    ''Private _advSearchKeyword As String
    ''Private _advSearchSW As Stopwatch
    ''Private Sub AdvancedSearch_Complete(ByVal SearchObject As Outlook.Search)
    ''    MessageBox.Show("AdvancedSearch_Complete 被呼叫了！結果：" & SearchObject.Results.Count & " 封")  ' ← 暫時加這行確認
    ''    If SearchObject.Tag <> "AttachSearch" Then Return
    ''    RemoveHandler AppOutlook.AdvancedSearchComplete, AddressOf AdvancedSearch_Complete  ' ✅ 最先執行
    ''    _advSearchSW.Stop()
    ''    DebugForm.AddMessage("AdvancedSearchComplete:", SearchObject.Results.Count & " 封")
    ''    Dim results As Outlook.Results = SearchObject.Results
    ''    Dim count As Integer = results.Count
    ''    Dim itemList As New List(Of ListViewItem)
    ''    For i As Integer = 1 To count
    ''        Try
    ''            Dim mail As MailItem = DirectCast(results.Item(i), MailItem)
    ''            itemList.Add(New ListViewItem({
    ''            mail.Subject,
    ''            mail.Size.ToString("###,###,##0"),
    ''            mail.ReceivedTime.ToShortDateString,
    ''            mail.SenderName,
    ''            mail.Attachments.Count.ToString,
    ''            mail.EntryID}))
    ''        Catch
    ''        End Try
    ''    Next
    ''    Me.Invoke(Sub()
    ''                  ListView3.BeginUpdate()
    ''                  ListView3.Items.AddRange(itemList.ToArray)
    ''                  ListView3.EndUpdate()
    ''                  lblStatus1.Text = count & " 封"
    ''                  lblStatus2.Text = $"AdvancedSearch 耗時 {_advSearchSW.Elapsed.TotalSeconds:0.00} 秒"
    ''                  Cursor = Cursors.Default
    ''              End Sub)
    ''End Sub
    Private Function GetFilterString(year As Integer) As String
        Throw New NotImplementedException()
    End Function
    Private Sub updateCounter(intCount As Integer)
        Throw New NotImplementedException()
    End Sub
    ''''===== DONE ======
    ''''Private Sub ShowResult(Of T)(yearCounts As Dictionary(Of Integer, Integer), control As T)
    ''''    'If yearCounts Is Nothing OrElse yearCounts.Count = 0 Then Invoke(Sub() control.Clear()) : Return
    ''''    'Dim sortedYearCounts = yearCounts.OrderBy(Function(pair) pair.Key).ToList
    ''''    'If sortedYearCounts.Count = 0 Then Return
    ''''    'If GetType(T) Is GetType(ListView) Then
    ''''    '    Dim listView As ListView = DirectCast(control, ListView)
    ''''    '    listView.Invoke(
    ''''    '    Sub()
    ''''    '        listView.Items.Clear()
    ''''    '        For Each pair As KeyValuePair(Of Integer, Integer) In sortedYearCounts
    ''''    '            listView.Items.Add(New ListViewItem({pair.Key.ToString, pair.Value.ToString("###,###,##0")}))
    ''''    '        Next
    ''''    '    End Sub)
    ''''    'ElseIf GetType(T) Is GetType(Chart) Then
    ''''    '    Dim chart As Chart = DirectCast(control, Chart)
    ''''    '    chart.Invoke(
    ''''    '    Sub()
    ''''    '        chart.Series(0).Points.Clear()
    ''''    '        chart.ChartAreas(0).AxisY.StripLines.Clear()
    ''''    '        Dim series As Series = chart.Series(0)
    ''''    '        For Each pair In sortedYearCounts
    ''''    '            series.Points.AddXY(pair.Key, pair.Value)
    ''''    '        Next
    ''''    '        If sortedYearCounts.Any() Then
    ''''    '            chart.ChartAreas(0).AxisX.Minimum = sortedYearCounts.Min(Function(pair) pair.Key) - 1
    ''''    '            chart.ChartAreas(0).AxisX.Maximum = sortedYearCounts.Max(Function(pair) pair.Key) + 1
    ''''    '        End If
    ''''    '        Dim average As Double = yearCounts.Average(Function(pair) pair.Value)
    ''''    '        Dim stripLine As New StripLine With {
    ''''    '            .Interval = 0,
    ''''    '            .IntervalOffset = average,
    ''''    '            .StripWidth = CHART_STRIP_LINE_WIDTH,
    ''''    '            .BackColor = CHART_STRIP_LINE_COLOR
    ''''    '        }
    ''''    '        chart.ChartAreas(0).AxisY.StripLines.Add(stripLine)
    ''''    '        chart.Invalidate()
    ''''    '    End Sub)
    ''''    'End If
    ''''End Sub
End Module