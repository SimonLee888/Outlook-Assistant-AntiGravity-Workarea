Imports System.Collections.Concurrent
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Outlook
Module Module_ToBeDelete

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
        '    Dim restrictFilter As String = $"[RcvTime] >= '{startDate}' AND [RcvTime] <= '{endDate}'"
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
        '                 Dim filter As String = $"[RcvTime] >= '{startDate}' AND [RcvTime] <= '{endDate}'"
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
        '            If intCountOfTheYear > 0 Then yearCounts.TryAdd(year, intCountOfTheYear)     ' 統計郵件數量並存入字典
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

End Module

