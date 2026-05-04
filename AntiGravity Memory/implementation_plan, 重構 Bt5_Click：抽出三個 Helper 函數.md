# 重構 Bt5_Click：抽出三個 Helper 函數

## 目標

從 `Bt5_Click` 中辨識並抽出三個**職責明確的 helper**，讓主流程只剩核心骨架。不是按照執行順序切塊，而是以「功能職責」為單位抽離。

---

## 抽出了什麼

| Helper 名稱 | 原本在哪裡 | 職責 |
|---|---|---|
| `BuildDuplicateHashKey()` | L3187~L3196（`If isExact Then...End If`） | 根據模式計算分組 key |
| `ComputeGroupSimScores()` | L3221~L3236（`If Not isExact Then...End If`） | 計算群組內每封信的相似度分數 |
| `CreateDuplicateLvi()` | L3243~L3249（`New ListViewItem({...})`） | 建立一個 ListView 項目 |

---

## Before（原始 Bt5_Click，節錄關鍵段落）

### 段落一：hashKey 計算（摻在 for 迴圈裡）
```vb
Dim hashKey As String
If isExact Then
    hashKey = $"{subject}|{size}|{recvTime:yyyyMMddHHmmss}|{senderName}"
Else
    Dim cleanSubj As String = GetCleanSubject(subject).Replace(" ", "").ToUpper()
    If cleanSubj.Length > 20 Then cleanSubj = cleanSubj.Substring(0, 20)
    hashKey = $"{cleanSubj}|{size}"
End If
If Not exactDict.ContainsKey(hashKey) Then exactDict(hashKey) = New List(Of MailItemInfo)()
exactDict(hashKey).Add(info)
```

### 段落二：相似度計算（摻在 for 迴圈裡）
```vb
Dim isValidGroup As Boolean = True
Dim simScores As New List(Of Double)()
If Not isExact Then
    Dim firstSubject As String = kvp.Value(0).Subject
    simScores.Add(1.0)
    For i As Integer = 1 To kvp.Value.Count - 1
        Dim sim As Double = JaccardSimilarity(firstSubject, kvp.Value(i).Subject)
        simScores.Add(sim)
        If sim < 0.6 Then isValidGroup = False : Exit For
    Next
Else
    For i As Integer = 0 To kvp.Value.Count - 1 : simScores.Add(1.0) : Next
End If
```

### 段落三：ListViewItem 建立（一個很長的 New）
```vb
Dim lvi As New ListViewItem({mailItem.Subject,
                             (mailItem.Size \ 1024L).ToString("N0") & "KB",
                              mailItem.ReceivedTime.ToString("yyyy/MM/dd"),
                              mailItem.SenderName,
                              "群組 " & groupID.ToString(),
                              simText,
                              mailItem.EntryID}) With {.BackColor = groupColor, .Tag = mailItem}
ListView5.Items.Add(lvi)
```

---

## After（重構後的樣子）

### Helper 1：`BuildDuplicateHashKey`（放進 `└ 輔助函數` Region）
```vb
Private Function BuildDuplicateHashKey(subject As String, size As Long,
        recvTime As DateTime, senderName As String, isExact As Boolean) As String
    ' 2026/05/04 by Gemini 3.1 Pro: 從 Bt5_Click 抽出，集中管理分組 Key 的計算邏輯
    ' Exact 模式：四欄精確比對（未來可在此加入時間/大小容差）
    ' Fuzzy 模式：清理前綴 + 前20字 + 大小，後由 Jaccard 二次過濾
    If isExact Then
        Return $"{subject}|{size}|{recvTime:yyyyMMddHHmmss}|{senderName}"
    Else
        Dim cleanSubj As String = GetCleanSubject(subject).Replace(" ", "").ToUpper()
        If cleanSubj.Length > 20 Then cleanSubj = cleanSubj.Substring(0, 20)
        Return $"{cleanSubj}|{size}"
    End If
End Function
```

### Helper 2：`ComputeGroupSimScores`（放進 `└ 輔助函數` Region）
```vb
Private Function ComputeGroupSimScores(mailGroup As List(Of MailItemInfo),
        isExact As Boolean) As (isValid As Boolean, scores As List(Of Double))
    ' 2026/05/04 by Gemini 3.1 Pro: 從 Bt5_Click 抽出，計算一個候選群組的相似度分數
    ' Exact 模式：全部給 100%，不做比對
    ' Fuzzy 模式：Jaccard 字元集比對，門檻 0.6；任一封低於門檻則整組無效
    Dim scores As New List(Of Double)()
    If isExact Then
        For i = 0 To mailGroup.Count - 1 : scores.Add(1.0) : Next
        Return (True, scores)
    End If

    Dim firstSubject As String = mailGroup(0).Subject
    scores.Add(1.0) ' 第一封與自己 = 100%
    For i As Integer = 1 To mailGroup.Count - 1
        Dim sim As Double = JaccardSimilarity(firstSubject, mailGroup(i).Subject)
        scores.Add(sim)
        If sim < 0.6 Then Return (False, scores) ' 不符門檻，整組廢棄
    Next
    Return (True, scores)
End Function
```

### Helper 3：`CreateDuplicateLvi`（放進 `└ 輔助函數` Region）
```vb
Private Function CreateDuplicateLvi(mailItem As MailItemInfo,
        groupID As Integer, simText As String, groupColor As Color) As ListViewItem
    ' 2026/05/04 by Gemini 3.1 Pro: 從 Bt5_Click 抽出，建立 ListView5 的一列資料
    Return New ListViewItem({mailItem.Subject,
                             (mailItem.Size \ 1024L).ToString("N0") & "KB",
                              mailItem.ReceivedTime.ToString("yyyy/MM/dd"),
                              mailItem.SenderName,
                              "群組 " & groupID.ToString(),
                              simText,
                              mailItem.EntryID}) With {.BackColor = groupColor, .Tag = mailItem}
End Function
```

### 重構後的 `Bt5_Click`（主體只剩骨架）
```vb
Private Async Sub Bt5_Click(sender As Object, e As EventArgs) Handles Button5.Click
    _dbg("開始")
    Dim cToken As CancellationToken = OkayNowYouHaveToken()
    Dim selectedNodes As List(Of TreeNode) = SimTree5.SelectedNodes
    If selectedNodes Is Nothing OrElse selectedNodes.Count = 0 Then
        MessageBox.Show("請先在左側選取要掃描的資料夾或 PST。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Return
    End If

    Button5.Enabled = False : Cursor = Cursors.WaitCursor
    ListView5.BeginUpdate() : ListView5.Items.Clear() : ListView5.EndUpdate()
    ProgressBar1.Text = "正在準備" : ProgressBar2.Text = "展開資料夾結構..."
    Dim sw As New Stopwatch() : sw.Start()
    Dim swThrottle As New Stopwatch() : swThrottle.Start()
    Dim progress5 As IProgress(Of ProgressReport) = New Progress(Of ProgressReport)(Sub(p) ProgressBar2.Text = p.Message)
    Dim exactDict As New Dictionary(Of String, List(Of MailItemInfo))(StringComparer.OrdinalIgnoreCase)
    Dim isExact As Boolean = rbExactMatch.Checked

    Try
        ' ── Step 1: 展開選取的資料夾（含子資料夾），去重 ──
        Dim folderList = Await GetUniqueFolderList(selectedNodes, includeSub:=True, cToken:=cToken, progress:=progress5)
        If folderList.Count = 0 Then Return

        ' ── Step 2: 逐資料夾掃 GetTable，建立分組字典 ──
        Dim totalFolders As Integer = folderList.Count
        For i As Integer = 0 To folderList.Count - 1
            Dim folder As Outlook.Folder = folderList(i).Folder
            Dim fPath As String = folderList(i).fPath
            Dim table As Outlook.Table = Nothing
            Try
                Const PR_MESSAGE_SIZE As String = "http://schemas.microsoft.com/mapi/proptag/0x0E080003"
                table = folder.GetTable()
                table.Columns.RemoveAll()
                table.Columns.Add("EntryID") : table.Columns.Add("Subject")
                table.Columns.Add(PR_MESSAGE_SIZE) : table.Columns.Add("ReceivedTime") : table.Columns.Add("SenderName")

                Do While Not table.EndOfTable
                    Dim arr As Object = table.GetArray(BATCH_SIZE)
                    If arr Is Nothing Then Exit Do
                    Dim data(,) As Object = DirectCast(arr, Object(,))
                    For r As Integer = 0 To data.GetUpperBound(0)
                        Dim entryID As String = SafeGet(Of String)(data, r, 0, "")
                        If entryID = "" Then Continue For
                        Dim subject As String = SafeGet(Of String)(data, r, 1, "")
                        Dim size As Long = SafeGet(Of Long)(data, r, 2, 0L)
                        Dim recvTime As DateTime = SafeGet(Of DateTime)(data, r, 3, DateTime.MinValue)
                        Dim senderName As String = SafeGet(Of String)(data, r, 4, "")
                        Dim info As New MailItemInfo With {.EntryID = entryID, .Subject = subject,
                                                          .Size = size, .ReceivedTime = recvTime,
                                                          .SenderName = senderName, .FolderPath = fPath}
                        ' ✅ 改用 helper，主流程不再關心 key 的計算細節
                        Dim key As String = BuildDuplicateHashKey(subject, size, recvTime, senderName, isExact)
                        If Not exactDict.ContainsKey(key) Then exactDict(key) = New List(Of MailItemInfo)()
                        exactDict(key).Add(info)
                    Next
                    Await Task.Yield()
                Loop
            Catch ex As System.Exception
                _dbg("錯誤", $"{folder.Name}: {ex.Message}")
            Finally
                TryMarshalRelease(table)
            End Try
            Await SmartThrottle(swThrottle, cToken:=cToken, ThrottleFreq.Hii,
                                 Sub() progress5.Report(New ProgressReport With {.Message = $"掃描中: {i + 1}/{totalFolders} 個資料夾..."}))
        Next

        ' ── Step 3: 過濾並渲染到 ListView5 ──
        ListView5.BeginUpdate()
        Dim groupID As Integer = 1
        Dim totalDuplicateMails As Integer = 0
        Dim swThrottleBuild As New Stopwatch() : swThrottleBuild.Start()
        For Each kvp In exactDict
            If kvp.Value.Count > 1 Then
                ' ✅ 改用 helper，主流程不再關心相似度計算細節
                Dim result = ComputeGroupSimScores(kvp.Value, isExact)
                If result.isValid Then
                    Dim groupColor As Color = If(groupID Mod 2 = 0, Color.FromArgb(240, 248, 255), Color.White)
                    For idx As Integer = 0 To kvp.Value.Count - 1
                        Dim simText As String = If(idx < result.scores.Count, $"{CInt(result.scores(idx) * 100)}%", "-")
                        ' ✅ 改用 helper，主流程不再關心欄位順序與格式細節
                        ListView5.Items.Add(CreateDuplicateLvi(kvp.Value(idx), groupID, simText, groupColor))
                        totalDuplicateMails += 1
                    Next
                    groupID += 1
                    Await SmartThrottle(swThrottleBuild, cToken:=cToken, ThrottleFreq.Hii,
                                         Sub() progress5.Report(New ProgressReport With {.Message = $"正在建立重複郵件清單: {groupID} 組..."}))
                End If
            End If
        Next
        ListView5.EndUpdate()
        sw.Stop()
        ProgressBar1.Text = $"找到 {groupID - 1} 組 ({totalDuplicateMails} 封) / 耗時 {sw.Elapsed.TotalSeconds:0.00} 秒"
        ProgressBar2.Text = ""
    Catch ex As OperationCanceledException
        _dbg("結束", "ESC 中斷") : ProgressBar1.Text = "已中斷。" : ProgressBar2.Text = ""
    Catch ex As System.Exception
        MessageBox.Show("掃描重複郵件時發生錯誤: " & ex.Message, "錯誤")
        _dbg("錯誤", ex.Message)
    Finally
        Button5.Enabled = True : Cursor = Cursors.Default
        _dbg("結束")
    End Try
End Sub
```

---

## 變化對比

| 項目 | Before | After |
|---|---|---|
| `Bt5_Click` 行數 | ~155 行 | ~75 行（縮短 ~50%） |
| hashKey 計算細節 | 在主迴圈內 | 移至 `BuildDuplicateHashKey` |
| Jaccard 比對邏輯 | 在主迴圈內 | 移至 `ComputeGroupSimScores` |
| LVI 欄位組裝 | 在主迴圈內 | 移至 `CreateDuplicateLvi` |
| 未來加 Exact 容差 | 需修改主流程 | 只改 `BuildDuplicateHashKey` |
| 未來換演算法 | 需修改主流程 | 只改 `ComputeGroupSimScores` |

## Verification Plan
- 確認編譯無錯誤。
- 執行 Tab5 的 Exact 模式與 Fuzzy 模式掃描，確認結果與重構前相同。
