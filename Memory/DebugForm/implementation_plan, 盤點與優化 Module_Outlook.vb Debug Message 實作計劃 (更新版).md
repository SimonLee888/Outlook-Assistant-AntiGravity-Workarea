# 盤點與優化 Module_Outlook.vb Debug Message 實作計劃 (更新版)

本計劃旨在根據最新討論原則，統一 `Module_Outlook.vb` 檔案中 **Layer2.5**、**Layer3 RDO** 以及 **Layer3 OOM** 這三個 Region 中的 Debug 訊息。

## 最新調整原則
1. **開始訊息**：
   - L2.5 進入端無條件顯示：`_dbg(" ├ 開始", fPath)`
   - L3 RDO / OOM 進入端在 `If _iLikeNoisy Then` 條件下顯示：`If _iLikeNoisy Then _dbg(" ├ 開始", fPath)`
2. **結束訊息（只加一行/只在出口加）**：
   - 提早 Return、skipCache 提早 Return、單行 `If...Then Return` 提早退出處**皆不加**結束訊息。
   - 僅在函數最後一個正常結尾的 Return 前，或是有 Finally 區塊的結尾處，加上**唯一的一行**結束訊息。
     - L2.5 結束端無條件顯示：`_dbg(" ├ 結束", fPath & " | 成果: " & [成果值])`
     - L3 RDO / OOM 結束端在 `If _iLikeNoisy` 條件下顯示：`If _iLikeNoisy Then _dbg(" ├ 結束", fPath & " | 成果: " & [成果值])`
3. **錯誤與例外訊息**：
   - L2.5 若有錯誤或例外，無條件顯示。
   - L3 RDO / OOM 過程中若有錯誤或例外，無條件顯示，**不須**加上 `_iLikeNoisy` 條件。
4. **註記規範**：
   - 修改處加註：`' by Gemini 3.5 Flash, 2026/07/01`

---

## 預計修改前後完整對照

以下列出所有涉及修改的函數對照。

### ■ Layer2.5 快取代理層 (部分範例，其餘依此原則套用)

#### 1. `GetMailCount` (folder) [L520]
```diff
     Private Function GetMailCount(folder As Folder, Optional fPath As String = "", Optional skipCache As Boolean = False) As Long
         fPath = SafeGetPath(folder, fPath)
+        _dbg(" ├ 開始", fPath) ' by Gemini 3.5 Flash, 2026/07/01
         Dim count As Long
         If Not skipCache Then
             If _cacheMailCount.TryGetValue(fPath, count) Then Return count       ' ① 記憶體命中 (提早退出，不加結束)
             Dim row = SafeGetDbRow(folder, fPath)                                ' ② DB lazy load
             If row IsNot Nothing AndAlso row.mc >= 0 Then Return row.mc          ' 提早退出，不加結束
         End If
 
         ' ③ 讀取派工: RDO 優先,失敗 fallback OOM
         count = GetMailCountRdo(fPath, folder.EntryID, folder.StoreID)
         If count < 0 Then count = GetMailCountOOM(folder, fPath:=fPath)
         If count >= 0 Then _cacheMailCount.TryAdd(fPath, count)
+        _dbg(" ├ 結束", fPath & " | 成果: " & count) ' by Gemini 3.5 Flash, 2026/07/01
         Return count
     End Function
```

#### 2. `GetMailCount` (fPath, eid, sid) [L548]
```diff
     Private Function GetMailCount(fPath As String, eid As String, sid As String, Optional skipCache As Boolean = False) As Long
+        _dbg(" ├ 開始", fPath) ' by Gemini 3.5 Flash, 2026/07/01
         Dim count As Long
         If Not skipCache Then
             If _cacheMailCount.TryGetValue(fPath, count) Then Return count      ' ① 記憶體快取命中 (提早退出)
             Dim dbRow = DbGetFolderStats(fPath)                                 ' ② DB lazy
             If dbRow IsNot Nothing AndAlso dbRow.mc >= 0 Then Return dbRow.mc
         End If
         count = GetMailCountRdo(fPath, eid, sid)                                ' ③ RDO 優先
         If count < 0 Then                                                       ' 底線: RDO 失敗 → OOM
             Dim f As Folder = GetFolderById(eid, sid)
             If f IsNot Nothing Then count = GetMailCountOOM(f, fPath:=fPath)
         End If
         If count >= 0 Then _cacheMailCount.TryAdd(fPath, count)
+        _dbg(" ├ 結束", fPath & " | 成果: " & count) ' by Gemini 3.5 Flash, 2026/07/01
         Return count
     End Function
```

#### 3. `GetFolderCount` (folder) [L566]
```diff
     Private Function GetFolderCount(folder As Folder, Optional fPath As String = "", Optional skipCache As Boolean = False) As Long
         fPath = SafeGetPath(folder, fPath)
+        _dbg(" ├ 開始", fPath) ' by Gemini 3.5 Flash, 2026/07/01
         Dim count As Long
         If Not skipCache Then
             If _cacheFolderCount.TryGetValue(fPath, count) Then Return count     ' ① 記憶體命中
             Dim row = SafeGetDbRow(folder, fPath)                                ' ② DB lazy load (fc 欄位)
             If row IsNot Nothing AndAlso row.fc >= 0 Then Return row.fc
         End If
 
         ' ③ 讀取派工: RDO 優先,失敗 fallback OOM
         count = GetFolderCountRdo(fPath, folder.EntryID, folder.StoreID)
         If count < 0 Then count = GetFolderCountOOM(folder, fPath:=fPath)
         If count >= 0 Then _cacheFolderCount.TryAdd(fPath, count)
+        _dbg(" ├ 結束", fPath & " | 成果: " & count) ' by Gemini 3.5 Flash, 2026/07/01
         Return count
     End Function
```

#### 4. `GetFolderSizeAll` [L652]
```diff
     Private Async Function GetFolderSizeAll(folder As Folder, Optional fPath As String = "", Optional skipCache As Boolean = False, Optional cToken As CancellationToken = Nothing) As Task(Of Long)
         fPath = SafeGetPath(folder, fPath)
-        If _iLikeNoisy Then _dbg(" ├ 開始", ExtractFolderName(fPath))
+        _dbg(" ├ 開始", fPath) ' by Gemini 3.5 Flash, 2026/07/01
         Dim size As Long
         If Not skipCache Then
             If _cacheFolderSizeAll.TryGetValue(fPath, size) Then Return size    ' ① 記憶體命中
             Dim row = SafeGetDbRow(folder, fPath)                               ' ② DB lazy load (fsa 欄位)
             If row IsNot Nothing AndAlso row.fsa >= 0 Then Return row.fsa
         End If
 
         ' ③ fallback: Layer3 呼叫
         size = Await GetFolderSizeAllOOM(folder, skipCache:=skipCache, cToken:=cToken)
         If size >= 0 Then _cacheFolderSizeAll.TryAdd(fPath, size)
+        _dbg(" ├ 結束", fPath & " | 成果: " & size) ' by Gemini 3.5 Flash, 2026/07/01
         Return size
     End Function
```

#### 5. `GetFolderBasicByEntryIDL3` [L910] (錯誤訊息無條件顯示，結束訊息加在 Finally 中)
```diff
     Private Async Function GetFolderBasicByEntryIDL3(fPath As String, ct As CancellationToken) As Task(Of Dictionary(Of String, MailItemInfo))
         Const PR_MESSAGE_SIZE As String = "http://schemas.microsoft.com/mapi/proptag/0x0E080003"
         Dim ids As (eid As String, sid As String, isMail As Boolean, hasCh As Boolean) = Nothing
         If Not _cacheFolderIDs.TryGetValue(fPath, ids) Then Return Nothing
         Dim folder As Folder = TryCast(_olNS.GetFolderFromID(ids.eid, ids.sid), Folder)
         If folder Is Nothing Then Return Nothing
 
+        _dbg(" ├ 開始", fPath) ' by Gemini 3.5 Flash, 2026/07/01
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
-            If _iLikeNoisy Then _dbg("    ├ GetFolderBasicByEntryIDL3 錯誤", $"{fPath} — {ex.Message}")
+            _dbg("    ├ GetFolderBasicByEntryIDL3 錯誤", $"{fPath} — {ex.Message}") ' by Gemini 3.5 Flash, 2026/07/01
             Return Nothing
         Finally
             TryMarshalRelease(table)
             TryMarshalRelease(folder)
+            _dbg(" ├ 結束", fPath & " | 成果: " & If(result IsNot Nothing, result.Count.ToString(), "Nothing")) ' by Gemini 3.5 Flash, 2026/07/01
         End Try
         Return result
     End Function
```

---

### ■ Layer3 RDO 直接存取底層 (範例)

#### 1. `GetMailCountRdo` [L1275]
```diff
     Private Function GetMailCountRdo(folderPath As String, eid As String, sid As String) As Long
+        If _iLikeNoisy Then _dbg(" ├ 開始", folderPath) ' by Gemini 3.5 Flash, 2026/07/01
         Dim store As Redemption.RDOStore = GetRdoStore(folderPath)
         If store Is Nothing Then Return -1
 
         Dim rdoFolder As Redemption.RDOFolder = Nothing
+        Dim count As Long = -1
         Try
             rdoFolder = TryCast(store.GetFolderFromID(eid), Redemption.RDOFolder)
             If rdoFolder Is Nothing Then Return -1
-            Return CLng(rdoFolder.Items.Count)
+            count = CLng(rdoFolder.Items.Count)
         Catch ex As System.Exception
-            If _iLikeNoisy Then _dbg("GetMailCountRdo 失敗", $"{ExtractFolderName(folderPath)} | {ex.Message}")
+            _dbg("GetMailCountRdo 失敗", $"{folderPath} | {ex.Message}") ' by Gemini 3.5 Flash, 2026/07/01
             Return -1
         Finally
             Dim o As Object = rdoFolder : TryMarshalRelease(o)
+            If _iLikeNoisy Then _dbg(" ├ 結束", folderPath & " | 成果: " & count) ' by Gemini 3.5 Flash, 2026/07/01
         End Try
+        Return count
     End Function
```

#### 2. `GetFolderSizeRdo` [L1319]
```diff
     Private Function GetFolderSizeRdo(folderPath As String, eid As String, sid As String) As Long
+        If _iLikeNoisy Then _dbg(" ├ 開始", folderPath) ' by Gemini 3.5 Flash, 2026/07/01
         Const PR_SIZE_LONG As String = "http://schemas.microsoft.com/mapi/proptag/0x0E080003"  ' PR_MESSAGE_SIZE (PT_LONG)
         Dim store As Redemption.RDOStore = GetRdoStore(folderPath)
         If store Is Nothing Then Return -1
 
         Dim rdoFolder As Redemption.RDOFolder = Nothing
         Dim items As Object = Nothing, tbl As Object = Nothing
+        Dim total As Long = -1
         Try
             rdoFolder = TryCast(store.GetFolderFromID(eid), Redemption.RDOFolder)
             If rdoFolder Is Nothing Then Return -1
             items = rdoFolder.Items
             tbl = items.MAPITable
-            If CInt(tbl.RowCount) = 0 Then Return 0   ' 空夾: 0 bytes (不走 GetRows,防空表邊界)
+            If CInt(tbl.RowCount) = 0 Then 
+                total = 0
+            Else
+                tbl.Columns = PR_SIZE_LONG
+                tbl.GoToFirst()
+                total = 0
+                Do
+                    Dim chunk As Array = TryCast(tbl.GetRows(5000), Array)
+                    If chunk Is Nothing Then Exit Do
+                    Dim got As Integer = 0
+                    For i As Integer = chunk.GetLowerBound(0) To chunk.GetUpperBound(0)
+                        got += 1
+                        Dim row As Array = TryCast(chunk.GetValue(i), Array)
+                        If row Is Nothing Then Continue For
+                        Dim v = row.GetValue(row.GetLowerBound(0))
+                        If v IsNot Nothing AndAlso Not IsDBNull(v) Then total += CLng(v)
+                    Next
+                    If got < 5000 Then Exit Do
+                Loop
+            End If
-
-            tbl.Columns = PR_SIZE_LONG
-            tbl.GoToFirst()
-            Dim total As Long = 0
-            Do
-                Dim chunk As Array = TryCast(tbl.GetRows(5000), Array)
-                If chunk Is Nothing Then Exit Do
-                Dim got As Integer = 0
-                For i As Integer = chunk.GetLowerBound(0) To chunk.GetUpperBound(0)
-                    got += 1
-                    Dim row As Array = TryCast(chunk.GetValue(i), Array)
-                    If row Is Nothing Then Continue For
-                    Dim v = row.GetValue(row.GetLowerBound(0))
-                    If v IsNot Nothing AndAlso Not IsDBNull(v) Then total += CLng(v)
-                Next
-                If got < 5000 Then Exit Do   ' 最後一批不足 → 到底
-            Loop
-            Return total
         Catch ex As System.Exception
-            If _iLikeNoisy Then _dbg("GetFolderSizeRdo 失敗", $"{ExtractFolderName(folderPath)} | {ex.Message}")
+            _dbg("GetFolderSizeRdo 失敗", $"{folderPath} | {ex.Message}") ' by Gemini 3.5 Flash, 2026/07/01
             Return -1
         Finally
             TryMarshalRelease(tbl) : TryMarshalRelease(items)
             Dim o As Object = rdoFolder : TryMarshalRelease(o)
+            If _iLikeNoisy Then _dbg(" ├ 結束", folderPath & " | 成果: " & total) ' by Gemini 3.5 Flash, 2026/07/01
         End Try
+        Return total
     End Function
```

---

### ■ Layer3 OOM 直接存取底層 (範例)

#### 1. `GetMailCountOOM` [L1682]
```diff
     Private Function GetMailCountOOM(folder As Folder, Optional fPath As String = "") As Long
         fPath = SafeGetPath(folder, fPath)
+        If _iLikeNoisy Then _dbg(" ├ 開始", fPath) ' by Gemini 3.5 Flash, 2026/07/01
         Dim count As Long = -1
         Try
             count = CLng(folder.Items.Count)
         Catch ex As System.Exception
-            ' 這裡原無 log
+            _dbg("GetMailCountOOM 失敗", fPath & " | " & ex.Message) ' by Gemini 3.5 Flash, 2026/07/01
         End Try
+        If _iLikeNoisy Then _dbg(" ├ 結束", fPath & " | 成果: " & count) ' by Gemini 3.5 Flash, 2026/07/01
         Return count
     End Function
```
