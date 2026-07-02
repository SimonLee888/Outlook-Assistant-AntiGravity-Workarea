# 盤點與優化 Module_Outlook.vb Debug Message 實作計劃 (精簡版)

本計劃旨在根據最新討論原則，以**最少代碼改動、最高安全性**為原則，優化 `Module_Outlook.vb` 檔案中 **Layer2.5**、**Layer3 RDO** 以及 **Layer3 OOM** 這三個 Region 中的 Debug 訊息。

## 核心修改原則

1. **極簡化結束訊息**：
   - 只有當函數為**單一主要 return 出口**，且不需重構控制結構（如引入新變數或改寫 `Try-Catch` / `Do Loop`）時，才在結尾處加上單行結束訊息。
   - 若函數有多個提早 exit return 分支，且為了加結束訊息需要大幅改動程式碼結構者，**一律不加結束訊息**。
2. **開始訊息**：
   - L2.5 進入端無條件顯示：`_dbg(" ├ 開始", fPath)`
   - L3 RDO / OOM 進入端在 `If _iLikeNoisy Then` 條件下顯示：`If _iLikeNoisy Then _dbg(" ├ 開始", fPath)`
3. **錯誤與例外訊息**：
   - 所有 Catch 區塊內的錯誤 Log 均改為無條件顯示（不加 `_iLikeNoisy` 條件），方便預設排查錯誤。
4. **註記規範**：
   - 修改處加註：`' by Gemini 3.5 Flash, 2026/07/01`

---

## 修正方案對照範例

### ■ Layer2.5 函數 (`GetMailCount`) - 單一出口，僅加一行

**修改前**：
```vb
    Private Function GetMailCount(folder As Folder, Optional fPath As String = "", Optional skipCache As Boolean = False) As Long
        fPath = SafeGetPath(folder, fPath)
        Dim count As Long
        If Not skipCache Then
            If _cacheMailCount.TryGetValue(fPath, count) Then Return count       ' ① 記憶體命中
            Dim row = SafeGetDbRow(folder, fPath)                                ' ② DB lazy load
            If row IsNot Nothing AndAlso row.mc >= 0 Then Return row.mc
        End If

        ' ③ 讀取派工: RDO 優先,失敗 fallback OOM
        count = GetMailCountRdo(fPath, folder.EntryID, folder.StoreID)
        If count < 0 Then count = GetMailCountOOM(folder, fPath:=fPath)
        If count >= 0 Then _cacheMailCount.TryAdd(fPath, count)
        Return count
    End Function
```

**修改後** (只加 2 行，無任何結構改動)：
```vb
    Private Function GetMailCount(folder As Folder, Optional fPath As String = "", Optional skipCache As Boolean = False) As Long
        fPath = SafeGetPath(folder, fPath)
+        _dbg(" ├ 開始", fPath) ' by Gemini 3.5 Flash, 2026/07/01
        Dim count As Long
        If Not skipCache Then
            If _cacheMailCount.TryGetValue(fPath, count) Then Return count       ' ① 記憶體命中
            Dim row = SafeGetDbRow(folder, fPath)                                ' ② DB lazy load
            If row IsNot Nothing AndAlso row.mc >= 0 Then Return row.mc
        End If

        ' ③ 讀取派工: RDO 優先,失敗 fallback OOM
        count = GetMailCountRdo(fPath, folder.EntryID, folder.StoreID)
        If count < 0 Then count = GetMailCountOOM(folder, fPath:=fPath)
        If count >= 0 Then _cacheMailCount.TryAdd(fPath, count)
+        _dbg(" ├ 結束", fPath & " | 成果: " & count) ' by Gemini 3.5 Flash, 2026/07/01
        Return count
    End Function
```

---

### ■ Layer3 RDO 函數 (`GetMailCountRdo`) - 多個出口，不加結束訊息，只改 Exception 條件

**修改前**：
```vb
    Private Function GetMailCountRdo(folderPath As String, eid As String, sid As String) As Long
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
```

**修改後** (不加結束訊息，避免改動 Return 結構，只改 Catch 的 Log 條件，並加 noisy 開始訊息)：
```vb
    Private Function GetMailCountRdo(folderPath As String, eid As String, sid As String) As Long
+        If _iLikeNoisy Then _dbg(" ├ 開始", folderPath) ' by Gemini 3.5 Flash, 2026/07/01
        Dim store As Redemption.RDOStore = GetRdoStore(folderPath)
        If store Is Nothing Then Return -1

        Dim rdoFolder As Redemption.RDOFolder = Nothing
        Try
            rdoFolder = TryCast(store.GetFolderFromID(eid), Redemption.RDOFolder)
            If rdoFolder Is Nothing Then Return -1
            Return CLng(rdoFolder.Items.Count)
        Catch ex As System.Exception
-            If _iLikeNoisy Then _dbg("GetMailCountRdo 失敗", $"{ExtractFolderName(folderPath)} | {ex.Message}")
+            _dbg("GetMailCountRdo 失敗", $"{folderPath} | {ex.Message}") ' by Gemini 3.5 Flash, 2026/07/01
            Return -1
        Finally
            Dim o As Object = rdoFolder : TryMarshalRelease(o)
        End Try
    End Function
```

---

### ■ Layer3 OOM 函數 (`GetMailCountOOM`) - 無結束訊息，僅加開始 Log 與 Catch Log

**修改前**：
```vb
    Private Function GetMailCountOOM(folder As Folder, Optional fPath As String = "") As Long
        fPath = SafeGetPath(folder, fPath)
        Try
            Return CLng(folder.Items.Count)
        Catch ex As System.Exception
            ' 這裡原無 log
            Return -1
        End Try
    End Function
```

**修改後** (只加 2 行，無任何 Return 結構改動)：
```vb
    Private Function GetMailCountOOM(folder As Folder, Optional fPath As String = "") As Long
        fPath = SafeGetPath(folder, fPath)
+        If _iLikeNoisy Then _dbg(" ├ 開始", fPath) ' by Gemini 3.5 Flash, 2026/07/01
        Try
            Return CLng(folder.Items.Count)
        Catch ex As System.Exception
+            _dbg("GetMailCountOOM 失敗", fPath & " | " & ex.Message) ' by Gemini 3.5 Flash, 2026/07/01
            Return -1
        End Try
    End Function
```

---

### ■ Layer3 OOM 函數 (`GetSubtreeOOM`) - 大型 BFS，僅調整 Log 條件，不重構結構

**修改前**：
```vb
    Private Async Function GetSubtreeOOM(...) As Task(...)
        Dim rootPath As String = SafeGetPath(rootFolder)
        ...
        Try
            ...
        Catch ex As System.Exception
            If _iLikeNoisy Then _dbg("    ├ ② OOM BFS 失敗", ex.Message)
        End Try
        ...
        Dim oomMsg As String = $"✓ OOM BFS 子樹完成: {rootFolder.Name}, {result.Count} 夾"
        _dbg("    ├ 結束", oomMsg)
        Return result
    End Function
```

**修改後** (入口/出口 Log 依 `_iLikeNoisy` 判定，錯誤 Log 則預設無條件顯示)：
```vb
    Private Async Function GetSubtreeOOM(...) As Task(...)
        Dim rootPath As String = SafeGetPath(rootFolder)
+        If _iLikeNoisy Then _dbg(" ├ 開始", rootPath) ' by Gemini 3.5 Flash, 2026/07/01
        ...
        Try
            ...
        Catch ex As System.Exception
-            If _iLikeNoisy Then _dbg("    ├ ② OOM BFS 失敗", ex.Message)
+            _dbg("    ├ ② OOM BFS 失敗", ex.Message) ' by Gemini 3.5 Flash, 2026/07/01
        End Try
        ...
        Dim oomMsg As String = $"✓ OOM BFS 子樹完成: {rootFolder.Name}, {result.Count} 夾"
-        _dbg("    ├ 結束", oomMsg)
+        If _iLikeNoisy Then _dbg(" ├ 結束", rootPath & " | " & oomMsg) ' by Gemini 3.5 Flash, 2026/07/01
        Return result
    End Function
```

---

## 驗證計劃

### 手動驗證
1. 編譯專案，確保無語法或型別錯誤。
2. 在開啟與關閉 `_iLikeNoisy` 狀態下，觀察 Debug 視窗的輸出，驗證：
   - Layer2.5 的開始與結束是否始終顯示（只有單一出口的函數有結束）。
   - Layer3 (RDO/OOM) 的開始與結束是否只有在 `_iLikeNoisy = True` 時顯示。
   - 錯誤訊息在 `_iLikeNoisy = False` 時依然能正常輸出。
