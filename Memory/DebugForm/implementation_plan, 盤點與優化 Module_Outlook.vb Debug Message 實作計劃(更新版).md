# 盤點與優化 Module_Outlook.vb Debug Message 實作計劃

本計劃旨在盤點並統一 `Module_Outlook.vb` 檔案中 **Layer2.5**、**Layer3 RDO** 以及 **Layer3 OOM** 這三個 Region 中的 Debug 訊息，以利於問題排查與效能監控。

## 調整原則

1. **Layer2.5 (快取代理層)**
   - **進入端**：無條件顯示 `_dbg(" ├ 開始", fPath)`（於 `SafeGetPath` 解析完成後）。
   - **結束端**：無條件顯示 `_dbg(" ├ 結束", fPath & " | 成果: " & [成果值])`。成果值必須為現成已計算之變數，不產生額外開銷。
   - **錯誤與例外**：若執行過程中發生錯誤或例外，無條件顯示相關的錯誤訊息（不加 `_iLikeNoisy` 判斷）。

2. **Layer3 RDO 與 Layer3 OOM (底層直接存取)**
   - **進入與結束端**：必須加上 `If _iLikeNoisy Then` 條件，只有打開該開關時才顯示。
     - 進入端格式：`If _iLikeNoisy Then _dbg(" ├ 開始", fPath)`
     - 結束端格式：`If _iLikeNoisy Then _dbg(" ├ 結束", fPath & " | 成果: " & [成果值])`
   - **過程中的錯誤與例外**：若發生錯誤或例外，預設顯示，**不須**加上 `_iLikeNoisy` 條件。

---

## 預計修改函數清單

### 1. Layer2.5 快取代理層 (共 23 個函數)
將在以下函數的入口處與結束/回傳處新增或修正 `_dbg` 輸出，並調整例外處理的 Log 輸出：
- `GetMailCount` (多載 1: folder, 多載 2: fPath)
- `GetFolderCount` (多載 1: folder, 多載 2: fPath)
- `GetFolderSize` (多載 1: folder, 多載 2: fPath)
- `GetFolderSizeAll`
- `GetYearCountsForFolder`
- `GetMonthCountsForYear`
- `GetAttachMailList`
- `GetAttachFilename`
- `GetMailBody`
- `GetBasicMailInfo` (多載 1: folder, 多載 2: fPath)
- `GetFolderBasicByEntryIDL3` (此函數包含 Try-Catch，將錯誤訊息改為無條件顯示)
- `PreLoadBasicMailCacheAsync` (維持現有的無條件開始/結束顯示)
- `GetSubtree`
- `GetSortedSubFolders`
- `GetSortedSubFolderIDs`
- `HasSubFoldersFast`
- `IsMailFolder`
- `InvalidateFolderTreeCache`
- `InvalidateBasicMailCache`

### 2. Layer3 RDO 底層 (共 11 個函數)
將在以下函數的入口與結束處新增/修改加上 `_iLikeNoisy` 限制的 `_dbg`，並將 Catch 區塊內的錯誤 Log 改為無條件輸出：
- `GetMailCountRdo` (錯誤 Log `GetMailCountRdo 失敗` 改為無條件顯示)
- `GetFolderCountRdo` (錯誤 Log `GetFolderCountRdo 失敗` 保留無條件顯示)
- `GetFolderSizeRdo` (錯誤 Log `GetFolderSizeRdo 失敗` 改為無條件顯示)
- `GetAttachFilenameRdo` (錯誤 Log `GetAttachFilenameRdo 失敗` 改為無條件顯示)
- `GetMailBodyRdo` (錯誤 Log `GetMailBodyRdo 失敗` 保留無條件顯示)
- `RefreshMailInfoL3` (Catch 中的錯誤與 NotFound 等 Log 改為無條件顯示)
- `GetRdoStore`
- `GetSubtreeRdo`
- `GetSubtreeRdoBatch` (錯誤 Log `RDO 批次失敗` 保留無條件顯示)
- `GetSubtreeRdoEnum` (錯誤 Log `RDO 枚舉失敗` 保留無條件顯示)
- `RdoTableEidToHex` (無 debug 訊息，不修改)

### 3. Layer3 OOM 底層 (共 14 個函數)
將在以下函數的入口與結束處新增/修改加上 `_iLikeNoisy` 限制的 `_dbg`，並將 Catch 區塊內的錯誤 Log 改為無條件輸出：
- `GetMailCountOOM`
- `GetFolderCountOOM`
- `GetMailCountAllOOM`
- `GetFolderCountAllOOM`
- `GetFolderSizeOOM`
- `GetFolderSizeAllOOM`
- `GetYearCountsForFolderL3`
- `GetMonthCountsForYearL3`
- `GetAttachMailListOOM`
- `GetAttachFilenameOOM`
- `GetMailBodyOOM`
- `GetBasicMailInfoOOM`
- `GetSubtreeOOM` (Catch 內的 `OOM 失敗`、`中斷`、`OOM BFS 失敗` 改為無條件顯示)
- `GetLiveFolderSnapOOM`

---

## 統一的訊息格式規範

1. **進入端 (開始)**
   - 格式：`_dbg(" ├ 開始", fPath)`
   - 說明：fPath 為完整的資料夾路徑。若是單一郵件操作，則為 `mail.FolderPath & "\" & mail.Subject`。
2. **結束端 (結束)**
   - 格式：`_dbg(" ├ 結束", fPath & " | 成果: " & [成果值])`
   - 成果值範例：
     - 數量或大小：`count`、`size`
     - 列表/字典：`If(result IsNot Nothing, result.Count.ToString(), "Nothing")`
     - 布林值：`isMail`
3. **註記規範**
   - 所有的修改處與新增註解將使用以下標記：
     `' by Gemini 3.5 Flash, 2026/07/01` (使用當天日期)

---

## 修改前後範例對照

### 範例 1：Layer2.5 快取代理層函數 (`GetMailCount`)

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

**修改後**：
```vb
    Private Function GetMailCount(folder As Folder, Optional fPath As String = "", Optional skipCache As Boolean = False) As Long
        fPath = SafeGetPath(folder, fPath)
        _dbg(" ├ 開始", fPath) ' by Gemini 3.5 Flash, 2026/07/01
        Dim count As Long
        If Not skipCache Then
            If _cacheMailCount.TryGetValue(fPath, count) Then 
                _dbg(" ├ 結束", fPath & " | 成果: " & count) ' by Gemini 3.5 Flash, 2026/07/01
                Return count       ' ① 記憶體命中
            End If
            Dim row = SafeGetDbRow(folder, fPath)                                ' ② DB lazy load
            If row IsNot Nothing AndAlso row.mc >= 0 Then 
                _dbg(" ├ 結束", fPath & " | 成果: " & row.mc) ' by Gemini 3.5 Flash, 2026/07/01
                Return row.mc
            End If
        End If

        ' ③ 讀取派工: RDO 優先,失敗 fallback OOM
        count = GetMailCountRdo(fPath, folder.EntryID, folder.StoreID)
        If count < 0 Then count = GetMailCountOOM(folder, fPath:=fPath)
        If count >= 0 Then _cacheMailCount.TryAdd(fPath, count)
        _dbg(" ├ 結束", fPath & " | 成果: " & count) ' by Gemini 3.5 Flash, 2026/07/01
        Return count
    End Function
```

### 範例 2：Layer3 RDO 函數 (`GetMailCountRdo`)

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

**修改後**：
```vb
    Private Function GetMailCountRdo(folderPath As String, eid As String, sid As String) As Long
        If _iLikeNoisy Then _dbg(" ├ 開始", folderPath) ' by Gemini 3.5 Flash, 2026/07/01
        Dim store As Redemption.RDOStore = GetRdoStore(folderPath)
        If store Is Nothing Then 
            If _iLikeNoisy Then _dbg(" ├ 結束", folderPath & " | 成果: -1") ' by Gemini 3.5 Flash, 2026/07/01
            Return -1
        End If

        Dim rdoFolder As Redemption.RDOFolder = Nothing
        Try
            rdoFolder = TryCast(store.GetFolderFromID(eid), Redemption.RDOFolder)
            If rdoFolder Is Nothing Then 
                If _iLikeNoisy Then _dbg(" ├ 結束", folderPath & " | 成果: -1") ' by Gemini 3.5 Flash, 2026/07/01
                Return -1
            End If
            Dim count As Long = CLng(rdoFolder.Items.Count)
            If _iLikeNoisy Then _dbg(" ├ 結束", folderPath & " | 成果: " & count) ' by Gemini 3.5 Flash, 2026/07/01
            Return count
        Catch ex As System.Exception
            ' 過程中的錯誤訊息，預設無條件顯示，不須加 _iLikeNoisy 條件
            _dbg("GetMailCountRdo 失敗", $"{folderPath} | {ex.Message}") ' by Gemini 3.5 Flash, 2026/07/01
            Return -1
        Finally
            Dim o As Object = rdoFolder : TryMarshalRelease(o)
        End Try
    End Function
```

---

## 驗證計劃

### 手動驗證
1. 編譯專案，確保無語法或型別錯誤。
2. 在開啟與關閉 `_iLikeNoisy` 狀態下，觀察 Debug 視窗的輸出，驗證：
   - Layer2.5 的開始與結束是否始終顯示。
   - Layer3 (RDO/OOM) 的開始與結束是否只有在 `_iLikeNoisy = True` 時顯示。
   - 錯誤訊息在 `_iLikeNoisy = False` 時依然能正常輸出。
