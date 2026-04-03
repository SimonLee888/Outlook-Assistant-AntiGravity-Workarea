# 非郵件資料夾與排序效能優化計劃 (更新版)

針對您剛才的提問，我已經在下方逐一回覆，並重新整理了實作方案。

## User Review Required (問題回覆)

> **Q: 一次的COM呼叫大約會花多久時間?**
**A:** 一次 OOM 跨行程 (Out-of-Process) COM 呼叫的開銷，在緩存模式下通常是 **幾毫秒 (1~5ms)**，但在連線 Exchange 或檔案龐大時可能飆升到 **10~20ms 以上**。
假設有 100 個子資料夾，`OrderBy` 的比較次數約為 `O(N log N) ≈ 460次`。如果每次比較都讀取兩遍 `.Name`，總計可能打將近 1000 次 COM。這就意味著單純一個排序就會造成 **主執行緒凍結 1 秒到數秒**！

> **Q: 方案 A 怎麼寫，先給我看一下**
**A:** 寫法是將 COM 取值限制為嚴格的 $O(N)$（也就是只有 100 次），放入純記憶體結構後再排序：
```vb
Private Structure FolderInfo
    Public FolderObj As Outlook.Folder
    Public Name As String
    Public HasChinese As Boolean
End Structure

' ... 在 GetSortedSubFolders 中：
Dim infoList As New List(Of FolderInfo)(folder.Folders.Count)
For Each subF As Outlook.Folder In folder.Folders
    Dim fName As String = subF.Name    ' [COM 呼叫] 僅一次
    infoList.Add(New FolderInfo With {
        .FolderObj = subF,
        .Name = fName,
        .HasChinese = TextHasChineseChar(fName)
    })
Next

' 針對記憶體中的變數排序 (不再觸發 COM)
Dim sortedInfos = infoList.OrderBy(Function(info) If(info.HasChinese, 1, 0)).
                           ThenBy(Function(info) info.Name).ToList()

' 再 Select 回原本的 Folder List
Dim sortedFolders = sortedInfos.Select(Function(info) info.FolderObj).ToList()
```

> **Q: OOM 也可以讀 MAPI table 嗎？/ RDO 轉 OOM 工具函數**
**A:** 
1. **OOM GetTable**: 可以！自 Outlook 2007 以後，OOM 支援 `folder.Folders.GetTable()`。我們可以將這當作 OOM 路徑下最高速的 Fallback！
2. **互相轉換工具**: 透過唯一的 `EntryID` 與 `StoreID` 就能完美雙向轉換。我會幫您加入以下兩個轉型工具：
```vb
' OOM 轉 RDO
Public Function GetRdoFolder(oomFolder As Outlook.Folder) As Redemption.RDOFolder
    If _rdo Is Nothing Then Return Nothing
    Return _rdo.GetFolderFromID(oomFolder.EntryID, oomFolder.StoreID)
End Function

' RDO 轉 OOM (需要在 Form1 內調用建立的 outlookApp 實例)
Public Function GetOomFolder(rdoFolder As Redemption.RDOFolder) As Outlook.Folder
    Dim ns As Outlook.NameSpace = globalOutlookApp.Session
    Return ns.GetFolderFromID(rdoFolder.EntryID, rdoFolder.StoreID)
End Function
```

> **Q: 過濾非郵件目錄應該在哪一層？**
**A: 絕對是 L3 (純資料獲取層)。**
在 L3 就把它擋掉（不放入回傳名單），這樣 L2.5 的快取、L2 的 UI 呈現與遞迴計算，就根本不會意識到它們的存在，省下極大的記憶體與迴圈開銷。
**過濾條件**：
- 在 OOM 裡就是判斷 `subF.DefaultItemType = Outlook.OlItemType.olMailItem`
- 在 MAPITable / RDO 裡就是看 `PR_CONTAINER_CLASS` 是否為 `"IPF.Note"` (包含郵件)。

## Proposed Changes

---

### [Form1.vb] 新增 OOM/RDO 轉換工具
- 將上述的兩個轉換函數封裝獨立，供全專案備用。

### [Form1.vb] 改造 GetSortedSubFolders (導入 MAPI Table 與過濾)
取代原有的 LINQ，我們將採用 OOM 的 `GetTable()` 做到高速全抓 + L3 過濾：

```vb
Private Function GetSortedSubFolders(folder As Outlook.Folder) As List(Of Outlook.Folder)
    Dim fPath As String = folder.FolderPath
    If _cacheFolderTree.ContainsKey(fPath) Then Return _cacheFolderTree(fPath)

    Dim resultList As New List(Of Outlook.Folder)()
    
    ' TODO: 利用 OOM 的 Table 高速查詢
    ' 1. 取得 Table: Dim tb As Outlook.Table = folder.Folders.GetTable()
    ' 2. 獲取資料夾 EntryID，只過濾出 "IPF.Note" 的欄位
    ' 3. 從 Application.Session.GetFolderFromID(ID) 返回對象
    ' 4. 依照名稱等條件進行記憶體內 Sorting
    
    ' 將結果存入 _cacheFolderTree
End Function
```
⚠️ *注意：雖然 Table 很快，但在最終還是得用 `GetFolderFromID` 把 `Outlook.Folder` 物件實例化出來放進 List 中給 UI Tree 綁定。但這已經省去了對「行事曆、聯絡人」的實例化與 COM Sorting 痛苦，速度至少提昇一半！*

---

## Open Questions
- 您希望在 `GetSortedSubFolders` 這個回傳 UI 綁定對象的函數裡，就直接強制過濾掉非郵件（行事曆等）目錄，讓它們**完全不顯示在左側樹狀清單裡**嗎？（這將大大簡化後續所有動作）。

## Verification Plan
1. 等您核准此計劃，我會著手撰寫 `GetSortedSubFolders` 的 MAPITable + Filtering 防護代碼，並建立 OOM/RDO 轉換函數。
