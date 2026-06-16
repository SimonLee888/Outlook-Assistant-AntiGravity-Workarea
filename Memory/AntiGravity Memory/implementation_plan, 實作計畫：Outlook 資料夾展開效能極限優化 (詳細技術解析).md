# 實作計畫：Outlook 資料夾展開效能極限優化 (詳細技術解析)

## 1. 核心瓶頸與優化原理
- **瓶頸探究**：在 `For Each subFolder In folder.Folders` 循環中，讀取 `subFolder.Name` 或 `subFolder.DefaultItemType` 都是一次跨進程或跨網路的同步呼叫。100 個資料夾會產生 200 次這類呼叫。
- **優化方案 (MAPITable)**：將「逐個詢問屬性」改為「批量拉取表格」。這是在資料庫層級 (MAPI Store) 進行的操作，效能提升可達 100 倍以上。

## 2. 互轉小函數 (RDO <-> OOM)
您要求的互轉函數，將放置於輔助函數區：

```vb
' [Helper] 將 OOM Folder 轉為 RDO Folder
Private Function OOMToRDO(oomFolder As Outlook.Folder) As Redemption.RDOFolder
    If _rdo Is Nothing OrElse oomFolder Is Nothing Then Return Nothing
    Return _rdo.GetFolderFromID(oomFolder.EntryID, oomFolder.StoreID)
End Function

' [Helper] 將 RDO Folder 轉為 OOM Folder (供 TreeView.Tag 使用)
Private Function RDOToOOM(rdoFolder As Redemption.RDOFolder) As Outlook.Folder
    If rdoFolder Is Nothing Then Return Nothing
    Return _olNS.GetFolderFromID(rdoFolder.EntryID, rdoFolder.StoreID)
End Function
```

## 3. MAPITable 實作細節說明

| 關鍵步驟 | 實作細節與用途 | 為什麼這樣做比較快？ |
| :--- | :--- | :--- |
| **GetHierarchyTable** | 取得該資料夾的「子目錄索引表」。 | 像資料庫的 `SELECT *`，不實體化對象就拿到索引。 |
| **Columns (屬性標籤)** | 指定 PR_DISPLAY_NAME, PR_ENTRYID, PR_CONTAINER_CLASS。 | 只抓我們需要的 3 個欄位，減少數據傳輸量。 |
| **Restrict (過濾)** | 僅讀取 `IPF.Note` (郵件) 資料夾。 | 在伺服器端就把非郵件目錄篩掉，不回傳到前端。 |
| **GetRows (2D 陣列)** | 將整張表一次抓進記憶體陣列中。 | 核心優化：將 N 次 COM 網返總結為 **1 次**。 |
| **EntryIDToString** | 將 MAPI 的二進位 ID 轉為 Hex 字串。 | 這是最後轉回 OOM Folder 給 TreeView 使用的關鍵索引。 |

## 4. 修改方案比較 (TryMarshalRelease)

### 修改前 (Before)
每個函數都要寫重複的 Try-Finally，且不保證對象是否已經失效：
```vb
Finally
    If rdoFolder IsNot Nothing Then Marshal.ReleaseComObject(rdoFolder)
End Try
```

### 修改後 (After)
統一處理，包含判斷物件是否為 COM 或是否已釋放：
```vb
' 用於每一處 Finally 區塊
Finally
    TryMarshalRelease(table)
    TryMarshalRelease(rdoFolder)
End Sub
```

## 5. 實施流程與清理

### 步驟 A：恢復 Form1.vb 的健康狀態 (優先執行)
- **清理**：移除 line 814-817 產生之語法錯誤。
- **合併**：確保 `GetSortedSubFolders` 只有一個單一入口。

### 步驟 B：實作核心優化代碼
1. 在 `GetSortedSubFolders` 中偵測 `_rdo` 是否就緒。
2. 就緒則進入 **MAPITable 軌道**，依序執行 Hierarchy -> Columns -> Restrict -> GetRows。
3. 利用 `RDOToOOM` 將結果轉換回 `Outlook.Folder` 並存入 `FolderSortInfo` 清單。
4. **記憶體排序**：依據快取到的 Name 屬性進行中文優先排序。

---

## 6. 驗證計畫
- **功能驗證**：確保展開資料夾後，內容與原本顯示一致，且「顯示全部」開關依然有效。
- **效能驗證**：在日誌中輸出耗時。目標是資料夾展開由「有感延遲」進度為「零感展開」。

**請查看以上規畫。如果您同意這個細節度，請回覆「確認開始」。**
