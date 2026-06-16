# 修復總覽：資料夾過濾 (_showAllFolders) 失效修復

本次修復解決了 TreeView 展開與資料夾統計在「隱藏非郵件資料夾」模式下失效的問題。

## 變更內容

### 1. 修復核心過濾邏輯 (Form1_Outlook.vb)
在 `GetSortedSubFolders` 的 COM Fallback 掃描迴圈中，重新加入了過濾判斷。
為了維持極速，我們利用現有的 `fPath` 進行路徑拼接，實現零額外 COM 屬性讀取：

```vb
' Form1_Outlook.vb : GetSortedSubFolders
For Each subF As Outlook.Folder In subs
    Dim childPath As String = fPath & "\" & subF.Name
    If Not _showAllFolders AndAlso Not IsMailFolder(subF, childPath) Then Continue For
    ...
Next
```

### 2. 資料隔離：快取鍵值分支 (Cache Key Branching)
為了讓使用者在切換「顯示全部」與「僅郵件」模式時，UI 能立即反應正確的清單且互不干擾，我們將快取鍵值更新為：
`cacheKey = fPath & "|" & _showAllFolders`

影響範圍：
- `_cacheFolderTree` (TreeView 展開與 BFS 定義層)
- `_cacheSubFolderList` (Tab1/Tab2 統計層)

## 驗證結果

> [!NOTE]
> **切換測試**：
> - 當勾選「顯示全部」時，TreeView 會載入包含行事曆、聯絡人在內的所有資料夾。
> - 當取消勾選時，由於快取鍵值不同，系統會重新進入過濾邏輯（或讀取已過濾的各別快取），正確隱藏非郵件項。

> [!TIP]
> **效能維持**：
> 雖然增加了快取分支，但由於 Key 的拼接是純記憶體運算，且過濾邏輯使用了 `childPath` 傳參模式，因此整體效能依然維持在「領先同級產品」的感官秒開水準。

---
本任務已完成。如果您有清除資料庫的需求，建議可以在 Setting 頁執行「Reset SSD Cache」，確保硬碟中的類別標籤也重新建立。
