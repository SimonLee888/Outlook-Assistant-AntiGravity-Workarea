# 修復資料夾過濾 (_showAllFolders) 失效計劃

修復目標：確保在未勾選「顯示所有資料夾」時，系統能正確隱藏非郵件類資料夾（如行事曆、聯絡人等）。

## User Review Required

> [!IMPORTANT]
> **快取鍵值 (Cache Key) 隔離**：
> `_cacheFolderTree` 的鍵值將從 `fPath` 改為 `fPath & "|" & _showAllFolders`。
> **目的**：避免「切換勾選狀態」後因為命中舊的記憶體快取而導致顯示結果錯誤。這是一個純記憶體運算，不會產生額外的 COM 開銷。

## Proposed Changes

---

### [Component] Form1_Outlook.vb

#### [MODIFY] [Form1_Outlook.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_Outlook.vb)

**1. `GetSortedSubFolders` 邏輯修復：**
- 修改快取讀取與寫入的 Key：
  - `Dim cacheKey As String = fPath & "|" & _showAllFolders`
  - 使用 `cacheKey` 取代原本的 `fPath` 進行 TryGetValue 與 賦值。
- 在底層 COM 遍歷迴圈（Fallback 路徑）中加入過濾判斷：
  ```vb
  For Each subF As Outlook.Folder In subs
      Dim name As String = subF.Name
      ' 修復：手動拼接子路徑傳入 IsMailFolder，實現「零額外 COM 呼叫」的過濾判斷
      Dim childPath As String = fPath & "\" & name
      If Not _showAllFolders AndAlso Not IsMailFolder(subF, childPath) Then Continue For
      
      infoList.Add(New FolderSortInfo With { ... })
  Next
  ```

**2. `IsMailFolder` 穩固性檢查：**
- 確認 `IsMailFolder` 已優先檢查 `_cacheIsMailFolder`，確保過濾判斷的效能。

---

### [Component] 其他連連鎖效應檢查

- **Form1_MainTabs.vb**: `BuildBfsFolderTree` 內部呼叫了 `GetSortedSubFolders`，將自動獲得過濾能力，無需額外改動。
- **Form1.vb**: `LoadSubFolderToTreeView` 內部呼叫了 `GetSortedSubFolders`，將自動獲得過濾能力，無需額外改動。

## Verification Plan

### Automated Tests
- 開啟 `_iLikeNoisy` 模式，觀察 Debug 視窗是否出現「過濾非郵件資料夾」的訊息。
- 測試「勾選」與「取消勾選」顯是所有資料夾，確認 TreeView 能即時反應正確內容。

### Manual Verification
- 展開含有非郵件資料夾（如：Calendar, Contacts, Tasks）的 Store。
- 確認在 `_showAllFolders = False` 時，這些資料夾不應該出現在 TreeView 也不應該出現在 Tab1 的統計清單中。
