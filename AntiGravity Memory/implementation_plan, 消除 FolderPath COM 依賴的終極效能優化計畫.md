# 消除 FolderPath COM 依賴的終極效能優化計畫

這是一份針對消除 `GetUniqueFolderList` 和 `GetCachedMailCount` 迴圈中，因為二次索取 `.FolderPath` 所造成的總計 **約 760ms** 延遲的優化計畫。

## 目標與範圍控制 (Impact Scope)

這次的變更主要限縮在底層快取機制與 `GetSubtreeToList` 這個核心函式。為了將「路徑字串」成功帶出，我們將全面替換它回傳的型態。

### 影響的核心範圍：
*   **1 個快取結構**: `_cacheSubFolderList`
*   **2 個核心函式**: `GetSubtreeToList`, `GetUniqueFolderList`
*   **10 個呼叫端點**: 所有呼叫 `GetSubtreeToList` 或 `GetUniqueFolderList` 的地方需要對齊新介面。

## Proposed Changes

### Form1_Outlook.vb

#### [MODIFY] 1. 更新快取字典宣告
將原本的 `_cacheSubFolderList` 升級為接受 Tuple 的字典。
```diff
-Private Shared _cacheSubFolderList As New ConcurrentDictionary(Of String, List(Of Outlook.Folder))
+Private Shared _cacheSubFolderList As New ConcurrentDictionary(Of String, List(Of (Folder As Outlook.Folder, FolderPath As String)))
```

#### [MODIFY] 2. 升級 GetSubtreeToList
將回傳值從單純的 Folder List 改為 Tuple List。利用其內部已經建構好 BFS 的字串優勢直接封裝。
```diff
-Private Async Function GetSubtreeToList(...) As Task(Of List(Of Outlook.Folder))
+Private Async Function GetSubtreeToList(...) As Task(Of List(Of (Folder As Outlook.Folder, FolderPath As String)))
...
' 內部 BFS 迴圈和 SSD 讀取部分，配合回傳型態組裝為 (subF, childPath) 的 Tuple。
```

#### [MODIFY] 3. 升級 GetUniqueFolderList
同步接收 Tuple，從大處攔截，並將這份包含預先計算好路徑的 Tuple 完整送交給 Tab2、Tab3 終端。
```diff
-Private Async Function GetUniqueFolderList(...) As Task(Of List(Of Outlook.Folder))
+Private Async Function GetUniqueFolderList(...) As Task(Of List(Of (Folder As Outlook.Folder, FolderPath As String)))
```
*在 HashSet 去重時，直接讀取 `subF.FolderPath` (此時為 Tuple 屬性，非 COM 操作)，瞬間完成。*

### Form1_MainTabs.vb

#### [MODIFY] Tab 2 & Tab 3 呼叫端對齊
利用新傳回來的 Tuple 完美避開 `.Select(Function(f) f.FolderPath).ToList()` 所造成的另一波連環 COM 呼叫。
```diff
-Dim folderPaths = folderList.Select(Function(f) f.FolderPath).ToList()
+Dim folderPaths = folderList.Select(Function(f) f.FolderPath).ToList() ' 完全零 COM 開銷，瞬間完成
```

### Form1_SQLite2.vb & Form1_Win32API.vb
#### [MODIFY] 餘下 8 個不關注路徑的 Caller 對齊
不影響這幾個地方原本的操作，他們只要加一句 LINQ 解開 Tuple 拿走 `Folder` 物件即可。
```diff
-Dim targetFolderList = Await GetSubtreeToList(root, True)
+Dim targetFolderList = (Await GetSubtreeToList(root, True)).Select(Function(x) x.Folder).ToList()
```

## User Review Required

> [!CAUTION]
> **簽章變更影響**：
> 將 `GetSubtreeToList` 回傳 Tuple 雖然完美解決了 COM 往返成本，但也牽動了 10 個周邊的呼叫區塊。我有信心乾淨俐落地處理好每一個依賴轉換。
> 我們將先進行修改，完成後並測試 Tab 3 搜尋時的 `效能 (1/4)` 和 `效能 (2/4)` 是否達到我們預測的 **近乎 0ms**。

請問計畫是否符合您的「影響範圍可控」期待，可以開始執行了呢？
