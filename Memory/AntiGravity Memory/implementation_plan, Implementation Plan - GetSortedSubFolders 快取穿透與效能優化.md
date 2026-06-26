# Implementation Plan - GetSortedSubFolders 快取穿透與效能優化

解決 `Form1_MainTab12.vb` 在強制重刷（F5）時無法繞過記憶體快取的問題，同時優化迴圈內部重複打 COM 讀取資料夾名稱的效能瓶頸，並改善 COM 物件的生命週期管理。

## User Review Required

> [!WARNING]
> **快取字典型別變更**
> 為了讓回傳的結果包含已經讀取好的 `Name`，我們必須將 `_cacheFolderTree` 儲存的值從 `List(Of Folder)` 升級為 `List(Of FolderSortInfo)`。這會牽動到其他清除快取的片段（如 `Form1.vb` 中的 `Clear` 與 `Module_Outlook.vb` 中的 `TryRemove`），這些地方的型別必須一併更新，請確認是否同意。

## Open Questions

目前沒有重大的未決問題，修改範圍已經限縮並確保向前相容性。

## Proposed Changes

---

### Module_Outlook.vb

優化記憶體快取的判斷邏輯，並提供帶有名稱屬性的擴充回傳函數，避免重複讀取 COM。

#### [MODIFY] [Module_Outlook.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/Module_Outlook.vb)
1. **型別與結構公開化**：
   將 `Private Structure FolderSortInfo` 修改為 `Friend Structure FolderSortInfo`，以便外部也能取用已讀取好的 `Name` 等屬性。
2. **快取字典型別更新**：
   將 `_cacheFolderTree` 從 `ConcurrentDictionary(Of String, List(Of Folder))` 改為 `ConcurrentDictionary(Of String, List(Of FolderSortInfo))`。
3. **修復 `skipCache` 穿透失效**：
   在 `TryGetValue` 取出快取時，加上 `AndAlso Not skipCache` 判斷。
   ```vb
   If Not skipCache AndAlso _cacheFolderTree.TryGetValue(cacheKey, cachedFolders) Then Return cachedFolders
   ```
4. **新增函數 `GetSortedSubFoldersWithInfo`**：
   將原先 `GetSortedSubFolders` 的主要邏輯移至此新函數，並讓它直接回傳 `List(Of FolderSortInfo)`（包含快取存取）。
5. **相容性包裝 `GetSortedSubFolders`**：
   讓原本的 `GetSortedSubFolders` 變成一個 Wrapper，內部呼叫 `GetSortedSubFoldersWithInfo`，再 `.Select(Function(x) x.FolderObj).ToList()` 回傳。這樣就不會破壞整個專案中十幾個依賴純 `List(Of Folder)` 的既有程式碼。

---

### Form1_MainTab12.vb

解決迴圈中頻繁讀取 `child.Name` 導致重複發送 COM 呼叫的效能浪費。

#### [MODIFY] [Form1_MainTab12.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/Form1_MainTab12.vb)
1. **替換為高效 API**：
   將原本的迴圈 `For Each child As Folder In GetSortedSubFolders(...)` 改為：
   ```vb
   ' 先存成變數，避免單行過長，並改用 WithInfo 函數以取得無 COM 的快取名稱
   Dim sortedInfos = GetSortedSubFoldersWithInfo(folder, rootPath, skipCache:=True)
   For Each info As FolderSortInfo In sortedInfos
       cToken.ThrowIfCancellationRequested()
       Dim child As Folder = info.FolderObj
       Dim childPath As String = rootPath & "\" & info.Name  ' ✅ 使用結構內已讀好的字串，省去 COM 呼叫
       ' 餘下 Await 邏輯不變...
   ```

## Verification Plan

### Manual Verification
1. 在 UI 點擊 "F5 重整" (Tab12 強制重刷)，觀察 Console/Debug 紀錄是否真正跳過記憶體快取，並看見 `[SSD Hit]` 或 `打 COM` 的紀錄，確保 `skipCache:=True` 發揮作用。
2. 檢查 F5 期間畫面是否卡頓，理論上減少了迴圈內部大量的 `child.Name` COM Interop，畫面反應速度與處理速度應該會明顯提升。
3. 測試展開樹狀節點（不會觸發 `skipCache`），確保原有的 `GetSortedSubFolders` 相容性包裝沒有破壞舊版延遲載入 (Lazy Load) 的功能。
