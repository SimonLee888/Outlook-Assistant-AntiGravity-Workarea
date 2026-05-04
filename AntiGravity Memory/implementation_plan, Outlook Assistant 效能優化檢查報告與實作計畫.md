# Outlook Assistant 效能優化檢查報告與實作計畫

根據您提供的 8 項優化重點，我已經掃描了專案內所有的 `*.vb` 檔案（排除了 `moduleStore.vb` 以及待刪除區塊），以下是逐一檢查的結果、詳細說明與後續實作計畫：

## 檢查報告

### 1. 其他小優化（很容易做）排序只做一次（不要每次呼叫都 Sort）
- **狀態**：✅ **已實作**
- **檢查結果**：在 `Form1_Outlook.vb` 的 `GetSortedSubFolders` 中，已經設計了 `_cacheFolderTree` 這個 `ConcurrentDictionary` 作為快取。
- **說明**：當第一次完成 `OrderBy(...).ThenBy(...)` 的 LINQ 排序後，排序結果會被存入記憶體快取。下一次展開或呼叫時直接返回快取清單，不會發生重複排序。

### 2. 用 ConcurrentDictionary 的 GetOrAdd 包裝整個函數，避免重複建置
- **狀態**：❌ **未實作**
- **檢查結果**：全域搜查 `GetOrAdd` 並未發現相關使用。
- **說明**：目前的快取層（例如 `GetMailCount`、`GetYearCountsForFolder` 等）廣泛使用了 `ConcurrentDictionary`，但邏輯多為：
  ```vb
  If _cache.TryGetValue(key, val) Then Return val
  ' ... 執行耗時的 DB 或 Layer3 查詢 ...
  _cache(key) = newVal
  ```
  在多執行緒環境下（如 Parallel 迴圈），多條執行緒可能「同時」發生 Cache Miss，導致耗時查詢被重複執行。
  **計畫**：應重構快取代理層，使用 `_cache.GetOrAdd(key, Function(k) ...)`，讓內部邏輯由底層確保執行緒安全與單次建置。

### 3. 如果資料夾非常多（>2000），考慮把整個 Store 的 folder hierarchy 一次建好快取，而不是每次只建 subtree
- **狀態**：❌ **未實作**
- **檢查結果**：目前的架構是以 `GetSubtreeToList` 為核心，針對特定的 `rootPath` 向 SQLite 發送 `LIKE` 查詢。
- **說明**：這代表程式依賴「按需加載」或「區塊加載」。如果 User 選了一個 Store 根節點，它會遞迴或透過 DB 取出該棵樹。但在程式最一開始啟動時，並沒有「一次性」將整顆幾千個節點的樹在記憶體中建立 mapping。
  **計畫**：可以在啟動階段或背景執行緒，下達 `SELECT * FROM folder_cache`，並在記憶體內直接建構完整 `Dictionary(Of String, Node)` 以徹底消除展開過程中的 DB 延遲。

### 4. `Dim queue As New Queue(Of Outlook.Folder)(512)` (預分配容量，減少 Resize)
- **狀態**：❌ **未實作**
- **檢查結果**：專案中的 `Queue`（例如在 `Form1_MainTabs.vb` 行 663、`Form1_Outlook.vb` 行 1590）皆未指定初始容量。
- **說明**：`Queue(Of T)` 在加入元素超過內部陣列大小時，會以 O(N) 的成本重新配置兩倍大小的陣列。若已知可能處理大量資料夾，給定如 `(512)` 這樣的預估容量能省下不必要的記憶體搬移。
  **計畫**：逐一將所有 BFS 用到的 `Queue` 與 `List` 宣告加上適當的 Capacity 參數。

### 5. 讓下次程式啟動時完全不跑 BFS，直接從 SSD 讀 EntryID 清單 → GetFolderFromID 重建
- **狀態**：✅ **已實作**
- **檢查結果**：在 `Form1_Outlook.vb` 的 `GetSubtreeToList` 與 `GetSortedSubFolders` 內，皆已實作了 DB Lazy Load。
- **說明**：程式會呼叫 `DbGetSubFolderIDList(rootPath, ...)` 取得 DB 中的 `eid` 與 `sid`，接著直接呼叫 `_olNS.GetFolderFromID(row.eid, row.sid)` 重建 COM 物件。這把耗時降到了只需 0.1~0.3 秒，成功繞過了 MAPI 的爬樹邏輯。

### 6. 【優化】把 Folders 集合一次取出，避免每次 For Each 都重取
- **狀態**：⚠️ **部分實作**
- **檢查結果**：在 `Form1_Outlook.vb` 已經有部分優化：`Dim subs As Outlook.Folders = pFolder.Folders`，並在 Finally 中 `TryMarshalRelease(subs)`。但其他很多地方依然直接寫 `For Each f As Outlook.Folder In rootFolder.Folders`。
- **說明**：在 COM 的世界中，每次 `.` 存取集合屬性都可能產生新的 RCW (Runtime Callable Wrapper) 物件。
  **計畫**：應全面巡檢 `*.vb`，統一改成先指派給區域變數再迭代，並確保正確釋放 RCW。

### 7. 將 GetSortedSubFolders(curr.folderObj, fPath) 取出存入變數，不要放在 For Each 之中
- **狀態**：⚠️ **有優化空間**
- **檢查結果**：在 `Form1_MainTabs.vb` 行 709 中有寫法：`For Each subFolder In GetSortedSubFolders(curr.folderObj, fPath)`。
- **說明**：其實在 VB.NET 中，`For Each` 結構的來源集合運算式**只會被求值一次**，所以效能上並不會每次迴圈都重算。但為了程式的可讀性、方便 Debug 時查看集合 Count，且避免在某些條件下可能引發的不可預期行為，最佳實踐仍是將其抽出。
  **計畫**：將這類 `For Each` 前先宣告 `Dim sortedFolders = GetSortedSubFolders(...)`。

### 8. 避免在展開子資料夾的迴圈內，呼叫 `Dim fPath As String = subF.FolderPath`
- **狀態**：✅ **已大幅實作**
- **檢查結果**：搜尋結果中滿滿的都是 `2026/04/16 by Gemini: 升級回傳 Tuple (Folder, fPath)，消除呼叫端對 COM .FolderPath 的依賴` 的註解。
- **說明**：在遞迴與迴圈中，您已經改為透過父節點路徑加上字串拼接 (`fPath & "\" & sName`) 往下傳遞，或者將路徑直接封裝進 `Tuple` 或 `MailItemInfo` 裡，有效避開了昂貴的 MAPI 往上爬樹（`.FolderPath`）開銷。

---

## 接下來的實作計畫 (Implementation Plan)

如果您同意，我將開始執行以下具體修改：

### 階段一：快速且安全的替換（針對第 4, 6, 7 項）
- **[修改]** 各處 BFS 迴圈內的 `Queue` 初始化，加入 `(512)` 容量。
- **[修改]** 將 `Form1_MainTabs.vb` 等處在 `For Each` 中呼叫的方法抽成獨立變數。
- **[修改]** 將所有直接呼叫 `For Each folder In root.Folders` 的寫法，改為先 `Dim subs = root.Folders`，並加上 `Finally TryMarshalRelease(subs)`。

### 階段二：快取層執行緒安全升級（針對第 2 項）
- **[修改]** 針對 `Form1_Outlook.vb` 內的 Layer 2.5 快取代理層（如 `GetMailCount`、`GetFolderCount`），導入 `GetOrAdd` 模式（或 `Lazy(Of T)` 等等價機制），確保平行查詢時不會出現重複撈取的情形。

> [!IMPORTANT]
> **開放問題與確認：**
> 1. 關於第 3 項「整棵樹一次建置」，這牽涉到整個架構從 Lazy Load 變為 Eager Load，風險較大。您希望在此次修改中一併將 SQLite 啟動時的初始化改為一次載入整個 Store 嗎？還是保留現有 Subtree 點選才載入的方式就好？
> 2. `GetOrAdd` 模式下如果處理非同步 (Async/Await) 方法，通常需要包裝成 `ConcurrentDictionary(Of Key, Lazy(Of Task(Of T)))` 才能真正避免重複執行。請問是否同意引進此種寫法？

等待您的審閱與指示！
