# ComputeFolderStatsAsync 重構計畫

在開始動手切分之前，我仔細評估了這 140 行程式碼拆分成小函數的潛在影響與副作用。這是在重構這類核心邏輯時**絕對必要**的思考過程。

## 分析結果：需要注意的影響與副作用

1. **陣列索引的敏感性 (Index Sensitivity)**
   - **潛在地雷**：目前的 BFS 邏輯非常依賴 `allEntries` 這個 List 中的**索引值** (`ParentIndex`)。
   - **解法**：我們傳遞 `allEntries` (ByVal List) 給後續的小函數時，後續的函數**絕對不能**改變 List 的長度或顛倒元素順序。只要我們確保 `allEntries` 只是用來被「循序讀寫屬性」，這種拆分就是絕對安全的。

2. **記憶體參照 (Reference Type)**
   - 我們不需要每個函數都 `Return allEntries`。因為 `List(Of T)` 傳遞的其實是記憶體的參考，我們在 `FetchFolderMailCountsAsync` 和 `AggregateFolderStatsBottomUp` 中修改 `entry.TotalMailCount`，主函數的 `allEntries` 也會同步反映更改。這是合理且正確的設計，能減少記憶體複製的開銷。

3. **非同步的傳染性 (Async Contagion)**
   - **潛在地雷**：目前的 `Step 2 (L3 讀取)` 中有 `Await Task.Yield()`，這是為了不卡死 UI。
   - **解法**：拆分出來的 `Step 2` 子函數必須標記為 `Async Function ... As Task(Of Boolean)`，而主函數也必須去 `Await` 它。

4. **ESC 中斷處理的轉移**
   - **影響**：全域的 `_cancelRequested` 主要在 Step 2 被觸發。
   - **解法**：`Step 2` 子函數遇到取消時回傳 `True` (表示 isCancelled)。主函數收到 `True` 後直接回傳空的 List。

---

## 擬議變更 (Proposed Changes)

基於上述分析，這是最安全、乾淨且符合您的命名風格的切分計畫。這些函數會接續放在原來的函數下方。

### [MODIFY] [Form1_Main.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_Main.vb)

**1. 主函數精簡 (將原本的 140 行縮減至不到 20 行)**
```vb
Private Async Function ComputeFolderStatsAsync(rootFolder As Outlook.Folder, progress As IProgress(Of L3ProgressReport)) As Task(Of List(Of FolderBfsEntry))
    Dbg("開始", rootFolder.Name)
    
    ' Step 1: 負責展開樹狀結構與初步快取剪枝
    Dim allEntries As List(Of FolderBfsEntry) = BuildFolderBfsTree(rootFolder)
    
    ' Step 2: 負責與 COM 溝通，取得基本數據 (回傳是否被使用者取消)
    Dim isCancelled As Boolean = Await FetchFolderMailCountsAsync(allEntries, progress)
    If isCancelled Then Return New List(Of FolderBfsEntry)()
    
    ' Step 3 & 4: 純記憶體運算與快取更新
    AggregateFolderStatsBottomUp(allEntries)
    UpdateFolderStatsCache(allEntries)
    
    ' Step 5: 提取 UI 所需的結果並回報最終進度
    Return ExtractRootAndSubFolders(allEntries, progress)
End Function
```

**2. 新增的五個小函數 (將使用 Chunked Edits 分批寫入)**
- **`BuildFolderBfsTree`**: 處理 BFS Queue 遍歷。
- **`FetchFolderMailCountsAsync`**: 執行迴圈調用 `GetCachedMailCount`、處理 `progress` 報告與 `Task.Yield()`。
- **`AggregateFolderStatsBottomUp`**: 反向 For 迭代加總 `TotalMailCount`。
- **`UpdateFolderStatsCache`**: 把結果 TryAdd 進 Cache 字典。
- **`ExtractRootAndSubFolders`**: 將 `ParentIndex = 0` 的目標萃取出來。

---

## 使用者回饋需求 (User Feedback Required)

> [!IMPORTANT]
> 經過上述詳細推演，我認為這個切分**沒有破壞性副作用**，且利用傳址 (Reference passing) 修改 `allEntries` 屬性是非常輕量的做法。
> 
> 請您檢視這份計畫。如果您覺得這些函數的命名與切分邏輯 OK，請直接回覆 **「同意」** 或給予我修改方向，我就會立即以**小區塊 (Chunk)** 的方式，一步一步拆解並更新您的程式碼！
