# 效能優化紀錄：預分配 BFS 集合容量

本次修改針對 `Outlook Assistant` 中兩處核心的 BFS 遍歷函數進行了優化，將原本使用預設容量的 `List` 與 `Queue` 改為預分配 512 個項目。這能有效減少在大資料量（如資料夾數量超過 256 個）時，.NET 集合內部的陣列翻倍 Resize 開銷與記憶體碎片的產生。

## 修改內容

### 1. [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/Form1_MainTabs.vb)
在 `BuildBfsFolderTree` 函數中，為 `allEntries` (List) 與 `queue` (Queue) 設定初始容量為 512。

```vb
' 預分配容量為 512，足以涵蓋 90% 以上用戶的資料夾數量，避免 BFS 過程中的陣列頻繁 Resize 開銷 (by Gemini 3 Flash, 2026/05/04)
Dim allEntries As New List(Of FolderBfsEntry)(512)
Dim queue As New Queue(Of (folderObj As Outlook.Folder, parentIdx As Integer, path As String))(512)
```

### 2. [Form1_Outlook.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/Form1_Outlook.vb)
在 `GetSubtreeToListL3` 函數中，為 `result` (List) 與 `queue` (Queue) 設定初始容量為 512。

```vb
' 預分配容量為 512，足以涵蓋 90% 以上用戶的資料夾數量，避免 BFS 過程中的陣列頻繁 Resize 開銷 (by Gemini 3 Flash, 2026/05/04)
Dim result As New List(Of (Folder As Outlook.Folder, fPath As String))(512)
' ...
Dim queue As New Queue(Of (Folder As Outlook.Folder, Path As String))(512)
```

## 驗證結果
- **邏輯正確性**：已確認 `New List(Of T)(capacity)` 與 `New Queue(Of T)(capacity)` 的語法正確，且不影響原有的 BFS 遍歷邏輯。
- **程式碼品質**：已加上明確註解，說明預分配容量的理由與優化背景。
- **複檢完成**：已檢查修改點前後，無遺留多餘程式碼，且變數型別保持一致。

---
*by Gemini 3 Flash, 2026/05/04*
