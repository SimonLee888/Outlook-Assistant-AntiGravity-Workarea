# Outlook Assistant 開發記憶摘要
## 供新 Chat 接手使用 — 產出於 2026-03-14

---

## 專案基本資訊

- **語言/框架**：VB.NET / WinForms
- **環境**：Outlook LTSC 2021（olNotExchange）、本機 PST 檔、VS 2022
- **主要檔案**：Form1.vb（~3000行）、DebugForm.vb、moduleStore.vb、MultiSelectTreeView.vb
- **5個分頁**：Tab1 資料夾統計、Tab2 依日期統計、Tab3 尋找附件、Tab4 系列郵件（未開發）、Tab5 重複郵件（未完成）
- **重要全域物件**：`AppOutlook`、`objNameSpace`、`PstStoreList`、`TreeView1/2/3`、`ListView1/2/3`

---

## 已確認的全域架構原則

### COM / 執行緒規則（非常重要）
- **所有 Outlook COM 呼叫必須在 UI 執行緒（STA）**
- **不可用 Task.Run / Parallel.ForEach 包住 COM 物件**
- **用 `Await Task.Yield()` 讓 UI 不凍結，不是切換執行緒**
- `Me.Invoke()` 只在從背景執行緒更新 UI 時才需要；在 UI 執行緒的 Async Sub/Function 裡直接更新即可

### 資料夾遍歷最佳實踐（Tab3 已實作，Tab1/Tab2 待套用）
```vb
' BFS 遍歷，把資料夾收集與業務邏輯完全分離
Private Function GetTargetFolders(rootFolder As Outlook.Folder,
                                   includeSubFolders As Boolean) As List(Of Outlook.Folder)
    Dim result As New List(Of Outlook.Folder)
    result.Add(rootFolder)
    If Not includeSubFolders Then Return result
    Dim queue As New Queue(Of Outlook.Folder)
    queue.Enqueue(rootFolder)
    While queue.Count > 0
        Dim current As Outlook.Folder = queue.Dequeue()
        For Each subFolder As Outlook.Folder In current.Folders
            result.Add(subFolder) : queue.Enqueue(subFolder)
        Next
    End While
    Return result
End Function
```

### DASL / Restrict 的根本限制（已確認，勿再測試）
- 附件檔名（`PR_ATTACH_LONG_FILENAME` 等）是 Attachment 子物件，PST 無索引，任何 DASL 語法都**無法**直接篩附件檔名
- `AdvancedSearch` 在 LTSC 失敗（`HResult=0x8007064F`）= Windows Search 未索引 PST，非語法問題
- **PST 可用的 DASL 頂層屬性**：`urn:schemas:httpmail:hasattachment`、`PR_MESSAGE_SIZE 0x0E080003`

---

## Tab3 尋找附件（已完成重構）

### 架構：兩階段搜尋
```
Button8_Click（新架構，主要使用）
├── BuildDaslFilter()           → 純字串，hasattachment + 大小範圍
├── GetTargetFolders()          → BFS 資料夾收集
├── ScanFolderWithGetTable()    → Phase1：GetTable MAPI binary table，快 5~10 倍
│    └── 產出 List(Of AttachCandidateInfo)（純 .NET struct，不帶 COM）
├── FilterByAttachDetailAsync() → Phase2：GetItemFromID 逐一細查附件名稱/數量
│    └── 只在有 keyword 或 count filter 時執行
├── BuildListItemsFromCandidates() → 無需細查時直接建 ListViewItem
└── DisplayTab3Results()        → BeginUpdate / AddRange / EndUpdate

Button3_Click（舊架構，保留作效能比對）
└── Items.Restrict() → GetFirst/GetNext → ShowResultToListview3Async
```

### AttachCandidateInfo Structure
```vb
Private Structure AttachCandidateInfo
    Dim EntryID As String
    Dim Subject As String
    Dim Size As Long
    Dim ReceivedTime As DateTime
    Dim SenderName As String
    ' 待加入（快取方案實作後）：
    ' Dim AttachFileNames As List(Of String)
End Structure
```

### 效能比較（Inbox 9840封 / 有附件5310封）
| 情境 | Button3 | Button8 |
|---|---|---|
| 不篩名稱不篩大小 | 基準 | 快約2倍（GetTable優勢）|
| 有篩名稱且有篩大小 | 接近 | 接近平手 |
| **有篩名稱不篩大小** | **反而較快** | 較慢 |
- Button8 有篩名稱時需要 GetItemFromID（隨機存取），Button3 用 GetFirst/GetNext（順序迭代），PST B-tree 對順序迭代更友善，所以特定情形 Button3 反而快

### 已知待修 Bug
```vb
' FilterByAttachDetailAsync 裡這行永遠是 True（bug）：
If processed Then   ' ← 應改為 If processed Mod 20 = 0 Then
```

### 附件檔名比對優化（待套用）
```vb
' 現在：att.FileName.ToLower.Contains(keyword) → 每次產生新字串
' 改為：
att.FileName.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0
```

### 快取方案（設計完成，待實作與前置驗證）
- **結構**：`Dictionary(Of String, FolderCacheEntry)`，key = `folder.FolderPath`
- **失效判斷**：`folder.Items.Count` 改變就清快取
- **前置驗證**：`PR_ATTACH_LONG_FILENAME` 能否加進 `GetTable.Columns`（PST 可能不支援，需實測）
- 若可行：快取 `AttachFileNames`，命中時完全跳過 Phase1 + Phase2，篩選純在 .NET 執行
```vb
Private Structure FolderCacheEntry
    Dim Candidates As List(Of AttachCandidateInfo)
    Dim ItemCountWhenCached As Integer
    Dim CachedAt As DateTime
End Structure
Private _tab3Cache As New Dictionary(Of String, FolderCacheEntry)
```

### UI 控制項名稱（Tab3）
- `TreeView3`：資料夾選取
- `CheckSubFolder3`：含子資料夾
- `CheckAttachName` + `TextBox3`：附件名稱關鍵字
- `CheckSize` + `NumberMin/Max` + `UnitMin/Max`：郵件大小範圍
- `CheckAttCount` + `CountMin/Max`：附件個數範圍
- `ListView3`：結果顯示（欄位順序：Subject, Size, ReceivedTime, SenderName, AttachCount, EntryID）
- `Button3`（舊）、`Button8`（新）、`Button3_Stop`
- `lblStatus1`（進度）、`lblStatus2`（總時間+速度）

---

## ListView 排序（已優化）

### 問題根源（已解決）
原本大小欄排序用 `Task.Run(Function() CompareNumbers(...)).Result`，68,000次比較 × 每次建立 Task = 10 秒以上。

### 現行正確架構
```vb
Public Class ListViewItemComparer
    Implements IComparer
    ' 全同步，無 Task.Run
    Public Function Compare(x As Object, y As Object) As Integer Implements IComparer.Compare
        Select Case columnIndex
            Case 1  ' 大小：Long.TryParse（或從 Tag 讀）
            Case 2  ' 日期：DateTime.TryParse
            Case 4  ' 附件數：Integer.TryParse
            Case Else  ' 文字：String.Compare OrdinalIgnoreCase
        End Select
    End Function
End Class
```
- `DefaultComparer` 在 ColumnClick 裡的多餘呼叫已移除
- 數千到一萬項目排序 < 0.1 秒

---

## Tab1 資料夾統計（下一個優化目標）

### 現有架構（問題清單）

**快取問題**：
- `folderCountCache`：`ConcurrentDictionary(Of Outlook.Folder, Integer)`，key 是 COM 物件（應改為 FolderPath 字串）
- `mailCountCache`：同上問題
- `folderSizeCache`：`Dictionary(Of Outlook.Folder, Long)`，非 Concurrent，key 也是 COM 物件
- `mailSizeCache`：同上

**`GetInfoForListview` 的問題**：
```vb
' 行 856：s4Task.Result 潛在 deadlock
Dim s4Task = Task.Run(Function() GetMailCountByMAPINew(folder))
' ...
Dim s3 As String = s4Task.Result.ToString(...)  ' ← .Result 阻塞 UI 執行緒
```
- `GetMailCountByMAPINew` 裡用 `Parallel.ForEach` 遍歷 COM Folder 子物件（STA 違規）
- `GetTotalFolderCountAsync` 裡也用 `Parallel.ForEach` 遍歷 COM Folder（STA 違規）

**`GetFolderSizeLINQ` 的問題**：
```vb
' Task.Run 包住 COM 物件（STA 違規）
Dim folderSize As Long = Await Task.Run(Function() folder.Items.Cast(Of Object)().Sum(Function(s) s.Size))
```
- 應改用 `GetTable + PR_MESSAGE_SIZE` 逐列加總，純在 UI 執行緒

**`GetInfoForListview` 架構問題**：
- 遞迴邏輯與顯示邏輯混在一起
- 資料夾遍歷沒有用 BFS 分層（應套用 GetTargetFolders 模式）
- 在 TreeView1_AfterSelect 裡直接對每個子資料夾呼叫（沒有收集再處理）

### Tab1 優化方向
1. 快取 key 全部改為 `folder.FolderPath`（字串，不帶 COM 物件）
2. `GetMailCountByMAPINew` 改為用 `GetTable + PR_CONTENT_COUNT` 在 UI 執行緒順序讀取（不用 Parallel）
3. `GetFolderSizeLINQ` 改為 `GetTable + PR_MESSAGE_SIZE` 加總（移除 Task.Run 包 COM）
4. 套用 `GetTargetFolders()` BFS 模式，把資料夾收集與業務邏輯分離
5. 解決 `s4Task.Result` deadlock（改為 Await）

### Tab1 UI 控制項名稱
- `TreeView1`：資料夾選取
- `ListView1`：結果顯示（欄位：資料夾名稱, 郵件數, 子資料夾數, 含子郵件總數, 大小）
- `_ctxListView1`：右鍵選單（已改為只建立一次，避免 memory leak）

---

## Tab2 依日期統計（已完成，待套用 BFS 優化）

- L1→L2（`ComputeYearCountsAsync`）→L3 分層架構完成
- `yearCountsCache` / `monthCountsCache` key 已是 FolderPath 字串（正確）
- SimTreeView2 多選控制項完成
- **待優化**：資料夾遍歷套用 GetTargetFolders BFS 模式（低優先）

---

## 其他全域待修項目

| 優先 | 位置 | 問題 |
|---|---|---|
| 高 | 行 862 `s4Task.Result` | 潛在 deadlock |
| 高 | `FilterByAttachDetailAsync` `If processed Then` | bug，應為 `Mod 20` |
| 高 | `GetMailCountByMAPINew` Parallel.ForEach | COM STA 違規 |
| 中 | `GetFolderSizeLINQ` Task.Run 包 COM | STA 違規，改 GetTable |
| 中 | Tab3 快取方案 | 待前置驗證後實作 |
| 低 | Tab1/Tab2 套用 GetTargetFolders BFS | 架構優化 |
| 低 | Tab4 系列郵件 | 未開發 |
| 低 | Tab5 重複郵件 | UI 未完成 |
