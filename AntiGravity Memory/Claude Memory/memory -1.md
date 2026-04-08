# Outlook Assistant — 開發記憶包 v2026-03-14

## 專案基本資訊
- **語言**：VB.NET WinForms
- **Outlook**：LTSC 16.0.19725 (64位元)、本機 PST、`olNotExchange`
- **主檔案**：Form1.vb (~2524行)、DebugForm.vb、moduleStore.vb（存檔用）、MultiSelectTreeView.vb
- **分頁**：Tab1資料夾統計、Tab2依日期統計、Tab3尋找附件、Tab4系列郵件(未開發)、Tab5重複郵件(Levenshtein實作但UI未完)

## 重要變數命名
```
AppOutlook        = Outlook.Application
objNameSpace      = Outlook.NameSpace
blnButton3_Stop   = Boolean 中止旗標 (Tab3)
sw5               = Stopwatch (Tab3計時)
```

## Cache 設計（重要決策）
```vb
' ConcurrentDictionary — key 仍用 Outlook.Folder 物件（使用者決定，接受 cache miss 風險）
Dim folderCountCache As New ConcurrentDictionary(Of Outlook.Folder, Integer)
Dim mailCountCache   As New ConcurrentDictionary(Of Outlook.Folder, Integer)
' 單執行緒，key 也是 Outlook.Folder
Dim folderTreeCache  As Dictionary(Of Outlook.Folder, TreeNode)
Dim folderSizeCache  As Dictionary(Of Outlook.Folder, Long)
' Tab2 — key = FolderPath 字串
Dim yearCountsCache  As Dictionary(Of String, Dictionary(Of Integer, Integer))
Dim monthCountsCache As Dictionary(Of String, Dictionary(Of Integer, Integer))
```

---

## 已完成改動（勿重複）

### Threading / COM
- `folderCountCache`/`mailCountCache` → `ConcurrentDictionary`，`.Add()` → `.TryAdd()`
- `GetInfoForListview` → `Async Function`，移除 `Task.Run` 包 COM，移除 `.Result`
- `TreeView1_AfterSelect` → `Async Sub + Await GetInfoForListview()`
- `CountMailByYearAsync_CL2` → 移除 `Task.Run` wrapper，每10年 `Task.Yield()`
- `GetYearCountsAsync_CL` → `UpdateCounterProgress` 改直接呼叫（移除 `Await Task.Run`）
- ContextMenu → 3個 ContextMenuStrip 改 Form 層級成員變數，初始化一次

### UI
- TreeView/ListView Hover → 共用 Handler，`Color.FromArgb(240,240,240)`
- 防閃爍 → Win32 雙緩衝 `TVM_SETEXTENDEDSTYLE` / `LVM_SETEXTENDEDLISTVIEWSTYLE`
- Chart2 平均線 → 獨立 `Series("平均線")`，`ChartDashStyle.Dash`，`TextAnnotation` 右端標籤
- Chart2 切換空資料夾 → `ShowResult()` 先清舊 Series/Annotation
- DebugForm → BeginUpdate/EndUpdate 防閃爍；雙擊複製單行、Ctrl+C 複製多行(Tab分隔)

---

## Pending 清單（還沒動）

| 優先 | 位置 | 問題 |
|------|------|------|
| 高 | 行 862 | `s4Task = Task.Run(GetMailCountByMAPINew)` + `.Result` 死結 |
| 高 | 行 1209 | `Invoke(Sub() lblStatus1.Text = GetMailCountByMAPINew(...))` COM在Invoke裡 |
| 高 | 行 1664 | `Await Task.Run(Sub() UpdateCounterProgress(...))` 第二個函數未修 |
| 中 | ListView排序 | `CompareNumbers` 仍用 Async+Task.Yield+.Result，應改背景整批排序 |
| 中 | `GetFolderSizeLINQ` | 改用 `GetTable + Marshal.ReleaseComObject` |
| 低 | 多處 | `Marshal.ReleaseComObject` 未釋放附件/Items物件 |

### LINQ 優化（已識別未實作）
- `keyword.ToLower` 在 FilterAttach 迴圈內重複計算 → 拉到迴圈外
- `findName.Replace(" - ", "")` 在 FindNodeByName 等迴圈內重複計算 → 拉到迴圈外
- `ShowResultToListView2` 不必要 Invoke + 缺 BeginUpdate/EndUpdate

---

## 使用者偏好（重要）
- ❌ 不移除 `Parallel.ForEach`（保留效能意圖，ConcurrentDictionary 降風險）
- ❌ cache key 不改成 FolderPath 字串（接受偶爾 cache miss）
- ✅ 偏好直接給程式碼，少解釋
- ✅ DebugForm clipboard 是主要除錯工具（雙擊=單行，Ctrl+C=多行Tab分隔）
- ✅ 所有 Catch 要寫 `System.Exception`（因 Imports Outlook 導致命名空間衝突）

---

## Tab3 探索結論（死路已確認，勿再嘗試）

**問題**：5310封逐一讀附件 → 154秒（29ms/封，COM正常但量大）

| 路線 | 方法 | 結果 |
|------|------|------|
| A | Restrict + attachmentfilename LIKE | ❌ 全回傳0封，PST不支援附件屬性索引 |
| A2 | AdvancedSearch | ❌ HResult=0x8007064F，LTSC無Windows Search索引 |
| B1 | SetColumns(PR_ATTACH_LONG_FILENAME) | ❌ 導致Attachments=null，附件是子物件不適用 |

**Tab3 下一步（之後再做）**：
1. 移除 `Parallel.ForEach` 改 `GetFirst/GetNext`，`keyword.ToLower` 拉迴圈外，Invoke改Mod 50
2. `ShowResultToListview3Async` 改用 `GetTable` 讀 Subject/Size/ReceivedTime/SenderName/EntryID

---

## 本次新Chat目標：Tab1 優化

### Tab1 現有架構
```
LoadStoreToTreeView()        → 載入根節點（PST Store 層級）
LoadSubFolderToTreeView()    → BeforeExpand 時 lazy load 子資料夾
TreeView1_AfterSelect()      → Async，呼叫 GetInfoForListview()
GetInfoForListview()         → 讀資料夾統計，顯示到 ListView1
GetTotalFolderCountAsync()   → Parallel.ForEach 遞迴計算子資料夾數
GetMailCountByMAPINew()      → 計算郵件數（含子資料夾選項）
GetFolderSizeLINQ()          → 計算資料夾大小（子資料夾遞迴，但用舊 Task.Run 方式）
```

### Tab1 已知問題
- `GetTotalFolderCountAsync` 用 `Parallel.ForEach` 操作 COM 物件（STA 違規）
- `GetFolderSizeLINQ` 用 `Task.Run(Function() folder.Items.Cast(Of Object)().Sum(...))` → 應改 `GetTable`
- BFS 資料夾遍歷模式已在 Tab3 設計好，建議套用到 Tab1/Tab2
- Tab1 cache key 維持 COM 物件（已決定不改）

### BFS 模式（Tab3已設計，Tab1可套用）
```vb
Private Function GetTargetFolders(rootFolder As Outlook.Folder, includeSubFolders As Boolean) As List(Of Outlook.Folder)
    Dim result As New List(Of Outlook.Folder)
    Dim queue As New Queue(Of Outlook.Folder)
    queue.Enqueue(rootFolder)
    While queue.Count > 0
        Dim f = queue.Dequeue()
        result.Add(f)
        If includeSubFolders Then
            For Each sub As Outlook.Folder In f.Folders
                queue.Enqueue(sub)
            Next
        End If
    End While
    Return result
End Function
```

### ListView 背景排序模式（已討論，適用Tab1/Tab3）
```vb
' ColumnClick 事件裡
Dim items = ListView1.Items.Cast(Of ListViewItem)().ToList()
Dim sortedItems = Await Task.Run(Function()
    ' 數字欄位用 Long.TryParse，文字欄位用 StringComparer.CurrentCultureIgnoreCase
    Return items.OrderBy(Function(item) ...).ToList()
End Function)
ListView1.BeginUpdate()
ListView1.Items.Clear()
ListView1.Items.AddRange(sortedItems.ToArray())
ListView1.EndUpdate()
```
