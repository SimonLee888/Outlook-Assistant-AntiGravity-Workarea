# Outlook Assistant — 開發記憶摘要
# 用途：新 Chat 啟動時附上此檔，讓 Claude 快速恢復上下文
# 最後更新：2026-03-13（Form1.vb 3207行 / SimTree.vb 474行）

---

## 專案環境
- VB.NET / WinForms，Visual Studio 2022
- Outlook LTSC 2021（olNotExchange），本地 PST 檔案
- 主要檔案：Form1.vb、SimTree.vb、DebugForm.vb、moduleStore.vb

---

## 整體架構：5個Tab

| Tab | 功能 | 狀態 |
|-----|------|------|
| Tab1 資料夾統計 | TreeView1 選資料夾 → ListView1 顯示子資料夾郵件數/大小 | **待重構**（下一個Chat的主題） |
| Tab2 依日期統計 | 選資料夾 → 年度/月份分布圖表 | **已完成**，L1→L2→L3分層架構 |
| Tab3 尋找附件 | DASL篩選 + 兩階段掃描 | 基本完成，部分細節待補 |
| Tab4 系列郵件 | Levenshtein相似度 | UI未完成 |
| Tab5 重複郵件 | | 未開發 |

---

## Tab2 架構（已完成，可作為Tab1重構範本）

```
L1（UI事件層）  : TreeView2_AfterSelect / SimTree2_AfterSelect
                   只取節點、組 folderList、呼叫L2、顯示結果
L2（流程協調層）: ComputeYearCountsAsync(folderList, totalMailCount, onProgress)
                   BFS遍歷、管理快取、驅動L3、合併結果、callback回報進度
L3（COM資料層） : CountYearsInSingleFolderAsync(folder)
                   只碰COM，回傳單一資料夾年份分布 ConcurrentDictionary(Of Integer,Integer)
```

**關鍵設計決策：**
- `yearCountsCache` / `monthCountsCache`：key = `folder.FolderPath`（純字串，避免COM物件當key的RCW殘留）
- `folderItems` 在L3迴圈外取一次，`Finally` 統一 `Marshal.ReleaseComObject`
- 進度用 callback `onProgress(processed, total)` 傳回L1更新UI，L2不碰UI控制項
- `Await Task.Yield()` 每個資料夾處理完讓出一次，保持UI回應
- `StatusUpdate(yearCounts, elapsed As TimeSpan)` 顯示速度統計
- `_tab2FolderList` / `_tab2IsMonthView` 記住目前視圖狀態供月份展開使用

**Tab2 月份視圖：**
- `ShowMonthViewAsync(year)` / `ShowYearViewAsync()`
- `UpdateChart2ForMonths` X軸要清除 `IntervalOffset=0` 和 `LabelStyle.Format=""`，避免年度格式殘留

---

## SimTree 控制項（已完成）

**檔案：SimTree.vb（474行）**

**核心設計：**
- `Shadows SelectedNode`：Get=`_lastClickedNode`，Set=`SelectSingleNode`（不觸發AfterSelect）
- `SelectedNodes As List(Of TreeNode)`：唯讀，多選清單
- `MouseDown` 記錄 `_pendingMouseUpNode`，`MouseUp` 才執行選取 + `FireAfterSelect`
- `OnBeforeSelect` 永遠 `e.Cancel = True`（阻止原生選取）
- `OnAfterSelect` 不呼叫基類（由 `FireAfterSelect` 手動觸發）
- Ctrl = 切換單節點；Shift = `SelectRange`（NextVisibleNode遍歷，跨層級正確）；普通 = 單選
- Space鍵：`e.SuppressKeyPress = True`（不只是 `Handled = True`，否則KeyPress會執行兩次）

**失焦/得焦：**
- 失焦：`SystemColors.InactiveCaption` / `InactiveCaptionText`
- 得焦：`SystemColors.Highlight` / `HighlightText`
- 沒有選取時 `OnGotFocus` 自動選 `TopNode`

**公共方法（供Form1呼叫）：**
- `AddSelectedNode(node)` — 不清除其他選取，不觸發AfterSelect（ExpandTreeToDefaultInbox用）
- `ClearSelectedNodes()` — 清除所有選取
- `SetSelectedNode(node)` — 等同 SelectedNode = node

---

## Form1 中 SimTree2 相關

**成員變數：**
```vb
Private _suppressSimTreeAfterSelect As Boolean = False
Private _ctxTreeView2 As ContextMenuStrip   ' Form1_Load 預建，不在MouseClick每次new
Private _ctxSimTree2 As ContextMenuStrip    ' 同上
```

**TabControl_SelectedIndexChanged "依日期統計"：**
- `LoadStoreToTreeView` 期間設 `_suppressSimTreeAfterSelect = True`
- `ExpandTreeToDefaultInbox` 必須在 `SimTree2.Focus()` **之後**才呼叫（隱藏狀態BeforeExpand不觸發）
- `ExpandTreeToDefaultInbox` 迴圈用 `treeview.Nodes(0).Nodes.Count - 1`（不是 `treeview.Nodes.Count - 1`）

**HandleTreeViewMouseMoveEvent（已修正）：**
- 還原hover時，若是SimTree已選取節點，根據 `sim.Focused` 還原成 Highlight 或 InactiveCaption
- 注意：還原色用 `SystemColors.InactiveCaption`，不是 `Color.FromArgb(240,240,240)`
- 判斷型別用 `TypeOf treeView Is SimTree`（小寫 treeView，不是大寫 TreeView 類別名）

**SimTree2_AfterSelect（待修正）：**
```vb
' ❌ 現在的寫法（f.Items 產生 RCW 未釋放，GC延遲0.01~0.02s）：
rootFolders.Sum(Function(f) f.Items.Count)

' ✅ 應改成：
For Each rf As Outlook.Folder In rootFolders
    Dim rfItems As Outlook.Items = rf.Items
    totalMailCount += rfItems.Count
    Marshal.ReleaseComObject(rfItems)
Next
```

---

## Tab1 現況（下一個Chat的主題）

**現有函數：**
- `TreeView1_AfterSelect` — 混合了COM呼叫、排序、UI更新，未分層
- `GetInfoForListview(folder, iamSub)` — 每個子資料夾都獨立呼叫，含 `s4Task.Result`（⚠️ 潛在deadlock）
- `GetTotalFolderCountAsync` — 用 `Parallel.ForEach` + 遞迴，COM物件當cache key（`folderCountCache As ConcurrentDictionary(Of Outlook.Folder, Integer)`）
- `GetMailCountByMAPINew` — 同樣用 `Parallel.ForEach` + 遞迴 + COM物件當cache key
- `GetFolderListByTierAsync` — BFS函數已存在但目前沒被Tab1使用

**已知問題：**
1. `folderCountCache` / `mailCountCache` / `folderSizeCache` key 是 `Outlook.Folder` COM物件 → RCW殘留
2. `s4Task.Result` 在 line 1134 — 在async context裡用 `.Result` 可能deadlock
3. `Parallel.ForEach` 呼叫COM（STA違規，偶發crash）
4. `GetInfoForListview` 每個資料夾sequential Await，無法並行

**Tab1重構方向（待確認後實作）：**
- L1：`TreeView1_AfterSelect` 只取節點、呼叫L2、把結果批次丟給ListView
- L2：`ComputeFolderStatsAsync(folder)` — BFS展開子資料夾清單，驅動L3，管理快取
- L3：`GetFolderStatsAsync(folder)` — 只碰COM，回傳單一資料夾的{郵件數,子資料夾數,大小}
- Cache key 改成 `folder.FolderPath`（純字串）
- 移除 `Parallel.ForEach`，改成 sequential + `Await Task.Yield()`

---

## 共用函數

```vb
' BFS展開（已存在，Tab2/3使用）
GetTargetFolders(rootFolder, includeSubFolders) As List(Of Outlook.Folder)

' 已存在但Tab1未使用的BFS函數
GetFolderListByTierAsync(targetFolder, maxDepth?) As List(Of Outlook.Folder)

' 排序子資料夾
GetSortedSubFolders(folder) As List(Of Outlook.Folder)

' 懶載入TreeView
LoadStoreToTreeView(storeList, treeview)     ' 只載入一層，子資料夾用":::"佔位
LoadSubFolderToTreeView(sender, e)           ' BeforeExpand時動態載入下一層
ExpandTreeToDefaultInbox(treeview)           ' 展開並選取"收件匣"/"Inbox"
```

---

## 待清理的舊程式碼

- `GetTotalFolderCount_Old` / `GetMailCountByMAPI_Old`（已註解）
- `ExpandTreeToDefaultInbox` 舊版（已註解在line 1054）
- Tab2 的舊版函數（已在comment裡標記取代關係）
- `GetFolderSizeOld_Async`（line 1492，已有新版LINQ替代）

---

## 未解決的討論

**問題4（父子多選重複計數）：**
多選時若選了父子資料夾，父BFS展開已包含子，子再跑BFS就重複統計。
方案B（選父自動清掉已選的子節點）已初步討論，尚未實作。

**Tab1啟動速度優化（新Chat主題）：**
重構分層架構後，考慮：
- cache key改純字串（最優先）
- 移除Parallel.ForEach的STA違規風險
- 背景預讀（BackgroundWorker或低優先Task）
