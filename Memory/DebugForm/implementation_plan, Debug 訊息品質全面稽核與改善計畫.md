# Debug 訊息品質全面稽核與改善計畫

## 背景說明

本計畫針對 `Form1.vb`、`Form1_Main.vb`、`Form1_ComL3.vb`、`DebugForm.vb` 四個檔案中所有 `Dbg()` 呼叫進行全面品質稽核，分析以下五個面向：

1. **選擇策略**是否適當（有無輸出的必要性）
2. **語意一致性**（開始/結束/錯誤/狀態的措辭是否統一）
3. **格式一致性**（`msg` 與 `detail` 兩個參數的使用慣例）
4. **傳入內容的重要性**（有沒有傳入最重要的監控資訊）
5. **高頻迴圈的降噪策略**（避免灌爆 DebugForm）

---

## 稽核結果總覽（問題分類）

---

### 問題 A：格式不一致（語意混亂）

**現況**：`Dbg()` 的第一參數 `msg` 目前混用了以下不同風格，導致 DebugForm 中的配對搜尋（FindSimilarPair）失效或誤配：

| 現有寫法 | 建議統一用法 |
|---|---|
| `Dbg("開始", folder.Name)` | ✅ 已是標準格式 |
| `Dbg("結束", "序號已不匹配，跳過更新")` | ✅ 可接受 |
| `Dbg("Error: TreeView1_AfterSelect", ex.Message)` | ❌ `msg` 混入了函數名 |
| `Dbg("FilterFolderWithAttachment Error: ", folder.Name & "—" & ex.Message)` | ❌ 格式混亂，函數名在 msg 裡 |
| `Dbg("Button4 GetTable Error: " & folder.Name, ex.Message)` | ❌ 字串串接在第一參數 |
| `Dbg("Button3_Click Error: ", ex.Message & vbCrLf & ex.StackTrace)` | ❌ StackTrace 混入 detail |
| `Dbg("GetMailSize ⓪ RDO 失敗，走MAPI fallback", ex.Message)` | ⚠️ 可接受，但措辭偏長 |
| `Dbg("縮合側邊欄: " & sc.Name & "...")` | ❌ 把狀態描述和值混進 msg |
| `Dbg("", root.Name)` | ❌ msg 為空，意義不明 |
| `Dbg("發現預設收件匣", node.FullPath)` | ✅ 語意清晰，可接受 |

**建議統一規則**：
- `msg`：永遠是 `"開始"` / `"結束"` / `"錯誤"` / `"快取命中"` / `"被取消"` 等**固定關鍵字**
- `detail`：放入具體的值或說明（資料夾名、計數、耗時等）
- 錯誤格式改為：`Dbg("錯誤", ex.Message)` — 函數名由 DebugForm 的 `GetCallerName()` 自動填入，不需手寫

---

### 問題 B：函數有「開始」沒有「結束」

以下函數已有 `Dbg("開始")` 但**缺少 `Dbg("結束")`**，導致 DebugForm 的配對計時功能失效：

| 函數 | 所在檔案 | 問題說明 |
|---|---|---|
| `GetMailCount()` | Form1_ComL3.vb | 只有開始，三條 fallback 路徑各自有「成功/失敗」但沒有統一結束點 |
| `GetFolderCount()` | Form1_ComL3.vb | 同上，各 fallback 分散在 Debug 內 |
| `IsMailFolder()` | Form1_ComL3.vb | 有「開始」但在 `Return True/False` 前沒有結束 |
| `GetSizeMultiplier()` | Form1_Main.vb | 有「開始/結束」但**位置錯誤**（結束在 `Return` 之後，永遠到不了） |
| `HandleListViewKeyPress()` | Form1.vb | 有「開始」，各分支內沒有「結束」 |
| `HandleTreeViewKeyPress()` | Form1.vb | 有「開始」，無「結束」 |
| `Button3_Stop_Click()` | Form1_Main.vb | 有開始/結束，但功能極簡，可考慮不用 |
| `ListView4_SelectedIndexChanged()` | Form1_Main.vb | 開始/結束緊接，函數本體是空的 todo |
| `FindNodeOrItemByName()` | Form1_ComL3.vb | 有「開始」，無「結束」 |

---

### 問題 C：函數中途中斷跳出前缺少時間戳記

以下函數在**提早 Return 時**沒有輸出 debug 訊息，導致追蹤路徑不完整：

| 函數 | 行號區段 | 問題說明 |
|---|---|---|
| `TreeView1_AfterSelect()` | ~L231 | `Return` 前有 `Dbg("結束", "未選定資料夾")` ✅ 好 |
| `TreeView1_AfterSelect()` | ~L241 | `If _tab1SelectSeq <> mySeq Then Return` — **沒有 Dbg 輸出** ❌ |
| `TreeView1_AfterSelect()` | ~L244-246 | `_cancelRequested` 路徑只有 `ProgressBar1.Text` 沒有 `Dbg` ❌ |
| `SimTree2_AfterSelect()` | ~L638 | `selectedNodes.Count = 0 Then Return` — 無 Dbg ❌ |
| `SimTree2_AfterSelect()` | ~L644 | `targetFolderList.Count = 0 Then Return` — 無 Dbg ❌ |
| `Button3_Click()` | ~L1434 | `Return` 前有 `Dbg("結束", "未選擇資料夾")` ✅ 好 |
| `Button3_Click()` | ~L1471-1480 | Phase 1 被停止時有 `Dbg("結束", "Phase 1 被停止")` ✅ 好 |
| `Button4_Click()` | ~L1822-1828 | `Return` 前有 `Dbg("結束", "未選擇資料夾")` ✅ 好 |
| `Button5_Click()` | ~L1977-1981 | `Return` 前有 `Dbg("結束", "PST 尚未載入")` ✅ 好 |
| `OpenMailByEntryID()` | ~L1782 | `Return` 前有 `Dbg("結束", "EntryID 為空")` ✅ 好 |
| `GetSubFolderList()` | Cache hit 路徑 | 有完整 Dbg ✅ 好 |
| `ExpandTreeToDefaultInbox()` | 找不到收件匣時 | 無 Dbg 輸出，只 `Exit Sub` ❌ |

---

### 問題 D：高頻迴圈噪音過多

以下函數在**高頻迴圈內**有 `Dbg()` 呼叫，可能造成 DebugForm 被灌爆：

| 函數 | 問題行 | 建議策略 |
|---|---|---|
| `LoadSubFolderToTreeView()` | `Dbg("", selectedFolder.Name & folder.Name)` 在 For Each 迴圈內 | **每個資料夾都輸出一行** → 改成只在開始和結束輸出 `(子資料夾數: N)` |
| `LoadStoreToTreeView()` | `Dbg("", root.Name)` 在 For Each 迴圈內 | **每個 Store 都輸出** → 改成只在結束輸出總計 |
| `GetMailCount()` | 每次呼叫都輸出開始/成功/失敗 | **Tab1 BFS N 個資料夾就有 N×3 行** → 高頻情境改用 L2 層統計 |
| `GetFolderCount()` | 同上 | 同上，已有 L2.5 快取代理，快取命中後不再進入這裡 |
| `IsMailFolder()` | 每次判斷都輸出 | 應靜默或只在非郵件資料夾時輸出 |
| `FindSimilarPair()` 中的 `CalculateSimilarity()` / `LevenshteinDistance()` | 兩個函數各自有「開始/結束」，Tab5 每封郵件都呼叫 | **N 封 × 2 函數 = 2N 行** → 完全移除這兩個函數的 Dbg |

---

### 問題 E：傳入內容品質不足（detail 沒有傳入最重要的監控資訊）

| 函數 | 現有 detail | 建議改善內容 |
|---|---|---|
| `BuildBfsFolderTree()` | 節點總計 ✅ | 可補：快取命中 vs 非命中的節點數比例 |
| `FetchDirectMailCountsAsync()` | 無結束訊息 | **缺少結束** → 補 `Dbg("結束", $"共讀取 {total} 個節點")` |
| `SummarizeSubTreeBottomUp()` | 無開始/結束 | 純記憶體操作，可不加（實際上正確） |
| `UpdateFolderStatsCache()` | 無開始/結束 | 同上（純記憶體操作，正確） |
| `ComputeYearCounts()` | 結束時有 `merged.Count & merged.Values.Sum` ✅ | 可補已命中快取資料夾的數量 |
| `GetYearCountsForFolder()` | 有完整開始/結束 ✅ | 可補「找到年份數 vs BATCH 讀取次數」 |
| `ShowResultTab2()` | 開始傳入 `yearCounts.Values.Sum` ✅ | 可接受，但可改成傳入年份範圍更有意義 |
| `ShowProgressTab2()` | 有開始/結束 ✅ | 可考慮移除（此函數很簡單值得精簡） |
| `CheckSubFolder2_CheckedChanged()` | 結束無 detail | 補上節點數量或勾選狀態 |
| `ListView2_SelectedIndexChanged()` | 有開始/結束 ✅ | 但「結束」在第一行 Return 前沒有輸出 |
| `Chart2_MouseClick()` | 有開始/結束 ✅ | OK |
| `ListView3_MouseClick()` | 有開始/結束 ✅ | 單擊功能很簡單，可考慮移除 |
| `ShowResultTab3()` | 開始傳入 `items.Count` ✅ | 可補 `totalProcessed` 以計算命中率 |
| `ScanAttachmentDetail()` | 開始無 detail | 補入 `targetMailList.Count` |
| `TreeView4_AfterSelect()` | 開始傳入 `e.Node.Text` ✅ | 結束補入 `mailList.Count` |

---

## 具體改善計畫（分優先順序）

> **說明**：以下每個 Issue 後標示「建議優先度」：
> - 🔴 高優先（影響診斷能力）
> - 🟡 中優先（影響格式品質）
> - 🟢 低優先（可跳過）

---

### Issue 1 — 高頻迴圈去噪 🔴

**影響函數**：`LoadSubFolderToTreeView()`, `LoadStoreToTreeView()`, `GetMailCount()`, `GetFolderCount()`, `IsMailFolder()`, `CalculateSimilarity()`, `LevenshteinDistance()`

**改動內容**：

#### Form1_ComL3.vb `LoadSubFolderToTreeView()`（約 L346）
```vb
' 現在（每個資料夾一行，高噪音）:
Dbg("", selectedFolder.Name & folder.Name)

' 改成（只在結束輸出節點總計）:
' 完全移除迴圈內的 Dbg，只保留函數開始/結束
```

#### Form1_ComL3.vb `LoadStoreToTreeView()`（約 L324）
```vb
' 現在:
Dbg("", root.Name)  ' 每個 Store 一行

' 改成: 完全移除，結束時已有 Dbg("結束", tv.Name)
```

#### Form1_ComL3.vb `GetMailCount()`（約 L504）
```vb
' 現在: Dbg("開始", folder.Name) — 每次被L2.5呼叫都輸出
' 策略: 快取命中時不進入此函數，但仍然高頻。
' 建議: 只保留失敗路徑的 Dbg（成功路徑靜默）
' 或: 完全在此函數靜默，讓 L2.5 的 GetCachedMailCount 處理 miss 情況
```

#### Form1_ComL3.vb `IsMailFolder()`（約 L1108）
```vb
' 現在: 每次呼叫都 Dbg("開始")，快取命不中時再 Dbg
' 建議: 移除「開始」，只在非郵件資料夾時 Dbg 記錄（因為過濾是重要事件）
Dbg("過濾非郵件資料夾", $"{folder.Name} = {itemType}")
```

#### Form1_Main.vb `CalculateSimilarity()` 和 `LevenshteinDistance()` (約 L2123, L2134)
```vb
' 建議: 完全移除兩個函數的 Dbg，Tab5 掃描時呼叫量=N封郵件×比對次數
' CalculateSimilarity 的結束可保留但移到外層呼叫者（Tab5）
```

---

### Issue 2 — 補充缺少結束的函數 🔴

**影響函數**：`FetchDirectMailCountsAsync()`, `HandleListViewKeyPress()`, `HandleTreeViewKeyPress()`, `FindNodeOrItemByName()`

#### Form1_Main.vb `FetchDirectMailCountsAsync()`（缺結束）
```vb
' 在 Return False 之前加入:
Dbg("結束", $"共讀取 {total} 個節點（非快取）")
```

#### Form1.vb `HandleListViewKeyPress()`
```vb
' 現在: 只有 Dbg("開始")，各分支 Return 前無輸出
' 此函數是高頻事件（每次按鍵都觸發），建議：
' 方案A: 完全移除 Dbg（按鍵事件太頻繁）
' 方案B: 只保留非平凡路徑的輸出（Enter/ESC 才記錄）
```

#### Form1.vb `HandleTreeViewKeyPress()`
```vb
' 建議: 移除 Dbg("開始")，只在 Enter/ESC 時輸出
Dbg("Enter鍵進入資料夾")  ' 或 Dbg("ESC鍵退回上層")
```

---

### Issue 3 — 修正 Return 前缺少 Dbg 的問題 🔴

#### Form1_Main.vb `TreeView1_AfterSelect()`（L241）
```vb
' 現在:
If _tab1SelectSeq <> mySeq Then Return

' 改成:
If _tab1SelectSeq <> mySeq Then
    Dbg("結束", "序號不匹配，已放棄（快速點選被丟棄）")
    Return
End If
```

#### Form1_Main.vb `TreeView1_AfterSelect()`（L244）
```vb
' 現在:
If _cancelRequested OrElse rows.Count = 0 Then
    ProgressBar1.Text = "已中斷。" : Cursor = Cursors.Default : Return

' 改成:
If _cancelRequested OrElse rows.Count = 0 Then
    Dbg("結束", If(_cancelRequested, "ESC 中斷", "結果為空"))
    ProgressBar1.Text = "已中斷。" : Cursor = Cursors.Default : Return
End If
```

#### Form1_Main.vb `SimTree2_AfterSelect()`（L637-644）
```vb
' 現在:
If selectedNodes Is Nothing OrElse selectedNodes.Count = 0 Then
    Cursor = Cursors.Default : Return

' 改成:
If selectedNodes Is Nothing OrElse selectedNodes.Count = 0 Then
    Dbg("結束", "無節點被選取")
    Cursor = Cursors.Default : Return
End If
```

#### Form1.vb `ExpandTreeToDefaultInbox()`（找不到收件匣時）
```vb
' 在迴圈後加:
Dbg("結束", $"找不到預設收件匣，節點總計: {rootNode.Nodes.Count}")
```

---

### Issue 4 — 格式標準化（msg 不應包含函數名或串接字串）🟡

#### 修正錯誤格式的 Dbg 呼叫

以下是需要格式化修正的位置：

**Form1_Main.vb**：
```vb
' 現在 (L261):
Dbg("Error: TreeView1_AfterSelect", ex.Message)
' 改成:
Dbg("錯誤", ex.Message)

' 現在 (L879):
Dbg("結束", $"共 {merged.Count} 個年份 | 郵件總計: {merged.Values.Sum}")
' OK，可接受

' 現在 (L925):
Dbg("錯誤", folder.Name & " - " & ex.Message)
' 改成:
Dbg("錯誤", $"{folder.Name}: {ex.Message}")

' 現在 (L977):
Dbg("GetMonthCountsForYear Error: ", folder.Name & $", year={year} - " & ex.Message)
' 改成:
Dbg("錯誤", $"{folder.Name}, year={year}: {ex.Message}")

' 現在 (L1520):
Dbg("開始", "排序列表")
Dbg("結束", "排序列表")
' 改成 (detail 沒意義):
Dbg("開始", $"{ListView3.Items.Count} 項，點選第 {e.Column+1} 欄")
Dbg("結束", $"{sw.Elapsed.TotalMilliseconds:0.0}ms")
```

**Form1_ComL3.vb**：
```vb
' 現在 (L183):
Dbg("GetSortedSubFolders 遍歷失敗", ex.Message)
' 改成:
Dbg("錯誤", ex.Message)

' 現在 (L238):
Dbg("GetSubFolderList ① OOM 失敗", current.Name & " - " & ex.Message)
' 改成:
Dbg("錯誤", $"OOM: {current.Name} — {ex.Message}")
```

**Form1.vb**：
```vb
' 現在 (L675):
Dbg("SafeGet(Row) 失敗", $"{column} | {ex.Message}")
' 改成:
Dbg("錯誤", $"SafeGet({column}): {ex.Message}")

' 現在 (L1008):
Dbg("縮合側邊欄: " & sc.Name & ", 原始寬度: " & sc.Tag.ToString)
' 改成:
Dbg("縮合側邊欄", $"{sc.Name} → 10px (原 {sc.Tag})")

' 現在 (L1014):
Dbg("恢復側邊欄: " & sc.Name & ", 目標寬度: " & prevDist)
' 改成:
Dbg("恢復側邊欄", $"{sc.Name} → {prevDist}px")
```

---

### Issue 5 — `GetSizeMultiplier()` 的「結束」永遠不會執行 🟡

**Form1_Main.vb**（約 L1766-1776）

```vb
Private Function GetSizeMultiplier(sizeUnit As String, Optional base1024 As Boolean = False) As Integer
    Dbg("開始")
    ...
    Select Case sizeUnit.ToLower()
        Case "kb" : Return multi     ' ← Return 後不會繼續執行
        Case "mb" : Return multi ^ 2
        ...
    End Select
    Dbg("結束")  ' ← 永遠到不了這行
End Function
```

**建議**：此函數極其簡單，完全移除兩個 Dbg。或改為：
```vb
Dim result As Integer = ...  ' 計算結果
Dbg("", $"{sizeUnit} → ×{result}")
Return result
```

---

### Issue 6 — 補強重要函數的 detail 內容 🟡

#### `ScanAttachmentDetail()` 開始補入處理量（Form1_Main.vb）
```vb
' 現在:
Dbg("開始")
' 改成:
Dbg("開始", $"候選郵件: {targetMailList.Count} 封")
```

#### `BuildBfsFolderTree()` 補入快取比例（Form1_Main.vb）
```vb
' 在結束之前補充:
Dim fromCache = allEntries.Count(Function(e) e.IsFromCache)
Dim fromScan = allEntries.Count - fromCache
Dbg("BFS 完成", $"節點: {allEntries.Count}（快取 {fromCache} / 新掃描 {fromScan}）")
```

#### `CheckTab3CacheOrRescan()` 補強開始 detail（Form1_Main.vb）
```vb
' 現在:
Dbg("開始", targetFolder.Name)
' 結束時補入找到的郵件數:
Dbg("結束", $"{targetFolder.Name} → {targetMailList.Count} 封有附件")
```

---

### Issue 7 — 冗餘的 Dbg 可以移除 🟢

以下 Dbg 可考慮整合或移除：

| 位置 | 函數 | 移除原因 |
|---|---|---|
| Form1_Main.vb ~L1224 | `ShowProgressTab2()` 開始/結束 | 函數體只有 5 行計算，不需要追蹤 |
| Form1_Main.vb ~L1950-1953 | `ListView4_SelectedIndexChanged()` | 函數體是空的（只有 todo 註解），開始/結束毫無意義 |
| Form1.vb ~L1766 | `GetSizeMultiplier()` | 上述 Issue 5，結束永遠到不了 |
| Form1_Main.vb ~L1779 | `OpenMailByEntryID()` | `"打開郵件"` 不是標準格式，改為 `"開始"` |
| Form1_ComL3.vb ~L1226 | `AutoDismissRedemptionDialog()` 內部 | Thread 內的 `Dbg()` 呼叫是跨執行緒的，可能有問題（`Dbg` 內部有 `_isDebugMode` 和 `DebugForm.AddMessage3` 呼叫） |
| Form1_Main.vb ~L2124 | `CalculateSimilarity()` 開始/結束 | Tab5 高頻呼叫 |
| Form1_Main.vb ~L2135 | `LevenshteinDistance()` 開始/結束 | 同上 |

---

## 關於 `AutoDismissRedemptionDialog()` 的特殊注意 ⚠️

此函數在 `Thread` 的 Lambda 內呼叫 `Dbg()`，而 `Dbg()` 最終會呼叫 `DebugForm.AddMessage3()`。

`AddMessage3()` 本身使用 `ConcurrentQueue` + Timer 批次更新，理論上是執行緒安全的。但 `_isDebugMode` 的讀取沒有任何同步保護（雖然布林讀取在 .NET 通常是原子操作）。

**建議**：保持現有的 `Dbg()` 呼叫不做改變，因為它已正確使用且現有架構支援跨執行緒。

---

## 建議保留的 Dbg 模式（不需改動）

以下函數的 Dbg 品質已達到良好標準，**不需要修改**：

- `ComputeFolderStatsAsync()` — 開始傳入 rootFolder.Name ✅
- `GetBfsResult()` — 結束傳入 result.Count 和直屬子資料夾數 ✅
- `GetSubFolderList()` — Cache Hit 路徑和結束都有完整資訊 ✅
- `GetMailCountAll()` — 各 fallback 路徑都有清晰的開始/成功/失敗標記 ✅
- `ScanFolderWithAttachment()` — 結束傳入找到的郵件數 ✅
- `ScanAttachmentDetail()` — 結束傳入篩選後郵件數 ✅（只是缺少開始的 detail）
- `GetSortedStores()` — 結束傳入 Profile 名稱和庫數量 ✅
- `GetSortedSubFolders()` — 開始/結束都完整 ✅
- `InitRedemptionSessionWithoutDeclaration()` — 各路徑都有清晰標記 ✅
- `WaitAndYieldIfBusy()` — 只在忙碌時輸出，不是固定呼叫 ✅

---

## 對 DebugForm 的 `FindSimilarPair` 影響說明

`FindSimilarPair` 目前依賴 `"開始"` / `"結束"` 關鍵字做配對。如果我們把某些函數內的：

```vb
Dbg("Error: GetXxx", ex.Message)
```

改成：

```vb
Dbg("錯誤", ex.Message)
```

這些「錯誤」行**不含「開始」或「結束」**，不會被 `FindSimilarPair` 嘗試配對，這是正確的行為。

---

## 實作優先順序建議

> [!IMPORTANT]
> 以下是建議的實作順序，請 Simon 確認哪些要加入、哪些可跳過。

| 優先度 | Issue | 函數數量 | 說明 |
|---|---|---|---|
| 🔴 立即處理 | Issue 1：高頻迴圈去噪 | 7個 | 防止 DebugForm 被灌爆 |
| 🔴 立即處理 | Issue 2：補結束 | 4個 | 讓 DebugForm 配對計時正常運作 |
| 🔴 立即處理 | Issue 3：早期Return補Dbg | 4個 | 讓診斷路徑不出現空白段 |
| 🟡 次要處理 | Issue 4：格式標準化 | ~12處 | 提升可讀性 |
| 🟡 次要處理 | Issue 5：GetSizeMultiplier | 1個 | 修正永遠到達不了的結束 |
| 🟡 次要處理 | Issue 6：補強 detail | 3個函數 | 提升監控資訊密度 |
| 🟢 可跳過 | Issue 7：移除冗餘 | 7處 | 精簡但非必要 |

