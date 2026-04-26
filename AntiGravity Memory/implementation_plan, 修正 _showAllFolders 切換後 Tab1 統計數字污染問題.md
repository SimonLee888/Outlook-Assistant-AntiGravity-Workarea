# 修正 `_showAllFolders` 切換後 Tab1 統計數字污染問題

## 問題背景

`_cacheMailCountAll` 和 `_cacheFolderCountAll` 是 BFS 加總的核心快取（含子孫的郵件/資料夾總數）。
這兩個字典的**鍵值只有 `fPath`**，沒有攜帶 `_showAllFolders` 的模式資訊。

**污染觸發情境：**
1. `_showAllFolders = False`：點選資料夾 A → BFS 只走郵件資料夾 → 寫入 `_cacheMailCountAll["\\PST\A"] = 100`
2. `_showAllFolders = True`：再次點選資料夾 A → BFS 的快取查詢用 `fPath` 命中 100 → **直接剪枝跳過深層節點**
3. 結果：勾選「顯示所有資料夾」後，深層（depth ≥ 2）的行事曆/連絡人加總**被遺漏**

> [!IMPORTANT]
> 現有的 `CheckShowAllFolders_CheckedChanged` 只清了 `_cacheFolderTree`（子資料夾清單），**沒有清 `_cacheMailCountAll` / `_cacheFolderCountAll`**。
> 因此切換 `_showAllFolders` 後，即使 BFS 重新展開 root 的直屬子資料夾，深層節點仍會讀到舊模式的加總值。

---

## 影響範圍分析

| 層次 | 受影響物件 | 問題描述 |
|------|-----------|---------|
| **記憶體快取** | `_cacheMailCountAll` | 鍵值無模式分支，跨模式污染 |
| **記憶體快取** | `_cacheFolderCountAll` | 同上 |
| **BFS Step 1** | `BuildBfsFolderTree` (Form1_MainTabs.vb L571) | 查詢時用 `fPath`，命中舊模式快取就剪枝 |
| **BFS Step 4** | `UpdateFolderStatsCache` (Form1_MainTabs.vb L675-676) | 寫入時用 `fPath`，覆蓋另一個模式的值 |
| **DB lazy load** | `BuildBfsFolderTree` (L576-581) | DB 的 `mca`/`fca` 欄位無模式欄，載入後放進無分支快取 |
| **DB 寫入** | `SaveFolderStatsInner` (Form1_SQLite2.vb L793-795) | 直接從 `_cacheMailCountAll` 遍歷，鍵值有分支後要對應取出 |
| **DB 全量讀取** | `LoadFolderStatsInner` (Form1_SQLite2.vb L1056-1058) | 讀出 `mca`/`fca` 直接塞入無分支快取 |
| **DB 個別讀取** | `FillFolderCacheFromDbRow` (Form1_Outlook.vb L2232-2234) | 同上 |
| **切換事件** | `CheckShowAllFolders_CheckedChanged` (Form1.vb L959) | 切換時沒有清 `_cacheMailCountAll` / `_cacheFolderCountAll` |
| **RenewCache** | `RenewCacheAsync` Phase 3/4 (Form1_SQLite2.vb L582-604) | 清除 dirty folder 與 ancestor 時只用 `fPath`，分支後要兩個都清 |

---

## 設計決策

> [!IMPORTANT]
> **DB schema 不改動**。`folder_stats` 表維持現狀，`mca`/`fca` 兩個欄位只存一個值。
> 
> 設計原則：**DB 只存「目前 `_showAllFolders` 模式」的計算結果**，不同時儲存兩種模式。
> 這與 `_cacheFolderTree` 的設計完全相同。

**理由：**
- DB schema 新增欄位代表整個儲存與讀取邏輯都要改，工程量大且 schema migration 需要 ALTER TABLE。
- 大部分使用者長期只用一種模式；DB 只存當前模式的值，切換後清掉聚合快取，重新點選即重算並寫入新模式值，行為正確且實現簡單。
- 與 `_cacheFolderTree`（記憶體分支鍵值，DB 不存）的設計一致。

---

## 修改方案

### 方案核心：記憶體快取鍵值加分支，DB 保持單值但切換時清空

**記憶體快取的鍵值**：`fPath` → `fPath & "|" & _showAllFolders`（與 `_cacheFolderTree` 完全相同的策略）

**DB 的處理**：切換 `_showAllFolders` 時，在 `CheckShowAllFolders_CheckedChanged` 中**同時清除** `_cacheMailCountAll` / `_cacheFolderCountAll`。DB 中的 `mca`/`fca` 值在切換後下一次 RenewCache 時自然被新模式值覆蓋（INSERT OR REPLACE）。

---

## 詳細修改清單

### 修改點 A — `BuildBfsFolderTree`（Form1_MainTabs.vb）

#### [MODIFY] 快取查詢（L571）與 DB lazy load 後的記憶體填入（L584-587）

```vb
' 舊
Dim cachedMail As Integer, cachedSub As Integer
Dim isHit As Boolean = False
If _cacheMailCountAll.TryGetValue(fPath, cachedMail) AndAlso _cacheFolderCountAll.TryGetValue(fPath, cachedSub) Then
    isHit = True    ' ① 記憶體命中
Else
    Dim row = DbGetFolderStats(fPath)
    If row IsNot Nothing AndAlso row.mca >= 0 AndAlso row.fca >= 0 Then
        cachedMail = row.mca : cachedSub = row.fca
        FillFolderCacheFromDbRow(fPath, row)
        isHit = True    ' ② DB 命中
    End If
End If

If isHit Then
    entry.TotalMailCount = cachedMail
    entry.TotalSubCount = cachedSub
    entry.IsFromCache = True
    ...
```

```vb
' 新（加入 cacheKey 分支）
Dim cacheKey As String = fPath & "|" & _showAllFolders   ' by Claude Sonnet 4.6, 2026/04/25
Dim cachedMail As Integer, cachedSub As Integer
Dim isHit As Boolean = False
If _cacheMailCountAll.TryGetValue(cacheKey, cachedMail) AndAlso _cacheFolderCountAll.TryGetValue(cacheKey, cachedSub) Then
    isHit = True    ' ① 記憶體命中
Else
    Dim row = DbGetFolderStats(fPath)
    ' DB lazy load：mca/fca 是當時寫入時的模式值
    ' 只在當時模式與現在模式相同（或無法判斷）時才信任 DB 值
    ' 設計決策：DB 不儲存模式資訊，切換後 CheckShowAllFolders 會清空記憶體快取，
    '           因此 DB 命中時記憶體必為空（未被舊模式污染），可安全使用 DB 值
    If row IsNot Nothing AndAlso row.mca >= 0 AndAlso row.fca >= 0 Then
        cachedMail = row.mca : cachedSub = row.fca
        FillFolderCacheFromDbRow(fPath, row, _showAllFolders)   ' 傳入 mode 寫入分支鍵值
        isHit = True    ' ② DB 命中
    End If
End If

If isHit Then
    entry.TotalMailCount = cachedMail
    entry.TotalSubCount = cachedSub
    entry.IsFromCache = True
    ...
```

---

### 修改點 B — `UpdateFolderStatsCache`（Form1_MainTabs.vb）

```vb
' 舊（L675-676）
_cacheMailCountAll.TryAdd(fPath, entry.TotalMailCount)
_cacheFolderCountAll.TryAdd(fPath, entry.TotalSubCount)

' 新
Dim cacheKey As String = entry.FolderPath & "|" & _showAllFolders   ' by Claude Sonnet 4.6, 2026/04/25
_cacheMailCountAll.TryAdd(cacheKey, entry.TotalMailCount)
_cacheFolderCountAll.TryAdd(cacheKey, entry.TotalSubCount)
```

---

### 修改點 C — `FillFolderCacheFromDbRow`（Form1_Outlook.vb）

加入 `Optional showAllMode As Boolean` 參數，讓呼叫端傳入目前的模式，寫入分支鍵值：

```vb
' 舊函數簽名
Private Sub FillFolderCacheFromDbRow(fPath As String, row As FolderStatsDbRow)

' 新函數簽名（加 Optional 參數，預設 Nothing = 用 fPath 原始鍵，向下相容）
Private Sub FillFolderCacheFromDbRow(fPath As String, row As FolderStatsDbRow, Optional showAllMode As Boolean? = Nothing)
    ...
    ' mca/fca 寫分支鍵值；其餘欄位（mc/fc/fs/fsa）維持原 fPath
    Dim mcaKey As String = If(showAllMode.HasValue, fPath & "|" & showAllMode.Value, fPath)
    If row.mca >= 0 Then _cacheMailCountAll.TryAdd(mcaKey, row.mca)
    If row.fca >= 0 Then _cacheFolderCountAll.TryAdd(mcaKey, row.fca)
    ...
```

> [!NOTE]
> `GetMailCountAllAsync` / `GetFolderCountAllAsync` 是死碼（目前無呼叫端），它們直接用 `fPath` 讀寫 `_cacheMailCountAll`。
> 因為是死碼，**本次不修改**，但加上註解說明鍵值已改為分支格式，未來若要啟用需同步更新。

---

### 修改點 D — `SaveFolderStatsInner`（Form1_SQLite2.vb）

`SaveFolderStatsInner` 從 `_cacheMailCountAll.Keys` 收集所有路徑，存入 DB 的 `mca` 欄位。

修改後鍵值為 `fPath|True` 或 `fPath|False`，需要在遍歷時**解析出真實 fPath** 再存入 DB：

```vb
' 舊（L760）
For Each k In _cacheMailCountAll.Keys : allPaths.Add(k) : Next
For Each k In _cacheFolderCountAll.Keys : allPaths.Add(k) : Next

' 新（解析出 fPath 再加入聯集）
For Each k In _cacheMailCountAll.Keys
    Dim rawPath = k.Split("|"c)(0)   ' 去掉 "|True" 或 "|False" 後綴
    allPaths.Add(rawPath)
Next
For Each k In _cacheFolderCountAll.Keys
    Dim rawPath = k.Split("|"c)(0)
    allPaths.Add(rawPath)
Next
```

在寫入每個 `path` 的 `mca`/`fca` 時，優先取現行模式的快取鍵值：

```vb
' 舊（L793-795）
Dim hasMca = _cacheMailCountAll.TryGetValue(path, mca)
Dim hasFca = _cacheFolderCountAll.TryGetValue(path, fca)

' 新（優先取當前模式的鍵值）
Dim currentModeKey = path & "|" & _showAllFolders
Dim hasMca = _cacheMailCountAll.TryGetValue(currentModeKey, mca)
If Not hasMca Then hasMca = _cacheMailCountAll.TryGetValue(path & "|False", mca)  ' fallback
If Not hasMca Then hasMca = _cacheMailCountAll.TryGetValue(path & "|True", mca)   ' fallback
Dim hasFca = _cacheFolderCountAll.TryGetValue(currentModeKey, fca)
If Not hasFca Then hasFca = _cacheFolderCountAll.TryGetValue(path & "|False", fca)
If Not hasFca Then hasFca = _cacheFolderCountAll.TryGetValue(path & "|True", fca)
```

> [!NOTE]
> Fallback 策略：若當前模式無資料，退而其次寫另一模式的值（避免 DB 空白）。
> 切換模式後 RenewCache 會以新模式值覆蓋。

---

### 修改點 E — `LoadFolderStatsInner`（Form1_SQLite2.vb）

DB 讀出 `mca`/`fca` 後，用**當前模式的分支鍵值**寫入記憶體：

```vb
' 舊（L1056-1058）
If Not reader.IsDBNull(2) Then _cacheMailCountAll.TryAdd(path, reader.GetInt32(2))
If Not reader.IsDBNull(4) Then _cacheFolderCountAll.TryAdd(path, reader.GetInt32(4))

' 新（分支鍵值）
Dim modeKey = path & "|" & _showAllFolders
If Not reader.IsDBNull(2) Then _cacheMailCountAll.TryAdd(modeKey, reader.GetInt32(2))
If Not reader.IsDBNull(4) Then _cacheFolderCountAll.TryAdd(modeKey, reader.GetInt32(4))
```

---

### 修改點 F — `CheckShowAllFolders_CheckedChanged`（Form1.vb）

切換時額外清除聚合快取（這是整個問題的根本修復點）：

```vb
' 舊（L969）
_showAllFolders = checkShowAllFolders.Checked
_cacheFolderTree.Clear()

' 新（多清兩個聚合快取）
_showAllFolders = checkShowAllFolders.Checked
_cacheFolderTree.Clear()
_cacheMailCountAll.Clear()      ' by Claude Sonnet 4.6, 2026/04/25: 聚合快取含模式，切換後必須清空
_cacheFolderCountAll.Clear()    ' 同上
```

> [!IMPORTANT]
> 修改點 F **本身就能解決當前污染問題**（切換後清空即可）。
> 但如果只做 F，下次切換回來又會有「跨模式填入」的問題，因此 A/B 的鍵值分支是長期正確性的保障。

---

### 修改點 G — `RenewCacheAsync` Phase 3/4（Form1_SQLite2.vb）

Phase 3 清除 dirty folder 的聚合快取，Phase 4 清除 ancestor 的聚合快取，都要改為雙鍵清除：

```vb
' Phase 3 舊（L582-583）
_cacheMailCountAll.TryRemove(fPath, Nothing)
_cacheFolderCountAll.TryRemove(fPath, Nothing)

' Phase 3 新（兩個模式都清）
_cacheMailCountAll.TryRemove(fPath & "|True", Nothing)
_cacheMailCountAll.TryRemove(fPath & "|False", Nothing)
_cacheFolderCountAll.TryRemove(fPath & "|True", Nothing)
_cacheFolderCountAll.TryRemove(fPath & "|False", Nothing)

' Phase 4 舊（L600-601）
_cacheMailCountAll.TryRemove(ancestor, Nothing)
_cacheFolderCountAll.TryRemove(ancestor, Nothing)

' Phase 4 新
_cacheMailCountAll.TryRemove(ancestor & "|True", Nothing)
_cacheMailCountAll.TryRemove(ancestor & "|False", Nothing)
_cacheFolderCountAll.TryRemove(ancestor & "|True", Nothing)
_cacheFolderCountAll.TryRemove(ancestor & "|False", Nothing)
```

---

### 修改點 H — 其他 TryRemove 呼叫

`Form1.vb ClearMemoryCachesInternal`（L1016-1018）：`.Clear()` 不受影響，直接清空整個字典，**不需要修改**。

`Form1_Win32API.vb`（L185,205,215,228）：這些是舊版函數入口，目前已無 Tab1 主流程呼叫，**本次不修改**，加上注意事項注解。

---

## 修改順序（執行時依序完成）

```
修改 F → 修改 A → 修改 B → 修改 C → 修改 D → 修改 E → 修改 G
```

先做 F（切換清空）確保立即可用，再依序完成架構性修正。

---

## 驗證計畫

1. **功能驗證**：
   - 以 `_showAllFolders = False` 點選含行事曆/聯絡人子資料夾的根目錄 → 確認數字只含郵件資料夾
   - 切換到 `_showAllFolders = True` 再次點選同一目錄 → 確認數字增加（包含非郵件資料夾）
   - 切換回 False → 確認數字回到郵件資料夾數字

2. **快取驗證**：
   - 切換後不應再看到舊數字（不需要 ClearCache 才能得到正確值）
   - SaveCache / LoadCache 後再切換，確認數字仍正確

3. **RenewCache 驗證**：
   - 執行 RenewCache → 確認不會將 Phase 3/4 的清除操作遺漏任一模式的鍵值

---

## 開放問題

> [!NOTE]
> **`GetMailCountAllAsync` / `GetFolderCountAllAsync`** 是目前死碼（`Form1_Outlook.vb` L479~547 有明確標注），直接讀 `_cacheMailCountAll[fPath]`（無分支）。
> 本計劃**不改這兩個函數**，只加上注解說明。若未來要復活這兩個函數，需同步改為讀分支鍵值。
