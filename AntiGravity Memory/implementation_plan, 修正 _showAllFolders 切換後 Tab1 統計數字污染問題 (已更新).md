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

| 層次 | 受影響物件 | 問題描述 | 本次狀態 |
|------|-----------|---------|----------|
| **記憶體快取** | `_cacheMailCountAll` | 鍵值無模式分支，跨模式污染 | ✅ 已解決（修改 F 切換時 Clear） |
| **記憶體快取** | `_cacheFolderCountAll` | 同上 | ✅ 已解決（修改 F） |
| **BFS Step 1** | `BuildBfsFolderTree` (Form1_MainTabs.vb L571) | 查詢時用 `fPath`，命中舊模式快取就剪枝 | ✅ 已解決（切換後快取已清空，BFS 查詢 miss → 重算） |
| **BFS Step 4** | `UpdateFolderStatsCache` (Form1_MainTabs.vb L675-676) | 寫入時用 `fPath`，覆蓋另一個模式的值 | ✅ 已解決（切換後快取清空，TryAdd 不覆蓋） |
| **DB lazy load** | `BuildBfsFolderTree` (L576-581) | DB 的 `mca`/`fca` 欄位無模式欄，切換後載入舊模式值填入記憶體快取，BFS 仍然剪枝 | ❌ **殘留**（見下方說明） |
| **DB 寫入** | `SaveFolderStatsInner` (Form1_SQLite2.vb L793-795) | 直接從 `_cacheMailCountAll` 遍歷，鍵值有分支後要對應取出 | ⬛ 不適用（本次未做鍵值分支，原行為不變） |
| **DB 全量讀取** | `LoadFolderStatsInner` (Form1_SQLite2.vb L1056-1058) | 讀出 `mca`/`fca` 直接塞入無分支快取，LoadCache 後模式不符時污染記憶體 | ❌ **殘留** |
| **DB 個別讀取** | `FillFolderCacheFromDbRow` (Form1_Outlook.vb L2232-2234) | 同 DB lazy load 問題 | ❌ **殘留** |
| **切換事件** | `CheckShowAllFolders_CheckedChanged` (Form1.vb L959) | 切換時沒有清 `_cacheMailCountAll` / `_cacheFolderCountAll` | ✅ 已解決（修改 F） |
| **RenewCache** | `RenewCacheAsync` Phase 3/4 (Form1_SQLite2.vb L582-604) | 清除 dirty folder 與 ancestor 時只用 `fPath`，分支後要兩個都清 | ✅ 已解決（修改 G，同時清 `\|True`/`\|False`/bare `fPath`） |
| **RenewCache** | `RenewCacheAsync` Phase 2/3 | 清空後 RenewCache 把所有資料夾視為全新 → 偷跑全量 attach_maillist 掃描 | ✅ 已解決（修改 Phase2 語意，`dirtyNewFolderSet` 跳過 attach_maillist） |

---

## ✅ 本次已解決（2026/04/25）

1. **切換 `_showAllFolders` 後記憶體快取污染** — `CheckShowAllFolders_CheckedChanged` 多清 `_cacheMailCountAll` / `_cacheFolderCountAll`
2. **RenewCache 清空後偷跑 2 萬筆 attach_maillist 掃描** — Phase 2 區分「全新（isNewFolder）」vs「真正 dirty（snapshot 不符）」，全新資料夾跳過 `RenewAttachMailListAsync`
3. **RenewCache Phase 3/4 清除遺漏模式鍵值** — 同時清 `|True` / `|False` / bare `fPath` 三種格式

---

## ❌ 殘留問題（DB lazy load 跨模式污染）

> [!WARNING]
> **觸發情境：**
> 1. 以 `_showAllFolders = False` 模式操作，SaveCache → DB 寫入 `mca = 100`（郵件資料夾加總）
> 2. 切換到 `_showAllFolders = True`，修改 F 清空了記憶體快取
> 3. 點選深層（depth ≥ 2）資料夾 → BFS 記憶體 miss → **DB lazy load 命中** → 讀出 `mca = 100`（舊 False 模式的值）
> 4. `FillFolderCacheFromDbRow` 把 100 填入 `_cacheMailCountAll[fPath]` → BFS 以為命中，跳過展開
> 5. 顯示數字仍是 100，但 True 模式應有行事曆/連絡人加入，正確值應更大

**受影響函數：**
- `BuildBfsFolderTree` L576-581（DB lazy load 路徑）
- `FillFolderCacheFromDbRow` L2232-2234（DB 值填入記憶體）
- `LoadFolderStatsInner` L1056-1058（LoadCache 按鈕批次讀取）

**現況評估：**
- 觸發條件需「曾在某個模式 SaveCache → 切換模式 → 點選深層資料夾」這個精確順序
- 如果使用者從不按 SaveCache，或每次切換後都 RenewCache，此問題不會觸發
- `mca/fca` 的 DB lazy load 讓深層節點免於完整 BFS 展開（效能），但同時帶入了這個模式無感知的問題

---

## 🔜 下次待處理

**選項 A（最小修正）：`BuildBfsFolderTree` 的 DB lazy load 不讀 `mca`/`fca` 做剪枝**

切換後第一次 BFS 完整展開，結果寫入記憶體快取，後續點選從記憶體命中（快）。

```vb
' 修改點：BuildBfsFolderTree L576-581
' 舊：DB lazy load 讀 mca/fca，命中即剪枝
Dim row = DbGetFolderStats(fPath)
If row IsNot Nothing AndAlso row.mca >= 0 AndAlso row.fca >= 0 Then
    cachedMail = row.mca : cachedSub = row.fca
    FillFolderCacheFromDbRow(fPath, row)
    isHit = True

' 新：DB lazy load 只填 mc/fc 等本層欄位，不以 mca/fca 做剪枝
' mca/fca 無論 DB 有無都不用，讓 BFS 自行展開計算後寫入記憶體
Dim row = DbGetFolderStats(fPath)
If row IsNot Nothing Then
    FillFolderCacheFromDbRow(fPath, row, skipAggregates:=True)   ' 不填 mca/fca 到記憶體
    ' isHit 仍為 False → BFS 繼續展開子資料夾
End If
```

**代價**：切換後第一次統計要完整展開（不能 DB 剪枝），但這是正確行為，且之後記憶體命中就快了。

**選項 B（長期架構）：DB schema 加 `show_all_mode` 欄位，`mca`/`fca` 依模式分行儲存**
- 工程量大，需 ALTER TABLE + 讀寫邏輯全改
- 除非有明確的多模式 DB 持久化需求，否則選項 A 已足夠
