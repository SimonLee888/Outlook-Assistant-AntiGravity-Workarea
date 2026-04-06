# Dbg() 品質改善 — 修改完成總覽

## 完成日期
2026/04/04

## 修改範圍

共修改三個檔案：`Form1_ComL3.vb`、`Form1_Main.vb`、`Form1.vb`

---

## Issue 1：高頻迴圈去噪

**目標**：避免在高頻執行路徑上導出大量 Dbg 行，灌爆 DebugForm 的 ListView。

| 函數 | 修改內容 |
|---|---|
| `LoadStoreToTreeView()` | 移除 For Each 迴圈內每筆 Dbg，改在結束輸出 `共 N 個 Store` |
| `LoadSubFolderToTreeView()` | 移除 For Each 迴圈內每筆 Dbg，改在結束輸出子資料夾數量 |
| `GetMailCount()` | 三條 fallback **成功路徑**全面靜默；失敗路徑改為 `Dbg("錯誤路徑", detail)` |
| `GetFolderCount()` | 兩條成功路徑靜默；失敗路徑標準化 |
| `IsMailFolder()` | 移除 `Dbg("開始")`，改為只在**非郵件資料夾**時輸出一行（真正有用的過濾事件） |
| `CalculateSimilarity()` | 完全移除開始/結束（Tab5 N封×2函數=2N行輸出） |
| `LevenshteinDistance()` | 完全移除開始/結束（同上） |

---

## Issue 2：補缺失的「結束」

| 函數 | 修改內容 |
|---|---|
| `FetchDirectMailCountsAsync()` | ESC 中斷路徑補 `Dbg("結束", "ESC 中斷")`；正常完成路徑補節點數量 |
| `HandleTreeViewKeyPress()` | 移除 `Dbg("開始")` — 高頻按鍵事件不需要追蹤 |
| `FindNodeOrItemByName()` | 找到/找不到兩條路徑各補 `Dbg("結束", ...)` |

---

## Issue 3：早期 Return 前補 Dbg

| 函數 | 補充位置 |
|---|---|
| `TreeView1_AfterSelect()` | 序號不匹配路徑補 `Dbg("結束", "序號已不匹配，市棄本次結果")` |
| `TreeView1_AfterSelect()` | `_cancelRequested` 路徑補 `Dbg("結束", ...)` |
| `SimTree2_AfterSelect()` | 無節點選取路徑補 `Dbg("結束", "無節點被選取")` |
| `SimTree2_AfterSelect()` | targetFolderList 為空路徑補 `Dbg("結束", "所有選定節點均無效資料夾")` |
| `ExpandTreeToDefaultInbox()` | 迴圈掃完找不到收件匣時補 `Dbg("結束", "找不到預設收件匣...")` |

---

## Issue 4：格式標準化

**規則**：`msg` 永遠使用固定關鍵字，不含函數名或字串串接。

| 修改位置 | 舊格式 | 新格式 |
|---|---|---|
| `TreeView1_AfterSelect()` | `Dbg("Error", ...)` | `Dbg("錯誤", ...)` |
| `GetMonthCountsForYear()` | `Dbg("GetMonthCountsForYear Error: ", ...)` | `Dbg("錯誤", ...)` |
| `Button3_Click()` | `Dbg("Button3_Click Error: ", ...)` | `Dbg("錯誤", ...)` |
| `Button4_Click()` | `Dbg("Button4 GetTable Error: " & folder.Name, ...)` | `Dbg("錯誤", $"...")` |
| `Button5_Click()` | `Dbg("Button5 GetTable Error: " & folder.Name, ...)` | `Dbg("錯誤", $"...")` |
| `Button5_Click()` | `Dbg("Button5 Store Error: ", ...)` | `Dbg("錯誤", $"store: msg")` |
| `OpenMailByEntryID()` | `Dbg("打開郵件", entryID)` | `Dbg("開始", entryID)` |
| 側邊欄縮合/恢復 | `Dbg("縮合側邊欄: " & sc.Name & "...")` | `Dbg("縮合側邊欄", $"...")` |

---

## Issue 5：GetSizeMultiplier 死碼修正

`GetSizeMultiplier()` 的 `Select Case + Return` 結構導致 `Dbg("結束")` 永遠無法被執行。

**修改**：移除 `Dbg("開始")` 和 `Dbg("結束")`，補上說明性註解。

---

## Issue 6：補強 detail 內容

| 函數 | 補充內容 |
|---|---|
| `ScanAttachmentDetail()` | 開始時補入 `targetMailList.Count` 作為最重要的監控資訊 |

---

## Issue 7：移除冗餘 Dbg

| 函數 | 修改內容 |
|---|---|
| `ShowProgressTab2()` | 移除開始/結束（函數體只有 5 行計算） |
| `ListView4_SelectedIndexChanged()` | 移除開始/結束（函數體為空的 todo） |

---

## 驗證結果

- 搜尋所有舊格式錯誤訊息（含函數名或字串串接）→ 零殘留
- 高頻迴圈函數成功路徑不再產生大量輸出
- DebugForm 的 `FindSimilarPair` 配對邏輯不受影響（開始/結束對稱結構完整）
- 所有早期 Return 路徑現在都有對應的 Dbg 輸出

## 備註：刻意保留的非標準格式

`GetMailCountAll()` 系列函數（含 `GetMailCountAllParallel`）的 Dbg 訊息仍保留函數名在 msg 中，這是**刻意為之**：這些函數有複雜的 4 層 fallback chain，識別是哪一條 fallback 路徑出問題比遵守格式規範更重要。
