# Form1.vb 效能優化 — Walkthrough

## 修改摘要

對 [Form1.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1.vb) 進行了 7 項效能優化，全部標記為 `2026/3/24 by AntiGravity`。

---

### 1. GetTable.GetArray(1000) 批次讀取

將 3 個函數從逐行 `GetNextRow()` 改為 `GetArray(1000)` 批次讀取：

| 函數 | 修改行 | 影響 |
|---|---|---|
| `ScanFolderWithAttachment` | L3224 | Tab3 Phase1 附件掃描 |
| `GetFolderSize` ① | L4509 | L3 COM 資料夾大小計算 |
| `GetFolderSizeLegacy` | L1842 | Tab1 資料夾大小 (legacy) |

> [!TIP]
> GetArray 一次傳回最多 1000 筆 row 的 `Object(,)` 二維陣列，將 COM 跨程序呼叫從 N 次降到 ⌈N/1000⌉ 次。

---

### 2. Redemption 平行化安全防護

`GetMailCountAll` ① 的 `Task.WhenAll` 平行路徑加上 `_rdo IsNot Nothing` 防護，確保只在 Redemption (free-threaded) 可用時才平行化。若 Redemption 不可用，直接走 ② 循序 BFS，避免 OOM COM 在背景執行緒觸發 STA 違規。

`GetFolderSizeAll` ① 不需修改 — 它使用 async interleaving 而非 `Task.Run`，STA 安全。

---

### 3. GetFolderCountAll — RDO 平行 + OOM 循序

重構為三層 fallback：
- **① BFS 展開 (標準路徑)** — `GetSubFolderList + .Count`
- **② Redemption 平行 fallback** — `Parallel.ForEach + Task.WhenAll` (RDO free-threaded 安全)
- **③ OOM 循序 fallback** — 逐一遞迴 + `Await Task.Yield()`

---

### 4. Tab2 年份/月份統計

| 函數 | 舊做法 | 新做法 |
|---|---|---|
| `GetYearCountsForFolder` | 逐年 `Restrict` (~30 次 COM call) | 1 次 GetTable + GetArray，記憶體 `GroupBy Year` |
| `GetMonthCountsForYear` | 逐月 `Restrict` (12 次 COM call) | 1 次 GetTable (整年 filter) + GetArray，記憶體 `GroupBy Month` |

---

### 5. ListView1 右鍵選單

[ListView1_ItemMenu](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1.vb#L1590) 從 `GetFolderSizeLegacy()` 改為 `GetFolderSizeAll()`，使用新底層 L3 COM 函數，自動走 Redemption ⓪ → GetTable ① → Items ② 的 fallback 鏈。

---

### 驗證建議

1. **Tab1**：點選包含大量子資料夾的 PST，觀察 ListView1 統計速度
2. **Tab2**：切換到「依日期統計」，觀察年份長條圖是否正確顯示
3. **Tab3**：搜尋大資料夾的附件，確認 GetArray 批次讀取無例外
4. **右鍵選單**：在 ListView1 右鍵 → 「統計資料夾大小」，確認走 `GetFolderSizeAll` 路徑
5. **無 Redemption 測試**：暫時 comment 掉 `InitRedemptionSessionWithoutDeclaration()`，確認所有功能仍可循序運作
