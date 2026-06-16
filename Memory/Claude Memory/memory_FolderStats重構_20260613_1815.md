# Outlook Assistant — folder_stats 子樹計數「去模式化」重構交棒 spec
# 產出: 2026/06/13 by Simon/Claude | 供新 chat 直接執行實作

---

## 0. 一句話任務
把「子樹郵件/資料夾計數鏈」改成 **骨架完整且唯一、模式(`_showAllFolders`)只活在呈現層**，
徹底消滅「folder_stats 殘缺 → 子樹靜默少算」的未爆彈，並順手修掉 RDO 計數比 OOM 多算的問題。

---

## 1. Bug 現象與根因 (背景，已確認)

**現象**: SimTree1 按 F5 強刷，右側 Lv1 的「資料夾數量 / 郵件總計」每次數字不同；只有 tab6 Renew Cache 後才正確。
範例: 收件匣正確值 11 子夾 / 26,919，F5 卻顯示 6 或 7 子夾、少算。直接郵件數(郵件數量欄)永遠正確。

**根因 (兩層)**:
- **直接機制**: L2.5 的子樹清單路徑 (`GetSubtreeToList` / `DbGetSubFolderIDList`) **沒有完整性驗證**，
  `folder_stats` 殘缺時，`LIKE rootPath%` 默默回傳殘缺子樹 → 繞過 L3 完整 COM 掃描 → 少算。
  (對比: 單夾快取 `GetMailCount`/`GetFolderCount` 有 snapshot 驗證，子樹清單沒有。)
- **時間軸根因 (為何 6/4 修好、改 SQLite 後復發)**: 6/4 的「修復點1/2」(走訪時註冊身分證) 只是「把表填滿」、
  **不是結構性修復**；表填滿後洞被遮住。這幾天改 SQLite schema 時重建/刪過 DB → `folder_stats` 歸零 →
  只走過部分分支(曾出現 12 筆) → 同一個老洞再次露出。`INSERT OR REPLACE` 只增不減、兩個 CleanupOrphanPath 呼叫端
  (SaveCachesToDB / RenewCacheToDB) 06/12 後皆安全，**沒有人主動砍表**，純粹是「重建後沒被完整重走」。

**Q1 釐清 (RDO vs OOM)**: 啟用 RDO 時資料夾數(27)比 Renew(22)多 5 個 —— 是 **RDO 多算**:
`GetSubtreeToListL3_Rdo` 完全無 `is_mail`/`_showAllFolders` 過濾，且 Redemption 走 MAPI 看得到 OOM 隱藏的系統/非-IPM 夾
(Recoverable Items、Conversation Action Settings 等)。OOM 的 22 才是「使用者可見郵件夾」正確值。

---

## 2. 已定案架構 (核心原則: 骨架完整且唯一、模式只在呈現層)

1. **快取去模式化**: `_cacheSubTreeList` 的鍵 **拿掉 `|_showAllFolders`**，只存一份**完整骨架**(含非郵件夾)。
   - `_cacheFolderCount` 已是單值(未過濾 fc)，無模式問題。
   - **本次不動 `_cacheFolderTree` / `GetSortedSubFolders`** (見 §4 範圍決定)。
2. **走訪/儲存層永不剪枝**: `GetSubtreeToListL3` 移除 `If Not _showAllFolders AndAlso Not isMail Then Continue For` 剪枝，
   永遠完整列舉整棵樹、完整註冊 (`_cacheFolderIDs` + 未過濾 `fc`)。→ folder_stats 永遠完整、不分模式。
3. **完整性檢查單純化 (無模式分支)**: 因骨架本就完整，閉包檢查對每個資料夾 F 一律要求
   「集合內 F 的直屬子夾數 == `fc(F)` (未過濾)」。`fc < 0`(未知) 或對不上 → 判殘缺 → fallback L3。
   (限定 rootPath 子樹範圍以避開 `LIKE` 前綴誤匹配 sibling，例如 Inbox 誤匹配 Inbox2。)
4. **計數/顯示層做模式過濾 (剪枝移到這裡)**: 依 `_showAllFolders` 從完整骨架即時派生，0 COM:
   - 全顯(True): 全數。
   - 關閉(False): **從 root 沿 `is_mail` 的夾往下剪枝走訪計數** (碰非郵件夾不往下數)。
     符合 Q2「理論上剪枝」，但剪枝在計數層、非骨架層。
   - `GetMailCountAllL3` / `GetFolderCountAllL3` 的 ② 加總路徑改成: 先拿完整骨架 → 依模式篩 → 加總。
5. **`forceRefresh` 一路 thread**: F5/discover 跳過記憶體+DB 快取，直接走 L3 完整重掃 + 回填。
   (補完 §1 提到「完整性檢查只驗內部一致、抓不到 Outlook 新增夾」的 staleness 缺口。)
6. **修 RDO 多算**: `GetSubtreeToListL3_Rdo` 改成完整列舉骨架(不過濾)，計數層用 is_mail 過濾，與 OOM 一致。
   (隱藏/非-IPM 夾的處理: 若要與 OOM 完全一致，RDO 列舉也需比照 OOM 的可見性/IsMailFolder 判斷。實作時注意。)

### 優點
folder_stats 永遠完整(消滅未爆彈)、切模式 0 COM、完整性檢查無模式 corner case、RDO/OOM 一致、快取兩份變一份。
### 代價
過濾模式走訪會多走非郵件夾子樹 — 但 Simon 資料幾乎無此結構(Q2)，實測近乎零差異；屬一次性建骨架成本。

---

## 3. 逐步實作清單 (file / function / 動作)

> 動工前先 grep 確認所有引用點，再改鍵。每步給 Simon 看 diff 再進下一步。

### Step A — `Form1_Outlook.vb : GetSubtreeToListL3` (約 line 1997+)
- **移除剪枝**: 刪掉 BFS 迴圈內 `If Not _showAllFolders AndAlso Not isMail Then Continue For`，永遠 `result.Add` + `queue.Enqueue`。
- 保留身分證註冊 + `fc` 回填 (見 §5 已改動狀態)。
- **去模式化快取鍵**: `cacheKey = rootPath`(拿掉 `& "|" & _showAllFolders`)，line ~2014 與寫入 `_cacheSubTreeList.TryAdd` 處 (line ~2085) 一併改。

### Step B — `Form1_Outlook.vb : GetSubtreeToListL3_Rdo` (約 line 2090+)
- 維持完整列舉(本就不過濾)，確認與 OOM 一致;若要排除隱藏/非-IPM 夾以對齊 OOM 可見性，於此處或計數層處理。

### Step C — `Form1_Outlook.vb : GetSubtreeToList` (L2.5, 約 line 484+)
- 簽章加 `Optional forceRefresh As Boolean = False`。
- 鍵改為 `rootPath`(去模式化)。
- 流程: `forceRefresh=True` → 直接 ③ L3;否則 ① 記憶體命中(完整骨架) → ② DB:
  `DbGetSubFolderIDList(rootPath, isIncludeAll:=True)` 撈**未過濾全集** → `IsSubtreeComplete(...)` →
  完整才重建 COM 物件清單回傳(回傳**完整骨架**，計數層才過濾) ;殘缺 → ③ L3。

### Step D — 新增 helper `Form1_Outlook.vb : IsSubtreeComplete(rows, rootPath) As Boolean`
- 限定子樹範圍(path==root 或 startsWith root&"\")。建 byPath + 集合內 childCnt。
- 對每個 F: `fc<0` → False;`childCnt(F) <> fc(F)` → False。全過 → True。**無模式分支**。

### Step E — 新增 helper `CountSubtreeByMode` / 計數派生 (呈現層剪枝)
- 全顯: 全數。關閉: 從 root 沿 is_mail 剪枝走訪計數。供 GetMailCountAllL3/GetFolderCountAllL3 的 ② 加總用。

### Step F — `Form1_Outlook.vb : GetMailCountAllL3 / GetFolderCountAllL3` (約 line 1168 / 1360)
- 簽章各加 `Optional forceRefresh As Boolean = False`，傳入 ② 路徑的 `GetSubtreeToList(...)`。
- ② 加總改成: 拿完整骨架 → 依 `_showAllFolders` 剪枝/過濾 → 加總 (folder count = 過濾後數量, mail = 過濾後各夾 GetMailCountL3 加總)。
- ⓪/① RDO 路徑: 同樣套用模式過濾以對齊 (修 RDO 多算)。

### Step G — `Form1_MainTab12.vb : CollectFolderStatsByL3ForceRefresh` (約 line 411+)
- 4 處呼叫 (line 418/419/437/438) 加 `forceRefresh:=True`。

### Step H — 清理 / 註解
- 修掉 `Form1_SQLite2.vb` line ~499 過時註解 (它說「RenewCacheToDB 完整 COM BFS」，但 RenewCacheToDB 已重構成「不再 BFS」)。
- 補滿新邏輯註解，沿用 `' YYYY/MM/DD by Simon/Claude:` 格式。

---

## 4. 範圍決定 (本次 NOT 動)
- **`GetSortedSubFolders` + `_cacheFolderTree`**: 概念同源(也是 `|_showAllFolders` 雙鍵 + 剪枝)，但它餵的是**樹顯示**，
  去模式化要改所有樹渲染呼叫點(`Form1_SimTree.vb`/樹填充)，blast radius 大且非本次 bug。**列為後續單獨清理**。
  (子樹計數鏈不經過 `GetSortedSubFolders`，兩者機械上分離，本次不動安全。)

---

## 5. 目前 repo 已改動狀態 (本次討論中已下的兩處 edit，新 chat 接手時注意)
> 注意: 這些是在 /mnt/project 唯讀副本上的示範性 edit，實際專案需重新套用;且 Step A 會再覆蓋其中一處。
1. **`Form1_Outlook.vb : GetSubtreeToListL3` BFS 迴圈** — 已把身分證 `_cacheFolderIDs.TryAdd` 移到過濾前，
   並在 `Next` 後加 `_cacheFolderCount(current.Path) = CLng(subFolders.Count)` (回填未過濾 fc)。
   → **Step A 要在此基礎上「再移除剪枝 `Continue For`」**;fc 回填保留;身分證移到過濾前在無剪枝後變無差別但無害。
2. **`Form1_SQLite2.vb : DbGetSubFolderIDList`** — SELECT 已補 `folder_count`，row 已填 `.fc`。**保留**。

---

## 6. Simon 的 coding 約束 (務必遵守)
- 循序思考、大任務拆塊、動手前先給 diff 確認、不浪費 token。
- 有疑問先問、不亂猜;看到更簡單做法要敢推回。
- **註解**: 原有註解/Debug 思考/日期紀錄**不可莫名刪除整段**;可修正/調整/補齊;太雜亂可把鄰近的整理在一起。
- 風格: 壓縮 VB.NET(單行 If/Then、冒號分隔);繁中行內註解 + `' YYYY/MM/DD by Simon/Claude:` 屬名。
- 區域結構: `#Region` + `■/├/└` 標記須保留。
- 讀檔: 上千行勿全讀,只讀需要的行段(grep + line range)。
- COM 必須留在 UI/STA 執行緒,勿包 `Task.Run`;`System.Exception` 必須完全限定。
- 回覆用繁中 + 英文,勿用韓文/簡中。

---

## 7. 驗證方式
- F5 強刷收件匣多次 → 數字穩定且 == Renew Cache 值(過濾模式 22 子夾類)。
- 切 `chkShowAllFolders` → 全顯較大、關閉較小,且皆穩定、即時(無 COM 重掃)。
- 啟用 RDO 後 F5 → 數字與 OOM 一致(不再多算 5 個隱藏夾)。
- 重建 DB(ZipAndRebuildDB)後立即 F5 → 仍正確(完整性檢查偵測殘缺 → fallback L3 → 正確 + 回填)。
- `_iLikeNoisy` 開啟可印列舉夾名 diff,驗證 RDO/OOM 一致。
