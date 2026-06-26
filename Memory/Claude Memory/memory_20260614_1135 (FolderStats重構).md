# Outlook Assistant — folder_stats 去模式化重構 + 樹崩潰 bug：結案紀錄
# 產出: 2026/06/14 by Simon/Claude | 取代 20260613_1815 版 (該版根因分析有誤，見 §3)
# 狀態: 全部完成並經 Simon 實測驗證通過。本輪 debug 結束。

---

## 0. 一句話結論
「子樹計數鏈去模式化」重構完成且通過驗證；過程中暴露的「重啟後樹只剩收件匣」是**獨立的長期潛伏 bug**，
真因是 **`BuildBfsFolderTree` 與 `GetSortedSubFolders` 都不註冊 `_cacheFolderIDs`（修復點2 被註解）**，
**不是** 先前懷疑的 CleanupOrphanPath，也不是 `INSERT OR REPLACE` 洗 entry_id。已修復並驗證。

---

## 1. 已完成並驗證的內容 (Tab1 計數鏈去模式化)

**動機**: 消滅「folder_stats 殘缺 → 子樹靜默少算」未爆彈、修 RDO 計數比 OOM 多算。

**已落地 (檔案 / 函數)**:
- `Form1_Outlook.vb : GetSubtreeToListL3` — 移除模式剪枝 (`If Not _showAllFolders AndAlso Not isMail Then Continue For`)，
  永遠完整列舉整棵骨架；`Next` 後回填 `_cacheFolderCount(current.Path)=subFolders.Count`(未過濾 fc)；快取鍵去 `|_showAllFolders`。
- `Form1_Outlook.vb : GetSubtreeToList` (L2.5) — 鍵去模式化；加 `Optional forceRefresh`；DB lazy 改撈未過濾全集
  (`DbGetSubFolderIDList(rootPath, isIncludeAll:=True)`) → 經 `IsSubtreeComplete` 驗證骨架完整才採用，否則 fallback L3；
  DB 重建時回填 `_cacheFolderIDs` + `_cacheFolderCount`。
- `Form1_Outlook.vb` 新增 helper：
  - `IsSubtreeComplete(rows, rootPath)` — 限定子樹範圍，逐夾要求「集合內直屬子夾數 == fc(未過濾)」；fc<0 或對不上 → 殘缺。無模式分支。
  - `FilterSubtreeByMode(skeleton, rootPath)` — 計數/顯示層的模式過濾(剪枝移到這)：全顯回全集；
    過濾則從 root 沿 `is_mail` 剪枝走訪(root 一律納入)；is_mail 來源 `_cacheFolderIDs`，查無保守視為 mail(納入)。
- `GetMailCountAllL3` / `GetFolderCountAllL3` — 加 `forceRefresh`；② 路徑改「取完整骨架 → `FilterSubtreeByMode` → 加總/Count-1」；
  ⓪/① RDO 快速路徑以 `_allCountUseRdoFastPath`(預設 False)關閉，一律走 ② OOM 骨架以結構性保證與 OOM 一致(修 RDO 多算)。
- `Form1_MainTab12.vb : CollectFolderStatsByL3ForceRefresh` (F5 入口) — 4 處呼叫加 `forceRefresh:=True`。
- `Form1_SQLite2.vb : DbGetSubFolderIDList` — SELECT 補 `folder_count` → 填 `row.fc`(供 IsSubtreeComplete)。

**驗證結果 (Simon 實測)**: 連續 F5 數字穩定 == Renew Cache 值；RDO 開啟 F5 數字與 OOM 一致(不再多算隱藏夾)。

**關鍵更正(舊註解)**: `GetMailCountAllL3/GetFolderCountAllL3` 開頭「2026/4/28 已成死碼」註解**已過時** —
它們已被 `CollectFolderStatsByL3ForceRefresh` 復用，是 F5/forceRefresh 的活路徑(已在碼中標註，未刪原註解)。

---

## 2. 樹崩潰 bug 修復 (本輪主要 debug 成果)

**現象**: 空 DB 重建 → 第1次啟動樹正確(COM) → do-nothing 關閉 → 重啟後 Gmail_2022 底下**只剩收件匣**；
F5 完整重掃可暫時修好，但下次 do-nothing 關閉重啟又壞。

**真因 (數據 + grep + 原碼三方確認)**:
- `_cacheFolderIDs` 只在這些點被寫入：L3 BFS(計數鏈/F5)、`FillCacheFromDbRow`(DB 命中且 eid 非空)、
  `LoadFolderStatsInner`(啟動載入、eid 非空)、RenewCache、以及**被註解掉的【修復點2】**。
- **`BuildBfsFolderTree`(Tab1 計算 BFS) 整段不寫 `_cacheFolderIDs`**，它取子夾是呼叫 `GetSortedSubFolders`，
  而後者 ③ COM 段的【修復點2】(`_cacheFolderIDs.TryAdd(...)`) 在 2026/6/2 被註解掉。
- 結果：只在樹上顯示/被 BFS 計算過的夾，存檔時 `entry_id`/`is_mail` 寫成 NULL；只顯示沒計算的夾(收件匣的兄弟)連快取都沒進、**完全沒存**。
- 重啟時 `DbGetOrderedSubFolderIDs` 帶 `entry_id IS NOT NULL AND is_mail=1` → 上述全被濾掉/不存在 → 樹崩。
- 收件匣本身有 entry_id，是因它是預設選取夾走了別的註冊路徑。

**決定性數據 (探針 B, 重啟後展開 Gmail_2022)**: `DB直屬子夾=1 | entry_id為NULL=0`(只有收件匣在 DB)；
收件匣的子孫 `DB直屬子夾=4 | entry_id為NULL=4 | is_mail≠1=4`(存了但身分證全 NULL)。

**修法 (已落地, `Form1_Outlook.vb : GetSortedSubFolders`)**:
- 還原【修復點2】：在 ③ COM 枚舉子夾時 `_cacheFolderIDs.TryAdd(childPath, (EntryID, StoreID, isMail, hasCh))`。
- `isMail` 改為先算一次(showAll 模式也算)，供過濾與註冊共用，不重複打 COM。
- `GetSortedSubFolders` 是「樹載入」與「BuildBfsFolderTree 計算」共用的子夾枚舉樞紐 → 一處修好兩條路徑。

**互補保險 (`Form1_SQLite2.vb : SaveFolderStatsInner`)**:
- `INSERT OR REPLACE` 改 `ON CONFLICT(folder_path) DO UPDATE`：統計欄照常覆寫，身分欄(entry_id/store_id/is_mail/has_chinese)
  用 `COALESCE(excluded.x, x)` 保留 DB 既有值。作用：DB 載入的後續存檔(快取無身分證時)不會把 entry_id 回退成 NULL。
- 注意：這層**單獨無法**修好本 bug(空 DB 首存就是 NULL，COALESCE(NULL,NULL)=NULL)；真正治本是修復點2。兩者互補保留。

**驗證結果 (Simon 實測)**: 多次重啟後收件匣與其兄弟夾都正常顯示；F5、RDO 重刷數字正確完整。

---

## 3. 兩個被推翻的錯誤根因 (記取教訓, 勿再回頭追)
1. **CleanupOrphanPath 刪 row** — 錯。它從 `SaveCachesToDB` 呼叫時 livePaths 已含全部 folder_stats 路徑(06/12 保護)，
   `stalePaths` 為空、不刪任何 row。非兇手。
2. **`INSERT OR REPLACE` 整列覆寫洗 entry_id** — 方向錯。真正問題是 entry_id **從一開始就沒被寫入**(修復點2 註解 + BFS 不註冊)，
   不是寫好後被洗掉。UPSERT 改法保留為互補保險，但非治本點。

**教訓**: 修一兩次沒解決時不可過度自信認定某合理推論就是根因到處亂修；應建立多假設、要求 Simon 設檢查點、多處放探針，
取得數據後再找唯一完全吻合的可能。本 bug 正是靠探針 A/B/D 的 DebugForm 輸出(尤其「DB直屬子夾=1」)才一刀定案。

---

## 4. 本輪附帶修正：Tab2-5 過濾模式掃 333 (= 去模式化副作用)
**現象**: Tab1 過濾模式資料夾 ~307、全顯 333；但 Tab2(年度)、Tab4(系列)等在過濾模式下狀態列卻掃 333(含非郵件夾)。
**真因**: 與 §4(GetSortedSubFolders/樹顯示)**無關**。是我把 `GetSubtreeToList` 去模式化成「永遠回完整骨架」後，
Tab2-5 共用的 `GetUniqueFolderList`(多選節點展開枚舉)沒補模式過濾 → 從 ~307 變 333。
**修法 (已落地)**: `Form1_Outlook.vb : GetUniqueFolderList` 對每個選取節點的子樹套 `FilterSubtreeByMode(subTree, SafeGetPath(rootF))`。
showAll 模式回全集、行為不變；過濾模式回 mail-only ~307(還原重構前行為)。

---

## 5. 探針 (本輪用畢已全部移除)
- 探針A: `SaveFolderStatsInner` 印「無身分證 path 數 + SQL前30字 + 範例」(`[DBG-SAVE]`)。
- 探針B: `DbGetOrderedSubFolderIDs` 印「過濾後回傳 vs DB 全部直屬子夾、entry_id NULL/is_mail≠1 分布」(`[DBG-TREE]`)。
- 探針D: `GetSortedSubFolders` ① 記憶體命中印 key + 子夾數。
→ 三者皆已於 2026/06/14 移除，最終版乾淨。

---

## 6. 仍待處理 (留給 Simon 決定)
- **`GetFolderSizeAllL3`(Form1_Outlook.vb ~1643)**: 同一去模式化遺漏 — ② 路徑取完整骨架後未套 `FilterSubtreeByMode`，
  過濾模式會把非郵件夾 size 也加總；且 ① RDO 路徑同理。因「資料夾大小」欄顯示「-」疑似未使用、改它牽涉 RDO，**本輪未動**。
  待 Simon 確認 size 是否在用，要修再比照 GetMailCountAllL3 補 FilterSubtreeByMode + 視需要 gate RDO。
- **`_allCountUseRdoFastPath` 恢復**: 若日後要恢復 RDO 快速路徑，須先在 `GetSubtreeToListL3_Rdo` 比照 OOM 可見性/IsMailFolder
  過濾隱藏/非-IPM 夾，再把開關設 True。
- **§4 原列「不動」項** `GetSortedSubFolders + _cacheFolderTree` 的 `|_showAllFolders` 雙鍵去模式化：仍未做(blast radius 大、非當前 bug)，列為後續單獨清理。
- horizon 其他項：Tab4 1-H、Tab7 Phase 2/3、`Dbg()` 遷移等(沿用既有 memory)。

---

## 7. 關鍵碼路徑速查 (供下次接手)
- 樹載入: `GetSortedSubFolders` (① `_cacheFolderTree` 記憶體[仍 `|showAll` 雙鍵] → ② `DbGetOrderedSubFolderIDs`[`entry_id IS NOT NULL AND is_mail=1`] → ③ COM[修復點2 在此註冊身分證])。
- Tab1 計算: `CollectFolderStatsByBFS → BuildBfsFolderTree`(自有 BFS，取子夾呼叫 `GetSortedSubFolders`，**不直接寫 _cacheFolderIDs**，靠修復點2 在枚舉時註冊) → `FetchDirectMailCountsAsync → SummarizeSubTreeBottomUp → UpdateFolderStatsCache`。
- Tab1 F5: `CollectFolderStatsByL3ForceRefresh → GetMailCountAllL3/GetFolderCountAllL3(forceRefresh) → GetSubtreeToList(forceRefresh) → L3`。
- Tab2-5 枚舉: `GetUniqueFolderList(selectedNodes, includeSub) → GetSubtreeToList → FilterSubtreeByMode`。
- 計數鏈子樹: `GetSubtreeToList`(L2.5: 記憶體[單鍵 rootPath] → DB[IsSubtreeComplete 驗證] → `GetSubtreeToListL3`[完整 BFS, 寫 _cacheFolderIDs+fc])。
- 存檔: `Form1_FormClosing → SaveCachesToDB → SaveFolderStatsInner`(UPSERT, 身分欄 COALESCE 保留) + `CleanupOrphanPath`(livePaths 含全部 folder_stats 路徑, 不誤刪)。

---

## 8. Simon coding 約束 (續遵守)
- 循序思考、大任務拆塊、動手前先給 diff 確認(除非 Simon 指示一次全做)、不浪費 token、上千行勿全讀(grep + line range)。
- 有疑問先問、不亂猜、看到更簡單做法敢推回、多種解讀攤出來給選；只動該動的、別人的 dead code 用講的。
- 原有註解/Debug/日期紀錄不可莫名整段刪除；可修正/調整/補齊；屬名 `' YYYY/MM/DD by Simon/Claude:`。
- 壓縮 VB.NET(單行 If/冒號分隔)；繁中行內註解；`#Region` + `■/├/└` 標記保留。
- COM 留 UI/STA 執行緒勿包 `Task.Run`；`System.Exception` 完全限定。回覆繁中+英文，勿韓文/簡中。
- **修不出來時建立多假設 + 設探針取數據，勿過度自信亂修(本輪血淚)。**
