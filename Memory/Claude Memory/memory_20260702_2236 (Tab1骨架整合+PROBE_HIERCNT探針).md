# Outlook Assistant — Tab1 BFS 骨架整合 + PROBE_HIERCNT 探針 進度記錄

> 本檔記錄 2026/07/02 晚間 session：Tab1 `BuildBfsFolderTree` 骨架整合(已實作、已編譯通過、待 Simon 實機驗證)，
> 以及 hierarchy table 順手撈 PR_CONTENT_COUNT 的 parity 探針(已實作、待跑)。
> 背景脈絡見 memory_20260702_1409(ExecSQL 年月計數) 與 memory_20260629_2330(Option A1 id-tuple BFS)。

---

## 一、問題盤點結論（Simon 四問的答案，供日後查閱）

### Q: Tab1 是否有兩組 BFS？
是。整個 codebase 的子樹展開有兩套系統：
- **A. `BuildBfsFolderTree`** (Form1_MainTab12.vb) — Tab1 標準路徑專用。自己走樹：每節點一次 `DbGetOrderedSubFolderIDs` + 一次 `DbGetFolderStats`，未知節點退 COM；邊走邊做 mca/fca 快取剪枝，靠 BFS 順序產生 ParentIndex 供 `SummarizeSubTreeBottomUp`。
- **B. `GetSubtree` L2.5** (Module_Outlook.vb) — Tab2/3/4/5(經 `GetUniqueFolderList`) 與 Tab1 F5(經 `GetMailCountAllOOM`/`GetFolderCountAllOOM`)、`GetFolderSizeAllOOM` 使用。四層：①記憶體 `_cacheSubTreeList` → ②DB 一條 LIKE 撈全樹 → ③`GetSubtreeRdo` 批次(~8ms 級) → ④`GetSubtreeOOM` BFS。

### Q: F5 為什麼走 GetMailCountAllOOM 而不是「GetMailCountAllRdo」？
1. **名字是歷史遺留，實質已 RDO 化**：`GetMailCountAllOOM` 內部 = `GetSubtree`(RDO 批次展開) + 逐夾 `GetMailCount`(L2.5 派工，`_rdo2` 在就走 `GetMailCountRdo`)。真正的 OOM 只剩最後保險。
2. **「單一 RDO call 拿子樹總數」的路(TotalItemCount)是刻意封殺的**：2026/06/13 實證它含 OOM 看不到的隱藏夾、無法 is_mail 過濾，與 OOM 數字不一致；06/25 已把 `_rdoFastPath` 死分支刪除。所以「整棵子樹總數」永遠 = 骨架展開 + 逐夾加總，不存在一發 RDO API。

### Q: ExecSQL 能替代 BFS 嗎？
**不能。** Redemption `ExecSQL` 作用於單一 MAPITable(某夾 Items.MAPITable 或某夾 Folders.MAPITable 的一層)，`FROM Folder` 是固定虛擬表名；MAPI 無 store 級遞迴表，逐層走訪不可避免——而逐層批次讀正是 `GetSubtreeRdoBatch` 用 `GetRows` 在做的事。同一張表上 ExecSQL 比 GetRows 多 SQL 解析+ADODB marshal 開銷，且 SUM 會 AV、無 GROUP BY。ExecSQL 的價值僅在「夾內聚合下推」(COUNT(*) 年月統計，已上線 4~10x)。SQLite 的「一條 SQL 取代 BFS」= `GetSubtree` ② 的 LIKE，本次整合讓 Tab1 也吃到。

---

## 二、已實作 #1：Tab1 骨架整合（production 變更）

**改動檔案**：
- `Form1_MainTab12.vb` — `BuildBfsFolderTree` 重寫；`CollectFolderStatsByBFS` 傳入 progress；兩處架構註解更新。
- `Module_Outlook.vb` — 僅 `GetSortedSubFolderIDs` 頭部加退役註記（無呼叫端，暫留，穩定一輪後可刪）。

**新流程**（`BuildBfsFolderTree` 三段式）：
1. `Await GetSubtree(rootFolder, True, progress, cToken)` 一次取完整骨架（含 root、含非郵件夾）。
2. `FilterSubtreeByMode` 模式過濾 → 建 `byPath`/`childrenOf` 記憶體樹 → 每層子夾按 `(TextHasChineseChar(name), path)` 排序，對齊 `DbGetOrderedSubFolderIDs` 的 `ORDER BY has_chinese ASC, folder_path ASC`（保 UI 顯示順序不變）。
3. 在記憶體樹上跑**原樣**的 BFS + 剪枝迴圈：非 root 雙快取(mca/fca)命中才剪枝、root 永遠展開(v4 fix)、未命中節點仍逐節點 `DbGetFolderStats` + `FillCacheFromDbRow(skipAggregates:=True)`。

**刻意保留的每節點 DbGetFolderStats**：目的是預熱 `_cacheMailCount`，讓 Step2 `GetMailCount` 走記憶體命中，避免觸發「DB lazy + snapshot 驗證」的逐夾 COM 讀取。若日後 PROBE_HIERCNT 通過、mc 由 hierarchy table 順手回填，這 N 次點查詢可再撤。

**退役**：`GetSortedSubFolderIDs` + `selfKnownToDb` 冷啟動特判（07/01 修的 regression 場景，改由 GetSubtree ③RDO/④OOM 全樹掃自然覆蓋）。

**預期收益**：
- 暖重啟：S1 從 N 次 `DbGetOrderedSubFolderIDs` → 1 條 LIKE(或記憶體命中 ~0ms)。
- 冷啟動(全新 DB)：S1 從逐節點 COM 物化(~2.8ms/夾) → `GetSubtreeRdo` 批次(探針實證 ~60-150×)。
- Tab1/Tab2-5 共用 `_cacheSubTreeList`，互相預熱。

**已驗證**：VS MSBuild 編譯通過(2026/07/02 22:3x)，無新警告。
**待 Simon 實機驗證**（新 PROBE_TIMING log 格式：`骨架={n} 過濾後={n} 節點={n} | 骨架 Xms + 剪枝 Yms`）：
1. 暖重啟點選常用子樹 → 數字與舊 log 完全一致、S1 應大降。
2. 冷啟動(刪/改名 DB)點選 → 子樹展得開(07/01 regression 場景)、數字正確。
3. 過濾模式(不勾顯示全部) → 非郵件夾不出現、加總正確。
4. ESC 中斷 → 不污染快取。
5. F5 → 行為不變(F5 不經 BuildBfsFolderTree，但共用 GetSubtree skipCache)。

---

## 三、已實作 #2：PROBE_HIERCNT 探針（Form1_MainTab56.vb debug 區，整塊可刪）

**驗證目標**（memory_20260702_1409 Q1 的未實作機會）：`GetSubtreeRdoBatch` COLS 加 `PR_CONTENT_COUNT (0x36020003)` 後——
- Q1 值 parity：hierarchy table 的 PR_CONTENT_COUNT vs OOM PropertyAccessor 同屬性（=GetMailCountOOM ① 路徑）逐夾全對拍。
- Q2 成本：5欄 vs 4欄(現行 production COLS) 批次走訪時間差。
- Q3 順手驗證 fc：RDO 每層 RowCount(production 已回填 `_cacheFolderCount`) vs OOM `Folders.Count`。

**跑法**：勾 CheckRDO → Tab3(或 Tab1) 樹選 1~n 個子樹 root → Tab6 DebugButton。結果進 Debug 視窗 + MessageBox。
**注意**：hierarchy table 的列描述「子夾」，root 自身 mc 不在表上（production 由既有 GetMailCount 負責，探針不對拍 root）。
已知陷阱已防：`Convert.ToInt64(Nothing)=0` 會偽裝成空夾，探針先判 Nothing 記 -999+cntMiss。

**判讀**：mc/fc 全一致 + 5欄 overhead 可忽略 → 下一輪把 production `GetSubtreeRdoBatch` 加欄回填 `_cacheMailCount`，
Tab1 Step2(佔 55~65%) 冷啟動時大量記憶體命中，並可考慮撤掉 BuildBfsFolderTree 迴圈內的逐節點 DbGetFolderStats。

---

## 四、情境×最快路徑規劃表（Q3 的設計答案，整合後的目標狀態）

| 情境 | 骨架來源 | 本層計數來源 | 備註 |
|---|---|---|---|
| 記憶體命中(第二次點選) | `_cacheSubTreeList` ~0ms | `_cacheMailCount`/`_cacheMailCountAll` 剪枝 | 最快路徑，整合後 Tab1/Tab2 互通 |
| 暖重啟(DB 健康) | DB 一條 LIKE + IsSubtreeComplete | DB 點查詢(FillCacheFromDbRow 預熱) | S1 一次查詢；S2 記憶體命中 |
| 冷啟動(全新 DB)+RDO | `GetSubtreeRdo` 批次 | 現況: 逐夾 GetMailCountRdo；PROBE_HIERCNT 通過後: hierarchy table 順手回填 | 這格是下一步最大紅利 |
| 冷啟動+無 RDO | `GetSubtreeOOM` BFS | 逐夾 PropertyAccessor | 安全網，維持純 OOM 不動 |
| F5 強刷 | GetSubtree(skipCache) 重掃+覆寫 | 逐夾 skipCache 直讀 | 語意=繞過快取，本來就不該吃快取加速 |

## 五、結果更新（2026/07/03 凌晨）
- **PROBE_HIERCNT 全數通過**：Simon 全選 26 個 PST root（~800+ 夾，含中文名/英文名/ePaper&RSS/iCloud 帳號）一次跑完，
  `VERDICT: PASS_ALL_PARITY | roots=26 | mc值不符=0 | mc缺夾=0 | fc值不符=0 | cnt讀取失敗=0 | 5欄-4欄總額外耗時=-16ms`。
  RDO 批次 vs OOM PropertyAccessor 走樹速度差 ~50-200×；第 5 欄成本為負值(雜訊等級)=免費。結果檔: PROBE_HIERCNT_log.txt(每次執行覆寫)。
- **production 加欄已上線**（由並行 session 完成，2026/07/03 註記在 GetSubtreeRdoBatch）：COLS 加 PR_CONTENT_COUNT，
  同一次 GetRows 順手回填 `_cacheMailCount`；讀不到值(Nothing/轉型失敗)不寫快取。全樹編譯通過。
- **暖重啟實測**：Simon 回報 Tab1/Tab2 都有變快，過濾模式正常、ESC 正常。

## 六、遺留 pending
- 冷啟動(刪/改名 DB)實測待做：驗證 GetSubtreeRdoBatch 回填 mc 後，Tab1 Step2 是否大量記憶體命中(看 PROBE_TIMING S2)。
- `GetSortedSubFolderIDs` 已退役無呼叫端，穩定一輪後可刪。
- PROBE_HIERCNT 探針區塊(Form1_MainTab56.vb)已完成使命，冷啟動驗證後可整塊刪除(搜 PROBE_HIERCNT)。
- BuildBfsFolderTree 迴圈內逐節點 DbGetFolderStats：**暖重啟(DB路徑)仍需要它預熱 mc/fc/fs，保留**；
  冷啟動時 RDO 批次已回填 mc，該迴圈查到的 row=Nothing 成本極低，不需再動。
