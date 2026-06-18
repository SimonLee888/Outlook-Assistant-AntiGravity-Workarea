# Memory — Tab5 內文模糊比對 (SimHash + Jaccard) 設計決策
輸出時間：2026/06/17 04:07
適用範圍：Outlook Assistant / Tab5 重複郵件偵測 / Fuzzy 模式重構

> **本文件的權威性說明**
> 以「本文件記載的設計」為唯一基準。前期若有與此衝突的設計筆記，一律作廢。

---

## ⚠ 已作廢的錯誤決策（勿再參考）

1. **❌「body SimHash 只對 subject 分桶後 >1 的群組成員計算」** — 完全錯誤。內文模糊比對必須對「全體郵件」兩兩比對（經 size 篩選後），不可限制在同主旨群組內。
2. **❌「先做 Subject SimHash 驗證，再決定要不要上 body」（方案 A）** — 已否決。主旨短→精確比對（Exact 模式既有功能）；內文大且不見得同主旨→模糊比對。兩者是完全不同的分類，不是分支。
3. **❌「Fuzzy 模式 = 主旨前綴分桶 + subject Jaccard」（現有實作）** — 確定要被取代。新 Fuzzy = 全體 body 兩兩比對 + union-find 分群。只保留 Exact + 新內文 Fuzzy 兩種。
4. **❌「FNV-1a + splitmix64」** — 已否決。Simon 原即指定 XxHash64；FNV-1a 手刻乘法在本專案預設溢位檢查下會拋 OverflowException。改回 XxHash64 庫（`System.IO.Hashing.XxHash64.HashToUInt64`）。
5. **❌「bigram 集合 size 是免費預篩（讀 body 前）」** — 誤解。bigram_count 必須讀完 body 才能得到，它的價值是（a）同一趟 build 順手算出、（b）在 size 視窗中做 O(1) 比對收斂、（c）後續跑完全免費。首次 build 必讀全部 body，size 篩不省這個成本。
6. **❌「D9 哨兵 Long.MinValue 加進 MailItemInfo」** — 改用獨立 db 後不需要。「算過沒？」直接用 `_simHashCache.ContainsKey(EntryID)` 判斷，MailItemInfo 不加欄位。

---

## ✅ 目前確定的設計（權威基準）

### 核心演算法（D1 / D7 定案）

- **bigram**：字元相鄰兩字，打包成 `(AscW(c1) << 16) Or AscW(c2)` → `HashSet(Of Integer)`（精確、零碰撞、無溢位）。
- **SimHash**：`ComputeSimHashFromSet(bigramSet)` — 對唯一 bigram 集合投票（非逐次出現），每 bigram 用 `XxHash64.HashToUInt64` 餵 4-byte buffer。
- **最終相似度（D1）**：`BigramJaccardSimilarity(setA, setB)` — `|A∩B|/|A∪B|`，與 SimHash、size 1/T 界線數學一致（取代舊的字元集 JaccardSimilarity）。
- **比對**：既有 `GetHammingDistance(Long, Long)` 直接沿用（`PopCount(XOR)`）。

### 已交付函式（S1，貼入 Form1_MainTab345.vb 1367 行後）

```
BuildBigramSet(text) As HashSet(Of Integer)
ComputeSimHashFromSet(bigramSet) As Long
ComputeSimHash(text) As Long              ← 便利進入點，build pass 請用前兩個避免重建集合
BigramJaccardSimilarity(setA, setB) As Double
```

**⚠ 舊的 `ComputeSimHash_core.vb` 必須從專案移除**（否則 `ComputeSimHash` 重複定義，編譯錯誤）。

### 儲存：獨立 db `OLAsimhash.db`（D9 簡化）

- `OLAcache.db` 的 `ZipAndRebuildDB` 整個砍檔重建 → simhash **必須放獨立 db** 才能存活。
- **同目錄**（cacheDir 下），自己的連線 `_dbSim`，PRAGMA WAL + NORMAL。
- Schema：`mail_simhash (entry_id BLOB PRIMARY KEY, simhash INTEGER NOT NULL, bigram_count INTEGER NOT NULL)`
- **entry_id 編碼**：沿用 `HexStringToByteArray` / `ByteArrayToHexString`（與其他表一致，含 EMPTY_ 哨兵）。
- **記憶體快取**：`_simHashCache As ConcurrentDictionary(Of String, (SimHash As Long, BigramCount As Integer))`；`LoadSimHashCache()` 一次全載（每列 ~16B，量小）。
- **哨兵**：不需要。「算過沒？」用 `_simHashCache.ContainsKey(EntryID)` 判斷；MailItemInfo 不加任何欄位。
- **MailItemInfo 完全不動**（不波及 Tab4 / Tab7 / basic_maillist 序列化）。
- **清除策略（UI）**：清快取對話框新增 `□ 一併清除 SimHash 快取（預設不勾）`，勾了才呼叫 `DeleteSimDatabase()`（關閉→刪檔→重建空表）。預設清快取不碰 OLAsimhash.db。（**S7 的一部份，尚未實作**）

### 已交付 DB helper（S2，加入 Form1_SQLite2.vb）

```
InitSimDatabase()          ← InitDatabase 末段呼叫
DeleteSimDatabase()        ← ClearCache 勾選 checkbox 才呼叫（S7 接線）
LoadSimHashCache()
SaveSimHashBatch(rows)
```

連線關閉接線：CloseDatabase 中 `_db` 關閉後補一行關 `_dbSim`（S7 確認是否已接）。

### Fuzzy 比對管線（三層篩選 + union-find 分群）

**流程（D2 / D3 / D4 定案）**

1. **S3 Build pass**（`PrecomputeFuzzySimHashAsync`）：`_simHashCache` 沒有的信 → 讀 body（`GetMailBodyL3`，不走 L2.5 避免撐爆記憶體）→ `BuildBigramSet` → 算 `simhash` + `bigram_count` → 進 `_simHashCache` → 每 500 封 `SaveSimHashBatch` flush。已算過直接跳過（暖快取）。`MIN_BIGRAM_FOR_FUZZY = 5`（極短信不納入，避免雜訊群）。
2. **S4 候選產生**（`GenerateFuzzyCandidatePairs`）：bigram_count 升冪排序 → size 1/T 滑動視窗 → Hamming 一階篩（D6 起始門檻）→ 探針記錄「視窗比對數 / Hamming 過關數」。純 CPU，放 `Task.Run`。**無自適應分支**（單一路徑，simhash 一次性都建好了，popcount 又極便宜）。
3. **S5 Jaccard 精算**（`FilterCandidatesByJaccardAsync`）：只重讀候選的 body（走 `GetMailBody` L2.5，快取這少數幾封）→ `BuildBigramSet` → `BigramJaccardSimilarity` → 過門檻配對。bigram 集合一起回傳給 S6（免重讀）。
4. **S6 分群**（`BuildFuzzyGroups`）：union-find（連通分量，D4）→ G1,G2…；每群選 `bigram_count` 最大者為代表；每封顯示「對代表的 bigram Jaccard %」（代表本身 100%）。含 `Uf_Find`（路徑壓縮）/ `Uf_Union`。

S3–S6 全部已交付，放在 `Form1_MainTab345.vb` 的新 region `"  ├ Fuzzy 內文比對引擎"`。**目前無呼叫端（不影響現有功能）**，等 S7 接線。

### D6 Hamming 一階門檻（起始值，v1.1 定案）

| 使用者門檻 T | Hamming ≤ |
|---|---|
| 87% | 14 |
| 92% | 10 |
| 95% | 7 |
| 98% | 4 |

- 寧鬆勿緊（誤殺=漏真重複，不可接受）。Jaccard 才是準確閘門。
- 探針（S4 `_dbg`）記錄 Hamming 過關數 vs Jaccard 最終過關數。真機看 yield 微調。v1.1 依實測定案。

### 候選產生收斂策略（D3）

- v1：bigram_count 排序 + size 1/T 滑動視窗（精確界線，非 byte size fudge）。
- **v2.0 TODO**：若真機顯示「尺寸高度聚集、視窗退化成接近 O(n²)」，補 LSH banding（64 bit 切段分桶）。

### 分群（D4）

Union-find 連通分量。A~B~C 同群（適合轉寄/回覆鏈）。漂移風險由高門檻壓制；v1.1 若真機出現巨群可加「群內最大跨度」護欄。Clique 嚴格群因 NP-hard 且分裂轉寄鏈而否決。

### 相似度門檻 UI（D10，v1 時實作）

- 四個選項：87% / 92% / 95% / 98%
- 控制項樣式（下拉 vs radio）留到 S7/S8 實作時再定，功能不影響核心。

### RenderLv5Group Fuzzy 分支（D8，S7 待實作）

- 舊的 [預留] 註解（描述的是「主旨分桶的 drop-in 過濾」）已過時，S7 時直接移除（Simon 確認：模糊內文是完整獨立分類，不是主旨搜尋的分支）。
- Fuzzy 分支改為消費 `scoreMap`（BuildFuzzyGroups 回傳）顯示 body Jaccard %，不在同步渲染函式裡讀 body。

---

## 已核對的 codebase 事實（2026/06/17，以檔案區最新為準）

- **`MailItemInfo`**（`Form1_Outlook.vb:80`）：`EntryID, Subject, Size, RcvTime, SenderName, AttachCount, FolderPath, MsgIDhash, SenderEmail`。**維持原樣，不加欄位。**
- **`GetHammingDistance(Long, Long) As Integer`**（`Form1_MainTab345.vb:1364`）：`PopCount`，直接沿用。
- **`JaccardSimilarity(strA, strB) As Double`**（`Form1_MainTab345.vb:1313`）：字元集 Jaccard，Fuzzy 模式**改用 BigramJaccardSimilarity**（舊函式留著、不刪，Exact 模式仍用它）。
- **`GetMailBody(entryID) As String`**（L2.5，`Form1_Outlook.vb:903`）；**`GetMailBodyL3(entryID)`**（L3 直取，`Form1_Outlook.vb:2077`）。Build pass(S3) 用 L3；候選精算(S5) 用 L2.5。
- **`RenderLv5Group(groupDict, isExact) As (GroupCount, MailCount)`**（`Form1_MainTab345.vb:1521`）：同步函式，被排序/刪除重渲染重複呼叫 → **嚴禁在此讀 body**。
- **`Bt5_Click`**（`Form1_MainTab345.vb:1373`）：body 預算/SimHash 預算插入點在 `ScanMailsToGroupDictAsync`(1399) 與 `RenderLv5Group`(1402) 之間（僅 Fuzzy）。
- **`ZipAndRebuildDB`**（`Form1_SQLite2.vb:348`）：`IO.File.Delete(_dbPath)` 砍整個 OLAcache.db → 必須用獨立檔 OLAsimhash.db。
- **`HexStringToByteArray` / `ByteArrayToHexString`**（`Form1_SQLite2.vb:404/419`）：EMPTY_ 哨兵 UTF-8、一般 hex 用 `Convert.FromHexString/ToHexString`。
- **`ClearCache_Click`**（`Form1_MainTab345.vb:2098`）：需加 checkbox 接線（S7）。
- **`InitDatabase`** / **`CloseDatabase`**：需各加一行呼叫 `InitSimDatabase` / 關 `_dbSim`（S7 確認）。
- **`ComputeSimHash_core.vb`**（專案中存在）：**必須從專案移除**（否則 ComputeSimHash 重複定義）。Simon 尚未移除，S7 開始前先確認。

---

## 下一輪要做的事（S7）

1. **確認 `ComputeSimHash_core.vb` 已從專案移除**（S1 前置，可能尚未做）。
2. **`InitDatabase` 末段加一行 `InitSimDatabase()`**（精準 diff）。
3. **`CloseDatabase` 加關 `_dbSim`**（精準 diff）。
4. **改寫 `Bt5_Click` Fuzzy 分支**：`ScanMailsToGroupDictAsync` 之前插入 `PrecomputeFuzzySimHashAsync`（S3）；`RenderLv5Group` 之前串入 S4→S5→S6，得到 `groupDict + scoreMap`。Exact 分支完全不動。
5. **改寫 `RenderLv5Group` Fuzzy 分支**：接受 `scoreMap`（或透過共享狀態），顯示 body Jaccard %；移除過時 [預留] 註解。（先讀現行 Fuzzy 分支 1555-1564 再做 diff）
6. **`ClearCache_Click` 對話框加 checkbox**：`□ 一併清除 SimHash 快取（下次比對需重新讀取全部內文）`，預設未勾；勾了在重置 SSD 時呼叫 `DeleteSimDatabase()`。
7. **Tab5 UI 加相似度門檻控制項**（S8，可與 S7 合併做）：87/92/95/98，控制項樣式開新對話再定。

---

## 工作守則提醒

- 不擅自修改既有可運作程式碼，動到他人函式先問。大改更要謹慎、逐步、勿一次大包。
- 只改該改的；保留所有既有註解 / `_dbg` / 日期作者標記。
- 死碼用說的，不自己刪。Dead code 讓 Simon 決定。
- 註解格式 `YYYY/MM/DD by Simon/Claude+模型版本:`（前導零、無逗號），繁中為主。
- 壓縮 VB.NET 風格（單行 If、冒號分隔）；不用 XML doc 註解（`<summary>` 除外若原本就有）。
- COM 留在 UI STA 執行緒，勿包進 `Task.Run`；`System.Exception` 全名。
- 輸出語言：繁體中文 + 英文，禁韓文與簡體。
- 讀大型檔案要分段讀（grep 定位行號、再 view 小範圍），不一次全讀。
- 困難 bug：若兩次未解，停下來建多假設、放探針、收集線索，找唯一符合的根因再修。
