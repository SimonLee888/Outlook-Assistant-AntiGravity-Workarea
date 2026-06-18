# Memory — Tab5 內文模糊比對 (SimHash + Jaccard) 設計決策
輸出時間：2026/06/16 09:54
適用範圍：Outlook Assistant / Tab5 重複郵件偵測 / Fuzzy 模式重構

> **本文件的權威性說明**
> 本輪對話經歷多次方向修正。**以「目前這份文件記載的設計」為唯一基準**。
> 前期（早於本文件）若有與此衝突的設計筆記，一律作廢，見下方「⚠ 已作廢的錯誤決策」。

---

## ⚠ 已作廢的錯誤決策（前期方向錯誤，勿再參考）

1. **❌「body SimHash 只對 subject 分桶後 >1 的群組成員計算」** — 完全錯誤。
   這違背內文比對的本意。內文模糊比對的價值正是找出「**主旨不同但內容相似**」的信（轉寄、回覆、改標題重發）。
   **正解：內文模糊比對必須對「全體郵件」兩兩比對（經 size 篩選後），不可限制在同主旨群組內。**

2. **❌「先做 Subject SimHash 驗證，再決定要不要上 body」（方案 A）** — 已否決。
   Simon 明確指出：主旨「短又明確」→ 直接做精確比對即可判斷是否同一封信（這是 Exact 模式既有功能）；
   內文「龐大又不見得同主旨」→ 才需要模糊比對。**兩者是完全不同的分類，不是分支。** 直接做 body（方案 B）。

3. **❌「Fuzzy 模式 = 主旨前綴分桶 + subject Jaccard」（現有實作）** — 確定要被取代。
   新 Fuzzy 模式 = 全體 body 兩兩比對 + 分群。**不保留主旨模糊模式**（Simon 確認：只要 Exact + 新內文 Fuzzy 兩種）。

4. **⚠ 概念澄清：SimHash 是「前期篩選器」，不是「最終相似度計算」。**
   SimHash 的 Hamming 距離只是**粗略估計**（64 bit 量化噪音 + 估計的是特徵向量夾角，非 Jaccard）。
   **要顯示給使用者的準確相似度 % 必須由 Jaccard 算出**，SimHash 只負責快速淘汰明顯不相似的配對。

5. **⚠ 概念澄清：FNV-1a「只算兩個字」的誤解。**
   SimHash 以 bigram（相鄰兩字）為特徵。一封 1000 字的信約 999 個 bigram，**每個 bigram（2 字）呼叫一次 FNV-1a**，整封信呼叫約 999 次。
   「2 字」指單次呼叫的輸入大小，不是總量。FNV-1a 贏 XxHash 的原因：每次只餵 2 字這種極短輸入，XxHash 的固定 setup/finalization 成本反成累贅；FNV-1a 對極短輸入近乎零成本且零記憶體配置。

---

## ✅ 目前確定的設計（權威基準）

### 核心演算法
- **特徵**：字元 bigram（相鄰兩字）。
- **每 bigram 的 hash**：**FNV-1a 64-bit + splitmix64 收尾混合**（Simon 同意 FNV-1a；splitmix64 由 Claude 建議，用來補 FNV-1a 較弱的位元雪崩，因 SimHash 需各 bit 獨立均勻）。
- **SimHash 函式命名**：`GetSimHash(text As String) As Long`（Simon 選定 `GetSimHash`；回傳 `Long` 以對齊既有 `GetHammingDistance(Long, Long)` 並可直接存 SQLite INTEGER）。
- **比對**：既有 `GetHammingDistance` 已用 `BitOperations.PopCount(hash1 Xor hash2)`，直接沿用。
- 核心函式已寫好，存於 `GetSimHash_core.vb`（含 `GetSimHash` + `SimHashMixBigram` + 兩個 FNV 常數）。插入位置：`Form1_MainTab345.vb` 第 1367 行後（`GetHammingDistance` 之後、`#End Region` 之前）。

### Fuzzy 模式比對管線（三層篩選 + 分群）
1. **Size 篩選（強制、免費）**：每封信只與「size 在容許範圍內」的信比對。
   - 精確界線（bigram 集合大小比）：相似度 T → 大小比上限 = 1/T。
     - 87% → 1.149×、92% → 1.087×、95% → 1.053×、98% → 1.020×（**比直覺的 2× 嚴格很多**）。
   - byte size ≠ 集合大小（重複內容會塌縮），故 **byte size 第一波粗篩用 `1/T × 1.5` 寬鬆窗口**（安全邊際，避免誤殺）；
     真正建出 bigram 集合後，再套**精確集合界線 1/T** 當免費早期退出。
2. **SimHash 篩選（快、近似、若快取則免費）**：size 範圍內用 Hamming 距離做 O(1) 配對篩選，淘汰明顯不相似者。
3. **Jaccard 精算（準、較貴、顯示值）**：只對通過①②的少數候選配對讀 body、算精確 body Jaccard。
   - **相似度欄位改顯示 body Jaccard %**（取代現在的 subject Jaccard）。Simon 第 5 點確認。
4. **分群**：通過 Jaccard 門檻的配對用 **union-find 分群**，產生 ListView5 的群組（G1, G2…）。

### 順序取捨結論（三選一的答案）
- 採用「**第三種（size → SimHash 配對篩選 → Jaccard 精算）**」為骨架，
  加「**第一種的自適應**」：若某 size 桶很小（< ~100 封），跳過 SimHash 直接 Jaccard 全比。
- **否決「第二種」（純 SimHash 分組不做 Jaccard）當主模式**：給不出準確 %。可列為未來的「極速粗略模式」選項。

### SSD 快取策略（Simon 構想，已採納, 需進行最終整合優化）
- **`basic_maillist` 新增 `simhash INTEGER` 欄位**（8 bytes/封，10 萬封僅 800KB）。
- 郵件內容不變 → **快取永久有效**，新進郵件套既有 snap 增量機制補算。
- **誠實成本**：
  - 算 SimHash 必須**先讀一次 body**（開信，昂貴），此為一次性主成本。只有在必要時才去讀mailbody運算simHash，候選郵件數量小於閾值時直接讀回運算Jaccard比較划算。
  - 最終 Jaccard 仍需讀**候選信**的 body（SimHash 不可逆，無法還原內容）。候選少 → 很快。
  - 只要是運算Jaccard時讀回的mailbody, 就可以用simHash存入SSD快取, 供下次lazy load讀回直接使用。
  - SimHash和Jaccard分別讀一次昂貴的COM, 取回mailbody, 其實是重複的浪費行為, 如何優化尚未有結論需要再進一步研究。
- 比對時序：size（免費）+ simhash（免費）→ 瞬間篩出候選 → 只讀幾十封候選 body 做 Jaccard → 秒級。
- simhash若已存入SSD快取, 使用lazy load讀回就是免費, 但若是為了用simHash分桶而浪費讀COM的開銷卻是得不償失。

### 閾值設定
- 使用者可選的相似度門檻：**87% / 92% / 95% / 98%**（需在 Tab5 UI 加下拉或選項控制項）。
- SimHash Hamming 門檻：**尚未最終定案**（前期討論 ≤3≈95%、≤5≈92%，但須注意這是線性近似；真實對應是 cosine。實作時 SimHash 門檻要比 Jaccard 目標**放寬**，避免一階就誤殺，二階 Jaccard 才是準確閘門）。

### 規模與可接受時間
- 兩個目標資料夾郵件數：**數千 ~ 10 萬封**。
- Size 排序 + 滑動視窗可把全體兩兩比對從 O(n²) 降為約 O(n × 視窗大小)。
- 一次性 SimHash build（讀全部 body）：數千封數秒~數分；10 萬封視 COM/RDO 速度可能數分以上（主瓶頸）。
- 每次比對（已快取 size+simhash）：秒級。

---

## 已核對的 codebase 事實（2026/06/16，以檔案區最新為準）

- **`MailItemInfo`**（`Form1_Outlook.vb:80`）欄位：`EntryID, Subject, Size, RcvTime, SenderName, AttachCount, FolderPath, MsgIDhash, SenderEmail`。
  - 注意：是 `RcvTime`（非 ReceivedTime）、`MsgIDhash`（已是 xxHash64 hex，非 MessageID）。
- **`GetHammingDistance(hash1 As Long, hash2 As Long) As Integer`** 已存在於 `Form1_MainTab345.vb:1364`，用 `System.Numerics.BitOperations.PopCount`。
- **`JaccardSimilarity(strA, strB) As Double`** 已存在於 `Form1_MainTab345.vb:1313`（字元集 Jaccard，Charles Wu 算法現代化版）。
- **`GetMailBody(entryID) As String`**（`Form1_Outlook.vb:903`）= L2.5 快取代理；`_lv4BodyCache`（`Form1_Outlook.vb:71`，ConcurrentDictionary，session 級）；L3 為 `GetMailBodyL3`（`Form1_Outlook.vb:2077`）。
- **`RenderLv5Group(groupDict, isExact) As (GroupCount, MailCount)`**（`Form1_MainTab345.vb:1521`）是**同步函式**，被多處重複呼叫（排序 1432、刪除重渲染 1667/1876/1907）→ **嚴禁在此讀 body**（會凍結 UI）。`[預留]` 佔位在 1566-1567。
- **`Bt5_Click`**（`Form1_MainTab345.vb:1373`）流程：`GetUniqueFolderList`(1396) → `ScanMailsToGroupDictAsync`(1399) → `RenderLv5Group`(1402)。body 預讀/SimHash 預算的插入點在 **1399 與 1402 之間**（僅 Fuzzy）。
- **`ScanMailsToGroupDictAsync`**（`Form1_MainTab345.vb` ~1436）+ **`BuildMailGroupKey`**（~1484）：現有 subject 分桶邏輯，Fuzzy 分支將被取代。
- **既有雜湊輔助**：`StringToXxHash64(String) As Long`、`StringToXxHash64Hex`、`FolderPathToHash64`（`Form1_SQLite2.vb:435-456`，用 `System.IO.Hashing.XxHash64`）。
- **`basic_maillist` schema**（`Form1_SQLite2.vb` CREATE ~249）：用 `folder_hash INTEGER`、`received_time INTEGER`、`msgid_hash BLOB`、`sender_id` 外鍵；migration ALTER 區塊在 ~144。
- Simon 提到「SQLite 裡已實作 FNV-1a 但未使用」——下一輪可確認位置，評估是否複用。

---

## 下一輪要先做的事（實作前）

1. **把大改拆成小步驟計畫給 Simon 過目**（Simon 要求：大改要小心謹慎，逐步實作，勿一次塞一大包）。
   建議步驟順序：
   - (a) 插入 `GetSimHash` 核心（低風險，已備好）。
   - (b) `basic_maillist` 加 `simhash` 欄位 + migration。
   - (c) body 讀取 + SimHash 預算的非同步預處理（寫入快取）。
   - (d) size 排序/分桶 + 滑動視窗候選產生。
   - (e) SimHash Hamming 篩選 + Jaccard 精算（取代 Fuzzy 分支）。
   - (f) union-find 分群 + RenderLv5Group Fuzzy 分支改寫（顯示 body Jaccard %）。
   - (g) Tab5 UI 加相似度門檻選項（87/92/95/98）。
2. **待定案**：SimHash Hamming 一階門檻的具體值（建議用「目標 T 對應的寬鬆 Hamming + 安全餘量」推導）。
3. **待定案**：是否一次性對全範圍（數萬~10萬）預建 SimHash 快取，或惰性（首次比對該範圍時建）。傾向惰性 + 永久快取。
4. **待確認**：union-find 分群的「相似」傳遞性處理（A~B、B~C 但 A≁C 是否同群）。

---

## 工作守則提醒（Simon 既有原則）
- 不擅自修改既有可運作程式碼，動到他人函式先問。大改更要謹慎、逐步、勿一次大包。
- 只改該改的；保留所有既有註解 / `_dbg` / 日期作者標記。
- 註解格式 `YYYY/MM/DD by Simon/Claude+模型版本:`（前導零、無逗號），繁中為主。
- 壓縮 VB.NET 風格（單行 If、冒號分隔）；不用 XML doc 註解。
- COM 留在 UI STA 執行緒，勿包進 `Task.Run`；`System.Exception` 全名（避免 Outlook interop 命名衝突）。
- 輸出語言：繁體中文 + 英文，禁韓文與簡體。
