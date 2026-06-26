# memory_20260622_1846 — RDO 獨立 session 與平行讀取（定案版）

## 本檔性質與閱讀方式（請先讀這段）

本檔是 `memory_20260621_1631（RDO平行讀取研究）` 的**接續定案版**。昨天那份標為「本輪結論未定案」，今天用一連串新 spike 把所有證據重驗，**多條昨天的待驗證推論已被實測推翻**。

沿用昨天的紀律，內容嚴格分四類，**不可混為一談**：

- **【實測事實】**：來自實際執行的 debug 輸出截圖，客觀證據。
- **【確認推論】**：經今天實測證實、可當定論引用。
- **【推翻推論】**：昨天或今天曾提出、今天被實測推翻者，**不可再引用**。
- **【未定推論】**：尚未直接量測，下個對話需驗證者，**不得當定論**。

> 一句話總結今天的最大翻案：**真正的加速槓桿不是「平行」，是「獨立 session vs 共用 `_rdo`」。** 單執行緒下，獨立 session 讀附件快 ~10×、讀內文快 ~2×；平行只是其次（內文 ~1.5×、附件幾乎 0）。

---

## 一、今天的測試順序（時間軸）

1. **P1+P2（`SpikeProfileAndStoreVisibility`）** — 探 live profile 名 + 驗獨立 session 可見 store。先誤在 Gmail profile 跑（假陽性），後在 Work profile 重跑。
2. **P3 第一版（`SpikeParallelReadBenchmark`，單參數 resolve）** — 平行讀取量測，**100% resolve 失敗、空轉**。
3. **P4（`SpikeResolveMethodFinder`）** — 三種 resolve 形式 × 同/跨 session 探測，找出 P3 失敗根因。
4. **P3 修正版（改 (c) store-scoped resolve）** — 拿到第一份真平行數據（附件 U 型、內文單調上升）。
5. **P3 + A 計量探針** — 加開 store 計時、實際讀取量、worker 重疊度，釘死 U 型謎團。
6. **B 附件版（`SpikeResolveFormCompare`）** — 同一批信讀三種形式，分離「resolve 形式」與「session 種類」兩變數。
7. **B 內文版（`SpikeBodyResolveCompare`）** — 同上但讀內文；中途因 OOM/IRM + UI 緒同步兩度凍結，改背景緒後成功。

---

## 二、環境關鍵事實（今天新確認）

- **有兩個 MAPI profile：**
  - `Gmail` profile：5 個 store，全在 `\Personal\`（Personal-1、Personal-2、Gmail_2022、寄件備份(個人)）＋ `過期ePaper`。
  - `Work` profile：26 個 store，全在 `\Work\`（27 個工作 PST）＋ iCloud（無檔）＋ `過期ePaper`。
  - 兩 profile 幾乎不相交，唯一交集是 `過期ePaper`。**昨天「獨立 session 只看到 5 個」的矛盾，根因就是這個 — 昨天 Outlook 跑 Gmail profile，今天 P1 也誤在 Gmail 跑。**
- `_rdo.ProfileName`（屬性名 **`ProfileName`**，非 `Profile`）回傳 live profile 名字串（實測得 `[Gmail]`、`[Work]`）。
- 獨立 session = `New RDOSession()` + `Logon(profileName, "", False, True)`。登入**正確 profile 名**時看得到該 profile 全部 store。
- 信件母體：約 **30 萬封**，其中 **84000 封有附件（約 28%）**。
- Redemption 版本 6.7.0.6412；`_rdo` 的 `MAPIOBJECT = _olNS.MAPIOBJECT`（即共用正在跑的 Outlook OOM session）。
- 內文 production 路徑 `GetMailBodyL3`（`Module_Outlook.vb` 第 2190 行）目前走 **OOM**（`_olNS.GetItemFromID` → `MailItem.Body`），**不是 RDO**。

---

## 三、各次實測的【實測事實】

### 3-1. P1+P2 on Gmail profile（誤跑，假陽性）
- `ProfileName=[Gmail]`；OOM 有檔路徑 store=5；獨立 session 可見=5；missing=0 → 印「主路線可行」。
- **此為假陽性**：只比到 Gmail 的 5 個小 store，完全沒測到 Work PST。

### 3-2. P1+P2 on Work profile（有效）
- `ProfileName=[Work]`；OOM 有檔路徑 store=**25**（iCloud 無檔被排除）；獨立 session 可見=**26**；**missing=0**。
- → 獨立 session 登 Work profile，看得到 OOM 全部有檔路徑 store。RDO 的 `PstPath` 與 OOM 的 `FilePath` 同格式（`D:\Users\Simon\Documents\Outlook 檔案\Work\...pst`），可當比對鍵。

### 3-3. P3 第一版（單參數 GetMessageFromID，N=1000，M=4）
- 收集階段成功（各 PST 收 6000 EntryID）。
- 量測階段：每個 (workload,K) 皆 `resolve失敗=4000/4000`、讀 0 封、`wall=0.1s`（4000 次呼叫瞬間回 Nothing）。
- Gmail 與 Work 兩次皆同。**量測迴圈空轉，無有效平行數據。**

### 3-4. P4（resolve 方法探測，Work profile）
- 取樣 PST=`Inbox_2009_Wistron_bak`，10 個 EntryID，`storeEid` 長度 380。
- **同 S1（同一 session 內反查）**：(a)單參數=10/10、(b)雙參數=10/10、(c)store-scoped=10/10。
- **跨 S2（新 session 用 S1 的 EntryID 反查）**：(a)單參數=**9/10** 且報 `MAPI_E_UNKNOWN_ENTRYID`；(b)雙參數=**10/10**；(c)store-scoped=**10/10**。

### 3-5. P3 修正版（改 (c) store-scoped，N=2000，M=4，Work）— 第一次拿到真數據
| workload | K=1 | K=2 | K=4 |
|---|---|---|---|
| 附件檔名 | 3486 封/s | 2267 | 2786 |
| 內文 | 239 封/s | 286 | 357 |
- `resolve失敗=0`。內文單調上升；附件呈 U 型（K=2 最低）。
- 浮現矛盾：附件 K=1=3486 vs production `_1/_2` 僅 200 多 → **25× 矛盾**。

### 3-6. P3 + A 計量探針（N=2000，M=4，Work）
| workload | K=1 | K=2 | K=4 |
|---|---|---|---|
| 附件吞吐 | 5589 | 4675 | 5569 封/s |
| 　有附件信 | 4441/8000 | 5240/8000 | 5108/8000 |
| 　開 store 均 | 34ms | 25ms | 10ms |
| 　重疊 | 1.00x | 1.91x | 3.36x |
| 內文吞吐 | 311 | 362 | 460 封/s |
| 　重疊 | 1.00x | 1.92x | 3.45x |
- CPU 全程 10~15%（OOM 原本 5~6%、RdoPreload 9~10%）；SSD 讀取 <5MB/s。**資源未打滿。**

### 3-7. B 附件版（`SpikeResolveFormCompare`，同批信讀三形式，單執行緒，Work）
| N | (1) 共用_rdo 單參數 | (2) 共用_rdo store-scoped | (3) 獨立session store-scoped |
|---|---|---|---|
| 2000 | 1135 封/s | 1178 | **11466**（附件301一致） |
| 6000 | 1079 | 1016 | **13946**（附件648一致） |
| 2000（複驗） | 911 | 875 | **9809**（附件301一致） |
- 三形式附件數一致（公平）。三次複驗結論穩定：(1)≈(2)；(3) 約 (2) 的 **10 倍**。

### 3-8. B 內文版（`SpikeBodyResolveCompare`，N=1000，Work）
- 原設計三條含 (1) OOM `.Body`，**兩度整個程式凍結**：
  - 第一次：OOM `.Body` 撞漏網 IRM 信、跳隱形授權 modal（已知地雷）。
  - 砍掉 OOM 後仍凍：**真因是「全程 UI 緒同步」**——收集階段逐封讀 `MessageClass` 過濾 IRM 花 **19.9 秒**，UI 緒被佔死、`_dbg` timer 無法刷新（看似當掉，實為在慢跑）。
  - 改成背景緒（`Task.Run`）後正常。
- 收集：取樣 1000 封（跨 2 個 PST），跳過疑似 IRM **7144 封**。
- 結果（只剩 (2)(3)）：(2) 共用_rdo `.Body`=**285 封/s**；(3) 獨立 session `.Body`=**578 封/s**；字元數 4165945 一致。**(2)→(3) 約 2×。**

---

## 四、【確認推論】（已實測，可當定論）

1. **最大槓桿 = 獨立 session vs 共用 `_rdo`（單執行緒即達成）**：
   - 附件 ~**10×**（共用 ~1000 → 獨立 ~12000+ 封/s）。
   - 內文 ~**2×**（共用 285 → 獨立 578 封/s）。
   - 倍率差異來自負載性質：附件每封小、瓶頸在共用 session 的 per-call 開銷，獨立 session 幾乎全解放；內文每封要搬整個 Body（I/O 重），獨立 session 救不了 I/O，只改善 per-call 那層。
2. **跨 session resolve 必須帶 store 資訊**：單參數 `GetMessageFromID(eid)` 跨 session 不穩（`MAPI_E_UNKNOWN_ENTRYID`）；(b) 雙參數 `GetMessageFromID(eid, storeEid)` 與 (c) `store.GetMessageFromID(eid)`（先 `FindStoreByPath` 開 store）跨 session 皆 10/10。
3. **平行（獨立 session 上再加 K=1/2/4）**：worker 時間軸**真實重疊**（K=2≈1.9x、K=4≈3.4x）。
   - 內文：吞吐隨 K 單調上升，K=4 約 **1.5×**（I/O bound，有等待可被重疊）。
   - 附件：重疊做到 3.36x 卻**零吞吐增益**（撞 per-call COM 天花板，無 I/O 等待可重疊）。平行對附件無益。
4. **獨立 session 必須登正確 profile**：`Logon(ProfileName, NewSession:=True)`。登錯（如預設 Gmail）只看到該 profile 的 store，看不到 Work PST。
5. **resolve 形式（單參數 vs store-scoped）在同一 session 上對吞吐無顯著差異**（B 附件 (1)≈(2)）。
6. **無撞鎖問題**：獨立 session 與正在跑的 Outlook 並存讀同檔，無鎖錯（昨天 v2 讀 31103 封已證，今天全程未再現鎖錯）。

---

## 五、【推翻推論】（曾提出、今天被推翻，不可再引用）

- ❌ 昨天「獨立 session 看不到 22 個 Work PST = 它們未寫 profile / process-only mounting」→ **實為兩個 profile**；Work profile 內含全部 26 個 PST（3-2 missing=0）。
- ❌ 昨天「冷 session 慢 3 倍 / CP 值低 / 建議 `_3` 退役」→ 該數據建立在 98.8% resolve 失敗的 run 上，**無效**。
- ❌ 今天「附件 workload 在空轉（取樣多為無附件信）是 U 型主因」→ **推翻**：有附件信達 55~65%，真的在讀。
- ❌ 今天「(c) store-scoped 的開 store 成本拖累 K=2/4」→ **推翻且方向相反**：K 越大開 store 越快（34>25>10ms）。
- ❌ 今天「換 resolve 形式（單參數→store-scoped）是 25× 的槓桿，production 換了就能 200→上千」→ **推翻**：(1)≈(2)，resolve 形式無關；槓桿是 session 種類。
- ❌ 今天「內文走獨立 session 也有 ~10×」（從附件外推）→ **推翻**：內文只 ~2×。
- ❌ 今天「B 內文版凍結是 OOM `.Body` 區塊造成」→ **推翻**：砍 OOM 仍凍；真因是收集階段在 UI 緒同步跑 19.9 秒。

---

## 六、【未定推論】（尚未直接量測，下個對話需驗證）

- 【未定】共用 `_rdo` 為何慢一個量級。推測：共用的是 Outlook live MAPI session，每次 COM call 跨行程 marshal 到 OUTLOOK.EXE 並與 Outlook 自身活動競爭同一 session；獨立 session 為本行程乾淨 MAPI session。**未直接分離量測「跨行程成本」與「session 競爭成本」。**
- 【未定】內文的 resolve 形式單獨效應（單參數 vs store-scoped）。B 內文版只測了 (2)(3)，未做內文版的 (1) 單參數對照（且 (1) 用 OOM 會撞 IRM）。
- 【未定】獨立 session + 雙參數 (b) 的**吞吐**未直接量測（P4 只驗了「能不能解開」=10/10，沒量速度）。B 用的是 (c)。production 若選 (b) 走無分組扁平平行，其吞吐是否等同 (c) 未知（推測相近，因槓桿在 session 種類）。
- 【未定】平行 K>4 的上限。同 profile K 條獨立 session ＋ 1 條 Outlook，幾條會 logon 失敗未測；內文吞吐天花板（K=6/8 能否再升）未測。CPU/SSD 未打滿，暗示仍有空間，但也可能被序列化限制。
- 【未定】U 型 / K=2 最低的完整解釋。推測為「每 config 讀不同冷 block，附件密度不對等 ＋ wall 太短（1.4~1.7s）使雜訊蓋過訊號」，未完全釘死（但附件平行無益的結論不受影響）。
- 【未定】IRM 信的確切 `MessageClass` 字串。目前過濾用「保守只收 `IPM.Note`、排除含 `rpmsg`/`SMIME`」，跳過 7144 封；此過濾是否精確未驗（OOM 讀 `.Body` 撞 IRM 必凍，故 production 內文走 RDO 才安全）。

---

## 七、下一步工作方向清單（帶到下個對話）

> 平行的問題已答完（內文值得、附件不值得）。真正高 CP 值的是「**共用 `_rdo` → 獨立 session**」這個單執行緒就 10×/2× 的改造。以下依優先序。

### A. Production 改造主線（最高價值，但動工作程式碼，建議獨立一輪）
1. **把 Tab3（附件）/ Tab5（內文）的 RDO 讀取從共用 `_rdo` 改成獨立 session**（`Logon(ProfileName, NewSession:=True)`）。預期附件 ~10×、內文 ~2×。這是整輪最重要的產出。
2. **resolve 形式抉擇（production）**：
   - (b) 雙參數 `GetMessageFromID(eid, storeEid)`：無需分組、可扁平平行；但需每封的 storeEid。
   - (c) store-scoped：需先把 `sourceList` 按 PST 分組、每 PST 開一次 store；同 PST 信須進同一 worker。
   - 傾向 (b)（架構簡單、切分自由），但 (b) 吞吐未直接量測，改前先補一支對照。
3. **內文 production 路徑要從 OOM 改 RDO**：`GetMailBodyL3`（第 2190 行）目前走 OOM，須改走獨立 session RDO `.Body`（順帶免疫 IRM 卡死）。
4. **profile 偵測**：production 需動態取 `_rdo.ProfileName` 餵 `Logon`，不可寫死（使用者可能跑 Gmail 或 Work profile）。
5. **DisplayName → PstPath 對照表**：若走 (c)，`MailItemInfo.FolderPath`（`\\StoreDisplayName\...`）需轉成實體 `PstPath` 才能 `FindStoreByPath`；啟動時掃一次 session 建 map。
6. **獨立 session 生命週期 / 清理**：常駐 vs 用完即關。memory 舊紀錄：RDO 持久 session 會讓 Outlook 關不乾淨。這是 production 化已知坑，須單獨設計（建議用完即 `Logon`→`Logoff`→`TryMarshalRelease`，如各 spike 的 Finally）。

### B. 改造前可先補的小驗證（低成本、補完未定推論）
7. 量測「獨立 session + 雙參數 (b)」的吞吐，確認與 (c) 相近（決定 production 用 (b) 還是 (c)）。
8. 平行 K=6/8：測內文吞吐天花板 + 同 profile session logon 上限。
9. （可選）直接量測共用 `_rdo` 慢的成因（跨行程 vs 競爭），但對決策非必要——事實已足夠。

### C. 善後 / 清理
10. **測試碼 Simon 自行清理**（已言明）。本輪新增的拋棄式 spike（供清理清單參考，勿當 production）：
    - `SpikeProfileAndStoreVisibility`、`SpikeParallelReadBenchmark`、`SpikeResolveMethodFinder`、`SpikeResolveFormCompare`、`SpikeBodyResolveCompare`
    - 輔助：`FindStoreByPath`、`DumpResolve`（注意 `FindStoreByPath` 被多支 spike 共用，清理時留意相依）
    - 昨天遺留：`SpikeRdoIndependentSession`、`SpikeDumpStoreIdentity`、`SpikeMountWorkPst`、`SpikeMountWorkPst2`
    - 呼叫端：`DebugButton_Click`（`Form1_Maintab56.vb` 約 line 1396）內的 `Await SpikeXXX()`
11. ⚠ 善後提醒（昨天遺留）：先前 spike 曾呼叫 `AddPSTStore`（疑似持久寫 profile）。下個對話開工前，建議確認 Outlook 資料檔清單無意外掛載的重複 PST。

---

## 八、本輪通用教訓（適用下個對話）

- **不要從一個 workload 外推另一個**：附件 10× ≠ 內文 10×（內文僅 2×）。負載性質（per-call bound vs I/O bound）決定一切。
- **量測前先確認「有沒有真的在讀東西」**：附件 workload 一度量到 3486 封/s 是空轉假象的鄰居（後證實非空轉，但「揪空轉」的計量欄位 `有附件信 X/8000`、`字元數` 是必備防呆）。
- **UI 緒同步跑長迴圈 = 假性凍結**：收集/走訪幾萬項放 UI 緒會佔死 timer、看似當掉。RDO 讀取一律進 `Task.Run` 背景緒（COM 仍須注意 STA，但 RDO 本行程 session 可背景跑；OOM 物件才必須留 UI 緒）。
- **OOM 碰 IRM `.Body` 必凍**：production 讀內文一律走 RDO，不走 OOM。
- **不憑推論直接寫 API**：本輪先用 P4 多形式探測再決定 resolve 寫法，避免重蹈昨天 `.StoreID` 憑猜致命。
- **印值要印完整、識別 store 用檔路徑（PstPath/FilePath）不用 Name**（昨天教訓，今天沿用，比對全用完整路徑、無踩雷）。
