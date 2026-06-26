# memory — RDO `_rdo2` 讀取層重構交接（phase 1 完成 → phase 2：folder / tree / count）

> 用途：交給下一個對話接手。phase 1（附件檔名 + 內文）已用 `_rdo2` 獨立 session 改造完成，並定出一套標準架構。phase 2 要把**同一套**搬到 folder tree 讀取與計數類底層函數。
> 撰寫：2026/06/23 by Simon/Claude。專案：Outlook Assistant（VB.NET .NET10 WinForms，Outlook LTSC 2021 OOM + Redemption RDO，本地 PST，olNotExchange）。

---

## 0. 鐵則（每次修改都要守）
1. 動既有 working code 前**先問 Simon**，不擅自改。
2. 一致、對稱、解耦、介面清楚乾淨抽象、呼叫路徑統一 —— 這是「多數呼叫端零修改」的根本原因，務必持續維持。
3. 先驗再 commit、不猜 API；同一問題修兩次未解就放探針、多假設、找唯一符合所有線索的根因。
4. 對話中貼/傳的程式碼最新，其次檔案區，皆優先於 memory。讀大檔分段讀，不整檔吞。
5. 既有註解 / debug 思考 / 日期標記不可遺失；dead code 用說的不自刪。
6. 小改動 → 直接貼 diff（bottom-up）；大範圍散落 → 才整檔下載。回覆用繁中 + 英文。

---

## 1. 已驗證的事實（probe 結果，**不要重測**）
- `_rdo2` = 獨立 `RDOSession`，登入 **[Work]** profile，看得到 **26 個 store**（含全部 PST）。
- OOM 取得的 **message EntryID** 在 `_rdo2` 上可解：雙參數 `GetMessageFromID(oomEid, oomStoreId)` **12/12**；store-scoped `store.GetMessageFromID(oomEid)` **12/12**；單參數 **0/12**（跨 session 必須帶 store，否則 `MAPI_E_UNKNOWN_ENTRYID`）。
- production 採 **store-scoped**（零新風險：`FolderPath → 解析 store 名 → 在 `_rdo2.Stores` 找 RDOStore`）。
- 單執行緒實測：**附件檔名 ~3000 封/s、內文 ~150 封/s**。Simon 認為夠用，多執行緒（phase C）延後。
- 「OOM EntryID 跨程式重啟穩定」是 SSD 快取一直依賴且實際運作正常的前提（被 move/copy 的信 EntryID 會變 → cache miss 自動重讀，自癒）。

---

## 2. phase 1 已建立的標準架構（**THE TEMPLATE，後續一律照此**）

**分層職責**
- L2：接 UI 要求，決定流程與分派順序。
- **L2.5 proxy**：決定資料來源，統一階梯（見下），多數呼叫端只認這層。
- L3：純資料層。`GetXxxL3()` = **純 OOM**；`GetXxxRdo()` = **純 RDO（`_rdo2` store-scoped）**。

**L2.5 統一階梯（附件 / 內文同形）**
```
① _cacheXxx 命中 → 回傳（0 開銷）                         （skipCache 時跳過）
② SSD（SQLite）命中 → lazy load 回填 → 回傳（微小開銷）    （skipCache 時跳過；body 目前無 SSD，預留佔位）
③ 讀取分派：_rdo2 在(且 store 可解) → GetXxxRdo()（中開銷高速）
            否則                    → GetXxxL3()（OOM，最貴中低速）
寫快取（skipCache 時不寫）→ 回傳
```
- `skipCache` 引數 = 「少數堅持直讀」用（如 build pass 一次掃數萬封，跳過 cache 讀寫避免撐爆記憶體，但**仍走 RDO 分派**）；未來 liveSnap / forceRefresh 也走這個。
- `GetXxxRdo()` 解析失敗（store 找不到 / 例外）一律**回 Nothing**，由 L2.5 fallback 到 `GetXxxL3()`。
- parity 鐵則：`GetXxxRdo` 的輸出必須與 `GetXxxL3` **逐筆一致**（相同過濾條件、相同正規化、相同屬性語意）。
- 命名慣例：RDO 讀取層用 `GetXxxRdo`；L3 維持 `GetXxxL3`；L2.5 proxy 保持原名、加 `Optional folderPath`（內文類需要）、`Optional skipCache`。
- 收尾：`GetXxxRdo` 確認成功後，**移除對應 `GetXxxL3` 內的 RDO 路徑**，只留一句註解說明「RDO 已上移至 L2.5」。

---

## 3. 共用基礎設施（**已建好，可直接重用**）
- `_rdo2StoreByName : Dictionary(String→RDOStore)`：權威表，首次掃 `_rdo2.Stores` 一次填滿 ~26 個，擁有 COM ref。
- `_rdo2StoreByPath : Dictionary(String→RDOStore)`：FolderPath→RDOStore 記憶化，免熱路徑重跑 `GetStoreNameFromPath` 解析；值為 byName 參考，不另釋放；含 Nothing 亦記。
- `GetRdo2Store(folderPath) As RDOStore`：phase 2 folder 解析直接重用（store 拿到後 `store.GetFolderFromID(eid)` 或 `_rdo2.GetFolderFromID(eid, sid)`）。
- `ReleaseRdo2Stores()`：在 `CheckRDO` 取消勾選 + `Form1_FormClosing` 兩處呼叫（獨立 session 須 `Logoff()` 再 release，否則 Outlook 關不乾淨）。
- 單執行緒（UI 緒）存取，用一般 Dictionary；phase C 平行才需評估執行緒安全。

---

## 4. phase 2 目標與切分

### 4.0 開工前置探針（**必做**）
folder-resolve 探針，比照昨天 message 版：取幾個 OOM folder EntryID + OOM StoreID + FolderPath，在 `_rdo2` 上測
①單參數 `GetFolderFromID(eid)`（預期失敗，基準）
②雙參數 `GetFolderFromID(eid, oomStoreId)`
③store-scoped（`GetRdo2Store(path).GetFolderFromID(eid)` 或等價）
每個成功就讀 `.Name` / `PR_CONTENT_COUNT` 比對。過了再建 folder 的 `GetXxxRdo`。信心高（per-folder L3 早就在共用 `_rdo` 上用 `GetFolderFromID` 成功），但照原則先驗。

### 4.1 目標函數分兩類
**(A) 「解析單一已知資料夾 + 讀屬性」→ 易，跟 message 一樣**
- `GetMailCountL3`（PR_CONTENT_COUNT）
- `GetFolderCountL3`
- `GetFolderSizeL3`（注意：PST 無彙總 size 屬性，需 `GetTable` 逐封加總 PR_MESSAGE_SIZE）

**(B) 「枚舉子樹」→ 難，有 visibility 陷阱（見 §5）**
- `GetSubtreeToListL3` / `GetSubTreeL3`（folder tree 讀取）
- `GetMailCountAllL3` / `GetFolderCountAllL3` / `GetFolderSizeAllL3`（子樹彙總，目前 `_rdoFastPath=False` 關閉中）

### 4.2 建議順序
先做 **(A)** 拿到 per-folder 加速（低風險、快速見效）→ 再把 **(B)** 另立子任務專門解 visibility 過濾。

---

## 5. ⚠ phase 2 最大風險：RDO 子樹枚舉的 visibility 不一致（**必讀**）
- Redemption 走 MAPI 會枚舉到 OOM **看不到**的隱藏 / 非-IPM 夾（Recoverable Items、Conversation Action Settings…），導致子樹資料夾數比 OOM 多（**實測 27 vs 22**）。
- 這正是 `_rdoFastPath` 目前設 **False** 的原因（`Module_Outlook.vb` 宣告區約 40–51 行有完整說明）。
- 因此：**(A) 類**（單一資料夾解析 + 讀屬性）可安全搬到 `_rdo2`；**(B) 類**（子樹枚舉）搬之前，`GetXxxRdo` 內**必須先做與 OOM 對齊的 visibility / IsMailFolder 過濾**（排除隱藏夾、非郵件夾），否則 tree 與 All 計數會與 OOM 不一致。
- 參考既有 `GetSubtreeToListL3_Rdo`（已存在的 RDO 子樹輔助）——恢復 RDO 快速路徑的條件就是「在它裡面比照 OOM 可見性過濾隱藏夾」後再把 `_rdoFastPath=True`。

---

## 6. 每支函數套用 template 的標準流程（checklist）
1. grep 該函數**所有呼叫端**：找出 L2.5 proxy 攔截點 + 是否有繞過 proxy 的直呼者。
2. 讀該 `GetXxxL3` 內部，記錄 parity 點（過濾條件、屬性名、release 模式、正規化）。
3. 寫 `GetXxxRdo`（store-scoped via `GetRdo2Store`；解析失敗回 Nothing；容量用手邊現成欄位精準預配置，如 `mail.AttachCount`）。
4. 在 L2.5 proxy 插入 ③ RDO tier（mem→SSD→RDO→OOM；`skipCache` 對稱）。
5. 多數呼叫端零修改；只有繞過 proxy 的直呼者改帶 `folderPath` + `skipCache`。
6. 確認 `GetXxxRdo` 成功後，移除 `GetXxxL3` 內 RDO 路徑，只留一句註解說明 RDO 已上移。
7. **先驗正確性**（與舊版逐筆一致）→ 再看速度。
8. 動到既有 working code 前先問 Simon。

---

## 7. Q3 待辦：單執行緒效率量測（晚一點做，已記入長期記憶）
目標：確認現行 `_rdo2` 單執行緒是否已達應有效率，或 pipeline 周邊仍有未優化處（不是 RDO call 本身）。用 Stopwatch 分段累計一個批次：
- RDO 解析（`store.GetMessageFromID`）vs 屬性讀取（`.Body` / `.Attachments`+`.FileName`）。
- 內文拆「`.Body` 讀取」vs「`NormalizeMailBody`（regex）」——150 封/s 比 3000 封/s 慢 20 倍，疑與 body payload + normalize 有關，要量出瓶頸在哪。
- L2.5 dispatch 開銷（dict 查找、`GetRdo2Store` path-memo 命中成本）。
- build pass 迴圈：ThrottledYield 頻率、`BuildBigramSet`/SimHash 計算 vs body 讀取占比。
- 產出各段 ms 與 %，判斷上多執行緒前是否先削周邊脂肪。

---

## 8. 檔案 / 函數位置索引（行號為 2026/06/23 版，可能微漂移）
- **Module_Outlook.vb**：`_rdo2` 宣告(44) 與 `_rdo2StoreByXxx` 欄位；`_rdoFastPath` 與 visibility 說明(40–51)；新增的 `GetRdo2Store`/`GetAttachFilenameRdo`/`GetMailBodyRdo`/`ReleaseRdo2Stores`(GetMailBodyL3 之後 ~2210)；proxy `GetAttachFilename`(885)/`GetMailBody`(909)；`GetAttachFilenameL3`(2123)/`GetMailBodyL3`(2175)；folder L3：`GetMailCountL3`(1313)/`GetFolderCountL3`(1367)/`GetFolderSizeL3`(1737)、All 版本(1428/1629/1841)、`GetSubtreeToListL3_Rdo`(2369)；`GetStoreNameFromPath`(2666)；`TryMarshalRelease`(296)。
- **Form1.vb**：`Form1_FormClosing`(294)、`CheckRDO_CheckedChanged`(1166)。
- **Form1_Maintab56.vb**：build pass `PreComputeFuzzySimHashAsync`(~440)、S5 `FilterCandidatesByJaccardAsync`(~497)。
- **Form1_MainTab34.vb**：Tab3 `GetAttachFilename` 呼叫(303)、Tab4 `GetMailBody`(838,1083)。

---

## 9. 待 Simon 確認 / 決定的開放項
- `Form1.vb` 的 `FormClosing` / `CheckRDO` 兩段 diff 是否已套用（上次依 /mnt/project 快照寫，需核對現行版本）。
- `GetAttachFilenameL3` 的 `New List(4096)` 是否也改 `(mail.AttachCount)` 與 Rdo 版對稱（Rdo 版已採 `mail.AttachCount`）。
- S5 是否也吃 RDO（目前 B1 只接 build pass；S5 上游有 MailItemInfo，改 projection 帶 folderPath 即可，低量、非必要）。
- preLoad 三兄弟（`RdoPreloadAttach_1/2/3`）與舊 `_rdo` 共用 session：待 RDO 完整取代、實測穩定後，依 Simon 指示整批移除。
