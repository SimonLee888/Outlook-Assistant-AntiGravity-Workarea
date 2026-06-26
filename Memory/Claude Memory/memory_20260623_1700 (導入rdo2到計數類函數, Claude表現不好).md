# memory — phase 2 (A) 計數類 RDO `_rdo2` 化交接

> 用途：交給下一個對話接手。承接 2026/06/23 上午的「phase 1 attach/body 完成」交接。本輪完成 (A) 類 per-folder 計數兩支(GetMailCount / GetFolderCount)的 `_rdo2` store-scoped RDO 化。
> 撰寫：2026/06/23 by Simon/Claude Opus 4.8。專案：Outlook Assistant（VB.NET .NET10 WinForms，Outlook LTSC 2021 OOM + Redemption RDO，本地 PST，olNotExchange）。

---

## 0. 鐵則（每次修改都要守，同上一份）
1. 動既有 working code 前先問 Simon，不擅自改。
2. 一致、對稱、解耦、介面清楚乾淨抽象、呼叫路徑統一 —— 「多數呼叫端零修改」的根本。
3. 先驗再 commit、不猜 API；同一問題修兩次未解就放探針、多假設、找唯一符合所有線索的根因。
4. 對話貼/傳的程式碼最新 > 檔案區 > memory。讀大檔分段讀。
5. 既有註解/debug/日期標記不可遺失；dead code 用說的不自刪。
6. 小改動 → 貼 diff(bottom-up)；大範圍散落 → 整檔下載。回覆繁中+英文。

### ⚠ 本輪新增鐵則(Simon 嚴正要求)
- **禁區**：`Module_Win32API.vb` 內 `#Region "■ 99 舊版備用 (勿刪)"` 的所有函數(GetTotalFolderCountAsync、第307行彙總等,平行主體均已註解掉)**永遠不看、不動、不列入重構考量**。grep 命中此 Region 直接視為死路徑略過,不得在動手前拿出來猶豫或詢問。
- **檢視一次到位**：grep 命中呼叫端後,必須在同一輪把每個呼叫端的「所在 Region / 死活 / 語意分類」全部看清並歸類完畢。不可把 grep 結果當待辦、臨到動手才逐項檢視,導致問題一個接一個爆、反覆打斷 Simon。死路徑(整段註解、位於舊版備用 Region)自行判定略過,不需停頓詢問。

---

## 1. 本輪已完成(2026/06/23 下午)

### 1.1 folder 跨 session 解析已驗證(探針 12/12,不要重測)
- 探針 `SpikeResolveFolderOnRdo2`(+ 輔助 `HarvestFoldersFromStore`)放在 `Form1_Maintab56.vb` 的 `#Region "  ├ Debug 測試區"`。
- 實測結果:**store-scoped 12/12、dual 12/12、single 0/12,且 .Name 全吻合、Items 全讀到**。
- production 採 **store-scoped 單參數** `store.GetFolderFromID(eid)`(經 `GetRdo2Store(folderPath)` 取 store),與 message 版 `store.GetMessageFromID(eid)` 同形。
- ※ 探針輔助碼曾有命名 bug:用了 VB 保留字 `sub` 當迴圈變數 → 一連串 BC30084/BC36673 連鎖錯。已改 `sf`。教訓:VB 保留字(Sub/Function/End...)不可當識別字。
- 探針使命已達成,可整段刪除或留 Debug 區。

### 1.2 兩支 RDO 讀取層已建(放 Module_Outlook.vb `#Region "  ├ Layer3 RDO 加速讀取層"`,接 GetMailBodyRdo 之後)
- `GetMailCountRdo(folderPath, eid, sid) As Long`:`GetRdo2Store(path).GetFolderFromID(eid).Items.Count`。
- `GetFolderCountRdo(folderPath, eid, sid) As Long`:同上 → `.Folders.Count`。
- **回傳型別約定**:count 類回 `Long`,解析失敗(store/folder Nothing/例外)回 **-1**(與 L3 既有 `fail→-1` 慣例一致),proxy 判 `<0` fallback。**不是**照 attach/body 的 `Nothing` 哨兵——型別要依改的目標對象自身慣例決定,不可慣性套早上的模板。
- `sid` 參數目前未用(store-scoped 單參數即可解),保留對稱備雙參數 fallback。

### 1.3 兩支 L2.5 proxy 已改造(Module_Outlook.vb)
- `GetMailCount` / `GetFolderCount` 皆加 `Optional skipCache As Boolean = False`。
- 階梯:`If Not skipCache Then ①記憶體 → ②DB lazy` → `③ GetXxxRdo() 失敗(<0)則 fallback GetXxxL3()` → `If Not skipCache Then 寫快取`。
- eid/sid 來源:`folder.EntryID` / `folder.StoreID`(與 L3 原 ⓪ tier 同做法,零新風險)。

### 1.4 L3 ⓪ tier 已移除留註解(照早上 attach/body 收尾慣例)
- `GetMailCountL3` / `GetFolderCountL3` 內原本各有一段「⓪ 用舊共用 `_rdo`」的 RDO 路徑,本輪移除,各保留 ⓪ 標題行 + 一行:
  `' ⓪ RDO 路徑已上移至 L2.5 GetXxxRdo(store-scoped on _rdo2),L3 回歸純 OOM。 2026/06/23 by Simon/Claude`
- L3 現為純 OOM(GetMailCountL3:PR_CONTENT_COUNT→Items.Count;GetFolderCountL3:Folders.Count)。

### 1.5 活路徑直呼者已改 skipCache:=True(維持原「繞過快取直讀」語意)
- `Form1_MainTab12.vb` 397/401/417/421:F5 forceRefresh 流程,4 處 `GetMailCountL3/GetFolderCountL3(...)` → `GetMailCount/GetFolderCount(..., skipCache:=True)`。
- `Module_SQLite2.vb` 745/746:snap 狀況 A 重讀,2 處同改。
- 理由:這些直呼者本就刻意繞過快取直讀 L3、讀完自己寫快取。若不傳 skipCache,函數內 default False → 會開始走 proxy 快取讀寫,語意改變。改 skipCache:=True 維持原語意,同時讓它們享受 RDO 派工。

---

## 2. 本輪「未動」清單(範圍邊界,重要)
- `Module_SQLite2.vb` 747 的 `GetFolderSizeL3` **未動**(見 §3 size 子任務)。
- `Module_Win32API.vb` 185/307 **未動**(死路徑,■99 禁區)。
- `GetMailCountAllL3` / `GetFolderCountAllL3` 內部仍直呼 `GetMailCountL3/GetFolderCountL3`(子樹彙總,屬 (B) 類,本輪範圍外)。
- 舊共用 `_rdo` 與 preLoad 三兄弟(RdoPreloadAttach_1/2/3)仍在,待 RDO 全面取代、實測穩定後依 Simon 指示整批移除。

---

## 3. 下一支獨立子任務:GetFolderSize 的 RDO 化(需小實驗)
- **效益不是零**(本輪查證推翻舊註解):Redemption 有 table 能力。`RDOFolder.Items` 暴露 `MAPITable` 物件,可預設 `MAPITable.Columns` 後不開啟訊息、批次讀屬性(等價 OOM GetTable+GetArray);另有 `MAPITable.ExecSQL` 支援 DATALENGTH/LEN(RES_SIZE)。來源:dimastr.com redemption mapitable.htm / RDOItems.htm / history.htm。
- 複雜度高於 count:要正確設 MAPITable.Columns=PR_MESSAGE_SIZE_EXTENDED、處理 variant error 列(MAPI_E_NOT_FOUND)、release MAPITable。故獨立為子任務,先做 column 設定小實驗驗證 parity(與 OOM GetTable 逐封加總結果一致)再上。
- PST 無彙總 size 屬性(PR_MESSAGE_SIZE_EXTENDED 在 folder object 回未知),只能逐封加總——此事實不變,RDO 化是換成 _rdo2 上的 MAPITable 逐封讀,非取得彙總屬性。

---

## 4. 之後仍在 horizon 的項目(承上一份)
- (B) 類子樹枚舉 RDO 化(GetSubtreeToList 等):有 visibility 陷阱(RDO 枚舉到 OOM 看不到的隱藏夾,27 vs 22,_rdoFastPath 至今 False)。CP 值低、風險高,Simon 與 Claude 共識:**暫緩**,等 (A) 類穩定後再評估是否值得。BuildBfsFolderTree/GetSortedSubFolders 的子夾枚舉走 OOM .Folders,無 visibility 陷阱,但純結構讀取、已有三層快取、熱路徑不打 COM,搬 RDO 效益低,亦暫不搬。
- 單執行緒效率 Stopwatch 分段量測(附件~3000封/s vs 內文~150封/s 差20倍,量 RDO 解析/屬性讀/normalize/dispatch 占比)。
- 多執行緒(phase C)延後至單執行緒優化完成。
- 終極目標:`_rdo2` 成為 cache miss 時的安全日常優先讀取路徑。原則:可用 _rdo2 加速的都排上,依難度/效益/影響/頻率排序。

---

## 5. 本輪 Claude 表現檢討(Simon 明確表達不滿,記取)
1. **檢視不徹底、分批冒問題**(最嚴重):未在首次 grep 時一輪內把呼叫端的 Region/死活/語意全部歸類,導致問題臨動手前一個接一個爆。
2. **死路徑判斷力不足**:■99 禁區內整段註解的函數,該自行判定略過,卻拿出來製造停頓詢問。
3. **慣性套模板、沒認真想**:把早上的 Nothing 哨兵慣性套到 Long 回傳(答案就在 L3 既有 -1 慣例);用沒查證的舊註解想跳過 GetFolderSize,被要求查證才發現 RDO 有 MAPITable 能力。
4. **沒善用既有資產**:沒 grep `spike` 就斷言探針沒留(其實都在 Debug 測試區)。
→ 已將「■99 禁區」與「檢視一次到位」記入長期記憶(memory #12, #13)。
