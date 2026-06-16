# Tab3 (尋找附件) 終極效能優化計畫

根據您的需求，我們需要針對「大量郵件讀取速度」、「有/無附件的過濾機制」、「GetArray 的應用」以及「是否能避免開啟 MailItem (Phase 2) 來讀取附件名稱」進行深度的 COM 效能考量。

## 效能耗損階層分析 (由快到慢)
在 Outlook COM 操作中，速度差異是數量級的：
1. **純 .NET 記憶體操作 (LINQ)**：極快 (微秒級)，成本近乎 0。
2. **`GetTable(Filter)` + `GetArray()`**：非常快。將過濾條件交給 Outlook PST 引擎處理，並批次傳回陣列，是撈取大數據的王道。
3. **`GetTable` 後逐筆檢查 MAPI 屬性欄位**：快。但只能取得「信件層級」的單一屬性 (如 Subject, Size)。
4. **`GetItemFromID` 實例化 `MailItem`**：**極度緩慢！** 尤其是還要再往下層實例化 `Attachments` 集合並遍歷 `FileName` 時，這就是目前 Tab3 卡頓的元凶。

## 核心優化策略：消滅 Phase 2 的依賴

目前的痛點在於：只要勾選了「附件名稱」，程式就會被迫進入 Phase 2，對*每一封*有附件的信呼叫 `GetItemFromID`。

**破局關鍵：將「附件名稱」比對，上推（Push-down）到 Phase 1 的 DASL 查詢中交給 PST 引擎做！**

Outlook DASL 支援使用 `urn:schemas:httpmail:attachmentfilename` 來搜尋附件名稱，這意味著我們根本不需要把信打開 (MailItem) 就能知道它有沒有包含特定檔名的附件！

---

## Proposed Changes (優化後的新工作流)

我們將嚴格遵守「盡可能不用 COM -> 不得已用 GetTable -> 絕對不開 MailItem」的過濾漏斗順序。

### 漏斗第 1 層：PST 引擎過濾 (Phase 1 - `GetTable` + `GetArray`)
*   **預設條件**：`"urn:schemas:httpmail:hasattachment" = True` (必定套用，最快)。
*   **動態條件 (附件名稱)**：如果 UI 有勾選「附件名稱」，直接將其加入至 SQL Filter 中。
    *   例如：`AND "urn:schemas:httpmail:attachmentfilename" LIKE '%pdf%'` 或使用 `ci_phrasematch`。
    *   *效益：將原本要在 Phase 2 慢慢開信比對檔名的工作，瞬間在底層完成。回傳的數量將從「所有有附件的信」驟降至「真正符合檔名關鍵字的信」。*

### 漏斗第 2 層：.NET 記憶體過濾 (Phase 1.5 - LINQ)
*   **附件大小**：將第一層拉回來的 `List<MailItemInfo>` (已包含 Size 屬性)，直接在記憶體中用 LINQ `.Where(Size >= min And Size <= max)` 篩選。
*   *效益：成本為 0，且保留了基礎快取重用的彈性。*

### 漏斗第 3 層：非不得已的 COM 實例化 (Phase 2 - 只針對「附件個數」)
*   **附件個數**：因為 `GetTable` 無法直接且穩定地回傳「附件確切數量」(沒有對應的輕量 MAPI 欄位)，只有當使用者 **確實勾選了「附件個數」** 限制時，我們才進入 Phase 2。
*   此時進入 Phase 2 的信件，已經過「有無附件」、「檔名」、「大小」三層嚴格篩選，數量極少（可能從 3000 封變成 10 封）。此時再對這 10 封執行 `GetItemFromID` 取 `Attachments.Count`，使用者幾乎感覺不到延遲。

---

## 修改範圍與架構切分

### 1. `Button3_Click` (主控流程) 邏輯重構
修改判斷邏輯，改變進入 Phase 2 (`ScanAttachmentByName`) 的條件：
```vb
' [舊邏輯] 只要有 Keyword 或 Count，就進 Phase 2 慢慢撈
Dim hasKeyword = CheckAttachName.Checked AndAlso TextBox3.Text.Trim.Length > 0
If hasKeyword OrElse CheckAttCount.Checked Then
    finalItems = Await ScanAttachmentByName(...)
...

' [新邏輯] 只有 Count 才需要進 Phase 2，Keyword 已經在 Phase 1 解決了！
If CheckAttCount.Checked Then
    finalItems = Await ScanAttachmentByCountOnly(...) ' 函式瘦身，只算個數
Else
    finalItems = BuildListViewItem_Tab3(...)
End If
```

### 2. `ScanFolderWithAttachment` 增強過濾字串
*   傳入 `keyword` 參數。如果 `keyword` 不為空，將其串接進 DASL Filter 字串中。
*   需要注意：帶有特定檔名條件的查詢結果，**不能**被存入目前的「無條件全集快取」中，或者快取的 Key 必須包含 Keyword。為求程式碼乾淨且避免找錯，當包含名稱搜尋時，我們可以選擇「繞過快取直接查引擎 (因為引擎查檔名也很快)」，或是「將 Keyword 加入 Cache Key」。

### 3. `ScanAttachmentByName` 降級為 `ScanAttachmentByCount` (Phase 2)
*   移除所有跟 `att.FileName.IndexOf` 有關的程式碼。
*   這個函數現在的唯一責任是：打開 `MailItem` -> 讀取 `Attachments.Count` -> 關閉 `MailItem` -> 根據 Count 決定是否保留。
*   速度將獲得二次提升，因為不需要遍歷 `Attachments` 集合中的每一個物件。

## Open Questions

> [!WARNING]
> 關於 DASL 找尋附件名稱的精準度，`"urn:schemas:httpmail:attachmentfilename" LIKE '%keyword%'` 在某些特殊編碼的郵件（例如 RTF 格式或是內嵌圖片的 cid 附件）可能會連同一些不可見的系統附件一起算進去。
> 您是否同意我們「先採用 DASL 全面下放」的策略？如果實測後發現名稱比對有誤差，我們再來微調 SQL 語法或加回第二道防線。

> [!IMPORTANT]
> 關於快取機制 (`_tab3Phase1Cache`)：如果使用者勾選了特定「附件名稱」，您傾向於：
> 1. **擴充快取 Key**：將 Keyword 當作 Cache Key 的一部分組合進去 (例如 `FolderPath_Keyword_pdf`)。
> 2. **直接繞過快取**：如果有檔名條件，就直接重新 `GetTable` (因為 GetTable 很快)，保持快取只儲存「該資料夾所有有附件的信」的乾淨狀態。
> （我個人推薦 **選項 2**，架構較單純不易出錯，且 GetTable 本身速度極佳）

請確認以上思路是否符合您的期待，或者有需要微調的地方？確認後我們即可動手開刀！
