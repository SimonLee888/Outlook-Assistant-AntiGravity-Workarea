# Outlook Assistant — 專案 Memory（延續 memory_20260702_1108）

> 輸出時間：2026/07/02 12:54 (Asia/Taipei)，末次更新 ~14:xx（見「七、」）
> 執行環境：Claude Code（非 claude.ai 對話），有本機檔案編輯 + bash/PowerShell 執行能力
>
> ⚠️ **命名已變更**：本文件二~六節內文寫的 `GetBasicMailInfo*`/`_cacheBasicMailInfo`/`DbGetBasicMailInfo*`/`basic_maillist` 等名稱，
> 在同一輪對話後段（見「七、」）已全部改名去掉 `Basic`（`GetMailInfo*`/`_cacheMailInfo`/`DbGetMailInfo*`/`mailinfo_list`）。
> 下方二~六節保留原文不回頭改，是因為那是決策當下的真實記錄；要找**目前**的正確名稱一律以「七、」為準。

---

## 一、本輪做了什麼（延續上一份 memory 的「⚠️ 目前真實狀態」）

上一份 memory (`memory_20260702_1108`) 結尾記錄的疑慮是：「探針從未被實際執行過，沒有任何真實 parity/benchmark 數據」「Claude Code 無法自動跑通這個探針拿數據（依賴 GUI 手動觸發 + DebugForm 視窗輸出）」。

本輪由 Claude Code 實際動手驗證，**這個結論已經過時**：Claude Code 改了程式碼、自己建立無 GUI 觸發機制、自己編譯、自己跑，全程自主拿到了真實數據，且發現並修正了兩個真實 bug。完整過程：

1. **建了無 GUI 自動化探針機制**（詳見二.1），讓 `TestProbeBasicInfoRdoParity` 可以不靠 Tab1 TreeView 手動選取、不靠人工按除錯按鈕觸發，改用命令列參數指定 root 資料夾，結果額外寫成文字檔（`%TEMP%\OutlookAssistant_ProbeResult.txt`），Claude Code 可直接讀檔拿結果。
2. **編譯環境踩坑**：`dotnet build`（.NET CLI）無法處理這個專案的 COM 參照（`ResolveComReference` 任務不支援 .NET Core 版 MSBuild），必須改用 VS 內建的 .NET Framework 版 MSBuild：
   `C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe`
3. **第一輪測試（小樣本，51 封信，`寄件備份, 2013~2018, Toshiba` store 的 `2013` 資料夾）：parity 沒過**：
   - `RcvTime`: 51/51 全部不符，固定差 -8.00 小時
   - `SenderName`: 51/51 全部不符
   - `Subject`: 5/51 不符（含日文亂碼）
   - （中途還踩了一個假陰性：PowerShell `Start-Process -ArgumentList` 傳含空格/逗號的字串沒包雙引號，argv 被 OS 拆散，`/autoprobe:寄件備份, 2013~2018, Toshiba|2013` 只留下 `/autoprobe:寄件備份,` 這個殘缺 token，探針默默 fallback 到空的「刪除的郵件」資料夾，回報 0/0 造成誤判「有跑但沒東西可比」。修法：`-ArgumentList` 的字串本身要包含字面雙引號 `'"/autoprobe:...|..."'`，讓它在 OS 層被當一個 token。）
4. **加診斷抽樣**（`fieldSamples`，每個不符欄位存前 5 筆 OOM vs RDO 實際值）後重跑，抓到兩個根因：
   - **時區**：MAPITable 讀到的 `PR_MESSAGE_DELIVERY_TIME` 是 UTC，OOM 的 `.ReceivedTime` 屬性已經是系統本地時區(UTC+8)，差值穩定 -8.00 小時。
   - **字元編碼**：`GetBasicMailInfoRdo` 裡宣告的字串 proptag（`PR_SUBJECT`/`PR_SENDER_NAME`/`PR_INTERNET_MESSAGE_ID`/`PR_SENDER_EMAIL`）字尾誤用 `001E`（**PT_STRING8, ANSI codepage**），日文字元讀出來變 `?`，全形括號 ［］ 被「最佳近似」置換成形似但不同的字元 〔〕。正確應為 `001F`（**PT_UNICODE**）。
5. **套用修正**（已寫入正式程式碼，非探針區）：
   - `Module_Outlook.vb` `GetBasicMailInfoRdo`（約 1409 行起）：四個字串 proptag 字尾 `001E`→`001F`；`RcvTime` 讀值改 `DateTime.SpecifyKind(raw, DateTimeKind.Utc).ToLocalTime()`。
6. **重新驗證，全部通過**：
   - 51 封信（單一年份夾 2013）：parity 全過
   - **15,865 封信、10 個資料夾（整個 `寄件備份, 2013~2018, Toshiba` store，含 2013~2018 六個年份夾）：parity 全過，且 RDO 比 OOM 快 2.0 倍（306ms vs 609ms 累加）**

**結論：`GetBasicMailInfoRdo` 的資料正確性已經在真實資料（15,865 封，含中日文、跨年份）上驗證通過，不再是「設計時的預判疑慮」，是有真實數據支撐的結論。**

---

## 二、新建的基礎設施（全部包在 `PROBE_BASICINFO_RDO` 可刪除標記內，grep 這個 token 可整批清除）

### 2.1 無 GUI 自動觸發機制

- **`Module_Outlook.vb`**：新增 `ResolveDefaultProbeFolder()`（預設探針資料夾＝第一個 PST 的「刪除的郵件」，失敗退回 Inbox）、`RunAutoProbeBasicInfoRdo(arg As String)`（等 `_pstStoreList`/`_rdo2` 就緒，逾時 30 秒；解析命令列參數選 store/folder；呼叫 `TestProbeBasicInfoRdoParity(target)`）、`RunListStoresDump()`（純列出所有 store + 第一層資料夾名稱到 `%TEMP%\OutlookAssistant_StoreList.txt`，不跑掃描，用來核對確切名稱避免猜錯）。
- **`Form1.vb` `Form1_Shown`**：尾端加兩行命令列判斷（純新增，不影響原有啟動流程）：
  - `/autoprobe` → 用預設資料夾跑探針
  - `/autoprobe:StoreName|FolderName` → 指定 store（DisplayName 包含比對，不用精確相等）+ 第一層子資料夾（Name 包含比對）
  - `/autoprobe:StoreName|*` → 指定 store 的**整個根目錄**（含全部子資料夾，用於較大樣本測試）
  - `/liststores` → 觸發 `RunListStoresDump()`
- **`Form1_Maintab56.vb` `TestProbeBasicInfoRdoParity`**：簽章加 `Optional rootOverride As Folder = Nothing`；有傳入時額外把所有輸出鏡射寫進文字檔（GUI 手動觸發路徑完全不受影響，不寫檔）；加了 `fieldSamples` 抽樣（每個不符欄位存前 5 筆 `OOM=... | RDO=...` 對照字串，RcvTime 額外印時差小時數）。

### 2.2 呼叫範例（供下次直接重用，不用重新設計）

```powershell
# 列出所有 store + 第一層資料夾名稱（核對用，不跑掃描）
Start-Process -FilePath $exe -ArgumentList '"/liststores"'

# 指定 store 下某個資料夾（fuzzy 比對）
Start-Process -FilePath $exe -ArgumentList '"/autoprobe:StoreDisplayName片段|資料夾名稱片段"'

# 整個 store 根目錄（較大樣本）
Start-Process -FilePath $exe -ArgumentList '"/autoprobe:StoreDisplayName片段|*"'
```

⚠️ **`-ArgumentList` 的字串本身要包含字面雙引號**（`'"..."'` 這種寫法），否則含空格/逗號的參數會被 OS 拆成多個 argv token，探針會靜默 fallback 到預設資料夾，回報 0/0 造成誤判。

結果讀取：`%TEMP%\OutlookAssistant_ProbeResult.txt`（parity 結果）、`%TEMP%\OutlookAssistant_StoreList.txt`（store 清單）。

### 2.3 前置條件

- Outlook.exe 必須已經在跑（`InitMapiNamespace` 檢查不到會跳 `MessageBox.Show` + `End`，這個對話框會卡住無人值守流程）
- 編譯必須用 VS 內建 MSBuild.exe，不能用 `dotnet build`

---

## 三、Task 2b 已完成 — RDO 已接上免-folder版 `GetBasicMailInfo`

**`GetBasicMailInfo(fPath, eid, sid, needTopic, cToken)` 免-folder多載版本**（`Module_Outlook.vb` 約 905 行，Layer2.5 快取存取點，Tab4/Tab5 用）的 ③ 這一步已改成 RDO 為主、OOM 為 fallback（不是「RDO=fallback」，是 Simon 訂正過的用詞：③ 本身是 Layer3 COM 讀取，RDO 是③裡的主要手段，OOM 才是 RDO 讀不到時的真正 fallback）：

```vb
' ③ Layer3 COM 讀取 — RDO 為主要手段
If _rdo2 IsNot Nothing Then
    Dim rdoResult = GetBasicMailInfoRdo(fPath, eid, sid)
    If rdoResult IsNot Nothing Then
        Dim converted = rdoResult.Select(Function(m) (m, GetCleanSubject(m.Subject))).ToList()
        _cacheBasicMailInfo(fPath) = (converted, currentSnap)
        Return converted
    End If
End If

' ③ Fallback: RDO 讀不到才 GetFolderById 物化 → OOM 掃描
Dim folder As Folder = GetFolderById(eid, sid)
Dim resultList = Await GetBasicMailInfoOOM(folder, needTopic, cToken, fPath)
If resultList IsNot Nothing Then _cacheBasicMailInfo(fPath) = (resultList, currentSnap)
Return resultList
```

寫法跟全專案既有的「RDO優先、OOM兜底」慣例（例如 `GetSubtree` 的 1006~1017 行）對稱，閘門一樣是 `_rdo2 IsNot Nothing`（= CheckRDO 勾選）。Topic 計算＝`GetCleanSubject(mail.Subject)`，與 `GetBasicMailInfoOOM` 2490 行一致。**已編譯通過（VS MSBuild.exe，無錯誤）**，尚未做這個 wrapper 層級的端對端 UI 驗證(Tab4/Tab5 實際畫面)，但底層 `GetBasicMailInfoRdo` 本身已在 15,865 封真實信件上驗證過 parity，wrapper 只是薄封裝(cache miss後dispatch + Topic轉換)，邏輯風險低。

**Task 2a 也在本輪同一次對話收尾了**（原本記錄「留給下一輪」已過時，見下方六、）。

---

## 四、關鍵學習（本輪新增，補充 memory_20260702_1108 的「四、關鍵學習與原則」）

- **MAPI proptag 字尾 `001E` = PT_STRING8(ANSI codepage)，`001F` = PT_UNICODE**。Redemption `MAPITable` 讀字串類欄位一律要用 `001F`，用 `001E` 在非 ASCII 內容(中日文等)會出現 `?` 亂碼或形近字置換(不一定報錯，容易被忽略)。（`Outlook.Table`／OOM 的 `GetTable` API 對同樣字尾似乎有不同容錯行為，本輪 baseline 是 OOM 版本，沒有深究為何 OOM 端沒出現同樣症狀，只需記得 **RDO MAPITable 這條路必須用 001F**。）
- **Redemption `MAPITable` 的 `PT_SYSTIME` 回傳 UTC**，OOM `.ReceivedTime`／`.SentOn` 等屬性回傳的是系統本地時區時間。RDO 路徑讀時間類欄位都要 `DateTime.SpecifyKind(raw, DateTimeKind.Utc).ToLocalTime()`。
- **這個專案的 COM 參照需要 .NET Framework 版 MSBuild**，`dotnet build`（.NET CLI）會在 `ResolveComReference` 直接失敗(`MSB4803`)。編譯指令：
  `"C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "Outlook Assistant.vbproj" -p:Configuration=Debug -m -v:minimal`
- **PowerShell `Start-Process -ArgumentList` 傳含空格的單一參數字串，必須自己在字串裡包字面雙引號**（`'"...".'`），否則命令列會被拆成多個 token，且失敗是靜默的(探針會 fallback 到預設值而非報錯)，很容易誤判成「有跑但沒資料」。這是本輪唯一一次抓假訊號的地方，值得記住當作以後類似命令列注入案例的直覺檢查點。
- **Claude Code 在 Windows 上跑無人值守的桌面 COM 應用完全可行**，只要：(1) 前置依賴(這裡是 Outlook.exe)已啟動、(2) 觸發機制不依賴滑鼠鍵盤操作(改用命令列參數)、(3) 輸出改寫檔案而非只進 GUI 視窗。三個條件都是探針程式碼本身要負責的，不需要 computer-use 之類的桌面操作工具。

---

## 五、與上一份 memory 的差異總結

| 項目 | memory_20260702_1108 | 本份 (1254) |
|---|---|---|
| 探針是否執行過 | 否，只是設計 | 是，3 輪真實執行(51封→修正→51封→15,865封) |
| parity 結果 | 未知 | 15,865 封全通過 |
| RDO 效能 | 未知 | 2.0x(15,865封樣本，OOM 609ms vs RDO 306ms) |
| Claude Code 可否自主跑通 | 判定「不行，依賴GUI」 | 判定過時，已證實可行(無GUI命令列觸發+文字檔輸出) |
| GetBasicMailInfoRdo 正確性 | 未驗證的骨架 | proptag型別+時區兩個真實bug已修正並驗證 |
| Task 2b 下一步 | 「Simon 套用→手動接除錯觸發→...」 | 「接上 L2.5 ③ Fallback」的插入點與寫法已確認，等 Simon 同意後動手 |

---

## 六、Task 2a 完成（同一輪對話內，Task 2b 接上之後緊接著做）

**共用骨架 `ScanFolderTable`**（`Module_Outlook.vb`，放在 `GetFolderById` 旁邊，歸類「其他輔助函數」，不是獨立 L3 entry point——判斷依據：跟 `GetFolderById` 一樣不被 L2.5 直接呼叫，是被其他 L3 函式內部呼叫的共用工具，工作內容是 L3 COM 的活但組織上算 helper）：

```vb
Private Async Function ScanFolderTable(folder As Folder, cToken As CancellationToken, throttleFreq As Integer, onThrottled As System.Action, rowHandler As System.Action(Of Object(,), Integer), ParamArray columns() As String) As Task
```

命名去掉了原提案的 `L3` 字尾（Simon 決定）。⚠️ **`Action` 這個型別名稱在這個檔案裡會撞到 `Microsoft.Office.Interop.Outlook.Action`（COM 的規則動作物件），必須寫 `System.Action` 完整限定名，跟既有 `System.Exception` 撞名是同一類坑。**

負責範圍：開 table(呼叫端傳 columns) → `GetArray` 分頁迴圈 → 節流(`SmartThrottle`) → 取消(`cToken.ThrowIfCancellationRequested()`，每筆資料抓取前檢查)。欄位解析透過 `rowHandler` 委派給呼叫端（closure 累加進呼叫端自己的容器，List 或 Dictionary 都行）。`table` 生命週期自己管(開+Finally釋放)；**folder 的物化/釋放責任不變，本函式不介入**——兩個呼叫端原本一個自己解析+釋放、一個呼叫端持有+不釋放，這個差異照舊保留。

**呼叫端 1：`GetBasicMailInfoOOM`** 改用 `ScanFolderTable`，欄位邏輯完全不變(7欄，含 Topic 動態計算)。回歸測試：15,865 封信重跑 parity，結果不變(全通過，OOM/RDO 郵件數一致)，確認重構沒有引入行為差異。

**呼叫端 2：原 `GetFolderBasicByEntryIDL3` 改名 `GetBasicMailInfoAsDict`**（Simon 定案，過程：`GetBasicMailInfoOOMDict`/`OOMByID` → 發現一旦這函式同時能走RDO+OOM，掛`OOM`字尾會誤導 → `GetBasicMailInfoDict`/`Direct` 二選一 → 最終選 `AsDict`）。命名判斷依據記錄下來供以後參考：
- 不叫 `GetBasicMailInfo()`（跟另兩個 L2.5 多載同名）：因為它沒有 ①記憶體②DB 快取層，只有③(RDO+OOM dispatch)，不符合這個專案「L2.5 = 快取代理」的定義，取同名會誤導呼叫端以為有快取保證
- 不掛 `OOM` 字尾：因為改完之後它會 RDO優先/OOM兜底都走，不是純 OOM 專屬
- `AsDict` 純粹標示回傳容器形狀（`Dictionary` vs 另兩個 `List(Mail,Topic)`），呼叫端寫法要對應調整這件事比「有沒有快取」更容易讓程式碼編不過，所以用回傳形狀命名優先

`GetBasicMailInfoAsDict` 內部也接上了 RDO優先/OOM兜底：直接借用 `GetBasicMailInfoRdo(fPath, eid, sid)`（不需要另外架一份 RDO 版本），`.ToDictionary(Function(m) m.EntryID, StringComparer.Ordinal)` 轉容器。RDO 分支刻意放在 OOM folder 物化**之前**，RDO 成功就完全不用物化 OOM folder，比照 Task 2b 免-folder版的省物化設計。**已確認 RDO 讀 7 欄比這個 Dictionary 版本原本讀的 5 欄多出的 2 欄(MsgID/SenderEmail)是無害的**——呼叫端(`Form1_Maintab56.vb:966~972`，Tab5 方法B刷新)只用 `Subject/Size/RcvTime/SenderName` 四欄，多出來的欄位單純被忽略。

已更新的地方：函式改名、唯一呼叫點(`Form1_Maintab56.vb:949`)、內部錯誤訊息字串、`GetBasicMailInfoRdo` 開頭註解裡的舊函式名參照，全部同步。**已編譯通過(VS MSBuild.exe，無錯誤)**，`GetFolderBasicByEntryIDL3` 這個名稱在全專案只剩一處刻意保留的「原名」歷史註解，其餘全部清乾淨。

---

## 七、全面改名 — 去掉 `Basic` 前綴 + SQL 表改名（Task 2a 收尾後，同一輪對話緊接著做）

**Simon 的指示**：`GetBasicMailInfo` 系列全部去掉 `Basic` 前綴改叫 `GetMailInfoXXX`；SQL 表 `basic_maillist` 也一起改名 `mailinfo_list`；改完要測寫入/讀出正確性。

**動手前的踩點（grep 全專案）**：一開始以為只有 5 個函式，實際 grep `BasicMailInfo` 抓到 **94 處出現、跨 6 個檔案**（`Module_Outlook.vb`/`Form1_Maintab56.vb`/`Form1.vb`/`Form1_MainTab34.vb`/`Module_SQLite2.vb`/`Form1_OST.vb`），涉及的識別字遠多於原本 5 個函式：`_cacheBasicMailInfo`(記憶體快取字典)、`DbGetBasicMailInfo`/`DbGetBasicMailInfoBatch`/`DbDeleteBasicMailInfoByPath`、`SaveBasicMailInfoInner`/`LoadBasicMailInfoInner`、`InvalidateBasicMailCache`、`PreLoadBasicMailCacheAsync`，以及 SQL 表 `basic_maillist` 本身(在 `Module_SQLite2.vb` 出現近 50 次，含 CREATE TABLE/INSERT/SELECT/索引)。範圍比字面「這幾個」大很多，動手前先攤開全部清單問過 Simon 才做。

**兩個關鍵確認（AskUserQuestion，動手前問清楚）：**
1. **SQL 表改名的資料風險**：`basic_maillist` 是 Simon 真實 `OLAcache.db` 裡已經有資料的表(350,000封信的 metadata)。單純改字串不做 migration，舊 DB 檔裡那個表還是叫 `basic_maillist`，程式啟動後查 `mailinfo_list` 會是空的——**Simon 選擇「直接改名，接受快取重建」**，不寫 migration。
2. **改名範圍**：只改 5 個 `Get` 開頭函式 vs 全部相關識別字一起改——**Simon 選「全部一起改」**，避免新舊命名混雜。

**執行方式**：這種大量(94處)、字面完全一致、不含語意判斷的識別字替換，改用 PowerShell 讀寫檔案做全域字串替換(而非一個個呼叫 Edit)，但**動手前先確認檔案編碼**(`[System.IO.File]::ReadAllBytes` 檢查開頭 3 bytes = `EF BB BF` = UTF-8 with BOM)，讀寫都用 `New-Object System.Text.UTF8Encoding($true)` 保留 BOM，避免破壞中文註解編碼。四個替換規則（互不重疊、可安全用單純字串取代，不需要 regex）：
```
"BasicMailInfo"   → "MailInfo"           (涵蓋 GetBasicMailInfo/DbGetBasicMailInfo/_cacheBasicMailInfo/Save·LoadBasicMailInfoInner/DbDeleteBasicMailInfoByPath/GetBasicMailInfoAsDict/GetBasicMailInfoRdo/GetBasicMailInfoOOM 全部)
"BasicMailCache"  → "MailCache"          (涵蓋 InvalidateBasicMailCache/PreLoadBasicMailCacheAsync)
"basic_maillist"  → "mailinfo_list"      (SQL 表名)
"idx_basic_folder"→ "idx_mailinfo_folder"(SQL 索引名，原本沒特別要求但為了跟表名一致一併改)
```
改完 grep 全專案確認零殘留，**編譯通過(VS MSBuild.exe，無錯誤)**，抽查過改名後的中文註解確認沒有亂碼(BOM/編碼沒被破壞)。

**寫入/讀出正確性驗證（新探針 `PROBE_MAILINFO_DB`，`/autoprobedb:Store|Folder` 觸發）**：

流程：① `GetMailInfoOOM` 拿 ground truth → ② 清 `_cacheMailInfo` 強制重填 → `GetMailInfo`(免-folder版，會走③RDO優先路徑)填快取 → ③ `SaveCachesToDB()` 落地寫進新表 `mailinfo_list` → 清記憶體模擬程式重啟 → ④ `DbGetMailInfo(fPath)` 直接讀 `mailinfo_list` → ⑤ 逐欄比對讀回結果 vs OOM ground truth。

同一個真實資料夾（`寄件備份, 2013~2018, Toshiba` 的 `2013`，51封信）跑出結果：
```
① OOM ground truth: 51 封
② GetMailInfo(填快取,應走RDO優先): 51 封
③ SaveCachesToDB 完成(寫入 mailinfo_list)
④ DbGetMailInfo 讀回: 51 封 | Snap=51
⑤ DB讀回 vs OOM ground truth: 缺漏=0
✓ mailinfo_list 寫入/讀出全部正確
```
**全部通過，零缺漏、零欄位不符**——確認改名沒有破壞任何存取路徑，新表名下的寫入/讀出資料完全正確。

**⚠️ 給下一輪的提醒**：Simon 電腦上原本的 `OLAcache.db` 裡仍然殘留舊的 `basic_maillist` 表(有資料，只是不再被程式碼參照)，沒有清除也沒有 migration。這是 Simon 明確選擇接受的取捨(快取重建一次)，不是遺漏；如果之後要做資料庫瘦身/清理孤兒表，才需要處理這個殘留表。

**這次額外學到的東西**：
- 大量(近百處)、無語意歧義的識別字重新命名，用 PowerShell 檔案級全域字串取代比一個個 Edit 呼叫實際很多，但**動手前一定要先驗證檔案編碼**(尤其是這種滿是繁體中文註解的檔案)，讀寫都指定同一個 encoding 物件，不能讓 PowerShell 用預設值(PowerShell 5.1 預設輸出是 UTF-16 LE with BOM，跟專案原本的 UTF-8 with BOM 不同，若不手動指定會整檔亂碼)
- 改 SQL 表名前一定要想清楚「這台機器上是不是已經有一個用舊表名存了真實資料的 DB 檔案」——這屬於「有真實資料風險的操作」，即使是 Simon 自己機器上的本機 DB，也要先問過確認取捨(migration vs 接受重建)，不能自己假設哪個選項比較好就直接做
- Lambda 委派型別在這個專案要用 `System.Action`／`System.Action(Of T)` 完整限定名，不能用裸 `Action`(跟 `Microsoft.Office.Interop.Outlook.Action` 撞名)；VB 單行 Lambda 語法(`Sub(x) : ... : End Sub`)不合法，多陳述式要用多行 Lambda(`Sub(x)` 換行 ... `End Sub`)
