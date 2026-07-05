# MAPI Proptag 常數集中化重構計畫

## Context

目前 `Module_Outlook.vb` 和 `Module_Win32API.vb` 裡，MAPI proptag（DASL URL 形式，如
`"http://schemas.microsoft.com/mapi/proptag/0x36020003"`）是用「每個函數各自宣告一份 local `Const`」
的方式重複寫。清點後規模比預期大：

- **40 個 local `Const` 宣告**，對應 **16 種不同的數值**（15 個標準 PR_ proptag + 1 個具名屬性
  `PidLidSmartNoAttach`），分散在 2 個檔案。
- 同一個 proptag 數值被取了不同名字，例如 `0x0E080003` (PR_MESSAGE_SIZE) 同時叫
  `PR_SIZE_LONG` / `PR_MESSAGE_SIZE` / `PR_SIZE` / `DASL_SIZE`（4 種名字）；
  `0x0E1B000B` (PR_HASATTACH) 同時叫 `PR_HASATTACH` / `DASL_HASATT`。
- 有兩組 Unicode/ANSI 變體 (`...001E` vs `...001F`) 共用同一個常數名稱
  (`PR_SENDER_EMAIL`、`PR_INTERNET_MESSAGE_ID`)，只能看 hex 尾碼分辨，容易看錯。
- 其中一個名字其實取錯了：`PR_SENDER_EMAIL` 的 MAPI 官方名稱應為 `PR_SENDER_EMAIL_ADDRESS`
  (`PidTagSenderEmailAddress`)。

這些風險（改錯值、漏改、看錯變體、非官方命名）在 40 處分散宣告下會隨時間累積，值得一次性集中管理。
本次重構是**純粹搬移常數定義，不改變任何 proptag 數值或程式邏輯**，因此行為應完全不變。

## 命名規則

採用 **MAPI 官方標準名稱**（mapitags.h / MS-OXPROPS 裡的正式名字），而非「目前出現頻率最高」的名字——
官方名字可以直接對照 MS 官方文件查到，之後維護不用猜是不是自創的。

- 同數值只保留一個名字，選官方名（多數現有名字已經是官方名，直接沿用）：
  - `PR_MESSAGE_SIZE` (0x0E080003) — 取代 `PR_SIZE_LONG`/`PR_SIZE`/`DASL_SIZE`
  - `PR_HASATTACH` (0x0E1B000B) — 取代 `DASL_HASATT`
  - `PR_MESSAGE_SIZE_EXTENDED` (0x0E080014) — 取代 `PR_SIZE_EX_STR`/`PR_SIZE_EXTENDED`
  - `PR_CONTENT_COUNT` (0x36020003)
  - `PR_SUBJECT` (0x0037001F)
  - `PR_MESSAGE_DELIVERY_TIME` (0x0E060040)
  - `PR_SENDER_NAME` (0x0C1A001F)
  - `PR_CONVERSATION_TOPIC` (0x0070001E)
  - `PR_LOCAL_COMMIT_TIME_MAX` (0x670A0040)
  - `PR_CONTAINER_CLASS` (0x3613001E)
  - `PR_SUBFOLDERS` (0x360A000B)
- **改名**：`PR_SENDER_EMAIL` → **`PR_SENDER_EMAIL_ADDRESS`**（官方正式名稱是這個，現有名字是自創的）。
- Unicode/ANSI 變體用 `_W`/`_A` 尾碼區分（這是 MAPI 標頭檔本身的慣例，non-invented）：
  - `PR_SENDER_EMAIL_ADDRESS_W` (0x0C1F001F，多數呼叫點用這個) / `PR_SENDER_EMAIL_ADDRESS_A` (0x0C1F001E)
  - `PR_INTERNET_MESSAGE_ID_W` (0x1035001F，多數呼叫點用這個) / `PR_INTERNET_MESSAGE_ID_A` (0x1035001E)
- 具名屬性（非標準 PR_ tag，`/mapi/id/{GUID}/hex` 定址）維持 `DASL_` 前綴以區別於固定 proptag：
  - `DASL_SMARTNOATTACH` (`PidLidSmartNoAttach`, PSETID_Common/0x8514/PT_BOOLEAN)

## 實作步驟

### 1. 在 `Module_Outlook.vb` 檔案最後新增一個 Region
不開新檔案。在 `Module_Outlook.vb` 結尾加一個區塊（如 `#Region "MAPI Proptag 常數"`），
把上述 16 個常數以 `Const` 宣告在 Module 層級（VB 同一 Module 內成員互相可見，不用加前綴）。
每個常數保留原本就有的中文/英文用途註解（如 `' PT_LONG`），不新增額外說明文字。

### 2. 改寫呼叫點
`Module_Outlook.vb`（本檔內，約 26 處：23 個 proptag + 2 個 SMARTNOATTACH + 1974 行內聯 COLS）
與 `Module_Win32API.vb`（14 處：13 個 proptag + 1 個 SMARTNOATTACH）：
- 刪除每個函數內的 local `Const PR_XXX As String = "..."` 那一行。
- 沿用不變名字的識別字不用改；用到被改名常數的地方
  （`PR_SIZE_LONG`/`PR_SIZE`/`DASL_SIZE`/`DASL_HASATT`/`PR_SIZE_EX_STR`/
  `PR_SENDER_EMAIL`→`PR_SENDER_EMAIL_ADDRESS`(_A/_W)/`PR_INTERNET_MESSAGE_ID`(_A/_W)）
  同步改成新名稱。
- `Module_Win32API.vb` 在 Module 層級的常數要跨檔可見，需要在 `Module_Outlook.vb` 那邊宣告時
  確保沒有 `Private` 修飾（VB Module 內 `Const` 預設就是可從其他檔案存取，不用額外處理，
  只要不是明確標了 `Private`）。
- 特別處理 [Module_Outlook.vb:1974](Module_Outlook.vb:1974) 的 `COLS` 字串：目前是把 3 個 proptag
  裸字串直接串在一起，改成用 `$"Name, EntryID, {PR_SUBFOLDERS}, {PR_CONTAINER_CLASS}, {PR_CONTENT_COUNT}"`
  組出來，維持字串內容完全一致。

### 3. 清死碼
`Module_ToBeDelete.vb` 裡的 `GetMailCountByLINQNew`（第 162–177 行）已確認只有自我遞迴呼叫，
唯一外部呼叫點在第 251 行已被註解掉——整個函數直接刪除，不搬進新常數區塊。

## 協作注意事項

- 這次會動到 **production 檔案** `Module_Outlook.vb`、`Module_Win32API.vb`、`Module_ToBeDelete.vb`。
  開始改之前會先提醒您，若這幾個檔案在 VS 裡有未存檔的修改，請先存檔或關閉，避免本機檔案
  last-writer-wins、彼此覆蓋。
- 建議改完後您在 VS 端也 `git commit` 一次作為安全點。

## 驗證方式

- 這是純搬移常數、不改數值/邏輯的重構，理論上行為零變化。
- 用 VS 的 `MSBuild.exe` 編譯（不能用 `dotnet build`，COM 參考的關係）確認能過編譯，
  這一步就能抓到「漏改成新名稱」或「打錯常數名」這類問題。
- 建議您實際跑一次 Tab1（資料夾樹/郵件數統計）驗證 `GetMailCount`、`GetSubtreeRdoBatch`、
  `GetLiveFolderSnapOOM` 這幾個直接用到 `PR_CONTENT_COUNT` 的路徑，以及 Tab3 SmartNoAttach
  相關功能（用到 `DASL_SMARTNOATTACH`）數字/行為沒有跑掉。
