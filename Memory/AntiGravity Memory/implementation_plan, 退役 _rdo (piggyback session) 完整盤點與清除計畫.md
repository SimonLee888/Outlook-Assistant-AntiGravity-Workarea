# 退役 `_rdo` (piggyback session) 完整盤點與清除計畫

## 背景

`_rdo` 是 piggyback 在 Outlook 原有 MAPI session 上的 `Redemption.RDOSession`，效能受限於共用 session 的序列化。所有生產路徑已切換到 `_rdo2`（獨立 session），`_rdo` 可以完全退役。

以下依檔案分區，列出所有包含 `_rdo`（排除 `_rdo2`）的引用，區分 **活程式碼**（必須改）與 **純註解**（可保留歷史紀錄但視需求更新）。

---

## 盤點摘要

| 檔案 | 活程式碼引用 | 純註解引用 | 改法 |
|------|:-----------:|:---------:|------|
| [Module_Outlook.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual Studio/Visual Studio 18 (2026)/Outlook Assistant - (AntiGravity測試區)/Module_Outlook.vb) | **11** | 7 | 刪欄位宣告、重構 Init/Release |
| [Module_Win32API.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual Studio/Visual Studio 18 (2026)/Outlook Assistant - (AntiGravity測試區)/Module_Win32API.vb) | **16** | 24 | 改用 `_rdo2` 或刪除死碼函數 |
| [Form1.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual Studio/Visual Studio 18 (2026)/Outlook Assistant - (AntiGravity測試區)/Form1.vb) | **2** | 0 | 改判 `_rdo2` + 移除 COM 釋放行 |
| [Form1_MainTab12.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual Studio/Visual Studio 18 (2026)/Outlook Assistant - (AntiGravity測試區)/Form1_MainTab12.vb) | 0 | 1 | 僅註解，可保留 |
| **合計** | **29** | 32 | |

---

## 詳細修改清單

### 元件 1：Module_Outlook.vb（核心宣告與生命週期）

#### [MODIFY] [Module_Outlook.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual Studio/Visual Studio 18 (2026)/Outlook Assistant - (AntiGravity測試區)/Module_Outlook.vb)

**1-1. 欄位宣告 (L43)** — 活程式碼
```vb
Private _rdo As Redemption.RDOSession = Nothing
```
→ **刪除**此行。`_rdo2` 及其 store 快取已足夠。

**1-2. InitRdoSession() (L150-168)** — 活程式碼
- L154-156: 已被註解的舊 piggyback 初始化（純註解，保留歷史）
- L160: 註解說明（純註解，保留）
- L162: `Dim unused = InitRdoSessionWithoutEULA()` — 這是呼叫入口
- L165: `TryMarshalRelease(_rdo)` — 活程式碼
- L166: `_rdo = Nothing` — 活程式碼

→ **整個 `InitRdoSession()` Sub 需要重構**：
  - Catch 區塊不再需要釋放 `_rdo`，改為只釋放 `_rdo2`
  - 或者，如果 `InitRdoSession()` 已不被外部呼叫（經確認，搜尋不到呼叫端），可以考慮**整個刪除**

**1-3. InitRdoSessionWithoutEULA() (L170-205)** — 活程式碼（核心！）
- L177: `If _rdo IsNot Nothing Then Return` — 改為判 `_rdo2`
- L189-192: 
  ```vb
  session.MAPIOBJECT = _olNS.MAPIOBJECT
  _rdo = session
  _dbg(" ├ _rdo init OK", ...)
  ```
  → **刪除這 3 行**（piggyback session 的建立邏輯）。`_rdo2` 的獨立 Logon 已在 L195-199 完成。
- L202: `_rdo = Nothing` — 刪除

> [!IMPORTANT]
> `InitRdoSessionWithoutEULA` 目前建兩條 session（`_rdo` piggyback + `_rdo2` 獨立）。退役 `_rdo` 後，L186-192 的 piggyback session 建立邏輯應整段移除，只保留 L195-199 的 `_rdo2` 獨立 session 建立。同時 L177 的 guard 改判 `_rdo2`。

**1-4. ReleaseRdoSession() (L309-314)** — 活程式碼
```vb
' L310: 註解 — 保留歷史
' L311-313: 活程式碼
If _rdo IsNot Nothing Then
    Dim r As Object = _rdo : TryMarshalRelease(r) : _rdo = Nothing
    _dbg(" ├ _rdo 釋放完成")
End If
```
→ **刪除** L309-314 整個 `_rdo` 釋放區塊

**1-5. objFolder2odoFolder() (L2922-2925)** — 活程式碼（但**已無呼叫端**）
```vb
Private Function objFolder2odoFolder(objFolder As Folder) As Redemption.RDOFolder
    If _rdo Is Nothing OrElse objFolder Is Nothing Then Return Nothing
    Return _rdo.GetFolderFromID(objFolder.EntryID, objFolder.StoreID)
End Function
```
→ 經搜尋確認無任何呼叫端，屬**死碼**，**整個函數刪除**

**1-6. 純註解**（建議保留歷史但不強制更動）
- L49: `_rdoFastPath` 註解 — 保留
- L154-156: 舊 piggyback 初始化的註解 — 保留
- L160-161: 效能說明 — 保留
- L1634, L1636: `_rdo → _rdo2` 遷移說明 — 保留
- L2057, L2154: 移除舊 tier 的說明 — 保留

---

### 元件 2：Module_Win32API.vb（生產函數與 Spike 測試）

#### [MODIFY] [Module_Win32API.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual Studio/Visual Studio 18 (2026)/Outlook Assistant - (AntiGravity測試區)/Module_Win32API.vb)

**2-1. GetMailSizeL3() (L253-284)** — 活程式碼區塊但**整個函數是死碼**（L253 註解確認無呼叫端）
- L268: `If _rdo IsNot Nothing Then`
- L273: `_rdo.GetMessageFromID(mail.EntryID, storeId)`
→ **整個函數已是死碼**，建議保留原狀或整函數刪除。如果想保留函數但退役 `_rdo`，則把 `_rdo` 改為 `_rdo2`。

**2-2. RdoPreloadAttach_1() (L543-596)** — 活程式碼但**呼叫端已被註解**
- L550: `If _rdo Is Nothing OrElse ...`
- L570: `_rdo.GetMessageFromID(mail.EntryID)`
→ 呼叫端 [Form1_MainTab34.vb L266](file:///d:/Users/Simon/Dropbox/私人文件/Visual Studio/Visual Studio 18 (2026)/Outlook Assistant - (AntiGravity測試區)/Form1_MainTab34.vb#L266) 已被 `'` 註解。此函數等同死碼。
→ 建議：**整個函數刪除**，或改 `_rdo` → `_rdo2` 留作備用。

**2-3. RdoPreloadAttach_2() (L598-653)** — 活程式碼但**呼叫端已被註解**
- L603: `If _rdo Is Nothing OrElse ...`
- L625: `_rdo.GetMessageFromID(mail.EntryID)`
→ 呼叫端 [Form1_MainTab34.vb L267](file:///d:/Users/Simon/Dropbox/私人文件/Visual Studio/Visual Studio 18 (2026)/Outlook Assistant - (AntiGravity測試區)/Form1_MainTab34.vb#L267) 已被 `'` 註解。此函數等同死碼。
→ 建議：**整個函數刪除**，或改 `_rdo` → `_rdo2`。

**2-4. RdoPreloadAttach_3() (L655-750)** — 活程式碼但**呼叫端已被註解**
- L666: `If _rdo Is Nothing OrElse ...` — `_rdo` 僅作可用性旗標
- L687: `sess.Logon(_rdo.ProfileName, ...)` — 借 `_rdo.ProfileName` 取 profile 名稱
→ 呼叫端 [Form1_MainTab34.vb L268](file:///d:/Users/Simon/Dropbox/私人文件/Visual Studio/Visual Studio 18 (2026)/Outlook Assistant - (AntiGravity測試區)/Form1_MainTab34.vb#L268) 已被 `'` 註解。此函數等同死碼。
→ 建議：**整個函數刪除**。若要保留，改判 `_rdo2`、`_rdo.ProfileName` → `_rdo2.ProfileName` 或 `_olNS.CurrentProfileName`。

**2-5. SpikeParallelReadBenchmark() (L1199-1389)** — 拋棄式 spike 測試
- L1210-1211: `If _rdo Is Nothing Then Await InitRdoSessionWithoutEULA()` / `If _rdo Is Nothing Then ...`
- L1214: `CallByName(_rdo, "ProfileName", ...)`
→ **拋棄式 spike**，無呼叫端，已完成使命。建議**整個函數刪除**。

**2-6. SpikeResolveFormCompare() (L1390-1508)** — 拋棄式 spike 測試
- L1398-1399: 判 `_rdo` 初始化
- L1401: `CallByName(_rdo, "ProfileName", ...)`
- L1405-1410: 用 `_rdo.Stores` 走訪
- L1448-1453: 用 `_rdo.GetMessageFromID`
- L1461-1465: 用 `_rdo` store-scoped
→ **拋棄式 spike**，無呼叫端。建議**整個函數刪除**。

**2-7. SpikeBodyResolveCompare() (L1509-1620左右)** — 拋棄式 spike 測試
- L1519-1520: 判 `_rdo`
- L1522: `CallByName(_rdo, "ProfileName", ...)`
- L1526-1532: 用 `_rdo.Stores` 走訪
- L1571-1575: 用 `_rdo` store-scoped
→ **拋棄式 spike**，無呼叫端。建議**整個函數刪除**。

**2-8. 純註解**（建議保留歷史）
- L237: `_rdo 未就緒時自動跳過此層`
- L948: `_rdo → _rdo2` 遷移說明
- L1393-1394, L1405, L1448, L1461, L1476, L1511, L1513, L1526, L1571, L1587, L1917: 各 spike 內的註解
→ 如果 spike 函數整個刪除，這些註解自然一起消失。如果保留死碼則做為歷史紀錄。

---

### 元件 3：Form1.vb（UI 事件處理）

#### [MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual Studio/Visual Studio 18 (2026)/Outlook Assistant - (AntiGravity測試區)/Form1.vb)

**3-1. FormClosing handler (L318)** — 活程式碼
```vb
If _rdo IsNot Nothing Then Marshal.FinalReleaseComObject(_rdo)
```
→ **刪除此行**。下一行 L319 `ReleaseRdoSession()` 已負責 `_rdo2` 的完整釋放。

**3-2. CheckRDO_CheckedChanged (L1172)** — 活程式碼
```vb
If _rdo Is Nothing Then Dim unused = InitRdoSessionWithoutEULA()
```
→ **改為** `If _rdo2 Is Nothing Then Dim unused = InitRdoSessionWithoutEULA()`

---

### 元件 4：Form1_MainTab12.vb（純註解）

#### [Form1_MainTab12.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual Studio/Visual Studio 18 (2026)/Outlook Assistant - (AntiGravity測試區)/Form1_MainTab12.vb)

- L722: `' 效能原理：有 RDO → GetMailCountAllOOM 內部呼叫 _rdo.TotalItemCount...`
→ **純歷史註解**，建議保留不動。

---

## Open Questions

> [!IMPORTANT]
> **死碼 Spike 函數的處置**：`SpikeParallelReadBenchmark`、`SpikeResolveFormCompare`、`SpikeBodyResolveCompare` 三個拋棄式 spike 函數加起來約 420 行，已無呼叫端。你想要：
> - (A) 整段刪除（推薦，清爽乾淨）
> - (B) 保留但把 `_rdo` 改為 `_rdo2`（保留測試基礎設施）

> [!IMPORTANT]
> **已廢棄的 `RdoPreloadAttach_1/2/3` 處置**：這三個函數的呼叫端在 Form1_MainTab34.vb 都已被註解。你想要：
> - (A) 整段刪除（推薦，加起來約 210 行）
> - (B) 改 `_rdo` → `_rdo2` 保留備用

> [!IMPORTANT]
> **已廢棄的 `GetMailSizeL3` 和 `objFolder2odoFolder` 處置**：兩者都已無呼叫端。你想要：
> - (A) 整段刪除
> - (B) 改 `_rdo` → `_rdo2` 保留備用

> [!IMPORTANT]
> **`InitRdoSession()` Sub (L150-168) 處置**：搜尋不到外部呼叫端，且 `InitRdoSessionWithoutEULA()` 才是真正的入口。你想要：
> - (A) 整段刪除
> - (B) 保留但改為只呼叫 `InitRdoSessionWithoutEULA()` 且 Catch 裡只清 `_rdo2`

---

## 風險評估

> [!WARNING]
> `InitRdoSessionWithoutEULA()` 重構時需注意：
> 1. 移除 `session.MAPIOBJECT = _olNS.MAPIOBJECT` 和 `_rdo = session` 後，**New RDOSession() 只會被 `_rdo2` 的 Logon 使用**，因此 L187 的 `Await Task.Run(Sub() session = New Redemption.RDOSession())` 可以整段移除（`_rdo2` 在 L196 自己 New）。
> 2. AutoDismissRdoEULA 仍需保留（`_rdo2` 初始化時也可能彈 EULA dialog）。
> 3. Guard `If _rdo IsNot Nothing Then Return` 改為 `If _rdo2 IsNot Nothing Then Return`。

---

## Verification Plan

### 編譯驗證
- 修改完成後執行 `dotnet build` 或 Visual Studio Build，確認無編譯錯誤
- 全域搜尋 `\b_rdo\b`（排除 `_rdo2`），確認不再有殘留引用

### 功能驗證
- 啟動程式，勾選 CheckRDO checkbox，確認 `_rdo2` 正常初始化
- 切換 Tab3/Tab5 確認附件/內文讀取走 `_rdo2` 路徑正常運作
- 關閉程式，確認 ReleaseRdoSession 正常釋放 `_rdo2`

### 複檢所有修改點確認正確、複檢修改點前後是否遺留多餘程式碼
