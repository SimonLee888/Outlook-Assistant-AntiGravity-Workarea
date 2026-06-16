# SimTree4 雙重模式拆分重構 — 局部補全實作計畫 (v2)

根據您的要求，我們將大幅簡化本次的重構範圍，**僅執行第 3 點、第 4 點**以及**在 `HandleTvKeyDown` 中以 `GetCurrentTv()` 限制 `ForceLv1Refresh()` 的呼叫條件**。其餘 Phase 1 與 Phase 2 項目皆不予改動。

## 使用者審查要求

> [!IMPORTANT]
> 1. **註解保留與標註**：所有修改點均保留原有的除錯歷程註解，新添加的註解統一加上標記：`' by Gemini 3.5 Flash, 2026/05/29`。
> 2. **小塊寫入 (Chunked Edits)**：本次修改涉及兩個檔案，將採用多個精準的 `ReplacementChunk` 進行小範圍替換，降低程式碼損壞的風險。
> 3. **改後立即複檢**：完成修改後，會主動使用 `view_file` 讀取並確認修改點的邏輯是否對齊。

## 本次修改細項與提案

---

### 1. 【1-F】修改 `Lv4_ColumnClick` 欄位排序 (對齊 3.)
* **位置**：[Form1_MainTab345.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_MainTab345.vb#L830-L880)
* **調整內容**：
  * **1-F-1**：增加 Guard Clause `If Not LvSearch4.Visible Then Return`。改讀 `_tv4SelectedTopicMailList` 替代 `SimTree4.SelectedNode?.Tag`。
  * **1-F-2**：將排序後產生的新 List 寫回 `_tv4SelectedTopicMailList`，不再觸碰 `SimTree4.SelectedNode.Tag`。
  * **1-F-3**：補上對應的修改註解與防禦設計說明（`' by Gemini 3.5 Flash, 2026/05/29`）。

---

### 2. 【1-G】修改 `RefreshLv4MailsAsync` (對齊 4.)
* **位置**：[Form1_MainTab345.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_MainTab345.vb#L1240-L1286)
* **調整內容**：
  * **1-G-1 & 1-G-2**：改讀並更新寫回 `_tv4SelectedTopicMailList`，不再去讀取 `SimTree4.SelectedNode?.Tag`。
  * **1-G-3**：加上修改說明註解（`' by Gemini 3.5 Flash, 2026/05/29`）。

---

### 3. 修改 `HandleTvKeyDown` 限制 `ForceLv1Refresh` 呼叫條件 (新增)
* **位置**：[Form1.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1.vb#L1285-L1289)
* **現狀**：按 F5 時，無論在哪個 tab 都會無條件執行 `Await ForceLv1Refresh()`。
* **調整內容**：
  * 呼叫 `GetCurrentTv()` 來取得當前活動中的 `SimTree`。
  * 限制僅當 `GetCurrentTv() Is SimTree1` (即處於 Tab1 資料夾統計) 時，才執行 `Await ForceLv1Refresh()`。
  * 加入修改說明註解（`' by Gemini 3.5 Flash, 2026/05/29`）。

---

## 驗證計畫

### 手動與功能性驗證
1. **右側列表排序 (ColumnClick)**：點擊右側 `ListView4` 欄位，確認排序能正確在 `_tv4SelectedTopicMailList` 上運作。
2. **右側列表重整 (F5)**：在 `ListView4` 上按 F5，確認能正常調用 `RefreshLv4MailsAsync` 改讀 `_tv4SelectedTopicMailList`，且郵件資訊成功刷新。
3. **左側 SimTree4 按 F5**：在 `SimTree4` (資料夾樹) 上按 F5，確認僅刷新該樹，不會去呼叫 `ForceLv1Refresh()`。
4. **左側 SimTree1 按 F5**：在 `SimTree1` (Tab1 統計樹) 上按 F5，確認會正確同時呼叫 `ForceTvRefresh(SimTree1)` 與 `ForceLv1Refresh()`。
