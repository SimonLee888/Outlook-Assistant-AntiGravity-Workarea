# 消除 FolderPath COM 依賴的補完計畫 (修復中斷損壞)

此計畫旨在修復因中斷導致的 `Form1_Outlook.vb` 損毀（目前有損壞的程式碼片段），並完成 `GetSubtreeToList` 回傳 Tuple 的全面對齊。

## User Review Required

> [!WARNING]
> **損毀修復**：
> `Form1_Outlook.vb` 中的 `GetSortedSubFolders` 目前內容已損毀（出現了 `Dim value` 後就斷掉的情況）。我將會根據之前的版本紀錄重新補回正確的邏輯。

> [!IMPORTANT]
> **對齊舊版代碼**：
> `Form1_Win32API.vb` 中被標記為「待刪區」的部分，我會直接使用 `.Select(Function(x) x.Folder).ToList()` 來讓它能通過編譯，而不會進行深度重構，以符合您「不用去改」的期待。

## Proposed Changes

### [Component] Form1_Outlook.vb

#### [MODIFY] [Form1_Outlook.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_Outlook.vb)
*   **修復 `GetSortedSubFolders`**：補回丟失的 `infoList` 收集與 `OrderBy` 邏輯。
*   **修復 `GetSubtreeToList`**：確保 `DbGetSubFolderIDList` 分支與 BFS 分支的 Tuple 建構正確且完整。

---

### [Component] Form1_SQLite2.vb

#### [MODIFY] [Form1_SQLite2.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_SQLite2.vb)
*   **優化 `RenewCacheAsync`**：根據先前討論，在 Phase 1 取得 Tuple 後，將後續 Phase 2~4 的 `folder.FolderPath` 呼叫全面替換為 Tuple 中的 `item.FolderPath`，達成全流程 0 COM 路徑讀取效能。

---

### [Component] Form1_Win32API.vb (僅編譯對齊)

#### [MODIFY] [Form1_Win32API.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_Win32API.vb)
*   **對齊 `GetMailCountAll_1` 與 `GetMailCountAll_2`**：這兩個函式在 #Region "99 舊版備用" 中。我將簡單呼叫 `.Select(Function(x) x.Folder).ToList()` 來解開 Tuple，僅為了解決編譯錯誤。

---

## Verification Plan

### Automated Tests
*   使用 `viewport` 檢查變更處的語法正確性。
*   檢查 `Form1_Outlook.vb` 是否還有斷頭代碼。

### Manual Verification
*   請使用者開啟專案確認是否能成功編譯。
*   測試 Tab 1 的資料夾統計功能，確認顯示正常。
*   測試 Settings 頁面的 `RenewCache` 功能，觀察 `_dbg` 是否顯示 Phase 2 比對的速度提升。
