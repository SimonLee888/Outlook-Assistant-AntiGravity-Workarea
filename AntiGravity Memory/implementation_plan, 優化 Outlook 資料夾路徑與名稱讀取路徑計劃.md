# 優化 Outlook 資料夾路徑與名稱讀取路徑計劃

目前系統在某些地方仍會呼叫昂貴的 `folder.FolderPath` (COM 呼叫約 1-2ms)，或者在已有 `fPath` 的情況下仍重複讀取。本計劃核心是利用「路徑拼接」取代「路徑讀取」，並將路徑快取在 UI 節點中。

## User Review Required

> [!IMPORTANT]
> **TreeNode.Tag 結構變更**：
> 我們將把所有 `TreeView` (Tab1~Tab5) 的 `TreeNode.Tag` 從單純的 `Outlook.Folder` 物件改為 `(Folder As Outlook.Folder, FolderPath As String)`。
> 這會影響到所有從 `Tag` 讀取資料的地方，我會一併修正。

> [!TIP]
> **路徑拼接規則**：
> 對於子資料夾，其路徑將統一由 `parentPath & "\" & folder.Name` 組合而成。這能省下大量的 COM 溝通開銷。

## Proposed Changes

---

### [Component] Form1_Outlook.vb (底層邏輯層)

#### [MODIFY] [Form1_Outlook.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_Outlook.vb)
- **`GetSortedSubFolders`**: 
    - 修改參數邏輯，確保當 `fPath` 缺失時，優先使用 `folder.Name` 進行日誌記錄，只有在真正需要 Cache Key 時才讀取 `.FolderPath`。
    - 在展開子資料夾時，利用傳入的 `fPath` 預先拼接好子路徑，避免內層再次觸發路徑讀取。
- **`GetCachedXxx` 系列**: 
    - 再次確認所有調用處是否都傳入了 `fPath`。

---

### [Component] Form1_MainTabs.vb (UI 邏輯層)

#### [MODIFY] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_MainTabs.vb)
- **`BuildBfsFolderTree`**: 
    - 目前已實現拼接邏輯，將進行微調以確保完全不依賴內層的 `.FolderPath`。
- **`EnterSelectedFolder` / `ComputeFolderSize`**: 
    - 配合 `Tag` 變更，從 Tuple 中直接解構出 `fPath`。

---

### [Component] Form1.vb (全域初始化層)

#### [MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1.vb)
- **`LoadStoreToTreeView`**:
    - 在建立 Root 節點時，呼叫一次 `store.GetRootFolder().FolderPath` 並存入 Tuple。
- **`LoadSubFolderToTreeView`**:
    - 從父節點的 Tuple 取得 `parentPath`，並與 `subFolder.Name` 拼接成子路徑，存入新節點的 Tuple。

---

### [Component] Form1_SQLite2.vb (持久化層)

#### [MODIFY] [Form1_SQLite2.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_SQLite2.vb)
- **`RenewAttachMailListAsync`**:
    - 確認 `fPath` 的傳遞路徑，確保不重複讀取。

## Open Questions

- **根資料夾特殊性**：`\\User@domain.com` 這種 Store Root 的 `Name` 有時與路徑不一致（例如 Name 是 "Simon Lee" 但路徑是 "\\simon@abc.com"）。目前的策略是 Root 呼叫一次 `.FolderPath`，隨後的子孫全部用拼接。這應該能解決問題，您是否認同？

## Verification Plan

### Automated Tests
- 觀察 `_dbg`輸出，確認在展開樹狀結構或執行統計時，不再出現重複的 `.FolderPath` 讀取警告（如果有設定）。
- 測試各個 Tab 的資料夾跳轉功能（Enter Folder）是否正常，確保 `Tag` 解析正確。

### Manual Verification
- 點擊 Tab1/Tab2 的資料夾，確認郵件數與路徑顯示正確。
- 檢查 SQLite 資料庫，確認存入的路徑依然完整且正確。
