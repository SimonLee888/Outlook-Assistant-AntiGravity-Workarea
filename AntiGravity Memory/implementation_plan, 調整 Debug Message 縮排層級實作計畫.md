# 調整 Debug Message 縮排層級實作計畫

根據使用者的要求，我們將調整專案中 `Dbg` 訊息的縮排，以反映函數的呼叫深度。

## 目的
提升 Debug Log 的可讀性，讓使用者能透過縮排一眼看出函數呼叫的層級關係。

## 縮排規則

| 層級 | 觸發位置 | 縮排格式 | 範例 |
| :--- | :--- | :--- | :--- |
| **Level 0 (頂級)** | UI 事件 (Click, Select) 或 Form 事件 | `無` | `Dbg("開始")` |
| **Level 1 (次級)** | 由 Level 0 直接呼叫的核心函數 | ` ├ ` | `Dbg(" ├ 開始")` |
| **Level 2 (三級)** | 迴圈內部、由 Level 1 呼叫的輔助函數 | `    ├ ` | `Dbg("    ├ 開始")` |

> [!IMPORTANT]
> **規則細節：**
> *   **不刪減、不新增、不改變文字內容**：僅在原始字串的前端調整縮排前綴。
> *   **保留開發註解**：原本的 debug 歷程與思考註解將完整保留，並在我的修改處標記 `by Gemini, 2026/04/10`。
> *   **不確定則跳過**：若呼叫關係不明確，暫時維持現狀。
> *   **小塊寫入 (Chunked Edits)**：確保檔案修改的安全與穩定。

---

## 預計修改清單 (初步分層)

### 1. Form1_Outlook.vb [L1/L2]
*   **Level 1 (` ├ `)**:
    *   `InitOutlookNamespace`, `InitRdoSession`, `InitRedemptionSessionWithoutDeclaration`
    *   `GetSubFolderList` (雖然有很多地方叫它，但它是掃描的核心 L1 起點)
    *   `GetCachedMailWithAttachs` (已手動加了一部分)
    *   `GetCachedAttachFilename`
    *   `PreloadAttachByRDOAsync1`
*   **Level 2 (`    ├ `)**:
    *   `AutoDismissRedemptionDialog` (由 L1 初始化呼叫)
    *   `GetSortedStores` (由 L1 初始化呼叫)
    *   `GetSortedSubFolders` (由 L1 的 BFS 內部呼叫時)
    *   `GetMailCount`, `GetFolderCount`, `GetFolderSize` (由 Layer 2.5 呼叫時)
    *   `GetSubFolderList` 內部的 `While` 迴圈。

### 2. Form1_MainTabs.vb [L0/L1/L2]
*   **Level 0 (不縮排)**:
    *   `Button3_Click`, `Button4_Click`, `Button5_Click` (搜尋按鈕)
    *   `TreeView1_AfterSelect`, `SimTree2_AfterSelect`, `ListView3_ColumnClick`
*   **Level 1 (` ├ `)**:
    *   `FilterBySize`, `FilterByAttachDetailsAsync`, `ShowResultTab3`
    *   `UpdateChart2forDefultView`, `ShowResultTab2`
    *   `OpenMailByEntryID`
*   **Level 2 (`    ├ `)**:
    *   `FilterByAttachDetailsAsync` 內的 `For Each` 迴圈進度訊息。

### 3. Form1_SQLite2.vb [L1/L2]
*   **Level 1 (` ├ `)**:
    *   `InitDatabase`, `CloseDatabase`, `SaveCachesToSQLiteAsync`, `LoadCachesFromSQLiteAsync`, `RenewCacheAsync`
*   **Level 2 (`    ├ `)**:
    *   `CleanupOrphanFolderPath`, `GetDatabaseSummary`
    *   所有的 `Inner` 寫入/載入子函數。

---

## 驗證計畫

### 手動與靜態檢查
1. **程式碼完整性**：修改後利用 `view_file` 複檢，確保 `Dbg` 語法正確（括號、逗號等）且未誤刪代碼。
2. **註解保留**：檢查原有的 debug 歷程註解是否完好。
3. **前綴正確性**：確認 ` ├ ` 與 `    ├ ` 的空格數量符合使用者要求。

## 開放問題
1. **空格數確認**：關於 Level 2，您提到「三個空白，加上 '    ├ '」，這在視覺上看起來是 4 個空格（3個新空格 + 1個原本的空格？）。我將統一定義為：
   *   Level 1: ` ├ ` (1 空格 + ├ + 1 空格)
   *   Level 2: `    ├ ` (4 空格 + ├ + 1 空格)
   如果您有不同看法請告知。
