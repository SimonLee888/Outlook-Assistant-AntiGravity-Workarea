# 啟動效能優化計劃 (更新版)

根據 Log 紀錄分析，程式啟動的前 1 秒內有幾個明顯的耗時點。本計劃旨在調整初始化順序與邏輯，減少 UI 執行緒的阻塞。

## 耗時點分析與優化對策

| 模組/函數 | 預計耗時 | 優化對策 (Phase 1) |
| :--- | :--- | :--- |
| `GetSortedStores` | ~130ms | 優先嘗試使用 Redemption (RDO) 獲取 Store，若未初始化則保留 OOM。 |
| `GetSortedSubfolders` | ~150ms | **實作跳過過濾**：若勾選「顯示全部」，直接不跑 `IsMailFolder` (COM 呼叫)。 |
| `DebugForm_Load` | ~120ms | **UI 延遲繪製**：計算邏輯移至 `Shown` 事件，訊息先由 Queue 緩衝。 |
| `InitListViews / Tabs` | ~150ms | **暫緩 (Phase 2)**：Simon 標記 Wait，待下階段執行。 |

## 優化實作細節

### 1. 資料夾遍歷優化 (Skip IsMailFolder)
*   **[MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1.vb)**
    *   在 `GetSortedSubFolders` 與 `GetSubFolderList` 中，重構 `If` 結構，確保當 `checkIncludeAllFolders.Checked = True` 時，完全不觸發 `IsMailFolder` 內部的屬性讀取。

### 2. DebugForm 啟動流程優化
*   **[MODIFY] [DebugForm.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/DebugForm.vb)**
    *   將 `DebugForm_Load` 中的座標計算、搜尋面板寬度設定等 UI 代碼移至 `Shown` 事件。
    *   確保在 `Shown` 觸發前，`AddMessage3` 的資料僅會在 Queue 中緩慢積壓，不會強制調用 `Refresh` 或 `Update`。

### 3. Store 清單取得效能 (待確認)
*   探索在 `Form1_Load` 早期是否能藉由 `_rdo.Stores` 快速取得清單。

---

## 暫緩執行清單 (Simon 標記：記下來，下一步再做)

*   **Lazy Initialization**：`InitListViews` 與 `InitTabXUI` 的延遲載入 (Tab2~Tab5)。
*   **TreeView 掃描優化**：移除 `GetAllTreeViews(Me)`，改為靜態列表操作。

## 待確認問題

> [!IMPORTANT]
> **關於 IsMailFolder 的過濾邏輯**：我將修改為只有在 `checkIncludeAllFolders` 為 `False` (預設不顯示全部) 時才會進行類型檢查。這樣在「顯示全部」模式下，啟動速度會得到最大化提升。

## 驗證計劃

### 自動與手動驗證
*   查看 DebugForm 中的啟動 Log，比對優化前後 `GetSortedSubfolders` 的豪秒數。
*   檢查 `DebugForm` 初始化時，主 UI 的卡頓感是否減輕。
