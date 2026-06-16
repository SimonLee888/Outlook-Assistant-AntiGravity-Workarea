# 啟動效能優化計劃

根據 Log 紀錄分析，程式啟動的前 1 秒內有幾個明顯的耗時點。本計劃旨在調整初始化順序與邏輯，減少 UI 執行緒的阻塞。

## 耗時點分析 (由高至低)

| 模組/函數 | 預估耗時 | 原因分析 |
| :--- | :--- | :--- |
| `GetSortedSubfolders` | ~150ms | 遍歷子資料夾時，每筆都會觸發 `IsMailFolder` (COM 呼叫讀取屬性)。 |
| `InitListViews` | ~150ms | 一次性初始化 5 個 ListView，包含欄位定義、雙緩衝設定與事件掛載。 |
| `GetSortedStores` | ~130ms | 同步呼叫 Outlook OOM 取得所有 Store 清單。 |
| `CheckDebug_CheckedChanged` | ~120ms | 開啟 `DebugForm` 時進行了複雜的 UI 佈局校算與控制項初始化。 |

## 提出的優化變更

### 1. 延遲初始化 (Lazy Initialization)
*   **[MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1.vb)**
    *   修改 `InitListViews`：僅初始化 `ListView1`。其餘 Tab2~Tab5 的 ListView 改在 `TabControl1_SelectedIndexChanged` 第一次切換時才初始化。
    *   修改 `InitTabXUI`：同樣改為延遲載入。目前 `InitLookAndFeel` 呼叫了所有 `InitTabXUI`，這在啟動時造成大量 UI 物件建立。

### 2. 優化資料夾過濾邏輯
*   **[MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1.vb)**
    *   在 `GetSortedSubFolders` 與 `GetSubFolderList` 中，如果 `checkIncludeAllFolders.Checked` 為 `True`，則**跳過** `IsMailFolder` 判斷。這能省下大量的 COM 屬性讀取時間。

### 3. 優化 DebugForm 初始化
*   **[MODIFY] [DebugForm.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/DebugForm.vb)**
    *   將 `DebugForm_Load` 中的 UI 佈局校算（例如 `txtDebug` 寬度計算）移至 `OnShown` 或是簡化佈局結構。
    *   減少啟動時不必要的 `Invalidate` 呼叫。

### 4. 其它細微優化
*   **[MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1.vb)**
    *   `InitTreeViews`：不再使用遞迴掃描全表單 (`GetAllTreeViews`)，改為針對已知的 `TreeView1`~`TreeView5` 列表進行處理。

## 待確認問題

> [!IMPORTANT]
> 1. 您是否同意將 Tab2~Tab5 的 UI 初始化延後到使用者點擊該分頁時？這會讓啟動變快，但第一次點擊分頁時會有極輕微的延遲感。
> 2. 關於 `IsMailFolder` 的過濾，目前是在 `checkIncludeAllFolders` 為 `False` 時才過濾。若勾選顯示全部，則不應檢查類型，這點邏輯是否正確？

## 驗證計劃

### 自動與手動驗證
*   查看 DebugForm 中的啟動 Log，比對優化前後的 `InitListViews` 與 `Form1_Load` 總耗時。
*   檢查切換 Tab2~Tab5 時，介面是否能正確動態生成且不影響功能。
