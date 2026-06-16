# 修復 DebugForm 重繪黑影與效能優化實作計畫

解決 `DebugForm` 在同步移動或縮放時出現的黑影與延遲作業。

## 核心問題分析

1.  **繪製效能瓶頸**：`lvwDebug_DrawSubItem` 在每次繪製時都執行 `String.Join`、`Cast`、`Select` 等高耗能操作，導致 UI 執行緒阻塞。
2.  **重繪機制衝突**：`SyncDebugFormPosition` 中的 `WM_SETREDRAW` 與 `RedrawWindow` 配合 `OwnerDraw` 效能低落，產生視覺空檔（黑影）。
3.  **背景清除問題**：當 `WM_SETREDRAW` 為 0 時，視窗完全停止繪製。在視窗移動或縮小時，暴露出的「新區域」若沒能即時由控制項填滿，就會顯示為黑影。

## 擬定變更

### 1. [DebugForm.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/DebugForm.vb) 優化

-   **快取機制**：在 `ListViewItem.Tag` 中存儲一個自訂物件，記錄該行是否「命中搜尋關鍵字」，不再於繪製時動態判斷。
-   **優化 `DrawSubItem`**：
    -   移除所有 LINQ 運算。
    -   只讀取預先算好的 `IsHit` 旗標來決定是否畫高亮。
-   **非同步/批次更新**：在 `txtDebug_TextChanged` 時，統一批次更新所有項目的 `IsHit` 狀態，一次 `Refresh()`。

### 2. [Form1.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1.vb) 優化

-   **`SyncDebugFormPosition` 精細化策略**：
    -   **移動 (Move)**：直接執行 `SetWindowPos`，不使用 `WM_SETREDRAW`。
    -   **縮放 (Resize)**：針對 `lvwDebug` 執行 `WM_SETREDRAW` 以防欄位寬度調整時閃爍。
    -   移除過於激進的 `RDW_UPDATENOW`，讓 Windows 繪製訊息隊列自然排程。

## 驗證計畫

-   **移動測試**：快速移動主視窗，確認 `DebugForm` 是否緊跟且無黑邊。
-   **縮放測試**：調整主視窗寬度，確認 `DebugForm` 寬度跟隨時內容無閃爍。
-   **搜尋測試**：輸入關鍵字，確認高亮顯示依舊正確且反應靈敏。
