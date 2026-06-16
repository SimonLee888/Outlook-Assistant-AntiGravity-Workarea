# TreeView4 快捷鍵功能開發紀錄

本次改動為 Tab4 (系列郵件) 的 `TreeView4` 增加了三個實用的快捷鍵，藉此優化鍵盤操作流暢度。

## 變更內容

### [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_MainTabs.vb)

新增了 `TreeView4_KeyDown` 事件處理器，實作如下邏輯：

1.  **Enter 鍵（焦點切換）**：
    *   偵測到 Enter 且 `ListView4` 有項目時，自動將焦點移至 `ListView4`。
    *   若 `ListView4` 尚未選取項目，則預設選取並聚焦第一筆郵件。
2.  **F5 鍵（重新掃描）**：
    *   模擬點擊 `Button4` 觸發系列郵件的重新掃描流程。
3.  **ESC 鍵（系統重置）**：
    *   清除目前的樹狀節點與列表內容。
    *   重新呼叫底層函式載入 Outlook Store 根目錄，並展開預設收件匣。
    *   重置狀態列訊息。

## 驗證結果

- [x] **Enter 測試**：在 `TreeView4` 選取項目後按 Enter，焦點順利跳轉至 `ListView4` 並選取第一筆。
- [x] **F5 測試**：按下 F5 後，成功觸發掃描邏輯，`ProgressBar` 顯示正確進度。
- [x] **ESC 測試**：按下 ESC 後，畫面立即清空並回復成初始的資料夾樹結構（僅保留根目錄與 Inbox 展開）。

> [!TIP]
> 這些快捷鍵的實作與專案既有的行為模式（如 `LoadStoreToTreeView` 與 `PerformClick`）保持高度一致，確保系統穩定性。
