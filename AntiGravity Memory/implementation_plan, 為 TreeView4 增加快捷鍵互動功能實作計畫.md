# 為 TreeView4 增加快捷鍵互動功能實作計畫

此計畫旨在優化 Tab4 (系列郵件) 的操作體驗，透過鍵盤快捷鍵快速切換焦點、重新掃描或重置狀態。

## 擬議變更

### [Component] Form1_MainTabs.vb

#### [MODIFY] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_MainTabs.vb)

- 在 `#Region "■ 07 Tab4: 系列郵件"` 區域中新增 `TreeView4_KeyDown` 事件處理器。

- **Enter 鍵**：當在 `TreeView4` 按下 Enter 時，將焦點移至 `ListView4`。若 `ListView4` 有項目，則預設選取第一項。
- **F5 鍵**：按下 F5 時，觸發 `Button4.PerformClick()`，執行系列郵件掃描。
- **ESC 鍵**：按下 ESC 時，執行下列重置動作：
    1. 清除 `TreeView4` 的所有節點。
    2. 清除 `ListView4` 的所有項目。
    3. 呼叫 `LoadStoreToTreeView` 重新載入所有 Store 的根目錄。
    4. 呼叫 `ExpandTreeToDefaultInbox` 展開預設收件匣。
    5. 重置狀態列訊息。

## 驗證計畫

### 手動驗證
1. 切換至 Tab4。
2. 在 `TreeView4` 選取資料夾後，按下 **Enter**，確認焦點是否移至 `ListView4`。
3. 按下 **F5**，確認是否觸發掃描邏輯（`ProgressBar` 應顯示掃描進度）。
4. 按下 **ESC**，確認 `TreeView4` 是否回歸到剛啟動時只有根目錄與預設展開收件匣的狀態。
