# 實作 F1 快速鍵切換側邊欄收合功能

使用者希望在應用程式的任何地方按下 **F1** 鍵，就能切換目前分頁側邊欄（Splitter）的收合與恢復狀態。這與目前雙擊 Splitter 分隔線的功能一致。

## 使用者評論與回饋要求

> [!IMPORTANT]
> 預設情況下，F1 鍵在 Windows 應用程式中通常用於開啟「說明 (Help)」。本實作將攔截 F1 並將其用於切換側邊欄，這會覆蓋系統預設行為。

## 擬議變更

### 表單與事件處理 (Form1.vb)

#### [MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1.vb)

1.  **實作核心切換邏輯**：
    *   新增一個私有方法 `PerformSplitterToggle(sc As SplitContainer)`，將原先在 `HandleSplitterMouseDown` 中的縮合/恢復邏輯抽取出來。
    *   修改 `HandleSplitterMouseDown` 改為呼叫 `PerformSplitterToggle`。

2.  **攔截 F1 快速鍵**：
    *   在 `Form1_KeyDown` 中偵測 `Keys.F1`。
    *   呼叫 `GetActiveSplitter()` 獲取當前分頁的 `SplitContainer` 並執行 `PerformSplitterToggle`。
    *   設定 `e.Handled = True` 與 `e.SuppressKeyPress = True` 以攔截系統行為。

3.  **輔助方法**：
    *   新增 `GetActiveSplitter()`：根據 `TabControl1.SelectedIndex` 傳回對應的 `SplitContainer1` ~ `SplitContainer5`。

---

## 驗證計畫

### 手動驗證
1.  **啟動程式**：在 Tab1 (資料夾統計) 按下 F1，確認左側目錄樹收合至 10px，再次按下 F1 確認恢復原寬度。
2.  **跨分頁測試**：切換到 Tab2、Tab3 等分頁，重複按下 F1，確認各自的 Splitter 都能正確獨立運作。
3.  **焦點測試**：將焦點放在 `TreeView`、`ListView` 或 `TextBox` 中按下 F1，確認側邊欄仍能觸發切換（歸功於 `KeyPreview = True`）。
4.  **連動測試**：手動雙擊分隔線收合後，按下 F1 確認其能正確「恢復」寬度（狀態應同步）。

### 自動化測試
*   使用 `_dbg` 記錄確認邏輯觸發順序。
