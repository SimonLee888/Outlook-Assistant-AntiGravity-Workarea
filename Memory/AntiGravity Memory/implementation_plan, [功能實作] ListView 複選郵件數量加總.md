# [功能實作] ListView 複選郵件數量加總

在 Tab1 (ListView1) 與 Tab2 (ListView2) 的 KeyPress 事件中增加 Enter 鍵的複選判斷。當選取多個項目時，彈出 MessageBox 顯示數值加總；若僅選取單一項目，則維持原有的導覽邏輯。

## 使用者審閱請求

> [!IMPORTANT]
> **欄位解析邏輯**：
> *   **Tab1**：加總 `SubItems(1)` (本層) 與 `SubItems(3)` (子樹)。
> *   **Tab2**：加總 `SubItems(1)` (郵件數)。
> *   **數值清理**：ListView 中的文字通常包含逗號 (如 `1,234`) 或結尾空格，我會使用 `.Replace(",", "").Trim()` 來確保數值能正確轉為 Long 進行加總。

## 擬議變更

### 核心邏輯層

#### [修改] [Form1.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/Form1.vb)

-   修改 `HandleListViewKeyPress` 函數：
    -   **ListView1 區塊**：
        -   檢查 `lv.SelectedItems.Count > 1`。
        -   若為複選，進入加總循環，讀取並累加數值。
        -   使用 `MessageBox.Show` 顯示兩組數值。
    -   **ListView2 區塊**：
        -   檢查 `lv.SelectedItems.Count > 1`。
        -   加總後顯示 MessageBox。
    -   加上註解：`by Gemini 3 Flash, 2026/04/13`。

## 驗證計畫

### 手動驗證
- [ ] **Tab1 複選測試**：選取多個資料夾，按 Enter，檢查 MessageBox 內容。
- [ ] **Tab1 單選測試**：選取單一資料夾，按 Enter，應正常進入下一層。
- [ ] **Tab2 複選測試**：選取多個年份/月份，按 Enter，檢查加總是否正確。
- [ ] **Tab2 單選測試**：維持原有的視圖切換行為。
