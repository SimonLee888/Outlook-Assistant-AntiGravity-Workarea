# DebugForm UI 與邏輯修正計畫

## 核心問題解決

### 1. 修正 Elapsed 自動顯示失效 (Bug Fix)
- **問題**：新加入的 Log 在 `Timer_Tick` 階段的 `Index` 為 `-1`，導致 `FindSimilarPair` 往回搜尋時索引越界或直接跳過。
- **方案**：在 `FindSimilarPair` 中處理 `Index = -1` 的情況，自動將起點設為 `lvwDebug.Items.Count - 1`。

### 2. 實現原生風格的 Header 粗體 (Premium UI)
- **要求**：不使用 `OwnerDraw` 以保留原生 Mouse Over 與顏色變化效果。
- **方案**：使用 Win32 API 直接對 Header 控制項發送 `WM_SETFONT`。
    - 使用 `LVM_GETHEADER` 取得 Header Handle。
    - 建立與 ListView 相同字型但帶有 `Bold` 屬性的 `hFont`。
    - 在 `DebugForm_Shown` 時套用。

## 擬定修改方案

### 1. [Modify] DebugForm.vb — Win32 API 擴充
- 新增 `WM_SETFONT` 常數。
- 在 `DebugForm_Shown` 執行 Header 字體替換邏輯。

### 2. [Modify] DebugForm.vb — `FindSimilarPair` 索引修正
- 增加邏輯判斷：`Dim startIdx As Integer = If(selectedItem.Index >= 0, selectedItem.Index - 1, lvwDebug.Items.Count - 1)`

## 驗證計畫

### 自動化測試
- 在 `Timer_Tick` 執行時偵測 `Elapsed` 欄位是否正確填入數值。

### 手動驗證
- 檢查 Debug 視窗開啟後，欄位標題是否為粗體。
- 將滑鼠移至欄位標題，確認是否依然有原生的藍色高亮或變色效果。
