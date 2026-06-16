# DebugForm 時間計算邏輯優化計畫

## 核心問題分析

目前的 `Dbg()` 系統在 Time Span 欄位有兩種不一致的計算邏輯：
1.  **預設（物理間隔）**：顯示「這一行 Log」與「上一行 Log」之間的物理時間差。這反映的是程式執行的「步進」速度。
2.  **雙擊後（邏輯耗時）**：顯示從「開始」到「結束」的總執行時間。

使用者發現這兩者經常不同，是因為大多數函數中間還會插進其他的 Log，導致預設值只計算了「最後一段」的時間，而非「整個函數」的時間。

## 預期目標

1.  **分流顯示**：將原有的 `Time Span` 拆分為兩個獨立欄位：
    *   **`Step (ms)`**：顯示上一行與這一行的物理間隔（維持原本的加總功能基準）。
    *   **`Elapsed (ms)`**：僅在「結束」訊息時，自動抓取對應的「開始」並計算總執行時間。
2.  **自動化**：當 `Dbg("結束")` 訊息出現時，系統應自動尋找配對的 `開始` 並計算總耗時，不再需要手動雙擊。
3.  **效能與準確率**：在 `Timer_Tick` 批次寫入時，即使 `開始` 與 `結束` 在同一個 100ms 內產生，也要能正確配對。
4.  **視覺對齊**：透過独立欄位實現靠右對齊，確保數值整齊美觀。

## 擬定修改方案

### 1. [Modify] DebugForm.vb — 介面初始化
- `DebugForm_Shown`: 
    - 修改原 `Time Span` 標題為 `Step (ms)`。
    - 新增欄位 `Elapsed (ms)` (寬度 85, 右對齊)。
- `RecalcColumnWidths`: 更新計算公式，保留兩個時間欄位的寬度。

### 2. [Modify] DebugForm.vb — `FindSimilarPair` 核心算法擴充
- 修改簽章，允許傳入一份「當前批次清單」(`List(Of ListViewItem)`)。
- 在搜尋 `lvwDebug.Items` 之前，先向後搜尋這份清單（因為新加入的 Log 往往跟最近的 `開始` 關聯度最高）。

### 3. [Modify] DebugForm.vb — `Timer_Tick` 邏輯升級
- 在批次處理取出待加入的 Log 後，先對其中的「結束」行進行自動配對。
- 調用擴充後的 `FindSimilarPair`，同時搜尋「已顯示」與「待顯示」的 Log。
- 若配對成功，將總耗時填入該項目的 `SubItems(3)` (Elapsed 欄位)。

### 4. [Modify] DebugForm.vb — 操作事件同步
- `lvwDebug_MouseDoubleClick`: 同步更新雙擊後的數值寫入位置 (從 SubItems(2) 改為 SubItems(3))。
- `CalculateSelectedTimeSpan`: 確認其依然讀取 `Tag.timeStamp` 或 `SubItems(2)`，確保選取範圍加總功能不受影響。

## 驗證計畫

### 自動化測試
- 在程式碼中連續呼叫 `Dbg("開始", "Test")`, `Dbg("中間", "Log")`, `Dbg("結束", "Test")`。
- 檢查 DebugForm 是否在「結束」行的 `Elapsed` 欄位自動顯示 520ms 左右的數值。
- 檢查 `Step` 欄位是否依然顯示每一行之間的物理間隔。

### 手動驗證
- 選取多行 Log（包含結束行），按 Enter 檢查加總視窗的數值是否依然為該物理區間的總和（應該與最後一行的 Elapsed 相同）。
