# 計畫：優化 Tab2 渲染與 UI 反應速度 (v2)

本計畫聚焦於減少 UI 執行緒在渲染統計結果時的無效操作。透過 `AddRange` 減輕 Windows 訊息負荷，並利用渲染節流避免重複繪圖。

## User Review Required

> [!NOTE]
> **優化重點：**
> 1. **ListView.AddRange()**: 取代逐列 `Items.Add`，將 12 個月份或數十個年份一次性置入。
> 2. **Idempotent Rendering**: 使用「數據指紋」判定內容是否變動，若無變動則跳過 `Clear()`。
> 3. **資源生命週期優化**: 將頻繁建立的 `Font` 改為全域靜態資源。
> 4. **大數據自動降級**: 當資料夾數量過多 (> 500) 時，自動關閉圖表數據標籤。

## Proposed Changes

### [MODIFY] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_MainTabs.vb)

#### 1. 全域資源宣告
- 定義 `ReadOnly _fItalic` 與 `ReadOnly _fBold` 為常駐字型。
- 定義 `_lastLv2RenderTag` 變數存放上一次渲染的數據指紋。

#### 2. RenderLvYearView / RenderLvMonthView
- **修改前**：
  ```vb
  ListView2.Items.Clear()
  For Each pair In sortedData
      ListView2.Items.Add(New ListViewItem(...))
  Next
  ```
- **修改後**：
  - 加入 `_lastLv2RenderTag` 判定，若數據未變則直接 return。
  - 使用 `List(Of ListViewItem)` 收集物件，最後呼叫 `ListView2.Items.AddRange(items.ToArray())`。
  - 套用預定義的 `_fItalic` 與 `_fBold`。

#### 3. RenderCtYearView / RenderCtMonthView
- 同步加入「指紋判定」，避免在已繪製相同內容時再次觸發沈重的 `Chart.Invalidate()`。
- 加入大數據條件式：若資料夾數量過多，則設定 `IsValueShownAsLabel = False`。

### [MODIFY] [Form1_Outlook.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_Outlook.vb)

- **L2.5 快取存取優化**：
  在 `CollectYearCounts` 與 `CollectMonthCounts` 中，若檢測到正在進行大批量讀取，則調用批次模式存取數據庫，減少 SQLite I/O 爭用。

## Verification Plan

### Automated Tests
- 使用 `Stopwatch` 監控 `RenderLvMonthView` 執行時間，預期從數十 ms 降至 < 5ms (命中節流時近乎 0ms)。
- 比對加入 `AddRange` 前後的繪圖閃爍感。

### Manual Verification
- 快速來回切換年度/月份視圖，確認畫面上數據正確且無明顯延遲。
- 在快取命中後雙擊月份分佈，確認切換瞬間完成。
