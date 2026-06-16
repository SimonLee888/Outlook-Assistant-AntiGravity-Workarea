# Chart2 初始化邏輯重構完成

我已經將 `Chart2` 的外觀與佈局設定從 `InitTab2UI` 成功遷移並整合至 `InitChart2` 函式中。

## 變更調整細節

### Form1.vb

#### [InitTab2UI](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1.vb#L440-L448)
- 移除了末尾關於 `Chart2.Borderline`、`ChartAreas(0).Position`、`InnerPlotPosition` 等重複設定代碼。
- 現在此函式專注於 Tab2 的控制項容器佈局（Panel, SplitContainer, Dock 順序）。

#### [InitChart2](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1.vb#L614-L689)
- **外觀整合**：補回 `BorderlineDashStyle` 與 `BorderlineColor` 設定。
- **樣式修正**：將 `ChartArea.BackColor` 修正為 `ThemeColors.bgColor`（原本誤設為 `Gray95`）。
- **佈局整合**：將 `Position` 與 `InnerPlotPosition` (X=8, Y=2, W=90, H=90) 整合進 `With mailChart` 區塊，確保「建立後即設定」。
- **安全性優化**：保留並優化了 `Legends.Count` 檢查，確保在清除所有圖例後不會發生索引越界。

## 驗證結果
- 已確認所有在 `InitTab2UI` 被刪除的屬性，都已對應地出現在 `InitChart2` 的正確位置（`mailChart` 建立之後，`Add` 之前）。
- 執行順序符合预期：先 `Clear()` -> 設定新物件屬性 -> `Add()`。
