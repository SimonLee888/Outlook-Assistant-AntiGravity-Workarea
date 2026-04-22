# 重構 Chart2 初始化邏輯

將 `Form1.vb` 中 `InitTab2UI()` 內一段關於 `Chart2` 外觀與佈局的設定代碼移至 `InitChart2()` 函式中。這樣可以讓 `InitChart2` 負責完整的圖表初始化工作，提高代碼的可讀性與可維護性。

## User Review Required

> [!IMPORTANT]
> **執行順序影響**：
> 在 `InitTab2UI` 中，`InitChart2` 會先被呼叫（清除並重新建立 Series/ChartArea），接著才是原本在 L446-L466 的細節設定。
> 將這些設定移入 `InitChart2` 後，必須確保它們作用在**新建立**的 `ChartArea` 上，否則設定會因為 `.ChartAreas.Clear()` 而消失。

## Proposed Changes

### Form1.vb (UI 初始化)

#### [MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1.vb)

- **移除** `InitTab2UI` 內的 L446-L466 區塊，此區塊原本是在 `InitChart2()` 呼叫之後進行額外的外觀與佈局補強。
- **整合至** `InitChart2` 並優化執行順序：

```vbnet
    Private Sub InitChart2()
        ' ... [1. Clear 階段] ...
        .Series.Clear()
        .Legends.Clear()
        .ChartAreas.Clear()

        ' ... [2. Chart 全域設定] (原本散落在 outside) ...
        .BorderlineDashStyle = ChartDashStyle.Solid
        .BorderlineColor = ThemeColors.AltoGray

        ' ... [3. 建立物件階段] ...
        Dim mailChart As New ChartArea With { ... }
        
        With mailChart
            ' --- [遷移] 將 Position 與 InnerPlotPosition 整合進物件宣告之後 ---
            .Position = New ElementPosition(1, 1, 99, 99)
            With .InnerPlotPosition
                .Auto = False
                .X = 8 : .Y = 2 : .Width = 90 : .Height = 90
            End With
            
            ' --- [遷移] 格線顏色統一設定 ---
            .AxisX.MajorGrid.LineColor = ThemeColors.gridLine
            .AxisY.MajorGrid.LineColor = ThemeColors.gridLine
        End With

        ' ... [4. Add 階段] ...
        .ChartAreas.Add(mailChart)

        ' ... [5. 安全檢查] ---
        If .Legends.Count > 0 Then .Legends(0).Enabled = False
    End Sub
```

---

## Open Questions
無。

## Verification Plan

### Manual Verification
- 切換至「依日期統計 (Tab2)」分頁。
- 確認 `Chart2` 的外觀是否正確（背景色、格線顏色、邊框風格）。
- 確認長條圖的佈局（Position 和 InnerPlotPosition）是否如同預期般填滿空間且正確留白給 Y 軸標籤。
- 確認不會引發 `ArgumentOutOfRangeException`（已包含 Legends 數量檢查）。
