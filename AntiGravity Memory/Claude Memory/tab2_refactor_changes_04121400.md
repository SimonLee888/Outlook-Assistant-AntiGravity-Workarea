# Tab2 重構改動清單 — 2026/04/12
# 套用方式: 每一節標題說明「在哪裡改什麼」，請逐節套用

---

## ① 成員變數區 (約 line 37)
### 刪除這一行:
```
Private _tab2CachedYearCounts As ConcurrentDictionary(Of Integer, Integer) = Nothing    ' 記住算好的年度資料，供月份視圖快速返回
```
### 在 _tab2MonthViewYear 那行下面插入這兩行:
```vb
    Private _lv2DataYear As ConcurrentDictionary(Of Integer, Integer) = Nothing         ' Tab2 年度視圖快取（已合併多資料夾），供月份進出快速 render
    Private _lv2DataMonth As ConcurrentDictionary(Of Integer, Integer) = Nothing        ' Tab2 月份視圖快取（已合併多資料夾）；_tab2MonthViewYear 記錄對應年份（方案A）
```

---

## ② Region header 的「替換說明」和「分層架構」註解 (line 636~673)
### 整段替換:
```vb
#Region "■ 05 Tab2: 依日期統計"
    ' ==============================================================
    ' 重構目標: COM/UI/流程邏輯與業務分離清晰分層，去除全域狀態，優化快取機制
    ' 1. 分層架構: 將原本混在一起的程式碼重構成三個明確的層次
    '    - Layer 1 (UI 事件層)  : 回應使用者操作，組裝參數後交給 Layer2 執行，最後把結果交給顯示函數
    '    - Layer 2 (流程協調層) : BFS 遍歷 folderList，管理快取，驅動 Layer3 計算，合併結果，回報進度
    '    - Layer 3 (COM 資料層) : 對 Outlook 發出 COM 呼叫，回傳單一資料夾的年份郵件分佈
    ' 2. 去除全域狀態: 原本的 _intTotalMailCount 和 _intProcessedCount 全域變數已改成局部變數，避免多次點選時的計數錯亂
    ' 3. 優化快取機制: 快取的 key 改為純字串 FolderPath，避免 COM 物件當 key 導致 RCW 殘留問題；快取只存單一資料夾的結果，由 Layer2 負責合併
    ' 4. 進度回報改為 callback 機制: Layer2 執行統計時，透過 onProgress callback 回報已處理的郵件數和總郵件數，Layer1 負責更新 UI 顯示，保持分層乾淨
    ' by: Claude AI (2026/3/10)
    ' ==============================================================
    '
    ' 替換說明:
    '   以下程式碼完整取代 Tab2 相關的所有邏輯函數。
    '   請同時刪除以下舊的函數與宣告:
    '     - Private _intTotalMailCount As Integer   (全域變數宣告，已改成局部)
    '     - Private _intProcessedCount As Integer   (全域變數宣告，已改成局部)
    '     - TreeView2_AfterSelect()                 (已重寫)
    '     - SimTree2_AfterSelect()                  (已重寫，不再 commented out)
    '     - CheckSubFolder2_CheckedChanged()        (已重寫)
    '     - GetYearCountsAsync_CL()                 (已由 ComputeYearCounts 取代)
    '     - CountMailByYearAsync_CLayer2()          (已由 GetYearCountsForFolderAsync 取代)
    '     - UpdateCounterProgress()                 (已改成 callback 機制，函數可刪除)
    '
    ' 2026/04/12 重構 v2 (render 層拆分):
    '   刪除: ShowYearView, ShowMonthView, ShowResultTab2, ShowProgressTab2
    '         UpdateChart2forYearView, UpdateChart2forMonthView
    '   新增: RenderLvwYearView, RenderChart2YearView      ← 年度視圖 render（純UI，不計算）
    '         CollectMonthCounts                            ← 月份資料收集（純計算，不碰UI render）
    '         RenderLvwMonthView, RenderChart2MonthView     ← 月份視圖 render（純UI，不計算）
    '   改動: SimTree2_AfterSelect — inline ShowProgressTab2，呼叫 RenderLvwYearView/RenderChart2YearView
    '         ListView2_MouseDoubleClick — inline ShowYearView/ShowMonthView，方案A _lv2DataMonth 快取
    '
    ' 分層架構 (更新後):
    '   Layer 1 (UI 事件層)       : SimTree2_AfterSelect, CheckSubFolder2_CheckStateChanged
    '                                ListView2_MouseDoubleClick (含返回/進入月份 inline)
    '   Layer 1.5 (render 層)     : RenderLvwYearView, RenderChart2YearView
    '                                RenderLvwMonthView, RenderChart2MonthView
    '   Layer 2 (流程協調層)      : ComputeYearCounts, CollectMonthCounts
    '   Layer 3 (COM 資料層)      : GetYearCountsForFolder (Form1_Outlook.vb，不動)
    '                                GetMonthCountsForYear  (Form1_Outlook.vb，不動)
    ' ==============================================================
```

---

## ③ SimTree2_AfterSelect 末尾 Try 區塊內（約 line 754~762）
### 找到這段:
```vb
            _tab2CachedYearCounts = Await ComputeYearCounts(folderList, totalMailCount, progressYear, cToken)
            ' 呼叫 Layer2 流程協調層執行統計 (跟單選模式走一樣的路徑，只是 folderList 不同) ' 2026/4/12, ✅ 改為存入全域快取變數

            ' --- 序號校驗點 3 (核心運算完成後) ---
            If _tab2SelectSeq <> mySeq Then Return          ' Dbg("結束", "序號已不匹配，丟棄本次結果（運算完畢中斷）")
            stopwatch.Stop()                                ' ✅ 統計完成後才停錶

            ShowResultTab2(_tab2CachedYearCounts)                      ' 顯示結果到 ListView2 和 Chart2
            ShowProgressTab2(_tab2CachedYearCounts, stopwatch.Elapsed) ' 顯示執行時間與處理速度到 ProgressBar2
```
### 替換成:
```vb
            _lv2DataYear = Await ComputeYearCounts(folderList, totalMailCount, progressYear, cToken)
            ' 呼叫 Layer2 流程協調層執行統計；結果存入 _lv2DataYear session 快取，月份進出時可直接 render 不重算

            ' --- 序號校驗點 3 (核心運算完成後) ---
            If _tab2SelectSeq <> mySeq Then Return          ' Dbg("結束", "序號已不匹配，丟棄本次結果（運算完畢中斷）")
            stopwatch.Stop()                                ' ✅ 統計完成後才停錶

            ' 2026/04/12: ShowResultTab2 + ShowProgressTab2 拆分為 Render 函數 + inline progress
            RenderLvwYearView(_lv2DataYear)
            RenderChart2YearView(_lv2DataYear)
            Dim _yTotal As Integer = _lv2DataYear.Values.Sum   ' Values.Sum 是最可靠的實際計數（含/不含子資料夾皆正確）
            Dim _ySpd As Double = If(stopwatch.Elapsed.TotalSeconds > 0, _yTotal / stopwatch.Elapsed.TotalSeconds, 0)
            ProgressBar1.Text = $"共 {_yTotal:###,###,##0} 封 / {stopwatch.Elapsed.TotalSeconds:0.00} 秒"
            ProgressBar2.Text = $"(年度統計完成 - 處理速度為 {_ySpd:###,##0}/sec)"
```

---

## ④ ListView2_MouseDoubleClick 內部 (約 line 787~803)
### 找到這段:
```vb
        Try
            ' 月份視圖 → 雙擊「← 返回年度統計」: 回到年度視圖
            If _tab2IsMonthView AndAlso clickedItem.Tag?.ToString() = "BACK" Then
                Await ShowYearView(cToken) : Return
            End If

            ' 年度視圖 → 雙擊某一年: 展開為月份視圖
            ' 2026/3/16: monthCountsCache 已在 GetMonthCountsForYear 內部實作，重複展開同一年直接命中快取
            Dim selectedYear As Integer = ParseYearFromText(clickedItem.Text)
            If selectedYear = 0 Then Return
            If _tab2FolderList Is Nothing OrElse _tab2FolderList.Count = 0 Then Return

            Await ShowMonthView(selectedYear, cToken)
            Dbg("結束", $"{selectedYear} 年") ' by Gemini, 2026/04/10
```
### 替換成:
```vb
        Try
            ' ── 月份視圖 → 雙擊「← 返回年度統計」: 回到年度視圖 ──
            ' 2026/04/12: inline 原 ShowYearView (已刪除)
            ' 注意: 刻意不 reset _tab2MonthViewYear，讓 _lv2DataMonth 快取跨 back-and-forth 繼續有效
            If _tab2IsMonthView AndAlso clickedItem.Tag?.ToString() = "BACK" Then
                Dbg(" ├ 返回年度視圖")
                Dim yearToRestore As Integer = _tab2MonthViewYear   ' 先記住要還原游標的年份
                _tab2IsMonthView = False
                Await Task.Yield()   ' 讓 UI 喘口氣，確保畫面流暢切換
                If _lv2DataYear IsNot Nothing AndAlso _tab2FolderList IsNot Nothing AndAlso _tab2FolderList.Count > 0 Then
                    Cursor = Cursors.WaitCursor
                    RenderLvwYearView(_lv2DataYear)
                    RenderChart2YearView(_lv2DataYear)
                    Dim _rTotal As Integer = _lv2DataYear.Values.Sum
                    ProgressBar1.Text = $"共 {_rTotal:###,###,##0} 封"
                    ProgressBar2.Text = "(返回年度統計)"
                    Cursor = Cursors.Default
                End If
                ' 還原游標到進入月份前的那一年，讓使用者感覺回到剛才看的地方
                If yearToRestore > 0 AndAlso ListView2.Items.Count > 0 Then
                    Dim tgt = FindListViewItemByYear(yearToRestore)
                    If tgt IsNot Nothing Then tgt.Selected = True : tgt.Focused = True : tgt.EnsureVisible() : ListView2.Focus()
                End If
                Dbg(" ├ 返回完成") : Return
            End If

            ' ── 年度視圖 → 雙擊某一年: 展開為月份視圖 ──
            ' 2026/04/12: inline 原 ShowMonthView (已刪除)；方案A _lv2DataMonth 快取 (_tab2MonthViewYear 為年份 tag)
            ' 快取命中（同一年份）: 純 render，不碰 COM/快取層；快取未命中: CollectMonthCounts → merge → render
            Dim selectedYear As Integer = ParseYearFromText(clickedItem.Text)
            If selectedYear = 0 Then Return
            If _tab2FolderList Is Nothing OrElse _tab2FolderList.Count = 0 Then Return

            Dbg(" ├ 進入月份視圖", selectedYear.ToString())
            Dim swM As New Stopwatch() : swM.Start()
            ProgressBar1.Text = "" : ProgressBar2.Text = "" : Cursor = Cursors.WaitCursor

            If _lv2DataMonth IsNot Nothing AndAlso _tab2MonthViewYear = selectedYear Then
                ' ★ _lv2DataMonth 快取命中：直接 render，不碰任何計算層
                Dbg(" ├ _lv2DataMonth 快取命中", selectedYear.ToString())
                _tab2IsMonthView = True
                RenderLvwMonthView(selectedYear, _lv2DataMonth)
                RenderChart2MonthView(_lv2DataMonth, selectedYear)
                Dim _mHit As Integer = _lv2DataMonth.Values.Sum
                ProgressBar1.Text = $"共 {_mHit:###,###,##0} 封"
                ProgressBar2.Text = $"({selectedYear} 年月份分佈 - 按 ESC 或雙擊標題橫列可返回視圖) "
                Cursor = Cursors.Default
            Else
                ' ★ 快取未命中：CollectMonthCounts → _monthCountsCache 一定命中 → merge → render
                Dbg(" ├ _lv2DataMonth 快取未命中，開始計算", selectedYear.ToString())
                Dim mc As ConcurrentDictionary(Of Integer, Integer) = Await CollectMonthCounts(selectedYear, cToken)
                _lv2DataMonth = mc : _tab2MonthViewYear = selectedYear : _tab2IsMonthView = True
                swM.Stop()
                RenderLvwMonthView(selectedYear, mc)
                RenderChart2MonthView(mc, selectedYear)
                Dim _mMiss As Integer = mc.Values.Sum
                Dim _mSpd As Double = If(swM.Elapsed.TotalSeconds > 0, _mMiss / swM.Elapsed.TotalSeconds, 0)
                ProgressBar1.Text = $"共 {_mMiss:###,###,##0} 封 / {swM.Elapsed.TotalSeconds:0.00} 秒"
                ProgressBar2.Text = $"({selectedYear} 年月份分佈讀取完成 - 按 ESC 或雙擊標題橫列可返回視圖) "
                Cursor = Cursors.Default
            End If

            ' 確保 SimTree2 的選取節點保持可見
            If SimTree2.Visible Then
                Dim nodes As List(Of TreeNode) = SimTree2.SelectedNodes
                If nodes IsNot Nothing AndAlso nodes.Count > 0 Then nodes(0).EnsureVisible()
            End If
            Dbg("結束", $"{selectedYear} 年")
```

---

## ⑤ ListView2_SelectedIndexChanged 內一行注解 (約 line 843)
### 找到:
```
            If selectedMonth > 0 Then targetIndex = selectedMonth - 1 ' UpdateChart2forMonthView 依 1~12 月順序加入 DataPoints，月份N = index N-1
```
### 替換成:
```vb
            If selectedMonth > 0 Then targetIndex = selectedMonth - 1 ' RenderChart2MonthView 依 1~12 月順序加入 DataPoints，月份N = index N-1
```

---

## ⑥ 刪除這六個函數（整個函數含 Dbg/注解一起刪）
- `Private Async Function ShowYearView(...)`  （約 line 995~1026）
- `Private Async Function ShowMonthView(...)`  （約 line 1027~1128）
- `Private Sub ShowResultTab2(...)`            （約 line 1129~1149）
- `Private Sub UpdateChart2forYearView(...)`   （約 line 1151~1205）
- `Private Sub UpdateChart2forMonthView(...)`  （約 line 1207~1244）
- `Private Sub ShowProgressTab2(...)`          （約 line 1246~1255）

---

## ⑦ 在 #End Region "└ 輔助函數" 前，插入五個新函數

### 插入位置: ShowProgressTab2 刪除後，#End Region "  ├ Layer1 UI事件層" 的 #End Region 之前

```vb
#Region "  ├ Layer1.5 Render層 (純UI，不計算)"
    Private Sub RenderLvwYearView(yearCounts As ConcurrentDictionary(Of Integer, Integer))
        ' ---------------------------------------------------------------
        ' 年度視圖 ListView2 渲染 (2026/04/12 由 ShowResultTab2 拆出)
        ' 職責: 純 UI render，不做計算，不查快取，不碰 COM
        ' 對稱: RenderChart2YearView 負責同一視圖的 Chart2 部分
        ' ---------------------------------------------------------------
        Dbg("開始", yearCounts?.Count)
        ListView2.Items.Clear()
        If yearCounts Is Nothing OrElse yearCounts.IsEmpty Then
            ClearChart2Series()     ' ★ 空資料夾時也要清除 Chart2，否則前一個資料夾的圖表會殘留
            ListView2.Items.Add(New ListViewItem("找不到郵件"))
        Else
            ListView2.BeginUpdate()     ' ✅ 批次更新，避免每次 Add 都觸發重繪
            Dim sortedYearCounts = yearCounts.OrderBy(Function(pair) pair.Key).ToList()
            For Each pair In sortedYearCounts
                ListView2.Items.Add(New ListViewItem({pair.Key, pair.Value.ToString("###,###,##0") & " "}))  ' 字串結尾一律補一格空白 (by Gemini, 2026/03/31)
            Next
            ListView2.EndUpdate()
        End If
        Dbg("結束")
    End Sub

    Private Sub RenderChart2YearView(yearCounts As ConcurrentDictionary(Of Integer, Integer))
        ' ---------------------------------------------------------------
        ' 年度視圖 Chart2 渲染 (2026/04/12 由 UpdateChart2forYearView 改名重構)
        ' 職責: 純 UI render；接受 ConcurrentDictionary，內部自行排序
        ' 原 UpdateChart2forYearView 接受已排好的 List，由 caller 負責排序；
        ' 改為自己排序，讓 caller 不需要了解排序的存在，介面更乾淨
        ' 對稱: RenderLvwYearView 負責同一視圖的 ListView2 部分
        ' ---------------------------------------------------------------
        Dbg("開始")
        ClearChart2Series()     ' 清除之前的統計結果，包括 Series Points、平均線 Series、平均值標籤 Annotation
        If yearCounts Is Nothing OrElse yearCounts.IsEmpty Then Return
        Dim sortedYearCounts = yearCounts.OrderBy(Function(p) p.Key).ToList()

        ' 添加數據到 Series
        Dim series As Series = Chart2.Series(0)
        For Each pair In sortedYearCounts
            series.Points.AddXY(pair.Key, pair.Value)
        Next

        ' 依內容大小設置 Chart2 的 X 軸上下限
        With Chart2.ChartAreas(0).AxisX
            .Minimum = sortedYearCounts.Min(Function(p) p.Key) - 0.5
            .Maximum = sortedYearCounts.Max(Function(p) p.Key) + 0.5
            .Interval = 1
            .IntervalOffset = 0                 ' ✅ 還原年度視圖的長條置中偏移
            .LabelStyle.Format = "####"         ' ✅ 還原年份格式
            .LabelStyle.Interval = 1
            .LabelStyle.IntervalOffset = 0.5    ' ✅ 校正還原上面 max/min 的 0.5 偏移
            .MajorTickMark.IntervalOffset = 0   ' ✅ 還原刻度偏移
        End With

        ' 添加一條代表平均值的線（獨立 Series 才能控制線型，StripLine 不支援虛線）
        ' 2026/3/6 by Claude Code；2026/04/12 移入 RenderChart2YearView
        Dim average As Double = sortedYearCounts.Average(Function(pair) pair.Value)
        Dim xMin As Double = sortedYearCounts.Min(Function(pair) pair.Key)
        Dim xMax As Double = sortedYearCounts.Max(Function(pair) pair.Key)

        Dim avgSeries As New Series("平均線") With {.ChartType = SeriesChartType.Line,
                                                    .Color = ThemeColors.avgLineColor,
                                                    .BorderWidth = 2,
                                                    .BorderDashStyle = ChartDashStyle.Dash,  ' ✅ 虛線
                                                    .ChartArea = Chart2.ChartAreas(0).Name,
                                                    .IsVisibleInLegend = False}
        avgSeries.Points.AddXY(xMin - 1, average)  ' 0: 從 X 軸最小值往左延伸
        avgSeries.Points.AddXY(xMax, average)       ' 1: 圖表最右邊長條的確切 X 座標（錨定用）
        avgSeries.Points.AddXY(xMax + 1, average)  ' 2: 到 X 軸最大值往右延伸

        ' 用 TextAnnotation 顯示平均值標籤 (by Gemini, 2026/04/04 改用 DeepAmber 提升辨識度)
        Dim avgLabel As New TextAnnotation With {.Name = "平均值標籤",
                                                 .Text = "AVG: " & average.ToString("#,###,##0"),
                                                 .ForeColor = ThemeColors.avgLineColor,
                                                 .Font = New Font("Tahoma", 10.0F, System.Drawing.FontStyle.Bold),
                                                 .AnchorDataPoint = avgSeries.Points(1),          ' 錨定在最右側長條的中間點 X 座標
                                                 .AnchorAlignment = ContentAlignment.BottomCenter, ' ★ 強制對齊點的正上方（避免 MS Chart 自動亂飄移）
                                                 .AnchorOffsetX = 0,    ' 保持置中
                                                 .AnchorOffsetY = -1,   ' 產生 1% 的空隙，確保不在線上
                                                 .BackColor = Color.Transparent,
                                                 .LineColor = Color.Transparent}
        Chart2.Series.Add(avgSeries)
        Chart2.Annotations.Add(avgLabel)
        Chart2.Invalidate()     ' 強制重新繪製圖表
        Dbg("結束")
    End Sub

    Private Sub RenderLvwMonthView(selectedYear As Integer, monthCounts As ConcurrentDictionary(Of Integer, Integer))
        ' ---------------------------------------------------------------
        ' 月份視圖 ListView2 渲染 (2026/04/12 由 ShowMonthView render 部分拆出)
        ' 職責: 純 UI render，不做計算，不查快取，不碰 COM
        ' 對稱: RenderChart2MonthView 負責同一視圖的 Chart2 部分
        ' ---------------------------------------------------------------
        Dbg(" ├ 開始", selectedYear.ToString())
        ListView2.BeginUpdate()
        ListView2.Items.Clear()

        ' 第一行: 返回按鈕
        Dim backItem As New ListViewItem("← 返回年度統計")
        backItem.SubItems.Add("") : backItem.Tag = "BACK"
        backItem.ForeColor = Color.Gray
        backItem.Font = New Font(_fontDefault, _fontItalic)
        ListView2.Items.Add(backItem)

        ' 第二行: 年份標題
        Dim titleItem As New ListViewItem($"── {selectedYear} 年月份分佈 ──")
        titleItem.SubItems.Add($"共 {monthCounts.Values.Sum:###,###,##0}  封")  ' 字串結尾補上空白防止選取時切邊，與下方對齊
        titleItem.ForeColor = Color.DimGray
        titleItem.Font = New Font(_fontDefault, _fontBold)
        ListView2.Items.Add(titleItem)

        ' 逐月顯示 (只顯示有郵件的月份)
        For month As Integer = 1 To 12
            Dim count As Integer = 0
            monthCounts.TryGetValue(month, count)
            If count > 0 Then
                Dim monthItem As New ListViewItem($"{selectedYear} /  {month:D2}月")
                monthItem.SubItems.Add(count.ToString("###,###,##0") & " ")  ' 字串結尾一律補一格空白
                ListView2.Items.Add(monthItem)
            End If
        Next
        ListView2.EndUpdate()
        Dbg(" ├ 結束", selectedYear.ToString())
    End Sub

    Private Sub RenderChart2MonthView(monthCounts As ConcurrentDictionary(Of Integer, Integer), year As Integer)
        ' ---------------------------------------------------------------
        ' 月份視圖 Chart2 渲染 (2026/04/12 由 UpdateChart2forMonthView 改名)
        ' 月份長條圖：只畫 1~12 月，X 軸標籤顯示「M月」，不畫平均線
        ' 完整替換 Chart2 的內容，與 RenderChart2YearView 平行存在
        ' 對稱: RenderLvwMonthView 負責同一視圖的 ListView2 部分
        ' ---------------------------------------------------------------
        Dbg("開始", year)
        ClearChart2Series()     ' 清除之前的所有圖表內容（同 RenderChart2YearView 的清除邏輯）

        ' 把 1~12 月的資料全部加入（沒有郵件的月份補 0，讓 X 軸保持完整 12 格）
        Dim series As Series = Chart2.Series(0)
        For month As Integer = 1 To 12
            Dim count As Integer = 0
            monthCounts.TryGetValue(month, count)
            Dim pt As New DataPoint()
            pt.SetValueXY(month, count)
            pt.AxisLabel = $"{month}月"     ' ✅ 用月份名稱當 X 軸標籤，比純數字 1~12 更易讀
            series.Points.Add(pt)
            pt.IsVisibleInLegend = True
        Next

        ' X 軸固定顯示 1~12，不根據資料範圍自動縮放
        ' X 軸重置所有從 InitChart2 繼承的年度設定，改成月份專用設定
        With Chart2.ChartAreas(0).AxisX
            .Minimum = 0.5
            .Maximum = 12.5
            .Interval = 1
            .IntervalOffset = 0                 ' ✅ 清除 InitChart2 的 0.5 偏移量
            .LabelStyle.Format = ""             ' ✅ 清除 "####" 年份格式，讓 AxisLabel 屬性生效
            .LabelStyle.Interval = 1
            .LabelStyle.IntervalOffset = 0.5
            .MajorTickMark.IntervalOffset = 0
        End With
        Chart2.Invalidate()
        Dbg("結束", year)
    End Sub
#End Region

#Region "  ├ Layer2 流程協調層"
    Private Async Function CollectMonthCounts(selectedYear As Integer, cToken As CancellationToken) As Task(Of ConcurrentDictionary(Of Integer, Integer))
        ' ---------------------------------------------------------------
        ' 月份資料收集 Layer2 (2026/04/12 由 ShowMonthView 計算部分拆出)
        ' 職責: 遍歷 _tab2FolderList，對每個資料夾呼叫 GetMonthCountsForYear，合併結果，回報進度
        '       不碰 UI render（render 由 caller 負責），進度透過直接寫 ProgressBar 文字（節流）
        '       OperationCanceledException 由 caller (ListView2_MouseDoubleClick) 的 Catch 攔截
        '       cToken 與 ComputeYearCounts 同理，都需要傳入以支援 ESC 中斷
        ' ---------------------------------------------------------------
        Dbg(" ├ 開始", selectedYear.ToString())
        Dim monthCounts As New ConcurrentDictionary(Of Integer, Integer)
        Dim totalFolders As Integer = _tab2FolderList.Count
        Dim processedFolders As Integer = 0
        Dim totalMailCount As Long = 0

        ' 先算信件總數作為進度顯示的分母（節流，避免頻繁更新 UI）
        Dim swThrottleA As New Stopwatch : swThrottleA.Start()
        For Each f In _tab2FolderList
            Dim c As Integer = GetCachedMailCount(f)
            If c > 0 Then totalMailCount += c
            If swThrottleA.ElapsedMilliseconds >= 100 Then
                swThrottleA.Restart() : Await Task.Delay(1, cToken)
            End If
        Next

        ' 逐資料夾取月份分布並合併
        Dim swThrottle As New Stopwatch() : swThrottle.Start()
        For Each folder As Outlook.Folder In _tab2FolderList
            processedFolders += 1
            Dim folderMonthCounts As ConcurrentDictionary(Of Integer, Integer) = Await GetMonthCountsForYear(folder, selectedYear)
            monthCounts = MergeDictionaries(monthCounts, folderMonthCounts)
            If swThrottle.ElapsedMilliseconds >= 100 OrElse processedFolders = totalFolders Then
                ProgressBar1.Text = "正在讀取..."
                ProgressBar2.Text = $"正在統計 {selectedYear} 年月份分佈: ({processedFolders}/{totalFolders})個資料夾 (相依包含共計 {totalMailCount:N0} 封信)。"
                swThrottle.Restart() : Await Task.Delay(1, cToken)
            End If
        Next
        Dbg(" ├ 結束", $"{selectedYear} 年 | 月份數: {monthCounts.Count}")
        Return monthCounts
    End Function
#End Region
```

---

## ✅ 自我檢查清單
- [x] `_tab2CachedYearCounts` 全部改成 `_lv2DataYear`（確認只有 ③ 的改動，其他地方已用搜尋確認無殘留）
- [x] `ShowResultTab2` / `UpdateChart2forYearView` 邏輯完整保留進 RenderLvwYearView / RenderChart2YearView（含所有舊注解）
- [x] `ShowMonthView` 的計算部分完整保留進 CollectMonthCounts；render 部分完整保留進 RenderLvwMonthView + RenderChart2MonthView
- [x] `ShowYearView` 的游標還原邏輯完整保留進 ListView2_MouseDoubleClick 返回分支
- [x] `_tab2MonthViewYear` 在返回時不 reset（方案A快取跨 back-and-forth 繼續有效）
- [x] `_tab2IsMonthView` 正確在兩個分支各自設定
- [x] `ListView2_SelectedIndexChanged` 的 `RenderChart2MonthView` 注解已更新（④⑤ 步驟）
- [x] `RenderChart2MonthView` 與舊 `UpdateChart2forMonthView` 參數順序相同 (monthCounts, year)
- [x] `RenderChart2YearView` 改為自己排序，不再依賴 caller 傳 sorted list
- [x] `CollectMonthCounts` 放在 Layer2 Region，與 `ComputeYearCounts` 同層
- [x] `Form1_Outlook.vb` 完全沒有異動
