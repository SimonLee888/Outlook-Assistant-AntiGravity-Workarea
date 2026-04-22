# 計畫：優化 Tab2 渲染與 UI 反應速度 (v3 - 詳細示範版)

本計畫旨在透過具體的技術手段，解決 Tab2 在快取命中時仍有的毫秒級延遲。核心策略是減少 UI Thread 的工作量，並優化數據填充與資源管理。

## User Review Required

> [!IMPORTANT]
> **效能優化核心：**
> 本次計畫將導入 `AddRange` 取代 `Add`，並透過「數據指紋」判定是否需要更新 UI。這能讓大數據量下的視圖切換（如 700+ 資料夾）達到近乎「秒開」的效果。

## Proposed Changes

### [MODIFY] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_MainTabs.vb)

#### 1. 預定義資源與狀態標記
*   **修改前**：
    在 `RenderLvMonthView` 內部每次渲染都 `New Font`。
*   **修改後**：
    在宣告區定義靜態資源，並增加指紋標記。
    ```vb
    Private ReadOnly _fItalic As New Font(_fontDefault, FontStyle.Italic)
    Private ReadOnly _fBold As New Font(_fontDefault, FontStyle.Bold)
    Private _lastLv2RenderTag As String = ""
    ```

#### 2. ListView2 `AddRange()` 應用
*   **修改前 (`Items.Add`)**：
    ```vb
    ListView2.BeginUpdate()
    ListView2.Items.Clear()
    For Each pair In sortedData
        ListView2.Items.Add(New ListViewItem({pair.Key, pair.Value.ToString("N0")}))
    Next
    ListView2.EndUpdate()
    ```
*   **修改後 (`AddRange`)**：
    使用緩衝 List，一次性填充，極大減少對底層 Windows 控制項的訊息往返。
    ```vb
    Dim items As New List(Of ListViewItem)
    For Each pair In sortedData
        items.Add(New ListViewItem({pair.Key, pair.Value.ToString("N0")}))
    Next
    ListView2.BeginUpdate()
    ListView2.Items.Clear()
    ListView2.Items.AddRange(items.ToArray()) 
    ListView2.EndUpdate()
    ```

#### 3. 渲染節流 (Idempotent Rendering)
*   **示範邏輯**：
    判定數據指紋是否變動。
    ```vb
    Dim currentTag = $"YEAR_{yearCounts.Count}_{yearCounts.Values.Sum()}"
    If _lastLv2RenderTag = currentTag Then Return ' 直接跳過重繪，省下數百毫秒
    _lastLv2RenderTag = currentTag
    ```

#### 4. Chart 數據降級策略 (針對大數據量)
*   **修改前**：數據標籤一律開啟，造成圖表重疊卡頓。
*   **修改後**：
    ```vb
    ' 當渲染點數超過 500 時
    If dataPoints.Count > 500 Then
        series.IsValueShownAsLabel = False ' 關閉標籤避免負擔
        series("PointWidth") = "0.6"
    Else
        series.IsValueShownAsLabel = True
    End If
    ```

### [MODIFY] [Form1_Outlook.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_Outlook.vb)

*   **L2.5 快取層**：優化 `GetYearCountsForFolder` 在大迴圈中的 SQLite 存取穩定性。

## Verification Plan

### Automated Tests
- 記錄 `RenderLvMonthView` 在命中指紋標籤時的毫秒數，預計 < 1ms。
- 比對 `Items.Add` 與 `AddRange` 在處理 1000 筆數據時的 UI 反應速度。

### Manual Verification
- 手動點選大量資料夾進度年度視圖。
- 反覆點進點出年份，觀察控制項是否完全不再閃爍。
