# 將 ListView 全選與複製功能抽離為獨立函數

這個計畫的目的是將散落在 `Lv1_KeyDown`、`Lv2_KeyDown` 與 `HandleLv3Lv4Lv5_KeyDown` 等各處的 `Ctrl+A` (全選) 與 `Ctrl+C` (複製到剪貼簿) 的邏輯抽離出來，集中到 `Form1.vb` 中的輔助函數區塊。

## User Review Required

> [!IMPORTANT]
> 請確認以下將要進行的修改：
> 1. 將在 `Form1.vb` 的 `#Region "  ├ ListView 格式工具"` 區塊內新增 `ListViewSelectAll(lv As ListView)` 與 `ListViewCopyToClipboard(lv As ListView)` 兩個輔助函數。
> 2. `Form1_MainTab12.vb` 中的 `Lv1_KeyDown` 與 `Lv2_KeyDown` 將修改為呼叫上述函數。
> 3. `Form1_MainTab345.vb` 中的 `HandleLv3Lv4Lv5_KeyDown` 將修改為呼叫上述函數，取代原有的虛擬模式 (VirtualMode) 選擇邏輯。
>
> 這樣可以簡化原本冗長且重複的 `KeyDown` 程式碼結構。

## 整合後函數結構設計 (已確認)

根據使用者的反饋，兩個整合後的函數設計如下，會加入對 `_isCtrl_A` 的內部判斷，以及對 `Ctrl+C` 匯出標題列 (Header) 的支援：

### `ListViewSelectAll`

負責處理所有 ListView 的 `Ctrl+A` 全選邏輯，自動相容虛擬模式與一般模式，並內建對特定 ListView 的效能優化旗標 (`_isCtrl_A`) 處理。

```vb
    Public Sub ListViewSelectAll(lv As ListView)
        ''' <summary>
        ''' by Gemini 3.0 flash, 2026/05/18: 抽離共用的 ListView 全選邏輯 (支援一般模式與虛擬模式)
        ''' </summary>
        If lv Is Nothing OrElse lv.Items.Count = 0 Then Return

        _dbg("開始", $"全選 {lv.Name}")

        ' 判斷是否需要開啟效能防呆旗標 (原本在 HandleLv3Lv4Lv5_KeyDown 針對大量計算的保護)
        Dim isPerformanceCritical As Boolean = (lv.Name = "ListView4" OrElse lv.Name = "ListView5")
        If isPerformanceCritical Then _isCtrl_A = True

        lv.BeginUpdate()
        Try
            If lv.VirtualMode Then
                ' 虛擬模式：直接將所有索引加進 SelectedIndices
                For i As Integer = 0 To lv.VirtualListSize - 1
                    lv.SelectedIndices.Add(i)
                Next
            Else
                ' 一般模式：遍歷實體項目並設為 Selected
                For Each item As ListViewItem In lv.Items
                    item.Selected = True
                Next
            End If
        Finally
            lv.EndUpdate()
            
            ' 如果啟用了防呆，要在 UI 更新後釋放旗標
            If isPerformanceCritical Then
                Application.DoEvents() ' 讓 UI 處理完選取變更帶來的重繪與事件
                _isCtrl_A = False
            End If
        End Try
        
        _dbg("結束", $"共選取 {lv.SelectedIndices.Count} 個項目")
    End Sub
```

### `ListViewCopyToClipboard`

負責將 ListView 選中的資料複製到剪貼簿。修改後會統一支援將「欄位標題 (Header)」作為第一列寫入剪貼簿，且會去除特殊的視覺裝飾字元（例如 `▸` 或 `-`），讓使用者直接貼上 Excel 時就有正確的欄位名稱。

```vb
    Public Sub ListViewCopyToClipboard(lv As ListView)
        ''' <summary>
        ''' by Gemini 3.0 flash, 2026/05/18: 抽離共用的 ListView 複製到剪貼簿邏輯
        ''' </summary>
        If lv Is Nothing OrElse lv.SelectedItems.Count = 0 Then Return

        _dbg("開始", $"複製 {lv.Name} 資料")
        
        Dim sb As New System.Text.StringBuilder()
        
        ' 1. 加入標題列 (Header)
        Dim headers As New List(Of String)(lv.Columns.Count)
        For Each col As ColumnHeader In lv.Columns
            headers.Add(col.Text.Trim())
        Next
        sb.AppendLine(String.Join(vbTab, headers))
        
        ' 2. 遍歷所有被選取的項目，把子欄位以 vbTab (Tab字元) 串接
        For Each item As ListViewItem In lv.SelectedItems
            Dim cols As New List(Of String)(item.SubItems.Count)
            For Each si As ListViewItem.ListViewSubItem In item.SubItems
                ' 去除頭尾空白，以及 "-"、"▸" 等視覺裝飾字元
                cols.Add(si.Text.Trim(" "c, "-"c, "▸"c, "▶"c))
            Next
            sb.AppendLine(String.Join(vbTab, cols))
        Next

        Try
            Clipboard.SetText(sb.ToString())
            ProgressBar2.Text = $"已複製標題與 {lv.SelectedItems.Count:N0} 列到剪貼簿。"
        Catch ex As System.Exception
            _dbg("剪貼簿存取失敗", ex.Message)
            MessageBox.Show("無法存取剪貼簿，可能被其他程式佔用。", "複製失敗", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
        
        _dbg("結束")
    End Sub
```

## Proposed Changes

### Form1.vb

#### [MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1.vb)
在 `#Region "  ├ ListView 格式工具"` 區塊中加入 `ListViewSelectAll` 與 `ListViewCopyToClipboard` 兩個函數。

---

### Form1_MainTab12.vb

#### [MODIFY] [Form1_MainTab12.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_MainTab12.vb)
- **Lv1_KeyDown**: 將原本 `Keys.A` 與 `Keys.C` 的實作替換為呼叫 `ListViewSelectAll(lv)` 與 `ListViewCopyToClipboard(lv)`。
- **Lv2_KeyDown**: 同上。

---

### Form1_MainTab345.vb

#### [MODIFY] [Form1_MainTab345.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_MainTab345.vb)
- **HandleLv3Lv4Lv5_KeyDown**: 將原本 `Keys.A` 結合虛擬模式防呆的冗長邏輯，替換為直接呼叫 `ListViewSelectAll(lv)` (函數內部會自動處理 `_isCtrl_A`)。
- **HandleLv3Lv4Lv5_KeyDown**: 同樣加上 `Keys.C` 複製的支援。

## Verification Plan

### Manual Verification
1. 啟動應用程式。
2. 進入 Tab1 點擊 ListView1，按下 `Ctrl+A` 確認是否可以全選，按下 `Ctrl+C` 後到 Notepad 貼上確認格式（包含 Header 表頭）是否正確。
3. 進入 Tab2 重複上述動作。
4. 進入 Tab3/Tab4，按下 `Ctrl+A` 確認是否可以快速全選且不會卡頓（確保 `_isCtrl_A` 防止了重複的相似度計算）。
