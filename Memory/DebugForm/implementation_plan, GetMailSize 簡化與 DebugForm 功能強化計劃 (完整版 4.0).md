# GetMailSize 簡化與 DebugForm 功能強化計劃 (完整版 4.0)

本計劃整合所有先前討論確認的細節，分為四大區塊。

---

## 一、[Form1.vb] GetMailSize 簡化

### 問題
`GetMailSize` 的 ① ② 兩段 MAPI 屬性讀取中，`TypeOf` 判斷在 `Try...Catch` 內是多餘的。`CLng()` 本身能處理 Long/Integer/甚至數值字串的轉型，失敗時拋出的例外會被外層 `Catch` 接住。

### 修改前 (L4567-4586)
```vb
' ① MAPI: PR_MESSAGE_SIZE_EXTENDED (0x0E080014, PT_I8) — 64-bit，無溢位風險
Try
    Const PR_SIZE_EXTENDED As String = "http://schemas.microsoft.com/mapi/proptag/0x0E080014"
    Dim val As Object = mail.PropertyAccessor.GetProperty(PR_SIZE_EXTENDED)
    If TypeOf val Is Long Then Return CLng(val)
    If TypeOf val Is Integer Then Return CLng(CInt(val))    ' 某些環境回傳 Integer，安全轉型
    ' todo: try/catch裡面包住的 TypeOf 都可以直接拿掉
Catch ex As System.Exception
    Dbg("GetMailSize ① PR_MESSAGE_SIZE_EXTENDED失敗", ex.Message)
End Try

' ② MAPI: PR_MESSAGE_SIZE (0x0E080003, PT_LONG) — 32-bit，超大郵件理論上溢位
Try
    Const PR_SIZE As String = "http://schemas.microsoft.com/mapi/proptag/0x0E080003"
    Dim val As Object = mail.PropertyAccessor.GetProperty(PR_SIZE)
    If TypeOf val Is Integer Then Return CLng(CInt(val))
    ' todo: try/catch裡面包住的 TypeOf 都可以直接拿掉
Catch ex As System.Exception
    Dbg("GetMailSize ② PR_MESSAGE_SIZE失敗", ex.Message)
End Try
```

### 修改後
```vb
' ① MAPI: PR_MESSAGE_SIZE_EXTENDED (0x0E080014, PT_I8) — 64-bit，無溢位風險
' by AntiGravity, 2026/03/29: 移除 TypeOf 判斷，CLng() 可自動處理 Long/Integer 轉型，
' 若屬性不存在或回傳 Nothing/DBNull，CLng 會拋例外進入 Catch
Try
    Const PR_SIZE_EXTENDED As String = "http://schemas.microsoft.com/mapi/proptag/0x0E080014"
    Return CLng(mail.PropertyAccessor.GetProperty(PR_SIZE_EXTENDED))
Catch ex As System.Exception
    Dbg("GetMailSize ① PR_MESSAGE_SIZE_EXTENDED失敗", ex.Message)
End Try

' ② MAPI: PR_MESSAGE_SIZE (0x0E080003, PT_LONG) — 32-bit，超大郵件理論上溢位
' by AntiGravity, 2026/03/29: 同上，移除 TypeOf 判斷
Try
    Const PR_SIZE As String = "http://schemas.microsoft.com/mapi/proptag/0x0E080003"
    Return CLng(mail.PropertyAccessor.GetProperty(PR_SIZE))
Catch ex As System.Exception
    Dbg("GetMailSize ② PR_MESSAGE_SIZE失敗", ex.Message)
End Try
```

---

## 二、[DebugForm.vb] 效能優化：Shift 多選卡頓修復

### 問題分析
當您用 Shift 多選 100 筆時，`ItemSelectionChanged` 事件會被觸發 **100 次**。
每次觸發都執行：
```vb
For Each item As ListViewItem In lvwDebug.Items   ' 假設目前有 5000 筆
    item.BackColor = Color.White
Next
```
總計 = `5000 × 100 = 500,000` 次背景色重設。**這就是多選時 UI 凍結的元凶。**

### 解決方案：O(1) 局部還原

#### 新增變數 (Region 02 成員變數)
```vb
Private _lastHighlightedPair As ListViewItem   ' 記錄上次被標記的配對行，用於 O(1) 還原
```

#### 修改 lvwDebug_ItemSelectionChanged (L208)
```vb
Private Sub lvwDebug_ItemSelectionChanged(sender As Object, e As ListViewItemSelectionChangedEventArgs)
    If Not e.IsSelected Then Return

    ' ── O(1) 還原上次標記 ──────────────────────────────────
    ' 取代原本的 For Each 全域清除，僅還原 1 個項目的顏色
    If _lastHighlightedPair IsNot Nothing Then
        _lastHighlightedPair.BackColor = Color.White
        _lastHighlightedPair = Nothing
    End If

    ' ── 雙向配對搜尋 ──────────────────────────────────
    Dim selectedItem As ListViewItem = e.Item
    Dim pair As ListViewItem = FindSimilarPair(selectedItem)
    If pair IsNot Nothing Then
        pair.BackColor = Color.LightCyan
        _lastHighlightedPair = pair          ' 記住這次標記的項目
    End If
End Sub
```

---

## 三、[DebugForm.vb] 巢狀雙向配對搜尋 (Stack 計數器)

### 問題
原本 `FindSimilarPair` 只會「向上找第一個 Begin」，無法處理同名函數的巢狀呼叫。
例如：
```
開始: GetMailCount        ← 第一層
  開始: GetMailCount      ← 第二層 (遞迴)
  結束: GetMailCount      ← 第二層結束
結束: GetMailCount        ← 第一層結束  ← 點選這裡應該配對到第一層的開始
```

### 新版 FindSimilarPair 演算法
```vb
Private Function FindSimilarPair(selectedItem As ListViewItem) As ListViewItem
    Dim txt As String = selectedItem.Text
    Dim coreName As String = RemoveBeginEnd(txt)

    ' 判斷方向
    Dim isBegin As Boolean = txt.Contains("開始")
    Dim isEnd As Boolean = txt.Contains("結束")
    If Not isBegin AndAlso Not isEnd Then Return Nothing

    Dim depth As Integer = 0

    If isBegin Then
        ' 向下搜尋配對的「結束」
        For i As Integer = selectedItem.Index + 1 To lvwDebug.Items.Count - 1
            Dim item As ListViewItem = lvwDebug.Items(i)
            Dim itemCore As String = RemoveBeginEnd(item.Text)
            If IsContentSimilar(coreName, itemCore) Then
                If item.Text.Contains("開始") Then
                    depth += 1          ' 同名的巢狀開始，深度 +1
                ElseIf item.Text.Contains("結束") Then
                    If depth = 0 Then Return item   ' 深度歸零 = 正確配對
                    depth -= 1          ' 消耗一層巢狀
                End If
            End If
        Next

    ElseIf isEnd Then
        ' 向上搜尋配對的「開始」
        For i As Integer = selectedItem.Index - 1 To 0 Step -1
            Dim item As ListViewItem = lvwDebug.Items(i)
            Dim itemCore As String = RemoveBeginEnd(item.Text)
            If IsContentSimilar(coreName, itemCore) Then
                If item.Text.Contains("結束") Then
                    depth += 1          ' 同名的巢狀結束，深度 +1
                ElseIf item.Text.Contains("開始") Then
                    If depth = 0 Then Return item   ' 深度歸零 = 正確配對
                    depth -= 1          ' 消耗一層巢狀
                End If
            End If
        Next
    End If

    Return Nothing
End Function
```

---

## 四、[DebugForm.vb] 右鍵選單 (3 個功能)

### 初始化 (DebugForm_Load 內新增)
```vb
' by AntiGravity, 2026/03/29: 右鍵管理選單
Dim ctx As New ContextMenuStrip()
ctx.Items.Add("清除所有項目", Nothing, Sub(s, ev) lvwDebug.Items.Clear())
ctx.Items.Add("計算選取耗時", Nothing, AddressOf CalculateSelectedTimeSpan)
ctx.Items.Add("刪除選取項目", Nothing, AddressOf DeleteSelectedItems)
lvwDebug.ContextMenuStrip = ctx
```

### 功能 1：清除所有項目
- 執行 `lvwDebug.Items.Clear()`。
- `AddMessage3` 內的 `Static lineCount` 不受影響，新 Log 會繼續累加。

### 功能 2：計算選取耗時
- **資料來源**：`item.Tag`（型別 `DebugItemTag`）的 `.timeStamp` 屬性。
- **計算方式**：取所有選取項目中最早與最晚的 `timeStamp`，相減得到總耗時。

```vb
Private Sub CalculateSelectedTimeSpan(sender As Object, e As EventArgs)
    If lvwDebug.SelectedItems.Count < 2 Then
        MessageBox.Show("請至少選取 2 個項目") : Return
    End If

    Dim earliest As Date = Date.MaxValue
    Dim latest As Date = Date.MinValue
    For Each item As ListViewItem In lvwDebug.SelectedItems
        Dim tag = TryCast(item.Tag, DebugItemTag)
        If tag IsNot Nothing Then
            If tag.timeStamp < earliest Then earliest = tag.timeStamp
            If tag.timeStamp > latest Then latest = tag.timeStamp
        End If
    Next

    Dim span As TimeSpan = latest - earliest
    MessageBox.Show(
        $"已選擇 {lvwDebug.SelectedItems.Count} 個項目" & vbCrLf &
        $"時間跨度: {span.TotalMilliseconds:N0} ms ({span.TotalSeconds:N2} s)",
        "計算結果")
End Sub
```

### 功能 3：刪除選取項目
- ListView 刪除後，剩餘項目自動**往上遞補**。
- 各行文字最前方的行號序號**保留不變**（不重新編號）。

```vb
Private Sub DeleteSelectedItems(sender As Object, e As EventArgs)
    lvwDebug.BeginUpdate()
    For Each item As ListViewItem In lvwDebug.SelectedItems
        lvwDebug.Items.Remove(item)
    Next
    lvwDebug.EndUpdate()
End Sub
```

---

## Open Questions
- 已回答：刪除後行號**保留原本**，不重新編號。 ✅

## Verification Plan

### Manual Verification
1. 按住 Shift 多選 100 行，確認不再有凍結卡頓。
2. 點選含遞迴呼叫的「開始」節點，確認標記的是**同一層級**的正確「結束」節點。
3. 右鍵「清除」後新增 Log，確認行號繼續累加。
4. 多選 5 行，右鍵「計算耗時」，確認顯示的毫秒數正確。
5. 多選 3 行，右鍵「刪除」，確認僅刪除選取行、其餘行往上遞補、行號不變。
6. 測試 GetMailSize，確認統計結果與修改前一致。
