# 修復虛擬模式下 Ctrl+A 全選導致的 InvalidOperationException

當 ListView（如 Tab3 的 ListView3）開啟 `VirtualMode = True` 時，若嘗試透過 `For Each item As ListViewItem In lv.Items` 進行列舉，會擲回 `System.InvalidOperationException`。這是因為在虛擬模式下，`Items` 集合並非實體存在，系統要求使用索引子存取。

## 使用者評論與回饋要求

> [!IMPORTANT]
> 此修正將影響 Tab3, Tab4, Tab5 共用的鍵盤處理邏輯。雖然目前僅 Tab3 明確開啟了虛擬模式，但此修正採用通用寫法，能同時相容實體模式與虛擬模式。

## 提出變更

### [Component] Form1_MainTab345.vb

#### [MODIFY] [Form1_MainTab345.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_MainTab345.vb)

修改 `HandleLv3Lv4Lv5_KeyDown` 函數中的全選邏輯：
1. **移除 `For Each item In lv.Items`**：這是造成 Exception 的主因。
2. **改用 `SelectedIndices` 加速**：
   - 針對虛擬模式：直接透過索引循環將索引加入 `SelectedIndices`。
   - 考慮效能：對於極大量數據（如 Tab3 可能有萬筆），使用 `lv.SelectedIndices.Add(i)` 依然比建立 `ListViewItem` 快得多。

修正後的代碼邏輯：
```vb
        ElseIf e.Control AndAlso e.KeyCode = Keys.A Then
            lv.BeginUpdate()
            If lv.VirtualMode Then
                ' 虛擬模式下不可枚舉 Items，改用索引循環或直接操作 SelectedIndices
                ' by Gemini 3 Flash, 2026/05/09: 修復虛擬模式全選當機問題
                For i As Integer = 0 To lv.Items.Count - 1
                    lv.SelectedIndices.Add(i)
                Next
            Else
                ' 實體模式維持原樣
                For Each item As ListViewItem In lv.Items
                    item.Selected = True
                Next
            End If
            lv.EndUpdate()
            e.Handled = True
            e.SuppressKeyPress = True
```

## 驗證計畫

### 手動驗證
1. 在 Tab3 執行搜尋，確保 ListView3 進入虛擬模式並顯示結果。
2. 在 ListView3 中按下 `Ctrl+A`。
3. 確認不再出現 `InvalidOperationException` 且所有項目正確被選取（背景變色）。
4. 測試 Tab4/Tab5（非虛擬模式）的 `Ctrl+A` 功能是否依然正常。
5. 在 ListView3 按下 `ESC` 確認取消全選功能正常。
