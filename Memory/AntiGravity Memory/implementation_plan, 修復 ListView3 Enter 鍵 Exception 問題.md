# 修復 ListView3 Enter 鍵 Exception 問題

使用者反應 ListView3 使用 `MouseDoubleClick` 開啟郵件正常，但按下 `Enter` 鍵會跳出 Exception。經查是因為 ListView3 已改為 `VirtualMode` (虛擬模式)，而全域的 `HandleListViewKeyPress` 仍試圖以傳統方式存取 `SubItems` 導致。

## 使用者回覆

1. **ListView3 的 MouseDoubleClick 事件在哪裡？**
   - 位於 `Form1_MainTabs.vb` 第 1808 行的 `ListView3_MouseDoubleClick` 子程序中。

2. **為什麼用 Enter 會跳 Exception？**
   - 因為 ListView3 目前處於 **VirtualMode (虛擬模式)**。在虛擬模式下，ListView 內部並不真正持有 `ListViewItem` 物件。
   - `Form1.vb` 中的 `HandleListViewKeyPress` 函式在處理 `ListView3` 的 `Enter` 鍵時，試圖透過 `lv.SelectedItems(0).SubItems(5).Text` 來取得 EntryID。
   - 在虛擬模式下，這樣的存取方式會因為 SubItems 尚未建立或物件不完整而觸發 `ArgumentOutOfRangeException` 或其他例外。
   - 而 `MouseDoubleClick` 能成功是因為它使用了 `GetItemFromPoint` 從底層資料源 (`_lv3MailList`) 直接取得資料，避開了對 `ListViewItem` 的依賴。

## 建議修改方案

### [Component: UI 事件處理]

#### [MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1.vb)

修改 `HandleListViewKeyPress` 函式中針對 `ListView3` 的處理邏輯：
- 改用 `lv.SelectedIndices(0)` 取得索引。
- 從 `_lv3MailList` 記憶體清單中直接讀取 `EntryID`。
- 加入必要的 Try-Catch 保護。

```vb
' 修改前 (約 1383 行)
If lv.SelectedItems.Count = 0 Then Return
OpenMailByEntryID(lv.SelectedItems(0).SubItems(5).Text)

' 修改後
If lv.SelectedIndices.Count = 0 Then Return
Try
    Dim idx As Integer = lv.SelectedIndices(0)
    If idx >= 0 AndAlso idx < _lv3MailList.Count Then
        OpenMailByEntryID(_lv3MailList(idx).EntryID)
    End If
Catch ex As Exception
    _dbg("錯誤", "Enter 開啟郵件失敗: " & ex.Message)
End Try
```

## 驗證計畫

### 手動測試
1. 開啟程式並切換至「尋找附件」分頁 (Tab3)。
2. 執行一次搜尋以填滿 ListView3。
3. 選取一封郵件，按下 `Enter` 鍵。
4. **預期結果**：郵件應順利開啟，不再出現 Exception。
5. 雙擊郵件，確認原有功能依然正常。
