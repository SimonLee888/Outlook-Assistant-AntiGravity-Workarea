# 重構 ListView KeyPress 事件處理計畫

根據您的建議，我們將把原本在 `Form1.vb` 中整合開發的 `HandleListViewKeyPress` 拆分回各自專屬的 `Handles` 事件處理器中。這能讓每個列表的行為更加獨立、易於維護，且符合 Windows Form 的標準事件掛載模式。

## 調整內容

### 1. 搬移邏輯至獨立 Handles 事件
- **ListView1_KeyPress**: 處理資料夾導覽 (Enter 進入、ESC 退回、Ctrl-A 全選)。
- **ListView2_KeyPress**: 處理年度/月份導覽與加總統計。
- **ListView3_KeyPress**: 處理郵件開啟與 ESC 清除選取。

### 2. 清理通用初始化邏輯
- 從 `Form1.vb` 的 `InitListView` 函式中移除對 `HandleListViewKeyPress` 的手動掛載 (`AddHandler`)。
- 刪除不再使用的通用函式 `HandleListViewKeyPress`。

---

## 建議修改方案

### [Component: UI 事件處理重構]

#### [MODIFY] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_MainTabs.vb)

在適當的 Region 內（分別對應 Tab1, Tab2, Tab3）新增各自的 `KeyPress` 處理器，並從 `Form1.vb` 搬回邏輯。

```vb
' ListView1 範例
Private Async Sub ListView1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles ListView1.KeyPress
    _dbg("開始", e.KeyChar)
    Dim cToken As CancellationToken = OkayNowYouHaveToken()
    ' (搬移 ListView1 專屬邏輯)
End Sub

' ListView2 範例
Private Async Sub ListView2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles ListView2.KeyPress
    _dbg("開始", e.KeyChar)
    Dim cToken As CancellationToken = OkayNowYouHaveToken()
    ' (搬移 ListView2 專屬邏輯)
End Sub

' ListView3 範例 (包含剛修復的邏輯)
Private Async Sub ListView3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles ListView3.KeyPress
    _dbg("開始", e.KeyChar)
    Dim cToken As CancellationToken = OkayNowYouHaveToken()
    ' (搬移 ListView3 專屬邏輯)
End Sub
```

#### [MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1.vb)

1.  **InitListView**: 移除 `AddHandler lv.KeyPress, AddressOf HandleListViewKeyPress`。
2.  **HandleListViewKeyPress**: 刪除整個函式內容。

---

## 驗證計畫

### 手動測試
1. **各分頁功能確認**：
   - Tab1：Enter 是否能進入資料夾、ESC 是否能退回。
   - Tab2：Enter 是否能切換年度/月份視圖。
   - Tab3：Enter 是否能多選開啟郵件、ESC 是否能清除選取。
2. **Debug 訊息檢查**：
   - 確認 `_dbg()` 輸出的「開始」/「結束」標記是否正確顯示對應的事件名稱。

### 程式碼品質檢查
- 使用 `view_file` 確認 `Handles` 語法正確且無邏輯遺漏。
