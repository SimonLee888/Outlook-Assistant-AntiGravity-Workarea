# 整合 ListView 懸停事件處理邏輯

本計畫旨在將 `ListView` 的 `MouseMove` 與 `MouseLeave` 兩個事件處理函式整合為一個單一的維護點 `HandleListViewHoverShared`。這將簡化程式碼維護，並確保滑鼠懸停效果在不同事件觸發時邏輯一致。

## 使用者評論要求
> [!IMPORTANT]
> 所有的修改將嚴格遵守 `by AntiGravity, 2026/04/03` 的標記規範，並保留現有的調試與歷史註解。

## 擬議變更

### Form1.vb

#### [MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1.vb)

1.  **修改 `InitListView` (L321-L322)**：
    *   將 `lv.MouseMove` 與 `lv.MouseLeave` 的事件處理程序從 `HandleListViewMouseMoveShared` / `HandleListViewMouseLeaveShared` 改為指向新的 `HandleListViewHoverShared`。

2.  **新增 `HandleListViewHoverShared` (替換 L921-L939)**：
    *   實作整合邏輯，使用 `EventArgs` 作為基底參數，並透過 `TryCast` 辨識事件類型。
    *   合併處理「清除舊背景色」與「設定新背景色」的邏輯。

3.  **移除 `HandleListViewMouseMoveShared` 與 `HandleListViewMouseLeaveShared`**。

## 開放性問題
無。目前的方案已經過評估，符合 WinForms 的事件處理機制。

## 驗證計畫

### 自動化測試
*   手動檢查編譯是否通過。

### 手動驗證
1.  **懸停測試**：移動滑鼠進入 各分頁的 ListView 項目，確保背景色正確變為 `ThemeColors.MercuryGray`。
2.  **離開測試**：滑鼠移出項目或移出整個 ListView 範圍，確保背景色恢復為 `Color.Empty`。
3.  **切換測試**：在不同分頁（Tab1~Tab5）間切換，驗證所有 ListView 是否都正常套用了新的整合處理器。
