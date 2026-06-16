# Tab3 搜尋功能快捷鍵優化計畫

此計畫旨在提升 Tab3 (依附件條件搜尋) 的使用者體驗，讓使用者在輸入大小限制 (最小值與最大值) 後，可以直接按 Enter 鍵啟動搜尋，而不需要移動滑鼠點擊按鈕。

## 使用者評論與回饋要求
> [!IMPORTANT]
> 預計在 `InitTab3UI` 中使用 `AddHandler` 進行事件掛載。若 `NumericUpDown` 的 `KeyDown` 事件在某些輸入狀態下無法正確攔截 Enter 鍵（例如編輯模式尚未退出），可能需要微調，但根據現有程式碼風格，使用 `KeyDown` 是最一致的做法。

## 擬議變更

### [表單初始化層]

#### [MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1.vb)
- 在 `InitTab3UI` 函數中，為 `NumberMin` 和 `NumberMax` 加入 `KeyDown` 事件處理器。
- 邏輯描述：當偵測到 `Keys.Enter` 時，呼叫 `Button3.PerformClick()`。

## 開放式問題
1. 暫無。

## 驗證計畫

### 自動化測試
- 無（本項目主要為 UI 邏輯變過）。

### 手動測試 (請使用者確認)
1. 切換至 Tab3。
2. 在「最小值」或「最大值」的數字框內點一下進入編輯狀態。
3. 修改數字後按下鍵盤上的 **Enter** 鍵。
4. 確認下方是否立即開始搜尋（ProgressBar 應顯示準備中或讀取中）。
