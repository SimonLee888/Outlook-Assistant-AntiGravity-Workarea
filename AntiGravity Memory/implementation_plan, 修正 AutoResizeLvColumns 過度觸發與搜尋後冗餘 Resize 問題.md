# 修正 AutoResizeLvColumns 過度觸發與搜尋後冗餘 Resize 問題

## 問題分析
根據使用者提供的 Debug Message 截圖與程式碼檢查，發現以下問題：

1.  **視窗最大化/還原觸發多次 AutoResize**：
    *   `Form1_Resize` 雖然有 `WindowState` 檢查，但 WinForms 在視窗狀態改變期間會發送多次 `Resize` 訊息。
    *   此外，`Form1.vb` 中的 `HandleLvResize` (由各 ListView 的 `Resize` 事件觸發) 雖然有 100ms 節流，但在視窗最大化時，每個 ListView 都會獨立觸發自己的 Resize 事件，導致 Debug 紀錄中出現大量 `ListView3` 的 Resize 紀錄。
    *   `Form1_Resize` 內的 `AutoResizeLvColumns(GetActiveListView)` 與 ListView 自身的 `HandleLvResize` 存在邏輯重疊。

2.  **搜尋結束觸發二次 Resize**：
    *   在 `ShowLv3Result` 中，程式碼呼叫了 `ListView3.Invalidate()`。
    *   最主要的原因是：`Bt3_Click` 本身並沒有直接呼叫 `AutoResize`，但截圖顯示在 `結束 Form1.Bt3_Click` 之後緊接著出現了兩次 `AutoResizeLvColumns`。這通常是因為搜尋結果填入後，ListView 的捲軸出現/消失觸發了 `Resize` 事件，進而執行了 `HandleLvResize`。

## 解決方案

### 1. 優化 Form1_Resize
*   移除 `Form1_Resize` 中重複的 `AutoResizeLvColumns` 呼叫。
*   既然所有 ListView 都已經掛載了 `HandleLvResize` (含 100ms 節流)，視窗縮放時自然會由各個 ListView 處理自己的欄寬調整，不需要在 Form 層級額外手動觸發。

### 2. 精簡搜尋結果呈現邏輯
*   在 `ShowLv3Result` 中移除手動的 `ListView3.Invalidate()`。設定 `VirtualListSize` 本身就會引發重繪，額外的 `Invalidate` 可能導致重複的繪製請求。

### 3. 強化 Resize 節流與過濾
*   在 `HandleLvResize` 中增加判定：如果寬度沒有實質改變，則不啟動計時器。
*   在 `AutoResizeLvColumns` 內部也增加「寬度未變則跳過」的防護。

## 擬定修改點

### [MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1.vb)

#### `Form1_Resize`
*   移除 `AutoResizeLvColumns(GetActiveListView)`，僅保留 `SyncDebugFormPosition()`。

#### `HandleLvResize`
*   增加 `lv.Width` 檢查，若寬度與上次紀錄相同則不觸發。

#### `AutoResizeLvColumns`
*   增加紀錄上一次處理的寬度，若寬度未變則提早結束。

### [MODIFY] [Form1_MainTab345.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_MainTab345.vb)

#### `ShowLv3Result`
*   移除 `ListView3.Invalidate()`。

## 驗證計畫
1.  **視窗縮放測試**：執行最大化與還原，觀察 Debug Message 是否減少。預期應只有 1-2 次（受計時器節流影響）。
2.  **搜尋觸發測試**：執行 Tab3 搜尋，觀察結尾是否仍有兩次 Resize。
3.  **效能檢查**：確認欄寬調整是否依然平滑。
