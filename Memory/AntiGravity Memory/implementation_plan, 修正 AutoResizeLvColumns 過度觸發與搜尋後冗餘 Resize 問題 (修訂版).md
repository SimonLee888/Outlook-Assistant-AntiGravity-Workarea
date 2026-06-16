# 修正 AutoResizeLvColumns 過度觸發與搜尋後冗餘 Resize 問題 (修訂版)

## 問題分析
根據截圖與程式碼巡檢，確認 `AutoResizeLvColumns` 被頻繁重複觸發的原因如下：

1.  **視窗最大化/還原觸發競爭**：
    *   `Form1_Resize` 偵測到 `WindowState` 改變時會主動呼叫 `AutoResizeLvColumns`。
    *   與此同時，視窗尺寸劇變會引發所有 ListView 的 `Resize` 事件，進而執行 `HandleLvResize` (100ms 節流)。
    *   這導致在最大化瞬間，同一個 ListView 會先被 Form 強制 Resize 一次，100ms 後又被計時器 Resize 第二次。

2.  **搜尋結束引發佈局連動**：
    *   `ShowLv3Result` 內的 `ListView3.Invalidate()` 在虛擬模式下可能觸發不必要的佈局檢查。
    *   搜尋結果填入後，若捲軸出現/消失，會引發 ListView `Resize` 事件，再次觸發 `HandleLvResize`。

3.  **無效的事件觸發**：
    *   目前的 `HandleLvResize` 只要收到事件就啟動計時器，沒有判斷寬度是否真的改變。

## 解決方案

### 1. 強化 `Form1_Resize` 的 WindowState 偵測
*   保留 `WindowState` 變化偵測，但改為「標記」而非「立即執行」。
*   當偵測到最大化/還原時，僅針對**當前活動的 ListView** 發送一次 Resize 請求，並透過統一的節流計時器處理，避免與 `HandleLvResize` 衝突。

### 2. 引入寬度比對機制 (防抖與過濾)
*   在 `HandleLvResize` 啟動計時器前，先比對目前的 `Width` 是否與上次執行時不同。若寬度未變，則無視該事件（例如僅高度改變或無效觸發）。

### 3. 精簡搜尋結果呈現邏輯
*   移除 `ShowLv3Result` 中冗餘的 `ListView3.Invalidate()`。

## 擬定修改點

### [MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1.vb)

#### `Form1_Resize`
*   修改 `WindowState` 檢查邏輯。
*   當狀態改變時，呼叫 `HandleLvResize(GetActiveListView(), EventArgs.Empty)`，利用 ListView 自身的節流機制來同步執行，不再繞過計時器直接呼叫 `AutoResizeLvColumns`。

#### `HandleLvResize`
*   在計時器啟動前，判斷 `sender.Width` 是否與該 ListView 紀錄的 `lastWidth` 不同。
*   (需在 ListView 的 `Tag` 或全域 Dictionary 暫存寬度)。

#### `AutoResizeLvColumns`
*   內部增加防護：若處理時的 `lv.Width` 與上次完成時一致，則提早 Return。

### [MODIFY] [Form1_MainTab345.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity測試區%29/Form1_MainTab345.vb)

#### `ShowLv3Result`
*   移除 `ListView3.Invalidate()`。

## 驗證計畫
1.  **最大化/還原測試**：雙擊標題列，確認 Debug Message 中 `AutoResizeLvColumns` 僅出現一次，且畫面欄位比例正確。
2.  **拖曳縮放測試**：緩慢拖曳邊框，確認計時器節流運作正常，且停止拖曳後才執行一次計算。
3.  **搜尋觸發測試**：執行 Tab3 搜尋，確認搜尋結束後不會因為 Invalidate 產生多餘的 Resize 紀錄。
