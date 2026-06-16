# 實作方案 4：Resize 期間暫停 OwnerDraw 繪製

## 問題分析
雖然使用了 `WM_SETREDRAW` 和 `BeginUpdate`，但在 `ListView` 欄位寬度劇烈變動時，大量的 `SubItems` 座標重算仍會導致 `EndUpdate` 瞬間出現文字「滑動」到新位置的視覺殘影（Layout Reflow 痕跡）。

## 解決方案
利用我們已開啟的 `OwnerDraw` 模式，在 `AutoResizeLvColumns` 執行期間透過旗標控制，暫停所有繪製動作，直到座標全部結算完成。

### 核心邏輯
1.  **定義全域旗標**：`Private _isResizingLv As Boolean = False`
2.  **鎖定繪製**：在 `AutoResizeLvColumns` 開始時將旗標設為 `True`。
3.  **攔截繪製事件**：在所有 ListView 的 `DrawSubItem` 事件開頭增加檢查：若 `_isResizingLv = True`，則直接退出不執行任何繪製動作。
4.  **恢復與重繪**：在 `AutoResizeLvColumns` 結尾將旗標設為 `False`，並呼叫一次完整的 `lv.Invalidate()`。

## 擬定修改點

### [MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1.vb)
*   **成員變數區域**：新增 `_isResizingLv`。
*   **`AutoResizeLvColumns`**：
    *   `Try` 區塊前：`_isResizingLv = True`。
    *   `Finally` 區塊中：`_isResizingLv = False`。
*   **各 ListView 的繪製事件**：
    *   `HandleLvDrawSubItem` (或相關繪製邏輯)：開頭增加 `If _isResizingLv Then Return`。

### [MODIFY] [Form1_Outlook.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_Outlook.vb) 等繪製邏輯所在文件
*   同步在 `DrawSubItem` 處理程序中加入旗標檢查。

## 驗證計畫
1.  **平滑度測試**：快速縮放視窗，觀察 ListView 內容是否在縮放期間保持「靜止」或「空白」，直到停止動作後才「砰」地一聲出現在正確位置。
2.  **防止黑屏測試**：確保 `Finally` 區塊一定會恢復旗標，否則 UI 會永久停止繪製。
