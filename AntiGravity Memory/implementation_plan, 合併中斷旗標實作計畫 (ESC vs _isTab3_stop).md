# 合併中斷旗標實作計畫

本計畫旨在將專案中冗餘的 `_isTab3_Stop` 旗標移除，並統一使用 `_cancelRequested` 作為全域的 ESC/停止按鈕中斷訊號。

## 使用者評論要求
> [!IMPORTANT]
> 此變更會將 Tab3 的「停止」按鈕與全域的 ESC 中斷邏輯綁定在一起。由於 UI 是單執行緒運作，這不會導致競爭問題，且能簡化程式碼邏輯。

## 擬議變更

### [Form1.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1.vb)

#### [MODIFY] 宣告與 KeyDown 事件
- 移除 `_isTab3_Stop` 的變數宣告。
- 在 `Form1_KeyDown` 中移除對 `_isTab3_Stop` 的賦值語句。

---

### [Form1_Main.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_Main.vb)

#### [MODIFY] Tab3 邏輯
- `Button3_Click`: 重置旗標時移除 `_isTab3_Stop = False`。
- `Button3_Stop_Click`: 改為設置 `_cancelRequested = True`。
- `FilterByAttachmentDetailsAsync`: 迴圈內的結束條件檢查改用 `_cancelRequested`。

---

### [Form1_ComL3.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_ComL3.vb)

#### [MODIFY] L3 底層檢查
- `PreloadAttachmentCacheRDOAsync`: 檢查改用 `_cancelRequested`。
- `GetMailWithAttachmentL3`: 檢查改用 `_cancelRequested`。

## 驗證計畫

### 手動驗證
- 在 Tab3 執行附件搜尋時，按下「停止」按鈕，確認搜尋能正確中止且 UI 回復正常。
- 在 Tab3 執行附件搜尋時，按下「ESC」鍵，確認搜尋能正確中止且 UI 回復正常。
- 檢查 Tab1 與 Tab2 的 ESC 中斷功能是否依然正常運作。
