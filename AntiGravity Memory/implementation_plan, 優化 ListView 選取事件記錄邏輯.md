# 優化 ListView 選取事件記錄邏輯

此計畫旨在優化 `ListView2` 與 `ListView4` 的 `SelectedIndexChanged` 事件處理流程。透過調整 `_dbg` 記錄位置與選取項檢查順序，消除因 Windows Forms 原生行為造成的重複日誌輸出，使調試訊息更精確。

## 擬議變更

### [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_MainTabs.vb)

#### [MODIFY] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_MainTabs.vb)

- **Lv2_SelectedIndexChanged**:
    - 將 `If ListView2.SelectedItems.Count = 0 Then Return` 提前至 `_dbg("開始")` 之前。
    - 確保只有在真正選中項目時才記錄日誌。

- **Lv4_SelectedIndexChanged**:
    - 將 `lv.SelectedItems.Count = 0` 的判斷移至最前面（甚至在 `Await Task.Delay` 之前）。
    - 將 `_dbg("開始")` 移至判斷之後。

## 驗證計畫

### 手動驗證
- 開啟 Debug 視窗。
- 在 Tab2 與 Tab4 中使用方向鍵移動游標。
- 確認每一格移動只會產生一組「開始/結束」紀錄。
