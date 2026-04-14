# 補全非同步中斷異常處理計畫

為了確保 ESC 中斷 (CancellationToken) 機制穩定運行，需要補齊以下函式中遺漏的 `Try...Catch OperationCanceledException` 區塊。

## 待修改組件

### 1. Form1_Outlook.vb
#### [MODIFY] [Form1_Outlook.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_Outlook.vb)
- **GetSubFolderList**: 
  - 在函式內部 `While` 迴圈外層加上 `Try...Catch OperationCanceledException`。
  - 中斷時記錄 Dbg 並回傳目前已蒐集的 `result` (部分結果)。

### 2. Form1_MainTabs.vb
#### [MODIFY] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_MainTabs.vb)
- **BuildBfsFolderTree**: 
  - 加上 `Try...Catch`，中斷時回傳已掃描的部分資料夾清單。
- **FetchDirectMailCountsAsync**: 
  - 加上 `Try...Catch`，中斷時回傳 `True` (表示已被取消)，讓呼叫者知道要停止後續計算。
- **ShowMonthView**: 
  - 加上 `Try...Catch`，確保月份統計被 ESC 中斷時，UI 能恢復 `Cursor = Default` 且不會跳出異常。

## 驗證計畫
1. 在 Tab1 顯示資料夾、Tab2 年度統計、Tab3 附件搜尋期間，頻繁按下 ESC 鍵。
2. 檢查 Debug 視窗是否顯示 `├ 中斷` 訊息。
3. 確認程式不會彈出 `TaskCanceledException` 或 `OperationCanceledException` 錯誤對話框，且 UI 滑鼠指標能正確恢復為箭頭。
