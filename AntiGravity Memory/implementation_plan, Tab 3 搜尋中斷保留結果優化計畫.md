# Tab 3 搜尋中斷保留結果優化計畫

## 目標
在 Tab 3 附件搜尋過程中按下 ESC 鍵時，不直接中斷並顯示「已中斷」，而是將目前已處理完成的郵件結果呈現給使用者。

## 待修改組件

### 1. Form1_MainTabs.vb
#### [MODIFY] [FilterByAttachDetailsAsync](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_MainTabs.vb)
- 在 `For` 迴圈外包裹 `Try...Catch OperationCanceledException`。
- **中斷時動作**：記錄 `Dbg` 指出中斷，並直接 `Return resultList`。這能讓呼叫端 `Button3_Click` 誤以為處理已完成（但只有部分），從而繼續執行顯示邏輯。

#### [MODIFY] [Button3_Click](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_MainTabs.vb)
- **Step 3 (收集郵件清單)**：目前這是一個 `For Each folder In folderList` 迴圈。如果在中途按 ESC，`Task.Delay` 或 `GetCachedAttachMailList` 可能拋出例外。
- 將該迴圈包裹在 `Try...Catch` 中，中斷時記錄 Dbg 並跳出迴圈 (`Exit For`)，保留 `targetMails` 內的既有內容，讓後續的 Pipeline 過濾和顯示邏輯能繼續執行。
- **Step 2 (資料夾收集)**：保持現狀。如果資料夾都還沒收集到任何一個就中斷，顯示清單意義不大。

## 驗證計畫
1. **手動測試 (Tab 3)**：
    - 啟動大型附件搜尋。
    - 在進度條跑動（比對 Phase 2）中途按下 ESC。
    - **預期結果**：進度條顯示中斷訊息，但 ListView3 依然會載入在按 ESC 前已經比對成功的郵件。
2. **手動測試 (清單蒐集階段)**：
    - 在 Phase 1（正在讀取資料夾郵件清單）時按下 ESC。
    - **預期結果**：同樣會顯示出目前已載入的郵件集。

---
> [!IMPORTANT]
> 此修改將改變「中斷即終止」的原有語義，轉變為「中斷並結算」。
