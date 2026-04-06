# 優化中斷響應靈敏度實作計畫

使用者反應 Tab1 與 Tab2 在長時間執行時 ESC 中斷反應遲鈍。經分析，主因是部分步驟為同步阻塞操作，或是非同步 Yield 頻率不足及檢查點遺漏。

## 使用者評論要求
> [!IMPORTANT]
> 此變更將把關鍵的 BFS 掃描改為非同步模式，並增加中斷檢查點，這將大幅提升 UI 線程在執行大任務時的響應能力。

## 擬議變更

### [Form1_Main.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_Main.vb)

#### [MODIFY] Tab1: ComputeFolderStatsAsync 相關
- **`BuildBfsFolderTree`**: 改為 `Async Task(Of List(Of FolderBfsEntry))`。
  - 在 `Do While` 迴圈內加入 `If _cancelRequested Then Return New List(Of FolderBfsEntry)`。
  - 加入 `Await Task.Yield()` 確保 UI 不會被大量資料夾遍歷卡死。
- **`FetchDirectMailCountsAsync`**: 
  - 將 `Await Task.Yield()` 改為 `Await Task.Delay(1)` (與 Tab3 對齊，提供更好的訊息泵處理)。
  - 確保 `_cancelRequested` 檢查緊跟在 Await 之後。

#### [MODIFY] Tab2: SimTree2_AfterSelect 與核心統計
- **`SimTree2_AfterSelect`**: 在計算 `totalMailCount` 的 `For Each` 迴圈中加入 `If _cancelRequested Then Return` 檢查點。
- **`GetYearCountsForFolder`**: 在批次讀取迴圈中，確保 `Await Task.Yield()` 或 `Task.Delay(1)` 的頻率足以處理 ESC 事件。

---

### [Form1_ComL3.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_ComL3.vb)

#### [MODIFY] GetSubFolderList (L3 BFS)
- 雖然此函數由多處調用，但在主執行緒中執行時仍可能阻塞。
- 檢查是否能在 `While queue.Count > 0` 迴圈中加入 `If _cancelRequested Then Return result` 以提早退出。

## 驗證計畫

### 手動驗證
1. **Tab1 壓力測試**: 點擊一個擁有數千個子資料夾的 PST 根目錄，在進度條跳動時按下 ESC，確認能立即顯示「已中斷」。
2. **Tab2 多選測試**: 在 SimTree2 同時選取多個 PST，在「正在預讀」階段（顯示總數計算時）按下 ESC，確認能立即中斷而不必等待預讀結束。
3. **Tab2 執行測試**: 在年度統計進行時按下 ESC，確認能立即停在當前進度。
