# 優化中斷響應靈敏度實作計畫 (v2 - 基於 100ms 節流)

使用者反應 Tab1 與 Tab2 在長時間執行時 ESC 中斷反應遲鈍。經分析，主因是部分步驟為同步阻塞操作，或是非同步 Yield 頻率不足及檢查點遺漏。此外，為了兼顧掃描效能，我們採用與進度條同步的 100ms 節流機制。

## 使用者回饋與設計說明
> [!IMPORTANT]
> **關於 GetMailCountAll 的修改理由**：
> 當使用者在 Tab 2 選取多個資料夾時，系統會先呼叫 `GetMailCountAll` 來計算「郵件總數」作為進度條分母。若不對此函式加入 Yield 與中斷點，在讀取大型資料夾的「預讀階段 (Read...)」按下 ESC จะ沒有反應，視窗會卡死直到讀取完畢。因此，我們必須在此處加入節流中斷機制。

## 擬議變更：100ms 節流策略
為了最小化 `Async/Await` 對效能的影響，我們統一採用以下邏輯：
```vb
Static swThrottle As New Stopwatch() : If Not swThrottle.IsRunning Then swThrottle.Start()
If swThrottle.ElapsedMilliseconds >= 100 Then
    ' 1. 進度回報 (如果有)
    ' 2. 重置計時器 swThrottle.Restart()
    ' 3. 讓出 UI 資源 Await Task.Delay(1)
    ' 4. 檢查中斷 _cancelRequested
End If
```

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

#### [EXCLUDED] GetMailCountAll (L3 Core)
- 依照使用者要求，此函式保持原始邏輯 (i Mod 10 Yield)，不進行節流優化。
- 後果：Tab 2 預讀大資料夾階段的 ESC 響應可能仍有遲鈍，但掃描本身邏輯不受影響。

#### [MODIFY] GetSubFolderList (L3 BFS)
- 雖然此函數目前是同步執行且由多處調用，但在 `While queue.Count > 0` 迴圈中加入 `If _cancelRequested Then Exit While` 可讓背景呼叫端更早得知取消訊號。

## 驗證計畫

### 手動驗證
1. **Tab1 壓力測試**: 點擊一個擁有數千個子資料夾的 PST 根目錄，在進度條跳動時按下 ESC，確認能立即顯示「已中斷」。
2. **Tab2 多選測試**: 在 SimTree2 同時選取多個 PST，在「正在預讀」階段（顯示總數計算時）按下 ESC，確認能立即中斷而不必等待預讀結束。
3. **Tab2 執行測試**: 在年度統計進行時按下 ESC，確認能立即停在當前進度。
