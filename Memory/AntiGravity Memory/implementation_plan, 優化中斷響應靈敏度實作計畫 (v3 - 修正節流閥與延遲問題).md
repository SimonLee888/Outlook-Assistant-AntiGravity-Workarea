# 優化中斷響應靈敏度實作計畫 (v3 - 修正節流閥與延遲問題)

針對使用者的最新修改，部分邏輯調整導致中斷機制失效。本計畫旨在修復這些問題並確認 Tab2 的運作狀態。

## 使用者回饋分析
使用者提到 Tab1 的 ESC 還是沒反應。經檢查，目前的程式碼有兩點關鍵問題：
1. **`Dim swThrottle` 在迴圈內部**：這導致每次進入迴圈計時器都會歸零，永遠達不到 100ms 的門檻，因此中斷檢查程式碼永遠不會被執行。
2. **`Await Task.Delay(0)`**：`Delay(0)` 實際上不會讓出執行權給 UI 訊息泵（Message Pump），導致系統無法捕捉到 ESC 按鍵。

## 擬議變更

### [Form1_Main.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_Main.vb)

#### [MODIFY] 修正 BuildBfsFolderTree
- 將 `Dim swThrottle` 移出 `Do While` 迴圈外部，或改回 `Static`。
- 將 `Await Task.Delay(0)` 改回 `Await Task.Delay(1)`。

#### [MODIFY] 修正 FetchDirectMailCountsAsync
- 確保迴圈內的 `Await Task.Delay(1)` 加強版邏輯正確運作。

#### [MODIFY] 修正 GetYearCountsForFolder (Tab2)
- 同樣修正迴圈內的 `swThrottle` 宣告位置與 `Delay(1)`，確保 Tab2 的 ESC 也能維持靈敏。

---

## 答覆使用者疑問
- **關於 Tab2 狀態**：`ComputeYearCounts` 此時確實是異步（Async）作業。但因為它內部迴圈的節流閥宣告方式（`Dim` 在迴圈內）與 `Delay` 的數值問題，會導致按下 ESC 時響應遲鈍。修正後，Tab2 的中斷表現將會與 Tab1/Tab3 一致。

## 驗證計畫
1. **測試 Tab1 中斷**：選取大資料夾後按下 ESC，確認能立即停止。
2. **測試 Tab2 中斷**：開始年度統計後按下 ESC，確認能立即停止。
