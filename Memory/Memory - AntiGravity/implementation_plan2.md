# RDO 平行處理加速 GetMailCountAll 與 GetFolderCountAll 計畫

使用者希望將 `GetMailCountAll` 與 `GetFolderCountAll` 也比照 `GetFolderSize` 的作法，導入 Redemption (RDO) 的多執行緒平行處理，以大幅提升遍歷與統計的速度。同時，更新 `ListView1` 右鍵選單以使用最新的平行計算函數。

## Proposed Changes

### 1. `GetFolderCountAll` 導入 RDO 平行遍歷
目前 `GetFolderCountAll` 是單純呼叫 OOM 的 `GetSubFolderList` 來進行 BFS 展開，所有的 COM 呼叫都在單一 UI 執行緒上排隊。
我們將在 `GetFolderCountAll` 中加入 RDO 的專屬平行路徑：
- 檢查 `_rdo IsNot Nothing`。
- 在 `Task.Run` 中使用基於 RDO 的遞迴或平行遍歷（Concurrent Queue + Parallel.ForEach）。
- 由於 RDO 是 free-threaded，可以運用多核同時遍歷子資料夾樹狀結構，完全避開 OOM 的 `GetSubFolderList`。

### 2. `GetMailCountAll` 升級為 RDO 平行遍歷
目前的 `GetMailCountAll` 只有試圖讀取 `TotalItemCount`，但這可能未能反映整個檔案樹。改寫方式：
- 將原本單點讀取改為與 `GetFolderCountAll` 一樣的 RDO 平行 BFS 展開。
- 對每一個展開的 `rdoFolder`，平行累加其 `rdoFolder.Items.Count`。

### 3. `GetSubFolderList` 平行化說明
由於 `GetSubFolderList` 的設計目的是回傳 OOM 的 `List(Of Outlook.Folder)`，而 OOM COM 物件會被限制在建立它們的 STA (UI) 執行緒上，因此「平行回傳 OOM 物件清單」容易引發執行緒違規。
**但是**，既然統計函數（如 `GetFolderCountAll` 和 `GetMailCountAll`）改用 RDO 平行計算，它們就**不再需要**呼叫 `GetSubFolderList` 了！因此直接平行的效益將由前面兩點完全吸收。

### 4. 更新 `ListView1_ItemMenu`
在 [Form1.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/AntiGravityTest/Form1.vb) 1567 行的 `ListView1_ItemMenu` 裡面，目前右鍵選單還在呼叫舊版的 `Await GetFolderSizeLegacy(folder)`。
我們將其修改為：
- `Await GetFolderSizeAll(folder)` 以利用新的極速平行架構，這也會更符合「計算資料夾大小 (包含子層)」的使用者期待。

#### 修改檔案清單
- **[MODIFY] Form1.vb**
  - 重構 `GetFolderCountAll` 增設 `⓪ Redemption 平行展開` 路徑。
  - 重構 `GetMailCountAll` 增設真實的 `平行 RDO 累加`。
  - 修改 `ListView1_ItemMenu`，將 `GetFolderSizeLegacy` 替換為 `GetFolderSizeAll`。

## Verification Plan

### 自動化 / 內部邏輯驗證
1. 確保程式編譯無誤，所有的 RDO 物件在底層 `Task.Run` 都有釋放 `Marshal.ReleaseComObject`。
2. 檢查 `GetFolderCountAll` 與 `GetMailCountAll` 統計總數正確無誤。

### 使用者手動驗證
使用者對 `ListView1` 的項目點選右鍵計算容量時，會使用最新版的 `GetFolderSizeAll()` 並且快速完成運算；切換資料夾造成的總數統計也會極速顯示。
