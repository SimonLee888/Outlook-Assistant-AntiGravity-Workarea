# 替換 GetFolderSizeAllAsync Fallback 路徑為 GetFolderSizeAsync

在 `GetFolderSizeAllAsync` (透過 `GetFolderSizeAllL3`) 的 OOM fallback 路徑中，目前使用的是 `GetFolderSizeL3` (Layer 3)，這會直接進行 COM 呼叫而不經過快取。為了優化效能，我們將其改為呼叫 `GetFolderSizeAsync` (Layer 2.5)，這樣當子資料夾的大小已經在快取中時，可以大幅提升速度。

## 使用者審查請求

> [!IMPORTANT]
> 此修改將 `GetFolderSizeAllL3` 的底層循環從直接呼叫 L3 改為呼叫 L2.5 (`GetFolderSizeAsync`)。這意味著在計算整棵子樹大小時，每個子資料夾都會先檢查記憶體和資料庫快取。

- **參數傳遞**: 我們將新增 `cToken` 參數至 `GetFolderSizeAsync` 並正確傳遞。
- **路徑優化**: 呼叫 `GetFolderSizeAsync` 時傳入已知的 `fPath`，避免再次讀取 COM 屬性。

## 擬議變更

### Form1_Outlook.vb

#### [MODIFY] [Form1_Outlook.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_Outlook.vb)

- **`GetFolderSizeAsync`**:
  - 新增 `Optional cToken As CancellationToken = Nothing`。
  - 呼叫 `GetFolderSizeL3` 時傳入 `cToken`。
- **`GetFolderSizeAllL3`**:
  - 將循環內的 `Await GetFolderSizeL3(f, progress, cToken)` 替換為 `Await GetFolderSizeAsync(f, cToken:=cToken, fPath:=targetFolderList(i).FolderPath)`。

## 實作步驟

1.  **修改 `GetFolderSizeAsync`**：
    - 更新函式簽名以包含 `cToken`。
    - 在對 `GetFolderSizeL3` 的 `Await` 呼叫中加入 `cToken:=cToken`。
2.  **修改 `GetFolderSizeAllL3`**：
    - 定位至 Line 1586。
    - 替換為調用 Layer 2.5 的 `GetFolderSizeAsync`。

## 驗證計畫

### 自動化測試
- 觀察 Log 中是否出現 `GetFolderSizeAsync` 被 `GetFolderSizeAllL3` 呼叫的跡象。
- 檢查重複計算同一棵樹時，子資料夾是否直接命中快取而不再進入 L3。

### 手動驗證
- 點擊 Tab 2 的資料夾，檢查統計結果是否準確且速度有提升。
- 測試過程中按下「ESC」或取消，確保作業能立刻停止。
