# OST CopyFolder 進度顯示優化 (速度與 ETA)

本計畫旨在優化 Tab7 在執行「複製資料夾」功能時的進度回饋。我們將在 `CopyFolder_Click` 執行期間，將原有的文字進度訊息增強為包含「每秒處理筆數」與「預估剩餘時間」的專業格式。

## 使用者審閱確認

> [!IMPORTANT]
> - **僅針對 CopyFolder 觸發**：解析 OST (Load) 時維持原有的簡約顯示。
> - **動態委派切換**：在複製開始時切換進度顯示邏輯，並在結束後（無論成功或失敗）還原為原本的簡約版本。
> - **計算邏輯**：使用正規表示式從核心庫訊息 `"Converting ... X out of Y"` 中提取數字，搭配 `Stopwatch` 計算秒速。

## 擬議變更

### 10 Tab7: OST/PST 解析 (UI 與流程層)

#### [MODIFY] [Form1_OST.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_OST.vb)

- **[ADD] 宣告區**：
  - 加入 `Private _tab7StatusSw As New Stopwatch()`。

- **[MODIFY] `CopyFolder_Click`**：
  - 在進入 `Try` 區塊前，啟動 `_tab7StatusSw.Restart()`。
  - 重新定義 `ost2pst.FM.StatusMsg`：
    - 使用 Regex `(\d+)\s+out\s+of\s+(\d+)` 解析當前進度與總筆數。
    - 仿照 Tab3 邏輯計算 `speed` (筆/秒) 與 `etaString` (剩餘 mm:ss)。
    - 更新 `ProgressBar2.Text` 為增強格式。
  - 在 `Finally` 區塊中：
    - 停止計時器。
    - 將 `ost2pst.FM.StatusMsg` 還原為原本簡單顯示訊息的版本，確保不影響其他功能的 UI 表現。

---

## 驗證計畫

### 自動化測試 (透過瀏覽器或單元測試)
- 無（主要為 UI 顯示邏輯）。

### 手動驗證
1. **Load OST 驗證**：點擊「解析 OST」，確認 `PB2` 顯示原始訊息（如：正在開啟...），無速度資訊。
2. **Copy Folder 驗證**：執行複製操作，確認 `PB2` 跳動顯示例如 `Converting ... 4500/7613 (120 筆/秒，預估剩餘 00:25)`。
3. **還原驗證**：複製結束後再次選取資料夾觸發載入，確認 `PB2` 回復為簡單文字顯示。
