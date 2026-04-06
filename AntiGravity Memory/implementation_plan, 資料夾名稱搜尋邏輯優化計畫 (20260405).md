# 資料夾名稱搜尋邏輯優化計畫 (2026/04/05)

此計畫旨在重構 `Form1_ComL3.vb` 中的名稱搜尋方法，解決 `Replace(" - ", "")` 重複運算導致的效能浪費及程式碼冗長問題。

## 擬議變更

### 1. 字串清理輔助方法 [Logic]

#### [NEW] [Form1_ComL3.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_ComL3.vb)

建立 `GetCleanText` 輔助方法：
- 統一移除 「 - 」 前綴與後綴。
- 統一執行 `Trim()`。
- (可選) 處理 Outlook 資料夾名稱中常見的特殊字元。

### 2. 優化搜尋函數 `FindNodeByName` [Logic]

#### [MODIFY] [Form1_ComL3.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_ComL3.vb)

- **移除重複 Replace**：
  在 `FindNodeByName` 的進入點（或是呼叫它的 `GetFolderByName`）就先洗好 `cleanTargetName`。
- **遞迴優化**：
  傳遞 `cleanTargetName` 進入遞迴，迴圈內僅對 `node.Text` 執行一次清理比對。
- **Case-Insensitive**：
  使用 `String.Equals(..., StringComparison.OrdinalIgnoreCase)` 提高穩定性。

### 3. 優化 `FindNodeOrItemByName` [Logic]

#### [MODIFY] [Form1_ComL3.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_ComL3.vb)

- 同樣採取「先洗好目標名稱」策略，讓迴圈內容從 `2N` 次 Replace 降為 `N` 次。

## 開放性問題 (Open Questions)

- **為什麼使用 " - " 作為取代目標？**
  目前程式碼中有 `.Replace(" - ", "")`。這看起來是為了處理 UI 上模擬縮排的空白與連字號。是否有其他格式（例如只有空格或是不同符號）也需要一併清理？我目前建議使用一個更通用的 `CleanNodeText` 方法來統一這類規則。

## 驗證計畫

### 手動測試 (Manual Verification)
1.  **測試導航**：在 Tab1 ListView 雙擊任何資料夾，確認仍能正確跳轉至 TreeView 對應位置（這會觸發 `FindNodeByName`）。
2.  **測試右鍵選單**：在 ListView 右鍵點選「統計此資料夾」，確認後端仍能透過名稱識別出正確的 Folder 物件。
3.  **效能對比**：在大型資料夾結構（例如有 200 個子資料夾）下執行搜尋，觀察 `Dbg()` 記錄的時間戳記是否有微幅進步。
