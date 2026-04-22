# 快取統計顯示優化記錄

我已經優化了 Setting 頁中快取統計數據的顯示格式，解決了數字排列不對齊的問題。

## 變更摘要

### [Form1.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/Form1.vb)

1.  **強制使用等寬字型**：
    在 `RefreshDatabaseStats` 中加入了判斷，確保 `txtDatabaseStats` 控制項使用 `Consolas` 字型。這是實現精確對齊的基礎。
2.  **優化字串格式化邏輯**：
    -   **標籤左對齊**：使用字串插補格式 `{label, -N}` 將標籤固定在左側。
    -   **數字右對齊**：統一使用 `{value, 8:N0}`，確保所有數字位元數對齊且包含千分位。
    -   **中英文混排處理**：針對「DB 檔案大小」等含中文字的標籤，特別調整了寬度係數，使其在等寬字型下與英文標籤視覺對齊。

## 驗證結果
- [x] 設定頁面字型已切換為 Consolas。
- [x] SQLite 快取統計區塊的冒號與數字垂直對齊。
- [x] Memory 快取統計區塊（標籤較長者）亦達成完全對齊。

> [!TIP]
> 之後若有新增統計欄位，只要沿用 `{"Name",-22} : {value,8:N0}` 的格式即可保持整齊。
