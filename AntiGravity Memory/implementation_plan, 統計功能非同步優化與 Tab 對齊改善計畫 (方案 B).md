# 統計功能非同步優化與 Tab 對齊改善計畫 (方案 B)

## Proposed Changes

### [Component] 統計功能非同步化

#### [MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/Form1.vb)
- 在 `TabControl1_SelectedIndexChanged` 中，將 `RefreshDatabaseStats()` 的呼叫改為 `Async` 執行路徑。

#### [MODIFY] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/Form1_MainTabs.vb)
- 重構 `RefreshDatabaseStats` 為 `Async Sub`。
- 使用 `Await Task.Run(...)` 非同步取得資料庫摘要。
- 顯示載入中狀態。

### [Component] UI 字型與 Tab 對齊

#### [MODIFY] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/Form1_MainTabs.vb)
- 設定字型為 `Noto Sans TC`。
- 移除原本的 `PadRight` 與手動空格計算。
- 改用 `\t` 字元，嘗試在比例字型下達成對齊。


---

## Open Questions

> [!WARNING]
> **您是否同意將 `txtDatabaseStats` 改為 `ListView`？**
> 如果同意，我會一併處理元件更正。如果不同意，使用 `Noto Sans TC` 時數字將無法整齊排列。

## Verification Plan

### 自動與手動測試
- 切換到 Setting 頁面，觀察是否還會卡頓 1-2 秒。
- 確認在統計資料未出爐前，畫面顯示「讀取中」提示。
- 確認數據載入後，`Noto Sans TC` 顯示效果美觀。
