# 修復 ListView2 群組縮合與連動功能計畫

根據使用者回報與程式碼分析，目前 ListView2 的群組收合功能無效，且點擊群組標題時的連動邏輯干擾了原生的收合行為。

## 使用者回報問題
1. **群組無法縮合**：雖然調用了 `SetGroupCollapsible`，但群組仍無法點擊收合。
2. **圖表連動失效**：點擊群組標題時，圖表應同步切換至該年份的月份分佈，但目前被 `MouseDown` 邏輯攔接且行為異常。

## 核心原因分析
1. **Win32 結構大小問題**：`LVGROUP` 結構在不同 Windows 版本（Vista 之後）有不同的 `cbSize` 要求。使用 `Marshal.SizeOf` 有時會因對齊或欄位定義導致 Win32 API 拒絕更新狀態（回傳 -1）。
2. **事件攔截衝突**：`ListView2_MouseDown` 中試圖透過 `HitTest` 且在 `hit.Item Is Nothing` 時執行連動，這會攔接到群組標題的點擊事件，導致觸發後不執行原生的收合切換。
3. **缺乏 GroupClick 處理**：真正的標題點擊連動應該在 `GroupClick` 事件中處理，而不是在 `MouseDown` 裡辛苦計算距離。

## 預計修改內容

### [Form1_Win32API.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_Win32API.vb) [MODIFY]
- 補齊 `LVGS_COLLAPSIBLE` (&H8) 等相關常數。
- 檢查 `LVGROUP` 結構宣告與系統版本相容性。

### [Form1.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1.vb) [MODIFY]
- **修正 `SetGroupCollapsible`**：將 `cbSize` 改為 Windows 10/11 通用的擴充大小（補齊所有 Vista+ 欄位）。
- **改進 Win32 調用**：確保 `LVM_SETGROUPINFO` 的 `wParam` 傳入正確的群組 ID。

### [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_MainTabs.vb) [MODIFY]
- **刪除 `ListView2_MouseDown` 冗餘邏輯**：移除手動計算距離以推斷群組的代碼，優先釋放由 WinForms 處理的點擊訊息。
- **新增 `ListView2_GroupClick` 事件處理**：
    - 正式掛載事件以處理標題點擊連動。
    - 取得點擊年份後更新 `Chart2` 月份分佈。

## 驗證計畫
### 手動測試路徑
1. **群組收合測試**：確認點擊年份標題左側小箭頭可正確縮合/展開。
2. **圖表連動測試**：確認點擊年份標題文字區域可同步更新圖表，且不影響縮合狀態。
