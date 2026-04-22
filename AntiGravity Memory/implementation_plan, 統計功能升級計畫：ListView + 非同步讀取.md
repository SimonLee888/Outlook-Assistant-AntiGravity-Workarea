# 統計功能升級計畫：ListView + 非同步讀取

本計畫將解決 Setting 頁面切換卡頓的問題，並透過更換 UI 元件（TextBox -> ListView）來達成 `Noto Sans TC` 字型下的完美對齊。

## Proposed Changes

### [Component] UI 元件升級 (TextBox 至 ListView)

#### [MODIFY] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/Form1_MainTabs.vb)
- **動態建立 ListView**：在 `RefreshDatabaseStats` 中偵測是否已存在 `ListView`，若不存在則動態建立 `lvDatabaseStats`。
- **屬性繼承**：`lvDatabaseStats` 將繼承 `txtDatabaseStats` 的位置、大小與父容器層級。
- **配置 ListView**：
    - 字型設為 `Noto Sans TC`。
    - `View = Details`, `FullRowSelect = True`, `HeaderStyle = None`。
    - 建立兩欄：[項目] (左對齊), [數值] (右對齊)。
- **隱藏舊組件**：將原本的 `txtDatabaseStats` 設為不可見。

### [Component] 非同步讀取重構

#### [MODIFY] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/Form1_MainTabs.vb)
- 將 `RefreshDatabaseStats` 修改為 `Async Sub`。
- 使用 `Await Task.Run(...)` 包裝資料庫讀取項目（`GetDatabaseSummary`）。
- 使用 `ListView.Items.BeginUpdate()` 與 `EndUpdate()` 確保介面更新流暢。

#### [MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/Form1.vb)
- 修改 `TabControl1_SelectedIndexChanged` 內部的呼叫，改為使用非同步方式觸發刷新。

## Verification Plan

### 自動與手動測試
- 切換至 Setting 頁面，確認是否還會卡頓 1-2 秒（預期應為零延遲切換，隨後數據異步載入）。
- 確認 `ListView` 的數值是否靠右對齊。
- 確認字型是否已成功套用 `Noto Sans TC` 且外觀優美。
