# ListView4 滑鼠停頓顯示資料夾路徑功能實作計劃

## 背景與目標
使用者希望在 Tab4 的郵件列表 (ListView4) 中，當滑鼠停留在某個項目上超過 **2 秒**時，以 ToolTip 顯示該郵件所屬的 `FolderPath`。

---

## 預計修改項目

### 1. 資料結構擴充 (MAPI 層)
為了解耦和加速，我們需要在內存資料中攜帶路徑資訊。

#### [MODIFY] [Form1_Outlook.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/Form1_Outlook.vb)
- 在 `MailItemInfo` 結構中新增 `FolderPath As String` 欄位。
- 在 `GetFolderBasicMailInfosL3` 函數中，讀取 Table 時直接將該次調用的 `folder.FolderPath` 存入每一個 `MailItemInfo` 物件。

---

### 2. UI 資料綁定優化 (UI 層)

#### [MODIFY] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/Form1_MainTabs.vb)
- 修改 `FillListView4`：在建立 `ListViewItem` 時，將 `mailItem.FolderPath` 存入 `lvi.Tag`，以便滑鼠懸停時能瞬間讀取。

---

### 3. 延遲 2 秒顯示邏輯 (互動層)

為了達成精確的 2 秒延遲，我們將使用一個單獨的 Timer 來控制。

#### [NEW] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/Form1_MainTabs.vb)
- **宣告變數**：
    - `Private WithEvents _tooltipLv4 As ToolTip`
    - `Private WithEvents _timerLv4Tooltip As System.Windows.Forms.Timer`
    - `Private _lastHoveredLv4Item As ListViewItem`
- **初始化** (在 `InitTab4UI` 中)：
    - 初始化 `_tooltipLv4`。
    - 初始化 `_timerLv4Tooltip` (設為 2000ms, Enabled=False)。
- **實作事件**：
    - `ListView4_MouseMove`: 偵測滑鼠位置。若進入新項目，則啟動 Timer；若移出項目，則重置 Timer 並隱藏 ToolTip。
    - `_timerLv4Tooltip_Tick`: 當 Timer 到達 2 秒且滑鼠仍停留在同一項目。
        - 從項目的 `Tag` 讀取 `FolderPath`。
        - 呼叫 `_tooltipLv4.Show()`。

---

## 實作細節與考量

> [!IMPORTANT]
> **為何不直接用 ListView.ShowItemToolTips?**
> Windows 原生的 ListView ToolTip 延遲時間通常由系統設定 (約 500ms)，且不易針對單一控制項修改為「特定時間 (如 2s)」。使用專屬 Timer 可提供精確控制。

> [!NOTE]
> **效能影響**：`MailItemInfo` 增加一個 String 欄位對記憶體消耗極低，且由於 `folder.FolderPath` 已預先讀取，此舉不會增加額外的 COM 往返。

---

## 驗證計畫

### 1. 資料驗證
- 中斷點檢查 `GetFolderBasicMailInfosL3` 回傳的 List 中，`FolderPath` 是否皆有正確填充。
- 檢查 `ListView4.Items(0).Tag` 是否包含正確路徑。

### 2. UI 互動驗證
- 滑鼠移入 ListView4 項目，心中默數 2 秒 -> 確認 ToolTip 出現。
- 滑鼠快速在項目間移動 -> 確認 ToolTip **不會** 出現（Timer 持續被重置）。
- 滑鼠在某一項目停頓出現 ToolTip 後，移開滑鼠 -> 確認 ToolTip 立刻消失。
