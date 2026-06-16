# ListView4 自訂延遲 ToolTip 實作計劃 (Ver. 2)

## 背景與問題分析
1. **原生 ToolTip 失敗**：系統預設 ToolTip 彈出速度快且內容不受控（顯示了收件時間），且無法設定為精確的 2 秒。
2. **前次修改未生效**：由於編輯工具傳輸錯誤，`FillListView4` 並未成功寫入路徑資訊。

---

## 預計修改項目

### 1. UI 初始化修正

#### [MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/Form1.vb)
- 將 `ListView4.ShowItemToolTips` 設回 `False` (或移除該設定)，避免原生的干擾。

---

### 2. 資料填充修正

#### [MODIFY] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/Form1_MainTabs.vb)
- 修改 `FillListView4`：確保執行 `lvi.Tag = mailItem.FolderPath`，讓滑鼠懸停時能瞬間取得資料夾路徑。

---

### 3. 2 秒延遲顯示邏輯 (核心更動)

#### [MODIFY] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/Form1_MainTabs.vb)
- **宣告控制項**：
  ```vb
  Private WithEvents _lv4TooltipTimer As New System.Windows.Forms.Timer With {.Interval = 2000}
  Private _lv4LastHoverItem As ListViewItem = Nothing
  Private _lv4Tooltip As New ToolTip()
  ```
- **實作 `ListView4_MouseMove` 事件**：
  - 使用 `HitTest` 或 `GetItemAt` 偵測滑鼠下方的 Item。
  - **如果滑鼠移到新項目**：
    - 停止並重新啟動 Timer (歸零計時)。
    - 隱藏目前的 ToolTip。
    - 紀錄 `_lv4LastHoverItem` 為目前項目。
  - **如果滑鼠移出任何項目**：
    - 停止 Timer 並隱藏 ToolTip。
- **實作 `_lv4TooltipTimer_Tick` 事件**：
  - 此事件會在滑鼠靜止 2 秒後觸發。
  - 從 `_lv4LastHoverItem.Tag` 讀取 `FolderPath`。
  - 在滑鼠位置顯示 `ToolTip`。
  - 停止 Timer (避免重複彈出)。

---

## 驗證計畫

### 1. 內容正確性驗證
- 確保出現的文字不再是時間，而是完整的 `FolderPath`。

### 2. 精確計時驗證
- 滑鼠移入項目，啟動碼表計時，確認是否接近 2 秒才彈出。
- 快速掃過多個項目，確認 ToolTip **不會** 頻繁彈出（應該要停下來才算）。

### 3. 穩定性驗證
- 點選 TreeView4 切换節點後，確認 ToolTip 功能依然正常（LVI.Tag 是否被正確填入）。
