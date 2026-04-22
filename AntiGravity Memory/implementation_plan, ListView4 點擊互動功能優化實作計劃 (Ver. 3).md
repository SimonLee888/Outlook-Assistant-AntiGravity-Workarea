# ListView4 點擊互動功能優化實作計劃 (Ver. 3)

## 背景與目標
之前的「滑鼠懸停 2 秒」方案因 ToolTip 效果不佳且穩定性問題，決定改為更直接的點擊觸發：
- **單擊**：彈出 ToolTip 並在 ProgressBar2 顯示 FolderPath。
- **雙擊**：開啟該郵件 (OpenMailByEntryID)。

---

## 預計修改項目 (精確執行策略)

### 1. 清理類別成員

#### [MODIFY] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/Form1_MainTabs.vb)
- 移除這三行自訂 ToolTip 變數：
    - `_lv4TooltipTimer`
    - `_lv4Tooltip`
    - `_lv4LastHoverItem`

### 2. 資料注入修復

#### [MODIFY] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/Form1_MainTabs.vb)
- 重寫 `FillListView4` 函數，確保：
    - `lvi.Tag = mailItem.FolderPath` 被正確執行。
    - 這樣單擊事件才能即時讀取到路徑資訊。

### 3. 事件處理重構

#### [MODIFY] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/Form1_MainTabs.vb)
- **實作 `ListView4_MouseClick`**：
    - 取得項目 -> 彈出 ToolTip。
    - 更新 `ProgressBar2.Text = "資料夾路徑: " & item.Tag`。
    - 原本的複製主旨功能保留。
- **實作 `ListView4_MouseDoubleClick`**：
    - 取得 EntryID。
    - 呼叫 `OpenMailByEntryID` 開啟郵件。
- **移除廢棄事件**：
    - 移除 `ListView4_MouseMove`
    - 移除 `ListView4_MouseLeave`
    - 移除 `_lv4TooltipTimer_Tick`

---

## 驗證計畫

### 1. 單擊測試
- 點選 ListView4 的任何一封郵件。
- 確認游標位置彈出路徑 ToolTip。
- 確認狀態列顯示 `資料夾路徑: [實際路徑]`。

### 2. 雙擊測試
- 雙擊系列郵件。
- 確認 Outlook 能夠成功跳出該封郵件的詳細視窗。
