# ListView4 點擊互動功能優化計劃

## 背景與目標
之前的「滑鼠懸停 2 秒」方案經使用者測試效果不佳。
為了提供更明確直覺的互動，我們決定改為：
- **單擊項目**：直接彈出 ToolTip 並在狀態列顯示路徑。
- **雙擊項目**：直接開啟 Outlook 郵件視窗（比照 Tab3 的行為）。

---

## 預計修改項目

### 1. 清除舊有懸停邏輯 (互動層)

#### [MODIFY] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/Form1_MainTabs.vb)
- 移除 `_lv4TooltipTimer` (Timer), `_lv4Tooltip` (ToolTip物件), `_lv4LastHoverItem` (變數)。
- 移除 `ListView4_MouseMove`, `ListView4_MouseLeave`, `_lv4TooltipTimer_Tick` 等事件。

---

### 2. 實作點擊顯示功能 (單擊)

#### [MODIFY] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/Form1_MainTabs.vb)
- **強化 FillListView4**：確保將 `mailItem.FolderPath` 存入 `lvi.Tag`。
- **修改 ListView4_MouseClick**：
    - 偵測滑鼠點選的項目。
    - 讀取 `lvi.Tag`。
    - 在滑鼠位置呼叫私有的 `ToolTip.Show` (或建立一個單次使用的 ToolTip)。
    - 設定 `ProgressBar2.Text = "資料夾路徑: " & folderPath`。

---

### 3. 實作開啟郵件功能 (雙擊)

#### [MODIFY] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/Form1_MainTabs.vb)
- **修改 ListView4_MouseDoubleClick**：
    - 取得選取項目的 EntryID（位於第 5 欄 SubItems(4)）。
    - 呼叫專案現成的 `OpenMailByEntryID(New List(Of String) From {eid})` 函數，實現在 Outlook 中開啟該郵件。

---

## 實作細節與考量

> [!NOTE]
> **為何不延續複製主旨功能？**
> 使用者目前的指示是雙擊開啟郵件。為了維持一致性，單擊顯示路徑更適合 Tab4「系列郵件辨識」的使用場景，方便確認郵件到底被收到了哪個資料夾。

> [!IMPORTANT]
> **修復警告**：我會先確認 `Form1_MainTabs.vb` 在前次編輯中是否有誤刪除代碼的情況，並在本次修改中一併修復至正確狀態。

---

## 驗證計畫

### 1. 單擊驗證
- 滑鼠單擊項目。
- 確認滑鼠位置出現 ToolTip 顯示路徑。
- 確認下方 ProgressBar2.Text 顯示相同路徑。

### 2. 雙擊驗證
- 滑鼠雙擊項目。
- 確認 Outlook 彈出該郵件內容視窗。
- 確認 EntryID 傳遞正確，不會開啟錯誤的郵件。
