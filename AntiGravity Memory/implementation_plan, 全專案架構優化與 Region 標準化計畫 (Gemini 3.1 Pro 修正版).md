# 全專案架構優化與 Region 標準化計畫 (Gemini 3.1 Pro 修正版)

收到您的回饋，我完全同意您的觀點：先前的方案過於武斷。`ComputeFolderStatsAsync` 既然與 Tab1 的邏輯和 UI 更新緊密綁定，強行移到 `Form1_Outlook.vb` 只會讓依賴關係變得混亂，我們應該尊重它目前在 `Form1_MainTabs.vb` 的歸屬。

另外，您提到的 `Form1.vb` 中的 `#Region "■ 03 共用控制項行為"` 確實已經變成了一個「大雜燴」。它裡面包含了全域狀態切換、快取按鈕事件、控制項樣式處理、甚至是複雜的樹狀節點搜尋邏輯，這正是它「很難分類」的原因——因為標題已經無法涵蓋其內容。

以下是我為您重新擬定的架構整理計畫：

## User Review Required

> [!IMPORTANT]
> 請檢視以下針對 `Form1.vb` 中「■ 03 共用控制項行為」的拆解方案。我計畫將其拆分為四個更具體、獨立的 Region。您覺得這樣的分類是否合理？

## 擬定變更

---

### [Component] Form1.vb (主視窗與全域共用邏輯)

#### [MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1.vb)
廢除現有的 `#Region "■ 03 共用控制項行為"`，將其內部龐大的函數群拆解重組為以下四個明確的 Region：

1. **`#Region "■ 03 全域狀態與快取管理"`**
   - **內容**：放置控制全域行為的 CheckBox 與快取維護按鈕。
   - **包含函數**：`CheckShowAllFolders_CheckedChanged`, `CheckRDO_CheckedChanged`, `SaveCache_Click`, `LoadCache_Click`, `ClearCache_Click`, `RenewCache_Click`, `ClearMemoryCachesInner`。
   - **理由**：這些是應用程式層級的狀態與生命週期管理，與「UI 控制項樣式/行為」無關。

2. **`#Region "■ 04 導覽與節點搜尋 (Navigation & Search)"`**
   - **內容**：放置所有用來找節點、切換分頁、載入特定路徑的輔助函數。
   - **包含函數**：`TabControl1_SelectedIndexChanged`, `ExpandTvToDefaultInbox`, `GetActiveTreeView`, `GetAllTreeViews`, `TriggerTvAfterSelect`, `RefreshAllTreeViews`, `FindNodeByFolderPath`, `GetSelectedFolderPath`, `SelectNodeByPath`, `SelectNodeByPathRecursive`, `FindLvItemByName`。
   - **理由**：這些函數的本質是「資料夾的導航」，把它們獨立出來能讓 UI 邏輯更乾淨。

3. **`#Region "■ 05 共用控制項事件與樣式 (UI Events & Styling)"`**
   - **內容**：真正屬於滑鼠、鍵盤、畫面重繪的共用事件處理。
   - **包含函數**：`HandleSplitterMouseDown`, `HandleTvMouseHover`, `HandleLvMouseHover`, `HandleTvKeyPress`, `HandleLvGotFocus`, `HandleLvResize`, `GetHeaderRowBackColor`, `AutoResizeLvColumns`, 以及 `HistoryListBox` 的專屬事件 (`MouseMove`, `MouseLeave`, `SelectedIndexChanged`, `DrawItem`)。
   - **理由**：名副其實的「控制項行為」。

4. **`#Region "■ 06 系統與非同步工具"`** (可選，視乎 `Form1.vb` 目前的狀況)
   - **內容**：放置 `OkayNowYouHaveToken`, `ThrottledYieldAsync`, `SyncDebugFormPosition` 等底層或偵錯輔助函數。

---

### [Component] Form1_Outlook.vb (資料與快取層)

#### [MODIFY] [Form1_Outlook.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_Outlook.vb)
您的手動調整已經將 `Layer 2.5` 等概念帶入註解中，但最外層的標題仍是 `■ 10 底層 COM 函數群 (新設計，現役主力)`。
- **變更**：將這個大標題依據您目前的子區域，正式拆分為獨立的 Region：
  - `#Region "■ 10 初始化與生命週期"` (包含 `InitOutlookNamespace`, `InitRdoSession` 等)
  - `#Region "■ 20 Layer 2.5 快取代理層 (Cache Proxy)"`
  - `#Region "■ 30 Layer 3 底層數據層 (COM/RDO)"`
- 此外，將 `FillFolderCacheFromDbRow` 從最下方的雜項輔助函數區，提上來放進 `Layer 2.5` 區塊中，因為它是快取機制不可分割的一部分。

---

### [Component] Form1_MainTabs.vb (Tab1~5 分頁邏輯)

#### [MODIFY] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_MainTabs.vb)
- 維持 `ComputeFolderStatsAsync` 及其他分頁專屬計算函數的現狀。
- 僅進行巡檢，確認 `#Region` 的命名是否符合 `■ 10 Tab 1: [名稱]` 的規範，確保整個檔案折疊後的層次清晰。

---

## 驗證計畫
1.  **結構複檢**：完成修改後，我會使用 `view_file` 工具重新折疊與展開這幾個檔案，確保所有的函數都已正確歸位，且沒有遺漏。
2.  **編譯與語法確認**：以 `Task` 搭配小塊修改，每次搬移函數後都會複檢頭尾的 `End Sub` / `End Function`，確保不會產生語法錯誤。
