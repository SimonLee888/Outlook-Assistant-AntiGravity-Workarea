# SimTree4 水平捲軸隱藏與位移防止 (強化版) 實作紀錄

針對先前實作後仍存在的「下方灰條空間」與「點選長項目微到位移」問題，我們已透過更深層的 Win32 API 呼叫進行了強化。

## 變更內容

### [SimTree 控制項](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_SimTree.vb)
- **引入 Win32 API**：
    - 宣告 `ShowScrollBar` (user32.dll)。
    - 新通常數 `SB_HORZ`, `WS_HSCROLL`, `TVM_ENSUREVISIBLE`。
- **徹底移除空間 (灰條)**：
    - **CreateParams**：在樣式中明確移除 `WS_HSCROLL`。
    - **OnHandleCreated**：在控制換代碼建立時，強行呼叫 `ShowScrollBar(IntPtr, SB_HORZ, False)`。
- **解決自動對齊位移**：
    - **WndProc**：除了 `WM_HSCROLL`，增加對 `TVM_ENSUREVISIBLE` 的處理。
    - > [!IMPORTANT]
      > `TVM_ENSUREVISIBLE` 是 TreeView 在點選長項目後，自動嘗試將內容捲入視圖的指令。攔截此訊號後，我們能確保左右偏移量永遠鎖定在 0。

### [Form1 初始化](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1.vb)
- （維持原狀）在 `InitTab4UI` 中啟用屬性。

## 新增功能：ListView4 Ctrl+A 全選

針對使用者在 `ListView4` 中進行批次操作的需求，我們新增了標準的快捷鍵支援：

### [Tab4 列表互動](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_MainTabs.vb)
- **事件處理**：在 `ListView4_KeyDown` 中攔截 `Ctrl+A`。
- **實作方式**：
    - 使用 `BeginUpdate` / `EndUpdate` 保護介面，避免大量選取時造成的重繪閃爍。
    - 遍歷所有 `ListViewItem` 並將其 `Selected` 屬性設為 `True`。
    - 成功執行後使用 `e.SuppressKeyPress = True` 阻止系統產生嗶聲。

## 待修項目 (暫時擱置)
- **SimTree4 水平位移問題**：強化的 Win32 API 方案目前仍未達預期，已記錄於待辦清單，留待後續進一步研究更高層級的 Layout 鎖定方案。
