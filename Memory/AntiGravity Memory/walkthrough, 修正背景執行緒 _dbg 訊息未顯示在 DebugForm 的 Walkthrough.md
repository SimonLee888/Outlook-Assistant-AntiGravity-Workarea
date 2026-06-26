# 修正背景執行緒 _dbg 訊息未顯示在 DebugForm 的 Walkthrough

我們已成功修改程式碼，解決了背景執行緒的除錯訊息無法在 `DebugForm` 顯示的問題。

## 修改內容

### 1. [Class_DebugForm.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Class_DebugForm.vb)
* 新增共享的 `ActiveInstance` 變數：
  ```vb
  Public Shared ActiveInstance As DebugForm = Nothing
  ```
* 於 `DebugForm_Load` 中將目前的 `DebugForm` 實例指派給 `ActiveInstance`。
* 於 `DebugForm_FormClosed` 中將 `ActiveInstance` 設回 `Nothing`。
* 於 `AddMessage3` 的說明區塊加註了重要的跨執行緒安全防範警語，要求未來修補此方法時必須確保其維持 `enqueue-only` 的非同步寫入特性，嚴禁在此方法中直接存取控制項。

### 2. [Form1.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1.vb)
* 修改 `_dbg` 方法中的路由邏輯。如果 `DebugForm.ActiveInstance` 不為 `Nothing`，優先透過其呼叫 `AddMessage3`，否則才走原本的預設實例呼叫：
  ```vb
  If DebugForm.ActiveInstance IsNot Nothing Then
      DebugForm.ActiveInstance.AddMessage3(msg, detail, realCaller)
  Else
      DebugForm.AddMessage3(msg, detail, realCaller)
  End If
  ```

## 驗證與結果

* **程式碼複檢**：我們已使用 `view_file` 仔細確認過所有修改行數，邏輯對齊，並無遺留多餘或損毀的程式碼。
* **背景執行緒安全性**：因為 `AddMessage3` 維持其 `ConcurrentQueue` 的寫入佇列架構，我們確保了背景執行緒不會觸發 UI 的跨執行緒存取崩潰，且能夠將訊息安全地遞交給主 UI 執行緒的 `QueueTimer` 進行批次渲染。
