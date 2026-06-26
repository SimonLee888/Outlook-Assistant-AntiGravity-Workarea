# 修正背景執行緒 _dbg 訊息未顯示在 DebugForm 的計劃

在 Tab5 Fuzzy 模式中，部分 SimHash 和 Jaccard 計算是透過 `Task.Run` 在背景執行緒中執行。
在 VB.NET 中，直接使用類別名稱 `DebugForm` 來存取表單方法（例如 `DebugForm.AddMessage3`）會使用編譯器生成的**預設實例 (Default Instance)**，而該實例在 WinForms 下是 **Thread-Local (執行緒獨立)** 的。
因此，當背景執行緒呼叫 `_dbg` 時，會隱式建立另一個全新且看不見的 `DebugForm` 實例，導致訊息只寫入該背景實例，而無法在主 UI 顯示。

## 解決方案

在 `DebugForm` 類別中導入一個全域共享的執行個體變數 `ActiveInstance`，在主 UI 執行緒的 `DebugForm` 載入時指向自己，並在關閉時釋放。隨後修改 `Form1` 的 `_dbg`，使其優先將訊息傳入 `DebugForm.ActiveInstance`。

## Proposed Changes

---

### DebugForm 模組

#### [MODIFY] [Class_DebugForm.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Class_DebugForm.vb)
- 新增 `Public Shared ActiveInstance As DebugForm = Nothing` 共享變數。
- 在 `DebugForm_Load` 中設定 `ActiveInstance = Me`。
- 在 `DebugForm_FormClosed` 中清除 `ActiveInstance = Nothing`。

---

### Form1 主畫面模組

#### [MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1.vb)
- 修改 `_dbg` 方法，在 `_isDebugMode` 通過後，優先呼叫 `DebugForm.ActiveInstance.AddMessage3`（如果存在），否則才 fallback 走預設實例。

## Verification Plan

### Manual Verification
- 啟動程式，開啟 Debug 視窗。
- 切換至 Tab5，進入 Fuzzy 模式執行掃描重複郵件。
- 確認 Debug 視窗是否正確顯示 `[探針:S4候選]`、`[探針:S5閘門]` 等背景執行緒所產生的訊息。
