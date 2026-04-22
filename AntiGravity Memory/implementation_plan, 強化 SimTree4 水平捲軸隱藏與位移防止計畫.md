# 強化 SimTree4 水平捲軸隱藏與位移防止計畫

目前的實作雖然禁用了功能，但捲軸空間（灰條）依然存在，且長項目仍有微小位移。本計畫將採用更底層的 Win32 呼叫來達成「完全消失」與「強制零偏移」。

## 使用者評論請求
> [!IMPORTANT]
> 為了徹底解決問題，我們將引入 `ShowScrollBar` API。這會強行移除水平捲軸佔用的 UI 空間（灰條）。

## 擬議變更

### [Component] SimTree 控制項 (Form1_SimTree.vb)

#### [MODIFY] [Form1_SimTree.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_SimTree.vb)

1.  **引入更多 Win32 常數與 API 宣告**：
    *   定義 `SB_HORZ = 0`
    *   宣告 `ShowScrollBar(hWnd As IntPtr, wBar As Integer, bShow As Boolean) As Integer`
    *   定義 `TVM_ENSUREVISIBLE = &H1114` (即 `TVM_FIRST + 20`)
2.  **覆寫 `OnHandleCreated`**：
    *   在控制項建立控制代碼時，立即呼叫 `ShowScrollBar(Me.Handle, SB_HORZ, False)`，確保捲軸不顯示。
3.  **強化 `WndProc` 攔截**：
    *   除了 `WM_HSCROLL`，增加攔截 `TVM_ENSUREVISIBLE`。
    *   許多時候 TreeView 自動捲動是因為 `EnsureVisible` 被觸發，攔截並自訂處理可防止水平方向的自動位移。

## 驗證計畫

### 手動驗證
1.  啟動程式，導航至 Tab4。
2.  確認 `SimTree4` 下方完全沒有任何捲軸空間（移除灰條）。
3.  點選內容極長的「(珍貴笑話篇)...」項目。
4.  確認文字**完全不會**向右偏移（保持左右對齊）。
