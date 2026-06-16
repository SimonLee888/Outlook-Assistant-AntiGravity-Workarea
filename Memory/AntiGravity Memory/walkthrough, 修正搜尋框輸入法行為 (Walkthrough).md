# 修正搜尋框輸入法行為 (Walkthrough)

## 變更摘要
我已修正了 `DebugForm` 中搜尋輸入框 (`txtDebug`) 在點選時會自動切換到中文輸入法的問題。

### 具體修改細節
 在 `Form1_DebugForm.vb` 的 `DebugForm_Shown` 事件中，為 `txtDebug` 新增了 `ImeMode` 屬性設定。

#### [MODIFY] [Form1_DebugForm.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual Studio/Visual Studio 18 (2026)/Outlook Assistant - (AntiGravity測試區)/Form1_DebugForm.vb)

```vb
' 4. 設定左側: txtDebug (搜尋輸入框) ──
txtDebug.ImeMode = ImeMode.Alpha            ' by AntiGravity, 2026/04/07: 強制預設英文/半形英數，解決輸入法自動切換中文問題
txtDebug.Location = New Point(8, targetTop) ' 距左側 8px
```

> [!TIP]
> 使用 `ImeMode.Alpha` 可以確保控制項獲取焦點時，輸入法會預設切換為「半形英數」狀態，這最符合除錯視窗搜尋函數名稱的使用情境。

## 驗證結果
- **程式碼檢視**: 已確認 `ImeMode` 設定語法正確，且緊鄰佈局邏輯，易於維護。
- **標記**: 已依規範加上 `by AntiGravity, 2026/04/07` 註記。
