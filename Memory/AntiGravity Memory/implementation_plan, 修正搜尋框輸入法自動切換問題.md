# 修正搜尋框輸入法自動切換問題

目前使用者點選 `DebugForm` 的搜尋框 (`txtDebug`) 時，系統會自動切換至中文輸入法。這是由於 WinForms 控制項預設的 `ImeMode` 行為所致。

## 提出變更

### DebugForm (Form1_DebugForm.vb)

在 `DebugForm_Shown` 事件中，明確設定 `txtDebug.ImeMode` 屬性。

#### [MODIFY] [Form1_DebugForm.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual Studio/Visual Studio 18 (2026)/Outlook Assistant - (AntiGravity測試區)/Form1_DebugForm.vb)

- 在 `txtDebug` 的佈局設定區塊加入 `ImeMode` 設定。
- 建議設定為 `ImeMode.Alpha` 或 `ImeMode.Off`。
    - `ImeMode.Alpha`: 強制為半型英數。
    - `ImeMode.Off`: 預設為英文，但允許使用者手動切更。
- 這裡採用 `ImeMode.Alpha` 以符合使用者「保持英文」的需求。

## 驗證計畫

### 手動驗證
1. 開啟 Outlook Assistant。
2. 開啟除錯視窗 (`DebugForm`)。
3. 點選 `txtDebug` 搜尋框，確認輸入法是否維持在英文狀態。
