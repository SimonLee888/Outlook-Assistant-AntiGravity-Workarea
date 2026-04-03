# 現代化配色系統整合計畫 (v3 - 修復與優化)

本計畫旨在優化目前的配色系統，解決 Tab1 閃爍問題、還原 Legacy 模式的原始外觀，並新增清爽與暗黑兩套新配色。

## User Review Required

> [!IMPORTANT]
> 1. **新增主題**: 我將新增「清爽現代 (FreshModern)」與「暗黑色系 (MidnightDark)」兩套主題。
> 2. **深度修復閃爍**: 
>    - 導入 `SetWindowTheme(handle, "explorer", Nothing)`：啟用 Windows 原生平滑渲染，這是解決 TreeView/ListView 懸停閃爍的標竿做法。
>    - 優化 `MouseMove` 邏輯：減少重複設定 `BackColor`，只有在顏色真正需要改變時才觸發，避免冗餘重繪。
>    - 視窗級別雙緩衝：評估在 `ApplyThemeColors` 中啟用 `WS_EX_COMPOSITED`。
> 3. **還原 Legacy 外觀**: 針對用戶反映的「形狀不對」，我會確保在 Legacy 模式下，`TabControl` 回復為 `Normal` 繪製模式，`Button` 回復為 `Standard/System` 樣式。

## Proposed Changes

### [主程式介面與配色邏輯]
#### [MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1.vb)
- **[MODIFY]** `Enum AppTheme`: 加入 `FreshModern`, `MidnightDark`。
- **[MODIFY]** `SetupTheme_Legacy()`: 修改參數，將 `isLegacy` 旗標傳入 `ApplyThemeColors` 以便還原原生控件樣式。
- **[NEW]** `SetupTheme_FreshModern()`, `SetupTheme_MidnightDark()`: 根據現代美學設計。
- **[MODIFY]** `ApplyThemeColors(...)`:
  - 增加一個 `isLegacy` 參數。
  - 當 `isLegacy = True` 時，`TabControl1.DrawMode = Normal`, `Button.FlatStyle = System`。
  - 當 `isLegacy = False` 時，才啟用 `OwnerDrawFixed` 繪製與 `Flat` 樣式。
- **[MODIFY]** `SwitchTheme(theme As AppTheme)`:
  - 加入 `SendMessage(WM_SETREDRAW, 0)` 與 `1` 來包夾切換邏輯，減少整體視窗閃爍。
- **[MODIFY]** `TabControl1_DrawItem`: 優化填色與邊框邏輯，解決「形狀異常」的問題。

## Verification Plan

### Automated Tests
- 無（主要涉及 UI 視覺呈現）。

### Manual Verification
1. **表格外觀檢查**: 切換至 Legacy 模式，確認 `TabControl` 的頁籤形狀是否回歸系統預設。
2. **閃爍測試**: 在 Tab1 (資料夾統計) 頁面下瘋狂切換主題，確認 `TreeView1` 是否不再劇烈閃爍。
3. **新主題預覽**:
   - 清爽現代: 檢查是否呈現淺藍/灰白的現代感。
   - 暗黑色系: 檢查是否呈現深灰/黑底且文字清晰。
4. **啟動預設**: 關閉程式重新開啟，確認第一眼看到的是否為您原本熟悉的 Classic 外觀與形狀。
