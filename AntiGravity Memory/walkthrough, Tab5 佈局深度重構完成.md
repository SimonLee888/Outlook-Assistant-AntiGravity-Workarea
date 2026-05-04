# Tab5 佈局深度重構完成

已解決 `SplitContainer5` 未正常顯示以及 `ListView5` 侵佔空間的問題。

## 變更摘要

### 1. 邏輯重構 (InitTab5UI)
- **容器化管理**: 將 Tab5 的所有控制項改為動態組裝，確保它們被放入 `SplitContainer5` 的正確面板中。
- **佈局對齊**: 
    - `Panel1`: 僅放置 `SimTree5` (取代 `TreeView5`)，設為 `Dock.Fill`。
    - `Panel2`: 頂部放置 `pnlOptions5` (含搜尋模式 RadioButtons 與按鈕)，其餘空間由 `ListView5` 填滿。
- **分割線修復**: 在程式碼中顯式設定 `Panel2Collapsed = False` 並指定 `SplitterDistance = 317`。

### 2. Designer 清理
- 移除 Designer 中衝突的控制項掛載動作，確保 `TabPage5` 只作為 `SplitContainer5` 的容器，具體排列交由 `InitTab5UI` 決定。
- 隱藏舊有的 `TreeView5`。

## 驗證點
- `SplitContainer5` 是否出現中間分割線。
- `SimTree5` 是否顯示在左側。
- `ListView5` 是否正確待在右側且上方留有選項空間。
- 複檢所有修改點確認正確、複檢修改點前後是否遺留多餘程式碼。

by Gemini 3 Flash, 2026/05/03
