# 自動化控制項遍歷與樣式套用計劃 (8.0)

本計劃旨在解決 `InitTreeViews` 中手動維護控制項清單（Hard-coded Array）帶來的維護風險與潛在崩潰問題。我們將引入動態遍歷機制，自動掃描 Form1 中的所有 `TreeView` 與 `SimTree`。

## User Review Required

> [!NOTE]
> **變更核心機制**：本改動會將 `InitTreeViews` 從「主動點名」改為「自動掃描」。這意味著如果您在表單深處（如隱藏的 Panel 裡）有任何不希望被套用共用樣式的 TreeView，它們也會被抓到。但我初步盤點過您的專案，目前所有的 TreeView 似乎都應遵循統一外觀，因此這是安全的。

## Proposed Changes

---

### [MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual Studio/Visual Studio 18 (2026)/Outlook Assistant - (AntiGravity測試區)/Form1.vb)

#### [NEW] 遞迴搜尋函數 `GetAllControlsOfType`
- 新增一個通用或專用的遞迴函數，用來找出 `Me.Controls` 中（含子容器）的所有對象。

#### [MODIFY] `InitTreeViews` 邏輯重構
- 移除 `{TreeView1, TreeView3, ...}` 的手動陣列。
- 改為呼叫 `GetAllControlsOfType(Of TreeView)(Me)`。
- 內部邏輯維持不變，但增加對 `SimTree` 的檢查，確保 `BeforeExpand` 事件只掛載給原生 `TreeView`。

---

## Open Questions

- **是否需要排除特定控制項？** 目前全表單除了 `TreeView1, 3, 4, 5` 與 `SimTree1-4` 外，是否有任何 TreeView 是您希望維持原樣、不要被自動上色的？（如果沒有，則此自動化方案最優）。

## Verification Plan

### Automated Tests
- 編譯並啟動程式：確認 `InitTreeViews` 執行時不會報錯（即使 `TreeView2` 已在 Designer 刪除）。

### Manual Verification
- 打開各個 Tab 頁面，確認各處的資料夾樹（不論是原生還是 SimTree）是否正確上色並具有雙緩衝效果。
