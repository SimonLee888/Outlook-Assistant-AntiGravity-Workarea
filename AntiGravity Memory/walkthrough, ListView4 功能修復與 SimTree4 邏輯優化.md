# ListView4 功能修復與 SimTree4 邏輯優化

本次更新解決了 Tab4（系列郵件）中開啟郵件功能失效的問題，並整理了 ESC 鍵的導航邏輯與控制項名稱誤植。

## 變動總覽

### 1. 解決 ListView4 開啟功能失效 (Index 錯位)

> [!IMPORTANT]
> 之前的邏輯依賴於 `item.Index` 來從資料源中抓取郵件，但在分組模式下索引會發生錯位，導致抓到錯誤的 EntryID 或報錯。

- **核心修復**: 在 `FillListView4` 中建立列表項時，將原始 `MailItemInfo` 直接賦值給 `lvi.Tag`。
- **受益範圍**: 
    - 滑鼠單擊 (顯示路徑)
    - 滑鼠雙擊 (開啟郵件)
    - Enter 鍵 (批次開啟選中郵件)

### 2. 修正控制項名稱誤植

- 在 `ListView4_KeyPress` 與 `RefreshListView4MailsAsync` 中，將殘留的 `TreeView4` 統一修正為 `SimTree4`，確保各個 Partial Class 對組件的存取同步。

### 3. 解除 SimTree4 ESC 鍵邏輯衝突

- **強化攔截**: 在 `SimTree4_KeyDown` 處理 ESC 鍵時加入 `e.SuppressKeyPress = True`，防止事件繼續傳遞。
- **通用排除**: 在 `Form1.vb` 的通用 `HandleTreeViewKeyPress` 中加入對 `SimTree4` 的排除邏輯，讓 Tab4 的導航模式切換不被全域行為干擾。

## 驗證細節

### 已執行的修正
- [x] 修改 `FillListView4` 注入 `Tag` 資料。
- [x] 修復滑鼠點擊與雙擊邏輯。
- [x] 修復 Enter 鍵鍵盤邏輯。
- [x] 修正 `RefreshListView4MailsAsync` 的資料來源。
- [x] 修復 ESC 鍵在多層事件處理器間的衝突。

---
**by Gemini 3.0 flash, 2026/04/21**
