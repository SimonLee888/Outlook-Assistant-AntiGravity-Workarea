# Outlook Assistant 初始化架構重構總結 (Phase 3.1 Pro)

本次任務成功將 `Form1` 的 UI 初始化流程從混亂的交叉調用，重構為明確、穩定的**三階段執行模型**。這解決了動態控制項（如 `SimTree2`） Z-Order 衝突導致的顯示與佈局異常。

## 🏗️ 核心架構：初始化三部曲

在 `InitLookAndFeel` 中，我們嚴格執行以下三個階段：

### 1. 掛載階段 (Mounting)
- **函數**：`InitTab2UI`, `InitTab3UI`, `InitTab4UI`, `InitTab5UI`
- **職責**：執行 `New` 實例化，並將控制項加入容器（`Controls.Add`）。
- **關鍵**：此階段**不設定** `Dock` 或層級屬性，避免 WinForms 提早渲染導致佈局破碎。

### 2. 渲染階段 (Theming)
- **函數**：`InitListViews`, `InitTreeViews`
- **職責**：利用**遞迴搜尋** (`GetAllListViews`, `GetAllTreeViews`) 遍歷全表單，統一套用字型、顏色、雙緩衝與共用事件。
- **改進**：即使是動態產生的 `SimTree2` 也能在此階段被自動納入管理，不需手動維護清單。

### 3. 佈局階段 (Final Layout)
- **函數**：`ExecuteFinalLayout` -> `LayoutTabXUI`
- **職責**：設定所有控制項的 `Dock`、`Height` 以及最重要的 **Z-Order** (`BringToFront` / `SendToBack`)。
- **成就**：確保佈局是在視窗顯示前的「最後定案」，徹底解決了 Chart2 遮疊 ListView 或 Panel 選項消失的問題。

---

## 🛠️ 分頁佈局修復詳情

| 分頁 | 變更說明 | 預期效果 |
| :--- | :--- | :--- |
| **Tab 2** | 分離 `LayoutTab2UI`，重新定義 Chart 與 ListView 的 Dock 順序。 | 統計圖表與列表層級正確，不再發生遮擋。 |
| **Tab 3** | 建立 `pnlOptions_tab3` 並在 Layout 階段執行 `SendToBack`。 | 頂部搜尋選項面板穩定顯示在上方。 |
| **Tab 4** | 將系列郵件按鈕放入 `pnlOptions_tab4`，佈局邏輯移至 `LayoutTab4UI`。 | 同時滿足按鈕置頂與 ListView 填滿。 |
| **Tab 5** | 將重複郵件選項放入 `pnlOptions5`，清理 `InitTab5UI`。 | UI 層級解耦，結構清晰。 |

---

## 📝 註解規範
所有新增或大幅改動的函數均已加上標記：
`' by AntiGravity, 2026/03/29`

> [!IMPORTANT]
> **後續維護建議**
> 若未來需要新增分頁或修改 UI 元素，請務必遵守「只在 Layout 函數設定 Dock/Z-order」的原則。
> `ExecuteFinalLayout` 應該是整個初始化流程中最後一個調整視覺位置的函數。

render_diffs(file:///d:/Users/Simon/Dropbox/私人文件/Visual Studio/Visual Studio 18 (2026)/Outlook Assistant - (AntiGravity測試區)/Form1.vb)
