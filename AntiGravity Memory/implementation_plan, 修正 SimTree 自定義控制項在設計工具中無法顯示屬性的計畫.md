# 修正 SimTree 自定義控制項在設計工具中無法顯示屬性的計畫

目前 `Form1_SimTree` 無法在設計工具中顯示屬性（如 Name, Size 等），主要是因為該控制項包含了不適當的 `DesignerGenerated` 屬性以及一個多餘的 `.Designer.vb` 檔案，導致設計工具將其視為獨立設計頁面而非繼承自 `TreeView` 的控制項。

## 使用者需求摘要
- [ ] 修正後的 `SimTree` 在表單上需能顯示 Name, Size 等屬性。
- [ ] 保留現有的所有註解、Debug 紀錄與思考演進過程。
- [ ] 回覆內容使用繁體中文。
- [ ] 註解標記 `by AntiGravity, 2026/03/29`。

## 擬議變更

### 1. 簡化控制項結構

#### [MODIFY] [Form1_SimTree.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_SimTree.vb)
- 在類別上方添加 `<ToolboxItem(True)>`, `<DesignTimeVisible(True)>` 以確保設計工具正確載入。
- 將原先在 `.Designer.vb` 中的 `Dispose` 方法移入主檔案。
- 移除對 `Partial Class` 的依賴（將其改為標準 `Public Class`），除非有特殊多檔案開發需求。

#### [DELETE] [Form1_SimTree.Designer.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_SimTree.Designer.vb)
- 刪除此檔案，避免其中的 `InitializeComponent`（包含 `ClientSize` 設定）干擾設計工具。這正是屬性視窗空白的元兇。

### 2. 環境與編譯優化
- 修改 `Outlook Assistant.vbproj`（如有必要）以移除對 `.Designer.vb` 的顯式引照（現代 SDK 專案會自動處理，但需確認是否有殘留配置）。

## 開放性問題
- **控制項註冊名稱**：如果您習慣在設計工具箱看到 `SimTree` 而非 `Form1_SimTree`，我們可以透過屬性來微調顯示名稱。

## 驗證計畫

### 自動與手動驗證
1. **編譯驗證**：確保原始碼合併後無語法錯誤，且原始註解完整保留。
2. **設計工具測試**：在 Visual Studio 中重新建置專案，打開 `Form1.vb` 設計介面。
3. **屬性視窗**：點選表單上的 `SimTree` 控制項，確認 `Name`, `Location`, `Size` 等基本屬性出現。
