# Tab3 UI 響應式佈局優化計畫

此計畫旨在解決 `GroupBox3` 在側邊欄縮合時無法正確顯示的問題。我們將判斷基準從「全域視窗寬度」改為「搜尋面板實際寬度」。

## 擬議變更

### 1. 表單變數與初始化 [Component]

#### [MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1.vb)

1.  **類別宣告區**：
    將 `pnlOptions_tab3` 提升為類別級別的變數 `_pnlOptionsTab3`，方便在不同事件中存取。

2.  **`InitTab3UI` 方法**：
    - 初始化 `_pnlOptionsTab3`。
    - 掛載 `Resize` 事件處理器（使用 Lambda 或獨立方法）。
    - 初始執行一次顯示判斷。

3.  **移除 `Form1_Resize` 中的舊邏輯**：
    刪除 L203-L205 中關於 `TabPage3` 與 `GroupBox3.Visible` 的判斷，減少不必要的重複計算。

### 2. 響應邏輯 [Logic]

- **閾值設定**：
  原設定 `SplitContainer3.Width >= 1100`。
  考慮到側邊欄預設寬度約為 280px，新的檢查條件將改為：
  `_pnlOptionsTab3.Width >= 820` (此數值可確保在側邊欄存在時若視窗夠大才顯示，且側邊欄縮合時能立即觸發顯示)。

## 開放性問題 (Open Questions)

- **其他 GroupBox 是否也需要類似處理？**
  目前 `GroupBox1` 與 `GroupBox2` 似乎是常駐顯示。如果未來功能更多，我們是否需要實作一整套「流式佈局 (Flow Layout)」？目前先針對 `GroupBox3` 進行這項優化。

## 驗證計畫

### 手動測試 (Manual Verification)
1.  **測試視窗縮放**：在 Tab3 正常模式下縮放寬度，確認 `GroupBox3` 在臨界點消失/出現。
2.  **測試分隔線縮合**：在視窗寬度不足（`GroupBox3` 隱藏狀態下），縮合左側 TreeView。
    - **預期結果**：右側麵板變寬後，`GroupBox3` 應自動彈出顯示。
3.  **測試分隔線展開**：將左側 TreeView 拉開（讓右側變窄）。
    - **預期結果**：面板寬度不足時，`GroupBox3` 應自動隱藏，確保不遮擋搜尋按鈕。
