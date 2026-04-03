# DebugForm 佈局優化計畫

為了提升 `DebugForm` 的佈局穩定性並與主程式（如 Tab5）的 UI 風格保持一致，我們將對其進行佈局重構。

## 修改目標

將原本依賴 Designer/Resource 的佈局設定改為在 `DebugForm_Load` 中顯式定義，確保穩定性與一致性。

### DebugForm.vb
(已完成佈局優化)

### Form1.vb
- **優化 Tab2 佈局 (容器化)**:
  - 為 `Tab2` 建立 `pnlOptions2` 面板並設為 `Dock = Top`。
  - 將 `CheckSub2` 移入 `pnlOptions2`，解決其在圖表上飄浮的問題。
  - 將 `ListView2` 設為 `Dock = Top` (高度 250)，`Chart2` 設為 `Dock = Fill`。
- **統一層次管理**: 確保所有面板 `SendToBack()`，所有 Fill 內容 `BringToFront()`。

### DebugForm.vb
- **優化 ListView 繪製一致性 (解決跳動)**:
  - 在 `DrawSubItem` 的所有 `TextRenderer.DrawText` 與 `MeasureText` 呼叫中，統一強制使用 `TextFormatFlags.NoPadding`。
  - 確保文字矩形 (`textRect`) 在所有情況下內縮量一致。
- **修復選取列高亮**:
  - 修正邏輯，確保當行被選取（藍底）時，關鍵字的高亮（黃底黑字）仍然正確疊加。
- **優化對齊**:
  - 調微調 `chkSearchLogic` 的 `Top` 偏移，使其視覺中心與 `txtDebug` 內的文字對齊。

## 驗證計畫
1. **繪製驗證**: 在搜尋狀態下點選不同行，確認黃色高亮在藍色選取背景上依然可見，且文字座標與未搜尋時完全重合。
2. **Tab 2 驗證**: 確認「含子資料夾」CheckBox 穩固在頂部面板，不再隨圖表縮放而飄移。


## 驗證計畫

### 手動驗證
1. **標頭可見性測試**: 確認 `DebugForm` 與 `Tab3` 的列表標頭（Column Headers）沒有被搜尋框遮擋。
2. **比例測試**: 確認 `Chart2` 在 `ListView2` 下方有足夠空間。


