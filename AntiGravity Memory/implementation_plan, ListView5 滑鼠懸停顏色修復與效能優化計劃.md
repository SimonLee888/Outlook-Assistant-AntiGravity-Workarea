# ListView5 滑鼠懸停顏色修復與效能優化計劃

## 問題描述
目前 Tab5 的 `ListView5`（重複郵件搜尋結果）存在以下問題：
1. **顏色消失**：當滑鼠移入項目時，自訂的群組背景色會被懸停灰色覆蓋；當滑鼠移出後，顏色會被重設為預設色，導致群組辨識功能失效。
2. **效能卡頓**：滑鼠移動時，程式會頻繁修改 `BackColor` 屬性，這在 WinForms 中會觸發 $O(N)$ 的全清單佈局重算，導致介面明顯延遲。

## 解決方案
比照 `ListView4` 的成熟方案，將 `ListView5` 改為 **OwnerDraw (自訂繪製)** 模式。
- **效能優化**：懸停時僅呼叫 `Invalidate()` 觸發重繪，不修改 `BackColor` 屬性，避免佈局重算。
- **顏色保留**：在繪製事件中，判斷若非懸停狀態，則畫出項目原本擁有的 `BackColor`（即群組顏色）。

## 預計變更內容

### Core Logic

#### [MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1.vb)

1. **`InitTab5UI`**: 
   - 顯式設定 `ListView5.OwnerDraw = True`。
2. **`InitListView`**:
   - 為 `ListView3`, `ListView4`, `ListView5` 統一掛載 `DrawItem` 事件（目前漏掉此項，可能導致某些狀態渲染不完全）。
3. **`HandleLv3Lv4Lv5_DrawSubItem`**:
   - 修正繪製邏輯：
     - 若為懸停項目：畫出懸停色 (`ThemeColors.MercuryGray`)。
     - 若非懸停項目且未選取：畫出該項目的 `e.Item.BackColor`（保留群組色）。
     - 若為選取項目：讓出由系統或專用邏輯處理。
4. **`HandleLv3Lv4Lv5_DrawItem` [NEW]**:
   - 實作基礎的 `DrawItem` 處理常式，設定 `e.DrawDefault = True` 以確保 Details 模式下的基本架構渲染正確。

## 驗證計劃

### 手動測試
- **顏色檢查**：執行重複郵件搜尋後，將滑鼠移入/移出 `ListView5` 的各個項目，確認群組背景色在滑鼠離開後會恢復。
- **效能檢查**：快速移動滑鼠，確認 UI 不會出現明顯卡頓或 FPS 下降。
- **選取狀態**：確認點選項目後，選取藍色（Highlight）能正常顯示，且不會被群組色蓋掉。

## 使用者確認事項
> [!NOTE]
> 目前 `ListView5` 的群組色是直接儲存在 `item.BackColor` 中。本計劃將會讓繪製引擎直接讀取該屬性進行渲染。
> 這樣的改動不需要更動 `Form1_MainTab345.vb` 中的搜尋邏輯，只需調整 UI 渲染層。
