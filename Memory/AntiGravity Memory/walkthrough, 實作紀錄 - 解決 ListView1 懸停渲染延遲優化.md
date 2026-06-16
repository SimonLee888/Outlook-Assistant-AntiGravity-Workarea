# 實作紀錄 - 解決 ListView1 懸停渲染延遲優化

本文件總結了針對 `ListView1` 在處理大量郵件資料夾數據時，滑鼠懸停（Mouse Hover）游標反應速度慢、不跟手的優化歷程與解決方案。

## 問題背景
當 `ListView1` 載入超過數百個項目時，原先的懸停邏輯會觸發顯著的 UI 延遲：
- **現象**：滑鼠移過每一行時，淡灰色的高亮區塊會延遲出現或消失，無法即時跟隨游標。
- **根因**：原先邏輯透過修改 `ListViewItem.BackColor` 屬性來實現 Hover 效果。在 WinForms 中，修改屬性會導致內部重新計算佈局（Layout Pass），在大量數據下，這種 O(N) 的開銷會導致嚴重的渲染阻塞。

## 優化方案
我們借鑒了之前在 `DebugForm` 處理大量 log 渲染時的優化模式，改採 **「零屬性變更 + 局部重繪」** 策略。

### 1. 隔離屬性與渲染 (Form1.vb)
在 `HandleListViewMouseHover` 事件中，我們不再寫入 `BackColor` 屬性：
- 檢查控制項是否開啟 `OwnerDraw`。
- 如果是 `ListView1`（OwnerDraw 模式），則僅呼叫 `listView.Invalidate(Bounds)` 通知系統該矩形區域需要重繪，而不是請求重新佈局。
- 這樣可以確保 UI 執行緒不會因為一次滑鼠移動就去遍歷幾千個項目的屬性狀態。

### 2. 手動繪製 Hover 狀態 (Form1_MainTabs.vb)
在 `ListView1_DrawSubItem` 中接管繪圖邏輯：
- 判斷目前繪製的項目是否為 `_lastHoveredListItem`。
- **自訂背景**：如果是 Hover 項目且未選取，我們手動用 `ThemeColors.MercuryGray` 填滿背景。
- **文字對齊**：使用 `TextRenderer.DrawText` (GDI) 確保文字對齊、邊距與系統原生渲染一致。

> [!TIP]
> **為什麼這樣比較快？**
> `Invalidate()` 只是在 Windows 訊息佇列中排入一個繪製標記，開銷極低。這種做法將渲染負擔從「全域佈局計算」轉向「單點像素繪圖」，在大量數據下能將反應速調提升到即時（O(1)）。

## 驗證結果
- **反應速度**：滑鼠在大量資料夾間滑動時，高亮區塊反應即時，完全跟手。
- **一致性**：此優化模式與 `DebugForm` 的 `lvwDebug_ItemSelectionChanged` 採用相同的效能優化思維，確保了 codebase 內對於大型列表處理的技術風格一致。
- **相容性**：不影響原有的群組標題行（淡藍色）與合计列（深藍色）的固定配色方案。

## 相關修改文件
- [Form1.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1.vb)
- [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_MainTabs.vb)

---
by Gemini 3.0 flash, 2026/04/14
