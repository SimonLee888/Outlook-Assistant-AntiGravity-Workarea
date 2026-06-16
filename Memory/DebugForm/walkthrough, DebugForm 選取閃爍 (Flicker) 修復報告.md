# DebugForm 選取閃爍 (Flicker) 修復報告

## 修正事項
針對您反映的「移動 ListView `SelectedItemChanged` 時，若找到配對 (Pair) 會發生閃爍」的問題，已完成效能優化與修復。

### 根本原因 (Root Cause)
之前程式尋找到配對時，會直接設定 `pair.BackColor = Color.Cyan`。
這觸發了 WinForms 的一個潛在副作用：當您更改一個 `ListViewItem` 的核心屬性（例如字型或背景色）時，底層的 ListView 會認為它的「佈局或是整體樣式改變了」，有時會觸發大範圍的重新計算甚至是整表的 `Invalidate()` 重繪。當您頻繁用方向鍵上/下移動時，這種全表的重繪就會形成明顯的「閃爍 (Flickering)」。

### 修正行動
我們不再直接依賴修改 `ListViewItem` 的專屬屬性，而是發揮 `OwnerDraw` 的最大優勢：**渲染期決定顏色，並嚴格限定重繪範圍 (局部重繪)**。

已修改 [DebugForm.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/DebugForm.vb)：
1. **取消 `BackColor` 賦值**：
   在 `lvwDebug_ItemSelectionChanged` 中，不再設定 `.BackColor = Color.White` 或 `.BackColor = Color.Cyan`。
   取而代之的是，我們僅記錄哪一個是剛找出來的 `_lastHighlightedPair`，並針對「上一組配對」與「新配對」所**佔用的特定區域 (Bounds)**，呼叫 `lvwDebug.Invalidate(bounds)`。這能精確告知作業系統：「請**只**重畫這細微的兩小塊區域」。

2. **渲染期處理**：
   在 `lvwDebug_DrawSubItem` 的第一步繪圖背景時，如果發現目前正在畫的 `e.Item` 就是我們記下來的 `_lastHighlightedPair`，就直接用 `Color.Cyan` 當作背景色填滿。

## 驗證建議
請在 `DebugForm` 產生多筆 `開始` 與 `結束` 的巢狀記錄後：
> [!TIP]
> 1. 用滑鼠點選一行 `開始`，確保他有連動反白 `結束`。
> 2. 接著使用鍵盤的 **向上/向下方向鍵**，快速在各行之間切換選取。
> 3. 您會發現配對的高光 (Cyan) 反應變得非常順暢、無延遲，且整個 ListView 不會再發生任何閃爍與抖動！
