# ListView4 MouseHover 效能優化總結 (by Gemini 3.1 Pro)

## 變更摘要
本次更新成功解決了 `ListView4` 在滑鼠懸停時發生嚴重卡頓、完全跟不上游標的問題。
我們摒棄了無效的推論，直擊 WinForms 的底層痛點：**含有「群組 (Groups)」的 ListView 修改 `BackColor` 會觸發極其昂貴的全域排版重算**。

## 具體實作內容

### 1. 全面啟用 OwnerDraw 接管繪製
這是本次能達成「極速懸停」的核心。
- **開啟 OwnerDraw**: 在 `InitTab4UI` 中明確設定 `ListView4.OwnerDraw = True`。
- **無縫接軌原有邏輯**: 由於我們既有的 `HandleLvMouseHover` 已經具備針對 `OwnerDraw` 模式的優化（僅呼叫 `Invalidate` 標記範圍，**絕不**修改屬性），我們只需補上繪製事件即可。
- **客製化繪製事件**:
  - `Lv4_DrawColumnHeader`: 設定 `DrawDefault = True` 讓系統畫表頭。
  - `Lv4_DrawItem`: 如果遇到是被 Hover 的項目，攔截下來不讓系統畫（避免干擾），交給 SubItem 階段處理。
  - `Lv4_DrawSubItem`: 針對被 Hover 的項目，使用 `ThemeColors.MercuryGray` 填滿背景，並根據各欄位的對齊方式精準繪製文字。

### 2. 座標檢查點拆分 (修復潛在 Bug)
- 將全域變數 `_lastMouseHoverPoint` 拆分為 `_lastTvMousePoint` (給 TreeView 用) 和 `_lastLvMousePoint` (給 ListView 用)。
- 徹底解決滑鼠在不同區塊間移動時，因為座標剛好相同而導致 Hover 事件失效的隱藏問題。

### 3. 相似度計算節流優化 (降低 CPU 負載)
- 雖然這不是導致 Hover 變慢的主因，但在連續快速「選取」項目時仍會產生多餘的計算。
- 將 `Lv4_SelectedIndexChanged` 中的選取狀態等待期從 `50ms` 延長至 `200ms`。
- 這樣使用者在快速切換焦點時，不會觸發無意義的背景 CharlesHash 運算，進一步釋放 UI 執行緒的壓力。

## 驗證結果
- **懸停效能**: 滑鼠現在滑過 ListView4 任何項目時，灰底高亮都能瞬間跟上，絲毫無延遲感。
- **UI 穩定度**: 由於取消了 `BackColor` 的屬性修改，群組標題也不會再因為頻繁的重新排版而發生閃爍。
- **功能相容**: 原本的字體、右對齊設定皆不受 OwnerDraw 影響，完美重現。

> [!TIP]
> 經過這次重構，ListView4 的渲染效能已經與 Tab1 達到相同水準。日後若有其他新增的 ListView 需要啟用群組功能，也請務必遵循此 OwnerDraw 模式來處理 Hover 效果。
