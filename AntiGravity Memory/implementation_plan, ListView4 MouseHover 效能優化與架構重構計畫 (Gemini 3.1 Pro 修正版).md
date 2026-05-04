# ListView4 MouseHover 效能優化與架構重構計畫 (Gemini 3.1 Pro 修正版)

## 背景與問題診斷
使用者反應 `ListView4` 的滑鼠懸停 (`MouseHover`) 反應極度遲鈍。前一模型 (3.0 Flash) 給出的解釋包含全域變數干擾、相似度計算負擔、以及雙緩衝未開啟等。

經過重新診斷，確認 3.0 Flash 的判斷存在重大誤區：
1. `MouseHover` 不會觸發 `SelectedIndexChanged`，因此相似度計算不是導致懸停變慢的原因。
2. `DoubleBuffered` 已經在 `InitListView` 統一為所有 ListView 啟用，ListView4 並未漏掉。
3. `_lastMouseHoverPoint` 的全域干擾只會導致提早 `Return` (Hover 沒反應)，並不會導致「變慢」。

**真正的效能瓶頸 (Root Cause)：**
ListView4 是全專案唯一啟用了 **「群組 (ListViewGroup)」** 功能的清單。在 WinForms 底層 (SysListView32) 的設計中，一旦啟用了群組，修改任何單一 Item 的 `BackColor` 屬性，都會強迫觸發整個 ListView 的重新排版 (Layout) 甚至全域重繪。當信件數量多時，滑鼠每移動一格就觸發一次 O(N) 的重算，這才是導致畫面完全凍結、跟不上滑鼠的元凶。

## 解決方案

### 1. 將 ListView4 改為 OwnerDraw 模式
比照 ListView1，不再直接修改 `BackColor` 屬性，而是透過 `OwnerDraw` 來自行渲染背景。
- 在 `InitTab4UI` 中開啟 `ListView4.OwnerDraw = True`。
- 修改 `HandleLvMouseHover`，當 `listView.OwnerDraw = True` 時，只對前一個和目前滑鼠所在的項目呼叫 `Invalidate(Bounds)`，**絕對不去修改** `BackColor` 屬性。
- 實作 ListView4 的 `DrawColumnHeader`, `DrawItem`, `DrawSubItem` 事件處理常式，自己畫上 `ThemeColors.MercuryGray` 的懸停底色。

### 2. 修復滑鼠座標檢查點干擾 (保留 3.0 Flash 正確的部分)
雖然不是變慢的主因，但 `_lastMouseHoverPoint` 全域共用的確是個 Bug。
- 將原本的 `_lastMouseHoverPoint` 拆分為 `_lastTvMousePoint` 與 `_lastLvMousePoint`，讓 TreeView 與 ListView 互不干擾。

### 3. [額外發現] ListView4 群組標題 (Group Header) 繪製
開啟 `OwnerDraw` 後，群組的標題繪製可能會變得比較棘手，因為 WinForms 原生不支援 `DrawListViewGroup`。但幸運的是，只要設定正確，`DrawDefault = True` 可以在多數情況下讓系統自己畫好標題，我們只需針對 `DrawSubItem` 介入一般項目的背景色即可。

## 預期修改範圍

### [MODIFY] `Form1.vb`
- **全域宣告**: 拆分 `_lastMouseHoverPoint` 為 `_lastTvMousePoint` 與 `_lastLvMousePoint`。
- **`HandleTvMouseHover`**: 改用 `_lastTvMousePoint` 判斷。
- **`HandleLvMouseHover`**: 改用 `_lastLvMousePoint` 判斷。

### [MODIFY] `Form1_MainTabs.vb`
- **`InitTab4UI`**: 加入 `ListView4.OwnerDraw = True`。
- **ListView4 繪製事件**: 新增 `Lv4_DrawColumnHeader` (DrawDefault = True)、`Lv4_DrawItem`、`Lv4_DrawSubItem` 來自行繪製 Hover 底色。

## 驗證計畫
1. 在 Tab4 載入包含多封信件的多個系列 (Groups)。
2. 快速移動滑鼠，確認 Hover 反應是否瞬間跟上，不再有任何卡頓。
3. 確認群組的標題 (Group Header) 顯示是否正常，沒有因為 OwnerDraw 破圖。
