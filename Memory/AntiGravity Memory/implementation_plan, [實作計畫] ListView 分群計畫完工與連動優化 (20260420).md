# [實作計畫] ListView 分群計畫完工與連動優化 (2026/04/20)

這份計畫旨在銜接並完成之前中斷的「ListView 分群計畫」，重點在於完善 Tab2 (日期統計) 的一覽式分群與圖表連動，以及檢核 Tab1 (資料夾統計) 的 PST 原生分群。

## User Review Required

> [!IMPORTANT]
> **關於 ListView2 (Tab2) 的操作邏輯變更：**
> 1. 目前已實作「一覽式報表」，即所有年份與月份直接以 Group 形式展開在 ListView 中，不再需要「雙擊進入月份」的切換感。
> 2. 計畫加入 **Group Header 點擊監聽**：點選年份群組標題 (Header) 時，右側 Chart2 將自動切換為該年份的月份分佈圖。
> 3. 目前 Phase C (Tab1) 已初步完成 PST 分群，我們將檢查摘要顯示的視覺效果。

## 待完成項目 (Proposed Changes)

### 0. COM 開銷稽核與修正 (Immediate Fix)
> [!CAUTION]
> **修正不必要的 .FolderPath 呼叫**: 
> 檢查 `SimTree1_AfterSelect` 與 `BuildListViewItem_Tab1`，將原本誤用的 `Folder.FolderPath` 屬性讀取改回使用 `FolderBfsEntry.FolderPath` 預存字串，確保大迴圈中零 COM 開銷。

### 1. Phase B: `ListView2` (日期統計) 一覽式報表完工
目前的 `ListView2_SelectedIndexChanged` 已能處理月份項目的連動，但缺乏對 Group Header 的點擊支援。

#### [MODIFY] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_MainTabs.vb)
- **新增 `ListView2_MouseDown`**: 利用 `HitTest` 檢測使用者是否點擊了群組標題。
- **實作 Header 點擊連動**: 點擊標題時，讀取群組名稱 (年份)，並呼叫 `RenderCtMonthView` 更新圖表。
- **優化背景加載**: 確保背景加載月份資料時，UI 響應流暢，不影響切換圖表。

### 2. Phase C: `ListView1` (資料夾統計) 視覺檢核
- 檢查 `SimTree1_AfterSelect` 產出的 Group Header 摘要是否包含 Size 資訊 (目前代碼中似乎已加入，但需確認快取同步)。
- 確保合計列 (Sum Row) 在分群模式下的顯示位置正確 (不屬於任何群組)。

### 3. 進度管理與完工
- 更新 `task.md` 並反映當前真實進度。
- 建立 `walkthrough.md` 紀錄最終視覺效果。

## 開放問題 (Open Questions)

- **ListView2 的「返回年度統計」按鈕**: 在一覽式報表中，您是否仍希望保留原本月份視圖中的「← 返回年度統計」項？或者改為點擊圖表空白處返回「年度總趨勢圖」？
- **ListView1 的 Header 背景**: 原生 ListViewGroup Header 無法輕易自定義背景顏色 (Gradient)。目前 Phase C 改用原生分群，原本精美的藍色漸層背景會消失，如果您強烈需要該視覺效果，我們可能需要維持「虛擬標題列 (OwnerDraw)」方式。但原生分群的好處是標準、可手動收合。**目前建議先使用原生分群。**

## 驗證計畫 (Verification Plan)

### 手動測試 (Manual Verification)
1. **Tab1 測試**: 多選資料夾，檢查 ListView1 是否按 PST Store 名稱正確分群，及其標題是否顯示 `(共 X 個資料夾, Y 封 / Z MB)`。
2. **Tab2 測試**: 
   - 點擊年份群組標題，檢查 Chart2 是否顯示該年月份分佈。
   - 點擊群組內的具體月份，檢查 Chart2 是否高亮對應長條。
   - 檢查背景加載過程是否有正確顯示「讀取中...」並隨後自動填充。
3. **性能測試**: 檢查大數據量下 `ListView1` 開啟 `ShowGroups` 後的渲染速度。
