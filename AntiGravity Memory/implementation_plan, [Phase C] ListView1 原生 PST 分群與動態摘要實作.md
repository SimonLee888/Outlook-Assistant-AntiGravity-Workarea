# [Phase C] ListView1 原生 PST 分群與動態摘要實作

此計畫旨在將 `ListView1` (資料夾大小統計) 從目前的「手動插入標題列」模式，升級為系統原生的 `ListViewGroup` 模式。分群將以 **PST 資料庫 (Store)** 為單位，並在群組標題中即時顯示該 PST 的統計總計（總封數與總大小）。

## 使用者評論與要求 (User Review Required)

> [!IMPORTANT]
> **歷史紀錄保留**：本計畫會嚴格遵守「保留函數層級歷史記錄」的要求。原本 v1~v5 的重構摘要與各開發者的 By-tag (Claude, Gemini 3.1 Pro) 都會被保留或濃縮至新版中。

> [!TIP]
> **OwnerDraw 策略**：由於 `ListView1` 啟用了 `OwnerDraw`，原生群組的加入可能會影響繪製行為。我將確保群組標題維持系統預設美感，而資料列則繼續沿用自訂的高亮與顏色邏輯。

## 預計提案變更 (Proposed Changes)

### [Core] 重構 Tab1 的渲染與分群邏輯

#### [MODIFY] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual Studio/Visual Studio 18 (2026)/Outlook Assistant - (AntiGravity測試區)/Form1_MainTabs.vb)

##### 1. 更新 `SimTree1_AfterSelect` 
- 修改 loop 邏輯：
    - 在填充 `allItems` 之前，先計算各個 PST (Store) 的匯總數據 (Total Items, Total Size)。
    - 使用 `ListView1.Groups` 根據 `folder.Store.DisplayName` 建立群組。
    - **動態更新群組標題**：`[Store Name] (共 x,xxx 封 / x.xx MB)`。
- **註解處理**：保留 2026-04-08/04-13 等關鍵重構節點的邏輯說明。

##### 2. 停用 `BuildGroupHeaderItem_Tab1` 
- 將原本用於偽裝標題的函數內容註解掉或標註為 [Legacy]，改由 `ListViewGroup` 取代其角色。
- 調整原本 root 資料夾的顯示方式，使其作為群組內的第一個顯眼項目。

##### 3. 修改 `BuildListViewItem_Tab1`
- 移除「▸」等手動縮排字元，改用更乾淨的原生列表樣式。
- 確保 `Tag` (ValueTuple) 結構保持不變，以免破壞雙擊與右鍵選單功能。

##### 4. 調整 `ListView1_DrawSubItem`
- 由於不再有 `Tag = Nothing` 的偽裝標題列（改用 native groups），這部分的繪製邏輯會簡化，僅保留「合計列」與一般列的高亮效果。

---

## 驗證計畫 (Verification Plan)

### 自動與手動測試
- [ ] **分群驗證**：選擇跨不同 PST (例如「個人資料夾」與「封存資料夾」) 的節點，確認 `ListView1` 是否正確依 PST 切分群組。
- [ ] **統計驗證**：比對群組標題內的總計數字，是否等於其下方所有子資料夾數據的總和。
- [ ] **交互驗證**：
    - 雙擊群組內的資料夾，確認仍能正確進入下層統計。
    - 右鍵點擊某個資料夾，確認「計算大小」功能仍能運作且只更新該項。
- [ ] **註解核對**：確認文件開頭的 v1~v6 演進紀錄是否完整。

### 回歸測試
- [ ] 確認 `ListView2` 的月份背景加載與 Chart2 連動是否在剛才修復後維持正常。
- [ ] 檢查 `lvwDebug` 的呼叫端分群是否依然清晰。
