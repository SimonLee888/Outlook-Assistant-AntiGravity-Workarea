# 任務清單：ListView 分群計畫 (A, B, C)

## 實作步驟

### Phase A: `lvwDebug` (除錯視窗) 分群改造
- [x] **基礎結構更新**
    - [x] 在 `DebugForm` 加入 `Dictionary(Of String, ListViewGroup)` 用於追蹤現有群組。
- [x] **訊息寫入調整**
    - [x] 修改 `AddMessage3`：根據 `callingMethod` 取得或建立群組。
    - [x] 修改 `Timer_Tick`：將 `item.Group` 關聯至對應群組。
- [x] **UI 與 OwnerDraw 調整**
    - [x] 確保 `lvwDebug_DrawSubItem` 的繪製座標在分群下正確。
    - [x] 測試搜尋高亮與配對高亮。
    - [x] 檢查游標選中、多選複製功能的相容性。
    - [x] **by Gemini 3.0 Flash, 2026/04/20**

### Phase B: `ListView2` (日期統計) 一覽式分群
- [ ] **顯示邏輯重構**
    - [ ] 重寫 `FillListView2` (或相關渲染函數) 一次載入所有資料。
    - [ ] 按年份建立 `ListViewGroup`。
- [ ] **圖表連動 (Chart2) 邏輯**
    - [ ] 修改 `ListView2_SelectedIndexChanged`。
    - [ ] 實作：點選 Group Header -> 圖表顯示「年度比對」。
    - [ ] 實作：點選 Month Item -> 圖表顯示「月份細節」。
    - [ ] **by Gemini 3.0 Flash, 2026/04/20**

### Phase C: `ListView1` (資料夾統計) 標題摘要化
- [ ] **分群實作**
    - [ ] 修改 `InitTab1UI` 或資料填充處，按 PST Store 分群。
- [ ] **標題摘要更新**
    - [ ] 實作摘要字串生成：`Folder (Mail: X, Size: Y)`。
    - [ ] 將摘要寫入 `ListViewGroup.Header`。
    - [ ] **by Gemini 3.0 Flash, 2026/04/20**

## 驗證與完工
- [ ] 檢查三者視覺一致性。
- [ ] 建立 `walkthrough.md`。
