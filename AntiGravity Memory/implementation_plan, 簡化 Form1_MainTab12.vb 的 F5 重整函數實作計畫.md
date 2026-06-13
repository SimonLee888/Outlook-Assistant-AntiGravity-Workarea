# 簡化 Form1_MainTab12.vb 的 F5 重整函數實作計畫

本計畫旨在利用 `Form1_SimTree.vb` 自訂控制項中新增的狀態備份與還原 API (`GetTreeState`、`RestoreTreeState`)，來簡化 `Form1_MainTab12.vb` 中的 `ForceRefreshSimTree` 函數。

## 使用者審查要求

- **繁體中文回覆**：此實作計畫、後續的 `task.md` 與 `walkthrough.md` 均使用繁體中文。
- **保留歷史註解**：我們將完整保留 `ForceRefreshSimTree` 中有關 `2026/05/17` 改回 `FireAfterSelect` 等珍貴的 debug 歷程與思考演進註解，絕不整段刪除。
- **添加新註解標記**：新增的修改點會明確附上 `by Gemini 3.0 Flash, 2026/05/18` 標記與當天日期。
- **過時函數標記**：舊有的 `CollectExpandedPaths` 與 `ReExpandNodeByPath` 私有輔助函數予以保留，但加上註解標明已被新控制項 API 取代，以維護程式碼演進完整性。

## 變更項目

### [MODIFY] [Form1_MainTab12.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_MainTab12.vb)

#### 1. 簡化 `ForceRefreshSimTree`
將 `ForceRefreshSimTree` 內部的狀態收集（步驟 ①）以及狀態還原（步驟 ④ 與 ⑤）邏輯，替換為高階控制項 API `tv.GetTreeState()` 與 `tv.RestoreTreeState()`。

**修改前 (大綱)**：
- 手動遍歷 Nodes 收集展開路徑
- 手動遍歷 `tv.SelectedNodes` 收集選取路徑
- 重新載入樹
- 手動由淺到深逐條展開路徑
- 手動在 Nodes 中搜尋並還原選取節點，呼叫 `tv.FireAfterSelect(firstNode)`
- 若找不到舊選取，呼叫 `ExpandTvToDefaultInbox`

**修改後 (大綱)**：
- 呼叫 `Dim state = tv.GetTreeState()` 備份狀態
- 重新載入樹
- 呼叫 `tv.RestoreTreeState(state.ExpandedPaths, state.SelectedPaths, fireEvent:=True)` 還原展開與選取，並觸發選取事件
- 若 `tv.SelectedNodes.Count = 0`，則呼叫 `ExpandTvToDefaultInbox(tv)` 退回預設 Inbox。
- 完整保留 2026/05/17 等 debug 註解，並增加 `by Gemini 3.0 Flash, 2026/05/18` 新修改標記。

#### 2. 標記舊有的備份/還原私有函數
在 `CollectExpandedPaths` 與 `ReExpandNodeByPath` 的 XML 註解或說明中，標明已被 `SimTree` 自帶 API 取代。

---

## 驗證計畫

### 手動驗證
1. 啟動應用程式。
2. 在 Tab 1 / Tab 2 展開多個資料夾，並選取其中一個資料夾以顯示信件清單。
3. 對 `SimTree1` 按下 **F5** 鍵，或者執行重整。
4. 檢查樹狀結構是否能夠在瞬間正確重新載入、展開原本已展開的資料夾、還原選取的資料夾，並且右側 `ListView` 能否正確同步刷新信件統計與內容。
5. 驗證過程中是否沒有發生任何例外，且焦點與選取狀態正確無誤。
