# Tab4 系列郵件佈局優化與 ListView4 分組排序實作計畫

此計畫旨在滿足使用者對 Tab4 UI 初始狀態的一致性需求，並增強 ListView4 的分組顯示與排序互動。

## 使用者評論與決策

- **TreeView4 初始狀態**：寬度需與左側 `SimTree4` 一致且不可顯示舊有或預設資料。
- **ListView4 分組排序**：
  - 實作「主旨分組」顯示。
  - **F6 快捷鍵**：在 `Subject` (主旨名稱) 與 `Count` (分組內郵件數量) 兩種排序方式間切換。

## 擬議變更

### [UI 佈局與初始化]

#### [MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1.vb)

- 在 `InitTab4UI` (或相關初始化處) 顯示設定 `scnrResults.SplitterDistance` 以符合 `SplitContainer4.SplitterDistance`，確保 `TreeView4` 初始寬度與 `SimTree4` 視覺一致。
- 確保 `TreeView4.Nodes.Clear()` 在啟動時呼叫。

### [ListView4 分組與快捷鍵實作]

#### [MODIFY] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_MainTabs.vb)

- **全域變數**：
  - [NEW] `Private _lv4GroupSortByCount As Boolean = False`：記錄目前是否按數量排序。
- **函數修改**：
  - `FillListView4`：
    - 加入 `ListViewGroup` 邏輯。
    - 根據 `_lv4GroupSortByCount` 對分組進行排序。
    - 如果按 `Count` 排序：分組數量多者在前。
    - 如果按 `Subject` 排序：按主旨字母排序。
  - `ListView4_KeyDown`：
    - 攔截 `Keys.F6`。
    - 切換 `_lv4GroupSortByCount` 並重新呼叫 `FillListView4`。
    - 更新 `ProgressBar2.Text` 提示目前排序模式。

## 驗證計畫

### 自動化/手動測試
- **佈局驗證**：啟動程式切換到 Tab4，確認 `TreeView4` 與 `SimTree4` 寬度比例正確且為空白。
- **分組功能**：搜尋系列郵件後，選取一個主題，確認 `ListView4` 出現主旨分組（例如標題包含 `Re:` 與不含 `Re:` 的分為兩組）。
- **F6 排序切換**：
    1. 按下 F6，確認 `ListView4` 的分組順序改變。
    2. 檢查 `ProgressBar2` 是否正確顯示「目前排序：按主旨名稱」或「目前排序：按郵件數量」。
    3. 確認分組標題格式為 `主旨名稱 (N 封)`。

## 開放問題

> [!NOTE]
> 關於「主旨分組」：同一個系列 (Topic) 內的郵件通常主旨極度相似。是否需要進行「主旨清理」（如移除 Re:, Fw:）後再分組，還是直接以原始主旨分組？
> 
> *目前預設按照「原始主旨」分組，若使用者有不同意見請告知。*
