# Tab4 快捷鍵功能修復與邏輯統一計畫

針對使用者回饋的 Tab4 (系列郵件) 快捷鍵問題，本計畫將修正 `ESC` 無效、`F6` 功能歸屬異常，並統一處理控制項名稱更動後的殘留程式碼。

## 使用者評論與主要問題

> [!IMPORTANT]
> 1. **ESC 不作用**：在 `ListView4_KeyPress` 內的 `ESC` 切換邏輯無效，且與 `KeyDown` 事件重複。
> 2. **F6 歸屬錯誤**：`F6` 目前在 `ListView4` 下觸發的是「郵件列表排序」，但需求應為針對左側 `SimTree4` 的「系列主題分組排序」。
> 3. **控制項混亂**：程式碼中仍有部分地方引用已移除或隱藏的 `TreeView4`，應統一改為 `SimTree4`。

## 擬議變更

### [Component] Tab4: 系列郵件 (Form1_MainTabs.vb)

#### [MODIFY] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_MainTabs.vb)

1. **清理 `ListView4_KeyPress`**：
   - 移除 `Keys.Escape` 判斷（因為 `KeyPress` 處理功能鍵效果不佳，且已在 `KeyDown` 實作）。
   - 移除 `Keys.Enter` 內部的 `TreeView4` 引用，改為 `SimTree4`。

2. **修正 `ListView4_KeyDown`**：
   - **ESC 鍵**：保留目前的 `SimTree4.Focus()` 邏輯，確保在列表端按下 ESC 能回到左側。
   - **F6 鍵**：修改邏輯，觸發 `_tab4SortGroupsByCount` 的切換，並呼叫 `RenderTab4Groups(_tab4LastTopicResults)` 重新對 `SimTree4` 進行排序渲染。

3. **修正 `RefreshListView4MailsAsync`**：
   - 將內部引用的 `TreeView4.SelectedNode` 改為 `SimTree4.SelectedNode`。

---

## 驗證計畫

### 手動測試
- **ESC 測試**：在 `ListView4` 選中郵件時按下 `ESC`，確認焦點是否回到 `SimTree4`。
- **F6 測試**：在 `ListView4` 獲得焦點時按下 `F6`，確認左側 `SimTree4` 的「系列主題」是否在「按次數排序」與「按主旨排序」間切換。
- **功能流測試**：點選 `SimTree4` 的系列主題，查看 `ListView4` 是否能正常顯示該主題下的郵件。
