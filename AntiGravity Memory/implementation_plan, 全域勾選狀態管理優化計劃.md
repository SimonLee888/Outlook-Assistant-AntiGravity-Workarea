# 全域勾選狀態管理優化計劃

此計劃旨在將 Tab2 與 Tab3 的「包含子資料夾」以及最後一頁的「顯示所有資料夾」勾選狀態，從直接讀取 UI 控制項改為使用全域變數存取。這樣做可以避免各個函數重複檢查 UI 狀態，並讓所有相關函數共用同一個布林值。我們將使用 `AddHandler` 在 UI 變更時同步更新這些變數。

## 使用者評論請求
> [!IMPORTANT]
> 此次修改涉及全域變數的引入與初始化邏輯的變更，請確認是否符合您的架構偏好。
> 所有的修改都將遵循「小塊寫入 (Chunked Edits)」原則。

## 擬議變更

### Form1 (全域變數定義與初始化)
- [MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1.vb)
    - 在「全域宣告」區域新增三個布林變數：
        - `_includeSubTab2` (對應 Tab2 `CheckSubFolder2`)
        - `_includeSubTab3` (對應 Tab3 `CheckSubFolder3`)
        - `_showAllFolders` (對應「顯示所有資料夾」`checkIncludeAllFolders`)
    - 在 `Form1_Load` 或相關初始化函數中，使用 `AddHandler` 綁定這些 CheckBox 的 `CheckedChanged` 事件，以便在狀態變動時即時更新全域變數。

### Tab 2 (依日期統計)
- [MODIFY] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_MainTabs.vb)
    - 修改 `SimTree2_AfterSelect`：將 `CheckSubFolder2.Checked` 改為讀取 `_includeSubTab2`。

### Tab 3 (附件搜尋)
- [MODIFY] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_MainTabs.vb)
    - 修改 `Button3_Click`：將 `CheckSubFolder3.Checked` 改為讀取 `_includeSubTab3`。

### Outlook 資料層 (資料夾遍歷)
- [MODIFY] [Form1_Outlook.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_Outlook.vb)
    - 修改 `GetSortedSubFolders`：將 `isIncludeAll` 的來源改為 `_showAllFolders`。
    - 修改 `GetSubFolderList`：將 `isIncludeAll` 的來源改為 `_showAllFolders`。

## 驗證計劃

### 手動驗證
1. 勾選/取消勾選 Tab2 的「包含子資料夾」，確認統計結果正確切換。
2. 勾選/取消勾選 Tab3 的「包含子資料夾」，執行搜尋，確認範圍正確。
3. 勾選/取消勾選「顯示所有資料夾」，確認 TreeView 展開時是否能正確過濾非郵件資料夾。
4. 使用 Dbg 訊息確認 `AddHandler` 觸發且變數值正確更新。
