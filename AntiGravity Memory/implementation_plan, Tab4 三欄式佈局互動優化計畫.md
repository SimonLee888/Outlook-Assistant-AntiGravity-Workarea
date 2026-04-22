# Tab4 三欄式佈局互動優化計畫

此計畫旨在精進 Tab4 的使用者體驗，透過自動化動作與權重設定，確保介面在不同解析度與操作階段下都能呈現最重要的資料。

## 擬議變更

### [Component] Form1.vb

#### [MODIFY] [InitTab4UI](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1.vb)

- **優化分欄權重**：
    - 將內層分欄 `scnrTab4Results` 的 `FixedPanel` 設為 `Panel1` (即中間的系列主題欄)。
    - **效果**：當外層分欄 (SplitContainer4) 縮合或視窗拉大時，中間攔寬度保持不變，所有額外空間將自動分配給右側的 `ListView4`。這解決了使用者提出的「放大的應該是 listview4」的問題。

### [Component] Form1_MainTabs.vb

#### [MODIFY] [Button4_Click](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_MainTabs.vb)

- **自動縮合邏輯**：
    - 在啟動搜尋時，檢查 `SplitContainer4.SplitterDistance`。
    - 如果寬度 > 20px，則自動將其縮合至 10px，並將原寬度存入 `Tag`。
    - **效果**：搜尋開始後，目錄樹會自動靠邊，讓出空間給搜尋結果。

#### [MODIFY] [TreeView4_KeyDown] (ESC 鍵處理)

- **自動回復邏輯**：
    - 在按下 ESC 清除結果後，檢查 `SplitContainer4.SplitterDistance`。
    - 如果寬度 <= 20px，則從 `Tag` 讀取原寬度並恢復。
    - **效果**：重置介面後，目錄樹會自動滑回原位，方便使用者重新選取資料夾。

## 驗證計畫

### 手動驗證
1. **擴展開行為測試**：
    - 雙擊 `SplitContainer4` 的分隔線進行縮合。
    - 觀察中間欄 (`TreeView4`) 寬度是否維持，且 `ListView4` 是否填滿剩餘空間。
2. **搜尋自動化測試**：
    - 選取資料夾後按下「搜尋系列郵件」。
    - 確認左側 `SimTree4` 是否自動推到最左邊 (10px)。
3. **ESC 自動化測試**：
    - 在搜尋結果頁面按下 ESC。
    - 確認側邊欄是否自動彈回原先的寬度，且焦點回到 `SimTree4`。
