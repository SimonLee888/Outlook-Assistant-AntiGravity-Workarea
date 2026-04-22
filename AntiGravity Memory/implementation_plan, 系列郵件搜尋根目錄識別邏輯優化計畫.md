# 系列郵件搜尋根目錄識別邏輯優化計畫

此計畫旨在解決「不論選取何處，掃描結果都一致」的問題。目前的邏輯固定讀取 Tab1 的選取項，改動後將能智慧識別使用者是想針對當前選取的資料夾進行新掃描，還是想針對目前的掃描結果進行重新整理。

## 擬議變更

### [Component] Form1_MainTabs.vb

#### [MODIFY] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_MainTabs.vb)

- **新增私有變數**：
    - `Private _tab4LastSearchRoot As Outlook.Folder = Nothing`：用於記錄最近一次成功啟動搜尋的根目錄。
- **修改 `Button4_Click` 邏輯**：
    1. **智慧識別 RootFolder**：
        - 檢查 `TreeView4.SelectedNode?.Tag`。如果它是一個 `Outlook.Folder`，則將其設為 `_tab4LastSearchRoot` 並開使新的搜尋。
        - 否則，如果 `_tab4LastSearchRoot` 已經有值，則繼續使用該目錄（這支援了在搜尋結果頁面按下 F5 進行重新整理）。
        - 如果上述兩者皆無效，則嘗試 fallback 到 Tab1 的 `SimTree1.SelectedNode?.Tag`。
    2. **使用者提示**：
        - 如果完全無法識別有效的 Outlook 資料夾，則彈出訊息提示使用者選擇一個資料夾後再開始。
- **優化 ESC 重置邏輯**：
    - 當按下 ESC 重置目錄樹時，同時清除 `_tab4LastSearchRoot`，確保使用者下一次點擊搜尋時是清新的判斷狀態。

## 驗證計畫

### 手動驗證
1. **跨資料夾搜尋測試**：
    - 在 Tab4 按下 ESC 回復初始目錄樹。
    - 選取資料夾 A，點擊搜尋。確認結果是 A 的系列。
    - 再次按下 ESC，選取資料夾 B，點擊搜尋。確認結果切換為 B 的系列。
2. **F5 重新整理測試**：
    - 在搜尋結果產出後，隨便選一個「系列主題」節點按下 **F5**。
    - 確認系統是否會針對當初選取的「同一個資料夾」重新啟動掃描。
3. **Fallback 測試**：
    - 按下 ESC 且不選取任何資料夾。
    - 到 Tab1 選取資料夾 C。
    - 回到 Tab4 直接點擊搜尋（不選 `TreeView4` 內容）。
    - 確認系統是否能自動去抓 Tab1 的資料夾 C。
