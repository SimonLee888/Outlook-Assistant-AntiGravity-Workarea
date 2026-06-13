# 移除 _isTv4ResultMode 改用 LvSearch4.Visible 判定之變更說明與複檢紀錄

本變更旨在徹底移除原本用於區分 Tab4 搜尋結果與資料夾樹的私有布林變數 `_isTv4ResultMode`，改為直接依據控制項顯示狀態 `LvSearch4.Visible` 來進行判定，落實 **Single Source of Truth (單一真實水源)** 設計，減少狀態同步的維護開銷。

## 變更項目

### 1. [Form1.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1.vb)
*   **第 675 行**：將 `_isTv4ResultMode = False` 初始化代碼註解掉，並留存修改註記。
    ```vb
    ' _isTv4ResultMode = False ' 初始為資料夾模式 ' by Claude Sonnet 4.6, 2026/05/29: 已廢棄，改用 LvSearch4.Visible 代替
    ```

### 2. [Form1_MainTab345.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_MainTab345.vb)
*   **第 14 行**：將私有變數 `_isTv4ResultMode` 的宣告註解掉，標記為廢棄。
*   **第 687 行** (`Tv4_AfterSelect`)：將原本 `If Not _isTv4ResultMode Then Return` 替換為 `If Not LvSearch4.Visible Then Return`。
*   **第 800 行** (`LvSearch4_KeyDown`)：
    *   Enter 鍵處理邏輯中，將原本 `If _isTv4ResultMode AndAlso ...` 改為 `If LvSearch4.Visible AndAlso ...`。
    *   Escape 鍵處理邏輯中，將 `_isTv4ResultMode = False` 設定代碼註解。
*   **第 1028 行** (`RenderLv4Group`)：將 `_isTv4ResultMode = True` 賦值代碼註解。

### 3. [Form1_Outlook.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_Outlook.vb)
*   **第 415 行** (`LoadSubFolderToTreeView`)：在 TreeView 的展開前置過濾中，將 `_isTv4ResultMode` 判斷改為對 `LvSearch4.Visible` 的判斷，確保若處於搜尋結果模式下時能正確退出。
    ```vb
    If sender Is SimTree4 AndAlso LvSearch4.Visible Then Exit Sub ' by Claude Sonnet 4.6, 2026/05/29: 將 _isTv4ResultMode 改為 LvSearch4.Visible
    ```

---

## 複檢與驗證結論

每處修改完成後，均已主動使用 `view_file` 工具重新讀取該程式碼行，確認：
1. **變數作用域一致**：原 `_isTv4ResultMode` 的操作皆已乾淨註解。
2. **邏輯流暢且對齊**：所有原先基於 `_isTv4ResultMode` 的分流均完美對齊至 `LvSearch4.Visible`。
3. **無多餘代碼殘留**：沒有遺留未註解的舊代碼片段。
