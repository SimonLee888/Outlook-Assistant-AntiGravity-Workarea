# Tab4 UI 二欄式復原與一樹兩用強化計畫

為了優化操作空間並徹底解決中間 `TreeView4` 被系統誤載入的問題，我們將 Tab4 的介面由「三欄」恢復為「二欄」，並讓左側的 `SimTree4` 同時具備顯示資料夾與顯示搜尋結果的能力。

## 使用者回饋重點 (User Review Required)

> [!IMPORTANT]
> **模式切換行為**：
> - 搜尋前：左側顯示「資料夾樹」。
> - 搜尋後：左側自動變更為「郵件系列清單」。
> - 若要回頭選資料夾，需按下 **ESC** 或點選 **Reset** 按鈕。

## 預計變更方案

### 1. UI 結構重組 (Form1.vb)

---
#### [MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1.vb)

- **移除 Nested SplitContainer**：廢棄 `_scnrTab4Results`，將其子控制項直接掛載。
- **重新配置 SplitContainer4**：
  - `Panel1`：直接放置 `SimTree4` (Dock=Fill)。
  - `Panel2`：放置 `pnlOptions_tab4` (Top) 與 `ListView4` (Fill)。
- **變數定義**：新增 `_isTab4ShowingResults` (Boolean) 來標記目前左側樹的狀態。

### 2. 搜尋與事件邏輯 (Form1_MainTabs.vb)

---
#### [MODIFY] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_MainTabs.vb)

- **`Button4_Click` (搜尋按鈕)**：
  - 在呈現結果前，將 `_isTab4ShowingResults` 設為 `True`。
  - 將搜尋到的系列 (Topics) 直接以 `TreeNode` 形式填入 `SimTree4` (取代原本的資料夾)。
- **`SimTree4_AfterSelect` (選取事件)**：
  - **模式判斷**：
    - `If Not _isTab4ShowingResults`：執行原本的資料夾邏輯。
    - `Else`：將選中的主旨節點資料，填入右側的 `ListView4`。
- **`SimTree4_BeforeExpand` (展開防治)**：
  - 在結果模式下，阻斷原本自動載入 Outlook 資料夾的行為。

### 3. 操作體驗優化 (鍵盤與切換)

- **ESC 恢復功能**：在 `Form1_KeyDown` 或是 `SimTree4` 的偵聽中加入 ESC 判斷。當按下 ESC 時，清空 `SimTree4` 並呼叫 `LoadStoreToTreeView` 恢復資料夾視圖。

## 驗證計畫

### 自動化/手動測試路徑
1. **啟動測試**：確認進入 Tab4 時左側為空白，且不會自動載入資料到預期之外的地方。
2. **搜尋測試**：
   - 選取資料夾 -> 按下搜尋。
   - 確認左側樹內容變為「郵件主旨」而非資料夾。
   - 點選左側主旨 -> 右側列出郵件。
3. **恢復測試**：按下 ESC，確認左側恢復為資料夾樹。
4. **佈局測試**：確認搜尋後縮合左側，右側 ListView 能獲得最大寬度。

## 檔案位置參考
- 初始化：`Form1.InitTab4UI`
- 核心邏輯：`Form1_MainTabs.vb` 中的 Tab4 區塊。
