# 修復 ListView4 開啟郵件功能與 SimTree4 ESC 邏輯衝突計畫

這是一個針對 Tab4（系列郵件）功能的修復計畫。使用者反映 `ListView4` 的 Enter 鍵與滑鼠點擊無法開啟郵件，且 `SimTree4` 的 ESC 鍵邏輯疑似重複。

## 使用者評論與需求
1. **ListView4 開啟功能失效**: 經查是因為分組排序後 Index 錯位，且代碼中存在誤植的控制項名稱（`TreeView4`）。
2. **SimTree4 ESC 重複**: 檢查是否存在多個事件處理器同時攔截 ESC 鍵。

## 方案設計

### 1. 解決 ListView4 Index 錯位問題
- **核心思維**: 不再依賴 `item.Index`，改為在建立 `ListViewItem` 時，將完整的 `MailItemInfo` 物件存入 `lvi.Tag`。
- **優點**: 無論如何分組、排序，`Tag` 永遠緊隨項目，取值路徑最短且最可靠。

### 2. 修正控制項名稱
- 將 `ListView4_KeyPress` 中誤植的 `TreeView4` 改為 `SimTree4`。

### 3. 清理 ESC 鍵重複處理
- 檢查 `Form1.vb` 與 `Form1_MainTabs.vb` 中的事件攔截。
- 確保導航邏輯清晰：
    - 在 `ListView4` 按 ESC -> 焦點回到 `SimTree4` (由 `ListView4_KeyDown` 處理)。
    - 在 `SimTree4` 按 ESC (若是搜尋結果模式) -> 退回資料夾模式 (由 `SimTree4_KeyDown` 處理)。

## 預定變動檔案

### Form1_MainTabs.vb [MODIFY]

#### [Layer1 UI事件層]
- **`ListView4_MouseClick`**: 改從 `item.Tag` 獲取郵件資訊更新路徑。
- **`ListView4_MouseDoubleClick`**: 改從 `item.Tag` 獲取 EntryID 並開啟。
- **`ListView4_KeyPress`**: 
    - 修正 `TreeView4` 誤植。
    - 改從 `item.Tag` 批量獲取 EntryID。
- **`SimTree4_KeyDown`**: 檢查並精簡 ESC 處理，避免與其他 Partial Class 衝突。

#### [Layer2 流程協調層]
- **`FillListView4`**: 在建立 `lvi` 時，將 `mailItem` 賦值給 `lvi.Tag`。
- **`RefreshListView4MailsAsync`**: 修正裡面引用的 `TreeView4` 並確保更新後的郵件資訊也能正確同步到 `Tag`。

## 驗證計畫

### 手動測試 (請使用者協助)
1. **開啟郵件**: 
    - 選取 Tab4，執行系列郵件搜尋。
    - 雙擊列表中的郵件，確認能正確開啟 Outlook 視窗。
    - 選取多封郵件按 Enter，確認能批次開啟（超過 10 封應有提示）。
2. **導航與 ESC**:
    - 在列表按 ESC，焦點應回到左側樹。
    - 在左側樹按 ESC，應能從系列主旨模式退回一般資料夾模式。
    - 確認不會觸發兩次行為或產生錯誤訊息。

## 開放性問題
> [!NOTE]
> 使用者之前的對話提到過 z-order 問題（ListView header 不動），目前的修改主要針對邏輯層，若 header 依然無法點擊排序，可能需檢查控制項重疊情況。
