# Tab3 & Tab4 互動邏輯整合計畫

本計畫旨在透過「代碼共用」來優化 Tab3 與 Tab4 的操作體驗，確保所有郵件清單控制項具備一致的快捷鍵與互動行為。

## 使用者評論要求
整合 Tab3 與 Tab4 的重複代碼，提升可維護性並確保操作手感一致。

## 擬議變更

### Form1_MainTabs

#### [MODIFY] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_MainTabs.vb)

1. **建立通用開啟邏輯 `HandleListViewOpeningMails`**：
   - 提取原本分散在各處的「檢查選取數 -> 確認開啟 -> 抓取 EntryID -> 呼叫 OpenMailByEntryID」流程。
   - 支援 `VirtualMode` (Tab3) 與 `NormalMode` (Tab4) 的資料抓取。

2. **統一事件處理**：
   - **KeyPress (Enter/ESC)**：建立 `CommonListViewKeyPress` 方法，供 `ListView3_KeyPress` 與 `ListView4_KeyPress` 呼叫。
   - **DoubleClick**：建立 `CommonListViewDoubleClick` 方法。
   - **MouseClick**：統一複製與進度條同步邏輯。

3. **清理冗餘代碼**：
   - 移除 `ListView4_KeyPress` 等地方手動寫入的 `TryCast` 檢索與開啟邏輯。

## 開放問題
無。

## 驗證計畫

### 手動測試
1. **開啟測試**：在 Tab3 與 Tab4 分別選取郵件後按 Enter，確認皆能正確開啟（含多選提示）。
2. **ESC 測試**：確認按下 ESC 皆能清除選取，且 Tab4 會將焦點移回資料夾樹。
3. **滑鼠測試**：確認雙擊皆能開啟郵件，單擊皆能同步路徑。
