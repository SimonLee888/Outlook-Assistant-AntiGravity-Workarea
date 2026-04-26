# Walkthrough: Tab7 UI 優化與互動功能更新紀錄

本次更新主要針對 Tab7 (OST/PST 解析) 進行了佈局自動縮放、郵件開啟互動以及資料夾刪除確認邏輯的實作，為接下來的大量測試做好準備。

## 1. UI 佈局優化 (Anchor 設置)
為了解決視窗放大後控制項無法自動延展的問題，我們實作了動態 Anchor 設置邏輯：
- **實作方式**：在 `LoadOST_Click` 時觸發 `InitTab7Layout()`。
- **效果**：
  - `SimTreeOST` (左上) / `SimTreePST` (左下)：固定於左側，隨高度變化。
  - `ListViewOST` (右上) / `ListViewPST` (右下)：水平與垂直雙向延展，填滿中間區域。
  - **右側按鈕群**：固定於右側邊緣，保持位置不變。

## 2. 郵件開啟互動 (雙擊 / Enter)
現在您可以直接從 OST 或 PST 的郵件清單中開啟項目。
- **OST 解析增強**：
  - 修改 `OstMailRow` 結構，加入 `EntryID` 欄位。
  - 在 `ReadOstFolderContentsL3` 中，透過 MAPI Tag `0x0FFF` (PidTagEntryId) 抓取原始二進位 ID 並轉為 Hex 字串。
- **事件整合**：
  - 為 `ListViewOST` 與 `ListViewPST` 綁定 `DoubleClick` 與 `KeyDown (Enter)` 事件。
  - 統一呼叫專案核心函式 `OpenMailByEntryID`，支援批次開啟 (超過 10 封會彈出警告)。
  - **支援類型**：郵件、工作、日曆、連絡人等，只要有 EntryID 即可透過 Outlook 開啟。

## 3. 資料夾刪除邏輯 (DeleteFolder)
新增了安全刪除資料夾的確認機制與移動邏輯。
- **安全確認**：點擊 `DeleteFolder` 時，會彈出 MessageBox 要求使用者確認，防止誤刪。
- **移動至回收桶**：
  - **PST 模式**：確認後，利用 Outlook OOM 將該資料夾移動到系統預設的「刪除的郵件」(Deleted Items) 資料夾，而非直接物理刪除，確保資料可救回。
  - **OST 模式**：由於目前採唯讀解析模式，系統會提示限制，保護原始檔案完整性。

## 4. 代碼結構優化 (By Gemini 3.0 Flash, 2026/04/23)
- **資料儲存**：在 `ShowOstItems` 與 `ShowPstItems` 中，將原始資料物件 (Row/MailInfo) 存入 `ListViewItem.Tag`，徹底解決了排序後 Index 錯位導致開啟錯誤項目的問題。
- **穩定性修正**：修正了 `SimTreeOST_AfterSelect` 的例外處理流程，確保 UI 狀態 (Cursor, ProgressBar) 在讀取失敗時能正確復原。

---
**測試建議**：
1. 嘗試拖拉視窗邊框，確認四格佈局是否如預期縮放。
2. 在 `ListViewOST` 選取一封郵件並按下 `Enter`，確認是否能彈出 Outlook 郵件視窗。
3. 測試選取 PST 資料夾後點擊 `Delete Folder`，確認資料夾是否正確移動到 Outlook 的回收桶。
