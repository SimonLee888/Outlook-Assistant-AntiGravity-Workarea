# Tab3 管線化架構重構任務清單

- [x] **第一階段：淨化 MAPI 操作，回歸 L3 底層**
  - [x] 在 `Form1_ComL3.vb` 中新增 `GetMailWithAttachmentL3`，並搬移原先在 L2 的 `GetTable` / `GetArray` 實作。
  - [x] 在 `Form1_ComL3.vb` 的 L2.5 層次中新增 `GetCachedMailWithAttachment` 處理 Phase1 快取。
- [x] **第二階段：淨化過濾引擎，解耦介面 UI (`Form1_Main.vb`)**
  - [x] 將 `ScanAttachmentDetail` 重命名為 `FilterByAttachmentDetailsAsync`，並修改其回傳為 `List(Of MailItemInfo)`。
  - [x] 將原先 `Button3_Click` 中的大小篩選邏輯抽出為獨立函數 `FilterBySize`。
  - [x] 將 `BuildListViewItem_Tab3` 收編擴展為 `MapToListViewItems_Tab3`，並在此實作決定傳回 ">0" 或明確 Count 的判定邏輯。
- [x] **第三階段：主控台 (Button3_Click) 管線化 (`Form1_Main.vb`)**
  - [x] 替換 `Button3_Click` 原有的複雜結構為純粹的 Pipeline 線性語法。
  - [x] 確保快取與所有的 Filter 正確串接執行。
- [x] **第四階段：註解與舊代碼清理**
  - [x] 檢查並保留重要演進歷史，清理殘餘冗餘的代碼塊。
