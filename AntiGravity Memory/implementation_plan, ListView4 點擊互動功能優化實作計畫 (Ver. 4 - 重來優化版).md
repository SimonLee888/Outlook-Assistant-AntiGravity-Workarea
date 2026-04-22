# ListView4 點擊互動功能優化實作計畫 (Ver. 4 - 重來優化版)

本計畫旨在移除原本不理想的 Timer 機制，改為更直覺的單擊顯示 ToolTip (路徑) 並同步至 ProgressBar2.Text，同時實作與 ListView3 一致的雙擊開啟郵件功能。

## 使用者評論與回饋要求
> [!IMPORTANT]
> 1. **移除 Timer**: 原本加在 ListView4 用於延遲顯示 ToolTip 的 Timer 將被徹底移除。
> 2. **單擊互動**: 滑鼠單點 ListView4 項目時，立即顯示 ToolTip (內容為資料夾路徑)，並將該路徑同步顯示在 `ProgressBar2.Text`。
> 3. **雙擊互動**: 雙擊 ListView4 項目，將呼叫 `OpenMailByEntryID` 開啟該郵件，行為模式參考 ListView3。

## 擬定變更

### [Component] Form1_MainTabs.vb

#### [MODIFY] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_MainTabs.vb)

- **移除 Timer 相關程式碼**: 如果檔案內仍殘留針對 ListView4 的 Timer 宣告或 Tick 事件，將予以移除。
- **修改 `ListView4_MouseClick`**:
  - 取得點擊的 `ListViewItem`。
  - 從 `TreeView4.SelectedNode.Tag` (List(Of MailItemInfo)) 中找出對應的 `MailItemInfo` (利用 Index)。
  - 取得 `MailItemInfo.FolderPath`。
  - 將路徑顯示在 `ToolTip1` 並指向 ListView4。
  - 將路徑設定給 `ProgressBar2.Text`。
  - 保留原有的「複製主旨到剪貼簿」功能。
- **修改 `ListView4_MouseDoubleClick`**:
  - 取得點擊點的 `EntryID`。
  - 呼叫 `OpenMailByEntryID(New List(Of String) From {entryID})`。

### [Component] Form1_Outlook.vb (已確認)

#### [MODIFY] [Form1_Outlook.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_Outlook.vb)
- **確認 `MailItemInfo` 結構**: 確保 `FolderPath` 欄位已存在 (目前已確認存在)。

## 驗證計畫

### 手動測試 (請使用者驗證)
1. 在 Tab4 搜尋系列郵件。
2. 點選 TreeView4 項目，展開 ListView4 列表。
3. **單擊測試**: 點擊 ListView4 中任一封郵件，確認是否出現 ToolTip 顯示路徑，且 ProgressBar2.Text 是否同步顯示路徑。
4. **雙擊測試**: 雙擊 ListView4 項目，確認是否正確在 Outlook 中開啟該封郵件。
5. **移除驗證**: 確認滑鼠停在項目上不再有 2 秒延遲的 Timer 觸發行為。
