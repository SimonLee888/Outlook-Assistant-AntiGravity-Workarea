# 批次讀取優化實作計畫

這個計畫旨在實作 `GetSubFolderList` 的「批次讀取」優化。雖然 OOM 不支援資料夾的 `GetTable`，但我們可以使用 `PropertyAccessor.GetProperties` 一次讀取多個屬性，並結合路徑字串計算來消除最耗時的 `subFolder.FolderPath` 與 `subFolder.Name` 呼叫。

## 使用者評論要求
> [!IMPORTANT]
> - 實作真正的「屬性批次讀取」(Batch Property Fetching)。
> - 避免在迴圈內多次觸發 COM 物件屬性讀取。
> - 利用路徑拼接優化，消除 recursive 的 `FolderPath` 計算。

## 擬定變更

### Form1_Outlook.vb

#### [MODIFY] [Form1_Outlook.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/Form1_Outlook.vb)

- 在 `GetSubFolderList` 的 OOM 遍歷迴圈中：
    - 定義 MAPI 屬性標記陣列：`PR_ENTRYID`, `PR_STORE_ENTRYID`, `PR_DISPLAY_NAME`。
    - 使用 `subFolder.PropertyAccessor.GetProperties` 一次取得這三個屬性。
    - 利用父目錄已知的 `currentPath` 加上 `\` 與子目錄 `fName` 拼接成新的 `fPath`，徹底取代 `subFolder.FolderPath`。
    - 補遺：在 `GetSortedSubFolders` 中也同步套用此優化，確保全域一致。

## 驗證計畫

### 自動化測試
- 啟動程式，點選大型 PST 根目錄。
- 觀察 Debug 視窗中的耗時輸出，確認「子樹展開」速度是否有感提升。
- 檢查 SQLite 資料庫，確認儲存的 `folder_path`, `entry_id` 等資訊與之前一致。

### 手動驗證
- 點選包含中英文字元的資料夾，確認 `fPath` 拼接邏輯在各種字元下皆正常。
- 驗證「搜尋」功能（依賴 `folder_path` 快取）是否仍能正確定位資料夾。
