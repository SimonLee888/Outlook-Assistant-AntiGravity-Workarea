# 修正 Tab4/Tab5 SSD 快取路徑污染計畫

## 核心問題
目前 `_cacheFolderBasicMailInfos` 使用 `fPath | needTopic` 作為 Key。
1. **寫入污染**：`SaveFolderBasicMailInfosInner` 直接將此 Key 寫入資料庫的 `folder_path` 欄位。
2. **查詢失敗**：`DbGetFolderBasicMailInfos(fPath)` 使用純路徑查詢，無法匹配資料庫中帶有後綴的路徑，導致 SSD 快取失效。

## 擬定修正方案

### 1. [MODIFY] [Form1_SQLite2.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_SQLite2.vb)
- **`SaveFolderBasicMailInfosInner`**: 在遍歷 `_cacheFolderBasicMailInfos` 時，利用 `Split("|"c)(0)` 提取純路徑再寫入 DB。
- **`LoadFolderBasicMailInfosInner`**: 從 DB 讀回時，統一重建為 `fPath | True` 的 Key 格式存入記憶體（因為 DB 已存有 Topic 欄位，視為最高權限資料）。

### 2. [MODIFY] [Form1_Outlook.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_Outlook.vb)
- **`GetFolderBasicMailInfos`**: 
    - 調整 L2.5 的查詢優先序。
    - 當 `dbResult` 命中時，無論請求的是 `needTopic=True` 或 `False`，都應視為有效，因為 DB 存儲的是完整資料。

---

## 驗證計畫

### 自動化/手動測試
- **快取寫入檢查**：執行 Save Cache 後，檢查 `basic_maillist` 資料表中的 `folder_path` 是否為純路徑。
- **重啟載入檢查**：重啟程式後，點選 Tab4 資料夾，檢查 Debug Log 是否出現 `命中 SSD` 且未觸發 L3 掃描。
- **Tab5 相容性**：確認 Tab5 在無 Topic 需求時也能正確使用此快取。
