# 實作計畫：全域套用 TryMarshalRelease() 優化

## 1. 目標
將專案中所有手動呼叫 `Marshal.ReleaseComObject` 的位置，統一替換為您剛才優化過的 `TryMarshalRelease()` 輔助函數。

## 2. 修改前後對照表 (Before vs. After)

### 場景 A：標準條件釋放 (多處)
這類修改能將三行代碼縮減為一行，且更安全。

| 檔案:行號 | 修改前 (Before) | 修改後 (After) |
| :--- | :--- | :--- |
| Form1.vb:1195 | `If _rdo IsNot Nothing Then Marshal.ReleaseComObject(_rdo)` | `TryMarshalRelease(_rdo)` |
| Form1.vb:2546 | `If table IsNot Nothing Then Marshal.ReleaseComObject(table)` | `TryMarshalRelease(table)` |
| Form1_ComL3:104 | `If rdoFolder IsNot Nothing Then Marshal.ReleaseComObject(rdoFolder)` | `TryMarshalRelease(rdoFolder)` |

### 場景 B：循環內的頻繁釋放 (效能關鍵點)
確保在大量資料遍歷時，物件指標能正確歸零。

| 檔案:行號 | 修改前 (Before) | 修改後 (After) |
| :--- | :--- | :--- |
| Form1_ComL3:1034 | `Finally : Marshal.ReleaseComObject(row)` | `Finally : TryMarshalRelease(row)` |
| Form1_ComL3:517 | `Marshal.ReleaseComObject(row)` | `TryMarshalRelease(row)` |

### 場景 C：多重物件連續釋放
提升代碼的可讀性，減少 `If` 嵌套。

| 檔案:行號 | 修改前 (Before) | 修改後 (After) |
| :--- | :--- | :--- |
| Form1.vb:2968-2970 | `If mail IsNot Nothing Then Marshal.ReleaseComObject(mail)`<br>`If validItems IsNot Nothing Then Marshal.ReleaseComObject(validItems)`<br>`If allItems IsNot Nothing Then Marshal.ReleaseComObject(allItems)` | `TryMarshalRelease(mail)`<br>`TryMarshalRelease(validItems)`<br>`TryMarshalRelease(allItems)` |
| Form1.vb:3902-3904 | `Marshal.ReleaseComObject(inbox)`<br>`Marshal.ReleaseComObject(ns)`<br>`Marshal.ReleaseComObject(outlookApp)` | `TryMarshalRelease(inbox)`<br>`TryMarshalRelease(ns)`<br>`TryMarshalRelease(outlookApp)` |

---

## 3. 擬議的變更清單

### [Component] UI 邏輯層 (Form1.vb)
- **範圍**：全文約 15 處修改。
- **重點**：`CheckRDO_CheckedChanged`、`ListView` 數據載入、郵件處理循環。

### [Component] 底層資料層 (Form1_ComL3.vb)
- **範圍**：全文約 25 處修改。
- **重點**：`GetMailCount`、`GetFolderSizeLegacy`、所有 RDO 相關的 Try-Finally 塊。

---

## 4. 驗證計畫
1. **編譯檢查**：確保所有修改均符合語法。
2. **資源監控**：開啟 Outlook 進程監控，確保在切換分頁或執行大規模統計後，記憶體占用能正常回落。
3. **穩定性測試**：在 `Debug` 頁面反覆開關 Redemption 以及切換分頁，確認 `TryMarshalRelease` 能正確捕捉並紀錄潛在的釋放異常。

**請查閱以上對照表。如果您同意，請回覆「確認開始」，我將執行全域替換。**
