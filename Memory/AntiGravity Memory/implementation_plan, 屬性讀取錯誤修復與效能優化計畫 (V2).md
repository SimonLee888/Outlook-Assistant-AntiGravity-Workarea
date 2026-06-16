# 屬性讀取錯誤修復與效能優化計畫 (V2)

這個計畫旨在修復因 `PropertyAccessor.GetProperties` 回傳屬性型別不一致（PT_BINARY 回傳 `Byte()` 而非 `String`）導致強轉失敗並掉入 `Catch` 的問題，同時落實全域屬性標記宣告與 `IsMailFolder` 瓶頸優化。

## 修復與優化重點
> [!IMPORTANT]
> - **型別修正**：MAPI 屬性 `PR_ENTRYID` 與 `PR_STORE_ENTRYID` 是 `PT_BINARY` 格式，`PropertyAccessor` 會將其回傳為 `Byte()`。目前的代碼直接 `DirectCast(..., String)` 會觸發 `InvalidCastException`，導致每一圈都進入 `Catch` 並改走慢速路徑，這正是耗時增加的原因。
> - **標記全域化**：將標記 URL 抽離至類別頂端作為 `Private Const`。
> - **路徑完全優化**：修改 `IsMailFolder` 以接受外部傳入的 `fPath`，消除最後一個核心迴圈內的 COM 呼叫。

## 擬定變更

### Form1_Outlook.vb

#### [MODIFY] [Form1_Outlook.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/Form1_Outlook.vb)

- **類別頂端**：宣告 `PR_EID`, `PR_SID`, `PR_NAME`, `PR_CC` 等常數。
- **新增輔助函數**：`Private Function BinToHex(val As Object) As String`。
    - 判斷輸入：如果是 `Byte()`，則呼叫 `PropertyAccessor.BinaryToString` (或手工轉 16 進位)；如果是 `String` 直接回傳。
- **`IsMailFolder` 函數**：
    - 修改簽章：`Private Function IsMailFolder(folder As Outlook.Folder, Optional fPath As String = "") As Boolean`。
    - 優先使用傳入的 `fPath` 進行快取查找。
- **`GetSortedSubFolders` & `GetSubFolderList`**：
    - 使用 `BinToHex` 處理批次讀取的結果，消除轉型錯誤。
    - 呼叫 `IsMailFolder` 時傳入預先算好的路徑，達成「零冗餘 COM 呼叫」遍歷。

## 驗證計畫

### 自動化測試
- 再次執行掃描並觀察計時。
- 若修復成功，耗時應大幅下降（預期從 450ms 下降至 200~300ms 區間，視 PST 大小而定）。

### 手動驗證
- 驗證 `eID` 和 `sID` 是否正確（可與舊版產生的 SQLite 比對）。
- 確認 TreeView 展開功能正常。
