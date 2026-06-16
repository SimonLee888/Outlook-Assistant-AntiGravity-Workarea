# 重構 RDO 資料夾獲取結構

在 `Form1_Outlook.vb` 中，從 Outlook `Folder` 物件轉換為 Redemption `RDOFolder` 的過程中，存在大量重複的結構（`Try...Catch...Finally` 搭配 `TryMarshalRelease`）。這不僅使代碼顯得冗長，也增加了漏掉釋放物件的風險。

## 使用者評論請求
> [!IMPORTANT]
> 本次重構會將多處（約 6-8 處）直接呼叫 `_rdo.GetFolderFromID` 的地方改為呼叫新的 `GetRdoFolderSafe` Helper 函數。這是一個純粹的結構優化，不應影響原有邏輯。

## 提議的變更

### Form1_Outlook.vb

#### [MODIFY] [Form1_Outlook.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_Outlook.vb)

1. **新增 Helper 函數**: 在檔案末尾的 Helper 區段新增 `GetRdoFolderSafe`。
2. **替換重複結構**: 搜尋並替換以下方法中的重複 RDO 獲取區塊：
   - `GetMailCountL3`
   - `GetFolderCountL3`
   - `GetFolderSizeL3`
   - `GetFolderSizeAllL3`
   - `GetMailCountAllL3`
   - `GetFolderCountAllL3`
   - `RdoPreloadAttach_1` & `RdoPreloadAttach_2` (獲取 Message 部分暫不更動，專注於 Folder)

---

### 新增的 Helper 函數預覽
```vb
    ''' <summary>
    ''' 安全地從 Outlook Folder 獲取 RDOFolder，並處理錯誤與 _rdo 檢查
    ''' </summary>
    Private Function GetRdoFolderSafe(folder As Folder, Optional logPrefix As String = "") As Redemption.RDOFolder
        ' by Gemini 3.1 Pro, 2026/05/09: 抽離重複的 RDO 獲取邏輯，確保 TryMarshalRelease 之前的獲取是安全的
        If _rdo Is Nothing OrElse folder Is Nothing Then Return Nothing
        Try
            Return _rdo.GetFolderFromID(folder.EntryID, folder.StoreID)
        Catch ex As System.Exception
            If _iLikeNoisy Then _dbg("    ├ RDO 獲取失敗", $"{logPrefix} | {folder.Name} | {ex.Message}")
            Return Nothing
        End Try
    End Function
```

### 修改後的結構示例 (以 GetMailCountL3 為例)
```vb
        ' ⓪ Redemption: RDOFolder.Items.Count
        Dim rdoFolder As Redemption.RDOFolder = GetRdoFolderSafe(folder, "GetMailCountL3")
        If rdoFolder IsNot Nothing Then
            Try
                Return CLng(rdoFolder.Items.Count)
            Catch ex As System.Exception
                If _iLikeNoisy Then _dbg("    ├ 錯誤路徑", $"GetMailCount ⓪ RDO: {fName} | {ex.Message}")
            Finally
                TryMarshalRelease(rdoFolder)
            End Try
        End If
```

## 驗證計畫

### 自動測試
- 編譯專案，確保無語法錯誤。

### 手動驗證
- 執行應用程式，測試 Tab 1 (資料夾計數)、Tab 2 (大小計算) 等功能，確認數據顯示正確。
- 檢查 Debug 視窗輸出，確認 RDO 獲取失敗時有正確記錄（若有發生）。
