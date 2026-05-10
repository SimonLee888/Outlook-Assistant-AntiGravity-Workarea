# 實施計畫 - 統一型別並精簡重構 (TryGetDbLong)

根據使用者的建議，我們將採取更徹底的重構方案：統一使用 `Long` 型別來簡化 Helper 函式的實作，並提高系統的穩健性。

## 使用者回饋與調整

> [!TIP]
> **型別統一化 (Long)**：改用 `Long` 後，Helper 函式不再需要 `TypeOf` 判斷，程式碼邏輯將簡化一半以上。雖然 `Integer` (32-bit) 對多數單一資料夾已足夠，但統一使用 `Long` 與資料夾大小 (Size) 對齊，且能避免子樹加總時的溢位風險。
> 
> **排除 `HasSubFoldersFast`**：該函式具有特殊的「極速預測」需求（略過 Snapshot 驗證且回傳 Boolean）。為了保持結構單純，我們將其排除在 Helper 套用範圍之外，維持其既有的效能特性。

## 擬議變更

### 1. Form1_Outlook.vb [MODIFY]

#### 調整快取字典型別
將下列 Shared 字典的型別從 `Integer` 改為 `Long`：
- `_cacheMailCount`
- `_cacheMailCountAll`
- `_cacheFolderCount`
- `_cacheFolderCountAll`

#### 新增 `TryGetDbLong` 輔助函式
```vbnet
    ''' <summary>
    ''' [L2.5 輔助函式] 從 DB 讀取數值並驗證。
    ''' 僅限 Long 型別欄位，且一律執行 Snapshot 驗證。
    ''' </summary>
    Private Function TryGetDbLong(folder As Folder, fPath As String, fieldSelector As Func(Of FolderStatsDbRow, Long), ByRef value As Long) As Boolean
        Dim row = DbGetFolderStats(fPath)
        ' ① row 不為空 且 ② Snapshot 吻合 (代表資料夾內容沒變)
        If row IsNot Nothing AndAlso GetLiveFolderSnapL3(folder) = row.snap Then
            Dim val = fieldSelector(row)
            If val >= 0 Then ' 數值有效性檢查 (所有欄位已統一為 Long)
                FillFolderCacheFromDbRow(fPath, row)
                value = val : Return True
            End If
        End If
        Return False
    End Function
```

#### 重構 L2.5 函式
更新簽章傳回值為 `Long` 並套用 Helper：
- `GetMailCount`
- `GetFolderCount`
- `GetMailCountAllAsync`
- `GetFolderCountAllAsync`
- `GetFolderSizeAsync` (已是 Long)
- `GetFolderSizeAllAsync` (已是 Long)

### 2. Form1_MainTab12.vb [MODIFY]
- **更新 `FolderBfsEntry` 類別**：將 `DirectMailCount`, `TotalMailCount`, `TotalSubCount` 欄位由 `Integer` 改為 `Long`。

### 3. Form1_SQLite2.vb [MODIFY]
- **更新 `FolderStatsDbRow` 類別**：將 `mc`, `mca`, `fc`, `fca`, `snap` 欄位由 `Integer` 改為 `Long`。
- **更新 `DbGetFolderStats` 讀取邏輯**：將 `reader.GetInt32` 調整為 `reader.GetInt64`。

## 驗證計畫

### 自動化測試
- 使用 `view_file` 確保所有數值計數函式的簽章均已正確更新為 `Long`。
- 檢查 `TryGetDbLong` 是否正確呼叫了 `FillFolderCacheFromDbRow`。

### 手動驗證
1. **數值顯示**：確認各分頁中的郵件數與資料夾數顯示正確，無遺失或錯誤格式。
2. **快取一致性**：在 Outlook 中改變郵件數量，確認 Helper 能偵測 Snapshot 不符並 Fallback。
3. **效能檢查**：確認 `HasSubFoldersFast` 的回應速度未受影響。
