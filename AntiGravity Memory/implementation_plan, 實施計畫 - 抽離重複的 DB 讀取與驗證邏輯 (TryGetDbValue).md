# 實施計畫 - 抽離重複的 DB 讀取與驗證邏輯 (TryGetDbValue)

本計畫旨在重構 `Form1_Outlook.vb` 中多次重複出現的「讀取 DB 統計、驗證 Snapshot、回填快取並回傳數值」的程式碼結構。我們將建立一個通用的泛型輔助函式 `TryGetDbValue(Of T)` 來簡化這些 Layer 2.5 快取代理函式。

## 使用者評論

> [!IMPORTANT]
> 此重構將影響多個核心計數函式（如 `GetMailCount`, `GetFolderCount`, `GetFolderSizeAsync` 等）。我們將使用泛型與 Lambda 選取器（fieldSelector）來處理不同的欄位名稱與資料型別（Integer/Long），確保程式碼的一致性與可維護性。

## 開放問題

> [!NOTE]
> `HasSubFoldersFast` 目前的實作略過了 Snapshot 驗證（推測是為了極速回應，避免一次 PropertyAccessor 呼叫）。在重構時，我會為 `TryGetDbValue` 加入一個可選參數 `checkSnap`，並在 `HasSubFoldersFast` 中將其設為 `False`，以維持既有的效能特性。

## 擬議變更

### Form1_Outlook.vb [MODIFY]

#### 1. [NEW] 新增 `TryGetDbValue(Of T)` 輔助函式
在 `Layer2.5 快取存取點` 區塊頂部新增此函式。它將封裝：
- 呼叫 `DbGetFolderStats(fPath)`。
- 檢查 `row` 是否存在。
- 根據 `checkSnap` 參數執行 Snapshot 驗證（利用 `GetLiveFolderSnapL3`）。
- 使用傳入的 `fieldSelector` Lambda 取得特定欄位值。
- 驗證數值是否有效（針對數值型別檢查 `>= 0`）。
- 驗證通過後，主動呼叫 `FillFolderCacheFromDbRow(fPath, row)` 以「順便」填滿記憶體中的其他統計欄位。
- 回傳結果。

#### 2. [MODIFY] 重構現有 L2.5 函式
將以下函式內部的「DB 讀取」區塊替換為 `TryGetDbValue` 呼叫：
- `GetMailCount`: 選取 `row.mc` (Integer)
- `GetFolderCount`: 選取 `row.fc` (Integer)
- `GetFolderSizeAsync`: 選取 `row.fs` (Long)
- `GetFolderSizeAllAsync`: 選取 `row.fsa` (Long)
- `GetMailCountAllAsync`: 選取 `row.mca` (Integer)
- `GetFolderCountAllAsync`: 選取 `row.fca` (Integer)
- `HasSubFoldersFast`: 選取 `row.fc` (Integer)，並設定 `checkSnap:=False`

## 驗證計畫

### 自動化測試
- 使用 `view_file` 複核重構後的邏輯是否與原邏輯對齊，特別是 `checkSnap` 的預設行為。

### 手動驗證
1. **資料準確性**：點選不同資料夾，確認 Tab1 的郵件數與資料夾數正確從 DB 載入且與 Outlook 顯示一致。
2. **快取效率**：觀察 `_dbg` 輸出，確認 `DbGetFolderStats` 在第二次點擊（快取未命中但 DB 有資料）時正確被呼叫且成功回傳。
3. **Snapshot 失效驗證**：在 Outlook 手動刪除郵件後（改變 Snapshot），點選該資料夾，確認 Helper 能正確偵測 Snapshot 不符並 Fallback 到 Layer3 重新計算。
4. **TreeView 相容性**：確認 TreeView 展開時，子資料夾的 `+` 號顯示邏輯不受影響。

---

## 程式碼對比預覽 (修改前 vs 修改後)

### 修改前 (以 GetMailCount 為例)
```vbnet
    Private Function GetMailCount(folder As Folder, Optional fPath As String = "") As Integer
        fPath = SafeGetPath(folder, fPath)
        Dim count As Integer
        If _cacheMailCount.TryGetValue(fPath, count) Then Return count

        ' 重複的結構 -------------------------------------------
        Dim row = DbGetFolderStats(fPath)
        If row IsNot Nothing AndAlso row.mc >= 0 AndAlso GetLiveFolderSnapL3(folder) = row.snap Then
            FillFolderCacheFromDbRow(fPath, row) : Return row.mc
        End If
        ' ------------------------------------------------------

        count = GetMailCountL3(folder, fPath:=fPath)
        _cacheMailCount.TryAdd(fPath, count)
        Return count
    End Function
```

### 修改後 (通用 Helper + 精簡後的 GetMailCount)
```vbnet
    ' --- 新增的 Helper ---
    Private Function TryGetDbValue(Of T)(folder As Folder, fPath As String, fieldSelector As Func(Of FolderStatsDbRow, T), ByRef value As T, Optional checkSnap As Boolean = True) As Boolean
        Dim row = DbGetFolderStats(fPath)
        If row IsNot Nothing Then
            ' 驗證 Snapshot (若需要)
            If checkSnap AndAlso GetLiveFolderSnapL3(folder) <> row.snap Then Return False

            Dim val = fieldSelector(row)
            ' 數值有效性檢查 (Integer/Long >= 0)
            Dim isValid As Boolean = False
            If TypeOf val Is Integer Then
                isValid = (CInt(DirectCast(val, Object)) >= 0)
            ElseIf TypeOf val Is Long Then
                isValid = (CLng(DirectCast(val, Object)) >= 0)
            End If

            If isValid Then
                FillFolderCacheFromDbRow(fPath, row)
                value = val : Return True
            End If
        End If
        Return False
    End Function

    ' --- 重構後的 GetMailCount ---
    Private Function GetMailCount(folder As Folder, Optional fPath As String = "") As Integer
        fPath = SafeGetPath(folder, fPath)
        Dim count As Integer
        If _cacheMailCount.TryGetValue(fPath, count) Then Return count

        ' 呼叫 Helper，一行搞定
        If TryGetDbValue(folder, fPath, Function(r) r.mc, count) Then Return count

        count = GetMailCountL3(folder, fPath:=fPath)
        _cacheMailCount.TryAdd(fPath, count)
        Return count
    End Function
```
