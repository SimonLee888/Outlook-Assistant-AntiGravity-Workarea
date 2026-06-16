# 優化 Outlook Store 排序效能計畫

這份計畫旨在加速 `Form1_Outlook.vb` 中 `GetSortedStores` 函數的執行速度，特別是針對大量 PST 檔掛載時的排序耗時。

## 使用者評論請求
這項變更涉及全域變數 `_pstStoreList` 的賦值邏輯，以及一個新的內部結構體 `StoreSortInfo`。

> [!IMPORTANT]
> 此優化將 `DisplayName` 的讀取從 O(N log N) 降至 O(N)，並將中文判定從 Regex/字串處理改為字元層級判定。雖然效能提升明顯，但如果您掛載的 Store 數量極少（如 1-3 個），體感可能不明顯。

## 擬議變更

### Form1_Outlook.vb (或全域定義區)

#### [NEW] StoreSortInfo 結構體
在 `Region "■ 01 全域宣告"` 內新增一個結構體，用於快取屬性：
```vbnet
Private Structure StoreSortInfo
    Dim StoreObj As Outlook.Store
    Dim DisplayName As String
    Dim HasChinese As Boolean
End Structure
```

---

### Form1_Outlook.vb

#### [MODIFY] [GetSortedStores](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_Outlook.vb#L259)

將原本的兩行 LINQ 替換為「屬性快取模式」：
1. **單次讀取屬性**：遍歷 `space.Stores` 時，一次性存入 `DisplayName` 到結構體。
2. **高效判定中文**：直接檢查字面編碼，避免進入 `TextHasChineseChar` 可能存在的複雜判定邏輯（如果該函數包含 Regex）。
3. **記憶體排序**：對結構體清單進行排序，最後再 `Select` 出原始物件。

**預計修改後的邏輯如下：**
```vbnet
Dim rawStores As Outlook.Stores = space.Stores
Dim infoList As New List(Of StoreSortInfo)
For Each st As Outlook.Store In rawStores
    Dim dn As String = st.DisplayName
    infoList.Add(New StoreSortInfo With {
        .StoreObj = st, 
        .DisplayName = dn, 
        .HasChinese = FastCheckChinese(dn)
    })
Next
' 排序後回傳
Return infoList.OrderBy(Function(i) If(i.HasChinese, 1, 0)).
                ThenBy(Function(i) i.DisplayName, StringComparer.OrdinalIgnoreCase).
                Select(Function(i) i.StoreObj).ToList()
```

## 開放性問題

> [!NOTE]
> 1. 您目前的 `TextHasChineseChar` 實作是否包含 Regex？如果是，改用 `AscW` 判定將會帶來數十倍的速度提升。
> 2. `space.Stores` 中的項目是否需要 `TryMarshalRelease`？通常 Store 物件由 Namespace 管理，但在大批次操作下，手動釋放 `rawStores` 集合物件是更安全的作法。

## 驗證計畫

### 手動驗證
- 在 `_dbg` 中記錄執行毫秒數（使用 `Stopwatch`），比對優化前後的差異。
- 觀察 TreeView 載入時是否仍能正確將中文標題排在前面。
