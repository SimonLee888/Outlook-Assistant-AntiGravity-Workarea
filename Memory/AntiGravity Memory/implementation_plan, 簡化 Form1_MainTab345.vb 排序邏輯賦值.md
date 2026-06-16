# 簡化 Form1_MainTab345.vb 排序邏輯賦值

此修改將 `ShowLv4Result` 函數中的群組排序邏輯從傳統的 `If...Then...Else` 結構重構為更簡潔的 `If` 運算子賦值方式，符合使用者要求的代碼風格。

## 提議的變更

### [Form1_MainTab345.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_MainTab345.vb)

#### [MODIFY] [Form1_MainTab345.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_MainTab345.vb)

將第 1017 行至 1023 行的代碼：
```vb
        If _lv4GroupSortByCount Then
            ' 模式：按組內數量遞減排序
            sortedGroups = groups.OrderByDescending(Function(g) g.Count()).ThenBy(Function(g) g.Key)
        Else
            ' 模式：按主旨字母順序排序
            sortedGroups = groups.OrderBy(Function(g) g.Key)
        End If
```
替換為：
```vb
        ' 2. 依照排序模式對「組」進行排序 (by Gemini 3.0 flash, 2026/05/11)
        Dim sortedGroups = If(_lv4GroupSortByCount,
                             groups.OrderByDescending(Function(g) g.Count()).ThenBy(Function(g) g.Key),
                             groups.OrderBy(Function(g) g.Key))
```

## 驗證計畫

### 手動驗證
- 在 Tab4 中執行搜尋後，觀察右側 `ListView4` 的分組排序是否仍能正確根據 `_lv4GroupSortByCount` 進行切換（數量遞減 vs 主旨字母順序）。
