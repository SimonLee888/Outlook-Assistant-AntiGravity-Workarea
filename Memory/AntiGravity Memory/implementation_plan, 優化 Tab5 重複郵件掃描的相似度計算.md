# 優化 Tab5 重複郵件掃描的相似度計算

目前 Tab5 的重複郵件掃描在符合條件後，部分情況會預設將相似度設為 100% (1.0)。為了讓使用者能更精確地觀察郵件間的微小差異（即使 Message-ID 相同，主旨仍可能因為轉寄或系統處理而有細微變動），我們決定對所有比對結果皆執行相似度計算，不再預設 100%。

## 使用者評論與回饋要求
> [!IMPORTANT]
> 不論是 `MID:` (Message-ID) 還是 `FB:` (Fallback) 或是 Fuzzy 模式，一律呼叫 `JaccardSimilarity` 計算第一封與後續郵件的主旨相似度並顯示。
> 在「精確比對」模式下，即使相似度未達 100%，也不會將其從群組中剔除（維持原本的篩選強度），僅將計算結果反映在 UI 上。

## 擬議變更

### Form1_MainTab345.vb

#### [MODIFY] [Form1_MainTab345.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_MainTab345.vb)

修改 `RenderGroupDictToLv5` 函數：
- 移除「精確比對模式下預設 1.0」的邏輯。
- 統一使用 `JaccardSimilarity` 計算主旨相似度。
- 僅在 `isExact = False` (模糊比對) 時才套用 `sim < 0.6` 的過濾門檻。

```vbnet
' 修改片段預覽
For Each kvp In groupDict
    If kvp.Value.Count <= 1 Then Continue For

    Dim simScores As New List(Of Double)(512)
    Dim isValidGroup As Boolean = True
    Dim firstSubject As String = kvp.Value(0).Subject
    simScores.Add(1.0) ' 第一封永遠是 100%

    For i As Integer = 1 To kvp.Value.Count - 1
        ' 2026/05/06 by Gemini 3 Flash: 不論模式，一律計算精確相似度
        Dim sim As Double = JaccardSimilarity(firstSubject, kvp.Value(i).Subject)
        simScores.Add(sim)
        
        ' 僅在「模糊比對」模式下才進行門檻過濾
        If Not isExact AndAlso sim < 0.6 Then 
            isValidGroup = False 
            Exit For 
        End If
    Next
    ' ... 略 ...
```

## 驗證計畫

### 手動驗證
1. 執行 Outlook Assistant。
2. 切換至 Tab5。
3. 選取包含主旨相似但略有不同（例如經過不同轉寄路徑導致主旨微變）的資料夾。
4. 使用「精確比對」模式進行掃描。
5. 確認顯示的相似度不再全是 100%，而是根據 Jaccard 演算法算出的結果。
6. 切換至「模糊比對」模式，確認原本的功能運作正常且仍有 0.6 的過濾門檻。
