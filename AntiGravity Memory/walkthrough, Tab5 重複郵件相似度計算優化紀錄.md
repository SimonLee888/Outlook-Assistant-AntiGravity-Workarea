# Tab5 重複郵件相似度計算優化紀錄

本次修改優化了 Tab5 在「精確比對」模式下的資訊回饋，讓使用者能觀察到郵件間細微的主旨差異。

## 變更內容

### 1. 統一相似度計算邏輯
在 `Form1_MainTab345.vb` 的 `RenderGroupDictToLv5` 函數中，我們移除了對 `isExact` 模式的特殊對待。現在，不論是透過 Message-ID 比對還是 Fallback 比對，系統都會：
- 呼叫 `JaccardSimilarity` 計算第一封郵件與後續郵件的主旨相似度。
- 將結果顯示在 ListView5 的「相似度」欄位中。

### 2. 精確度與顯示分離
- **精確比對模式 (Exact)**：計算出的相似度僅供 UI 顯示參考。即使相似度低於 0.6，也不會將郵件從群組中移除，確保掃描結果的嚴謹性。
- **模糊比對模式 (Fuzzy)**：維持原有的 0.6 相似度過濾門檻，自動過濾掉相似度過低的假陽性結果。

## 驗證結果

### 程式碼複檢
- [x] 確認 `simScores` 列表正確收集了所有郵件的計算結果。
- [x] 確認 `If Not isExact AndAlso sim < 0.6` 邏輯正確將過濾門檻限制在模糊模式。
- [x] 確認 `simText` 格式化邏輯 `CInt(simScores(idx) * 100)%` 運作正常。
- [x] 確認修改點前後無遺留多餘程式碼。

### 關鍵程式碼片段

```vbnet
' Form1_MainTab345.vb:L1536-1542
For i As Integer = 1 To kvp.Value.Count - 1
    Dim sim As Double = JaccardSimilarity(firstSubject, kvp.Value(i).Subject)
    simScores.Add(sim)
    
    ' 僅在模糊模式下才套用門檻過濾 (0.6)
    If Not isExact AndAlso sim < 0.6 Then isValidGroup = False : Exit For
Next
```
