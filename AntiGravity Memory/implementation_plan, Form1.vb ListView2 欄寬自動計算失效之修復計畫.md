# Form1.vb ListView2 欄寬自動計算失效之修復計畫 (優化方案)

本計畫針對 `ListView2`（Tab2 依日期統計）之自動欄寬計算失效問題進行更精準且低開銷的修正。

## 根本原因分析

1. **防線字典過早記錄（核心 Bug 原因）**：
   在 `Form1.vb` 的 `CalculateLvColumnSize()` 中，原先的寬度防護線設計如下：
   ```vb
   Static lastProcessedWidths As New Dictionary(Of String, Integer)
   If lastProcessedWidths.ContainsKey(lv.Name) AndAlso lastProcessedWidths(lv.Name) = lv.Width Then Return
   lastProcessedWidths(lv.Name) = lv.Width ' ⚠️ 尚未計算，就先記錄當前寬度！
   ```
   在 Form 剛啟動或 `ListView2` 初始化時，`CalculateLvColumnSize` 會被觸發。此時雖然 `lv.Width` 有值，但因為 `ListView2` 尚未在畫面上實質渲染（位於非活動 Tab），其 `lv.ClientSize.Width`（`w`）拿到的值是 `0`（或無效值）。
   此時程式碼執行到後半段的 `If w <= 0 Then Exit Try`，直接跳出了計算與賦值過程，導致**欄寬並未被正確調整**。
   但由於前面已經將 `lastProcessedWidths("ListView2") = lv.Width` 寫入，導致之後不論是切換 Tab 還是載入資料，只要 `ListView2` 的寬度沒有發生改變，這條優化防線就會永遠在最上方直接 `Return` 攔截，使得自動欄寬計算在「有了正確 ClientSize.Width 後」也永遠無法再執行一次！

## 優化解決方案

我們完全**不需要**在每次切換 Tabpage 或每次 Render 資料時重複呼叫，這樣確實太浪費開銷。
最優雅且無額外開銷的方案是：
**將 `lastProcessedWidths(lv.Name) = lv.Width` 的寫入時機，移到「成功計算並套用寬度之後」！**

這樣一來：
1. 最初初始化 `w <= 0` 失敗時，防線字典**不會**記錄該寬度。
2. 當使用者第一次切換到該 Tabpage，`ListView2` 擁有了正確的實質寬度，並在觸發 Resize 或重算時，因為防線字典中尚無此寬度的成功記錄，將會**成功執行且僅執行這一次精準的欄寬調整**，並寫入防線記錄。
3. 此後，不論是切換 Tab 還是填入資料，只要寬度不變，防線字典就會在最上方直接 `Return` 攔截，**完全不會產生任何重複計算的效能浪費**。

## User Review Required

> [!NOTE]
> 本方案完全不增加額外的定時器或無意義的重複呼叫，僅修正防線字典的寫入時序，是最符合 WinForms 佈局機制的低開銷解法。

## Proposed Changes

### 1. Form1.vb Core

#### [MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual Studio/Visual Studio 18 (2026)/Outlook Assistant - (AntiGravity測試區)/Form1.vb)
- **調整 `CalculateLvColumnSize` 函數內部防線寫入時機**：
  將最前方的 `lastProcessedWidths(lv.Name) = lv.Width` 移除，並改移至 `Try` 區塊內批量賦值迴圈結束之後（`For i As Integer = 0 To lv.Columns.Count - 1 ... Next` 之後）。

---

## Verification Plan

### Manual Verification
1. 啟動 Outlook Assistant 程式。
2. 點選「2.依日期統計」Tab，確認切換分頁後，`ListView2` 的欄位（年份、郵件數量、空白欄位）寬度已**自動且正確**依比例縮放（年份與郵件數量各有合理的寬度，且第三欄吸收了剩餘空間）。
3. 載入統計數據或切換年度/月份視圖，確認欄寬維持正確，且在拖動或縮放視窗時，自動調整欄寬邏輯依然正常運作。
