# Debug 訊息呈現優化計畫

由於目前的 `DebugForm` 已能自動抓取呼叫者函數名稱（Calling Method 欄位），原本在 `strA` 中重複包含函數名稱的做法會導致資訊冗餘。本計畫將優化資料呈現方式，並同步增強後台的「自動耗時計算」配對邏輯。

## User Review Required

> [!IMPORTANT]
> **訊息格式變更**：`Debug Message` 欄位將不再顯示 `開始: MethodName()`，而是改為更簡潔的 `開始 [MethodName] (參數)`。這會讓搜尋與閱讀更佳直覺。
> **配對邏輯調整**：為了讓 `開始` 與 `結束` 能在參數不同（例如開始顯示「名稱」，結束顯示「數量」）的情況下依然正確配對並計算總時長，我將調整 `RemoveBeginEnd` 邏輯，使其在「比對核心」時忽略括號內的變動參數。

## Proposed Changes

### DebugForm 核心優化

#### [MODIFY] [DebugForm.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/DebugForm.vb)
*   **優化 `RemoveBeginEnd`**：
    *   在產生「比對用核心字串」時，除了移除行號與「開始/結束」標籤外，進一步移除結尾的整個括號區塊 `(...)`。
    *   這能確保即使開始訊息是 `開始 MethodA (Folder1)`，結束訊息是 `結束 MethodA (Count: 10)`，系統仍能識別出它們是同一對，從而正確計算總耗時。

---

### Form1 訊息格式重構 (L2430-L2721 範疇)

#### [MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1.vb)
將目標區域（Tab2 統計核心邏輯）的 `Dbg` 呼叫進行精簡：
*   **ComputeYearCounts**:
    *   開始：`Dbg("開始", $"資料夾數: {folderList.Count}")`
    *   結束：`Dbg("結束", $"年份數: {merged.Count}, 總計: {merged.Values.Sum}")`
*   **GetYearCountsForFolder** / **GetMonthCountsForYear**:
    *   開始：`Dbg("開始", folder.Name)`
    *   結束：`Dbg("結束", $"結果: {count} 筆")`
*   **ShowYearView** / **ShowMonthView**:
    *   同步清理掉字串中多餘的 `()` 或重複的函數名。

## Open Questions

1. **配對關鍵字**：您目前使用的配對標籤是「開始: 」與「結束: 」。我計畫將其簡化為「開始」與「結束」（移除冒號），並讓邏輯同時相容兩者。是否同意？

## Verification Plan

### 自動化/手動驗證
- [ ] **配對測試**：確認調整為新格式後，雙擊「結束」行依然能精確找到對應的「開始」行，且 `Time Span` 欄位能正確顯示總耗時。
- [ ] **視覺檢視**：觀察 `DebugForm` 列表，確認資訊不再擁擠且重點突出。
