# 修改 Form1_MainTab345.vb 排序邏輯

將 `ShowLv4Result` 中手動判斷 `_lv4GroupSortByCount` 並分開處理排序邏輯的部分，改為使用單一的 `sortedGroups` 變數搭配 `If` 運算子（Ternary Operator 概念）進行簡潔賦值，以提升代碼可讀性與維護性。

## 任務清單

- [ ] 修改 `Form1_MainTab345.vb` 的 `ShowLv4Result` 函數
    - [ ] 將 `If _lv4GroupSortByCount Then ... Else ... End If` 區塊替換為 `sortedGroups = If(_lv4GroupSortByCount, ...)` 形式
- [ ] 複檢所有修改點確認正確、複檢修改點前後是否遺留多餘程式碼
