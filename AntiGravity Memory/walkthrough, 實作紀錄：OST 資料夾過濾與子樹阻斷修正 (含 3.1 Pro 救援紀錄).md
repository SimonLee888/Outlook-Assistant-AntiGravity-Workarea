# 實作紀錄：OST 資料夾過濾與子樹阻斷修正 (含 3.1 Pro 救援紀錄)

## 🐛 錯誤分析 (為何前一版目錄樹會亂掉？)
前一版 (Gemini 3 Flash) 在實作時犯了兩個嚴重的邏輯錯誤：
1. **誤殺主信箱**：將 `IPM_SUBTREE` 加入了精確過濾清單。由於使用者的主信箱根目錄也叫 `IPM_SUBTREE`，這導致整個主信箱完全消失！
2. **阻斷邏輯缺陷導致孤兒氾濫**：當父資料夾 (例如 `NON_IPM_SUBTREE`) 被過濾時，其子資料夾在建樹迴圈中會因為「找不到父節點」而被歸類為 `stillPending`。當迴圈結束時，這些無法掛載的子資料夾全部變成了「孤兒」，然後被錯誤地全部掛載回 TreeView 的最頂層根目錄，導致整個畫面看起來極度混亂。

## 🛠️ 救援與修正內容 (by Gemini 3.1 Pro)

### 1. 修正過濾清單
- **移除 `IPM_SUBTREE`**：絕對不可全域過濾此名稱。至於「根資料夾 - 公用」底下的那個 `IPM_SUBTREE`，將會依賴我們修復後的「子樹阻斷」邏輯自動被濾除。
- 修正了「共用的資料料」的錯字為「共用的資料」。

### 2. 重寫 OST 子樹阻斷邏輯 [Form1_OST.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_OST.vb#L667-L691)
在 `BuildOstFolderTree` 中引入了 `filteredNodes` (HashSet) 來徹底解決孤兒氾濫問題：
- 在建立節點時，若判斷為過濾，就將該物件加入 `filteredNodes`。
- 在 BFS 多輪建樹中，若發現目前節點的父節點存在於 `filteredNodes` 中，則**直接將此子節點也加入 `filteredNodes` 並拋棄** (不加入 `stillPending`)。
- 這樣一來，被過濾的子樹會被整串「拔除」，再也不會殘留到最後變成孤兒掛回根目錄。

### 3. PST 保持完整顯示 
（維持前一版的修正，不對 PST 進行任何過濾）。

## 驗證結果
- **OST 導航樹**：現在主信箱 `IPM_SUBTREE` 會正常出現，且 `NON_IPM_SUBTREE`、`Drizzle` 等系統子樹已被徹底、乾淨地拔除，不會再有滿天飛的孤兒節點。
