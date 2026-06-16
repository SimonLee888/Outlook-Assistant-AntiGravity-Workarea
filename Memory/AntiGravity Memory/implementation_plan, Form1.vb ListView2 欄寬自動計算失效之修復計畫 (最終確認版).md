# Form1.vb ListView2 欄寬自動計算失效之修復計畫 (最終確認版)

本計畫針對 `ListView2`（Tab2 依日期統計）之自動欄寬計算失效問題進行修正。經審視與討論，已確認先前的「時序型 Bug」分析完全正確，並根據使用者「不應無謂浪費開銷」的原則，去除了多餘的觸發點。

## 根本原因與優化分析

1. **防線過早記錄 (核心 Bug)**：
   原有的 `CalculateLvColumnSize` 在剛進入函式時就立刻記錄了 `lastProcessedWidths(lv.Name) = lv.Width`。但在程式啟動背景載入時，ListView 尚未繪製，`ClientSize.Width` 為 0，導致計算跳出。結果是：**欄寬沒有更新，但系統卻記住了「已經處理過此寬度」**。往後不論怎麼切換，只要 Width 不變，就永遠不會再計算。
2. **零開銷的解決策略**：
   正如使用者所指出的，欄寬比例只跟「容器寬度」有關，跟「載入什麼資料 (Render)」毫無關聯。因此，在 `RenderLv2` 中反覆呼叫不僅沒必要，還浪費開銷。
   最適當的時機是：**在 ListView 第一次正確顯示、且有了實質 ClientSize 的時候呼叫一次。**

## Proposed Changes

### 1. Form1.vb Core

#### [MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual Studio/Visual Studio 18 (2026)/Outlook Assistant - (AntiGravity測試區)/Form1.vb)
- **延後防線字典寫入時機**：
  在 `CalculateLvColumnSize` 內部，將 `lastProcessedWidths(lv.Name) = lv.Width` 移到 `Try...Finally` 迴圈結束、成功賦值之後。如果 `ClientSize <= 0` 而提早 `Exit Try`，就不會被記錄，避免鎖死防線。
- **於 Tab 切換時觸發 (零開銷)**：
  在 `TabControl1_SelectedIndexChanged` 函式結尾處，加上 `Dim currentLv = GetCurrentLv() : If currentLv IsNot Nothing Then CalculateLvColumnSize(currentLv)`。
  - 當第一次切換到該 Tab，會精準計算一次並記錄。
  - 後續切換回該 Tab，由於寬度未變，防線會在第一行直接 Return，**開銷趨近於零**。

此方案完美達成了「正確初始化後只在需要時計算」，完全不浪費效能。

---

## Verification Plan

### Manual Verification
1. 啟動 Outlook Assistant。
2. 切換至「2.依日期統計」Tab，確認 ListView2 的欄寬立即正確依比例展開。
3. 拖曳縮放視窗，確認欄寬跟隨調整。
4. 載入資料並在年度與月份視圖間切換，確認欄寬維持正常且不再產生多餘計算開銷。
