# 重構完成報告：大量郵件資料掃描優化

已經順利完成 [Form1.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/AntiGravityTest/Form1.vb) 中核心資料掃描邏輯的改寫！本次重構完全依循您規劃的 Fallback 鏈，保留了原本的安全性設計與註解，並在效能瓶頸處導入了極速的批次與平行存取方法。

## 變更摘要

1. **`GetFolderSize` 升級為 GetArray() 批次處理**
   - **原本**：使用 `Items.GetFirst/GetNext` 或 `GetTable()` 的 `GetNextRow()` 逐筆進行 COM 提取，對含有萬筆郵件的 PST 負擔龐大。
   - **優化**：改用 `GetTable().GetArray(1000)` 一次性將 1000 筆 `PR_MESSAGE_SIZE_EXTENDED` (0x0E080014) 取回到原生的 .NET 2D 陣列中，**徹底消滅了千百次的跨行程 COM 交談**，速度將有爆發性的提升。
   - **RDO 平行化**：在 Redemption (⓪ 階段)，充分利用其 MTA-Safe 特性，把原本單執行的 `rdoFolder.Items` 迭代安全包裹進 `Task.Run`，並運用 `Parallel.For` 達成多核同步加總 (Interlocked.Add)，壓榨所有硬體效能。

2. **`GetYearCountsForFolder` 與 `GetMonthCountsForYear` 的 GetArray() 極速重寫**
   - **原本**：呼叫 `folderItems.Restrict` 迴圈 30 次來計算每年的資料量，或 12 次來計算月份資料，這會在底層建立多個 COM Filters。
   - **優化**：加入 **GetTable() + GetArray()** 的首選路徑，直接提取 `PR_MESSAGE_DELIVERY_TIME` (0x0E060040) 屬性。
   - 將所有收集到的日期陣列在純 .NET 的記憶體空間中透過 `For` 迴圈解析 `DateTime.Year` / `Month` 並寫入 `ConcurrentDictionary`，比連續發送 30 個 `Restrict` 指令更為快速流暢。
   - **安全後援**：原本的 `Restrict` 與 `BuildFilterDateRangeTab2` 完全保留，並做為 GetTable 若因為某些系統資料夾報錯時的 `② Fallback` 安全備案。

3. **完整保留了您的註解及 fallback 架構**
   - 包含 L2/L3 中您對於 `_cancelRequested` ESC 中斷、Yield 的設計，以及 Redemption -> MAPI -> OOM 的Fallback鏈等，皆原封不動地保留。
   - 所有 AI 的改動及思路，皆加上了特別醒目的 `' ✅ Agent 補充:` 開頭註解，方便您隨時檢閱或調整。

## 驗證建議

經過這些重構後，原本會讓 UI 卡頓的 STA 線程負載應該已被大幅釋放。  
請您在 Visual Studio 18 (2026) 中建置專案並測試：
1. 觀察讀取裝滿數萬封郵件之大型資料夾 (PST) 的總 Size 讀取速度變化。
2. 隨機點擊 Tab2 觀看長條圖與 `ListView2` 渲染，確認年份/月份郵件分佈統計速度。
3. 測試在掃描途中狂點 ESC 或中斷，確認防呆機制仍正常運作。

若測試沒問題，我們未來還可以針對其他讀取密集的區塊套用相同的 `GetArray()` 理論。
