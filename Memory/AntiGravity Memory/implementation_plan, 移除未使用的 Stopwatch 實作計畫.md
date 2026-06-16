# 移除未使用的 Stopwatch 實作計畫

檢查專案中所有 `Stopwatch` 的用法，找出並移除宣告了但未停止、未顯示或未被引用的物件。

## 使用者評論要求
- ** moduleStore.vb**: 該檔案包含大量歷史測試程式碼，多數 `Stopwatch` 位於註解區塊。我將移除無用的宣告，但盡量保留註解中的邏輯說明以符合「保留歷程紀錄」的規則。
- ** Form1_DebugForm.vb**: 移除 7 個完全未使用的私有計時器變數。

## 擬議變更

### [Outlook Assistant]

---

#### [MODIFY] [Form1_DebugForm.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/Form1_DebugForm.vb)
- 移除第 76 行未使用的私有變數宣告：`Private sw0, sw1, sw2, sw3, sw4, sw5, sw6 As New Stopwatch`。

#### [MODIFY] [moduleStore.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/moduleStore.vb)
- 移除多處僅存在於註解中或已無實際作用的 `Stopwatch` 宣告與起始指令，包括：
    - `Private sw0 As Object` (第 12 行)
    - `tmrPreCache_Tick` 註解區塊內的 `swa`, `swb` (第 16 行)
    - 其他散落在檔案各處被註解掉的 `swa`, `swb`, `stopwatch` 等。

## 驗證計畫

### 自動測試
- 執行建置，確保移除變數宣告後沒有引發編譯錯誤。

### 手動驗證
- 檢查剩餘的 `Stopwatch` 用法，確認皆有對應的 `.Stop()` 呼叫或用於 `SmartThrottle` 節流邏輯中。
