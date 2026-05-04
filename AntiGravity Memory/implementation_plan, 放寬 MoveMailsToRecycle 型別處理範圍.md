# 放寬 MoveMailsToRecycle 型別處理範圍

## 目的
目前 `MoveMailsToRecycle` 函數在取得 Outlook 項目時，固定將其轉型為 `Outlook.MailItem`。這導致 RSS 摘要（通常是 `PostItem`）、會議邀請（`MeetingItem`）或其他非標準郵件類型的項目在刪除時會發生轉型錯誤或失敗。

## 擬議變更

### [Outlook Assistant]

---

#### [MODIFY] [Form1_MainTab345.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_MainTab345.vb)

- 修改 `MoveMailsToRecycle` 函數：
    - 將 `mail` 變數型別改為 `Object`。
    - 移除明確的 `CType(..., Outlook.MailItem)` 轉型。
    - 檢查取得的物件是否具有 `Delete` 方法（通常所有 Outlook Item 都有此方法），或直接透過 `Object` 呼叫 `Delete()`。
    - 增加防錯邏輯，確保各種 Item 都能正常執行 `.Delete()`。

## 驗證計畫

### 手動驗證
1. 針對 RSS 摘要項目執行刪除。
2. 針對草稿、電子報項目執行刪除。
3. 確認這些項目能正確移至對應 Store 的「刪除郵件」資料夾。
