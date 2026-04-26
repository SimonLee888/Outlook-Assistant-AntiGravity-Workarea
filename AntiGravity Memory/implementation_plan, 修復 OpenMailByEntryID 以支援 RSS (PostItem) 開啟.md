# 修復 OpenMailByEntryID 以支援 RSS (PostItem) 開啟

目前的 `OpenMailByEntryID` 實作中，強制將 `GetItemFromID` 取得的物件轉型為 `Outlook.MailItem`。當使用者試圖開啟 RSS 項目（其類型為 `Outlook.PostItem`）時，會發生 `InvalidCastException` 導致開啟失敗。

## 擬定變更

### Form1_MainTabs.vb

#### [MODIFY] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_MainTabs.vb)

- 修改 `OpenMailByEntryID` 內部的迴圈邏輯。
- 將 `Dim mail As Outlook.MailItem` 改為更通用的 `Dim olItem As Object`。
- 使用 `TryCast` 嘗試轉型為 `MailItem` 或 `PostItem`。
- 若皆非上述兩者，則使用 `Microsoft.VisualBasic.Interaction.CallByName` 作為最後手段嘗試呼叫 `Display`。
- 確保 `TryMarshalRelease` 正確處理該物件。

## 驗證計畫

### 手動測試
1. 在 Tab3 或 Tab4 中選取一個 RSS 項目（PostItem）。
2. 執行雙擊或按下 Enter 鍵。
3. 確認 Outlook 視窗正常彈出顯示該 RSS 內容。
4. 同時選取一封普通郵件與一個 RSS 項目，確認兩者皆能正常批次開啟。
