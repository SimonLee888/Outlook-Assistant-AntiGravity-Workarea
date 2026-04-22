# ListView4 F5 刷新功能實作計畫

此計畫旨在為 Tab4 (系列郵件) 的 `ListView4` 增加 F5 刷新機制，當使用者在列表中按下 F5 時，系統將重新從 Outlook 取得這些郵件的最新屬性（如主旨、大小、日期等）並更新介面。

## 擬議變更

### [Component] Form1_MainTabs.vb

#### [MODIFY] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_MainTabs.vb)

- 將 `ListView4_KeyDown` 事件處理器修改為 `Async Sub`。
- 在 `ListView4_KeyDown` 中新增對 `Keys.F5` 的偵測。
- **實作 `RefreshListView4MailsAsync` 函式**：
    1. 從 `TreeView4.SelectedNode.Tag` 取得目前的 `mailList`。
    2. 顯示等待游標並更新 `ProgressBar2` 進度文字。
    3. 遍歷 `mailList` 中的每一筆郵件：
        - 透過 `EntryID` 呼叫 `_olNS.GetItemFromID` 取得 Outlook 實體郵件對象。
        - 更新 `MailItemInfo` 結構中的 `Subject`、`Size`、`ReceivedTime` 與 `SenderName`。
        - 確保透過 `Marshal.ReleaseComObject` 釋放 COM 物件。
    4. 刷新完成後，呼叫 `FillListView4` 重新填入列表。
    5. 回復游標狀態並更新狀態列。

## 驗證計畫

### 手動驗證
1. 在 Tab4 執行掃描並選取一個系列。
2. 在 `ListView4` 按下 **F5**。
3. 觀察 `ProgressBar2` 是否顯示「正在重新讀取郵件資訊...」以及進度。
4. 確認列表內容是否有正常顯示，且排序狀態維持不變（或依據新資料重新排序）。
5. 驗證大約 20-50 封郵件的刷新速度與 UI 響應。

> [!NOTE]
> 重新讀取郵件屬性會涉及 COM 呼叫，對於大量郵件（超過 100 封）可能會有感官輕微延遲，但在「系列郵件」的場景下（通常數量不多），此做法最為直觀且能保證資料準確性。
