# Tab3 搜尋功能架構重構 (Pipeline Refactoring)

本次重構成功地將 Tab3 附件搜尋功能的內部骨架打掉重練，建立起壁壘分明的 **管線化處理流程 (Pipeline Processing)**。我們解決了舊版架構中 L1 / L2 / L3 職責混淆的問題，並徹底隔離了 MAPI 存取與介面 (UI) 的耦合。

> [!NOTE]
> 本次重構完全保留了所有的功能邏輯與使用體驗，主要是針對內部代碼的「關注點分離 (Separation of Concerns)」進行治理。

## 重點架構演進

### 1. `Button3_Click` 退居高層指揮官 (L1)
我們將原先散落在 Button 內的條件邏輯、快取判斷全部抽離。現在它只包含了乾淨俐落的線性 Pipeline：
```vb
' 【載入層 Load】
Dim targetMails = 透過 GetCachedMailWithAttachment 取回

' 【過濾管線 Pipeline】
targetMails = FilterBySize(targetMails)

If 條件成立 Then
    targetMails = Await FilterByAttachmentDetailsAsync(...)
End If

' 【介面映射 Mapping】
Dim finalItems = MapToListViewItems_Tab3(targetMails)
```
這種語法不僅意圖清晰，未來若需要插入新的過濾條件（例如依據寄件人篩選），只需像積木一樣插在中間即可，防呆且高擴展。

### 2. 徹底淨化 MAPI 操作 (L3)
原先在 Form1_Main 中負責 Phase 1 核心掃描的 `ScanFolderWithAttachment`，因為包含了 `table.GetArray` 等極度底層的 COM 呼叫，被我們安全地搬移到了 `Form1_ComL3.vb`。
*   **新函數**：`GetMailWithAttachmentL3` (專責跟 Outlook 溝通)
*   **效益**：這確保了所有的 `TryMarshalRelease` 與 COM 例外都鎖死在 L3，業務邏輯層將再也看不到煩人的 Outlook MAPI 宣告。

### 3. 解耦 UI 繪製與過濾器
過去最詬病的 `ScanAttachmentDetail` 會在一邊撈資料的同時直接 `New ListViewItem`。
*   現在它改名為 **`FilterByAttachmentDetailsAsync`**，回傳極為單純的 `List(Of MailItemInfo)` 純資料清單。
*   產生 ListViewItem 的責任全權交給 **`MapToListViewItems_Tab3`**。
*   **終極智慧解耦**：映射函數會去敲一敲 L2.5 Cache `_cacheAttachFilename` 的門。如果裡面有該封信的資料，它就知道要顯示精準的數字；如果沒有跑過 Phase 2，它就聰明地印出 ">0"。我們不再需要在過濾迴圈中辛苦地轉傳狀態！

### 4. L2.5 雙快取代理層收斂
將全新的 `GetCachedMailWithAttachment` 移回 `Form1_ComL3.vb` 與大家團聚。這使得 L2.5 的 Cache 家族 (`FolderSize`, `FolderCount`, `MailCount`, `AttachFilename`, `Phase1`) 終於全數在用同一個門面面對上級呼叫。

這套代碼已經如絲般順滑，架構上的重構也已經補齊！
