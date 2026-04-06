# Tab3 搜尋架構重構計畫：管線化 (Pipeline) 與乾淨分層

為了根除「業務邏輯與 UI 耦合」、「MAPI 底層邏輯越界」以及「主控端管太多」的三大結構性痛點，我為 Tab3 規劃了全新的 **管線化處理流程 (Pipeline Processing)**。此計畫嚴格遵守我們確立的 L1 / L2 / L2.5 / L3 階層。

## User Review Required

請檢視以下重構方案的職責分佈是否符合您對這個系統架構的美學要求。如果有想要調整命名的函數，或是有其他希望合併/切分的邏輯，請隨時提出。

## 📍 Proposed Changes

### 第一刀：淨化 MAPI 操作，回歸 L3 底層
目前在 `Form1_Main.vb` 中做為 Phase 1 核心的 `ScanFolderWithAttachment`，裡面有大量的 `table.GetArray`、`table.Columns.Add`，這是標準的底層 COM 存取，將全面移出主流層。

#### [MODIFY] `Form1_ComL3.vb`
- **[NEW] L3 方法 `GetMailWithAttachmentL3`**
  - 將原 `ScanFolderWithAttachment` 的實作完全搬移到 L3 區域。
- **[NEW] L2.5 方法 `GetCachedMailWithAttachment`**
  - 將原本放在 `CheckTab3CacheOrRescan` 這幾行讀取 `_cachePhase1tab3` 的快取判斷邏輯，向下封裝到 L2.5 層次，完美補齊缺角。

---

### 第二刀：淨化過濾引擎，解耦介面 UI
我們要終結 `ScanAttachmentDetail` 裡面會噴出 `ListViewItem` 的原罪。過濾器 (`Filter`) 就該只回傳資料 (`Model`)。

#### [MODIFY] `Form1_Main.vb` (Phase 2 重構)
- **[RENAME + MODIFY] `ScanAttachmentDetail` ➝ `FilterByAttachmentDetailsAsync`**
  - 將回傳型別從 `Task(Of List(Of ListViewItem))` 降級為 `Task(Of List(Of MailItemInfo))`。
  - 專注於跑 L2.5 快取比對與名稱、數量的條件篩選，通過篩選就 `Add(mail)`，不要自己組裝控制項。
- **[NEW] `FilterBySize`**
  - 將原先在 `Button3_Click` 裡面自行撰寫的 LINQ `Size` 過濾邏輯，封裝為獨立的商業邏輯函數。

#### [MODIFY] `Form1_Main.vb` (視覺化映射 Mapping)
- **[RENAME + MODIFY] `BuildListViewItem_Tab3` ➝ `MapToListViewItems_Tab3`**
  - 統一負責所有的 UI 綁定。
  - **核心亮點**：要怎麼知道「明確的附件個數 (Count) 還是未知的 ">0"」？不需要在 Filter 裡記算傳遞。此函數直接判定：`如果 L2.5 快取裡有這封信的附件名稱 => 顯示真實 Count`；`如果不在快取內 => 顯示 ">0"`。這讓 UI 層與過濾層實現終極完美解耦！

---

### 第三刀：主控台 (Button3_Click) 退居指揮官

#### [MODIFY] `Form1_Main.vb` (L1 重構)
將原本肥大的控制流程 (God Object) 簡化為以下極度清爽的 Pipeline 線性文法：

```vb
' 【載入層 Load】
Dim targetMails As New List(Of MailItemInfo)
For Each folder in folderList
    ' 向 L2.5 索取基礎郵件全集
    targetMails.AddRange(Await GetCachedMailWithAttachment(folder, progressPhase1))
Next

' 【過濾管線 Pipeline】
' 過濾 1: 大小條件 (瞬間)
targetMails = FilterBySize(targetMails)

' 過濾 2: 附件名稱與數量的深層篩選
If 需要條件篩選 Then
    targetMails = Await FilterByAttachmentDetailsAsync(targetMails, progressPhase2)
End If

' 【介面映射 Mapping】
Dim listViewItems = MapToListViewItems_Tab3(targetMails)
ShowResultTab3(listViewItems, ...)
```

## 預期效益
1. **一目了然的主函數**：`Button3_Click` 會變成只有幾十行的指揮官代碼。
2. **遵守單一職責 (SOLID)**：L3 專注挖資料，L2.5 專心當門神，L2 專注跑條件過濾，L1 專注畫介面。
3. **無痛維護**：日後如果 Tab3 要新增條件 (例如按寄件人過濾)，只要多加一層 `targetMails = FilterBySender(...)` 插在管線中間即可，完全不會動到上下游。
