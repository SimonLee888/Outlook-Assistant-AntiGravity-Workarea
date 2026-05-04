# ListView4 相似度百分比計算實作計畫

本計畫旨在實作 Tab4 (系列郵件) 中 ListView4 的相似度計算功能。當使用者選取不同郵件時，程式會自動計算該群組內其他郵件與選中郵件主旨的相似度，並顯示在「相似」欄位。

## 使用者評論與回饋要求
> [!IMPORTANT]
> 1. 計算邏輯將直接套用現有的 `CalculateSimilarity()` 函式。
> 2. 以目前選取的 `ListViewItem` 為 100% 基準。
> 3. 修改過程將遵循「小塊寫入 (Chunked Edits)」原則，並保留歷史註解。

## 修改範圍

### 1. Form1.vb (UI 初始化)
- 確認 `InitTab4UI` 中的 `ListView4` 欄位定義。
- 目前 `ListView4` 已定義欄位：`主旨`, `郵件大小`, `收到日期`, `寄件者`, `相似`, `EntryID`。
- 「相似」欄位於 Index 4。

### 2. Form1_MainTabs.vb (邏輯實作)
#### [MODIFY] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_MainTabs.vb)
- **新增獨立事件處理函式**：`Lv4_SelectedIndexChanged`。
    - 理由：避免污染共用的 `ShowPathToProgressBar`，並確保 Tab3 不受影響。
- **實作計算邏輯**：
    1. 取得目前選取的 `ListViewItem` (基準)。
    2. 取得該項目的 `Group` (同一話題的群組)。
    3. 遍歷該 `Group` 中的所有 `Items`。
    4. 呼叫 `CalculateSimilarity(基準主旨, 目標主旨)`。
    5. 更新目標項目的 SubItems(4) 為百分比格式 (例如 `95%`)。
    6. 將基準項目標示為 `100%`。

## 預計步驟 (Task)
1. 檢視 `Form1_MainTabs.vb` 的 `ShowPathToProgressBar`，決定是否在此擴充或新開事件。
2. 實作 `ListView4_UpdateSimilarity` 邏輯。
3. 複檢所有修改點確認正確。
4. 複檢修改點前後是否遺留多餘程式碼。

## 驗證計畫
### 手動驗證
- 開啟 Tab4 並執行系列郵件掃描。
- 選取某個群組中的郵件，觀察「相似」欄位是否即時更新。
- 選取同群組內不同郵件，確認 100% 基準隨之切換，其他項目的相似度數值隨之改變。
- 測試空選取或切換群組時的表現。
