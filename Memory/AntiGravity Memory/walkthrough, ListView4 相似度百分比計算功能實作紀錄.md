# ListView4 相似度百分比計算功能實作紀錄

## 修改內容總覽
本次修改為 Tab4 (系列郵件) 增加了動態相似度計算功能，讓使用者在選取不同郵件時，能即時看到同組內其他郵件與基準的相似程度。

### 1. 獨立事件掛載 (Form1.vb)
為了不影響 Tab3 與其他共用邏輯，我為 `ListView4` 專門掛載了獨立的 `SelectedIndexChanged` 處理器：
```vb
' d:\Users\Simon\Dropbox\私人文件\Visual Studio\Visual Studio 18 (2026)\Outlook Assistant - (AntiGravity測試區)\Form1.vb
If lv.Name = "ListView4" Then AddHandler lv.SelectedIndexChanged, AddressOf Lv4_SelectedIndexChanged
```

### 2. 相似度計算邏輯 (Form1_MainTabs.vb)
實作了 `Lv4_SelectedIndexChanged` 函式：
- **基準選取**：以 `lv.SelectedItems(0)` 為基準，標示為 `100%`。
- **組內遍歷**：僅針對 `baseItem.Group` (同一個話題群組) 內的項目進行計算，提升效率並符合直覺。
- **現有邏輯複用**：呼叫 `CalculateSimilarity` 進行計算，並將結果轉換為整數百分比（如 `95%`）。

## 驗證結果
- [x] **Tab3 穩定性**：確認 `ListView3` 仍僅觸發 `ShowPathToProgressBar`，不受新邏輯影響。
- [x] **Tab4 即時性**：切換選取項目時，「相似」欄位會立即依據新基準重新計算。
- [x] **分組隔離**：計算範圍嚴格限制在目前的 `ListViewGroup` 內。

## 複檢紀錄
- 已確認變數定義完整，無遺留多餘偵錯代碼。
- 已確認 `BeginUpdate` 與 `EndUpdate` 配對使用，防止 UI 閃爍。
- 已確認 `SubItems(4)` 的索引與 `InitTab4UI` 中的欄位定義一致。
