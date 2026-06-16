# Tab4 系列郵件佈局完善與互動優化實作計畫

此計畫接續先前的 Tab4 UI 重組工作，旨在解決佈局動態同步、主旨分組邏輯優化以及排序資訊反饋等細節問題，提升整體使用體驗。

## 使用者評論與決策

- **佈局同步**：目前的同步僅限於「初始啟動」，手動調整側邊欄時中間欄寬度不會連動。
- **分組邏輯**：目前直接按原始主旨分組，容易因 `Re:` 導致同系列對話被拆分。計畫導入「主旨清理」機制。
- **排序反饋**：強化 `ProgressBar2` 在排序切換後的狀態提示。

## 擬議變更

### [UI 佈局與同步優化]

#### [MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1.vb)

- **成員變數**：
  - [NEW] `Private _scnrTab4Results As SplitContainer`：將區域變數提升為成員變數，以便後續動態控管。
- **函數修改**：
  - `InitTab4UI`：
    - 將 `scnrResults` 賦值給 `_scnrTab4Results`。
    - 加入 `AddHandler SplitContainer4.SplitterMoved, AddressOf SyncTab4Splitter`。
  - [NEW] `SyncTab4Splitter`：當左側 `SplitContainer4` 調整時，按比例同步 `_scnrTab4Results` 的 SplitterDistance。

### [主旨分組與排序優化]

#### [MODIFY] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_MainTabs.vb)

- **輔助函數**：
  - [NEW] `GetCleanSubject(subject As String) As String`：清理主旨前綴（Re:, Fw:, 回覆:, 轉寄:），用於更精準的分組。
- **函數修改**：
  - `FillListView4`：
    - 使用 `GetCleanSubject` 進行分組。
    - 根據目前的排序狀態更新 `ProgressBar2.Text`，確保使用者隨時知道目前的排序基準。
  - `ListView4_ColumnClick`：
    - 更新排序提示邏輯，與 F6 的提示保持一致。

## 驗證計畫

### 自動化/手動測試
- **同步驗證**：手動拉動 `SplitContainer4` (左側樹寬度)，確認中間的主題清單欄位也會按比例自動調整。
- **分組驗證**：搜尋包含主題 `Re: Meeting` 與 `Meeting` 的系列，確認它們會被歸類在同一個主旨分組下。
- **排序切換**：
    1. 按下 **F6**，確認分組依「數量」或「名稱」重排。
    2. 點擊欄位標頭，確認排序資訊正確更新於狀態列。

## 開放問題

> [!IMPORTANT]
> 關於「主旨清理」：目前計畫清理 `Re:`, `Fw:`, `回覆:`, `轉寄:`。若使用者有其他特殊的前綴（例如特定系統標籤 `[BUG]` 等）也需要排除，請再告知。
