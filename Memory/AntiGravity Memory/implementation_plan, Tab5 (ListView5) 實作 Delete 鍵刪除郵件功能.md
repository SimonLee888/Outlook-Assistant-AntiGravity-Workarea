# Tab5 (ListView5) 實作 Delete 鍵刪除郵件功能

本計畫旨在為 Tab5 的 `ListView5` 新增 Delete 快捷鍵支援。當使用者按下 Delete 鍵時，程式將確認並把選中的郵件移動到 Outlook 的預設刪除郵件資料夾，並即時從 `ListView5` 列表中移除這些郵件，行為與 Tab4 的 `ListView4` 保持一致。

## 使用者評論與回饋要求

> [!IMPORTANT]
> - `ListView4` 與 `ListView5` 的欄位定義不同（`ListView5` 多了「群組」與「相似度」欄位），因此我們需要調整渲染邏輯以支援從原始資料清單中移除後重新整理 UI。
> - 我們將建立一個專屬於 `ListView5` 的刪除處理函數 `HandleLv5Delete`，並在 `Form1_MainTab345.vb` 中實作相關邏輯。

## 待解決問題 (Open Questions)

- 無。目前的實作計畫已涵蓋欄位差異與資料同步邏輯。

## 擬議變更 Proposed Changes

### [Outlook Assistant]

---

#### [MODIFY] [Form1_MainTab345.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_MainTab345.vb)

- **新增 `Lv5_KeyDown`**: 偵測 Delete 鍵並呼叫 `HandleLv5Delete`。
- **新增 `HandleLv5Delete`**: 
    - 收集選中項目的 EntryID。
    - 從緩存的 `groupDict` 中移除對應郵件。
    - 呼叫 `MoveMailsToRecycle` 執行實體移動。
    - 呼叫 `RenderLv5Group` 重新渲染 UI 以反映刪除後的狀態。
- **修改 `Bt5_Click`**: 
    - 儲存掃描結果 `groupDict` 到類別成員變數 `_tab5LastGroupResults`。
    - 儲存當前比對模式 `isExact` 到類別成員變數 `_tab5LastIsExact`。
- **修改 `RenderLv5Group`**: 確保渲染時能正確處理刪除後的資料。

## 驗證計畫 Verification Plan

### 自動測試
- 無

### 手動驗證
1. 切換至 Tab5。
2. 執行重複郵件掃描。
3. 選取一封或多封郵件。
4. 按下鍵盤 Delete 鍵。
5. 確認彈出確認視窗。
6. 確認點選「是」後，郵件從列表消失。
7. 開啟 Outlook 確認郵件已移至「刪除郵件」資料夾。
8. 檢查剩餘郵件的群組背景顏色是否依然正確（交替顯示）。
