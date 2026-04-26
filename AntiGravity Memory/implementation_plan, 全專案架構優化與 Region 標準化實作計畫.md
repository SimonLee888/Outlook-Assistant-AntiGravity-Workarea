# 全專案架構優化與 Region 標準化實作計畫

本計畫旨在透過明確的層次劃分 (Layers) 與 Region 重新定義，提升程式碼的可讀性與維護性，並解決部分函數定位不明的問題。

## 使用者審核請求

> [!IMPORTANT]
> **關於 `ComputeFolderStatsAsync` 的歸屬：**
> 我建議將此函數從 `Form1_MainTabs.vb` 移至 `Form1_Outlook.vb`。
> **理由：** 它是全專案核心的數據處理流程，雖然目前只有 Tab1 在用，但性質上屬於「業務邏輯 (Layer 2)」，而非「UI 事件 (Layer 1)」。將其移出後，`MainTabs` 將更專注於 UI 互動。

> [!NOTE]
> **Region 命名標準：**
> 統一採用 `■ [編號] Layer [層級]: [功能名稱]` 的格式，例如：
> - `■ 10 Layer 1: UI 事件與渲染`
> - `■ 20 Layer 2: 流程協調層`
> - `■ 30 Layer 2.5: 快取代理層`
> - `■ 40 Layer 3: 底層數據存取`

## 待解決問題 (Open Questions)

1.  除了 `ComputeFolderStatsAsync`，是否還有其他 Tab2/Tab3 的複雜計算邏輯需要一併移入邏輯層？
2.  `Form1_OST.vb` 是否也要同步進行嚴格的三層劃分，還是維持目前的專屬結構？

## 擬定變更

---

### [Component] Form1.vb (主架構與工具)

#### [MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1.vb)
- **新增 Region**: `■ 03 全域非同步工具`
- **移動函數**: 將 `OkayNowYouHaveToken` 與 `ThrottledYieldAsync` 從生命週期區域移至此區。
- **重命名 Region**: 將原本模糊的輔助函數區域重新分類。

---

### [Component] Form1_Outlook.vb (核心邏輯層)

#### [MODIFY] [Form1_Outlook.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_Outlook.vb)
- **結構重組**: 廢止 `■ 10 底層 COM 函數群`，拆解為以下四個 Region：
  1.  `■ 10 初始化與生命週期`: `InitOutlookNamespace`, `InitRdoSession` 等。
  2.  `■ 20 Layer 2: 流程協調層`: `ComputeFolderStatsAsync` (從 MainTabs 移入), `GetUniqueFolderList`。
  3.  `■ 30 Layer 2.5: 快取代理層`: `GetMailCount`, `GetFolderCount`, `GetYearCountsForFolder` 等快取邏輯。
  4.  `■ 40 Layer 3: 底層數據存取`: 所有的 `xxxL3` 結尾函數。
- **移動位置**: 將 `FillFolderCacheFromDbRow` 從末尾移至 `Layer 2.5` 區塊。

---

### [Component] Form1_MainTabs.vb (分頁事件層)

#### [MODIFY] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_MainTabs.vb)
- **區域化**: 根據 Tab1 ~ Tab5 重新標記 `■ Layer 1: Tab [X] UI 事件`。
- **清理**: 移除已移至 Outlook.vb 的 `ComputeFolderStatsAsync` 及其私有輔助結構（若僅為邏輯用途）。
- **簡化**: 確保事件處理器 (如 `SimTree1_AfterSelect`) 只負責「調度」與「呈現」，不含計算邏輯。

---

## 驗證計畫

### 自動化測試
- 使用 `view_file` 複檢修改後的行號。
- 編譯專案，確保跨檔案呼叫的路徑正確。

### 手動驗證
- 點選 Tab1 樹狀目錄，確認統計數字依然正確顯示（驗證 `ComputeFolderStatsAsync` 搬移後運作正常）。
- 執行「儲存快取」，確認 `FillFolderCacheFromDbRow` 運作正常。
- 點選快速統計，確認 `ThrottledYieldAsync` 的節流與中斷機制依然靈敏。
