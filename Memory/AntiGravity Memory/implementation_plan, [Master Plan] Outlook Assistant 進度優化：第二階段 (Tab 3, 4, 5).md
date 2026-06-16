# [Master Plan] Outlook Assistant 進度優化：第二階段 (Tab 3, 4, 5)

本計畫將延續之前的優化工作，完成剩餘分頁（附件、系列、重複郵件）的進度回報標準化，並修正 `GetFolderSize` 缺失的進度回報邏輯。

## User Review Required

> [!IMPORTANT]
> **1. 進度回報標準化 (IProgress)**
> 所有掃描功能將統一使用 `IProgress(Of L3ProgressReport)`，並由 L1 (UI 層) 決定如何顯示文字，確保 L2/L3 (邏輯層) 與 UI 完全解耦。
> 
> **2. 100ms 時間節流 (Throttle Balance)**
> 在所有高頻率迴圈（如讀取附件、建構重複列表）中，嚴格遵守 100ms 節流，防止訊息泵過載造成 UI 凍結。
> 
> **3. 性能保護：絕對禁止 Dbg()**
> 掃描郵件或資料夾的密集迴圈中，將絕對移除或標記掉屬性讀取相關的 `Dbg()`，避免 `DebugForm` 被洗板導致程式卡死。

## Proposed Changes

### [Component] 底層數據層 (Form1_ComL3.vb)

#### [MODIFY] [Form1_ComL3.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_ComL3.vb)
- **`GetFolderSize`**: 
    - 補上 `GetArray` 批次迴圈內的 `progress.Report()`。
    - 雖然是單一資料夾，但在大資料夾（數萬封）掃描時仍需回報內部進度。
- **`GetFolderCountAll`**:
    - 在資料夾樹展開過程中加入更細緻的進度回報與 100ms 節流。
- **`GetFolderSizeAll`**: 
    - 確保在 OOM 循序路徑中，正確傳遞 `progress` 給內部的 `GetFolderSize`。

---

### [Component] UI 與掃描流程層 (Form1_Main.vb)

#### [MODIFY] [Form1_Main.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_Main.vb)

- **Tab 3 (附件搜尋)**:
    - 優化 `ScanFolderWithAttachment` (Phase 1) 與 `ScanAttachmentByName` (Phase 2) 的進度回報顆粒度。
    - Phase 2 每處理一個郵件後加入 `Await Task.Delay(1)` 的時機點微調，確保 ESC 響應更快。
- **Tab 4 (系列郵件)**:
    - 將 `Button4_Click` 的掃描與分析邏輯導入 `IProgress`。
    - 針對「搜尋資料夾」與「建立 TreeView 節點」的過程增設 100ms 節流。
- **Tab 5 (重複郵件)**:
    - 將 `Button5_Click` 的全信箱掃描（跨 Store）邏輯導入 `IProgress`。
    - 統一使用 `Stopwatch` 做節流，取代原有的 `totalProcessed Mod 10`。
    - 在「重複群組列表建構」階段也加入進度回報，預防大數據量時建表瞬間的假死現象。

---

## Open Questions

> [!NOTE]
> **關於 Tab 5 的模糊比對 (Levenshtein)**
> 目前主要針對「掃描資料夾」與「建構列表」做進度回報。如果郵件量極大，純記憶體的模糊比對運算耗時可能較長，若使用者反映卡頓，後續可考慮再針對比對迴圈進行拆解。

## Verification Plan

### 自動化測試
- 無（主要涉及 UI 互動與 COM 非同步操作）。

### 手動功能測試
1.  **Tab 1 - GetFolderSize 測試**：對超大資料夾點擊右鍵選單「Show This Folder Size」，確認 `lblStatus1` 每 100ms 正常跳動。
2.  **Tab 3 - 附件搜尋**：執行大範圍搜尋，確認 Phase 1/2 的狀態顯示一致。
3.  **Tab 4 - 系列掃描**：確認建立長列表節點時不再導致主 UI 「沒有回應」。
4.  **Tab 5 - 重複掃描**：觀察跨 Store 掃描時狀態列是否能正確反映 Store 名稱。
5.  **全域 ESC 測試**：在上述任何掃描期間按下 ESC，確認搜尋能立刻終止並回傳「已中斷」。
