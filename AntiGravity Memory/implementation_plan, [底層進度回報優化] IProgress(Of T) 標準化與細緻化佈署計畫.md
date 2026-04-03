# [底層進度回報優化] IProgress(Of T) 標準化與細緻化佈署計畫

本計畫旨在解決 `Form1_ComL3.vb` 底層計數及大小計算函數在執行大規模 PST/Exchange 掃描時，因 UI 回報機制不夠頻繁或機制較老舊（Action 回呼），導致使用者體驗不佳、疑似當機的問題。

## [核心變更說明]

將原本的 `Action(Of Integer, Integer)` (processed, total) 進度回報，全面升級為 **`.NET IProgress(Of T)`** 模式。

> [!TIP]
> `IProgress(Of T)` 的主要優勢在於它能自動捕捉與排程到 UI 的 `SynchronizationContext`，簡化了背景執行緒更新 UI 的 `Invoke` 邏輯，並能與 `Async/Await` 完美結合。

---

## [擬議變更概要]

### 1. [定義通用進度結構]
由於底層函數眾多，為了統一路徑與格式，將定義一個 `L3ProgressReport` 結構體。

*   **欄位**：`CurrentCount` (目前完成數)、`TotalCount` (總數)、`Message` (提示訊息，如「正在計算 A 大小...」)。

### 2. [優化 L3 底層函數] ([Form1_ComL3.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_ComL3.vb))

我們將對以下關鍵函數進行改裝，加入 `IProgress` 參數：

#### [MODIFY] `GetMailCountAll`
*   將 `onProgress As Action(Of Integer, Integer)` 替換為 `progress As IProgress(Of L3ProgressReport)`。
*   在不同的 Fallback 路徑中（RDO 平行或 OOM BFS）確實上報進度。

#### [MODIFY] `GetFolderSize` (包含 GetTable 迴圈)
*   **關鍵優化點**：目前的 `GetTable().GetArray(1000)` 迴圈中雖然有 `Await Task.Yield()`時，但完全沒有對外回報進度。
*   **新做法**：在分批（每 1000 封）處理郵件大小時，主動呼叫 `progress.Report`，讓使用者看見數字在跳動。

#### [MODIFY] `GetFolderSizeAll` (整棵資料夾樹的大小計算)
*   新增 `progress As IProgress(Of L3ProgressReport)` 參數。
*   在遍歷資料夾清單時，同步回報「目前正在計算第幾個資料夾」，提供宏觀進度。

---

### 3. [對接 UI (Layer 1/2)] ([Form1_Main.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_Main.vb))

#### [MODIFY] `ComputeFolderStatsAsync`
*   調整對 L3 函數的呼叫，傳入 UI 層建立的 `Progress` 處理器。
*   更新 `lblStatus1` 的邏輯。

#### [MODIFY] `ListView1_ItemMenu` (右鍵計算大小)
*   在長時間計算時，使用正確的進度處理機制更新 UI。

---

## [使用者 review 確認事項]

> [!IMPORTANT]
> 1. **進度更新密度**：在高頻率迴圈（如數萬封郵件）中，如果每一封都 report 會卡死。我計畫 L3 底層採「分批 Report」（如每 20 個資料夾或每 1000 封郵件上報一次）。
> 2. **中文回覆規範**：本計畫與衍生代碼註解均會依照規範，保留原始註解並添加 `by AntiGravity, 2026/04/02`。

---

## [驗證計畫]

### 手動功能測試 (Manual Verification)
1.  **Tab 1 資料夾點選**：點選一個包含數百個資料夾的 PST，確認 `lblStatus1` 狀態文字有隨著掃描節奏跳動。
2.  **右鍵計算資料夾大小**：選取多個大資料夾執行計算，確認 `ListView1` 的子項目列或狀態列能即時反映計算進度。
3.  **ESC 取消測試**：確認在回報進度的同時，點選 ESC 依然能準確中止背景運算，不影響進度條清空。

### 效能監控
*   確認加入 `IProgress` 後，在 PST 統計時沒有顯著的效能下降（若有，將調整 Report 的觸發閾值）。
