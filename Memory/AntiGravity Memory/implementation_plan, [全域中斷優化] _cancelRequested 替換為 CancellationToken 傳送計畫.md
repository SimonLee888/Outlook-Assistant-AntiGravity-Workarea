# [全域中斷優化] _cancelRequested 替換為 CancellationToken 傳送計畫

本計畫旨在回應「將全部使用 `_cancelRequested` 的地方都改成傳送 `ct`」之目標，徹底消滅容易發生競爭狀態（Race Condition）與狀態殘留的全域布林旗標，改用 .NET 現代化的合作式中斷模式 (Cooperative Cancellation)。

## User Review Required

> [!IMPORTANT]
> - 大範圍重構：這將是一次影響超過 50 處的替換作業，涵蓋 `Form1_MainTabs.vb`（UI與層次控制）、`Form1_Outlook.vb`（COM 資料操作）、`Form1_SQLite2.vb` 及 `Form1_Win32API.vb` 等。
> - 會透過小塊寫入（Chunked Edits）慢慢替換，過程中請盡量不要去編輯被修改的這幾個檔案，避免 `Edit Failed (File locked)` 的狀況。
> - 針對某些原本利用 `Dim savedCancel = _cancelRequested` 暫時無視中斷的機制，我們將以直接呼叫方法並傳入 `CancellationToken.None` 來完美取代原先易錯的旗標蓋寫。

## Proposed Changes

---

### UI 層與協調層 (Form1_MainTabs.vb / Form1.vb)

#### [MODIFY] Form1.vb
- 移除全域布林變數 `Private _cancelRequested As Boolean = False`
- 清除 `Form1_KeyDown` 中已註解或過時的 `_cancelRequested = True` 操作。

#### [MODIFY] Form1_MainTabs.vb
1. 確保所有的進入點（例如 `SimTree2_AfterSelect`, `SimTree3_AfterSelect`, `Button3_Click`, `ListView2_MouseDoubleClick` 等）都有呼叫 `Dim ct As CancellationToken = PrepareNewTaskToken()` 並往下傳遞。
2. 刪除這些進入點中舊有的 `_cancelRequested = False` 重置敘述。
3. 把所有的協調層函式如 `ComputeYearCounts`, `ShowMonthView`, `ShowYearView`, `SearchOutlookAttachmentsAsync`, `SearchOutlookThreadsAsync`... 修改簽章，加入 `ct As CancellationToken`。
4. 把所有 `If _cancelRequested Then...` 取代為 `If ct.IsCancellationRequested Then...`。

---

### 核心資料層與工具層 (Form1_Outlook.vb / Form1_SQLite2.vb / Form1_Win32API.vb)

#### [MODIFY] Form1_Outlook.vb
- 針對 `GetYearCountsForFolder`, `GetMonthCountsForYear`, `SearchAttachmentsInFolder`, `ComputeFolderSizeInternal` 以及其他各式耗時長的迴圈遞迴，將 `ct As CancellationToken` 加入其參數。
- 將 `If _cancelRequested Then Exit For / Return` 全面取代為檢查 `ct.IsCancellationRequested` 確保資源能在第一時間跳出並結束 COM 操作。

#### [MODIFY] Form1_SQLite2.vb
- 同步在如 `WaitThenLoad`、背景快取預先讀取或其他可能阻擋的操作中，接收並檢查 `ct.IsCancellationRequested`。

#### [MODIFY] Form1_Win32API.vb
- API 層如果提供可中斷的目錄遍歷或資料轉換，同樣變更其參數，將原本依靠外層全域旗標 `_cancelRequested` 反映的情況，改成讀取傳入的 `ct` 訊號。

## Verification Plan

### 自動測試
- 呼叫編譯確保 `_cancelRequested` 變數移除後，全專案無參考遺落，成功建置。

### 手動驗證流程
- 啟動程式後，隨意點選樹狀節點（Tab2 或 Tab3）發動長時間運算，並立刻按下 ESC 鍵。
- 確認：
  1. 程式立刻停止轉圈並顯示「已中斷」。
  2. 沒有因切換造成下一個點選事件無法運作（舊時殘留旗標臭蟲不會重現）。
  3. 各個層次都不再依賴全域狀態，讓傳入的 `ct` 控制自己專屬的背景生命週期。
