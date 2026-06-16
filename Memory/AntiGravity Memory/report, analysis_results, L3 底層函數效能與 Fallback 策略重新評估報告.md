# L3 底層函數效能與 Fallback 策略重新評估報告

根據您提出的三個關鍵限制條件與目前 [Form1.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1.vb) 內的實作，我重新審視了 7 個 L3 函數（`GetMailCount`, `GetMailCountAll`, `GetFolderCount`, `GetFolderCountAll`, `GetFolderSize`, `GetFolderSizeAll`, `GetSubFolderList`）。以下為評估結果與結論。

## 限制條件與因應策略回顧

1. **RDO 的 `TotalItemCount` 與全樹讀取不保證成功**：
   在 PST 檔案上，有時 MAPI 快取屬性無法直接在 Root 層級反映整棵子樹的總量。因應策略：當單一屬性讀取失敗時，必須退而求其次進行展開（BFS），逐個資料夾加總。
   
2. **OOM 不能平行處理，但有 `GetTable().GetArray()`**：
   OOM 受限於 STA (Single-Threaded Apartment)，在 `Task.Run` 或 `Parallel.ForEach` 中呼叫會引發致命的 COM 錯誤。但它擁有 `GetArray` 的殺手級功能，可以一次把數千筆資料轉換為二維陣列帶回記憶體，是 OOM 大量資料傳輸的極速解。

3. **RDO 可以平行處理，但缺乏 `GetArray`**：
   Redemption (RDO) 是 free-threaded (MTA)，完美的平行處理對象。缺點是它沒有 OOM 那種一次全拉回陣列的 API；如果要讀每一封信的大小，必須依靠讀取其 Property/Fields。但在資料夾層級的操作（例如取得子資料夾清單或加總 Item.Count），RDO 的多執行緒效能輾壓 OOM。

---

## 各函數策略檢視與評估

### 1. 計數類：`GetMailCount`, `GetMailCountAll`, `GetFolderCount`, `GetFolderCountAll`
這些函數的目標是「算數量」，不需要讀取清單內每一封信的詳細欄位。

*   **當前順序**：
    `RDO 單次屬性直讀` ➔ `RDO Parallel BFS (平行處理)` ➔ `OOM 循序 BFS`
*   **評估結果：【絕佳】**
    *   `GetMailCountAll` 中，如果 ⓪ `rdoFolder.TotalItemCount` 失敗（符合您的限制要求 1），程式會無縫降級到 ① RDO 平行 BFS（使用 `GetSubFolderList_RDO` + `Parallel.ForEach`）。
    *   因為只問數量（`Items.Count` 或 `Folders.Count`），**不需要 `GetArray`**，充分發揮了 RDO 可以平行處理（限制條件 3）的優勢。就算 PST 資料夾高達上千個，多執行緒同時詢問 `Count` 也會在一瞬間完成。
    *   若連 RDO 都失敗，最後安全退回 ② OOM 循序 BFS，符合限制條件 2。

### 2. 容量計算：`GetFolderSize`, `GetFolderSizeAll`
容量計算最嚴苛，因為 PST 的資料夾層級 `PR_MESSAGE_SIZE_EXTENDED` 屬性經常無法讀取，導致必須統計每一封郵件。

*   **當前順序**：
    `RDO 平行讀取資料夾屬性` ➔ `OOM 循序 GetTable().GetArray(1000)` ➔ `OOM GetNextRow()`
*   **評估結果：【極度合理且為當前最佳解】**
    *   在 `GetFolderSizeAll` 中，程式會先嘗試用 RDO 的 `Parallel.ForEach` 快速讀取每個資料夾的 `PR_MESSAGE_SIZE_EXTENDED`。如果這條路（因為 PST 不支援等原因）失敗了，程式果斷放棄 RDO。
    *   降級到 OOM 後，非常精準地遵守了不用平行處理的鐵律，改用傳統迴圈，但**利用了 `GetTable().GetArray(1000)`** 來狂飆速度。這完美呼應了限制條件 2 與 3：既然 RDO 不能 GetArray，而 OOM 可以，在面臨需要讀取逐筆資料的場景（FolderSize）時，切換到 OOM 的 GetArray 是最明智的 fallback。

### 3. 目錄展開：`GetSubFolderList` 與 `GetSubFolderList_RDO`
*   **評估結果：【職責分離明確】**
    針對 OOM 和 RDO 分別實作了安全的 Queue BFS 和高併發的 ConcurrentBag BFS，各司其職，不互相干擾，完美支撐了上層函數的 Fallback 體系。

---

## 總結結論

現在的降級順序：
1. **極速：RDO 單一屬性（無展開）**
2. **高速：RDO 平行展開 BFS**（最適合只需要 Count 的場合）
3. **中速巨量：OOM GetTable + GetArray**（最適合需要掃過所有 Item、加總 Size 的場合）
4. **低速安全：OOM 基本屬性與循序展開**（最後的保險底線）

**這個順序是非常正確且最優的。** 
完美繞過了 OOM 不能平行的死穴，也補足了 RDO 沒有 GetArray 的短板。這套由快到慢、兼顧「MTA 並行」與「STA 陣列批次處理」的防禦性降級邏輯，在目前架構下無須再作大改，可視為目前的最佳實踐。
