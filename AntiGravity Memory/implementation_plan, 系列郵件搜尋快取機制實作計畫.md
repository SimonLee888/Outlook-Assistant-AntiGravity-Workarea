# 系列郵件搜尋快取機制實作計畫

此計畫旨在優化 Tab4 (系列郵件) 的搜尋效能。透過在資料夾層級快取郵件的基本資訊與 `ConversationTopic`，當使用者重複搜尋或按下 F5 重新整理時，系統能跳過重複的 Outlook COM 屬性讀取，達到光速掃描的效果。

## 擬議變更

### [Component] Form1_Outlook.vb

#### [MODIFY] [Form1_Outlook.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_Outlook.vb)

- **新增全域快取變數**：
    - 定義 `_cacheFolderMailsWithTopic` (ConcurrentDictionary)，其 Key 為資料夾路徑，Value 為一個包含郵件資訊清單與 Snapshot (郵件數) 的 Tuple。
- **修改 `GetFolderBasicMailInfosL3`**：
    1. 在進入 `GetTable` 掃描前，先嘗試從 `_cacheFolderMailsWithTopic` 獲取資料。
    2. 使用 `GetLiveFolderSnapL3` 取得該資料夾當下的郵件總數作為 **Snapshot**。
    3. 如果快取中的 Snapshot 與目前系統中的 Snapshot 一致，則直接回傳快取的郵件清單。
    4. 如果快取不存在或 Snapshot 已失效（例如郵件增減），則執行完整的 `GetTable` 掃描，並在完成後將結果（含目前的 Snapshot）存入快取。

### [Component] Form1_MainTabs.vb

#### [VERIFY] [Button4_Click](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_MainTabs.vb)

- 確認 `Button4_Click` 維持現有的資料夾遍歷流程。由於底層 `GetFolderBasicMailInfosL3` 已實作快取，當 `F5` 觸發 `Button4.PerformClick()` 時，整體的搜尋速度將會顯著提升。

## 驗證計畫

### 手動驗證
1. 在 Tab1 選取一個擁有多個子資料夾與大量郵件的根目錄。
2. 切換至 Tab4 按下「搜尋系列郵件」。紀錄第一次掃描的耗時（應顯示在 `ProgressBar1`）。
3. 掃描完成後，按下 **F5**。
4. 比對第二次掃描的耗時。預期第二次掃描（命中快取）的耗時應遠小於第一次掃描。
5. 在 Outlook 手動刪除一封郵件，再次按下 F5。
6. 確認系統是否能偵測到 Snapshot (郵件數) 改變並正確觸發重新掃描，以保證資料準確性。

> [!TIP]
> 使用資料夾郵件數 (`PR_CONTENT_COUNT`) 作為 Snapshot 是一種高效且低成本的快取校驗方法。只要資料夾內的郵件總量沒變，我們就假設內容基本沒變。
