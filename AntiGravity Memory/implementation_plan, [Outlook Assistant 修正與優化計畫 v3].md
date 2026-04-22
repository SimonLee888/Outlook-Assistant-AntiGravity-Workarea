# [Outlook Assistant 修正與優化計畫 v3]

本計畫根據使用者最新的指示（方案 B：維持資料表獨立）進行調整。

## User Review Required

> [!IMPORTANT]
> **關於 Tab4/Tab5 SSD 快取 (方案 B)**：
> 確定採行**新增獨立資料表**的方案，以保持 Tab3 與 Tab4/5 的資料邊界清晰：
> 1. 新增 `basic_maillist` 資料表，專門儲存 Tab4/Tab5 掃描到的全部郵件基本資訊與主題 (Topic)。
> 2. Tab3 繼續使用原本的 `attach_maillist`。
> 3. 雖然會增加一點點儲存空間，但兩者的掃描邏輯與資料用途可以完全獨立，互不干擾。

## Proposed Changes

---

### 1. 修復 Lv1 導航 (有子資料夾卻進不去的問題)

#### [MODIFY] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_MainTabs.vb)
- 在 `EnterSelectedFolder` 加入詳細的 `_dbg` 日誌，追蹤：
    - `parentNode.Text` 與 `parentNode.Nodes.Count`
    - 比對中的 `subject` 與迴圈中每個 `node.Text`
    - `GetFolderCount(targetFolder)` 的實際回傳值。
- **修復重點**：確認比對邏輯是否因為空格或大小寫失效，並確保 `parentNode.Expand()` 後有正確等待節點載入。

---

### 2. Tab4/Tab5 SSD 快取 (新增資料表)

#### [MODIFY] [Form1_SQLite2.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_SQLite2.vb)
- **資料庫擴充**：
  - `GetCreateTablesSql`：新增 `basic_maillist` 表 (欄位：`entry_id, folder_path, subject, msg_size, received_time, sender_name, topic, item_count_snap, updated_at`)。
  - 建立索引 `idx_basic_folder ON basic_maillist(folder_path)`。
- **存取邏輯**：
  - 新增 `SaveFolderBasicMailInfosInner(txn)`：實作 `_cacheFolderBasicMailInfos` 的批次寫入。
  - 新增 `DbGetFolderBasicMailInfos(fPath)`：實作單一資料夾的 Lazy Load 讀取。
  - 修改 `SaveCachesToSQLiteAsync`：加入 `SaveFolderBasicMailInfosInner` 呼叫。
  - 修改 `CleanupOrphanFolderPath`：加入清理 `basic_maillist` 孤兒資料的 SQL。

#### [MODIFY] [Form1_Outlook.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_Outlook.vb)
- **Layer 2.5 快取代理**：
  - 實作 `GetFolderBasicMailInfos` (取代直接呼叫 L3)：邏輯為 記憶體 -> DB Lazy Load -> L3。
  - 修改 `GetFolderBasicMailInfosL3`：將其回傳結果存入記憶體快取。

---

### 3. Lv3 (Tab3) 路徑不顯示問題

#### [MODIFY] [Form1_SQLite2.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_SQLite2.vb)
- 在 `DbGetAttachMailList` 與 `LoadAttachMailListInner` 中，補上 `mail.FolderPath = fPath` (或 `fp`) 的賦值邏輯。

## Verification Plan
1. **Lv1 驗證**：透過日誌確認 `EnterSelectedFolder` 在有子資料夾時能正確找到 `foundNode`。
2. **Tab4/5 驗證**：重啟程式後在 Tab4 按 F5，確認是否能秒開結果，並檢查資料庫中 `basic_maillist` 是否有值。
3. **Lv3 驗證**：確認重啟後搜尋，路徑欄位不再是空白。
