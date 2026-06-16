# 修復 Tab4 SSD 快取 Lazy Load 失效計畫

抱歉，我前一次判斷錯誤。你的觀察完全正確，DB Lazy Load 機制確實存在，但它在重啟後「失效」了，導致重新掃描。

我深入分析了寫入 (`SaveCachesToDB`) 與讀出 (`DbGetFolderBasicMailInfos`) 的底層程式碼，終於發現了這個機制的盲點：**「掃描結果為 0 筆的資料夾，根本沒有被寫入資料庫！」**

### 問題根源 (Root Cause)
1. **寫入的盲點**：`basic_maillist` 資料表的設計是每一列對應「一封信」。在存入 SSD 的 `SaveFolderBasicMailInfosInner` 迴圈中，如果某個資料夾**沒有任何符合條件的系列郵件**（`Mails.Count = 0`），它就不會執行任何 `INSERT` 動作。
2. **讀出的誤判**：下次重啟時，`GetFolderBasicMailInfos` 去呼叫 DB Lazy Load，`DbGetFolderBasicMailInfos` 會去資料庫找這個資料夾。因為上次沒寫入，所以找到 `0` 筆。
3. **重新掃描**：因為 DB 找不到該資料夾的任何紀錄（連 `snapshot` 都拿不到），代理層只好判斷「快取不存在」，於是再次呼叫 COM (`GetFolderBasicMailInfosL3`) 進行掃描。
這就是為什麼當你全選所有 PST 檔，大部分裡面沒有系列郵件的資料夾，每次重啟都會再被重新掃描一次的原因。

Tab3 的 `attach_maillist` 也有著一模一樣的問題（沒有附件的資料夾，重啟後會再次掃描）。

## 擬定變更 (Proposed Changes)

為了讓資料庫能夠記住「這資料夾已經掃過了，只是結果為 0」，我們不需要改變資料表結構，只需要在存入空結果時，寫入一筆**特殊的佔位符 (Dummy Row)** 即可。

### [Component] SQLite 快取寫入層

#### [MODIFY] [Form1_SQLite2.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_SQLite2.vb)
*   **修改 `SaveFolderBasicMailInfosInner`**：
    *   如果 `kvp.Value.Mails.Count = 0`，則 `INSERT` 一筆 `entry_id` 為 `"EMPTY_DIR_" & folder_path` 的紀錄，並寫入當下的 `item_count_snap`。
*   **修改 `DbGetFolderBasicMailInfos`**：
    *   讀取時，加入 `hasRecord` 布林值標記。只要有讀到任何列就算命中 DB。
    *   如果讀到的 `entry_id` 開頭是 `"EMPTY_DIR_"`，則跳過不加入郵件清單，但會保留 `snap`。
    *   最後只要 `hasRecord` 為 True，即使 `result.Count = 0` 也能成功回傳空的 List 與正確的 Snapshot，讓 Lazy Load 成功命中。
*   **同步修改 Tab3 的快取**：
    *   對 `SaveAttachMailListInner` 和 `DbGetAttachMailList` 進行完全相同的 Dummy Row 邏輯處理，一併修復 Tab3 附件快取的重新掃描問題。

## 驗證計畫
1. 全選包含空資料夾的目錄，點擊掃描 Tab4。
2. 存入 SSD 後，關閉程式重啟。
3. 再次掃描相同的目錄，觀察 `_dbg` 是否顯示命中 DB，且不再觸發 COM 掃描。
