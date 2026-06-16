# RenewCache 增強計畫：支援新資料夾偵測與非同步 UI 刷新

這份計畫旨在修正 TabPage6 "Setting" 頁面的「更新快取」功能。目前該功能僅會更新已存在的快取，若使用者在 Outlook 中新增了資料夾，快取更新流程會跳過它，導致新資料夾無法顯示在左側的 `SimTree` 中。

## User Review Required

> [!IMPORTANT]
> **UI 行為變更**：執行「更新快取」後，所有頁籤的 TreeView 將會被重置回「摺疊並僅顯示根節點」的狀態。這是為了確保所有新資料夾都能被正確載入，但也意味著使用者目前展開到一半的路徑會被收合。

## Proposed Changes

### 1. 資料處理層 (Data Layer)

#### [MODIFY] Form1_SQLite2.vb
*   **RenewCacheAsync**:
    *   在 **Phase 2 (比對)** 邏輯中，新增路徑存在性檢查：
        ```vbnet
        ' 如果 fPath 不在資料庫中 (row Is Nothing)，則判定為新資料夾
        ' 強制加入 dirtyDict 並處理父資料夾結構失效
        ```
    *   新增「父路徑快取失效」邏輯：當發現新資料夾時，從 `_cacheFolderTree` 中移除其父資料夾的快取，確保樹狀結構重新讀取。
    *   **Phase 3 (更新)**：確保新資料夾的 `GetMailCountL3` 與 `GetFolderCountL3` 被正確執行。

### 2. UI 互動層 (UI Layer)

#### [MODIFY] Form1.vb
*   **RenewCache_Click**: 
    *   在 `Await RenewCacheAsync(...)` 之後，新增 `Await RefreshAllTreeViewsAsync()` 呼叫。
*   **[NEW] RefreshAllTreeViewsAsync**:
    *   建立此非同步函數，循環處理 `SimTree1` ~ `SimTree4`。
    *   使用 `BeginUpdate/EndUpdate` 包裹 `Nodes.Clear()` 與 `LoadStoreToTreeView`，並適時插入 `Task.Yield()` 釋放 UI 執行緒。

### 3. 持久化層 (Persistence Layer)

#### [MODIFY] Form1_SQLite2.vb
*   **SaveFolderStatsInner**: 
    *   檢查新資料夾加入 `allPaths` 的邏輯。
    *   確保 `_cacheFolderIDs` 捕獲到的新 EntryID/StoreID 能被正確 INSERT OR REPLACE 到 `folder_stats` 表中。

---

## Verification Plan

### Automated Tests (Manual Sequence)
1. **環境準備**：開啟一個 PST 並記錄目前的資料夾結構。
2. **新增測試**：在 Outlook 中手動新增資料夾 `Folder_AntiGravity_Test`。
3. **執行更新**：到 Setting 頁點擊 `RenewCache` (不勾選 Include Size 以加快速度)。
4. **觀察 Log**：
    * 檢查是否出現 `Phase 2: 偵測到新資料夾` 相關 Log。
    * 檢查是否顯示 `UI 刷新開始`。
5. **最終確認**：
    * 所有 TreeView 是否自動收回到根節點。
    * 展開目標位置，確認 `Folder_AntiGravity_Test` 已出現。
    * 重啟程式，確認該資料夾依然存在於 TreeView 中（代表已成功寫入 SSD）。

### Manual Verification
* 確認在大約 500+ 資料夾的環境下，刷新過程不會造成 UI 無回應 (Not Responding)。
