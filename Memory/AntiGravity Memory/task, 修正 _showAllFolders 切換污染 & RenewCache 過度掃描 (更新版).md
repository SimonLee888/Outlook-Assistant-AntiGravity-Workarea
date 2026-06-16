# 修正 _showAllFolders 切換污染 & RenewCache 過度掃描

## 修改清單

- `[x]` **修改 F** — Form1.vb `CheckShowAllFolders_CheckedChanged`：切換時多清 `_cacheMailCountAll` / `_cacheFolderCountAll`
- `[x]` **修改 Phase2 語意** — Form1_SQLite2.vb `RenewCacheAsync`：區分「全新資料夾（DB 從未記錄）」與「真正 dirty（snapshot 不符）」，全新資料夾跳過 `attach_maillist` 掃描
- `[x]` **修改 G** — Form1_SQLite2.vb `RenewCacheAsync` Phase 3/4：清除聚合快取時兩個模式鍵值都清（為 _showAllFolders 鍵值分支預留）
- `[x]` **選項 A — FillFolderCacheFromDbRow** (Form1_Outlook.vb)：加入 `skipAggregates` 參數，`True` 時跳過填 `mca`/`fca`/`fsa`
- `[x]` **選項 A — BuildBfsFolderTree DB lazy load** (Form1_MainTabs.vb)：改用 `skipAggregates:=True` 呼叫，不設 `isHit = True`，BFS 自行展開重算
- `[x]` 複檢所有修改點確認正確、複檢修改點前後是否遺留多餘程式碼
