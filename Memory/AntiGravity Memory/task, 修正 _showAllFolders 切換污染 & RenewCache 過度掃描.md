# 修正 _showAllFolders 切換污染 & RenewCache 過度掃描

## 修改清單

- `[ ]` **修改 F** — Form1.vb `CheckShowAllFolders_CheckedChanged`：切換時多清 `_cacheMailCountAll` / `_cacheFolderCountAll`
- `[ ]` **修改 Phase2 語意** — Form1_SQLite2.vb `RenewCacheAsync`：區分「全新資料夾（DB 從未記錄）」與「真正 dirty（snapshot 不符）」，全新資料夾跳過 `attach_maillist` 掃描
- `[ ]` **修改 G** — Form1_SQLite2.vb `RenewCacheAsync` Phase 3/4：清除聚合快取時兩個模式鍵值都清（為 _showAllFolders 鍵值分支預留）
- `[ ]` 複檢所有修改點，確認正確
