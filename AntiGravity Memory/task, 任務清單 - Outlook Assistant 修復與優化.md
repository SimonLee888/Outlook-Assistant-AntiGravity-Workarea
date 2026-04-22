# 任務清單 - Outlook Assistant 修復與優化

- [x] 修復 Lv1 導航問題 (`Form1_MainTabs.vb`)
    - [x] 將 `EnterSelectedFolder` 改為使用 `EntryID` 進行節點比對
    - [x] 驗證比對邏輯是否能正確找到含有子資料夾的節點
- [x] 修復 Lv3 路徑顯示問題 (`Form1_SQLite2.vb`)
    - [x] 在 `DbGetAttachMailList` 中回填 `mail.FolderPath`
    - [x] 在 `LoadAttachMailListInner` 中回填 `mail.FolderPath`
- [/] 實作 Tab4/Tab5 SSD 快取 (方案 B)
    - [ ] 在 `Form1_SQLite2.vb` 新增 `basic_maillist` 資料表與索引
    - [ ] 在 `Form1_SQLite2.vb` 實作 `DbGetFolderBasicMailInfos` (Lazy Load)
    - [ ] 在 `Form1_SQLite2.vb` 實作 `SaveFolderBasicMailInfosInner` (批次儲存)
    - [ ] 將儲存邏輯整合進 `SaveCachesToSQLiteAsync`
    - [ ] 在 `CleanupOrphanFolderPath` 加入清理 `basic_maillist` 的邏輯
    - [ ] 在 `Form1_Outlook.vb` 重構 `GetFolderBasicMailInfos` 以支援 SSD 快取層
- [ ] 最終驗證與整合測試
    - [ ] 驗證 Lv1 導航
    - [ ] 驗證 Lv3 重啟後路徑顯示
    - [ ] 驗證 Tab4 重啟後 F5 速度 (SSD 命中)
