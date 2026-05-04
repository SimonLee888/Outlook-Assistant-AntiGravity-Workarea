# Task: 修正 Tab4 SSD 快取雙重失效問題

## Root Cause A：|needTopic 後綴污染 DB folder_path
- [x] SaveFolderBasicMailInfosInner：`Dim fp = kvp.Key` → `kvp.Key.Split("|",c)(0)` 剝離後綴取純路徑
- [x] LoadFolderBasicMailInfosInner：從 DB 讀回時，重建 key 為 `fp & "|True"` 對齊 GetFolderBasicMailInfos 格式
- [x] InitDatabase：新增一次性 migration，清除 basic_maillist 中舊版污染的帶後綴路徑資料

## Root Cause B：空資料夾掃過後不留紀錄，重啟後重掃
（已由前任開發者部份實作 EMPTY_BASIC_ sentinel 機制，本次補齊漏掉的 Load 端邏輯）
- [x] SaveFolderBasicMailInfosInner：mails.Count=0 時寫入 EMPTY_BASIC_ sentinel row（已有）
- [x] DbGetFolderBasicMailInfos：sentinel row 設 hasRecord=True 確保回傳非 Nothing（已有）
- [x] LoadFolderBasicMailInfosInner：讀回時偵測 EMPTY_BASIC_ 並 Continue While，避免建立假 MailItemInfo（本次修正）

## 最後
- [x] 複檢所有修改點確認正確、複檢修改點前後是否遺留多餘程式碼
