# 任務：重構 DB 讀取與驗證邏輯

- [ ] 準備階段：更新相關資料結構與型別
    - [ ] [MODIFY] `Form1_SQLite2.vb`: 將 `FolderStatsDbRow` 欄位改為 `Long` 並更新 `DbGetFolderStats` 讀取邏輯
    - [ ] [MODIFY] `Form1_MainTab12.vb`: 將 `FolderBfsEntry` 欄位改為 `Long`
    - [ ] [MODIFY] `Form1_Outlook.vb`: 將 `_cacheMailCount` 等字典改為 `Long`
- [ ] 實作階段：建立 Helper 與重構 L2.5 函式
    - [ ] [NEW] `Form1_Outlook.vb`: 建立 `TryGetValidDbRow` 輔助函式
    - [ ] [MODIFY] `Form1_Outlook.vb`: 重構 `GetMailCount`
    - [ ] [MODIFY] `Form1_Outlook.vb`: 重構 `GetFolderCount`
    - [ ] [MODIFY] `Form1_Outlook.vb`: 重構 `GetFolderSizeAsync` (簽章已是 Long，套用 Helper)
    - [ ] [MODIFY] `Form1_Outlook.vb`: 重構 `GetFolderSizeAllAsync` (簽章已是 Long，套用 Helper)
    - [ ] [MODIFY] `Form1_Outlook.vb`: 重構 `GetMailCountAllAsync`
    - [ ] [MODIFY] `Form1_Outlook.vb`: 重構 `GetFolderCountAllAsync`
- [ ] 複檢階段
    - [x] 複檢所有修改點確認正確、複檢修改點前後是否遺留多餘程式碼
