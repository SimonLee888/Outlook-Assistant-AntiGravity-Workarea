# Outlook Assistant 修正與優化任務清單

- `[x]` **1. 修復 Lv1 導航 Bug**
  - `[x]` 在 `EnterSelectedFolder` 增加 Debug 輸出以追蹤匹配失敗原因
  - `[x]` 修正名稱匹配或等待展開的邏輯問題
- `[x]` **2. 實作 Tab4/Tab5 SSD 快取 (方案 B: 獨立資料表)**
  - `[x]` `Form1_SQLite2.vb`: 修改 `GetCreateTablesSql` 新增 `basic_maillist` 資料表與索引
  - `[x]` `Form1_SQLite2.vb`: 實作 `SaveFolderBasicMailInfosInner` 與 `DbGetFolderBasicMailInfos`
  - `[x]` `Form1_SQLite2.vb`: 修改 `SaveCachesToSQLiteAsync` 以寫入快取，並在 `CleanupOrphanFolderPath` 清除孤兒資料
  - `[x]` `Form1_Outlook.vb`: 新增 `GetFolderBasicMailInfos` (L2.5) 並重構 `GetFolderBasicMailInfosL3` (L3)
  - `[x]` `Form1_MainTabs.vb`: 更新呼叫端改為呼叫 L2.5 函式
- `[x]` **3. 修復 Lv3 路徑顯示 Bug**
  - `[x]` `Form1_SQLite2.vb`: 於 `DbGetAttachMailList` 和 `LoadAttachMailListInner` 補上 `mail.FolderPath` 賦值
- `[/]` **4. 驗證與收尾**
  - `[ ]` 驗證 Lv1 跳轉是否正常
  - `[ ]` 驗證 Tab4 重啟後按 F5 是否達到秒開
  - `[ ]` 驗證 Tab3 重啟後路徑是否顯示
