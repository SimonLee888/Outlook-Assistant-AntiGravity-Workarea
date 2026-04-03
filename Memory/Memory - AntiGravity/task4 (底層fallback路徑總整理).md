# Form1.vb 效能優化與 Fallback 鏈重建任務清單

## 1. 基礎巡覽函數 (GetSubFolderList)
- [x] 實作 `GetSubFolderList_RDO()`: 使用 ConcurrentQueue + Parallel.ForEach 回傳 `List(Of RDOFolder)`
- [x] 確認原版 `GetSubFolderList()` 保留為 OOM 專用循序版

## 2. 郵件計數 (GetMailCount / GetMailCountAll)
- [x] 改寫 `GetMailCount()`: 1. RDO Items.Count 2. MAPI PR_CONTENT_COUNT 3. OOM Items.Count
- [x] 改寫 `GetMailCountAll()`: 1. RDO TotalItemCount 2. RDO平行 (GetSubFolderList_RDO + Parallel.ForEach + Interlocked.Add) 3. OOM 循序 4. 舊版遞迴 (移至備用區)

## 3. 資料夾計數 (GetFolderCount / GetFolderCountAll)
- [x] 改寫 `GetFolderCount()`: 1. RDO Folders.Count 2. OOM Folders.Count
- [x] 改寫 `GetFolderCountAll()`: 1. RDO (若支援) 2. RDO 平行 3. OOM 循序

## 4. 大小計算 (GetFolderSize / GetFolderSizeAll)
- [x] 改寫 `GetFolderSize()`: 1. RDO 讀取屬性 PR_MESSAGE_SIZE_EXTENDED 2. OOM GetTable.GetArray 3. OOM GetTable.GetNextRow
- [x] 改寫 `GetFolderSizeAll()`: 1. RDO 平行讀取屬性 2. OOM 循序 + 內部呼叫 GetFolderSize (GetTable.GetArray)

## 5. 日期統計 (Tab2)
- [x] 確認 `GetYearCountsForFolder` 與 `GetMonthCountsForYear` 已從 Restrict 改為 GetTable.GetArray (先前階段已優化完成)

## 6. 註解規範
- [x] 嚴格保留原始 debug 紀錄與改版註解，所有修改加上 `' 2026/3/24 by AntiGravity` 並擴充說明最新設計
