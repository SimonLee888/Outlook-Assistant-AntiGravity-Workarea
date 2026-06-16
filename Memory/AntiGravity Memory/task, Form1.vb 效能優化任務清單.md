# Form1.vb 效能優化任務清單

## 1. GetTable.GetArray() 批次讀取優化
- [x] `ScanFolderWithAttachment` (Tab3 Phase1): While GetNextRow → GetArray(1000)
- [x] `GetFolderSize` ① 路徑: Do While GetNextRow → GetArray(1000)
- [x] `GetFolderSizeLegacy`: Do While GetNextRow → GetArray(1000)

## 2. Redemption 平行化 (保留 TotalItemCount 不動)
- [x] `GetMailCountAll` ①: Task.WhenAll 只在 `_rdo IsNot Nothing` 時走平行，無 RDO 直接走循序
- [x] `GetFolderSizeAll` ①: 無需修改 (已使用 Async interleaving，不走 Task.Run，STA 安全)

## 3. GetFolderCountAll Redemption 平行 + 循序 fallback
- [x] `GetFolderCountAll`: 新增 ② RDO 平行 fallback + ③ OOM 循序 fallback

## 4. 註解規範
- [x] 所有新增註解使用 `' 2026/3/24 by AntiGravity` 標記

## 5. Tab2 年份/月份統計 Restrict → GetTable.GetArray
- [x] `GetYearCountsForFolder`: Restrict 逐年 → GetTable + GetArray 一次讀完、記憶體分組
- [x] `GetMonthCountsForYear`: Restrict 逐月 → GetTable (日期範圍 filter) + GetArray、記憶體分組

## 6. ListView1 右鍵選單改用新底層 L3 函數
- [x] `ListView1_ItemMenu`: `GetFolderSizeLegacy()` → `GetFolderSizeAll()`
