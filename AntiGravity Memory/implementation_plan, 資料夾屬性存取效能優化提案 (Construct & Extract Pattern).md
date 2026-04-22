# 資料夾屬性存取效能優化提案 (Construct/Extract Pattern)

## 優化目標
將資料夾相關的 COM 屬性存取次數降低 50% 以上，特別是在大規模掃描 (BFS) 與頻繁 Debug 記錄的場景。

## 擬議改動清單

### 1. [核心掃描] Form1_Outlook.vb : `FlattenSubtreeToList`
- **現狀**: 迴圈內對每個 `subF` 讀取 `.FolderPath` (L404) 與 `.Name` (L407)。
- **改動**: 
    - 將佇列改為 `Queue(Of (Folder, String))`，內含已計算好的路徑。
    - 迴圈內僅讀取 `subF.Name`。
    - 使用 `parentPath & "\" & subF.Name` 生成子路徑。
- **效益**: **極大**。資料夾掃描速度預計提升 30%~50%。

### 2. [顯示邏輯] Form1_Outlook.vb : `GetSortedSubFolders`
- **現狀**: 
    - L279 讀取 `folder.Name` 用於 Log。
    - L320/L325 迴圈內同時讀取子資料夾的 `subF.Name` 與 `subF.FolderPath`。
- **改動**: 
    - 預先讀取父資料夾 `folder.FolderPath`。
    - 迴圈內只讀子資料夾 `.Name`，拼出路徑。
    - Log 用的名稱改從路徑切出。

### 3. [快取剪枝] Form1_MainTabs.vb : `BuildBfsFolderTree` (Tab 1 核心)
- **現狀**: 
    - L479 讀取 `curr.folderObj.FolderPath`。
- **改動**: 
    - 配合 `FlattenSubtreeToList` 的優化，讓 BFS 本身攜帶路徑，省去重複讀取。

### 4. [Debug 記錄] Form1_SQLite2.vb : 多個函數
- **現狀**:
    - `RenewAttachMailListAsync` (L606) 為了 _dbg 讀取 `folder.Name`。
    - `RenewCacheAsync` (L492) 同上。
- **改動**: 
    - 這些函數通常已帶入 `fPath` 參數，直接使用 .NET 函數從 `fPath` 切出最後一段作為名稱，**完全移除 `folder.Name` 呼叫**。

### 5. [輔助工具] 新增路徑處理 Helper
- **改動**: 在 `moduleStore.vb` 或 `Form1_Outlook.vb` 新增一個 `ExtractNameFromPath(path)` 的高效字串工具函式，封裝切分邏輯，處理 Root 路徑字串的特殊情況。
