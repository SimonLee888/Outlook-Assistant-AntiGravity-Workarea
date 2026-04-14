# 全域變數與物件存取權限優化計畫

此計畫旨在稽核並收縮 `Form1` 及其關聯 Partial Class 中的變數存取權限。目前部分內部資料結構（如資料庫 Row 對象）使用了 `Friend` 級別，若這些結構僅在 `Form1` 及其 Partial Class 內部使用，應收縮為 `Private` 以符合封裝原則。

## 使用者評論與回饋要求
> [!IMPORTANT]
> 由於專案使用 `Partial Class` 拆分檔案，所有在 `Form1` 不同檔案間共用的變數（例如快取 `_cacheMailCount` 或資料庫連線 `_db`）**必須保持在 `Private` 級別**（Partial Class 內部對 `Private` 是互通的）。
> 若有任何變數需要被外部 Form（如 `DebugForm`）或 Module 讀取，則不應收縮。

## 擬議變更

### [封裝優化 - 資料結構層]

#### [MODIFY] [Form1_Outlook.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_Outlook.vb)
- 將 `Friend Structure MailItemInfo` 改為 `Private Structure MailItemInfo`。
- 將 `Friend Structure ProgressReport` 改為 `Private Structure ProgressReport`。
- 將 `Friend Async Function PreloadAttachByRDOAsync1/2` 改為 `Private`。

#### [MODIFY] [Form1_SQLite2.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_SQLite2.vb)
- 將 `Friend Class FolderStatsDbRow` 改為 `Private Class FolderStatsDbRow`。
- 將 `Friend Class MailwithAttachsDbResult` 改為 `Private Class MailwithAttachsDbResult`。
- 將 `Friend Sub InitDatabase` 改為 `Private Sub InitDatabase`。
- 將 `Friend Sub CloseDatabase` 改為 `Private Sub CloseDatabase`。
- 將 `Friend Async Function SaveCachesToSQLiteAsync` 改為 `Private`。
- 將 `Friend Async Function LoadCachesFromSQLiteAsync` 改為 `Private`。
- 將 `Friend Function GetDatabaseSummary` 改為 `Private` (若無外部 UI 需要)。
- 將 `Friend Sub CleanupOrphanFolderPath` 改為 `Private`。

#### [MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1.vb)
- 將 `Friend Structure StatusHistoryItem` 改為 `Private Structure StatusHistoryItem`。

## 開放式問題
1. **已解：DebugForm 引用狀況**：經稽核，DebugForm 僅存取 `Dbg` 與其內部字串，不存取上述結構，可放手收縮。
2. **已解：Module 引用**：`moduleStore.vb` 無現役外部引用，收縮安全。

## 驗證計畫

### 自動化測試
- 無（主要為編譯期權限變更）。

### 手動驗證
1. 執行編譯 (Build)，確認沒有發生「不可存取成員」的錯誤。
2. 特別檢查不同 Partial Class 檔案之間的互操作性是否依然正常。
