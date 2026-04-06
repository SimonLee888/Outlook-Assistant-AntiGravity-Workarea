# Outlook Assistant — 開發記憶包
## 供新 Chat 接手使用 — 產出於 2026-04-05 11:20

---

## 專案基本資訊

- **語言/框架**：VB.NET / WinForms，.NET 10 (`net10.0-windows10.0.17763.0`)
- **環境**：Outlook LTSC 2021（olNotExchange）、本機 PST 檔、VS 2022
- **主要檔案（2026-04-05 重構後）**：
  - `Form1.vb`（1438行）：全域宣告、Form 生命週期、外觀初始化
  - `Form1_Main.vb`（2256行）：Tab1~Tab5 事件與業務邏輯
  - `Form1_ComL3.vb`（1749行）：Outlook 物件初始化、L2.5 快取代理、L3 COM 底層函數
  - `Form1_SimTree.Designer.vb`（26行）：SimTree 設計工具描述
  - `Form1_SimTree.vb`（416行）：多選 TreeView 控制項（原 SimTree.vb，改名為 Form1_SimTree，類別名稱也改為 `Form1_SimTree`）
  - `Form1_Win32API.vb`（110行）：Win32 API 宣告集中管理
  - `Form1.Designer.vb`（1177行）：UI 設計工具產生
  - `DebugForm.vb`（761行）：大幅擴充，新增 `GetCallerName()` 支援 Async 解析
  - `moduleStore.vb`（1416行）：舊版程式碼保留區（大部分已 comment out）

---

## ⚠️ 重大架構變更（與舊 memory 不同之處）

### 1. SimTree 類別已改名
- **舊名稱**：`SimTree`（SimTree.vb）
- **新名稱**：`Form1_SimTree`（Form1_SimTree.vb，Inherits TreeView）
- Form1 中宣告方式：
  ```vb
  Private WithEvents SimTree1 As New Form1_SimTree
  Private WithEvents SimTree2 As New Form1_SimTree
  Private WithEvents SimTree3 As New Form1_SimTree
  Private WithEvents SimTree4 As New Form1_SimTree
  ```
- `Form1_SimTree` 新增 `<Browsable(False)>` 和 `<DesignerSerializationVisibility(Hidden)>` attribute 在 `Shadows SelectedNode` 上，修正設計工具序列化紅色警告

### 2. 狀態列控制項改名
- **舊名稱**：`lblStatus1`、`lblStatus2`
- **新名稱**：`ProgressBar1`（300px）、`ProgressBar2`（480px）
- 類型仍是 `ToolStripStatusLabel`，只是改名反映其語意用途

### 3. 快取字典全部改名（key 已全部改為 FolderPath 字串）
```vb
' Form1_ComL3.vb 中定義（Shared ReadOnly ConcurrentDictionary）
Private Shared ReadOnly _cacheMailCount As New ConcurrentDictionary(Of String, Integer)
Private Shared ReadOnly _cacheMailCountAll As New ConcurrentDictionary(Of String, Integer)
Private Shared ReadOnly _cacheFolderCount As New ConcurrentDictionary(Of String, Integer)
Private Shared ReadOnly _cacheFolderCountAll As New ConcurrentDictionary(Of String, Integer)
Private Shared ReadOnly _cacheFolderSize As New ConcurrentDictionary(Of String, Long)
Private Shared ReadOnly _cacheFolderSizeAll As New ConcurrentDictionary(Of String, Long)
Private Shared ReadOnly _cacheIsMailFolder As New ConcurrentDictionary(Of String, Boolean)
Private Shared ReadOnly _cacheFolderTree As New ConcurrentDictionary(Of String, List(Of Outlook.Folder))
Private Shared ReadOnly _cacheSubFolderList As New ConcurrentDictionary(Of String, List(Of Outlook.Folder))
Private Shared ReadOnly _cacheAttachFilename As New ConcurrentDictionary(Of String, List(Of String))

' Form1_Main.vb 中定義
Private Shared ReadOnly _yearCountsCache As New ConcurrentDictionary(Of String, ConcurrentDictionary(Of Integer, Integer))
Private Shared ReadOnly _monthCountsCache As New ConcurrentDictionary(Of String, ConcurrentDictionary(Of Integer, Integer))
Private _cachePhase1tab3 As New Dictionary(Of String, FolderCacheTab3)
```
**重要**：所有快取 key 現在都是 FolderPath 字串，不再用 COM 物件當 key（已解決舊版 RCW 殘留問題）

### 4. L2.5 快取代理層正式成型
所有 L3 COM 呼叫都透過 L2.5 proxy function，L1/L2 不直接呼叫 L3：
```vb
GetCachedMailCount(folder)           ' 單一資料夾郵件數
GetCachedFolderCount(folder)         ' 單一資料夾子資料夾數
GetCachedMailCountAllAsync(folder)   ' 整棵子樹郵件總數 (Async)
GetCachedFolderCountAllAsync(folder) ' 整棵子樹資料夾總數 (Async)
GetCachedFolderSizeAsync(folder)     ' 單一資料夾大小 (Async)
GetCachedFolderSizeAllAsync(folder)  ' 整棵子樹大小 (Async)
GetCachedMailWithAttachment(folder, progress)  ' Tab3 Phase1 含附件郵件 (Async)
GetCachedAttachFilename(mail)        ' 附件檔名清單 (同步，因有 RDO 平行預載)
```

### 5. Redemption 整合狀態（目前關閉）
- `InitRdoSession()` 存在但在 `Form1_Load` 中被 comment out（`'InitRdoSession()`）
- 保留完整實作（`AutoDismissRedemptionDialog`、`ManualResetEventSlim` 同步點）
- 原因：效能測試後發現演算法改進的效果比 RDO 更顯著，RDO 徒增複雜度與啟動時間
- `PreloadAttachmentCacheRDOAsync` 仍保留，當 `_rdo IsNot Nothing` 時自動啟用 RDO 平行預載

### 6. Tab1 架構升級至 v5
`ComputeFolderStatsAsync` 從百行巨型函數拆分為五個子函數：
- `BuildBfsFolderTree`：BFS 展開 + 快取剪枝
- `FetchDirectMailCountsAsync`：呼叫 L3，含 progress 與 `_cancelRequested` 支援
- `SummarizeSubTreeBottomUp`：純記憶體底部向上加總
- `UpdateFolderStatsCache`：寫入 L2.5 快取字典
- `GetBfsResult`：提取 root + 直屬子資料夾供 L1 顯示

### 7. Tab2 移除 TreeView2_AfterSelect
- 完全由 `SimTree2_AfterSelect` 取代
- `CheckSub2` 更名為 `CheckSubFolder2`（需確認 Designer 中的實際名稱）

### 8. Tab3 架構升級至 v3（管線化 Pipeline）
```
Button3_Click 管線:
  Step 1. 驗證大小參數
  Step 2. GetSubFolderList (BFS)
  Step 3. GetCachedMailWithAttachment → L2.5 快取 or L3 GetMailWithAttachment
  Step 4. FilterBySize (純記憶體 LINQ)
  Step 5. FilterByAttachmentDetailsAsync → PreloadAttachmentCacheRDOAsync (RDO 平行預載) + _cacheAttachFilename
  Step 6. ShowResultTab3
```
- `blnButton3_Stop` 已移除，統一使用 `_cancelRequested`
- `Button8` 舊架構已廢棄（Button3 現在就是新架構）

### 9. Tab4 系列郵件已初步實作
- 使用 `Button4_Click`，讀取 `GetSubFolderList + GetTable` 掃描 `PR_CONVERSATION_TOPIC`
- 依主旨分群，顯示系列郵件 topic → ListView4

### 10. 新增 Tab5 重複郵件（初步可用）
- `Button5_Click`，比對相似度（Exact / Fuzzy 兩種模式）
- Fuzzy 模式用 `LevenshteinDistance` 計算編輯距離（閾值 0.8 相似度）
- UI：`ListView5`（動態建立）、`rbExactMatch`、`rbFuzzyMatch`

### 11. 全域新增 checkIncludeAllFolders
- 勾選時 `GetSortedSubFolders` 顯示全部資料夾（含行事曆/聯絡人等非郵件資料夾）
- 未勾選時只顯示 `IsMailFolder()` 回傳 True 的資料夾

### 12. 新增 ThemeColors 靜態色彩管理
```vb
ThemeColors.Gray95        ' #F2F2F2 主背景
ThemeColors.MercuryGray   ' #E5E5E5 Hover 背景
ThemeColors.AltoGray      ' #E0E0E0 格線/邊框
ThemeColors.Brand_Blue    ' #0078D4 品牌藍
ThemeColors.CoralRed      ' #D83933
ThemeColors.RustRed       ' #A22C29
ThemeColors.DeepAmber     ' #E67E22 平均線/參考線
```

### 13. StatusHistory 彈出歷史功能
- `_statusHistory As List(Of StatusHistoryItem)`，最多 100 筆
- 點擊 `ProgressBar1` 或 `ProgressBar2` 標籤彈出 ToolStripDropDown 顯示歷史
- 最新一筆在底部，自動捲動到底部

### 14. Form_Shown 延遲載入各 Tab
```vb
Private Async Sub Form1_Shown(...)
    ' 依序延遲初始化 Tab2~Tab5，避免啟動時搶資源
    ' _isTabInitialized(N) 記錄每個 Tab 的初始化狀態
    ' WaitAndYieldIfBusy() 確保使用者操作時暫緩預載
```
- `_isTabInitialized(0)`：Form 第一次啟動中旗標（True=啟動中）
- `_isTabInitialized(1~5)`：Tab1~5 UI 是否已初始化
- `_isUserBusy`：使用者操作忙碌旗標，背景預載時讓步

### 15. Dbg() 升級
```vb
<System.Diagnostics.Conditional("DEBUG")>
Private Sub Dbg(Optional msg As String = "", Optional detail As String = "")
    Dim realCaller As String = DebugForm.GetCallerName()  ' ← 新：支援 Async 方法名稱解析
    If _isDebugMode Then DebugForm.AddMessage3(msg, detail, realCaller)
End Sub
```
- `_isDebugMode` 由 `#If DEBUG Then` 自動設定
- `_iLikeNoisy`：控制是否顯示高頻迴圈訊息（預設 False）
- `OKiLikeNoisy` 勾選框動態切換

### 16. Win32API 集中管理
`Form1_Win32API.vb` 集中所有 DllImport 宣告（SendMessage, PostMessage, FindWindow, FindWindowEx, ShowWindow 等）

---

## 全域架構原則（不變）

### COM / 執行緒規則
- 所有 Outlook COM 呼叫必須在 UI 執行緒（STA）
- 不可用 Task.Run 包住 COM 物件
- 用 `Await Task.Delay(1)` 讓 UI 不凍結（注意：`Task.Yield()` 在高頻迴圈中過重，改用 `Delay(1)`）
- Redemption 是 free-threaded，可在 Task.Run / Parallel.ForEach 中呼叫
- `_cancelRequested` 全域旗標統一控制 ESC 中斷，不再有各 Tab 專屬旗標

### Outlook Namespace 衝突
- `System.Exception` 必須寫全名（不可只寫 `Exception`）
- `System.Windows.Forms.View.Details` 不可縮寫
- ListView `View` 屬性等衝突名稱都要加全名前綴

### L3 函數設計原則
- Fallback 鏈：⓪ Redemption → ① MAPI PropertyAccessor → ② OOM → ③ return -1
- 失敗統一回傳 -1（不回 0），讓 L2 能區分「真的是 0」或「讀取失敗」
- Finally 中統一釋放 COM 物件（`TryMarshalRelease()`）
- 成功路徑靜默（不輸出 Dbg），只在錯誤路徑輸出

---

## 結構體與型別定義（Form1_ComL3.vb）

```vb
Public Structure MailItemInfo   ' Tab3 候選郵件（純 .NET，不帶 COM）
    Dim EntryID As String
    Dim Subject As String
    Dim Size As Long
    Dim ReceivedTime As DateTime
    Dim SenderName As String
    Dim AttachCount As Integer
End Structure

Public Structure L3ProgressReport  ' 統一進度回報
    Dim CurrentCount As Integer
    Dim TotalCount As Integer
    Dim Message As String
    Dim IsIndeterminate As Boolean
End Structure

Private Class FolderBfsEntry       ' BFS 遍歷用容器
    Public Folder As Outlook.Folder
    Public ParentIndex As Integer   ' -1=root
    Public DirectMailCount As Integer
    Public TotalMailCount As Integer
    Public TotalSubCount As Integer
    Public IsFromCache As Boolean
End Class

Private Structure FolderCacheTab3  ' Tab3 Phase1 快取
    Dim mailWithAttachment As List(Of MailItemInfo)
    Dim ItemCountWhenCached As Integer
End Structure
```

---

## 已確認的死路（PST 限制，勿再嘗試）

- DASL/Restrict 無法篩附件檔名（PST 無索引）
- `AdvancedSearch` 在 LTSC 失敗（HResult=0x8007064F）
- `SetColumns` 加 `PR_ATTACH_LONG_FILENAME` → NullReference（附件是子物件，非 MailItem 頂層屬性）
- `PR_MESSAGE_SIZE_EXTENDED (0x0E080014)` 和 `PR_MESSAGE_SIZE (0x0E080003)` 在 PST folder object 均回「未知或找不到」
- `GetTable + GetArray()` 無法取附件檔名（附件在 message row 的下層）

---

## 待辦清單（依優先順序）

| 優先 | 項目 | 說明 |
|---|---|---|
| 高 | **SQLite 持久化快取** | 詳見下方「SQLite 快取設計討論」 |
| 中 | Tab2 Bug A | 首次切 Tab2 焦點停最上面，多次嘗試未解 |
| 中 | CacheSniffer 恢復 | Form1_Load 末尾那行被 comment out，待確認啟用 |
| 低 | Tab4 系列郵件 UI | 基本功能已有，但 UI 細節未完成 |
| 低 | Tab5 重複郵件 | Levenshtein 已實作，UI 尚未完善 |
| 低 | FastStringSimilarity.vb | SimHash + Bigram Jaccard，設計完成但尚未整合進 Tab5 |

---

## SQLite 快取設計討論（2026-04-05 待決策）

以下三個問題需使用者確認後才能動手：

**Q1. Tab3 Phase1（`_cachePhase1tab3`）與 Phase2（`_cacheAttachFilename`）分開或合一張表？**
- 建議分開：`mail_basic`（Phase1）、`mail_attachments`（Phase2）
- Phase2 不是每次執行，合一張會讓 Phase1 快取的讀寫複雜化

**Q2. Tab2 年份分布快取（`_yearCountsCache`、`_monthCountsCache`）是否一起存 SQLite？**
- 這是衍生計算結果（非直接 MAPI 屬性），可從 mail_basic 重新計算
- 存進 SQLite 可省計算時間，但增加 schema 複雜度

**Q3. 目前 PST 數量大概幾個？**
- 影響 SQLite schema 是否需要 `store_id` 欄位區分不同 PST 的資料

**已確認的設計方向：**
- 格式：SQLite（`Microsoft.Data.Sqlite` NuGet）
- 架構：L2.5 proxy function 內插入 SQLite 讀寫，L1/L2/L3 完全不知道 SQLite 的存在
- 寫入時機：增量式（L2.5 完成一個資料夾就立刻 commit），FormClosing 確保 flush
- 快取有效性：`GetArray()` 快速取 EntryID + `PR_LAST_MODIFICATION_TIME` 比對
  - 新郵件（EntryID 不在快取）→ 待抓取
  - 已刪除（快取有但 GetArray 沒有）→ 移除
  - 已修改（EntryID 相同但 ModTime 不同）→ 待更新
  - 只針對「待抓取」和「待更新」的郵件做 Phase2 COM 呼叫
- 啟動時：背景非同步讀入記憶體快取，不阻塞 UI（50-150ms，使用者無感）
- 容量估算：500-1000 資料夾 + 10,000 封郵件 ≈ 5-7 MB SQLite 檔案（無需壓縮）
- 不加 PR_SEARCH_KEY：PST 郵件移動後 EntryID 改變，比對邏輯中「新郵件+已刪除」就能正確處理，不需要額外一層

---

## 使用者偏好（重要）

- ✅ 偏好直接給程式碼，在瀏覽器頁面解釋做法跟原理，在程式碼的註解記錄清楚
- ✅ 詳細中文註解，說明設計意圖與決策理由，附日期
- ✅ 壓縮 VB.NET 語法（單行 If/Then、冒號分隔），不要展開
- ✅ 修改時原有 Dbg/註解不可遺失
- ✅ `System.Exception` 永遠寫全名
- ✅ 不可縮寫 namespace（`System.Windows.Forms.View.Details` 等）
- ✅ 每次輸出前確認沒有遺漏原有內容
- ❌ 不用 `<summary>/<remarks>` XML doc comment 格式
- ❌ 不使用韓文或簡體中文，回覆用繁體中文+英文
- ❌ 討論時不要浪費 token，大型任務先拆分讓使用者確認再動手
