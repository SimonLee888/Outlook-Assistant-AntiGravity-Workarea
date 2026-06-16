# Outlook 效能優化精鍊計畫：殘餘任務與清理

根據對「Outlook 資料夾屬性讀取效能優化計畫」的查核，核心效能紅利已全數取得。本計畫旨在處理剩餘的微量優化點以及因架構演進而產生的冗餘程式碼。

## 使用者回顧與決策

> [!IMPORTANT]
> **主要優化方向：**
> 1. **極速 RDO (Redemption)**：將 RDO 分支的資料夾展開邏輯也改為 Tuple 模式，實現與 OOM 分支持平的「零重複屬性讀取」。
> 2. **專案瘦身 (Dead Code Removal)**：移除已被 BFS 剪枝邏輯取代的舊型遞迴統計函數（如 `GetMailCountAllAsync`），減少維護負擔。

## 擬議改動內容

### 1. [核心通訊層] Form1_Outlook.vb

#### [MODIFY] [GetSubtreeToListL3_Rdo](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_Outlook.vb)
- **改動內容**：
    - 將 `Queue(Of RDOFolder)` 提升為 `Queue(Of (Folder, String))`。
    - 在 BFS 展開時使用路徑拼接，而非重複讀取 RDO 屬性。
    - 將回傳類型改為 `List(Of (Folder As RDOFolder, Path As String))`。
- **效益**：當 RDO 掃描數萬個資料夾時，效能可再提升 15~20%。

#### [DELETE] [GetMailCountAllAsync / GetFolderCountAllAsync](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_Outlook.vb)
- **改動內容**：
    - 移除已標註為死碼的 L2.5 與 L3 相關函數。
- **效益**：確保開發者不會誤用舊的、較慢的統計方法。

---

### 2. [持久化層] Form1_SQLite2.vb

#### [MODIFY] [RenewAttachMailListAsync](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_SQLite2.vb)
- **改動內容**：
    - 檢查內部迴圈，確保所有輔助訊息輸出皆使用傳入的 `fPath` 與 `ExtractFolderName`，而非呼叫 `folder.Name`。

## 開放問題

> [!NOTE]
> 1. **RDO 相依性**: 目前 RDO 在部分環境可能未掛載，優化後的 Tuple 回傳是否需要針對全域穩定性做額外檢查？（預計維持目前的 Try-Catch 包裝即可）。

## 驗證計畫

### 自動化測試
- 測試針對同一批資料夾，RDO 版與 OOM 版產出的路徑清單完全一致（除根路徑撇線外）。

### 手動驗證
- 執行「更新資料庫快取」，確認掃描過程無任何 `.Name` 或 `.FolderPath` 引發的 COM 例外或延遲感。
