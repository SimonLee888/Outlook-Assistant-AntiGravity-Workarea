# Outlook 效能優化精鍊計畫：殘餘任務 (v2)

本計畫旨在處理剩餘的微量優化點。依據使用者指令，**不進行死碼清理**，保留過往程式碼。

## 使用者 review 重點

> [!IMPORTANT]
> **本次優化亮點：**
> 1. **極速 RDO (Redemption) Tuple 化**：將 RDO 分支的資料夾展開邏輯也改為 Tuple 模式，實現與 OOM 分支持平的「零重複屬性讀取」。
> 2. **跨層級屬性 Bundle 傳遞**：在 `RenewCacheAsync` 流程中，將單個資料夾的多個 COM 讀取動作，整合成一次性讀取。

## 擬議改動內容

### 1. [核心通訊層] Form1_Outlook.vb

#### [MODIFY] [GetSubtreeToListL3_Rdo](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_Outlook.vb)
- **改動細節**：
    - 將 Queue 與回傳 List 改為 `(Folder As RDOFolder, Path As String)`。
    - 在遍歷 `current.Folders` 時，立即拼接 `childPath = current.Path & "\" & subFolder.Name`。
- **效益**：杜絕後續對 `.FolderPath` 的 COM 讀取，尤其在數萬資料夾情境下提速顯著。

#### [MODIFY] [GetAttachMailListL3](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_Outlook.vb)
- **改動細節**：加入 `Optional fPath As String = ""` 參數，內部的 `fName` 改由 `ExtractFolderName(fPath)` 取得。

---

### 2. [持久化層] Form1_SQLite2.vb

#### [MODIFY] [RenewCacheAsync / RenewAttachMailListAsync](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_SQLite2.vb)
- **改動細節**：
    - 在 `RenewCacheAsync` 的 Phase 3 迴圈內，傳遞 `fPath` 給 `RenewAttachMailListAsync`。
    - 優化 Phase 3 的多重屬性讀取。

## 驗證計畫

### 自動化測試
- 測試針對同一批資料夾，RDO 版與 OOM 版產出的路徑清單完全一致。

### 手動驗證
- 執行「更新資料庫快取」，觀察 Debug 視窗是否還有不必要的 `.Name` 或 `.FolderPath` 讀取紀錄。
