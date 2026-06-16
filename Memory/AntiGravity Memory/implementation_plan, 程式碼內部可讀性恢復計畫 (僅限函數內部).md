# 程式碼內部可讀性恢復計畫 (僅限函數內部)

根據您的最新指示，我將修正實施方針：**不變動 Region、不變動方法 (Sub/Function) 之間的間距**（因為您已經手動調整好了），而是專注於**恢復 Sub 與 Function 內部的邏輯空行**。

## 使用者核心指示 (User Review Required)

> [!IMPORTANT]
> **執行限制：**
> 1. **Region 不動**：保持 `#Region` 與 `#End Region` 及其周邊的現有間距。
> 2. **方法之間不動**：`End Sub` / `End Function` 與下一個方法開始之間的空行保持現狀，由您掌控。
> 3. **僅處理內部**：只在 `Sub...End Sub` 或 `Function...End Function` 的**內部**插入空行，用以區隔不同的邏輯步驟。
> 4. **保留註解**：絕對不刪除任何既有的註解（思考過程、debug 紀錄）。
> 5. **AntiGravity 註解標記**：若我有添加或大幅調整註解說明，會加上 `by AntiGravity, 2026/04/03`。

## 擬議變更 (Proposed Changes)

我將依序以「小塊寫入 (Chunked Edits)」處理以下檔案：

### [Outlook Assistant 專案]

#### [MODIFY] [DebugForm.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/DebugForm.vb)
- **優先處理**。例如在 `DebugForm_Shown` 內部，將各個步驟（1. 搜尋列, 2. 搜尋框對齊...）之間插入空行，增加層次感。
- 恢復 `AddMessage3` 內部邏輯（時間計算、行號遞增、Tag 產生）之間的間距。

#### [MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1.vb)
- 處理 `InitTabUI` 系列函數內部的控制項配置邏輯空行。
- 處理 `Form1_Load` 內部的初始化順序空行。

#### [MODIFY] [Form1_Main.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_Main.vb)
- 在龐大的統計邏輯（如 `ComputeFolderStatsAsync`）內部，於 BFS、讀取、彙總、寫快取各階段之間插入空行。

#### [MODIFY] [Form1_ComL3.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_ComL3.vb)
- 在 L3 底層函數內部，將 RDO、MAPI、OOM 各種 Fallback 嘗試區塊區隔開來。

## 開放性問題 (Open Questions)

> [!NOTE]
> 目前針對「單行短函數（One-liners）」仍維持不留內部空行的原則。

## 驗證計畫 (Verification Plan)

### 手動驗證
- 我將在每完成一個檔案的小塊修改後，確認其語法正確且視覺上符合「函數內部邏輯區分」的要求。
- 確認所有原本的開發筆記均完整保留。
