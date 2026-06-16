# Form1.vb Dbg() 偵錯紀錄標準化實作計畫

本計畫旨在優化與標準化 [Form1.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1.vb) 中的 `Dbg()` 偵錯輸出，確保紀錄語意一致、格式統一，並能提供精確的效能監控數據。

## 使用者審核請求

> [!IMPORTANT]
> **標準化規範與時間紀錄策略：**
> 1. **時間戳記 (Timestamp)**：由 `DebugForm` 在接收訊息時自動擷取系統時間，主程式 **不需要** 傳送當前時間。
> 2. **執行耗時 (Duration)**：建議由主程式使用 `Stopwatch` 紀錄關鍵段落的精確耗時，並在「結束」紀錄中傳出（例如：`Dbg("結束", $"耗時 {sw.ElapsedMilliseconds}ms")`）。
>    - *原因*：`DebugForm` 雖有紀錄間隔時間，但在非同步或併發環境下，間隔時間不等於單一函式的執行長度。
> 3. **移除冗餘名稱**：移除 `msg` 參數中重複手寫的函式名稱。
> 4. **統一耗時單位**：全域偵錯輸出改用毫秒 (`ms`)。

## 擬定變更內容

### 1. 全域規範統一
- **大小寫修正**：將所有 `dbg(...)` 統一替換為 `Dbg(...)`。
- **標記格式**：
    - 起始：`Dbg("開始", [關鍵上下文參數])`
    - 結束：`Dbg("結束", [統計資訊/耗時])`
    - 錯誤：`Dbg("Error: [描述]", ex.Message)`

---

### 2. 分層實作

#### [L1 UI 事件層]
- **目標**：確保使用者操作的起點與終點皆有紀錄。
- **異動檔案**：[Form1.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1.vb) (L1 Region)
    - `TreeView1_AfterSelect`: 補強起始/結束紀錄，確保 `sw.ElapsedMilliseconds` 被正確紀錄。
    - `ListView1_MouseDoubleClick` / `Redemption_Click` 等：補開起始紀錄。

#### [L2 流程協調層 (Compute...Async)]
- **目標**：精確監控複雜流程的進度與總體耗時。
- **異動檔案**：[Form1.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1.vb) (L2 Region)
    - `ComputeFolderStatsAsync`: 標準化 BFS 階段性的 `Dbg` 輸出。
    - `ComputeYearCounts`: 統一快取命中 (Cache Hit/Miss) 的語義。

#### [L3 底層資料層 & 輔助函數]
- **目標**：在 Fallback 鏈的每一層提供清晰的成功/流轉紀錄，並附帶 `Stopwatch` 數據。
- **異動檔案**：[Form1.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1.vb) (L3 & 輔助函數 Region)
    - `GetMailCount`, `GetFolderSize`: 統一 `⓪ RDO`, `① MAPI`, `② OOM` 的成功與失敗格式。
    - `FindNodeByName`, `GetFolderByName`: 加入遺漏的結束紀錄。

---

## 驗證計畫

### 自動化驗證 (偵錯主控台檢查)
1. **啟動偵錯模式**：確保 `CheckDebug` 已勾選。
2. **操作 Tab1 點選資料夾**：
    - 檢查 `DebugForm` 是否出現完整的 `開始` -> `結束` 序列。
    - 確認 `Source` 欄位顯示正確的函式名，且 `Msg` 欄位不再有重複的函式名。
    - 確認 `Details` 欄位包含正確的資料夾名稱與 `ms` 耗時。
3. **測試中斷功能 (ESC)**：
    - 在耗時統計期間按下 ESC。
    - 確認 `DebugForm` 紀錄中出現 `已中斷`。

### 手動驗證 (UI 一致性)
- 確認 `lblStatus2.Text` 顯示的秒數與 `Dbg()` 紀錄中的毫秒數在語意上對應（例如 2.50s 相對應於約 2500ms）。
