# 偵錯訊息格式標準化實作計畫

此計畫接續先前中斷的工作，目標是將全專案（主要為 `Form1.vb` 與 `Form1_ComL3.vb`）中的 `Dbg` 呼叫格式進行標準化。

## 使用者評論與回饋

> [!IMPORTANT]
> 根據先前討論，我們將移除 `Dbg` 訊息中冗餘的函數名稱與 `: ` 冒號分隔符。
> 由於 `DebugForm.GetCallerName()` 已經能自動識別呼叫者，因此 `Dbg("FunctionName: 開始")` 應簡化為 `Dbg("開始")`。

## 擬定變更

### 1. 全域偵錯邏輯擴充
目前 `Dbg` 函數已經整合了 `GetCallerName()`，我們將確保其在各處的調用都能發揮最大效用。

### 2. Form1.vb 標準化
將檔案中剩餘的舊格式 `Dbg` 呼叫進行清理。

#### [MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1.vb)
- 將 `Dbg("開始", sender.Name)` 修改為 `Dbg("開始")`。
- 將 `Dbg("結束", sender.Name)` 修改為 `Dbg("結束")`。
- 移除所有 `Dbg("開始: ...")` 中的冒號，並將冗餘的函數名移至第二參數（若仍有必要）。

### 3. Form1_ComL3.vb 標準化
此檔案包含許多底層資料存取邏輯，其 `Dbg` 訊息較為複雜，需逐一調整。

#### [MODIFY] [Form1_ComL3.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_ComL3.vb)
- 將 `Dbg("開始: FunctionName", folder.Name)` 修改為 `Dbg("開始", folder.Name)`。
- 處理 `Dbg("結束: ... (FAIL)", ...)` 等特殊訊息，優化其可讀性。

## 開放性問題

> [!NOTE]
> 在事件處理器（如 `Button_Click`）中，`sender.Name` 通常能提供是哪個按鈕被按下的資訊。雖然 `GetCallerName()` 能抓到函數名 `Button1_Click`，但如果多個按鈕共享同一個處理器，`sender.Name` 仍有其價值。
> **建議：** 若是共享處理器，保留 `sender.Name` 於第二參數；若是專屬處理器，則移除。

## 驗證計畫

### 自動化測試
- 使用 `grep` 或 `powershell` 指令確認檔案中是否仍殘留帶有冒號的 `Dbg` 呼叫。

### 手動驗證
- 執行程式並開啟 `DebugForm`，確認輸出的訊息是否如預期般乾淨且帶有正確的呼叫者名稱。
