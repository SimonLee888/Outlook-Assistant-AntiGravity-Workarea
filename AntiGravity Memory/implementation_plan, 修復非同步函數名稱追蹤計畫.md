# 修復非同步函數名稱追蹤計畫

目前 `WhoCallsMe` 邏輯在遇到 `Async` 方法時，會因為編譯器生成的 `MoveNext` 狀態機器而無法正確抓取原始的方法名稱。本計畫將依照之前的設計，將追蹤邏輯集中至 `DebugForm.vb` 並強化非同步解析能力。

## User Review Required

> [!IMPORTANT]
> 此變更會將原先散落在 `Form1.vb` 的追蹤邏輯移除，改為統一呼叫 `DebugForm.GetCallerName`。這有助於後續維護並減少重複代碼。

## Proposed Changes

### DebugForm 組件

#### [MODIFY] [DebugForm.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/DebugForm.vb)
*   **實作 `Public Shared Function GetCallerName`**：
    *   遍歷 `StackTrace`。
    *   偵測是否為 `MoveNext` 且其 `DeclaringType` 實作了 `IAsyncStateMachine`。
    *   使用 Regex 解析 `<MethodName>d__XX` 格式，從中擷取真正的 `MethodName`。
    *   返回格式：`ClassName.MethodName` (若是 Async 則額外標記，例如 `ClassName.MethodName [Async]`)。
*   **刪除舊有的 `GetActualCallingMethod`**（私人方法且邏輯過時）。

### 主表單組件

#### [MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1.vb)
*   **更新 `Dbg` 方法**：直接呼叫 `DebugForm.GetCallerName()`。
*   **刪除本地 `WhoCallsMe` 函數**。

## Open Questions

1. **Async 標記方式**：在回傳的名稱後方是否需要標註 `[Async]`？目前的計畫是加上標註以便識別。
2. **效能考慮**：解析名稱會用到簡單的字串處理或 Regex，對於大量 Dbg 呼叫是否可以接受？（通常 Debug 模式下效能非首要考量，且解析僅在 Dbg 觸發時執行一次）。

## Verification Plan

### 自動化/手動驗證
- [ ] **同步方法測試**：呼叫 `Dbg()`，確認 `DebugForm` 顯示正確的 `Form1.MethodName`。
- [ ] **非同步方法測試**：在 `Async Function` 內呼叫 `Await ...` 之後再呼叫 `Dbg()`，確認能解析出原始方法名而非 `MoveNext`。
- [ ] **多層呼叫測試**：測試從不同的類別呼叫，確保 `ClassName` 切換正確。
- [ ] **Debug 模式驗證**：確認這些變更僅在 `DEBUG` 條件編譯下運作穩定。
