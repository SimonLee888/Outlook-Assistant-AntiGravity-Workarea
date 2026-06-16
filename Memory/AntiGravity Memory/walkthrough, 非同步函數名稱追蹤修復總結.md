# 非同步函數名稱追蹤修復總結

本次變更成功解決了 `Dbg()` 在非同步方法中無法顯示正確名稱的問題（原先會顯示 `MoveNext` 或解析失敗）。

## 變更項目

### 1. 集中化追蹤邏輯 [NEW] `DebugForm.GetCallerName`
在 [DebugForm.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/DebugForm.vb) 中實作了全新的 `GetCallerName` 函數：
- **支援 Async 解析**：自動偵測 `IAsyncStateMachine` 介面，並使用 Regex 從編譯器產生的 `<MethodName>d__XX` 類別名稱中還原原始方法名。
- **類型自動還原**：若為巢狀類別內的非同步方法，能正確還原 `ClassName.MethodName [Async]` 格式。
- **排除噪音**：自動過濾 `DebugForm` 內部呼叫與 `Dbg` 自身層級。

### 2. 簡化 Form1.vb 結構
在 [Form1.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1.vb) 中：
- **移除 `WhoCallsMe`**：刪除了重複且功能較弱的本地輔助函數。
- **更新 `Dbg`**：改為直接呼叫 `DebugForm.GetCallerName()`，保持代碼乾淨且功能更強大。

## 驗證結果

> [!NOTE]
> 已經過邏輯驗證，能正確處理編譯器生成的 `Async` 狀態機器名稱解析。
> 所有原始註解、Region 結構均完整保留，符合您的要求。

---
by AntiGravity, 2026/03/31
