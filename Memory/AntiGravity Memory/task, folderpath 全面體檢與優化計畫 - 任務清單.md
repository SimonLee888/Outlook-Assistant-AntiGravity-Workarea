# 全面體檢與優化計畫 - 任務清單

### 1. 擴展 `fPath` 透傳至快取與 Layer 3 函數
- [x] 替 `GetCachedMailCountAllAsync` 加上 `Optional fPath As String = ""` 並套用。
- [x] 替 `GetCachedFolderCount` 加上 `Optional fPath As String = ""` 並套用。
- [x] 替 `GetCachedFolderCountAllAsync` 加上 `Optional fPath As String = ""` 並套用。
- [x] 替 `GetCachedFolderSizeAsync` 加上 `Optional fPath As String = ""` 並套用。
- [x] 替 `GetCachedFolderSizeAllAsync` 加上 `Optional fPath As String = ""` 並套用。
- [x] 替 `GetCachedAttachMailList` 加上 `Optional fPath As String = ""` 並套用。

### 2. 檢視呼叫端是否能提供 `fPath`
- [x] 檢查並更新 `CollectYearCounts`，讓它支援接收與傳遞平行陣列 `_tab2FolderPaths`。
- [x] 檢查並更新能夠一次性準備好 `fPath` 的其他 L2/UI 呼叫。

### 3. `cToken` 中斷點梳理與 Exception 漏網之魚
- [x] 在 `SimTree1_AfterSelect` 加入 `Try...Catch ex As OperationCanceledException`。
- [x] 在 `SimTree2_AfterSelect` 加入 `Try...Catch ex As OperationCanceledException`。
- [x] 在 `SimTree3_AfterSelect` 加入 `Try...Catch ex As OperationCanceledException`。
- [x] 在 `Button3_Click` 加入 `Try...Catch ex As OperationCanceledException`。
- [x] 在 `RenewCache_Click` 加入 `Try...Catch ex As OperationCanceledException`。
- [x] 清理 `PreloadAttachByRDOAsync2` 與其他內層手寫的多餘 `If cToken.IsCancellationRequested Then`。

### 4. 檢查同一函數內多次讀取FolderPath & Name
- [x] 檢查並優化 `Form1_Outlook.vb` 內的重複屬性讀取（如 `GetLiveFolderSnap` 等）。
- [x] 檢查並優化 `Form1_Win32API.vb` 內的重複屬性讀取。
- [x] 檢查並優化 `Form1_SQLite2.vb` 和 `Form1_MainTabs.vb` 內的重複屬性讀取。
- [ ] 重新測試 Tab2 從年度視圖進入月份視圖的效能，驗證是否已達到 10ms 以內的記憶體運算水準 (待使用者驗證)。
