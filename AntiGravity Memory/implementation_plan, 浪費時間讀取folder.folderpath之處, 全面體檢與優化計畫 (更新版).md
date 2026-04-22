# 全面體檢與優化計畫 (更新版)

根據您的指導，我修正並細化了體檢與優化目標如下：

## 1. 擴展 `fPath` 透傳至快取與 Layer 3 函數
- **Layer 2.5 (`GetCached...`)：** 會全部加上 `Optional fPath As String = ""`。
- **Layer 3 (原生存取)：** 會在直接被 L2.5 呼叫的「單層查詢接口」(如 `GetMailCount`) 加上 `Optional fPath`，因為這時候 L2.5 手上**已經擁有**這串字串，直接傳下去就能節省一次 COM 呼叫。
- **不勉強加的地方 (L3 的遞迴內部)：** 例如 `GetMailCountAll` 內部循序鑽取 `subFolders` 時，由於那些子資料夾尚未轉為字串路徑，強行索取反而還是要花 COM 的時間。維持這是「不可避免的初次連線成本」，所以不再硬求。

## 2. 檢視呼叫端是否能提供 `fPath`
- 已確認 `CollectMonthCounts` 完美受惠。
- 我會同步檢查 `CollectYearCounts` 是否能在 `Tab2` 點擊樹狀圖獲得平行陣列 `_tab2FolderPaths` 時，將這個陣列一併代入進去，避開千次 COM。
- 其餘只給出 `Outlook.Folder` 原生呼叫的地方 (例如 Tab 1 剛載入時的展開)，就不強求傳入字串。

## 3. `cToken` 中斷點梳理與 Exception 漏網之魚
您的觀察非常正確。以前因為用全域旗標 `_cancelRequested`，所以才會到處自己埋 `If _cancelRequested Then...`。現在改用真正的 `CancellationToken`，只要遇到 `Await Task.Delay(1, cToken)` 或 `Await Task.Run(..., cToken)` 就會自動拋出 `TaskCanceledException`，我們不需要自己到處寫 `If cToken.IsCancellationRequested` 來防堵了！

- **【精簡】刪除多餘中斷點**：
  我會巡視並刪除藏在內層迴圈 (如 `PreloadAttachByRDOAsync2` 等函數) 裡那些手寫的冗餘 `If cToken.IsCancellationRequested Then`，只在回報進度條或 `Task.Delay` 時讓系統原生中斷。

- **【補洞】必須加上 `Try...Catch` 的對象 ( UI 進入點全列出 )**：
  以下所有被標為 `Async Sub` 的最外層觸發器，我都必須包覆 `Try...Catch ex As OperationCanceledException` 來防止 WinForms 崩潰：
  1. `SimTree1_AfterSelect` (Tab 1 樹狀展開)
  2. `SimTree2_AfterSelect` (Tab 2 樹狀展開)
  3. `SimTree3_AfterSelect` (Tab 3 樹狀展開)
  4. `ListView2_MouseDoubleClick` (舊版已加過)
  5. `HandleListViewKeyPress` (剛剛已加過)
  6. `Button3_Click` (Tab 3 搜尋)
  7. `RenewCache_Click` (刷新快取按鈕)

## 4. 檢查同一函數內多次讀取 `FolderPath` & `Name`
包含但不限於 `Form1_Win32API.vb` 中的背景快取或是 `Form1_Outlook.vb` 裡的 `GetLiveFolderSnap` 等原生函式，若同一區塊內看見 `f.FolderPath` 又看見 `f.Name`：
我將把它們統一為：
```vb
Dim fPath As String = folder.FolderPath
Dim fName As String = fPath.Substring(fPath.LastIndexOf("\"c) + 1)
```
確保任何獨立執行的函數最多只花 $1$ 毫秒在 COM 字串解析上。

---
### 接下來的動作
這個修正計畫符合您的預期嗎？如果沒問題，我將立即動手為上述四點進行程式碼梳理！
