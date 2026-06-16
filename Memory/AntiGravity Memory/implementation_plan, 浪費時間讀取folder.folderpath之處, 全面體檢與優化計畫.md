# 全面體檢與優化計畫

根據您的需求，我將對整個系統針對以下四點進行全面體檢與修正：

## 1. 擴展 `fPath` 透傳至所有 `GetCachedxxxx` 函數 
目前 `_cache` 系列全依賴 `FolderPath` 當作鍵值，為了避免所有 L2.5 的查詢在命中的情況下仍浪費時間去呼叫 `folder.FolderPath` 的 COM 屬性，我會將 `Optional fPath As String = ""` 推行到所有的 `GetCached...`：
- `GetCachedMailCountAllAsync`
- `GetCachedFolderCount`
- `GetCachedFolderCountAllAsync`
- `GetCachedFolderSizeAsync`
- `GetCachedFolderSizeAllAsync`
- `GetCachedAttachMailList`

## 2. 檢視呼叫端是否能提供 `fPath`
- **可以提供的地方**：`CollectMonthCounts` 已經透過我們的修正提供了 `_tab2FolderPaths(i)` 給 `GetMonthCountsForYear`。連同內層呼叫的 `GetCachedMailCount` 也已經受惠。
- **後續能加上透傳的地方**：例如 `CollectYearCounts`，如果在其上層（如 `SimTree2_AfterSelect`）我們也是用陣列迭代的，我們會嘗試把先前產生的 `_tab2FolderPaths` 餵給 `GetCachedMailCount` 或是 `GetYearCountsForFolder`，將此邏輯橫向套用以降低 Tab2 點選初期的延遲。
- **較難以提供的地方**：像某些 DFS/BFS 遞迴或 `Tab 1` 的資料展開，若傳入引數依然只有 `Outlook.Folder` 物件而沒有維護平行的字串陣列，則該當下的第一次 `fPath` 解析是無法避免也無需勉強避免的（這是初始連線成本）。

## 3. 檢查 `cToken` Exception 漏網之魚
`OperationCanceledException` 或是 `TaskCanceledException` 只有在 `cToken` 觸發中斷時才會產生。
- **需要捕捉的地方**：任何作為 UI Event Handler 的最上層入口 (被標為 `Async Sub` 的函數)。如果這些底下的 Await 丟出 Cancel Exception 卻沒有被 Try Catch 接住，程式會閃退。
  我們已幫 `HandleListViewKeyPress` 修復，我會進一步盤點 `SimTree1_AfterSelect`、`SimTree3_AfterSelect` 等所有含有 `Await` 及傳遞 `cToken` 的第一線 UI Task 進入點。
- **不用捕捉的地方**：任何屬於內層商業邏輯的 `Async Function` (例如 `CollectMonthCounts` 或 `GetYearCountsForFolder`)。這些只要不在裡面 Try-Catch，例外就會自然「向上冒泡 (Bubble Up)」，由最外層的 UI Event (例如 DoubleClick 或 KeyPress) 統一接住並顯示「已中斷」，這是最乾淨的寫法。

## 4. 檢查同一函數內多次讀取 `FolderPath` & `Name`
在剛才的 grep 掃描中，我發現在 `Form1_Outlook.vb` 以及少部分 `Form1_Win32API.vb` 中，仍有直接用 `folder.Name` 或是呼叫 `folder.FolderPath` 兩次的痕跡。
- 我將巡查像是 `GetFolderStats` 或是 `GetLiveFolderSnap` 這些函式，把它們統一為：
  ```vb
  Dim fPath As String = folder.FolderPath
  Dim fName As String = fPath.Substring(fPath.LastIndexOf("\"c) + 1)
  ```
避免 `folder.Name` 這個完全能用字串切割取代的屬性再度觸發冗餘的 COM 呼叫。

### 接下來的動作
如果您同意這個計畫，我會立即開始執行程式碼的修改。請讓我知道！
