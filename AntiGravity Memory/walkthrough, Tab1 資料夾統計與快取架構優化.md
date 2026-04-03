# Tab1 資料夾統計與快取架構優化

這份文件總結了為解決 Tab1 TreeView 展開延遲與 A/B 目錄切換時重新讀取 RDO 的問題，所進行的一系列架構與快取優化。

## 變更摘要 (Changes Made)

本次共進行了三項主要的快取架構更新，徹底解決了無效的 COM 呼叫並提升了多次切換時的 UI 反應速度：

### 1. 將原有快取的 Key 由 `Outlook.Folder` 全面改為 `String` (FolderPath)
由於 Outlook COM 的 RCW 機制導致相同資料夾每次被讀取時會在記憶體產生不同參考位址的 Wrapper，使得 `_mailCountCache.ContainsKey(folderObj)` 經常失效 (Cache Miss)。
因此已將 [Form1.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1.vb) 中所有與資料夾相關的集合，全面採用 `folder.FolderPath` 作為唯一字串鍵值：
* `_mailCountCache`
* `_folderCountCache`
* `_folderSizeCache`
* `_folderTreeCache`
* 此變更同時同步修改了 `GetTotalFolderCountAsync`, `GetMailCountByMAPINew`, `GetFolderSizeAsync` 等所有依賴這些快取的函數。

### 2. 建立 L2 的直屬資料夾數量快取 (`_directFolderCountCache`)
* **問題**：過去 `LoadSubFolderToTreeView` 為了決定是否繪製 `+` 號，會對每個子節點呼叫 `GetFolderCount(folder)`，而這個函數因為沒有快取，每次都會向 RDO 發出 COM 請求。
* **解法**：為保持 L3 函數的「純資料讀取」性質，我們在 L2 新增了 `GetCachedFolderCount` 來進行快取管理。現在展開節點時會優先檢查字串快取 `_directFolderCountCache`，只有第一次會呼叫 RDO。

### 3. 建立 L2 的直屬郵件數量快取 (`_directMailCountCache`)
* **問題**：`ComputeFolderStatsAsync` 的 BFS 掃描原先只會快取「總信件數 (含子孫)」。當命中快取進入 Step 5 回傳給 ListView 顯示時，卻發現缺少「本層信件數」，導致被迫再向 RDO 發出 `GetMailCount` 呼叫。
* **解法**：新增 `_directMailCountCache` 專門記錄本層信件數。Step 2 掃描時順便寫入這份快取；Step 5 回傳時，即優先從記憶體讀取本層數值。切換 A/B 目錄時即可瞬間載入，完全零 COM 呼叫。

### 4. 實作統一的 L2 快取代理層 (Cache Proxy Layer) 重構
* **問題**：原本在業務邏輯如 `CacheSniffer` 或 `ComputeFolderStatsAsync` 中，充滿了手動判斷 `_cache.TryGetValue` 與寫入 `_cache.TryAdd` 的邏輯，讓程式碼顯得雜亂且責任界線不清。
* **解法**：我們將快取邏輯抽出，全面封裝成獨立的 `GetCachedXxx` 函數群。
  * `GetCachedFolderCount`
  * `GetCachedMailCount`
  * `GetCachedTotalMailCountAsync`
  * `GetCachedTotalFolderCountAsync`
* 上層業務邏輯只需要向代理層「拿資料」，完全不用管是不是拿快取；如果快取沒命中，代理層會自動原封不動地把 `onProgress` 等委派遞交給下層 L3 的 `GetMailCount` 或 `GetMailCountAll` 去向 RDO 或 OOM 取值。如此兼顧了效能、平行處理安全與架構乾淨度。

## 程式碼註解政策
所有新增或變更的邏輯都已遵照您的規定加上了類似以下的註解以利未來追蹤：
* `(by AntiGravity, 2026/03/27 修正快取鍵值型別)`
* `(by AntiGravity, 2026/03/27 修正 A/B 目錄切換再讀取問題)`

## 後續驗證回報 (Validation Results)
請您編譯並執行應用程式，測試以下操作：
1. **點擊並展開任意大型子資料夾**：只會在第一次時看見 Debug 視窗吐出 `GetFolderCount ⓪ RDO 成功`。後續重複收合與展開時，UI 會瞬間反應。
2. **切換多個已載入的資料夾 (例如 A 與 B)**：來回點選 A 與 B 時，Debug 視窗不會再次湧現 `GetMailCount ⓪ RDO 成功` 的訊息，ListView1 能夠瞬間顯示精確的信件和資料夾數字。
