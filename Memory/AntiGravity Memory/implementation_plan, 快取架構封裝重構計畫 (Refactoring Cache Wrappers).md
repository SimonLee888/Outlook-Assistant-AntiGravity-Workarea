# 快取架構封裝重構計畫 (Refactoring Cache Wrappers)

回應您的意見：「既然在 Tab1 加了這個薄層來輔助 L2 跟 L3 之間的順利轉換，是否其他的 cache 也應比照作法，讓整個程式的邏輯一致？」

**這個提議非常優秀且切中軟體工程的架構精神！** 
目前我們的程式碼中，快取邏輯在某些地方是被「薄層 Wrapper」包裝 (`GetCachedFolderCount`)，但在另一些地方（如 `ComputeFolderStatsAsync`, `GetTotalFolderCountAsync`, `GetMailCountByMAPINew`, `GetFolderSizeAsync`）卻是直接把 `_cache.TryGetValue` 與 `_cache.TryAdd` 混雜在業務邏輯或底層資料探索的迴圈之中。這會造成程式碼職責不清、不好維護。

## 重構目標 (Refactoring Goals)

我們將實行一個 **「快取代理層 (Cache Proxy Layer)」** 模式，將所有的字典操作封裝在專屬的 `GetCachedXxx` 函數中。所有的上層 L1 或是 L2 業務流程，一律只能透過這些代理函數來取得資料；而代理函數底下再透過單純的資料收集函數 (L3) 取得真正原本的值。

### 重構步驟 (Execution Plan)

1. **建立統一的 `GetCachedXxx` 輔助函數群**：
   - [x] (已完成) `GetCachedFolderCount`
   - [ ] 提取 `GetCachedMailCount`，對應封裝 `_directMailCountCache` 與 `GetMailCount`。
   - [ ] 提取 `GetCachedFolderSizeAsync`，對應封裝 `_folderSizeCache` 與重新命名的 `CalculateFolderSizeAsync`。
   - [ ] 提取 `GetCachedTotalFolderCountAsync`，對應封裝 `_folderCountCache` 與負責遞迴計算的 `CalculateTotalFolderCountAsync`。
   - [ ] 提取 `GetCachedTotalMailCountAsync`，對應封裝 `_mailCountCache` 與負責遞迴計算的主邏輯。

2. **淨化業務與 L3 函數 (Purifying Functions)**：
   - **(依據您的建議)** 既然 `GetTotalFolderCountAsync`, `GetMailCountByMAPINew`, `GetFolderSizeAsync` 都已經是準備淘汰的舊函數，我就**完全不再去動它們**了，保持現狀讓您日後自行刪除。
   - 簡化 `ComputeFolderStatsAsync` 的複雜度：不再手動管理 `_directMailCountCache.TryAdd`，直接呼叫 `GetCachedMailCount` 即可。

3. **統一外部呼叫點**：
   - 舉例來說，在全域的 `CacheSniffer`（背景預讀）迴圈中，不再自己寫 `If Not _mailCountCache.ContainsKey... Then Await`，而是乾淨俐落地直接呼叫 `Await GetCachedTotalMailCountAsync(folder)`。

## 預期效益 (Expected Benefits)
- **架構優雅**：L1/L2 只要資料，不必管來源；L2 Wrapper 管理快取；底層單純讀取 COM。
- **程式碼減重**：移除各處散落、重複的 Dictionary 讀寫與 `FolderPath` 對應。
## 針對您的疑問與深入分析 (Addressing Your Questions)

> **Redis 是什麼？比快取更快嗎？**
> Redis 是業界最有名的一種「獨立記憶體資料庫 (In-Memory Database)」，常被大型網站與後端伺服器用來當作跨機器的分散式快取。
> 它的速度超級快，這只是我在跟您舉例軟體工程中「抽換底層引擎」的觀念而已。對於我們這套跑在個人 Windows 上的 Outlook 桌面小程式來說，現在用的 `ConcurrentDictionary` 就是最頂級、最零延遲的極致效能了，完全不需要（也不適合）去裝大砲等級的 Redis！

> **這樣切分架構，會不會有不同層級之間重複動作的疑慮？效能是否折損？**
> 完全不會折損，甚至會更好。
> `GetCachedXxx` 這層薄層在電腦底層的執行代價只有幾個「奈秒 (nanosecond)」，就是進去查一下 Dictionary 而已。因為它把判斷收攏在一個地方，反而絕對杜絕了「不小心在兩個不同地方又打了一次 COM」的重複動作發生。

> **對於平行處理有幫助還是有害？在流程迴圈中要 ESC 中斷或是用 onProgress 來更新狀態有益還是有害？**
> 1. **平行處理 (Parallel)**：有益且安全。我們用的是 `ConcurrentDictionary`，本身就是專門給多執行緒 (Multi-threading) 搶佔用的。如果在平行的 `Task.WhenAll` 中有兩根執行緒同時摸到同一個資料夾，代理層會正確處理。
> 2. **ESC 中斷與 onProgress 回報**：完全不影響，因為代理層只是「路過」。如果快取沒中，代理層就把您的 `onProgress` 和 `_cancelRequested` (或 `CancellationToken`) 原封不動地往地下一層 (L3) 遞過去，該取消就取消，該更新狀態列就更新。

## 結論
既然釐清了那些待淘汰函數不需要動，這個計畫的修改範圍變得更小且更安全了。我們只需針對您日常真正在跑的核心函數進行 Wrapper 包裝即可。
