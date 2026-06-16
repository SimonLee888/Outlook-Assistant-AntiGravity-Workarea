# Todo / Debug 標註完整盤點

`Form1.vb` 共 **28** 處 todo、**5** 處 debug  
`DebugForm.vb` 共 **3** 處 todo

---

## ✅ 已完成，可安全移除或標記 Done

| # | 檔案 | 行號 | 內容 | 判斷理由 |
|---|------|------|------|----------|
| 1 | Form1 | L1152 | `_cacheFolderSizeAll.Clear() ' todo: 尚未使用` | 剛才已新增 `GetCachedFolderSizeAllAsync`，已在使用了 |
| 2 | Form1 | L4934 | `todo: 統一成一個函數供各處呼叫` | 已有 `GetMailSize()` (L4517) 統一處理，Fallback 鏈完整 |
| 3 | Form1 | L4573/L4583 | `todo: try/catch裡面包住的 TypeOf 都可以直接拿掉` | **已在 try/catch 裡**，`TypeOf` 不會拋例外，但移掉也無妨。純程式碼風格，不影響功能 |

---

## 🟡 簡單可做，可以立刻動手

| # | 檔案 | 行號 | 內容 | 工作量 | 建議 |
|---|------|------|------|--------|------|
| 4 | Form1 | L2097-2098 | `todo: 改成從 L2.5 cache proxy 讀取` | ⚡ 1 分鐘 | 把 `GetMailCountAll(rf)` 換成 `GetCachedMailCountAllAsync(rf)` 即可 |
| 5 | Form1 | L615/618/622 | `debug: 但是不成功` (SplitContainer 不可選取) | ⚡ 1 分鐘 | 這三段已用 `sc.TabStop=False + Enabled 開關` 成功解決，被 comment 掉的失敗嘗試可以直接刪除 |
| 6 | Form1 | L883 | `debug: 首次切到tab2時 Gmail_2022 不會展開` | ⚡ 1 分鐘 | 看起來是歷史除錯紀錄，下方已有快取邏輯正常運作，可轉為普通註解 |
| 7 | Form1 | L1129 | `debug: 卸載後再重新載入第二次不會成功` | 📝 僅標記 | RDO 的 COM 生命週期問題，這是已知限制不是 bug，可改為「已知限制」註解 |

---

## 🟠 有價值但需要設計決策

| # | 檔案 | 行號 | 內容 | 複雜度 | 我的看法 |
|---|------|------|------|--------|----------|
| 8 | Form1 | L374 | `todo: 真正做到 lazy loading` | 🔧 中 | 目前已是 lazy loading (:::佔位 + BeforeExpand)，此 todo 應已完成。建議確認後標記 Done |
| 9 | Form1 | L376 | `todo: PST 數量多時延遲載入` | 🔧 中 | 與 CacheSniffer 概念相近，可考慮用 BeginInvoke 依序載入 |
| 10 | Form1 | L385-387 | `todo: 背景預讀 foldercount/mailcount/foldersize` | 🔧 中 | CacheSniffer (L392) 已實作但被 comment 掉。如果要啟用，取消註解並測試即可 |
| 11 | Form1 | L403 | `todo: ESC 全域中斷有時管用有時不管用` | 🔧 中 | 問題在於某些 COM call 是阻塞的 (如 `GetTable`)，ESC 只能在 `Await Task.Yield()` 點生效。根本性解法需要改用 CancellationToken |
| 12 | Form1 | L2395 | `todo: 移除 TreeView2 相關程式碼` | 🔧 中 | SimTree2 已穩定，TreeView2 事件邏輯也已被 comment 掉，可以安全移除 |
| 13 | Form1 | L4025 | `todo: GetMailCountAll 改成平行處理跟 GetArray() 的 v4.0` | 🔧 高 | 目前 RDO ⓪ 路徑已極快 (TotalItemCount)，此優化的價值僅在無 RDO 時才體現 |
| 14 | Form1 | L1611-1615 | `todo: debugForm 開啟時 addmessage 拖累速度` 等多項 | 🔧 中 | DebugForm 的 QueueTimer 已改善但未完全解決。可考慮 DebugForm.AddMessage 改為 fire-and-forget |
| 15 | DebugForm | L26-27 | `todo: 點 begin/end 自動 highlight 配對` | 🔧 中 | 可在 `ItemSelectionChanged` 中搜尋配對行，目前的黃色標示框架已存在 |
| 16 | DebugForm | L209 | `todo: 選取變更向前搜尋配對的 Begin: 行` | 🔧 中 | 與上面 #15 是同一個功能需求 |

---

## 🔴 長期規劃 / 暫時不動

| # | 檔案 | 行號 | 內容 | 備註 |
|---|------|------|------|------|
| 17 | Form1 | L172 | `todo: ESC 全域中斷旗標` | 已實作完成，但 todo 文字描述的是設計說明，不是待辦 |
| 18 | Form1 | L346 | `todo: 如何設置版本號自動遞增` | VS 專案層級的設定，與程式碼無關 |
| 19 | Form1 | L357 | `todo: debugform 自動上色, 可多選, 正確減去時間差` | 自動上色已實作 (DrawSubItem)。多選與時間差是獨立功能 |
| 20 | Form1 | L377 | `todo: 第一次 formload 時 RDO 一直還沒 init 完` | RDO init 是非同步的，如果 Load 時 RDO 還沒好自然走 OOM fallback，這是設計行為 |
| 21 | Form1 | L1044 | `todo: SafeGet 拿來替換許多 COM Exception 的地方` | 大規模重構，暫不動 |
| 22 | Form1 | L1526 | `todo: 但我還是想要再嚐試看看 Task.WhenAll` | 研究型 todo，暫不動 |
| 23 | Form1 | L4058/4165 | `todo: 若 ③ 常被觸發需檢查根本原因` | 防禦性註記，保留 |
| 24 | Form1 | L4612 | `todo: 暫時先保留 (GetMailCountRecursive)` | 舊版備用區的標記，保留 |
| 25 | Form1 | L4642 | `todo: 有沒有地方可以用上 BFS 共用清單的好處` | 設計思考，暫不動 |
| 26 | Form1 | L4662/4666 | `todo: cancelRequested / onProgress 如何使用` | 這是 GetMailCountAll_1 的參數說明，不是待辦 |
| 27 | Form1 | L4713 | `todo: 這個進度回報如何使用` | 同上 |
| 28 | Form1 | L4730 | `todo: 遞迴會重複呼叫 GetSubFolderList` | GetMailCountAll_1 的已知限制，保留 |
| 29 | DebugForm | L519 | `todo: 處理項目新增的非同步方法` | OnItemAddedAsync 目前只有 Task.Yield，看起來是未完成的佔位 |

---

## 建議優先順序

1. **立刻做** (#4)：L2097 改用 cache proxy — 一行改動
2. **立刻做** (#1)：L1152 移除「尚未使用」的 todo 註解
3. **清理** (#5)：L615/618/622 刪除已失敗的 debug 嘗試
4. **決策** (#8)：確認 L374 的 lazy loading 是否已完成
5. **決策** (#12)：是否正式移除 TreeView2 相關程式碼
