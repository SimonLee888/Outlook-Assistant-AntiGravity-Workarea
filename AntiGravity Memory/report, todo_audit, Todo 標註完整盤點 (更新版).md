# Todo / Debug 標註完整盤點 (更新版)

`Form1.vb` 目前剩餘約 **15** 處有效待辦。

---

## ✅ 已完成項目

| # | 檔案 | 原行號 | 內容 / 描述 | 完成方式 |
|---|------|------|----------|----------|
| 1 | Form1 | L1152 | `_cacheFolderSizeAll.Clear() ' todo: 尚未使用` | 已啟用快取清空邏輯。 |
| 2 | Form1 | L4934 | `todo: 統一成一個函數供各處呼叫` | 已由 `GetMailSize()` 統一處理。 |
| 3 | Form1 | L4573 | `todo: try/catch裡面包住的 TypeOf 都可以直接拿掉` | 已移除，直接使用 `CLng()`。 |
| 4 | Form1 | L2092 | `todo: 改成從L2.5 cache proxy 讀取` | 已換成 `GetCachedMailCountAllAsync`。 (USER 手動完成) |
| 5 | Form1 | L615+ | `debug: 但是不成功` (SplitContainer 嘗試) | 已刪除失敗代碼並改為 Enabled 邏輯。 (USER 手動完成) |
| 6 | Form1 | L883 | `debug: Gmail_2022 不會展開` | 已確認快取正常後刪除。 (USER 手動完成) |
| 7 | Form1 | L1129 | `debug: 卸載後再重新載入第二次不會成功` | 已標記為「已知限制」。 (USER 手動完成) |
| 8 | Form1 | L344 | `Simon Lee Studio (build ...)` 格式化與修正 | 已改為內插字串並修正「尸」字。 (USER 手動完成) |
| 9 | DebugForm | L26-27 | `todo: 點 begin/end 自動 highlight 配對` | 已由 AntiGravity 完成 O(1) 效能版。 |
| 10 | DebugForm | L209 | `todo: 選取變更向前搜尋配對的 Begin: 行` | 已由 AntiGravity 完成雙向搜尋。 |
| 11 | Form1 | L1611 | `todo: debugForm 開啟時 addmessage 拖累速度` | 已透過 _lastHighlightedPair 解決多選停頓問題。 |
| 12 | Form1 | L357 | `todo: debugform 自動上色, 可多選, 正確減去時間差` | 配對高亮、計算時間差功能已完整實作。 |

---

## 🟡 簡單、中優先級 (建議接著處理)

| # | 檔案 | 原行號 | 內容 | 建議 |
|---|------|------|------|------|
| 13 | Form1 | L887 | `todo: high: 目前最耗時的函數, 首要優化目標 (非郵件目錄占最多時間?)` | **🔥 下一個目標！** 針對非郵件資料夾 (例如行事曆、聯絡人) 的效能進行專屬優化。 |
| 14 | Form1 | L2395 | `todo: 移除 TreeView2 相關程式碼` | 清理舊 TreeView1/2 的殘留邏輯。 |
| 15 | Form1 | L374 | `todo: 真正做到 lazy loading` | 確認目前 BeforeExpand 是否已達標，若是則可 Done。 |

---

## 🟠 需要設計決策 (或中長線規劃)

| # | 檔案 | 原行號 | 內容 | 複雜度 | 我的看法 |
|---|------|------|------|--------|----------|
| 16 | Form1 | L403 | `todo: ESC 全域中斷有時管用有時不管用` | 🔧 中 | 歸因於 COM 阻塞呼叫，需評估是否導入 `CancellationToken` (工程較大)。 |
| 17 | Form1 | L4025 | `todo: GetMailCountAll 改成平行處理跟 GetArray() 的 v4.0` | 🔧 高 | 目前 RDO 路徑已達極速，考慮開發資源分配，此項可延後。 |
| 18 | DebugForm | L519 | `todo: 處理項目新增的非同步方法` | 🔧 中 | 當前 queue 處理已優化，此項屬架構美化。 |

---

## 🔴 其他已知現況 / 註記

- **L346 (版本號遞增)**: 需專案屬性設定，非代碼改動可解。
- **L4058/4165 (備用方案觸發檢查)**: 屬防禦性監控，持續保留。
- **L4662+ (參數說明)**: 屬於文件，不需移除。
- **L1129 (RDO 二次載入限制)**: 已確認為 COM 限制。
