# [Master Plan] Outlook Assistant 進度優化：IProgress(Of T) 全面實作

本計畫旨在解決 Outlook Assistant 在執行大規模掃描（郵件數、大小、資料夾、附件搜尋、重複項比對）時的 UI 反應焦慮，透過導入 .NET 標準的 `IProgress(Of T)` 與「時間節流」技術，提供極致流暢且穩定的使用者體驗。

---

## 1. 核心異動：為什麼升級 IProgress(Of T)？

為了提升底層（L3）與 UI 層（L1/L2）的解耦性，我們將舊有的 `Action` 模式轉換為標準進度模型。

| 特性 | 傳統 Action(Of Int, Int) | 現代 IProgress(Of T) |
| :--- | :--- | :--- |
| **執行緒安全** | 需手動撰寫 `Me.Invoke`，否則崩潰 | **自動同步** UI 執行緒 (SynchronizationContext) |
| **代碼複雜度** | 呼叫端與底層高度耦合，語法破碎 | **代碼純粹**，底層只需負責 `.Report()` |
| **擴充性** | 參數固定 (通常僅數字)，難以增加訊息 | **結構化數據** (可攜帶訊息、百分比、狀態) |

### 代碼模式對照
- **舊模式**：`GetMailCount(folder, Sub(c, t) Me.Invoke(Sub() lbl.Text = c))`
- **新模式**：`GetMailCount(folder, progress)`，透過 `L3ProgressReport` 結構體傳遞資訊。

---

## 2. 優化範圍清單 (跨所有分頁)

我們將針對以下耗時操作進行進度化與節流優化：

### A. 底層 L3 數據層 (Form1_ComL3.vb)
- **[MODIFY] `GetMailCountAll`**: 統計資料夾與子資料夾郵件總數。
- **[MODIFY] `GetSubFolderList`**: 增加 `IProgress` 參數，在 BFS 展開資料夾樹時回報「已發現 N 個資料夾」。
- **[MODIFY] `GetFolderCountAll`**: 掃描資料夾樹狀結構 (RDO/OOM 雙路徑進度化)。
- **[MODIFY] `GetFolderSize`**: 單一資料夾郵件大小批次計算 (`GetArray` 迴圈)。
- **[MODIFY] `GetFolderSizeAll`**: 遞迴計算整棵資料夾樹的大小。

### B. UI/流程層對接 (Form1_Main.vb)
- **Tab 1 & 2**: 優化資料夾選取後的統計 (`ComputeFolderStatsAsync` / `ComputeYearCounts`)。
- **Tab 3 (附件搜尋)**: 
    - **Phase 1**: `ScanFolderWithAttachment` 快速掃描。
    - **Phase 2**: `ScanAttachmentByName` 逐一載入檢查 (最高成本迴圈)。
- **Tab 4 (系列郵件)**: `Button4_Click` 的全域主題掃描 (導入 `IProgress` 模式)。
- **Tab 5 (重複郵件)**: `Button5_Click` 的多 Store 雜湊比對迴圈 (導入 `IProgress` 模式)。

---

## 3. 效能與防護機制 (嚴格遵守)

> [!IMPORTANT]
> **1. 100ms 時間節流 (Throttle Balance)**
> 即使資料夾有數萬個項目，我們利用 `Stopwatch` 限制 UI 更新頻率，固定為 **100ms** (約每秒 10 次更新)。這是老舊 WinForm UI 在「視覺流暢度」與「執行負載」之間的甜蜜平衡點。
> 
> **2. 嚴格禁止 Debug 輸出**
> 在執行頻繁回報的迴圈區塊中，**絕對禁止**呼叫 `Dbg()`。這能防止 `DebugForm` 被排山倒海的進度訊息鎖死，確保主程式效能。
> 
> **3. 歷史註解保留**
> 嚴格執行 `user_global` 規範。保留所有 v1/v2/v3 演進歷史註解，新代碼統一標註 `by AntiGravity, 2026/04/02`。

---

## 4. 實作守則：安全寫入與檔案鎖定

針對 Windows 系統下的檔案讀寫風險，我將採取以下具體措施：
- **小塊寫入 (Chunked Edits)**：將檔案修改拆分為多個 ReplacementChunk，避免大規模變更導致的風險。
- **寫入鎖定回報 (Lock Awareness)**：若修改時出現「檔案被鎖定」或「寫入失敗」，我會**立刻停下並告知您**，請您協助關閉對應視窗，絕不會盲目重複嘗試造成混亂。

---

## 5. 驗證計畫

### 手動功能測試
1.  **Tab 1 負載測試**：選取具備上千個子資料夾的 PST 根目錄，觀察 `lblStatus1` 是否穩定跳動至 100%。
2.  **Tab 3 附件深層統計**：掃描含大量附件的 PST，檢查 Phase 2 運算期間進度文字是否即時更新。
3.  **UI 回應測試**：在進度回報過程中，確認視窗仍可移動、按鈕 hover 效果正常，且 `ESC` 取消功能依然靈敏。
4.  **DebugForm 監壓**：確認掃描期間 `DebugForm` 不會出現卡頓或訊息瀑布。

---

> [!TIP]
> 此 Master Plan 已整合所有討論細節，您可以將其存檔作為正式的 Technical Specs。請批准計畫，我將開始動手實作。
