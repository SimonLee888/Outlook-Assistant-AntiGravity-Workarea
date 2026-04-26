# Implementation Plan - 全面導入 SafeGetPath 安全機制

## 概述
為了解決 Outlook COM 物件不穩定導致的崩潰問題，本計畫將全面替換專案中直接存取 `.FolderPath` 的程式碼，改為使用已驗證的 `SafeGetPath` 工具函數。

## 修改範圍與細節

### 1. Form1_SQLite2.vb (快取核心)
*   **RenewCacheAsync (Phase 1)**: 在掃描現有資料夾建立比對字典時，改用 `SafeGetPath`。
*   **CleanupOrphanFolderPath**: 清理資料庫孤兒時，確保路徑比對來源安全。

### 2. Form1_Outlook.vb (掃描與判斷)
*   **GetSubtreeToListL3**: BFS 掃描的入口點與循環體內部，全面改用安全路徑取得。
*   **GetSortedSubFolders**: 排序邏輯中的路徑讀取。
*   **HasSubFoldersFast**: 檢查資料夾狀態時的安全性加固。

### 3. Form1_MainTabs.vb (UI 交互)
*   **頁籤切換邏輯**: 當使用者在 TreeView 選取資料夾時，取得路徑的操作應經過 `SafeGetPath` 保護。

## 第三點專門優化：IsMailFolder 與呼叫鏈效能提升

### 優化策略：
目前的程式碼常發生重複讀取路徑的情況。我們將推行以下模式：
1. **外部預讀**: 在進入循環或判斷式前，先透過 `fPath = SafeGetPath(folder)` 取得路徑。
2. **參數傳遞**: 將 `fPath` 傳入 `IsMailFolder(folder, fPath)`。
3. **COM 閉鎖**: `IsMailFolder` 內部偵測到 `fPath` 已存在，將直接跳過 COM 呼叫，僅執行邏輯判斷。

## 驗證計畫
1. **靜態分析**: 全域搜尋 `.FolderPath` 確保無遺漏。
2. **運行驗證**: 執行 RenewCache 功能，確認在大規模 PST 掃描下不會觸發編譯錯誤 BC30451 或執行期例外。
3. **Log 監控**: 透過 `_dbg` 確認路徑解析是否依然正確。

---
**by AntiGravity (Gemini 3.0 Flash), 2026/04/24**
