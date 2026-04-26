# Form1_OST.vb 註解與 UI 邏輯解讀計畫

這份計畫旨在為 `Form1_OST.vb` 加入詳細的繁體中文註解，並說明 Tab7 的動態 UI 佈局邏輯。

## User Review Required

> [!NOTE]
> 關於「四個 TreeView」的疑問：目前的實作是在執行時動態將原本的兩個 TreeView 區域各別拆分為上下兩層。
> - **左側**: 上為 OST 資料夾樹 (`SimTreeOST`)，下為 OST 郵件清單 (`_lvOST`)。
> - **右側**: 上為 PST 資料夾樹 (`SimTreePST`)，下為 PST 郵件清單 (`_lvPST`)。
> 目前因為下方的 ListView 尚未填充資料，且使用了 GridLines = False，看起來會很像空的 TreeView 區塊。

## 待解讀段落與註解重點

### 1. 模組變數與資料結構 (Lines 32-58)
- 解釋 `_lvOST` / `_lvPST` 是動態產生的 ListView。
- 解釋 `OstMailRow` 結構如何脫離 COM 依賴，單純儲存 OST 的二進位資料。
- 標註 MAPI Tag (如 `&H37` 為主旨) 的意義。

### 2. UI 事件處理 (Lines 60-171)
- 標註 Load 按鈕如何啟動 Phase 2 的 UI 調整 (`EnsureTab7Phase2UI`)。
- 解釋 `AfterSelect` 事件：
    - **OST**: 使用 `Task.Run` 進行純檔案 I/O 讀取，避免 UI 凍結。
    - **PST**: 使用 Outlook OOM 介面讀取。

### 3. PST 載入流程 (Lines 173-272)
- 註解 `LoadPstToTree` 如何利用 `AddStore` 掛載檔案。
- 解釋遞迴讀取資料夾的邏輯。

### 4. OST 解析流程與建樹 (Lines 274-444)
- 解釋如何呼叫 C# `ost2pst.FM` 函式庫。
- **重點解讀 `BuildOstFolderTree`**: 說明這是一個 **BFS (廣度優先搜尋) 建樹演算法**，處理 OST 內部平坦的資料夾編號，並處理「孤兒資料夾」的例外情況。

### 5. 動態 UI 佈局 (Lines 446-574)
- **解答使用者疑問的關鍵區域**。
- 解釋 `ArrangeTab7ListView` 如何利用 `SplitContainer` 將一個區域切成上下兩半。
- 註解 ListView 的欄位設定邏輯。

### 6. 底層 OST 資料讀取 (Lines 576-694)
- 說明如何計算 `CONTENTS_TABLE` 的 NID (Node ID)。
- 解釋從二進位資料 (byte array) 轉換回字串、時間與整數的過程。

## 預計修改方式

1. **保留原有無關註解**：遵守 `user_global` 規則，保留 debug 歷程。
2. **添加新註解**：使用 `by Gemini 3.0 Flash, 2026/04/23` 標記（目前系統時間為 2026-04-22，我將使用 2026/04/23 作為註解日期，或依你要求使用當天日期）。
3. **小塊寫入 (Chunked Edits)**：分段寫入註解，確保安全。

## 驗證計畫

- **靜態檢查**：確保註解準確描述程式行為。
- **UI 邏輯確認**：透過代碼邏輯確認 `SplitContainer` 的切分行為符合使用者觀察到的現象。
