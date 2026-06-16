# ProgressBar 狀態歷史紀錄功能實作總結 (2026/04/02)

本階段成功為 Outlook Assistant 狀態列增加了「點擊查看歷史訊息」的 Popup 介面。這項微幅更新讓使用者可以隨時回溯過去的執行進度與耗時，不用擔心錯過稍縱即逝的系統訊息。

## 實作內容亮點

### 1. 無痛攔截的歷史紀錄 (Smart Overwrite)
*   **全域攔截**: 透過直接監聽 `ProgressBar1.TextChanged` 與 `ProgressBar2.TextChanged`，無需修改現有的任何指派 (Assignment) 程式碼，成功攔截所有狀態文字。
*   **PB2 智慧覆寫**: 為了解決 `ProgressBar2` (10/100, 20/100...) 快速洗版的問題，設計了前綴比對機制：若新進來的文字跟歷史第一筆的開頭前 10 字元相同 (如皆為「正在統計郵件數:」)，便直接覆寫取代，而不是產生新的歷史紀錄。這確保了只記錄該階段的**最後終極進度**與耗時總結。

### 2. 優化 UX 的動態 ListBox Popup 
*   **自訂下拉選單**: 利用 `ToolStripDropDown` 包裝了一個動態大小的 `ListBox`，而不是死板的 `ContextMenuStrip`。 
*   **完美體驗**: 
    - 自動計算最適合的寬度。
    - 最多顯示 15 筆高度，超過自動展現捲動條。
    - **支援滑鼠滾輪**、最新的時間在最上方，操作極度順暢自然。
*   **精準定位**: 自動擷取滑鼠點擊 `ProgressBar` 的 X 座標，將選單漂亮地繪製在狀態列的「正上方」。

### 3. 一鍵複製
*   掛載了 `_historyListBox_SelectedIndexChanged`，只要滑鼠單擊清單內的任何一筆紀錄（例如 `[21:12:35] 統計花費 3.2 毫秒`），就會自動複製到剪貼簿，並優雅地自動關閉該清單。

## 代碼調整範圍
*   `Form1.vb` 
    *   新增 `StatusHistoryItem` 資料結構。
    *   新增 `_statusHistory` 變數 (上限 100 筆)。
    *   新增 `AppendStatusHistory` 等輔助邏輯。
    *   掛載 Popup 的視窗產生相關邏輯。

> [!TIP]
> 此功能現在通用於 `ProgressBar1` 與 `ProgressBar2`。您可以用滑鼠點擊左側或右側的狀態列，都能召喚出這份精心優化的歷史選單！

---
by AntiGravity, 2026/04/02
