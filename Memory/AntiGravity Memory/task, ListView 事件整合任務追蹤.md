# ListView 事件整合任務追蹤

- [x] 尋找並分析 `Form1.vb` 中的 ListView 初始化邏輯。
- [x] 將 `ListView3` 和 `ListView4` 中簡單的 1~2 行事件處理（例如滑鼠點擊同步路徑、全選等）透過 `AddHandler` 合併。
- [x] 提取並統一在 `Form1_MainTabs.vb` 中的 `GetSelectedEntryIDs`、共通事件邏輯（包含 `Ctrl+A` 等）。
- [x] 清理並移除 `ListView4_MouseDoubleClick` 等靜態的 `Handles` 綁定，改用 `AddHandler`。
- [x] 補齊歷史註解 `by Gemini 3.1 Pro, 2026/04/21`。
