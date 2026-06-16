# Tab4 改造手術清單

- `[x]` 基礎建設與 UI 結構調整
    - [x] 定義模式標記變數 `_isTab4ShowingResults`
    - [x] 修改 `InitTab4UI`：拆除三欄結構，恢復二欄佈局
    - [x] 移除 `TreeView4` 相關代碼
- `[x]` 事件邏輯重導向
    - [x] 修改 `Button4_Click`：搜尋結果改填入 `SimTree4`
    - [x] 修改 `SimTree4_AfterSelect`：支援「資料夾/結果」雙模式切換
- `[x]` 操作體驗優化
    - [x] 實作 ESC 鍵恢復資料夾樹功能
    - [x] 修正 Tab 切換時的自動載入排除邏輯
- `[x]` 最終複檢與驗證
    - [x] 檢查分割線 (Splitter) 同步邏輯是否已安全移除
    - [x] 驗證註解標記與代碼整潔度
