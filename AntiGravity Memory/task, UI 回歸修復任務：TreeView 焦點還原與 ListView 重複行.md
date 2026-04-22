# UI 回歸修復任務：TreeView 焦點還原與 ListView 重複行

## 準備工作
- [x] 研究 SimTree 控制項內部選取機制
- [x] 研究 ExpandTreeToDefaultInbox 選取邏輯
- [x] 獲得使用者對實作計畫的批准 (2026/04/17)

## 核心修改
- [x] **[SimTree] 強化對齊與安全性檢查**
    - [x] 修改 `SelectedNodes` 屬性，過濾 `node.TreeView IsNot Me` 的節點 (by Gemini 3.0 Flash, 2026/04/17)
    - [x] 確保 `ClearSelectedNodes` 完整重置 `_lastClickedNode`
- [x] **[Form1] 建立路徑導航輔助函數**
    - [x] 實作 `GetSelectedFolderPath(tv)`
    - [x] 實作 `SelectNodeByPath(tv, path)`
- [x] **[Form1] 重構與實行 CheckedChanged 事件**
    - [x] 將 `checkShowAllFolders` 從 `AddHandler` 移出，建立獨立的 `Sub CheckShowAllFolders_CheckedChanged`
    - [x] 在事件中實作：記住路徑 -> 清除狀態 -> 重載樹 -> 還原路徑

## 驗證
- [x] 測試切換過濾後，焦點是否停留在原處
- [x] 確認 ListView1 重複顯示「收件匣」的問題已解決
- [x] 確認若路徑被過濾掉時，能正確 fallback 回收件匣
