# Dbg() 品質改善 — 任務清單

## Issue 1：高頻迴圈去噪（Form1_ComL3.vb）
- [ ] `LoadStoreToTreeView()` — 移除迴圈內的 `Dbg("", root.Name)`
- [ ] `LoadSubFolderToTreeView()` — 移除迴圈內的 `Dbg("", selectedFolder.Name & folder.Name)`
- [ ] `GetMailCount()` — 靜默成功路徑，只保留失敗 Dbg
- [ ] `GetFolderCount()` — 靜默成功路徑，只保留失敗 Dbg
- [ ] `IsMailFolder()` — 移除每次呼叫的開始/結束，只在非郵件資料夾時輸出
- [ ] `CalculateSimilarity()` (Form1_Main.vb) — 完全移除開始/結束 Dbg
- [ ] `LevenshteinDistance()` (Form1_Main.vb) — 完全移除開始/結束 Dbg

## Issue 2：補缺失的「結束」
- [ ] `FetchDirectMailCountsAsync()` (Form1_Main.vb) — 補 `Dbg("結束", ...)`
- [ ] `HandleListViewKeyPress()` (Form1.vb) — 重構為只在 Enter/ESC 時輸出，移除無意義的開始
- [ ] `HandleTreeViewKeyPress()` (Form1.vb) — 同上
- [ ] `FindNodeOrItemByName()` (Form1_ComL3.vb) — 補 `Dbg("結束", ...)`

## Issue 3：早期 Return 前補 Dbg
- [ ] `TreeView1_AfterSelect()` 序號不匹配路徑 (Form1_Main.vb)
- [ ] `TreeView1_AfterSelect()` cancelRequested 路徑 (Form1_Main.vb)
- [ ] `SimTree2_AfterSelect()` 無節點選取路徑 (Form1_Main.vb)
- [ ] `ExpandTreeToDefaultInbox()` 找不到收件匣路徑 (Form1.vb)

## Issue 4：格式標準化（msg 不含函數名或串接字串）
- [x] `TreeView1_AfterSelect()` 的 `Dbg("Error", ...)` → `Dbg("錯誤", ...)` (Form1_Main.vb)
- [x] `GetYearCountsForFolder()` 的錯誤格式 (Form1_Main.vb)
- [x] `GetMonthCountsForYear()` 的 `Dbg("GetMonthCountsForYear Error: ", ...)` (Form1_Main.vb)
- [x] `Button3_Click()` 的 `Dbg("Button3_Click Error: ", ...)` (Form1_Main.vb)
- [x] `Button4_Click()` 的 `Dbg("Button4 GetTable Error: " & folder.Name, ...)` (Form1_Main.vb)
- [x] `Button5_Click()` 的 `Dbg("Button5 GetTable Error: " & folder.Name, ...)` (Form1_Main.vb)
- [x] `OpenMailByEntryID()` 的 `Dbg("打開郵件", ...)` (Form1_Main.vb)
- [x] Form1.vb：`HandleSplitContainerMouseDown()` 側邊欄縮合/恢復格式

## Issue 5：GetSizeMultiplier 永遠到達不了的「結束」
- [x] `GetSizeMultiplier()` (Form1_Main.vb) — 移除無效的 Dbg 開始/結束

## Issue 6：補強重要函數的 detail 內容
- [x] `ScanAttachmentDetail()` 開始補入 `targetMailList.Count` (Form1_Main.vb)
- [x] `CheckTab3CacheOrRescan()` — 已有完整開始/結束 ✅
- [x] `ScanFolderWithAttachment()` — 已有「開始 folder.Name」，可接受 ✅

## Issue 7：移除冗餘 Dbg
- [ ] `ShowProgressTab2()` — 移除開始/結束（函數體只有 5 行）(Form1_Main.vb)
- [ ] `ListView4_SelectedIndexChanged()` — 移除開始/結束（函數體為空）(Form1_Main.vb)
- [ ] `Button4_Click()` — `swThrottle.Start()` 位置有誤（應在初始化時啟動）
