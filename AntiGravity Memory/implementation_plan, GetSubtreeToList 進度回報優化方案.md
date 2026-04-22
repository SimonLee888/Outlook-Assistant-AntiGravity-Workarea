# GetSubtreeToList 進度回報優化方案

此計畫旨在讓資料夾子樹展開（GetSubtreeToList）的過程透明化，讓使用者在處理大型 PST 檔案時能透過 ProgressBar1/2 看到即時進度。

## User Review Required

> [!NOTE]
> 進度回報會稍微增加 UI 執行緒的負擔（每 100ms 更新一次），但在大型資料夾展開時對 UX 有極大幫助。

## Proposed Changes

---

### UI 與流程層 (Form1_MainTabs.vb)

#### [MODIFY] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_MainTabs.vb)

- **Button3_Click**: 在收集資料夾清單階段加入 `progress` 處理。
- **GoToMonthView**: 確保在展開資料夾時也能顯示進度。

---

### MAPI 操作層 (Form1_Outlook.vb)

#### [MODIFY] [Form1_Outlook.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_Outlook.vb)

- **GetUniqueFolderList**: 修改簽名，新增 `Optional progress As IProgress(Of ProgressReport) = Nothing`。
- **GetSubtreeToList**: 保持現有的 `progress?.Report` 邏輯，確保訊息格式統一。

## Verification Plan

### Automated Tests
- 使用 `Search Web` 或 `Run Command` 確認編譯無誤 (由於是 UI 變動，主要依賴手動驗證)。

### Manual Verification
1. 在 Tab 3 選擇一個包含大量子資料夾的根目錄。
2. 點擊「搜尋」按鈕。
3. 觀察 ProgressBar2 是否顯示「正在展開資料夾結構: 已發現 XXX 個資料夾...」。
4. 驗證 ESC 鍵是否仍能正常中斷。
