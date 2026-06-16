# 實作搜尋框 (txtDebug) 關鍵字歷史回溯功能

使用者希望在 `txtDebug` 搜尋框按下「上/下」鍵時能回溯之前的搜尋紀錄，且不希望增加額外的 UI 負載。

## 提出變更

### DebugForm (Form1_DebugForm.vb)

#### [MODIFY] [Form1_DebugForm.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual Studio/Visual Studio 18 (2026)/Outlook Assistant - (AntiGravity測試區)/Form1_DebugForm.vb)

- **成員變數區塊**:
    - 新增 `_historyDebug` (List of String)。
    - 新增 `_historyIndex` (Integer)。
    - 新增 `_tempInput` (String) 用於存放在開始回溯前使用者目前的輸入。
- **Shown 事件**:
    - 加入 `AddHandler txtDebug.KeyDown, AddressOf txtDebug_KeyDown`。
- **新增 txtDebug_KeyDown 事件處理程序**:
    - 攔截 `Keys.Enter`: 呼叫 `AddToHistory()` 並模擬搜尋結束。
    - 攔截 `Keys.Up`: 向上導覽歷史紀錄。
    - 攔截 `Keys.Down`: 向下導覽歷史紀錄。
- **新增輔助邏輯**:
    - `AddToHistory(query As String)`: 檢查重複性，避免存入空值。
    - `NavigateHistory(step As Integer)`: 處理索引位移與文字替換。

## 預期行為

1. **紀錄**: 使用者輸入文字並按下 `Enter` 後，該字串會被存入歷史末端。
2. **回溯**:
   - 按下 `Up` 時，若目前正在輸入新內容，會先將新內容存入 `_tempInput`。
   - 繼續按 `Up` 會在歷史紀錄中往舊的項目移動。
   - 按下 `Down` 會往新的紀錄移動，直到回到 `_tempInput` (原本正在打的東西)。
3. **去重**: 連續輸入相同的搜尋詞不會產生多筆重複紀錄。

## 驗證計畫

### 手動驗證
1. 開啟 DebugForm。
2. 在搜尋框輸入 "keyword1" 並按下 `Enter`。
3. 在搜尋框輸入 "keyword2" 並按下 `Enter`。
4. 清空搜尋框，輸入 "pending..." (不按 Enter)。
5. 按下「上鍵」，搜尋框應顯示 "keyword2"。
6. 再次按下「上鍵」，搜尋框應顯示 "keyword1"。
7. 按下「下鍵」，搜尋框應顯示 "keyword2"。
8. 再次按下「下鍵」，搜尋框應顯示 "pending..."。
