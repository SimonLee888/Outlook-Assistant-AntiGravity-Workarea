# ListView3 與 ListView4 互動邏輯整合實作計畫

此計畫旨在解決 ListView3（Tab3 附件搜尋）與 ListView4（Tab4 系列郵件搜尋）中散落且重複的事件處理邏輯，將其完全收攏至「共通搜尋結果連動邏輯」區塊。

## User Review Required

> [!NOTE]
> 此次重構將會大量修改 `Form1_MainTabs.vb` 內的這幾個事件處理：
> - `ListView4_MouseDoubleClick`
> - `ListView4_KeyPress`
> - `ListView4_KeyDown`
> - `ListView4_SelectedIndexChanged`
> - `ListView4_MouseClick`
> 
> 修改後將完全改用通用方法處理，並為這些調整補上明確的註解 `by Gemini 3.1 Pro, 2026/04/21`。
> 請您檢閱下列變更計畫。

## Proposed Changes

### Form1_MainTabs.vb

我們將修改並統一 `Form1_MainTabs.vb` 中的程式碼：

#### 1. 強化通用邏輯區塊：`#Region "■ 共通搜尋結果連動邏輯"`
- 擴充或修改現有的 `CommonSearchResult_KeyPress`、`CommonSearchResult_DoubleClick`，確保它們能完美兼容 `ListView3` 與 `ListView4` 的行為。
- 提取「路徑同步至 StatusBar (ProgressBar2)」的通用邏輯，命名為 `CommonSearchResult_UpdateFolderPath`。
- 新增統一的剪貼簿複製邏輯。

#### 2. 修改 ListView4 的事件處理器
#### [MODIFY] Form1_MainTabs.vb
1. **`ListView4_MouseDoubleClick`**: 移除內部手動尋找 ID 的邏輯，直接呼叫 `CommonSearchResult_DoubleClick(sender, e)`。
2. **`ListView4_KeyPress`**: 移除超過 10 封的確認邏輯（因已實作在底層 `OpenMailByEntryID` 內）與抓取 ID 邏輯，直接呼叫 `CommonSearchResult_KeyPress(sender, e)`。
3. **`ListView4_SelectedIndexChanged`** 與 **`ListView4_MouseClick`**: 移除冗餘的 `ProgressBar2.Text` 更新邏輯，改為呼叫 `CommonSearchResult_UpdateFolderPath(sender)`，並保留滑鼠點擊專屬的複製行為。
4. **`ListView4_KeyDown`**: 統一化，若有需要與 `ListView3` 共享 ESC 邏輯，將移至通用方法或精簡之。

#### 3. 確認 ListView3 的綁定
- 檢查 `ListView3.DoubleClick`、`ListView3.KeyPress` 及選取變更事件，若尚未綁定，統一將它們掛載（Handles）到通用的 `CommonSearchResult_` 方法上，使代碼行為絕對一致。

## Open Questions

> [!TIP]
> 請問您希望將 `Ctrl+A` (全選) 以同樣的通用方式擴展給 `ListView3` 使用嗎？目前 `ListView4` 的 KeyDown 有實作 `Ctrl+A`。如果可以，我也會把它放到通用 KeyDown 函式內。

## Verification Plan

### Manual Verification
1. **編譯驗證**: 確保重構後沒有發生編譯錯誤。
2. **功能驗證**:
   - 在 Tab4 選取郵件後按下 Enter 鍵，確認信件是否成功開啟。
   - 雙擊 Tab4 的信件，確認信件是否成功開啟。
   - 點擊 Tab4 的信件，確認視窗底部 `ProgressBar2` 是否正確顯示信件夾路徑。
   - 檢查 Tab3 是否功能依然正常，不會被改壞。
