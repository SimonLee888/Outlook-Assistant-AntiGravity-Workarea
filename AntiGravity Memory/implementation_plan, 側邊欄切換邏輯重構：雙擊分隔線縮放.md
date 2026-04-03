# 側邊欄切換邏輯重構：雙擊分隔線縮放

將原本的「側邊欄切換按鈕」移除，改為**雙擊 SplitContainer 分隔線**即可自動收合/展開左側面板。
收合時會保留 10 像素的觸控區，確保使用者可以輕易連按兩下恢復原狀。

## 使用者回饋

- **保留觸控區**：收合後寬度設為 10 像素，而不是完全歸零。
- **直覺設計**：連按兩下分隔線即可切換，無需額外按鈕。

## 擬定的變更

### [Form1.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1.vb)

#### [MODIFY] `InitSplitContainer`
- 將所有 `SplitContainer` 的 `Panel1MinSize` 設為 `0`。
- 移除舊的 `AddSidebarToggle` 調用邏輯。
- 為所有 `SplitContainer` 掛載 `MouseDown` 事件來偵測雙擊。

#### [NEW] `HandleSplitContainerMouseDownShared`
- 偵測左鍵連按二下 (`e.Clicks = 2`)。
- 判斷目前寬度：
  - 若大於臨界值（如 20px），紀錄當前值（`Tag`）並縮小至 10px。
  - 若小於等於臨界值，讀取 `Tag` 紀錄並恢復（預設恢復至 250px）。

#### [DELETE] `AddSidebarToggle` & `HandleToggleSidebarShared`
- 完整刪除這兩個副程式，包含其 UI 產生邏輯。

#### [CLEANUP] `InitListViews`, `InitTab3UI`, `InitTab4UI`, `InitTab5UI`
- 移除原本為了放置切換按鈕而調整的 Panel 設定（如果不再需要）。

## 驗證計畫

### 手動驗證
1.  啟動程式，在各分頁的 `SplitContainer` 分隔線上連按兩下。
2.  確認左側 TreeView 縮小至寬度 10 像素。
3.  再次於 10 像素處連按兩下，確認還原為之前的寬度。
4.  確認原本的「顯示/隱藏側邊欄」按鈕已經消失，UI 佈局保持整潔。
