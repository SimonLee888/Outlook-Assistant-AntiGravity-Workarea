# SimTree4 水平捲軸隱藏與位移防止 實作紀錄

我們已成功實作了針對 `SimTree4` 的水平捲軸隱藏功能，並解決了點選長項目時控制項會自動向右位移的問題。

## 變更內容

### [SimTree 控制項](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_SimTree.vb)
- **新增屬性**：`HideHorizontalScrollBar` (預設為 `False`)。
- **Win32 樣式控制**：
    - 覆寫 `CreateParams`，當屬性為 `True` 時加入 `TVS_NOHSCROLL (&H8000)`。這能直接停用系統層級的水平捲軸。
- **訊息攔截**：
    - 覆寫 `WndProc`，當屬性為 `True` 時攔截 `WM_HSCROLL (&H114)` 訊息。
    - > [!TIP]
      > 即使隱藏了捲軸，TreeView 的 `EnsureVisible` 行為仍可能觸發水平捲動。攔截此訊息可確保視圖永遠固定在最左側，不會因為項目標題過長而位移。

### [Form1 初始化](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1.vb)
- 在 `InitTab4UI` 方法中，將 `SimTree4.HideHorizontalScrollBar` 設為 `True`。
- 其他 `SimTree` 實例（如 `SimTree1`, `SimTree2`, `SimTree3`）均保持預設值 `False`，功能不受影響。

## 驗證結果
- **SimTree4**：下方不再出現水平捲軸，且點選長標題郵件系列時，控制項不再產生水平偏移。
- **其他 SimTree**：水平捲軸功能維持正常。
