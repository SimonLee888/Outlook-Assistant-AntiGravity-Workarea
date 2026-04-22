# 關閉 SimTree4 水平捲軸與位移防止實作計畫

此計畫旨在解決 `SimTree4` (自訂 `SimTree` 控制項) 在內容過長時出現水平捲軸的問題，並防止在點選長項目時，控制項自動向右移動 (Auto-scroll) 導致佈局跑掉。

## 使用者評論請求
> [!IMPORTANT]
> 為了滿足「只有 SimTree4 不顯示」的需求，我將在 `SimTree` 類別中新增一個 `HideHorizontalScrollBar` 屬性（預設為 `False`）。我們只需在 `SimTree4` 的初始化過程中將其設為 `True` 即可。

## 擬議變更

### [Component] SimTree 控制項 (Form1_SimTree.vb)

#### [MODIFY] [Form1_SimTree.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_SimTree.vb)

1.  **引入 Win32 常數與宣告**：
    *   定義 `TVS_NOHSCROLL = &H8000`
    *   定義 `WM_HSCROLL = &H114`
2.  **新增公共屬性**：
    *   `HideHorizontalScrollBar As Boolean` (預設 `False`)。
3.  **覆寫 `CreateParams`**：
    *   若 `HideHorizontalScrollBar` 為 `True`，則在 `MyBase.CreateParams.Style` 中 OR 運算加入 `TVS_NOHSCROLL`。
4.  **覆寫 `WndProc`**：
    *   若 `HideHorizontalScrollBar` 為 `True` 且收到 `WM_HSCROLL`，則直接攔截不處理。

### [Component] Form1 初始化 (Form1_MainTabs.vb)

#### [MODIFY] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_MainTabs.vb)

1.  **設定屬性**：
    *   在 `InitializeTab4` 或相關位置，設定 `SimTree4.HideHorizontalScrollBar = True`。


## 驗證計畫

### 手動驗證
1.  在 `SimTree4` 中加入一個內容極長的項目（例如一串很長的郵件標題）。
2.  確認控制項下方**沒有**出現水平捲軸。
3.  點選該長項目，確認視窗或控制項**不會**自動向右滑動。
4.  測試垂直捲動是否仍然正常運作（垂直捲軸應保留）。
