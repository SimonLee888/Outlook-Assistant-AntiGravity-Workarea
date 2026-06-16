# 三階段初始化架構實作計劃 (9.1)

本計劃旨在重構 `Form1` 的 UI 初始化邏輯，解決 `ListView2` 佈局衝突並統整所有 `TreeView` 與 `ListView` 的外觀設定。

## 核心設計理念：職責切分 (Separation of Concerns)

將原本混雜在各個 `InitTabXUI` 中的邏輯拆分為三個清晰的階段：

1.  **Phase 1: Mounting (掛載)** - 建立動態元件並決定它們「屬於誰」(Controls.Add / Parent)。
2.  **Phase 2: Theming (渲染)** - 透過遞迴搜尋 (Recursive Search)，對全畫面已掛載的控制項統一套用樣式與事件。
3.  **Phase 3: Final Layout (佈局)** - 最後才設定 `Dock`、`Size` 並進行 `SendToBack/BringToFront` 的 Z-Order 排序。由於排序在最後執行，絕對不會被屬性設定干擾。

---

## 預計修改內容

### [MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual Studio/Visual Studio 18 (2026)/Outlook Assistant - (AntiGravity測試區)/Form1.vb)

#### 1. 職責拆分：修改各 Tab 初始化函數
將 `InitTab2UI` ~ `InitTab5UI` 的內容重整：
- 保留 `New` 控制項與 `Controls.Add` 在函數內。
- 將 `Dock` 與 `BringToFront/SendToBack` 相關邏輯收納在一起，確保在 Phase 3 執行。

#### 2. 全域渲染：優化 `InitListViews` 與 `InitTreeViews`
- **[NEW]** `GetAllListViews(container)`: 遞迴搜尋容器內所有 ListView。
- **[MODIFY]** `InitListViews()`: 改用遞迴搜尋，不再寫死 `ListView1~5`。
- **[MODIFY]** `InitTreeViews()`: 確保在 Phase 2 執行，此時 `SimTree2` 已掛載進容器，必能抓到。

#### 3. 統籌：重構 `InitLookAndFeel()`
調整調用順序為：
```vb
' 1. Phase 1: Mount (成家)
InitTab2UI()
InitTab3UI()
...

' 2. Phase 2: Theme (立業)
InitListViews()    ' 統一設定字型、雙緩衝、View屬性
InitTreeViews()    ' 統一設定字型、縮排、BeforeExpand事件

' 3. Phase 3: Final Layout (排隊)
ApplyFinalLayout() ' 確保 Tab2 的 ListView2(Top) 與 Chart2(Fill) 順序正確
InitSplitContainer() ' 設定游標
```

---

## 使用者評論區 (User Comments)
> [!NOTE]
> 請在此處留下您的意見或建議。

---

## 驗證計劃

### 手動測試 (Manual Verification)
1.  **Tab 2 佈局檢查**：ListView2 的高度是否維持在 250px？下方的 Chart2 是否正常顯示（沒被遮擋）？
2.  **動態控制項檢查**：切換到 Tab 2 時，`SimTree2` 的資料夾樹是否能正常點擊展開？
3.  **一貫性檢查**：Tab 5 的 `ListView5`（程式建立）是否具有與其他分頁一致的字型與網格設定？

### 自動化檢查
- 檢查 `Dbg` 日誌，確認三階段的執行順序是否符合預期。
