# Tab3 & Tab4 搜尋結果邏輯整合計畫 (修正版)

## 核心目標
將 Tab3 (附件搜尋) 與 Tab4 (系列郵件) 兩者高度重複的「搜尋結果操作」邏輯（開啟郵件、同步路徑、ESC 處理）進行整合，但不影響 Tab1/Tab2 的特殊業務邏輯。

## 修改範圍

### 1. 職責分離
- **Tab3 & Tab4**: 行為是「搜尋結果清單」，主要操作是批次開啟、ESC 回退。
- **Tab1 & Tab2**: 行為是「導覽與分析」，涉及圖表連動，維持原本獨立的事件處理。

### 2. 實作細節

#### [NEW] 在 `Form1_MainTabs.vb` 建立共通處理器
不再使用全域 `InitListView` 掛載，改為在 `Form1_Load` 或初始化處「手動」對 3 和 4 加入。

```vb
' 只針對 3 和 4 的共通分發邏輯
Private Sub HandleSearchResultClick(sender As Object, e As EventArgs)
    ' 僅處理路徑同步至 ProgressBar2
End Sub

Private Sub HandleSearchResultKeyPress(sender As Object, e As KeyPressEventArgs)
    ' 僅處理 Enter (開啟) 與 ESC (回退)
End Sub
```

#### [MODIFY] 移除冗餘
刪除目前在 `ListView3` 與 `ListView4` 中重複寫入的 `OpenMailByEntryID` 呼叫邏輯，統一委派。

## 開放問題
- **ProgressBar2 顯示**：目前 Tab1/Tab2 也會使用 ProgressBar2 顯示狀態嗎？如果是，則同步路徑的邏輯需要更小心地判斷當前 Tab。

## 驗證計畫
1. 測試 Tab1/Tab2：確認點選、雙擊時原本的報表或圖表連動功能完好無損。
2. 測試 Tab3/Tab4：確認 Enter 開啟多封郵件的功能在兩處皆能正常運作。
