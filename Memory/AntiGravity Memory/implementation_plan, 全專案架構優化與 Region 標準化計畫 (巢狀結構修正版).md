# 全專案架構優化與 Region 標準化計畫 (巢狀結構修正版)

非常抱歉，我先前完全忽略了您精心設計的巢狀 Region 結構（`■` 大標題搭配 `├` 和 `└` 子標題）。這種結構非常清晰，我們絕對應該保留並延續這種風格！

既然 `ComputeFolderStatsAsync` 與 Tab1 緊密相連，我們就讓它留在 `Form1_MainTabs.vb`，不再強行搬移。

針對您覺得「很難分類」的 `#Region "■ 03 共用控制項行為"`，我現在理解是因為它底下的子分類（`├ 共用 UI控制項`、`├ 滑鼠 & 鍵盤操作事件`、`└ 其他輔助事件`）隨著專案發展，已經塞滿了跨越單純 UI 範圍的邏輯（例如全域快取控制、節點導航等）。

以下是我為您重新擬定的架構整理計畫，這次將嚴格遵守您的巢狀標籤結構：

## User Review Required

> [!IMPORTANT]
> 請檢視以下針對 `Form1.vb` 與 `Form1_Outlook.vb` 巢狀 Region 的重新命名與拆解方案。
> 我維持了您的兩層式結構，但對標題與內容物進行了更精確的洗牌。您覺得這個新分類符合您的直覺嗎？

## 擬定變更

---

### [Component] Form1.vb (主視窗與全域共用邏輯)

#### [MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1.vb)

**策略**：將原本龐大的 `■ 03` 區塊，依據實際邏輯屬性拆分為 `■ 03` 與 `■ 04` 兩個大群組，並重構子節點。

*   **保留與微調**：`■ 01 全域宣告` 與 `■ 02 Form 生命週期 & 外觀初始化` 維持不變。但會將非同步工具 (`OkayNowYouHaveToken`, `ThrottledYieldAsync`) 從生命週期區移至下方更合適的地方。
*   **重組與新建**：

```vb
#Region "■ 03 全域狀態與視圖導航"
#Region "  ├ 全域切換與快取管理"
    ' CheckShowAllFolders, CheckRDO 切換事件
    ' SaveCache, LoadCache, ClearCache, RenewCache 按鈕事件
    ' ClearMemoryCachesInner 等快取清理輔助
#End Region
#Region "  ├ 樹狀視圖導航與節點搜尋"
    ' ExpandTvToDefaultInbox, GetActiveTreeView, GetAllTreeViews, TriggerTvAfterSelect
    ' FindNodeByFolderPath, GetSelectedFolderPath, SelectNodeByPath, SelectNodeByPathRecursive
#End Region
#Region "  └ 分頁切換行為"
    ' TabControl1_SelectedIndexChanged, RefreshAllTreeViews
#End Region
#End Region

#Region "■ 04 共用控制項事件與輔助工具"
#Region "  ├ 控制項樣式與滑鼠事件"
    ' HandleSplitterMouseDown, HandleTvMouseHover, HandleLvMouseHover, HandleLvGotFocus, HandleLvResize
    ' AutoResizeLvColumns, GetHeaderRowBackColor
#End Region
#Region "  ├ 鍵盤與內容操作"
    ' HandleTvKeyPress, FindLvItemByName, HistoryListBox 相關事件
#End Region
#Region "  └ 底層非同步與系統輔助"
    ' OkayNowYouHaveToken, ThrottledYieldAsync (從 02 移來)
    ' SyncDebugFormPosition
#End Region
#End Region
```
**理由**：把「資料夾的切換與尋找」以及「全域快取的開關」從「UI 控制項樣式」中抽離出來，賦予獨立的大標題 `■ 03`，剩下的純 UI 行為和工具函數歸入 `■ 04`，這樣邏輯就清晰了。

---

### [Component] Form1_Outlook.vb (資料與快取層)

#### [MODIFY] [Form1_Outlook.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_Outlook.vb)

您的手動調整已經將 `Layer 2.5` 等概念帶入子節點 (`├`)，但外層的 `■ 10` 大標題依然是籠統的「底層 COM 函數群」。

**策略**：正式將大標題依據分層架構拆解，並維持您的巢狀風格。

```vb
#Region "■ 10 初始化與流程協調 (Layer 2)"
#Region "  ├ Outlook 與 COM 初始化"
    ' InitOutlookNamespace, InitRdoSession 等
#End Region
#Region "  └ 流程協調層"
    ' GetUniqueFolderList 等涉及多個呼叫的流程函數
#End Region
#End Region

#Region "■ 20 快取代理層 (Layer 2.5)"
#Region "  ├ 資料夾統計快取"
    ' GetMailCount, GetFolderCount, GetFolderSizeAsync 等
#End Region
#Region "  ├ 陣列與子樹快取"
    ' GetAttachMailList, GetSubtreeToList 等
#End Region
#Region "  └ 快取同步輔助"
    ' FillFolderCacheFromDbRow (從雜項移入這裡，因為它是核心快取機制)
#End Region
#End Region

#Region "■ 30 底層數據層 (Layer 3)"
#Region "  ├ 單一屬性讀取 (COM)"
    ' GetMailCountL3, GetLiveFolderSnapL3 等
#End Region
#Region "  └ 陣列與列表讀取 (COM/Table)"
    ' GetAttachMailListL3, GetSubtreeToListL3 等
#End Region
#End Region
```

---

## 驗證計畫
1.  **結構複檢**：完成修改後，使用 `view_file` 工具重新展開這幾個檔案，確保所有的 `■` 和 `├`/`└` 對齊完美，沒有落單的函數。
2.  **語法確認**：使用 `multi_replace_file_content` 進行小塊搬移，每次修改後確認無語法中斷。
