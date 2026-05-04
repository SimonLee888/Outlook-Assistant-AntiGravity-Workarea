# 重構 Bt5_Click (Tab5 重複郵件掃描)

`Bt5_Click` 目前長達約 160 行，雖然邏輯清晰（分為 Step 1, 2, 3），但全部擠在同一個 Event Handler 裡確實讓程式碼顯得擁擠。將其重構切分為較小的獨立函數可以大幅提升可讀性，未來如果要在背景執行掃描（例如建立排程），也會更容易抽出共用。

## Proposed Changes

### [Outlook Assistant]

我們將把 `Bt5_Click` 的主要階段抽離到 `#Region "  ├ Layer2 流程協調層"` 下的新函數中：

#### [MODIFY] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_MainTabs.vb)

1. **`ScanFoldersForDuplicatesAsync`** (處理 Step 2)：
   - **輸入**：`folderList`, `isExact`, `cToken`, `swThrottle`, `progress5`
   - **輸出**：`Task(Of Dictionary(Of String, List(Of MailItemInfo)))` (即 `exactDict`)
   - **功能**：負責使用 `GetTable` 掃描各個資料夾，解析並建立基礎的分組字典 (`exactDict`)。

2. **`BuildDuplicateListViewAsync`** (處理 Step 3)：
   - **輸入**：`exactDict`, `isExact`, `cToken`, `progress5`
   - **輸出**：`Task(Of Integer)` (回傳找到的群組數量，供 UI 更新提示)
   - **功能**：負責 Jaccard 二次過濾，並將結果寫入 `ListView5`。

3. **`Bt5_Click`** (重構為純粹的 UI 協調者)：
   - 只負責：UI 鎖定準備、呼叫 `GetUniqueFolderList`、呼叫 `ScanFoldersForDuplicatesAsync`、呼叫 `BuildDuplicateListViewAsync`、以及最後的 UI 解鎖與狀態報告。

> [!TIP]
> 這樣重構後，`Bt5_Click` 預計會縮減至 40 行以內，流程會變成非常直觀的 `Await Step1() -> Await Step2() -> Await Step3()` 結構。

## User Review Required
請問這樣的切分方式（將資料收集與 UI 渲染分開）是否符合您的重構期望？若是，我將開始執行重構。
