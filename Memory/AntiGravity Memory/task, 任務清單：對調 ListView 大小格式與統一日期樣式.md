# 任務清單：對調 ListView 大小格式與統一日期樣式

## 實作步驟

- [x] **修改 `Form1_MainTabs.vb`**
    - [x] 修改 `ListView3_RetrieveVirtualItem` (L1858-1859)
        - [x] 大小改為 `(mail.Size \ 1024L).ToString("N0") & " KB"`
        - [x] 日期改為 `mail.ReceivedTime.ToString("yyyy/MM/dd")`
    - [x] 修改 `FillListView4` (L2608-2609)
        - [x] 大小改為 `mailItem.Size.ToString("N0")`
        - [x] 日期改為 `mailItem.ReceivedTime.ToString("yyyy/MM/dd")`

- [x] **修改 `Form1.vb`**
    - [x] 修改 `InitTab4UI` (L639)
        - [x] 加入 `.Columns("收到日期").TextAlign = HorizontalAlignment.Center`
    - [x] 修改 `AutoResizeListViewColumns` (L1558-1561)
        - [x] 將 `ListView3` 的寄件者比例從 `0.15` 改為 `0.2`

- [x] **驗證內容**
    - [x] 檢查 Tab 3 顯示 (KB, 日期置中)
    - [x] 檢查 Tab 4 顯示 (Bytes, 日期置中)
    - [x] 檢查欄位對齊一致性

- [x] **總結文件**
    - [x] 建立 `walkthrough.md`
