# 為 lvOst 欄位標題新增排序功能實作計畫

本計畫旨在為 Tab7 的 `LvOST` (與 `LvPST`) 提供欄位點選排序功能。將實作一個通用的比較器，優先使用 Tag 中的原始數據進行精準排序（如大小、日期），避免字串解析誤差。

## 使用者評論請求

> [!NOTE]
> 為了保持程式碼整潔，我將實作一個通用的 `ListViewItemComparer` 類別，它能識別 `OstMailRow` 與 `MailItemInfo` 這兩種結構。這比解析顯示的字串（如 "1,234" 或 "2026/04/23"）更可靠。

## 擬議變更

### Form1_OST.vb

#### [MODIFY] [Form1_OST.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_OST.vb)

1.  **`InitTab7UI`**: 在循環中新增 `AddHandler lv.ColumnClick, AddressOf Lv_ColumnClick`。
2.  **`Lv_ColumnClick`**: 新增此事件處理函式。它將：
    *   檢查點選的欄位是否與上次相同。
    *   切換排序順序 (Ascending / Descending)。
    *   建立 `Tab7ListViewItemComparer` 並指派給 `lv.ListViewItemSorter`。
    *   呼叫 `lv.Sort()`。
3.  **`Tab7ListViewItemComparer`**: 在 `Form1` 類別內（或 `Form1_OST.vb` 檔尾）新增此私有類別。
    *   實作 `IComparer` 介面。
    *   根據欄位索引 (0-4) 進行比較。
    *   針對「大小」(Index 1) 與「日期」(Index 2)，從 `ListViewItem.Tag` 中提取 `OstMailRow` 或 `MailItemInfo` 進行數值比較。

## 驗證計畫

### 手動測試
1. 載入一個 OST 檔案。
2. 點擊「郵件大小」標題，確認是否正確按位元組大小排序（而非字串）。
3. 點擊「收到日期」標題，確認日期順序正確。
4. 點擊「主旨」或「寄件者」，確認按字母順序排序。
5. 再次點擊同一欄位，確認切換遞增/遞減。
6. 對 `LvPST` 進行同樣的操作。
