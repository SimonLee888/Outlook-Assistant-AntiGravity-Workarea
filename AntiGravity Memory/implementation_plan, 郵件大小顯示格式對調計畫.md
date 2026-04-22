# 郵件大小顯示格式對調計畫

目前 `ListView3` 的郵件大小是「精細到位元」，而 `ListView4` 是顯示為 「xx KB」。根據使用者要求，我們需要將這兩者的顯示格式進行反轉。

## 使用者審核點
> [!NOTE]
> 1. 此變更僅涉及顯示格式，不影響排序引擎（排序仍將基於原始位元組數值與日期物件進行）。
> 2. 日期格式改為 `yyyy/MM/dd` (例如 `2026/04/09`) 並置中對齊。
> 3. `ListView3` 與 `ListView4` 的欄位比例將統一致化（主要調整 `ListView3` 的寄件者比例至 20%）。

## 預計變更內容

### Form1_MainTabs.vb

#### [MODIFY] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_MainTabs.vb)

1.  **ListView3 (附件搜尋 - 虛擬模式)**
    - 修改位置：`ListView3_RetrieveVirtualItem` (約於 L1858)
    - 變更：
        - 大小：將 `mail.Size.ToString("N0")` 改為 `(mail.Size \ 1024L).ToString("N0") & " KB"`。
        - 日期：將 `mail.ReceivedTime.ToShortDateString()` 改為 `mail.ReceivedTime.ToString("yyyy/MM/dd")`。
    - 對齊設置：在 `Form1.vb` 之 `InitTab3UI` 或 `Form1.Designer.vb` 中確保日期欄位 (Index 2) 設為 `HorizontalAlignment.Center`。
    - **by Gemini 3.0 Flash, 2026/04/20**

2.  **ListView4 (系列搜尋 - 分組顯示)**
    - 修改位置：`FillListView4` (約於 L2608)
    - 變更：
        - 大小：將 `(mailItem.Size \ 1024L).ToString("N0") & "KB"` 改為 `mailItem.Size.ToString("N0")`。
        - 日期：將 `mailItem.ReceivedTime.ToShortDateString()` 改為 `mailItem.ReceivedTime.ToString("yyyy/MM/dd")`。
    - 對齊設置：在 `Form1.vb` 之 `InitTab4UI` (約於 L639) 將日期欄位 (`"收到日期"`) 設為 `HorizontalAlignment.Center`。
    - **by Gemini 3.0 Flash, 2026/04/20**

### Form1.vb

#### [MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1.vb)

1.  **InitTab4UI**
    - 在欄位初始化處補上 `HorizontalAlignment.Center` 設定。

2.  **AutoResizeListViewColumns** (約於 L1558)
    - 將 `ListView3` 的寄件者（Index 3）寬度從 `0.15` (15%) 調升至 `0.2` (20%)，與 `ListView4` 保持一致。
    - **by Gemini 3.0 Flash, 2026/04/20**

### 手動驗證
- **Tab 3 & Tab 4**: 
    - 確認郵件大小欄位與日期欄位的顯示格式符合預期。
    - 觀察兩者的欄位比例（尤其是寄件者與日期）是否視覺上達到一致。
