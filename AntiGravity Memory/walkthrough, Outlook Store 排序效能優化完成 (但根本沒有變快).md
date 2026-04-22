# Outlook Store 排序效能優化完成

我已經完成了 `GetSortedStores` 函數的重構。這次修改將原本在高頻排序中重複讀取 COM 屬性的行為，改為單次讀取並快取的模式，能大幅提升掛載多個 PST 檔時的介面載入速度。

## 修改內容

### 1. 引入 StoreSortInfo 結構體
在 `Form1_Outlook.vb` 的宣告區新增了 `StoreSortInfo`，用來在排序期間存放暫時性的資料。
[Form1_Outlook.vb:L69-75](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_Outlook.vb#L69-L75)

### 2. 重構 GetSortedStores
[Form1_Outlook.vb:L269-299](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_Outlook.vb#L269-L299)

- **屬性快取 (Property Caching)**：現在會在迴圈中一次性讀取所有 `st.DisplayName`，排序過程完全在記憶體中進行。
- **內聯中文判定**：直接使用您提供的 `ChrW(&H4E00)` 邏輯進行內聯檢查，省去跨函數呼叫。
- **大小寫忽略排序**：加入 `StringComparer.OrdinalIgnoreCase` 確保與 Windows 檔案系統排序行為一致且速度最快。

## 效能對比 (預估)

| 項目 | 舊版 (LINQ) | 新版 (Caching) | 優化重點 |
| :--- | :--- | :--- | :--- |
| COM 讀取次數 | $O(N \log N)$ | **$O(N)$** | 減少跨進程開銷 |
| 排序穩定性 | 取決於 LINQ 實作 | 穩定 | 且不觸發 COM 例外 |
| 中文判定 | 函數呼叫開銷 | **內聯處理** | 減少 Stack 負擔 |

> [!TIP]
> 這次優化後，即便掛載幾十個 PST 檔，`InitOutlookNamespace` 的執行體感應該會接近無瞬間延遲。

## 驗證紀錄
- 確認 `StoreSortInfo` 以 Private 形式定義在檔案內。
- 確認排序邏輯依然保持「中文優先」且「字母排序」。
