# Lambda 風格統一化方案：實作計畫

本計畫旨在針對 `Form1.vb` 與 `Form1_ComL3.vb` 中散亂的 Lambda 語法進行「大掃除」，建立一套兼顧資訊含量、簡潔度與閱讀效率的標準風格。

## 使用者審閱要求
> [!IMPORTANT]
> 1. **保留歷史註解**：必須完整保留所有原始註解與開發歷程紀錄。
> 2. **AntiGravity 標記**：所有修改處需加上 `by AntiGravity, 2026/04/03`。

## 擬議變更規範 (Coding Standards)

1.  **事件參數 (Event Args)**：
    *   統一使用 `(s, e)`。`s`=sender, `e`=event。
    *   移除冗長的 `As Object`, `As EventArgs`（利用型別推論）。
2.  **LINQ 參數 (LINQ Args)**：
    *   改用代表物件類型的首字母縮寫 (如 `f`=Folder, `h`=History, `st`=Store)。
3.  **排版結構 (Formatting)**：
    *   **單行**：僅限單一呼叫或賦值。
    *   **多行**：包含 `If` 判斷或多步邏輯時，強制換行並縮齊。

---

## 擬議變更對照表 (Before vs After)

### [Component 1] Form1.vb 事件與狀態掛載

#### [MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1.vb)

| 位置 (預計行號) | 類別 | 修改前 (Before) | 修改後 (建議統一) | 優點 |
| :--- | :--- | :--- | :--- | :--- |
| **L330, L464, L1223** | 事件處理 | `Sub(s, ev)` | `Sub(s, e)` | 標準化參數命名，減少認知噪音。 |
| **L763, L767** | 事件代理 | `Sub(sender As Object, e As EventArgs)` | `Sub(s, e)` | 減少視覺雜訊，利用推論提升閱讀速度。 |
| **L1202** | LINQ 篩選 | `Function(x) x.Source = source` | `Function(h) h.Source = source` | `h` 直覺連結為 `history`。 |

### [Component 2] Form1_ComL3.vb 資料處理重構

#### [MODIFY] [Form1_ComL3.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_ComL3.vb)

| 位置 (預計行號) | 類別 | 修改前 (Before) | 修改後 (建議統一) | 優點 |
| :--- | :--- | :--- | :--- | :--- |
| **L139** | Store 排序 | `Function(s) s.DisplayName` | `Function(st) st.DisplayName` | 避開與 `sender` 混淆的 `s`。 |
| **L183** | Folder 排序 | `Function(i) i.Name` | `Function(fi) fi.Name` | `fi` 代表 `FolderInfo`，增加語意。 |
| **L269** | Folder 歷遍 | `Sub(current)` | `Sub(f)` | `f` 代表 `Folder`，符合專案常用縮寫。 |

---

## 開放性問題 (Open Questions)

> [!IMPORTANT]
> 1. **關於 RDO 物件 (Redemption)**：在 `Form1_ComL3.vb` L612 等處，目前使用 `Sub(rdoF As Redemption.RDOFolder)`。
>    *   **選項 A**：保留完整註解與型別（為了顯示 RDO 的特殊性）。
>    *   **選項 B**：簡化為 `Sub(rf)` 或 `Sub(rdoF)` 省略型別。
>    *   **您的建議？**
> 2. **參數 s 的衝突**：在處理 Store 的 LINQ 中，您覺得 `st` 本身是否已經足夠區別於 `sender` 的 `s`？

---

## 驗證計畫

### 自動化驗證
- 執行重新編譯，確保所有隱式型別推論 (Implicit Typing) 在 Lambda 中皆能正確運作。

### 手動驗證
- 點擊 UI 各項功能（滑鼠移動、Progress Bar 歷史），確保 Event Handler 依然正常運轉。
