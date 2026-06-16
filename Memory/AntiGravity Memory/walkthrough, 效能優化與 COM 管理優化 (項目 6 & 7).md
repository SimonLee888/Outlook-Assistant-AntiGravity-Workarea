# 效能優化與 COM 管理優化 (項目 6 & 7)

本階段已完成對 `Outlook Assistant` 核心代碼的 COM 資源管理優化，重點在於防止長期運作下的 RCW (Runtime Callable Wrapper) 記憶體洩漏與減少不必要的 COM 屬性存取。

## 🛠 修改摘要

### 1. COM 集合提取與顯式釋放 (項目 6)
在 `Outlook.Folders`、`Outlook.Items` 與 `Outlook.Stores` 迭代中，改為先將集合指定給區域變數，並使用 `Try...Finally` 確保在迴圈結束後呼叫 `TryMarshalRelease`。這在遞迴 (如 `GetMailCountAllL3`) 與 BFS (如 `GetSubtreeToListL3`) 中至關重要。

*   **[Form1_Outlook.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/Form1_Outlook.vb)**:
    *   `GetSubtreeToListL3`: 優化 BFS 過程中的 `Folders` 集合釋放。
    *   `GetMailCountAllL3`: 優化遞迴過程中的 `Folders` 集合釋放。
*   **[Form1_OST.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/Form1_OST.vb)**:
    *   `CopyFolder_Click`: 針對目標 PST 資料夾檢查、Stores 遍歷、以及 Temp PST 資料夾尋找邏輯進行優化。
    *   `LoadPstSubFoldersRecursive`: 遞迴載入 PST 時的 `Folders` 安全釋放。
    *   `OpenSelectedOstMailViaTempPST`: 開啟 OST 郵件時的 `Stores`、`Folders` 與 `Items` 集合管理。

### 2. 迴圈函數呼叫優化 (項目 7)
將迴圈 `In` 運算式中的重度函數呼叫提取至區域變數，提升代碼可讀性並利於除錯。

*   **[Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/Form1_MainTabs.vb)**:
    *   `BuildBfsFolderTree`: 將 `GetSortedSubFolders` 提取至 `sortedSubs` 變數。

---

## ✅ 驗證結果
- **修復回退**：修正了 `Form1_Outlook.vb` 在修改過程中因 `replace_file_content` 誤匹配導致的 `GetMonthCountsForYear` 損壞問題，已確認該函數恢復正常。
- **邏輯複檢**：所有 `Try...Finally` 結構均正確包裝了對應的 COM 集合變數，確保無論是否發生異常都能正確釋放資源。
- **代碼整潔**：移除了多餘的空行並統一了註解格式。

> [!TIP]
> 這些修改能顯著提升處理包含數千個資料夾的巨型 OST/PST 檔案時的系統穩定性，降低 Outlook 崩潰或出現 "Out of memory" 錯誤的風險。
