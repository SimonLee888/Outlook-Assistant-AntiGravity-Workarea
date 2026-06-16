# Outlook Assistant 集合容量優化實作計畫

根據效能分析結果，將針對全專案中處理大量資料的 `List(Of T)` 進行初始容量預分配優化。

## 擬改善目標

### 1. 高價值目標 (初始容量 1024 ~ 2048)
處理郵件資訊或大量 UI 資料列，此類優化對減少記憶體碎片與提升大資料量下的響應速度最有感。

#### [MODIFY] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/Form1_MainTabs.vb)
- `_lv3MailList` (2048)
- `targetMails` (1024)
- `resultList` (1024)
- `mBodyList` (1024)
- `groupItems` (1024)

#### [MODIFY] [Form1_Outlook.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/Form1_Outlook.vb)
- `result` (1024)
- `resultList` (1024)

#### [MODIFY] [Form1_SQLite2.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/Form1_SQLite2.vb)
- `Mails` (1024)

---

### 2. 中等價值目標 (初始容量 512)
處理資料夾結構、樹狀節點或暫存的 UI 項目，確保在處理數百個節點時達成零 Resize。

#### [MODIFY] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/Form1_MainTabs.vb)
- `_tab4LastSearchFolders` (512)
- `_tab4FolderTreeNodesBackup` (512)
- `allItems` (512)
- `items` (512)
- `itemsList` (512)

#### [MODIFY] [Form1_Outlook.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/Form1_Outlook.vb)
- `infoList` (512)
- `nodeList` (512)
- `fList` (512)
- `dbResults` (512)

---

## 驗證計畫
### 自動化檢查
- 透過編譯檢查確保語法正確。
- 透過 View 工具確認所有修改點皆已加上正確的容量參數與註解標記。

### 手動驗證
- 檢查各 Tab 頁面切換與資料載入是否正常。
- 觀察大資料量（如數千封郵件）掃描時的記憶體佔用與流暢度。

> [!NOTE]
> 註解將統一標註為 `by Gemini 3 Flash, 2026/05/04`，並簡述選擇該容量的理由。
