# SimTree4 雙重模式拆分重構 — Phase 1 補全實作計畫

本計畫旨在補全 `Phase 1：新增 LvSearch4` 中尚未完成的所有項目，確保左側系列郵件搜尋結果 ListView (`LvSearch4`) 的功能完全對齊，並徹底與資料夾目錄樹 (`SimTree4`) 的職責分離。

## 使用者審查要求

> [!IMPORTANT]
> 1. **註解保留與標註**：所有修改點均保留原有的除錯歷程註解，新添加的註解統一加上標記：`' by Gemini 3.5 Flash, 2026/05/29`。
> 2. **小塊寫入 (Chunked Edits)**：本次修改涉及兩個檔案，將採用多個精準的 `ReplacementChunk` 進行小範圍替換，降低程式碼損壞的風險。
> 3. **改後立即複檢**：完成修改後，會主動使用 `view_file` 讀取並確認修改點的邏輯是否對齊。

## Phase 1 未完成項目檢查結果與提案

經過對程式碼的深入搜尋與研讀，以下為 Phase 1 尚未完成的細項及其具體修改提案：

---

### 1. 【1-B-5】加入清楚的修改註解，說明雙控制項並存的設計理由
* **位置**：[Form1.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1.vb#L622-L630) 中的 `InitTab4UI`
* **說明**：目前僅有一行簡單的切換顯示註解。
* **提案**：補上詳細的架構說明註解，闡述為何需要以 `LvSearch4` (主旨清單) 與 `SimTree4` (資料夾樹) 雙軌並存來取代舊版的雙重模式。

---

### 2. 【1-D-2】SaveTreeNodeSnap 改判斷條件
* **位置**：[Form1_MainTab345.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_MainTab345.vb#L661-L663) 中的 `Bt4_Click`
* **現狀**：
  ```vb
  'If Not _isTv4ResultMode Then SimTree4.SaveTreeNodeSnap("folder-view")
  ```
  目前被註解，未發揮效用。
* **提案**：取消註解，將條件改為 `If Not LvSearch4.Visible Then`，確保在資料夾視圖下才儲存快照，並附上註解。

---

### 3. 【1-F】修改 `Lv4_ColumnClick` 欄位排序
* **位置**：[Form1_MainTab345.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_MainTab345.vb#L830-L880)
* **現狀**：
  * 第 837 行讀取 `SimTree4.SelectedNode?.Tag`；
  * 第 875 行寫回 `SimTree4.SelectedNode.Tag`。
* **提案**：
  * **1-F-1**：增加 Guard Clause `If Not LvSearch4.Visible Then Return`。改讀 `_tv4SelectedTopicMailList` 替代 `SimTree4` 節點。
  * **1-F-2**：排序完後，將結果寫回 `_tv4SelectedTopicMailList`，不再觸碰 `SimTree4.SelectedNode.Tag`。
  * **1-F-3**：補上對應的修改註解與防禦設計說明。

---

### 4. 【1-G】修改 `RefreshLv4MailsAsync`
* **位置**：[Form1_MainTab345.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_MainTab345.vb#L1240-L1286)
* **現狀**：第 1245 行仍讀取 `SimTree4.SelectedNode?.Tag`。
* **提案**：
  * **1-G-1 & 1-G-2**：改讀並寫入 `_tv4SelectedTopicMailList`。
  * **1-G-3**：加上修改說明註解。

---

### 5. 【1-H-3 & 1-H-6】補全 `LvSearch4_KeyDown` 中的 F6 與說明註解
* **位置**：[Form1_MainTab345.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_MainTab345.vb#L781-L821)
* **現狀**：F6 分支在 `LvSearch4_KeyDown` 中為空，且缺乏獨立 handler 的理由註解。
* **提案**：
  * **1-H-3**：實作 F6 功能，切換 `_tv4GroupSortByCount` 並呼叫 `RenderLvSearch4(_tv4PrevTopicResults)`。
  * **1-H-6**：在函數開頭加入大型說明註解，解釋為何此事件處理器必須獨立於共通的 `HandleLv3Lv4Lv5`。

---

### 6. 【1-I】`HandleLv3Lv4Lv5_MouseClick` 整合 `LvSearch4`
* **位置**：[Form1_MainTab345.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_MainTab345.vb#L317)
* **現狀**：僅處理動態掛載的 ListView，漏掉了手動宣告的 `LvSearch4`。
* **提案**：
  * **1-I-1**：將 `Handles LvSearch4.MouseClick` 加入 `HandleLv3Lv4Lv5_MouseClick` 宣告中。
  * **1-I-2**：加入修改說明註解。

---

### 7. 【1-J】修改 `HandleTvKeyDown` 條件
* **位置**：[Form1.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1.vb#L1272-L1299)
* **現狀**：第 1283 行的結果模式判斷被註解。
* **提案**：
  * **1-J-1**：取消註解，將條件從 `_isTv4ResultMode` 改為 `LvSearch4.Visible`。
  * **1-J-2**：加入修改說明註解。

---

## 驗證計畫

### 手動與功能性驗證 (Tab4 LV與TV)
1. **點擊開始掃描**：確認 `LvSearch4` 順利顯示，`SimTree4` 隱藏。
2. **切換系列郵件**：點擊左側 `LvSearch4` 主旨，右側 `ListView4` 能正確更新與載入。
3. **滑鼠單擊事件**：在 `LvSearch4` 上單擊左鍵，確認主旨能順利複製至剪貼簿 (對齊 1-I 整合)。
4. **鍵盤熱鍵測試**：
   * **ESC**：確認 `LvSearch4` 隱藏，`SimTree4` 還原並重新獲得焦點。
   * **F5**：在 `LvSearch4` 上按 F5 觸發重新掃描。
   * **F6**：在 `LvSearch4` 上按 F6 觸發排序順序切換並重繪。
5. **右側列表排序**：點擊右側 `ListView4` 欄位 (如主旨、時間)，確認能正確在 `_tv4SelectedTopicMailList` 上排序並重繪。
6. **右側列表重整**：在 `ListView4` 上按 F5，確認能從 Outlook 正確重新整理該系列郵件。
