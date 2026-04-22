# [Outlook Assistant 效能優化與 Bug 修復計畫]

本計畫旨在解決三個主要問題：
1. **Lv1 (ListView1)**: 有子資料夾的項目按 Enter 或雙擊時，無法正確跳轉焦點的問題。
2. **Lv3 (Tab3)**: 不勾選包含子資料夾時，路徑資訊在 SSD 快取讀回後消失的問題。
3. **SimTree4 (Tab4)**: F5 刷新（以及 Tab5 的搜尋）沒有設計 SSD 快取，導致重新啟動程式後每次都需要重新掃描 COM，耗時過長的問題。

## User Review Required
關於您提到的「沒有原本的 table 可以借用或修改嗎？」以及「優先考慮現有的 save/load 函數」，我仔細分析了資料庫結構：

目前現有的 `attach_maillist` 資料表與對應的 Save/Load 函數（`SaveAttachMailListInner` / `LoadAttachMailListInner`）的確非常接近 Tab4/Tab5 的需求。**但是，這兩者的「掃描邊界」不同**：
- Tab3 (`GetAttachMailListL3`) 為了極致效能，使用了 `@SQL=hasattachment=True` 過濾，所以一個有 1 萬封信的資料夾，可能只會掃描並存入 **5 封有附件的信** 到 `attach_maillist`。
- Tab4/Tab5 需要**所有信件**的主旨與 Topic 才能找系列/重複信，如果借用同一個資料表，Tab4 讀取快取時會以為這資料夾「真的只有 5 封信」，導致嚴重漏信。

**解決方案抉擇：**
- **方案 A (統一掃描，共用資料表)**：修改 `attach_maillist` 增加 `topic` 欄位。並且**廢除 Tab3 的附件過濾**，讓 Tab3、Tab4、Tab5 全部共用同一個「掃描所有信件」的底層函數。
  - *優點*：完美共用現有的 `attach_maillist` 資料表與 Save/Load 函數。
  - *缺點*：如果使用者「只用」Tab3 搜尋，原本只要掃描 5 封信的時間，現在會被強迫掃描 1 萬封信，導致 Tab3 初次掃描變慢。
- **方案 B (維持獨立，新增資料表)**：也就是我原先提議的，新增 `basic_maillist`。Tab3 維持極速的附件過濾掃描，Tab4/Tab5 才去掃描全資料夾並存在新表。
  - *優點*：Tab3 效能不受影響，各司其職不互相污染。
  - *缺點*：需要新增資料表與專屬的 Save/Load 函數（但我會模仿現有的寫，不會增加太多複雜度）。

> [!IMPORTANT]
> **請問您傾向選擇 方案 A 還是 方案 B？** 如果您認為 Tab3 稍微變慢沒關係（因為最終都會快取），那我們可以選方案 A 達到最精簡的代碼庫。

---

## Proposed Changes

### Form1_MainTabs.vb

#### [MODIFY] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_MainTabs.vb)
- **Lv1 導航問題**：您提到「有 subfolder 還進不去」。追查發現 `node.Text = subject` 在字串比對時失效（可能是因為 `SimTree1` 的節點文字帶有空白或特殊字元差異）。我會改用絕對精準的 **`EntryID`** 來做節點比對 (`f.EntryID = t.SubFolder.EntryID`)，並保留原本「沒有子資料夾就不進去」的合理判斷邏輯。

---

### Form1_SQLite2.vb

#### [MODIFY] [Form1_SQLite2.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_SQLite2.vb)
- **Lv3 快取路徑遺失問題**：
  - 在 `DbGetAttachMailList` 中，於讀出資料並建構 `MailItemInfo` 時，補上 `mail.FolderPath = fPath` 賦值。
  - 在 `LoadAttachMailListInner` 中，於建構 `MailItemInfo` 時，補上 `mail.FolderPath = fp` 賦值。
- **Tab4 SSD 快取 (視方案而定)**：
  - 方案 A: 對 `attach_maillist` 進行 `ALTER TABLE` 加上 `topic` 欄位，並修改相關 Save/Load 加入 topic 讀寫。
  - 方案 B: 建立 `basic_maillist` 資料表與專屬 Save/Load。

---

## Verification Plan
1. **Lv1 測試**：進入有子資料夾的目錄，在右側 ListView1 選擇一個有子資料夾的項目，雙擊或按 Enter，確認左側樹狀結構能展開且精準選取該節點。
2. **Lv3 測試**：取消勾選「包含子資料夾」執行搜尋，觀察 ProgressBar 或 Item tooltip 是否有正確顯示路徑；接著關閉程式再重啟，重新執行同樣條件搜尋 (此時必定從 DB 讀取)，確認路徑依然正確顯示。
3. **SimTree4 / Tab4 測試**：執行搜尋並產出結果後，到 Tab7 按下 SaveCache 寫入資料庫。接著重新啟動程式，在 Tab4 按下 F5，觀察是否能立刻瞬間跑完 (代表成功命中 SQLite lazy load)。
