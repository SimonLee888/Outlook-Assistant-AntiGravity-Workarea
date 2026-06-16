# Debug 訊息輸出一致性與完善化計劃

此計劃旨在優化 `Outlook Assistant` 中 `dbg()` (實為 `Dbg()`) 的調用策略，確保訊息語意一致、格式統一，並補全缺失的結束標記與重要監控資訊。

## 使用者評論要求 (User Review Required)

> [!IMPORTANT]
> 1. **格式統一**：我將統一使用「開始」與「結束」作為函數進入與離開的關鍵字。這能讓 `DebugForm` 的 `FindSimilarPair` 邏輯正確執行「開始/結束」配對，並自動計算函數總耗時。
> 2. **結束標記補全**：許多 UI 事件（如 `MouseClick`）目前只有「開始」而無「結束」，或是早期的 `Return` 路徑漏掉了「結束」。我將逐一補全。
> 3. **參數優化**：對於資料夾操作，我會確保 `Dbg` 的第二個參數 (detail) 帶入 `FolderPath` 或 `Name`；對於運算操作，則帶入處理數量。
> 4. **標記標註**：我添加的註解會加上 `by AntiGravity, 2026/03/31` 的標記，並保留原有的 debug 紀錄與思考過程。

## 擬議變更 (Proposed Changes)

### 1. 核心規則定義
- **進入點**：`Dbg("開始", [關鍵參數])`
- **離開點**：`Dbg("結束", [執行結果/耗時])`
- **中斷點**：`Dbg("已中斷", [原因])`
- **錯誤點**：`Dbg("失敗", [錯誤訊息])`

---

### 2. Form1.vb
#### [MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1.vb)
- **GetSortedStores**: 補上 `Dbg("結束", space.CurrentProfileName)`。
- **Form1_ResizeBegin / ResizeEnd**: 保持現狀 (已有詳細寬高資訊)。
- **InitRedemptionSession**: 補強成功與失敗的語意連貫性。
- **InitLookAndFeel**: 補齊內部調用的子函數進入與離開 Log。

### 3. Form1_Main.vb
#### [MODIFY] [Form1_Main.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_Main.vb)
- **TreeView1_AfterSelect**: 檢查多個 `Return` 路徑 (如 `selectedFolder Is Nothing`)，補上「結束」或「已取消」。
- **ListView1_MouseClick / DoubleClick**: 補上「結束」。
- **ListView1_ItemMenu**: 補上「結束」，並在 detail 加入選取的項目數量。
- **Tab1_EnterSelectedFolder**: 補全 `Return` 點的訊息，確保標記對稱。
- **SimTree2_AfterSelect**: 目前做得不錯，但需檢查所有早鳥 `Return` 點，確保都有「結束」。

### 4. Form1_ComL3.vb
#### [MODIFY] [Form1_ComL3.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_ComL3.vb)
- **GetMailCount / GetFolderCount**: 解除部分重要的註解掉的 `Dbg`，或統一格式。
- **GetMailCountAll**: 確保 ⓪ (RDO), ① (Parallel), ② (Sequential), ③ (Recursive) 的成功路徑最終都回報「結束」或「成功」字眼，且 detail 包含數量。
- **GetFolderSize / GetFolderSizeAll**: 補強大小轉換後的顯示資訊 (如 MB/GB)。

## 開放問題 (Open Questions)

> [!TIP]
> 1. 有些超高頻率調用的函數 (如每封郵件都要跑一遍的 `GetMailSize`)，如果開啟 `Dbg` 可能會造成 Log 視窗爆炸。我建議這些函數保留註解狀態，或僅在「失敗」時才輸出 Log？
> 2. `DebugForm.AddMessage3` 目前會將 `detail` 參數放在括號 `()` 中。如果 `detail` 為空，會產生空的括號。是否需要我順便修掉 `DebugForm` 裡面的這個小細節 (只有不為空才顯示括號)？

## 驗證計劃 (Verification Plan)

### 自動化測試 (模擬操作)
- 運行程式，開啟 `DebugForm`。
- 點選樹狀目錄、ListView、執行統計。
- **檢查重點**：
    - 所有的「開始」行是否都能與「結束」行變色配對。
    - 點選「結束」行時，`Time Span` 欄位是否正確顯示了該函數的總毫秒數。
    - 搜尋結果是否能正確過濾出特定的函數執行過程。

### 手動驗證
- 觀察 Log 內容語意是否通順 (例如：不再出現 "Start" 與 "結束" 混用的情況)。
- 確認高耗時函數 (如載入所有 PST) 是否都有頭有尾。
