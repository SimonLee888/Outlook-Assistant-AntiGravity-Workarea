# 平行預載快取擴充 (Parallel Cache Pre-warming) 實作計畫

## 目標說明
在不改變既有 L1 介面邏輯，以及維持 L3 循序檢索之安全性的前提下，於 L2（流程協調）與 L2.5（快取代理）層間，加入基於 Redemption (`_rdo`) 的多執行緒快取預載機制。藉由背景預先填滿快取，將原本循序存取的延遲時間移除，大幅提升 Phase 2 的效率。

## 需要使用者確認的重大變更
> [!IMPORTANT]
> 此次實作將會在 `Form1_ComL3.vb` 引入 `Parallel.ForEach`，因為 `_rdo` 設計上為 Free-Threaded，因此平行處理不會有 STA 執行緒問題。若運行環境中沒有 `_rdo`，程式會立刻返回，安全退回原有的 OOM 循序處理。請確認此處理流程與預期一致。

## Proposed Changes

### L2.5 快取代理層
負責實際的預載邏輯。當有 RDO 物件時，將發動多執行緒快速建立 `_cacheAttachFilename` 字典資料，不會變更既有 L3 取出附件檔名的底層邏輯。

#### [MODIFY] Form1_ComL3.vb
- **位置**: `#Region "  ├ L2.5 快取代理層"`
- **新增函數**: `PreloadAttachmentCacheRDOAsync(sourceList As List(Of MailItemInfo)) As Task`
    - 檢查 `_rdo` 狀態，若為 `Nothing` 則直接 `Return`。
    - 使用 `Await Task.Run(...)` 發起背景執行緒。
    - 內部使用 `Parallel.ForEach` 遍歷傳入的 `sourceList`。
    - 在迴圈內，略過已存在於 `_cacheAttachFilename` 的 EntryID。
    - 透過 `_rdo.GetMessageFromID(mail.EntryID)` 取得郵件並擷取附件檔名。
    - 釋放 COM 物件 (`TryMarshalRelease`) 並統一加入快取中。
    - 使用 `Try...Catch` 防禦不可預期的 COM 例外，出現單筆例外時安靜結束，交由原先的 OOM 處理安全網接手。

---

### L2 業務過濾層
負責發號施令。不改變任何原有的 `For` 迴圈架構與 `Continue For` 邏輯，僅在進入迴圈前，對外要求預熱。

#### [MODIFY] Form1_Main.vb
- **位置**: `Private Async Function FilterByAttachmentDetailsAsync(...)` 的開頭。
- **變更**:
    - 在初始化迴圈相關變數（`Dim swTotal`、`Dim keyword` 等）後，插入新行 `Await PreloadAttachmentCacheRDOAsync(sourceList)`。
    - 利用既有的 `progress` 顯示一條非特定的「正在預載...」之類的狀態文字（或者不顯示，維持原本 UI 的無痕切換）。

## 預期結果與驗證
*   **架構層面**: `Button3_Click` 及所有的 UI 元件皆無須配合修改。
*   **效能方面**: 若環境中已載入 Redemption，系統會自動在幾百毫秒內將幾百封信的附件名稱載入至快取，接下來的 Phase 2 迴圈能幾乎瞬間完成且不再有 UI 卡頓現象。
*   **安全網測試**: 若主動將 `_rdo` 設為 `Nothing`，功能應當與目前完全一樣照常運行，退回原始速度。

## Open Questions
不需要額外提問便可實作，若您同意此計畫，我們將依照此文件進行調整。
