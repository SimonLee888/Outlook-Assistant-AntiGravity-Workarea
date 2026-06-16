# 將 ListView 全選與複製功能抽離為獨立函數

這個計畫的目的是將散落在 `Lv1_KeyDown`、`Lv2_KeyDown` 與 `HandleLv3Lv4Lv5_KeyDown` 等各處的 `Ctrl+A` (全選) 與 `Ctrl+C` (複製到剪貼簿) 的邏輯抽離出來，集中到 `Form1.vb` 中的輔助函數區塊。

## User Review Required

> [!IMPORTANT]
> 請確認以下將要進行的修改：
> 1. 將在 `Form1.vb` 的 `#Region "  ├ ListView 格式工具"` 區塊內新增 `ListViewSelectAll(lv As ListView)` 與 `ListViewCopyToClipboard(lv As ListView)` 兩個輔助函數。
> 2. `Form1_MainTab12.vb` 中的 `Lv1_KeyDown` 與 `Lv2_KeyDown` 將修改為呼叫上述函數。
> 3. `Form1_MainTab345.vb` 中的 `HandleLv3Lv4Lv5_KeyDown` 將修改為呼叫上述函數，取代原有的虛擬模式 (VirtualMode) 選擇邏輯。
>
> 這樣可以簡化原本冗長且重複的 `KeyDown` 程式碼結構。

## Open Questions

> [!NOTE]
> 目前 `HandleLv3Lv4Lv5_KeyDown` 有特別處理一個全域變數 `_isCtrl_A` 來防止大量選取時觸發過多的計算（例如 Jaccard 相似度）。
> 我預計會把設定 `_isCtrl_A` 的邏輯保留在 `HandleLv3Lv4Lv5_KeyDown` 中，或者直接放進 `ListViewSelectAll` 中，但若是放進全域輔助函數，可能會影響到不需要這個變數的 Tab1/Tab2。
> 預計方案：在 `ListViewSelectAll` 裡判斷如果控制項是 ListView4 或 ListView5 才去操作 `_isCtrl_A`，或是由呼叫端自己處理。為了最少依賴，我會在 `ListViewSelectAll` 加上一個對 `_isCtrl_A` 變數的處理，因為這是之前針對效能特別設定的。是否同意這樣的處理？

## Proposed Changes

### Form1.vb

#### [MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1.vb)
在 `#Region "  ├ ListView 格式工具"` 區塊中加入以下兩個函數：
- `Public Sub ListViewSelectAll(lv As ListView)`: 處理一般與虛擬模式的 `Ctrl+A` 全選邏輯。
- `Public Sub ListViewCopyToClipboard(lv As ListView)`: 處理 `Ctrl+C` 的字串組成與剪貼簿寫入。

---

### Form1_MainTab12.vb

#### [MODIFY] [Form1_MainTab12.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_MainTab12.vb)
- **Lv1_KeyDown**: 將原本 `Keys.A` 與 `Keys.C` 的實作替換為呼叫 `ListViewSelectAll(lv)` 與 `ListViewCopyToClipboard(lv)`。
- **Lv2_KeyDown**: 同上。

---

### Form1_MainTab345.vb

#### [MODIFY] [Form1_MainTab345.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_MainTab345.vb)
- **HandleLv3Lv4Lv5_KeyDown**: 將原本 `Keys.A` 結合虛擬模式防呆的冗長邏輯，替換為在設置 `_isCtrl_A` 後直接呼叫 `ListViewSelectAll(lv)`。

## Verification Plan

### Manual Verification
1. 啟動應用程式。
2. 進入 Tab1 點擊 ListView1，按下 `Ctrl+A` 確認是否可以全選，按下 `Ctrl+C` 後到 Notepad 貼上確認格式是否正確。
3. 進入 Tab2 重複上述動作。
4. 進入 Tab3/Tab4，按下 `Ctrl+A` 確認是否可以快速全選且不會卡頓（確保 `_isCtrl_A` 防止了重複的相似度計算）。
