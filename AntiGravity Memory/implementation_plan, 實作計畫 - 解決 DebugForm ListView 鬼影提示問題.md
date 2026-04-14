# 實作計畫 - 解決 DebugForm ListView 鬼影提示問題

## 問題背景與原因分析 (Root Cause)
在 `DebugForm` 中，`lvwDebug` 啟用了 `OwnerDraw = True` 並負責自訂訊息的渲染（包括高亮搜尋）。

1. **LabelTip 衝突**：Windows ListView 預設具有一項名為 `LVS_EX_LABELTIP` 的擴充功能。當滑鼠懸停 (hover) 在文字長度超過欄寬的項目上時，系統會自動彈出一個透明或白底的浮動標籤（LabelTip）來顯示完整內容。
2. **OwnerDraw 與系統渲染的競爭**：當 `OwnerDraw` 開啟時，系統產生的 LabelTip 會與我們自繪的文字產生「重疊」現象。這就是用戶看到「鬼影」的原因。
3. **Handle 生命週期問題**：現有程式碼雖然在 `Shown` 事件嘗試移除 `LVS_EX_LABELTIP`，但 WinForms 的 `ListView` 常因屬性變動（如 `ColumnWidth` 或 `View` 改變）而重新建立底層 Handle。一旦重新建立，之前透過 `SendMessage` 設定的 Win32 擴充樣式就會遺失，導致 LabelTip 再次出現。
4. **繪製溢出風險**：在高亮搜尋模式下，程式碼透過手動分段（Normal/Match/Remaining）繪製文字，若字串極長且未在邊界做嚴格截取，即便設定了 `SetClip`，依舊可能在邊界處理不當時造成渲染殘影。

## 擬議變更詳情 (Detailed Implementation)

---

### [Component] DebugForm UI (Form1_DebugForm.vb)

#### [MODIFY] [Form1_DebugForm.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_DebugForm.vb)

##### 1. 擴充 Win32 定義 (Region 01)
- 新增 `LVM_SETTOOLTIPS = (&H1000 + 74)`。
- 確認 `LVS_EX_LABELTIP = &H4000` 存在。

##### 2. 建立穩固的樣式修復機制 (ApplyListViewFixes)
- 建立私有方法 `ApplyListViewFixes()`：
  - 呼叫 `SendMessage(..., LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_LABELTIP, 0)` 移除標籤提示旗標。
  - 呼叫 `SendMessage(..., LVM_SETTOOLTIPS, 0, 0)` 徹底解除 ListView 與內部 ToolTip 控制項的關聯。
  - 確保 `LVS_EX_DOUBLEBUFFER` 也在此重新套用，提升視窗縮放時的渲染穩定性。

##### 3. 攔截 HandleCreated 事件
- 在 `DebugForm_Load` 中，新增 `AddHandler lvwDebug.HandleCreated, AddressOf OnLvwHandleCreated`。
- `OnLvwHandleCreated` 負責再次觸發 `ApplyListViewFixes()`，確保修復隨 Handle 持久存在。

##### 4. 優化 DrawSubItem 渲染邏輯
- **關鍵修復**：在 `isHitCell = True` 的文字分段繪製循環 (`For Each m In matches`) 中：
  - 在繪製每個分塊前，計算該塊的寬度。
  - 若 `currentX + blockWidth` 超過 `textRect.Right`（考慮邊距後），則該塊僅繪製部分內容或直接停止繪製並加入省略號。
  - 確保不論字串多長，繪製內容絕不跳出 SubItem 的範圍。
- **改用 EndEllipsis**：將單塊繪製的 `TextFormatFlags` 加入 `EndEllipsis`。

---

1. **擴充 Win32 API 宣告**:
   - 新增 `LVM_SETTOOLTIPS = (&H1000 + 74)` 常數。
   - 確保 `LVS_EX_LABELTIP` 相關邏輯能正確執行。

2. **強化樣式套用時機**:
   - 建立 `ApplyListViewWin32Fixes()` 方法，整合原有的樣式修改邏輯。
   - 在 `DebugForm_Shown` 呼叫此方法。
   - 額外監聽 `lvwDebug.HandleCreated` 事件，並在事件中再次呼叫此方法。這能確保若 ListView 因屬性變動重新建立 Handle 時，修復不會遺失。

3. **徹底關閉 ToolTip 控制項**:
   - 呼叫 `SendMessage(lvwDebug.Handle, LVM_SETTOOLTIPS, IntPtr.Zero, IntPtr.Zero)`，手動解除 ListView 與 Windows 預設 ToolTip 調度器的關聯。

4. **優化 `lvwDebug_DrawSubItem` 繪製邏輯**:
   - 修正 `isHitCell = True` (搜尋高亮模式) 下的繪製循環。
   - 在繪製每個分段（普通文字與高亮關鍵字）前，檢查 `currentX` 是否已超出 `textRect.Right`。若溢出則停止繪製或強制加入省略號。
   - 將原本的 `TextFormatFlags.WordEllipsis` 調整為 `TextFormatFlags.EndEllipsis`。

## 驗證計畫

### 手動驗證
1. **基本測試**: 開啟除錯視窗，將滑鼠懸停在長訊息上，確認不再顯示白色浮動標籤。
2. **高亮測試**: 輸入搜尋關鍵字，確認高亮狀態下 Hover 也不會出現鬼影，且文字在邊界處正確截斷。
3. **穩定性測試**: 最大化/縮小視窗，確認修復依然有效（Handle 重建後的行為）。

> [!IMPORTANT]
> 此修改主要涉及底層 Windows 訊息攔截與樣式設定，不影響現有的資料排隊與搜尋邏輯。
