# 實作計畫 v2 - 解決 DebugForm ListView 鬼影提示與繪製溢出問題

## 問題背景與原因分析 (Root Cause) - 已整合 3.1 Pro 補充
在 `DebugForm` 中，`lvwDebug` 的 UI 渲染問題由兩個層級的缺陷組成：

1. **系統層級：LVS_EX_LABELTIP 與 Handle 重建 (由 3.0 Flash 發現)**
   - Windows ListView 預設會顯示 `LabelTip` 浮動標籤，這與 `OwnerDraw` 自繪文字重疊。
   - WinForms 的 `ListView` 經常因為屬性異動（如欄寬調整）而重新建立底層 Handle，導致原本在 `Shown` 設定的移除樣式失效，LabelTip 隨即復發。

2. **繪製層級：搜尋高亮模式缺失 Clipping (由 3.1 Pro 發現)**
   - **嚴重 Bug**：在 `isHitCell = True` (搜尋命中) 時，手動分塊繪製 (Normal/Match/Remaining) 完全漏掉了 `TextFormatFlags.PreserveGraphicsClipping` 旗標。
   - 這代表一旦我們拔除了系統的 LabelTip，加長的文字會直接「畫穿」目前的 SubItem 格子，蓋掉後面欄位的內容。這是因為手動拼接繪製時缺乏數學邊界計算。

## 擬議變更詳情 (Detailed Implementation)

---

### [Component] DebugForm UI (Form1_DebugForm.vb)

#### [MODIFY] [Form1_DebugForm.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_DebugForm.vb)

##### 1. 擴充 Win32 定義 (Region 01)
- 新增 `LVM_SETTOOLTIPS = (&H1000 + 74)`。
- 確保 `LVS_EX_LABELTIP = &H4000` 與 `LVS_EX_DOUBLEBUFFER` 等常數正確。

##### 2. 建立穩固的樣式修復與隔離機制
- 實作 `ApplyListViewFixes()`：
  - 移除 `LVS_EX_LABELTIP` 擴充樣式。
  - **切斷連結**：呼叫 `SendMessage(lvwDebug.Handle, LVM_SETTOOLTIPS, 0, 0)` 徹底撤銷 ListView 的 Tooltip 調度器。
  - 重新佈署 `LVS_EX_DOUBLEBUFFER` 確保雙緩衝穩定。
- **持久化修復**：修改 `DebugForm_Load`，註冊 `AddHandler lvwDebug.HandleCreated, AddressOf OnLvwHandleCreated`。只要控制項重建 Handle，立即重新套用修復。

##### 3. 優化 DrawSubItem 邊界防禦邏輯
- **補回 Clipping 旗標**：在高亮分塊繪製的 `TextRenderer.DrawText` 參數中，強制加入 `TextFormatFlags.PreserveGraphicsClipping`。
- **手動邊距安全檢查**：
  - 在拼接分塊前，預判 `currentX + blockWidth` 是否超出 `textRect.Right`。
  - 若超出，則該塊強制使用 `TextFormatFlags.EndEllipsis` 且不再繪製後續塊。
  - 確保不論字串長度，自繪內容絕對被「鎖」在格子內。

---

## 驗證計畫

### 手動驗證
1. **鬼影測試**：Hover 長字串，確認不再出現系統浮動標籤（LabelTip）。
2. **溢出測試**：輸入搜尋字串使長文字高亮，確認高亮文字不會蓋到右側的「Timestamp」或其他欄位。
3. **Handle 重建測試**：手動拖拉欄位寬度或縮放視窗（觸發 Handle 重新分配），確認上述功能依然正常。

> [!IMPORTANT]
> 所有的修改將標註 `by Gemini 3 Flash, 2026/04/13` 並保留既有的 Debug 註釋歷程。
