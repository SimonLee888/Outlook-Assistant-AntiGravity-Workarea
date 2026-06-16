# DebugForm 高度異常與卷軸消失修復計畫

## 目的
解決 `DebugForm` 當高度 (Height) 拉大到 2000px 以上時，縱向卷軸 (Vertical Scrollbar) 消失及 ListView 內容消失的問題。修正先前錯誤將焦點放在寬度的邏輯。

## 使用者確認事項
> [!IMPORTANT]
> 此次修正將移除透過反射設置的 `lvwDebug.DoubleBuffered` 屬性，改為純粹依賴 Win32 `LVS_EX_DOUBLEBUFFER` 擴充樣式。這應能解決極大解析度/尺寸下的緩衝區溢位問題，但需確認是否會導致些微閃爍（預期不會，因為原生樣式通常更穩定）。

## 變動事項

### [DebugForm.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/DebugForm.vb)

#### [MODIFY] `DebugForm_Load`
- **移除冗餘雙緩衝**：刪除使用反射強制開啟 `DoubleBuffered` 的代碼。
- **保留 API 樣式**：維持 `SendMessage` 開啟 `LVS_EX_DOUBLEBUFFER` 的邏輯。

#### [MODIFY] `RecalcColumnWidths`
- **修正高度監測**：更新註解，明確註記 2000px+ 的問題發生在「高度」。
- **優化填滿計算**：調整第一欄寬度計算公式，精確計算與縱向卷軸空間的關係，避免因高度拉長導致誤觸水平卷軸，進而干擾縱向卷軸的顯示。

#### [MODIFY] `lvwDebug_DrawSubItem`
- **增加繪製安全墊**：確保在繪製背景與文字時，即使座標因控制項極大而產生邊際誤差，也能正確處理渲染。

## 驗證計畫

### 手動驗證 (代碼審查)
- [ ] 確認 `RecalcColumnWidths` 不會因為高度變動而陷入佈局死迴圈。
- [ ] 確認雙緩衝部分只有一種機制在運作。
- [ ] 檢查是否所有 `2000px` 相關的註解都已正名為「高度」。

### UI 驗證
- 請使用者在 4K 螢幕或拉長視窗至 2000px 以上，確認縱向卷軸是否穩定存在，且內容不再消失。
