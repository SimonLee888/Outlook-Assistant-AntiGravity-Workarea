# 目標描述
我們的目標是讓 `DebugForm` 能夠跟隨 `Form1` 的移動與縮放，並自動將 `DebugForm` 的寬度延展到螢幕最右側邊緣。

## 需要使用者確認的項目
> [!NOTE]
> 1. 當 `Form1` 移動或縮放時，`DebugForm` 會自動對齊至 `Form1` 右側並填滿剩餘螢幕空間。
> 2. 我將一併處理 `Resize` 事件，確保視窗放大縮小時也能同步。

## 預計修改項目

### [Form1 同步邏輯]
#### [MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1.vb)
- 更新 `CheckDebug_CheckedChanged`：在開啟 Debug 視窗時，立即呼叫 `SyncDebugFormPosition()`。
- 新增 `Form1_Move` 與 `Form1_Resize` 事件 handler：
  - 當 Debug 視窗可見時，呼叫 `SyncDebugFormPosition()`。
- 建立私有輔助函數 `SyncDebugFormPosition()`：
  - 更新 `DebugForm.Left` 貼齊 `Form1` 右側 (Me.Left + Me.Width - 12)。
  - 設定 `DebugForm.Top` 與 `DebugForm.Height` 等於 `Form1`。
  - 計算目前的 `Screen` 工作區域。
  - 設定 `DebugForm.Width = Screen.FromControl(Me).WorkingArea.Right - DebugForm.Left` 以強行貼齊螢幕最右側。

### [DebugForm 欄位自動縮放]
#### [MODIFY] [DebugForm.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/DebugForm.vb)
- 更新 `DebugForm_Resize`：
  - 計算 `lvwDebug.ClientSize.Width` 扣除「時間」與「間隔」兩欄固定寬度後的剩餘空間。
  - 動態設定第一欄 (`Debug Message`) 的寬度，使其填滿左右空間。
  - 保留現有的 `EnsureVisible()` 邏輯。

## 驗證計畫

### 手動驗證流程
1. 開啟 Debug 視窗。
2. 拖曳 `Form1` 移動位置，觀察 `DebugForm` 是否同步移動且右側始終貼齊螢幕邊緣。
3. 縮放 `Form1` 高度，觀察 `DebugForm` 高度是否跟隨變化。
4. 關閉並重新開啟 Debug 視窗，觀察初始位置與大小是否正確。

