# 優化 ListView4 全選時的相似度計算觸發問題

## 問題描述
在 `ListView4` 執行全選操作（Ctrl+A 或程式碼觸發）時，會導致 `SelectedIndexChanged` 事件被大量觸發。由於該事件中包含耗時的 Jaccard Similarity 相似度計算邏輯，短時間內產生的大量異步任務會導致 UI 卡頓並消耗大量 CPU 資源。

## 解決方案
引入一個模組等級的布林旗標 `_isSelectingAll`。
1. 在執行「全選」邏輯前後，分別將此旗標設為 `True` 與 `False`。
2. 在 `Lv4_SelectedIndexChanged` 事件進入點，檢查此旗標。若為 `True` 則直接返回，跳過計算。
3. 同時優化 `HandleLv3Lv4Lv5_KeyDown` 內的全選邏輯。

## 擬議變更

### [Form1_MainTab345.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_MainTab345.vb)

#### [MODIFY] 宣告區
- 新增 `Private _isSelectingAll As Boolean = False` 用於控制全選狀態。

#### [MODIFY] `HandleLv3Lv4Lv5_KeyDown`
- 在 Ctrl+A 處理邏輯中，針對 `ListView4` 設定 `_isSelectingAll` 旗標。

#### [MODIFY] `Lv4_SelectedIndexChanged`
- 在函數開頭增加判斷：`If _isSelectingAll Then Return`。

## 驗證計畫

### 手動測試
1. 在 Tab4 的 ListView4 中按下 `Ctrl+A`，觀察是否仍會瘋狂觸發「計算中」或日誌輸出。
2. 確認全選後，UI 是否維持流暢。
3. 單擊單一郵件，確認相似度計算邏輯在非全選狀態下依然運作正常。
