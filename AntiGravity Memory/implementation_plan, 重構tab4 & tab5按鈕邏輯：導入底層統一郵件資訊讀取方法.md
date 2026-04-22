# 重構按鈕邏輯：導入底層統一郵件資訊讀取方法

本計畫旨在優化 `Form1_MainTabs.vb` 中的 `Button4_Click` (系列郵件) 與 `Button5_Click` (重複郵件)，將原有的重複 Outlook Table 遍歷邏輯替換為已在 `Form1_Outlook.vb` 中定義的底層方法 `GetFolderBasicMailInfosL3()`。這將提高程式碼的可維護性，並確保掃描邏輯、節流機制與錯誤處理的一致性。

## 使用者評論要求
> [!IMPORTANT]
> 1. 本計畫將替換 `Button4` 與 `Button5` 內部的 `GetTable` 遍歷迴圈。
> 2. 將使用底層內建的 200ms 節流機制 (`ThrottleFreq.Mid`)。
> 3. 所有註解將標註 `by Gemini 3.0 Flash, 2026/04/19` 並保留開發歷程。

## 擬議變更

### Form1_MainTabs.vb

#### [MODIFY] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_MainTabs.vb)

- **Button4_Click (Tab 4: 系列郵件)**
    - 移除原本在 `For Each folder In targetFolderList` 內部的 `table = folder.GetTable()` 到 `Loop` 的整段邏輯。
    - 替換為 `Dim infoList = Await GetFolderBasicMailInfosL3(folder, needTopic:=True, ct:=cToken)`。
    - 遍歷 `infoList` 並將結果存入 `topicDict`。
    - 保留資料夾層級的進度報告與節流邏輯。

- **Button5_Click (Tab 5: 重複郵件)**
    - 移除原本在 `For Each folder In targetFolderList` 內部的 `table = folder.GetTable()` 到 `Loop` 的整段邏輯。
    - 替換為 `Dim infoList = Await GetFolderBasicMailInfosL3(folder, needTopic:=False, ct:=cToken)`。
    - 遍歷 `infoList` 並根據 `isExact` 模式生成 `hashKey` 存入 `exactDict`。
    - 保留資料夾層級的進度報告與節流邏輯。

## 開放性問題
1. **節流頻率確認**：`GetFolderBasicMailInfosL3` 內部固定使用 `ThrottleFreq.Mid` (200ms)。而按鈕外部仍保有 `ThrottleFreq.Hi` (100ms) 的資料夾間隔。這會導致在處理極小資料夾時有額外的 100ms 延遲，這是否符合預期？(通常影響微乎其微)

## 驗證計畫

### 自動測試
- 編譯並檢查 `Form1_MainTabs.vb` 是否有任何語法錯誤。

### 手動驗證
- 執行程式並點選 `Tab 4` 的「掃描系列郵件」，驗證是否能正確分群並顯示。
- 執行程式並點選 `Tab 5` 的「掃描重複郵件」，驗證 Exact / Fuzzy 模式是否運作正常。
- 點選 ESC 鍵驗證中斷邏輯是否依然有效（透過 `CancellationToken`）。
