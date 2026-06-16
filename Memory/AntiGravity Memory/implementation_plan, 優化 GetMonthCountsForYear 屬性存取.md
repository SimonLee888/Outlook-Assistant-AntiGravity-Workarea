# 優化 GetMonthCountsForYear 屬性存取

## 目的
在 `GetMonthCountsForYear` 函數中，重複讀取 `folder.FolderPath` 和 `folder.Name` 會造成不必要的 COM 物件屬性讀取開銷。本計畫旨在將這些屬性預先讀入局部變數中，以提升效能並減少與 Outlook 的互動頻率。

## 使用者回饋要求
> [!NOTE]
> 此變更純屬效能優化，不影響現有邏輯或快取機制。

## 擬議變更

### 1. Form1_Outlook.vb

#### [MODIFY] [Form1_Outlook.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_Outlook.vb)

- 在 `GetMonthCountsForYear` 函數最上方新增 `fName` 與 `fPath` 變數。
- 將所有 `folder.Name` 替換為 `fName`。
- 將所有 `folder.FolderPath` 替換為 `fPath`。
- 新增註解標註此優化。

## 驗證計畫

### 自動化測試
- 無，此為效能細節優化。

### 手動驗證
- 執行程式並執行「年度/月份統計」，確認 `_dbg` 輸出的內容與之前的 `Name` 和 `FolderPath` 一致。
- 確認快取與資料庫儲存邏輯依然運作正常。
