# OST 解析器整合計畫

此計畫旨在實作按下 "Load OST" 按鈕時，自動尋找目錄下唯一的 `.ost` 檔案，利用 `ost2pst` 函式庫解析其目錄結構，並將結果顯示在 `SimTreeOST` 控制項中。

## 使用者回饋需求

> [!IMPORTANT]
> 目前 `Form1_OST.vb` 中使用的 `ost2pst` 命名空間假設已在專案中正確引用。如果編譯時找不到該命名空間，可能需要手動將 `Backup/OST Parser/Niv2023 ost2pst/` 下的相關 `.vb` 檔案加入專案中。

## 擬議變動

### [Component] OST 解析與 UI 顯示

#### [MODIFY] [Form1_OST.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_OST.vb)
- **修正路徑獲取**：不再使用硬編碼路徑，改為根據當前專案結構動態定位 `Backup/OST Parser` 目錄。
- **整合 `ost2pst.FM`**：
    - 呼叫 `FM.OpenSourceFile(targetFile)` 開啟唯一 OST 檔。
    - 呼叫 `FM.GetFolderList()` 掃描資料夾。
- **填充 `SimTreeOST`**：
    - 清除舊節點。
    - 使用 `Dictionary` 加速父子節點查找。
    - 修正根目錄判斷（處理 `f.parent Is f` 或 `f.parent Is Nothing` 的情況），避免 TreeView 發生無限迴圈或節點遺失。
    - 標註修改人與日期（by Gemini 3.5 Sonnet, 2026/04/15）。

## 驗證計畫

### 手動驗證
1. 啟動程式，切換至 「OST 解析」 分頁。
2. 點擊 「Load OST」 按鈕。
3. 檢查 `SimTreeOST` 是否正確顯示 `Inbox_2011_GLI_OST.ost` 的資料夾結構。
4. 檢查 `_dbg` 輸出是否包含預期的解析成功訊息。
