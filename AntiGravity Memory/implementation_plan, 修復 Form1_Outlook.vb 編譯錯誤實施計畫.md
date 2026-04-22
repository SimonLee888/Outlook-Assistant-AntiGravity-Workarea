# 修復 Form1_Outlook.vb 編譯錯誤實施計畫

本計畫旨在解決 `Form1_Outlook.vb` 中出現的三項 BC 編譯錯誤，這些錯誤主要源於最近對 `GetSubtreeToList` 進行 Tuple 化重構時，導致的型別不匹配以及資料庫讀取物件屬性缺失。

## User Review Required

> [!IMPORTANT]
> 此修改將變更 `FolderStatsDbRow` 類別的定義，新增 `path` 欄位以儲存資料夾路徑。這是一項破壞性變更，旨在讓資料庫讀取層級能完整支援新版的 Tuple 重構。

## Proposed Changes

### [Component] 資料庫快取層 (SQLite)

#### [MODIFY] [Form1_SQLite2.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_SQLite2.vb)

- 在 `FolderStatsDbRow` 類別中新增 `Public path As String = ""` 欄位。
- 更新 `DbGetSubFolderIDList` 函數，在讀取資料時將 `folder_path` (reader(0)) 賦值給 `row.path`。

---

### [Component] Outlook 資料存取層 (Outlook Assistant)

#### [MODIFY] [Form1_Outlook.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_Outlook.vb)

- **BC30456 修復 (Line 406):** 確保 `row.path` 已定義。
- **BC30311 修復 (Line 1216 & 1539):**
    - 在 `GetMailCountAll` 與 `GetFolderSizeAll` 函數中，將原本預期為 `List(Of Outlook.Folder)` 的 `targetFolderList` 變數型別更正為 `List(Of (Folder As Outlook.Folder, FolderPath As String))`，以匹配 `GetSubtreeToList` 的新回傳型別。

## Verification Plan

### Manual Verification
- 手動檢查程式碼語法。
- 請使用者在 Visual Studio 中重新建置專案，確認上述三項 BC 錯誤已消除。
- 驗證「資料夾統計」與「掃描」功能是否運作正常，特別是從資料庫載入(SSD Hit) 的路徑。
