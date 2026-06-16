# OST/PST 目錄樹資料夾過濾功能實作計畫

這個計畫旨在優化 OST 解析分頁 (Tab7) 的目錄樹顯示，過濾掉一些系統或不需要顯示的資料夾（例如 `IPM_SUBTREE`, `~MAPISP(Internal)`, `Drizzle` 等），讓介面更簡潔。

## 使用者確認事項

> [!IMPORTANT]
> 預計過濾的資料夾清單如下（已排除 `IPM_SUBTREE`）：
> - `NON_IPM_SUBTREE`
> - `~MAPISP(Internal)`
> - `Drizzle`
> - `ItemProcSearch`
> - `SPAM Search Folder 2`
> - 以 `~` 開頭的隱藏資料夾

## 擬議變更

### [Form1_OST.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_OST.vb)

#### [MODIFY] `Form1_OST.vb`

1.  **定義過濾函式**：新增一個私有函式 `IsFolderFiltered(name As String)` 用於集中管理過濾邏輯。
2.  **修改 `BuildOstFolderTree`**：在根節點與子節點的建立循環中，加入 `If IsFolderFiltered(displayName) Then Continue For`。
3.  **修改 `LoadPstSubFoldersRecursive`**：同步套用過濾邏輯，保持 OST 與 PST 顯示一致性。

## 驗證計畫

### 手動測試
1. 開啟程式並切換至 **OST 解析** 分頁。
2. 點擊 **Load OST file** 載入一個 OST 檔案。
3. 檢查 TreeView 是否已隱藏上述提到的資料夾。
4. 點擊 **Load PST file** 載入一個 PST 檔案，確認過濾功能在 PST 下也生效。
