# 重構 FolderPath 取得邏輯為 SafeGetPath 函數

這個計畫旨在解決程式碼中多次出現的 `If String.IsNullOrEmpty(fPath) Then Try : fPath = folder.FolderPath : Catch : End Try` 模式，將其重構為一個統一且安全的輔助函數 `SafeGetPath()`。

## 使用者評論與回饋要求

> [!IMPORTANT]
> 為了讓所有 `Form1` 的 partial 檔案以及 `moduleStore.vb` 都能存取，我將在 `Form1` 類別中定義一個 `Friend Shared` 的 `SafeGetPath` 函數。

## 擬議變更

### Form1 (核心邏輯)

#### [MODIFY] [Form1_Outlook.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_Outlook.vb)

1.  **新增函數**：在 `輔助函數` 區域新增 `SafeGetPath`。
    ```vb
    ''' <summary>
    ''' 安全取得資料夾路徑 (FolderPath)，防止 COM 物件失效或 Nothing 引發的例外。
    ''' </summary>
    ''' <param name="folder">Outlook 資料夾物件</param>
    ''' <param name="existingPath">選擇性：若已存在路徑則直接回傳，減少 COM 呼叫開銷</param>
    Friend Shared Function SafeGetPath(folder As Outlook.Folder, Optional existingPath As String = "") As String
        ' by Gemini 3 Flash, 2026/04/23
        ' 邏輯：優先使用傳入的路徑，若無則嘗試從物件讀取並捕捉所有潛在例外
        If Not String.IsNullOrEmpty(existingPath) Then Return existingPath
        If folder Is Nothing Then Return ""
        Try
            Return folder.FolderPath
        Catch
            ' 捕捉 RCW 已釋放或物件失效的例外
            Return ""
        End Try
    End Function
    ```

2.  **替換現有模式**：
    將所有類似 `If String.IsNullOrEmpty(fPath) Then Try : fPath = folder.FolderPath : Catch : End Try` 的程式碼替換為：
    `fPath = SafeGetPath(folder, fPath)`

    涉及行號（約略）：291, 439, 463, 498, 533, 558, 584, 752, 934, 1058, 1125, 1457, 1678, 1741, 1917, 2019, 2090, 2142。

### 其他組件

#### [MODIFY] [moduleStore.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/moduleStore.vb)
將直接呼叫 `folder.FolderPath` 的地方改為 `Form1.SafeGetPath(folder)`，以增強穩定性。

#### [MODIFY] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_MainTabs.vb)
替換對 `folder.FolderPath` 的直接存取。

## 驗證計畫

### 自動測試
- 編譯專案，確保沒有語法錯誤。
- 檢查 `SafeGetPath` 是否能正確處理 `Nothing` 傳入。

### 手動驗證
- 在 Outlook 環境中運行，切換資料夾，確保路徑顯示正常且無例外崩潰。
- 特別檢查在頻繁切換資料夾導致 COM 物件釋放時，是否能平穩捕捉例外而不彈出錯誤視窗。
