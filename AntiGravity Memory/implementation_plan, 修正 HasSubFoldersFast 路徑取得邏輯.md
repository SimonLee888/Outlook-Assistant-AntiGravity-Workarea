# 修正 HasSubFoldersFast 路徑取得邏輯

在 `HasSubFoldersFast` 函數中，目前仍使用手動的 `Try...Catch` 來獲取 `folder.FolderPath`。為了保持代碼一致性並提升魯棒性，我們將改用 `SafeGetPath` 工具。

## 使用者評論請求
請確認以下邏輯處理是否符合預期：
> [!IMPORTANT]
> 如果 `SafeGetPath` 回傳空字串（代表 COM 物件失效或路徑無法讀取），`HasSubFoldersFast` 將直接回傳 `False`。這是因為在沒有路徑的情況下，我們無法查詢快取或資料庫，且該資料夾極可能已損壞，視為「無子資料夾」是相對安全的做法。

## 提議的變更

### [Component Name] Outlook Assistant Core

#### [MODIFY] [Form1_Outlook.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_Outlook.vb)

將 `HasSubFoldersFast` 中的路徑取得邏輯替換為 `SafeGetPath`。

```vb
' 修改前
If String.IsNullOrEmpty(fPath) Then
    Try : fPath = folder.FolderPath
    Catch : Return False : End Try
End If

' 修改後
fPath = SafeGetPath(folder, fPath)
If String.IsNullOrEmpty(fPath) Then Return False ' 2026/04/23 by Gemini 3.0 flash: 若無法取得路徑，視為無子資料夾以求安全
```

## 驗證計畫

### 自動化測試
- 無（本環境不支援自動化單元測試）

### 手動驗證
1. 觀察 TreeView 展開時的行為，確認具有子資料夾的項目仍能正確顯示 "+" 號。
2. 檢查 Debug 視窗是否出現非預期的例外訊息。
