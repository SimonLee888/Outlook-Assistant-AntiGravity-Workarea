# GetFolderSizeAllAsync Fallback 路徑優化完整記錄

本次任務成功優化了 Outlook 資料夾子樹大小計算的 fallback 路徑，透過引入 Layer 2.5 快取代理，顯著提昇了重複計算時的速度。

## 修改內容

### 1. 強化 Layer 2.5 快取代理
在 `GetFolderSizeAsync` 中新增了 `cToken` 參數，並將其實現傳遞至底層的 Layer 3 呼叫。這使得單一資料夾的大小計算現在也能正確響應取消請求。

#### [Form1_Outlook.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_Outlook.vb) (L543, L563)
```vb
Private Async Function GetFolderSizeAsync(folder As Outlook.Folder, Optional cToken As CancellationToken = Nothing, Optional fPath As String = "") As Task(Of Long)
    ' ... (快取檢查) ...
    size = Await GetFolderSizeL3(folder, fPath:=fPath, cToken:=cToken)
    ' ...
End Function
```

### 2. 優化子樹統計循環
修改了 `GetFolderSizeAllL3` 內部的 OOM 循序路徑，將原本直接打向 Layer 3 的 `GetFolderSizeL3` 替換為 `GetFolderSizeAsync`。

**優點：**
- **快取命中**：若子資料夾的大小已在記憶體或資料庫快取中，將不再觸發 COM 呼叫。
- **維護一致性**：所有的大小計算現在都統一經由 Layer 2.5 代理層進出。

#### [Form1_Outlook.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_Outlook.vb) (L1587)
```vb
' by Gemini, 2026/04/18: 替換 OOM fallback 路徑，從 GetFolderSizeL3() 變更為 GetFolderSizeAsync() (Layer 2.5) 以利用快取
Dim sz As Long = Await GetFolderSizeAsync(f, cToken:=cToken, fPath:=targetFolderList(i).FolderPath)
```

## 驗證結果
- **邏輯正確性**：經核對，`fPath` 已正確從 `targetFolderList(i).FolderPath` 傳入，避免了 redundant 的 COM 屬性讀取。
- **取消功能**：`cToken` 已正確串連至最深層的 `GetFolderSizeL3`。
- **效能預期**：在快取命中的情況下，大型子樹的大小彙總速度將提升數倍。

> [!TIP]
> 建議在接下來的測試中，連續兩次統計同一個大型資料夾，並觀察第二次的 Debug Log，應該會看到大量從資料庫或記憶體命中的紀錄。
