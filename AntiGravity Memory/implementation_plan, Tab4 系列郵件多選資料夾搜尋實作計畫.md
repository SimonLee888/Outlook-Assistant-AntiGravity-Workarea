# Tab4 系列郵件多選資料夾搜尋實作計畫

本計畫旨在讓 Tab4 (系列郵件) 能夠同時對多個選定的資料夾進行系列主題分析，利用 `SimTree4` 的多選特性提升搜尋效率。

## 使用者評論要求
讓 `Button4` 支援多選資料夾比對，提升搜尋靈活性。

## 擬議變更

### Form1_MainTabs

#### [MODIFY] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_MainTabs.vb)

1. **狀態變更**：
   - 將 `Private _tab4LastSearchFolder As Outlook.Folder` 改為 `Private _tab4LastSearchFolders As List(Of Outlook.Folder)`。

2. **`Button4_Click` 邏輯升級**：
   - 抓取 `SimTree4.SelectedNodes` 中所有具備 `Tag` (Outlook.Folder) 的節點。
   - 若目前無選取，則套用 `_tab4LastSearchFolders` 歷史清單。
   - 遍歷所有選定資料夾，合併其子資料夾清單。

```vb
    ' 修改前
    Dim rootFolder As Outlook.Folder = TryCast(SimTree4.SelectedNode?.Tag, Outlook.Folder)
    
    ' 修改後 (示意)
    Dim selectedFolders As New List(Of Outlook.Folder)
    For Each node In SimTree4.SelectedNodes
        Dim f = TryCast(node.Tag, Outlook.Folder)
        If f IsNot Nothing Then selectedFolders.Add(f)
    Next
    ' ... 歷史紀錄處理 ...
```

3. **掃描邏輯優化**：
   - 使用 `HashSet` 或 `Distinct` 確保合併多個資料夾樹時，重複的子資料夾不會被重複掃描。

## 開放問題
無。

## 驗證計畫

### 手動測試
1. **多選測試**：在 `SimTree4` 中使用 Ctrl 選取兩個不同的資料夾，按下搜尋，檢查結果是否包含這兩個資料夾下的所有系列郵件。
2. **F5 刷新測試**：搜尋完多選資料夾後，清除選取（點空白處），按下 F5，確認是否能正確抓回上一次選取的多個資料夾進行刷新。
3. **單選相容測試**：確保原本的單選操作依然正常運作。
