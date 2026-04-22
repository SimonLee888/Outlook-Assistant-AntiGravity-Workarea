# OST 解析器整合完成

我已經完成了按下 **Load OST** 時，讀取唯一的 OST 檔案並解析其資料夾結構顯示在左側 `SimTreeOST` 的功能。

## 變更摘要

### [Form1_OST.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_OST.vb)

- **動態路徑獲取**：
    - 使用 `Application.StartupPath` 結合相對路徑動態尋找 `Backup\OST Parser` 目錄，增加了程式的靈活性。
    - 保留了原本的絕對路徑作為備援（Fallback），確保在不同開發環境下都能運作。

- **解析邏輯整合**：
    - 正確呼叫 `ost2pst.FM.OpenSourceFile` 開啟目標 OST 檔。
    - 呼叫 `ost2pst.FM.GetFolderList` 獲取解析後的資料夾清單。

- **TreeView 填充優化**：
    - **層級建立**：使用 `Dictionary(Of ost2pst.Folder, TreeNode)` 進行快速搜尋，將子資料夾正確歸類於父資料夾之下。
    - **無限迴圈防護**：修正了根目錄的判斷邏輯。由於解析器會將根目錄的父節點設為自己（`f.parent Is f`），我加入了明確的檢查以防止 TreeView 節點嵌套產生死循環。
    - **根目錄命名**：若根目錄名稱為空，自動標示為 `"Root [OST]"` 以提高可讀性。

## 修改細節 (by Gemini 3.5 Sonnet, 2026/04/15)

```vb
' 修正根目錄判斷與填充邏輯，避免無限遞迴 (by Gemini 3.5 Sonnet, 2026/04/15)
For Each f In ost2pst.FM.folders
    Dim nodeName As String = If(String.IsNullOrEmpty(f.name), "Root [OST]", f.name)
    Dim node As New TreeNode(nodeName) With {.Tag = f}
    folderNodes(f) = node

    ' 判斷是否為根目錄：parent 為 Nothing 或者是其父節點等於自己
    If f.parent Is Nothing OrElse f.parent Is f Then
        tv.Nodes.Add(node)
    Else
        ' 如果父節點已經在字典中，則加入父節點下
        If folderNodes.ContainsKey(f.parent) Then
            folderNodes(f.parent).Nodes.Add(node)
        Else
            ' 暫掛或視為根節點
            tv.Nodes.Add(node)
        End If
    End If
Next
```

## 驗證結果
- 已確認 `SimTreeOST` 能清空舊有節點並載入新的結構。
- `_dbg` 已加入載入完成的統計訊息（資料夾總數）。
