# 重構 Form1_MainTab12.vb 中的巢狀尋找邏輯

本計畫旨在將 `Form1_MainTab12.vb` 中 `Lv1_KeyDown()` 事件函式處理 `Keys.Escape` 時的巢狀搜尋與選取邏輯，抽離成一個獨立且具重用性的輔助子程序（`Sub`），以提高程式碼的可讀性與維護性。

## 使用者審查請求

> [!NOTE]
> 1. 本重構不會改變任何原有的邏輯行為，僅做程式碼結構的優化。
> 2. 抽離出來的獨立 Sub 將包含適當的型別安全轉型保護 (`Try-Catch` 結構)，以防止未來可能加入的其他型別 Tag 導致 `InvalidCastException`。
> 3. 修改後的程式碼中會附帶標記：`by Gemini 3.5 Flash, 2026/05/27`。

## 預計變更

### Form1_MainTab12.vb

#### [MODIFY] [Form1_MainTab12.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_MainTab12.vb)

- **新增獨立輔助程序**：在 `Form1_MainTab12.vb` 中適當位置（例如 `Lv1_KeyDown` 下方或輔助函數區）新增 `SelectFolderInListView` 子程序。
- **重構 `Lv1_KeyDown`**：將 `Keys.Escape` 分支中的巢狀 For 迴圈替換為對 `SelectFolderInListView` 的呼叫。

```diff
             Dim currentNode As TreeNode = SimTree1.SelectedNode     ' 2026/04/13 by Simon/Claude: Tab1 改用 SimTree1
             If currentNode IsNot Nothing AndAlso currentNode.Parent IsNot Nothing Then
                 ' 記下當前 Folder 物件，用於回到上層後在 ListView1 定位游標
                 Dim currentFolder As Folder = TryCast(currentNode.Tag, Folder)
                 Dim parentNode As TreeNode = currentNode.Parent
 
                 ' 用 SimTree1 正確選取父節點 (不呼叫 FireAfterSelect，避免與下方手動計算重複觸發)
                 SimTree1.ClearSelectedNodes()
                 SimTree1.AddSelectedNode(parentNode)
 
                 ' 手動計算統計並渲染 (等同 SimTree1_AfterSelect 的流程)
                 Dim dedupedNodes As List(Of TreeNode) = GetDedupedNodes(SimTree1.SelectedNodes)
                 Dim items As List(Of ListViewItem) = Await ComputeTab1Stats(dedupedNodes, cToken)
                 RenderLv1(items)
 
                 ' 在 ListView1 中找到代表「剛才那個資料夾」的列並移去高亮
-                ' todo: 改用FindLvItemByName()?
-                If currentFolder IsNot Nothing Then
-                    For Each item As ListViewItem In lv.Items
-                        If item.Tag IsNot Nothing Then
-                            Dim t = DirectCast(item.Tag, (SubFolder As Folder, ParentNode As TreeNode))
-                            If t.SubFolder IsNot Nothing AndAlso t.SubFolder.EntryID = currentFolder.EntryID Then
-                                item.Selected = True : item.Focused = True : item.EnsureVisible()
-                                Exit For
-                            End If
-                        End If
-                    Next
-                End If
+                ' by Gemini 3.5 Flash, 2026/05/27: 重構抽離至獨立的輔助子程序
+                SelectFolderInListView(lv, currentFolder)
                 lv.Focus()
             End If
```

## 驗證計畫

### 手動驗證
1. 進入子資料夾後按下 `ESC` 退回上一層，驗證游標是否能正確移回並高亮選取剛才退出的子資料夾項目。
2. 檢查程式編譯是否完全正常，且無編譯錯誤。
