# 進入資料夾後自動計算大小實作計畫 (更新版)

本計畫旨在使用者在 `ListView1` 中按下 `Enter` 進入單一資料夾後，自動計算該層所有子資料夾的大小，且不產生 UI 全選高亮。

## 變更點與協同運作方式

為了達到「進入後自動全算」且「不需要全選高亮」的效果，**兩處都需要進行修改並協同運作**：

1. **修改 `ComputeFolderSize` 內部實作**：
   使其在 `ListView1.SelectedItems.Count = 0`（無選取項目）時，預設將計算對象設為 `ListView1.Items`（該層所有資料夾）。
2. **修改 `EnterSelectedFolder` 的宣告**：
   將 `Private Async Sub` 改為 `Private Async Function ... As Task`，以便在鍵盤事件中可以被 `Await`。
3. **在 `Lv1_KeyDown` 呼叫端進行 Await 與觸發計算**：
   當按下 `Enter` 進入資料夾後，先 `Await EnterSelectedFolder` 確保新資料夾的子目錄渲染完成，接著直接呼叫 `ComputeFolderSize(Nothing, Nothing)`。此時因為剛進入新目錄且無項目被選取，會自動觸發步驟 1 的邏輯，計算該層所有資料夾大小。

---

## 預計修改程式碼

### `Form1_MainTab12.vb`

#### [MODIFY] `EnterSelectedFolder` 宣告 (L895)
```diff
-    Private Async Sub EnterSelectedFolder(selectedItem As ListViewItem)
+    Private Async Function EnterSelectedFolder(selectedItem As ListViewItem) As Task
```

#### [MODIFY] `Lv1_KeyDown` 中的 `Keys.Enter` 處理邏輯 (L210)
```diff
         If e.KeyCode = Keys.Enter Then
             If lv.SelectedItems.Count = 0 Then Return
 
             ' by Gemini 3 Flash, 2026/04/13: 選取多個項目時，改用 MessageBox 顯示數量加總
             If lv.SelectedItems.Count > 1 Then
                 Dim sumDirect As Long = 0 : Dim sumTotal As Long = 0
                 For Each item As ListViewItem In lv.SelectedItems
                     ' 2026/04/13 v2: 移除「所屬父資料夾」欄後，索引回歸原位
                     ' SubItems(1): 郵件數量(直屬)；SubItems(3): 郵件總計(含子孫)
                     Dim strDirect As String = item.SubItems(1).Text.Replace(",", "").Trim()
                     Dim strTotal As String = item.SubItems(3).Text.Replace(",", "").Trim()
                     Dim valDirect As Long = 0 : Dim valTotal As Long = 0
                     Long.TryParse(strDirect, valDirect) : Long.TryParse(strTotal, valTotal)
                     sumDirect += valDirect : sumTotal += valTotal
                 Next
                 MessageBox.Show($"已選取 {lv.SelectedItems.Count:N0} 個資料夾統計結果：" & vbCrLf & vbCrLf &
                                 $"【本層郵件】加總：{sumDirect:N0} 封" & vbCrLf &
                                 $"【包含子樹】加總：{sumTotal:N0} 封", "複選數量加總", MessageBoxButtons.OK, MessageBoxIcon.Information)
                 Return
             End If
 
-            Dim selectedItem As ListViewItem = lv.SelectedItems(0)          ' 獲取點選的資料夾並進入 (單選時維持原邏輯)
-            If selectedItem IsNot Nothing Then EnterSelectedFolder(selectedItem)
+            Dim selectedItem As ListViewItem = lv.SelectedItems(0)          ' 獲取點選的資料夾並進入 (單選時維持原邏輯)
+            If selectedItem IsNot Nothing Then
+                Await EnterSelectedFolder(selectedItem)
+                ' by Gemini 3.5 Flash, 2026/06/27: 進入資料夾後，直接呼叫 ComputeFolderSize 計算該層所有資料夾大小
+                ComputeFolderSize(Nothing, Nothing)
+            End If
             e.Handled = True : e.SuppressKeyPress = True
```

#### [MODIFY] `ComputeFolderSize` 方法 (L972)
```diff
-    Private Async Sub ComputeFolderSize(sender As Object, e As EventArgs)
-        _isUserBusy = True
-        _dbg(" ├ 開始", $"選取項目數: {ListView1.SelectedItems.Count}") ' by Gemini, 2026/04/10: 調整縮排層級為 Level 1
-
-        Try
-            Dim stopwatch As Stopwatch = Stopwatch.StartNew()  ' by Claude Sonnet 4.6, 2026/06/07
-            Dim selectedItems As ListView.SelectedListViewItemCollection = ListView1.SelectedItems  ' 如果有選中項目, 獲取所選中的項目
-            If selectedItems.Count > 0 Then
-                Dim cToken As CancellationToken = OkayNowYouHaveToken() ' ✅ 取得新 Token
-                For Each s As ListViewItem In selectedItems
-                    'If s.Index = 0 Then Continue For ' 若選中本體目錄則跳過 (之前統計速度很慢的時候, 怕計算量太大跑太久)
-                    If s.SubItems.Count > 4 Then s.SubItems(4).Text = "計算中..." Else s.SubItems.Add("計算中...")
-                    ' 提高反應速度, 先占位 (如果已經有FolderSize的子項目就先把它改成「計算中...」, 如果還沒有就先加一個占位用的子項目)
-                Next
-
-                Dim swThrottle As Stopwatch = Stopwatch.StartNew()  ' by Gemini, 2026/04/11; refactored by Claude Sonnet 4.6, 2026/06/07
-                Dim totalCount As Integer = selectedItems.Count
-                Dim processedCount As Integer = 0
-
-                For Each s As ListViewItem In selectedItems
-                    'If s.Index = 0 Then Continue For ' 一樣, 若選中本體目錄則跳過 (之前統計速度還很慢的時候, 怕計算量太大跑太久)
-                    ' 2026/04/13 by Simon/Claude: Tag 升級為 ValueTuple (SubFolder, ParentNode)；群組標題行 / 合計列 Tag=Nothing，直接跳過
-                    If s.Tag Is Nothing Then Continue For
-
-                    Dim t As (SubFolder As Folder, ParentNode As TreeNode) = DirectCast(s.Tag, (SubFolder As Folder, ParentNode As TreeNode))
-                    Dim folder As Folder = t.SubFolder
-                    If folder Is Nothing Then Continue For
-
-                    Dim folderSize As Long = Await GetFolderSizeAllAsync(folder, cToken:=cToken)  ' 2026/3/29 by Gemini: 改為存取 Layer2.5 快取代理，第二次點擊同一資料夾直接命中快取; 2026/04/15 by Claude: 加入 cToken
-
-                    Dim strFolderSize As String
-                    ' by Gemini 3 Flash, 2026/04/20: 資料大小單位統一改為 MB (保留兩位小數)，更能直觀反映 Outlook 佔用情況
-                    ' 2026/6/27 by simon: 根據 mbSize 是否大於等於 1，動態決定格式是要 "N0" 還是 "N2"
-                    If folderSize < 0 Then : strFolderSize = "計算失敗"
-                    Else : strFolderSize = (folderSize / 1024 ^ 2).ToString(If(folderSize >= 1024 ^ 2, "N0", "N2")) & " MB" ' 2026/6/27 by simon: 根據 mbSize 是否大於等於 1，動態決定格式是要 "N0" 還是 "N2"
-                    End If
-                    If s.SubItems.Count > 4 Then s.SubItems(4).Text = strFolderSize Else s.SubItems.Add(strFolderSize)
-
-                    processedCount += 1
-                    ' 2026/04/16 by Gemini 3.0 flash: 改用 ThrottleFreq.Hii + SmartThrottle 與 onThrottled 委派
-                    Await SmartThrottle(swThrottle, cToken:=cToken, ThrottleFreq.Hii,
-                                              Sub() PgrsBar2.Text = $"正在計算資料夾大小: {processedCount:N0} / {totalCount:N0} ({folder.Name})")
-                Next
-            End If
-
-            PgrsBar2.Text = "統計資料夾大小花費了 " & stopwatch.Elapsed.TotalSeconds.ToString("0.00") & " 秒。"
+    Private Async Sub ComputeFolderSize(sender As Object, e As EventArgs)
+        _isUserBusy = True
+        _dbg(" ├ 開始", $"選取項目數: {ListView1.SelectedItems.Count}")
+
+        Try
+            Dim stopwatch As Stopwatch = Stopwatch.StartNew()
+            
+            ' by Gemini 3.5 Flash, 2026/06/27: 若有選取項目則僅計算選取者；若無選取項目，則預設計算 ListView1 內的所有項目
+            Dim targetItems As New List(Of ListViewItem)()
+            If ListView1.SelectedItems.Count > 0 Then
+                For Each item As ListViewItem In ListView1.SelectedItems
+                    targetItems.Add(item)
+                Next
+            Else
+                For Each item As ListViewItem In ListView1.Items
+                    targetItems.Add(item)
+                Next
+            End If
+
+            If targetItems.Count > 0 Then
+                Dim cToken As CancellationToken = OkayNowYouHaveToken()
+                For Each s As ListViewItem In targetItems
+                    If s.Tag Is Nothing Then Continue For ' 排除標題列或合計列
+                    If s.SubItems.Count > 4 Then s.SubItems(4).Text = "計算中..." Else s.SubItems.Add("計算中...")
+                Next
+
+                Dim swThrottle As Stopwatch = Stopwatch.StartNew()
+                ' 僅統計有 Tag（有效資料夾）的項目數量
+                Dim totalCount As Integer = 0
+                For Each s As ListViewItem In targetItems
+                    If s.Tag IsNot Nothing Then totalCount += 1
+                Next
+                Dim processedCount As Integer = 0
+
+                For Each s As ListViewItem In targetItems
+                    If s.Tag Is Nothing Then Continue For
+
+                    Dim t As (SubFolder As Folder, ParentNode As TreeNode) = DirectCast(s.Tag, (SubFolder As Folder, ParentNode As TreeNode))
+                    Dim folder As Folder = t.SubFolder
+                    If folder Is Nothing Then Continue For
+
+                    Dim folderSize As Long = Await GetFolderSizeAllAsync(folder, cToken:=cToken)
+
+                    Dim strFolderSize As String
+                    ' by Gemini 3 Flash, 2026/04/20: 資料大小單位統一改為 MB (保留兩位小數)，更能直觀反映 Outlook 佔用情況
+                    ' 2026/6/27 by simon: 根據 mbSize 是否大於等於 1，動態決定格式是要 "N0" 還是 "N2" (保留原有註解)
+                    If folderSize < 0 Then : strFolderSize = "計算失敗"
+                    Else : strFolderSize = (folderSize / 1024 ^ 2).ToString(If(folderSize >= 1024 ^ 2, "N0", "N2")) & " MB"
+                    End If
+                    If s.SubItems.Count > 4 Then s.SubItems(4).Text = strFolderSize Else s.SubItems.Add(strFolderSize)
+
+                    processedCount += 1
+                    ' 2026/04/16 by Gemini 3.0 flash: 改用 ThrottleFreq.Hii + SmartThrottle 與 onThrottled 委派 (保留原有註解)
+                    Await SmartThrottle(swThrottle, cToken:=cToken, ThrottleFreq.Hii,
+                                              Sub() PgrsBar2.Text = $"正在計算資料夾大小: {processedCount:N0} / {totalCount:N0} ({folder.Name})")
+                Next
+            End If
+
+            PgrsBar2.Text = "統計資料夾大小花費了 " & stopwatch.Elapsed.TotalSeconds.ToString("0.00") & " 秒。"
```

---

## 驗證計畫

### 手動測試步驟
1. 切換至 Tab1。
2. 尋找並選取一個含有複數子資料夾的資料夾。
3. 按下 `Enter` 鍵進入。
4. **預期行為**：進入新資料夾後，新資料夾旗下的所有子資料夾均沒有被藍色高亮全選，但其最右側會自動顯示「計算中...」，並依序完成各個資料夾的大小統計。
