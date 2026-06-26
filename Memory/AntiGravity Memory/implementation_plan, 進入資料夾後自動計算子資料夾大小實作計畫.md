# 進入資料夾後自動計算子資料夾大小實作計畫

本計畫旨在當使用者在 `ListView1` (資料夾導覽) 中選取單一項目並按下 `Enter` 鍵進入資料夾後，自動全選所有子資料夾，並觸發大小計算。

## 變更項目

### 1. 將 `EnterSelectedFolder` 宣告改為 `Task`
目前 `EnterSelectedFolder` 為 `Async Sub`，無法被 Await。為了在進入資料夾並渲染完畢後再執行全選與計算，我們將其改為 `Async Function ... As Task`。

### 2. 修改 `Lv1_KeyDown` 的 `Keys.Enter` 分支
在 `Lv1_KeyDown` 中，若單選且按 Enter 進入資料夾後，使用 `Await` 等待 `EnterSelectedFolder` 執行完畢，接著呼叫 `LviSelectAll` 全選子資料夾，最後呼叫 `ComputeFolderSize` 計算它們的大小。

---

## 預計修改程式碼

### `Form1_MainTab12.vb`

#### [MODIFY] `EnterSelectedFolder` 宣告
```diff
-    Private Async Sub EnterSelectedFolder(selectedItem As ListViewItem)
+    Private Async Function EnterSelectedFolder(selectedItem As ListViewItem) As Task
```

#### [MODIFY] `Lv1_KeyDown` 中的 `Keys.Enter` 處理邏輯
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
+                ' by Gemini 3.5 Flash, 2026/06/27: 進入資料夾後，全選所有新渲染出來的子資料夾，並呼叫 ComputeFolderSize 計算大小
+                LviSelectAll(lv, Nothing)
+                ComputeFolderSize(Nothing, Nothing)
+            End If
             e.Handled = True : e.SuppressKeyPress = True
```

---

## 驗證計畫

### 手動測試步驟
1. 開啟程式並切換至 Tab1。
2. 在 ListView1 中選取一個含有子資料夾的資料夾。
3. 按下 `Enter` 鍵進入。
4. 預期行為：進入該資料夾後，ListView1 會顯示所有子資料夾，並且這些子資料夾會被自動全選，且最右側會自動顯示「計算中...」並開始更新各資料夾的大小。
