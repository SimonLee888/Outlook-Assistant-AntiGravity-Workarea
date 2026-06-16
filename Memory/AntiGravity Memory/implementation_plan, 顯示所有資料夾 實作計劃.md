# 顯示所有資料夾 (Debug 選項) 實作計劃

本計劃將在 Debug 標籤頁中導入 `checkIncludeAllFolders` 功能，允許使用者動態切換是否過濾非郵件目錄，並在 UI 上對過濾掉的目錄進行特殊標記。

## User Review Required

> [!IMPORTANT]
> **快取機制變動**：因為「顯示全部」與「只顯示郵件」會產生完全不同的 `_cacheFolderTree`，我將在 CheckBox 改變狀態時主動清空快取，確保 UI 立即反應新設定。

## Proposed Changes

---

### [Form1.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual Studio/Visual Studio 18 (2026)/Outlook Assistant - (AntiGravity測試區)/Form1.vb)

#### [MODIFY] GetSortedSubFolders (L890 附近)
- 修改過濾條件：
  ```vb
  If Not checkIncludeAllFolders.Checked AndAlso Not IsMailFolder(subF) Then Continue For
  ```
  這表示：**只有在「沒勾選顯示全部」且「不是郵件資料夾」時才排除。**

#### [MODIFY] GetSubFolderList (L950 附近)
- 同步修改 BFS 遍歷邏輯，確保背景統計數值的正確性（與 UI 同步）。

#### [MODIFY] ListView1 填入邏輯 (尋找 ListView1.Items.Add 位置)
- 判斷目標資料夾：
  ```vb
  Dim lItem = ListView1.Items.Add(folder.Name)
  If Not IsMailFolder(folder) Then
      lItem.ForeColor = Color.DarkGray ' 或可用 Color.FromArgb(160, 160, 160)
      lItem.Font = New Font(ListView1.Font, FontStyle.Italic)
  End If
  ```

#### [NEW] CheckBox 事件處理
- 在 `Form1_Load` 或設計工具建立事件：
  ```vb
  Private Sub checkIncludeAllFolders_CheckedChanged(sender As Object, e As EventArgs) Handles checkIncludeAllFolders.CheckedChanged
      _cacheFolderTree.Clear() ' 核心：清空快取
      ' 觸發目前選取節點的重新整理或是提示使用者需重新選取
  End Sub
  ```

---

## Open Questions

- **關於字體與顏色**：`LightGray` 在白色背景上可能太淡不易閱讀，建議使用 `DarkGray` 或自訂比例的灰色（如 R:180, G:180, B:180），您有偏好嗎？
- **統計數據**：當顯示「所有資料夾」時，統計數值（郵件數、大小）對於非郵件目錄可能都是 0。這是預期行為嗎？

## Verification Plan

### Manual Verification
1. 在 Debug 頁面勾選「顯示所有資料夾」。
2. 重新展開 `SimTree`，確認「行事曆」、「連絡人」等目錄已出現。
3. 檢查右側 `ListView1`，確認非郵件目錄是否呈現為淡灰色斜體。
4. 取消勾選，確認非郵件目錄再次消失。
