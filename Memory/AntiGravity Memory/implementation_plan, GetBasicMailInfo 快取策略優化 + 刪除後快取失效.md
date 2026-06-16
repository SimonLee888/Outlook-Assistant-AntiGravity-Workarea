# GetBasicMailInfo 快取策略優化 + 刪除後快取失效

## 目標
同時達成：
- **選項 A**：`GetBasicMailInfo` 改用 `GetMailCount`（L2.5 快取）取代 `GetLiveFolderSnapL3`（直打 COM），讓記憶體命中時完全 0 COM
- **選項 B**：刪除郵件後主動 invalidate 受影響 fPath 的記憶體快取與 DB 記錄，防止快取污染

---

## 修改清單

### Form1_SQLite2.vb — 新增 DB 刪除 Helper

#### [MODIFY] [Form1_SQLite2.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity測試區%29/Form1_SQLite2.vb)

在 L1716 `DbSaveMonthCountsSingle` End Sub 之後、L1717 `#End Region` 之前，新增一個 helper：

```vb
    Friend Sub DbDeleteBasicMailInfoByPath(fPath As String)
        ' ---------------------------------------------------------------
        ' DbDeleteBasicMailInfoByPath — 刪除郵件後立即清除指定 fPath 的 basic_maillist 記錄
        ' 只針對被刪除郵件所在的資料夾，不影響其他 fPath
        ' 配合 InvalidateBasicMailCache 一起使用，確保 DB lazy load 不會回傳舊資料
        ' 2026/05/11 by Claude Sonnet 4.6
        ' ---------------------------------------------------------------
        If _db Is Nothing OrElse String.IsNullOrEmpty(fPath) Then Return
        Try
            Using cmd As New SqliteCommand("DELETE FROM basic_maillist WHERE folder_path=@p", _db)
                cmd.Parameters.AddWithValue("@p", fPath)
                Dim rows = cmd.ExecuteNonQuery()
                _dbg("DbDeleteBasicMailInfoByPath", $"{ExtractFolderName(fPath)}: 清除 {rows} 筆")
            End Using
        Catch ex As Exception
            _dbg("DbDeleteBasicMailInfoByPath 錯誤", $"{ExtractFolderName(fPath)}: {ex.Message}")
        End Try
    End Sub
```

---

### Form1_Outlook.vb — 兩處修改

#### [MODIFY] [Form1_Outlook.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity測試區%29/Form1_Outlook.vb)

**修改點 1：`GetBasicMailInfo` L763 — 選項 A**

將 `GetLiveFolderSnapL3` 換成 `GetMailCount`：

```vb
' 改前:
Dim currentSnap As Long = GetLiveFolderSnapL3(folder, fPath)

' 改後:
Dim currentSnap As Long = GetMailCount(folder, fPath)  ' 2026/05/11 by Claude Sonnet 4.6: 改用 L2.5 快取，記憶體命中時 0 COM；配合刪除後主動 invalidate _cacheMailCount 確保不污染
```

> **注意**：`currentSnap` 的型別維持 `Long`，`GetMailCount` 回傳 `Long`，型別完全相符，無需額外轉型。

---

**修改點 2：L2.5 Region 新增 `InvalidateBasicMailCache` helper — 選項 B 用**

位置：`GetBasicMailInfo` 函式結束後（L786 之後），`RdoPreloadAttach_1` 之前（L787 之前）。

```vb
    Friend Sub InvalidateBasicMailCache(fPath As String)
        ' ---------------------------------------------------------------
        ' InvalidateBasicMailCache — 刪除郵件後，主動清除指定 fPath 的記憶體快取
        ' 只清 _cacheBasicMailInfo 和 _cacheMailCount 兩個 key，不影響其他資料夾
        ' 配合 DbDeleteBasicMailInfoByPath 一起呼叫，確保記憶體與 DB 兩層同步失效
        ' 2026/05/11 by Claude Sonnet 4.6
        ' ---------------------------------------------------------------
        If String.IsNullOrEmpty(fPath) Then Return
        Dim dummy1 As (Mails As List(Of (Mail As MailItemInfo, Topic As String)), Snap As Long) = Nothing
        _cacheBasicMailInfo.TryRemove(fPath, dummy1)
        Dim dummy2 As Long
        _cacheMailCount.TryRemove(fPath, dummy2)
        _dbg("InvalidateBasicMailCache", ExtractFolderName(fPath))
    End Sub
```

---

### Form1_MainTab345.vb — 兩處修改

#### [MODIFY] [Form1_MainTab345.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity測試區%29/Form1_MainTab345.vb)

**修改點 3：`HandleLv4Delete` (L1071) — 刪除後 invalidate**

在現有的 `For Each item` 迴圈中，收集 fPaths（與 entryIDs 同一個迴圈），在 `MoveMailsToRecycle` 之後清除快取：

```vb
' 在 Dim entryIDs 之後加一行:
Dim affectedPaths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

' 在 For Each 迴圈裡，entryIDs.Add 下方加一行:
If Not String.IsNullOrEmpty(info.FolderPath) Then affectedPaths.Add(info.FolderPath)

' 在 MoveMailsToRecycle 呼叫之後、ShowLv4Result 之前加:
For Each fPath In affectedPaths
    InvalidateBasicMailCache(fPath)     ' 清記憶體
    DbDeleteBasicMailInfoByPath(fPath)  ' 清 DB
Next
```

> Tab4 實際上只有一個 fPath（單資料夾掃描），`HashSet` 保留是為了和 Tab5 保持一致的模式。

---

**修改點 4：`HandleLv5Delete` (L1558) — 刪除後 invalidate**

相同模式，`For Each item` 迴圈中收集 fPaths，`MoveMailsToRecycle` 之後清除：

```vb
' 在 Dim entryIDs 之後加一行:
Dim affectedPaths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

' 在 For Each 迴圈的 entryIDs.Add 之後加一行:
If Not String.IsNullOrEmpty(info.FolderPath) Then affectedPaths.Add(info.FolderPath)

' 在 MoveMailsToRecycle 呼叫之後、RenderLv5Group 之前加:
For Each fPath In affectedPaths
    InvalidateBasicMailCache(fPath)     ' 清記憶體
    DbDeleteBasicMailInfoByPath(fPath)  ' 清 DB
Next
```

> Tab5 可能跨多個資料夾刪除，`HashSet` 確保每個 fPath 只清一次。

---

## 快取行為分析

| 場景 | 行為 |
|---|---|
| **一般情況（無刪除）** | `GetMailCount` 記憶體命中 → `GetBasicMailInfo` 記憶體命中 → **完全 0 COM** |
| **刪除後首次進入** | `_cacheMailCount` 被清 → `GetMailCount` 走 SafeGetDbRow → snap 不符 → L3 更新計數 → 存入記憶體；`_cacheBasicMailInfo` 被清，DB 行被刪 → L3 重掃 → 存入記憶體 → **正確新資料** |
| **SSD 寫入（SaveCache）** | L3 重掃後新資料在記憶體 → 下次 SaveCache 寫入正確新資料 → **無副作用** |
| **folder_stats 表** | 不需動：snap 驗證自動排除舊值 → **自動自癒** |
| **其他 fPath** | 完全不受影響 → **隔離安全** |

---

## 驗證計畫

### 程式碼驗證
- [ ] 確認 `GetMailCount` 回傳型別是 `Long`，與 `currentSnap As Long` 相符
- [ ] 確認 `MailItemInfo.FolderPath` 在 Tab4/Tab5 的刪除流程中確實不為空
- [ ] 確認 `InvalidateBasicMailCache` 和 `DbDeleteBasicMailInfoByPath` 呼叫時機在 `MoveMailsToRecycle` 之後

### 執行期驗證
1. Tab4 刪除一封郵件 → 立即切換到同資料夾重新掃描 → 應不顯示已刪除的郵件
2. Tab5 刪除一封郵件 → 重新掃描同資料夾 → 應不顯示已刪除的郵件
3. 未刪除的情況下正常瀏覽 → `_iLikeNoisy` 開啟確認無 `GetLiveFolderSnapL3` 被呼叫於 Tab4/Tab5 流程中

---

## 執行順序
1. Form1_SQLite2.vb — 新增 `DbDeleteBasicMailInfoByPath`（最無風險，純新增）
2. Form1_Outlook.vb — 新增 `InvalidateBasicMailCache`（純新增）
3. Form1_Outlook.vb — 修改 L763 `GetLiveFolderSnapL3` → `GetMailCount`（單行替換）
4. Form1_MainTab345.vb — 修改 `HandleLv4Delete`（加 HashSet + For Each）
5. Form1_MainTab345.vb — 修改 `HandleLv5Delete`（加 HashSet + For Each）
6. 複檢所有修改點確認正確、複檢修改點前後是否遺留多餘程式碼
