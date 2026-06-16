# 替換 Tab5 手動清理主旨邏輯為 GetCleanSubject()

在 `Form1_MainTabs.vb` 的 Tab5 (重複郵件) 掃描邏輯中，目前有一段手動處理 `RE:`, `FW:` 等前綴的程式碼。為了保持邏輯一致性並減少重複程式碼，建議改用 Tab4 已經實作好的 `GetCleanSubject()` 函數。

## Proposed Changes

### [Outlook Assistant]

#### [MODIFY] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_MainTabs.vb)

- 在 `Bt5_Click` 事件處理器中（約 L3192 處），將手動清理 `subject` 的邏輯替換為 `GetCleanSubject(subject)`。
- 移除原本 `.ToUpper().Replace(...).Replace(...).Trim()` 的鏈式呼叫。

```vb
' 原本的程式碼：
' Dim cleanSubj As String = subject.ToUpper().Replace("RE:", "").Replace("FW:", "").Replace("回覆:", "").Replace("轉寄:", "").Replace(" ", "").Trim()

' 修改後的程式碼：
' Dim cleanSubj As String = GetCleanSubject(subject).Replace(" ", "").ToUpper()
```
> [!NOTE]
> `GetCleanSubject` 內部已經處理了 `While` 迴圈清理巢狀前綴，比原本的單次 `Replace` 更精準。
> 原本程式碼中有 `.Replace(" ", "")` 去除所有空白，我也會予以保留或整合。

## Verification Plan

### Manual Verification
- 執行 Tab5 的「掃描重複郵件」功能。
- 測試包含 `Re:`, `Fw:`, `回覆:`, `轉寄:` 等前綴的郵件是否仍能正確被分組為重複郵件。
- 確認 Fuzzy 模式下的掃描結果與修改前一致（或更精準）。
