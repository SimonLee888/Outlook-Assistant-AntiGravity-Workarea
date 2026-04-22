# 多選郵件開啟效能優化：STA 批次處理架構完成

已經聽從建議，將原本存在效能隱患與可能違反 COM 規範的實作，徹底升級為「單一 STA 執行緒批次處理」架構。

## 問題回顧與解法比較
- **最初的寫法 (`New Thread` x 10)**：選取 10 封信會同時開出 10 個執行緒爭搶 `_olApp` 資源。造成 CPU 滿載與卡頓。
- **3.0 Flash 提出的寫法 (`Task.Run`)**：將工作放到 MTA (多執行緒單元) 的 ThreadPool 中，嚴重違反了 Outlook COM 必須使用 STA (單執行緒單元) 的通訊規範。
- **現在的 3.1 Pro 解法 (單一 `STA` 執行緒)**：將所有選取的 `EntryID` 集結成一個清單，交給**唯一一個**背景執行緒去處理。該執行緒被我們主動掛載了合法的 `ApartmentState.STA`。

## 主要修改內容

### 1. `OpenMailsByEntryIDs(entryIDs)` 核心實作
建立了一個專門負責「批次派送」的新函式。它確保：
- `GetNamespace("MAPI")` 只會呼叫**一次**，省下昂貴的進程內登入開銷。
- 在一個有規律的 `For Each` 迴圈中，有秩序地呼叫 `mail.Display()`。
- `TryMarshalRelease` 能安全、依序地正確釋放每一個 COM 物件。

### 2. 在 `ListView3_KeyPress` 中的改動
當使用者按下 `Enter` 且選取多封郵件時，不會再呼叫多次開啟指令：
```vb
Dim entryIDs As New List(Of String)
For Each idx As Integer In lv.SelectedIndices
    If idx >= 0 AndAlso idx < _lv3MailList.Count Then
        entryIDs.Add(_lv3MailList(idx).EntryID)
    End If
Next
OpenMailsByEntryIDs(entryIDs)
```

## 預期效益
- 若您再次全選 10 封信按下 Enter，CPU 不再會突然飆高，因為我們只額外建立了一個執行緒。
- 郵件的視窗開啟速度會感覺更為連貫且不會互相卡死。
- 不用害怕出現難解的 COM Marshalling Exception！
