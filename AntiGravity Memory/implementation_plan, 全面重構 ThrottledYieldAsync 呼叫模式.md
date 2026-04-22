# 全面重構 ThrottledYieldAsync 呼叫模式

使用者發現 `ThrottledYieldAsync` 在呼叫前手動檢查 `Stopwatch` 的模式過於冗贅。我們已在 `Form1.vb` 中擴充了 `ThrottledYieldAsync` 支援 `onThrottled` 委派。本計畫將全面掃描整個專案，將所有類似的「手動判斷 + 呼叫」模式統一更換為更簡潔的委派寫法。

## User Review Required

> [!NOTE]
> 這次重構主要是程式碼風格的優化（Syntactic Sugar），邏輯上與原本完全一致，但在可讀性與維護性上有顯著提升。

## Proposed Changes

### 核心異動：模式轉換
原本的寫法：
```vb
If sw.ElapsedMilliseconds >= ThrottleFreq.Hi Then
    progress?.Report(...)
End If
Await ThrottledYieldAsync(sw, cToken, ThrottleFreq.Hi)
```
將統一轉換為：
```vb
Await ThrottledYieldAsync(sw, cToken, ThrottleFreq.Hi, Sub()
    progress?.Report(...)
End Sub)
```

### [核心組件] Form1_MainTabs.vb

掃描並修改以下位置：
- #### [MODIFY] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_MainTabs.vb)
    - `ComputeFolderSize` (~L488)
    - `SimTree2_AfterSelect` (~L957)
    - `CollectYearCounts` (~L1253)
    - `CollectMonthCounts` (~L1293)
    - `GetCachedAttachMailList` (~L2015)
    - `ScanSeriesAsync` 相關邏輯 (~L2225, L2243, L2385, L2429)

### [MAPI 邏輯] Form1_Outlook.vb

掃描並修改以下位置：
- #### [MODIFY] [Form1_Outlook.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_Outlook.vb)
    - `GetMailCountRecursive` (~L488)
    - `GetFolderSizeAsync` (~L867, L926)
    - `GetMailCountByMAPINew` (~L1270)
    - `ExtractFolderSizeOld` (~L1359)
    - `GetUniqueFolderList` (~L1460, L1596)
    - `GetYearCountsForFolder` (~L1828)

### [SQLite 邏輯] Form1_SQLite2.vb

掃描並修改以下位置：
- #### [MODIFY] [Form1_SQLite2.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_SQLite2.vb)
    - `FilterByAttachDetailsAsync` (~L497, L540)

## Open Questions
目前無重大設計問題。所有變動僅限於 UI 進度回報與 CPU 讓出的簡化。

## Verification Plan

### Automated Tests
- 編譯專案，確保 Lambda 語法 (Sub() ... End Sub) 在 VB.NET 中正確。
- 使用 `view_file` 隨機檢查 3-5 個修改過的區塊，確認邏輯對其。

### Manual Verification
- 執行程式，進行「資料夾統計」、「年度統計」與「附件搜尋」，確認進度條 (ProgressBar) 仍能正常更新且沒有閃爍。
- 在統計過程中按下 `ESC`，確認 `cToken` 仍然能透過 `ThrottledYieldAsync` 正確引發 OCE 中斷。
