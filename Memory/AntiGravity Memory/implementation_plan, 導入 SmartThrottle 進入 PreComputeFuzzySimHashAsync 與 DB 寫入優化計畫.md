# 導入 SmartThrottle 進入 PreComputeFuzzySimHashAsync 與 DB 寫入優化計畫

您剛才的觀察非常精準，一針見血指出了效能的盲點！

## 分析：為什麼您是對的？

1. **UI 更新的浪費 (i And 15)**：
   正如您所算，若每秒 400 封，`i And 15` (16封) 意味著每秒要更新 UI高達 25 次！進度條重繪與字串處理 (`$"..."`) 是很昂貴的，這確實反而拖累了效能。
   若改用 `SmartThrottle` 的時間閘門 (100ms)，每秒只更新 10 次，UI 負擔反而大幅減輕。我們在計畫中的「雙重閘門法」正是要解決這個問題，不僅大幅降低 UI 更新次數，連檢查 Stopwatch 的微小開銷都省了。

2. **DB 寫入的頻繁 I/O (batch.Count >= 500)**：
   這是極好的洞察！每秒 400 封代表每 1.25 秒就觸發一次 `SaveDbMail(batch)`。SQLite 雖然有 Transaction 撐腰，但頻繁的磁碟寫入 (Commit) 絕對會造成每秒一小次的 I/O 停頓 (Stutter)。這是拖慢整體速度的真正隱形殺手。

## Proposed Changes

### Form1_Maintab56.vb

#### [MODIFY] Form1_Maintab56.vb

1. **優化 SmartThrottle 節流與 UI 更新 (避免 closure 與頻繁重繪)**
   我們將 `i And 15` 加大為 `i And 63 = 0` (每 64 封查一次)，並且裡面再包一層 `SmartThrottle` 時間檢查。
   這保證了：CPU 不用一直查時間，而 UI 則穩定每 100ms (1秒10次) 更新，徹底消除您提到的 UI 浪費。

2. **拉高 DB 批次寫入的緩衝區 (Batch Size)**
   將 `batch.Count >= 500` 修改為 `batch.Count >= 2000` 或 `3000`。
   這能讓 DB 大約每 5~8 秒才寫入一次，大幅降低磁碟 I/O 的頻繁鎖定與寫入開銷，讓 CPU 能更專注在計算 SimHash 上。

**修改後的程式碼預期：**
```vbnet
        Dim swThrottle As Stopwatch = Stopwatch.StartNew() ' 2026/06/25 by Gemini: 獨立時間閘門
        Dim batch As New List(Of (EntryID As String, SimHash As Long, BigramCount As Integer))(3072)
        ' ... 略 ...
        For i As Integer = 0 To todo.Count - 1
            ' ... 略 ...
            
            ' 2026/06/25 by Gemini: 將 Batch 拉高到 3000，大幅降低每秒都在 Commit DB 的 I/O 停頓
            If batch.Count >= 3000 Then SaveDbMail(batch) : batch.Clear()

            ' 2026/06/25 by Gemini: 雙重閘門。先擋 63/64，減少查表。
            ' 時間到了才更新 UI，解決原本 i And 15 在高速下每秒狂刷 UI 25次造成的效能浪費
            If (i And 63) = 0 Then
                If swThrottle.ElapsedMilliseconds >= ThrottleFreq.Hii Then
                    Dim eta = CalculateSpeedAndETA(todo.Count, i + 1, swEta.Elapsed.TotalSeconds)
                    progress?.Report(New ProgressReport With {.Message = $"計算內文指紋: {i + 1}/{todo.Count} ({eta.Speed:F0} 個/秒{eta.EtaString})"})
                    Await SmartThrottle(swThrottle, cToken, ThrottleFreq.Hii) ' 僅作時間重置與讓出，不傳入委派避免 GC
                End If
            End If
        Next
```

## User Review Required

> [!TIP] 
> 將 Batch Size 從 500 提高到 3000，能顯著改善大量掃描時的 I/O 順暢度。如果您覺得 3000 太多或太少，我們也可以隨時調整。這個更新版的計畫您覺得可以開始執行了嗎？
