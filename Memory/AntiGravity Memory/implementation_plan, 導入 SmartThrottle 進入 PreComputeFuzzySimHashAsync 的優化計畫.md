# 導入 SmartThrottle 進入 PreComputeFuzzySimHashAsync 的優化計畫

根據我們先前的討論，`PreComputeFuzzySimHashAsync` 作為熱路徑迴圈，原本使用 `(i And 15) = 0` 主要是為了效能與避免 GC 壓力。但我們現在可以透過一個更聰明的方式，將 `SmartThrottle` 導入，達到**「兩全其美」**的境界：既擁有極致的 CPU 效能與零記憶體分配，又能享受 `SmartThrottle` 帶來的穩定 UI 刷新與取消機制。

## Proposed Changes

### Form1_Maintab56.vb

#### [MODIFY] Form1_Maintab56.vb

我們將修改 `PreComputeFuzzySimHashAsync` 內的節流邏輯。

**變更重點：**
1. **宣告一個外部的 `swThrottle`**：作為時間節流的計時器。
2. **結合次數檢查與時間檢查**：保留 `If (i And 15) = 0 Then` 作為第一道閘門。這道閘門幾乎零成本。
3. **第二道時間閘門**：在第一道閘門內，檢查 `swThrottle.ElapsedMilliseconds >= ThrottleFreq.Hii`。
4. **無閉包的 SmartThrottle 呼叫**：符合時間條件時，才進行字串安插、進度回報，並且直接 `Await SmartThrottle(swThrottle, cToken, ThrottleFreq.Hii)`。我們**不使用** `onThrottled` 委派，因此完全不會產生 Closure (閉包) 的記憶體分配問題。

**修改前的程式碼：**
```vbnet
            If (i And 15) = 0 Then                                          ' 每 16 封讓出訊息泵 + 回報進度(ESC 可即時中斷)
                Dim eta = CalculateSpeedAndETA(todo.Count, i + 1, swEta.Elapsed.TotalSeconds)   ' 2026/06/17 by Simon/Claude: 加入速度與 ETA 顯示，對齊 Tab3/Tab4 做法
                progress?.Report(New ProgressReport With {.Message = $"計算內文指紋: {i + 1}/{todo.Count} ({eta.Speed:F0} 個/秒{eta.EtaString})"})
                Await Task.Delay(1, cToken)
            End If
```

**修改後的程式碼預期：**
```vbnet
            ' 外層迴圈前加入: Dim swThrottle As Stopwatch = Stopwatch.StartNew()
            ...
            ' 迴圈內:
            ' 雙重閘門：先用極低成本的位元運算擋掉 15/16 的 Stopwatch 檢查
            If (i And 15) = 0 Then
                ' 再用 Stopwatch 確保只有時間到了 (例如 100ms) 才更新 UI 並讓出
                If swThrottle.ElapsedMilliseconds >= ThrottleFreq.Hii Then
                    Dim eta = CalculateSpeedAndETA(todo.Count, i + 1, swEta.Elapsed.TotalSeconds)
                    progress?.Report(New ProgressReport With {.Message = $"計算內文指紋: {i + 1}/{todo.Count} ({eta.Speed:F0} 個/秒{eta.EtaString})"})
                    ' 不傳入 onThrottled 委派，避免迴圈產生 closure allocation，單純使用其 Delay 與 OCE 機制
                    Await SmartThrottle(swThrottle, cToken, ThrottleFreq.Hii)
                End If
            End If
```

## User Review Required

> [!NOTE] 
> 這樣的改法保留了您原本 `(i And 15) = 0` 的極致效能精神，也補上了我們統一管理 `SmartThrottle` 節流與拋出 `OperationCanceledException` 的安全機制。您覺得這個方向可以嗎？

## Verification Plan

### Manual Verification
- 編譯後，執行 S3 的 Build pass 掃描。
- 確認 UI 進度條能夠平順跳動，並且沒有效能下降的感覺。
- 在掃描途中按下 ESC 中斷，確認 `SmartThrottle` 能夠如預期般拋出取消例外並終止處理。
