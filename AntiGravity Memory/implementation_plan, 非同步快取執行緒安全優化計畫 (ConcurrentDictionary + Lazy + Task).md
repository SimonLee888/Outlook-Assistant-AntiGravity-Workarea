# 非同步快取執行緒安全優化計畫 (ConcurrentDictionary + Lazy + Task)

此計畫針對效能優化重點的第二點進行深入探討。旨在解決高併發情境下，多個執行緒同時請求同一個未快取的資料時，所引發的「重複計算 (Thundering Herd)」問題。

## 💥 核心問題：傳統 GetOrAdd 在 Async 下的缺陷

目前專案中的 Layer 2.5 快取代理層大量使用了 `ConcurrentDictionary`。在實作快取時，我們通常會寫出類似以下的邏輯：

```vb
' 寫法 A (TryGetValue + 判斷)
Dim result As Long
If Not _cacheFolderSizeAll.TryGetValue(key, result) Then
    result = Await 讀取資料庫或COM()   ' <--- 🚨 多個執行緒可能同時跑到這裡
    _cacheFolderSizeAll.TryAdd(key, result)
End If
```

或者使用 `GetOrAdd`，但在非同步中會有困難：
```vb
' 寫法 B (傳統 GetOrAdd)
' 🚨 GetOrAdd 不支援直接 Await 傳入的委派。
' 即使將委派標記為 Async，它也會回傳 Task，導致字典快取的是「正在執行的任務」。
' 若多個執行緒同時觸發 GetOrAdd，.NET 的底層機制會「多次執行」委派，然後只保留最後一個回傳的結果，導致嚴重的資源浪費。
```

當使用者透過 `Parallel.ForEach` 或 `Task.WhenAll` 大量掃描資料夾時，若快取未命中，就會發生多個執行緒**同時**去查 DB 或打 COM 的情況，拖慢整體效能。

## 💡 解決方案：Lazy(Of Task(Of T)) 模式

為了解決這個問題，業界標準做法是將 `ConcurrentDictionary` 的 Value 型別改為 `Lazy(Of Task(Of T))`。

### 運作原理：
1. **Lazy 的特性**：`Lazy(Of T)` 保證傳入的建構函數只會被執行**一次**。
2. **完美結合**：當多個執行緒同時呼叫 `GetOrAdd` 時，`ConcurrentDictionary` 可能會快速建立多個 `Lazy` 物件，但最終**只有一個** `Lazy` 物件會被成功存入字典。
3. **單一觸發**：所有執行緒接下來都會去讀取那個「獲勝的」`Lazy` 物件的 `.Value` 屬性。此時，`Lazy` 內部才會**真正啟動** `Task`。所有執行緒最終都會 `Await` 同一個正在執行中的 `Task`。

### 實際程式碼結構示範

我們需要先修改宣告（以 `_cacheFolderSizeAll` 為例）：

#### [舊代碼]
```vb
Private Shared _cacheFolderSizeAll As New ConcurrentDictionary(Of String, Long)
```

#### [新代碼]
```vb
Private Shared _cacheFolderSizeAll As New ConcurrentDictionary(Of String, Lazy(Of Task(Of Long)))
```

#### [封裝呼叫方法]
在讀取快取時，我們會建立一個輔助方法：

```vb
Private Async Function GetFolderSizeAllSafeAsync(key As String, pFolder As Outlook.Folder) As Task(Of Long)
    ' 1. 準備一個產生 Task 的委派
    Dim taskFactory As Func(Of Task(Of Long)) = 
        Async Function()
            ' 這裡放真正的重度計算 (DB 查詢或 COM 讀取)
            Dim size = Await 真正的計算邏輯()
            Return size
        End Function

    ' 2. 透過 GetOrAdd 放入 Lazy 物件
    Dim lazyTask As Lazy(Of Task(Of Long)) = _cacheFolderSizeAll.GetOrAdd(
        key, 
        New Lazy(Of Task(Of Long))(taskFactory, System.Threading.LazyThreadSafetyMode.ExecutionAndPublication)
    )

    ' 3. 取得並 Await 這個唯一的 Task
    Return Await lazyTask.Value
End Function
```

---

## 🛠 建議實作範圍 (後續可分階進行)

若您未來決定引入此模式，建議優先針對以下 **高成本 I/O 或 COM 操作** 的快取字典進行重構：

1. `_cacheSubTreeList` (展開子樹清單，涉及 DB 與 COM)
2. `_cacheFolderBasicMailInfos` (郵件基礎資訊，極度吃重)
3. `_cacheAttachMailList` (附件郵件掃描)
4. 各種耗時的統計資料：`_cacheFolderSizeAll`、`_cacheMailCountAll`

## 結論
引入 `Lazy(Of Task(Of T))` 會讓程式碼稍微複雜一些，但它是解決非同步高併發快取穿透 (Cache Stampede) 最優雅且絕對安全的方案。

您可以將此文件保存下來，待後續有效能瓶頸或多執行緒優化需求時，再決定是否實行。
