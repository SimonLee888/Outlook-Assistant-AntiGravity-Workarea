# DebugForm 新增訊息閃爍修復計畫

## 目的
修復在執行 `AddMessage3` (Timer批次新增) 時，清單畫面會發生閃爍的問題。

## 原因分析
在我們先前修復「高度超過 2000px 導致項目消失」的 Bug 時，曾經嘗試拔除了 `DebugForm_Load` 中使用反射設定 `DoubleBuffered = True` 的代碼，因為當時懷疑是雙層緩衝機制衝突導致 GDI 崩潰。

但在隨後的深度排查中，我們已經確認了 2000px Bug 的 **真正元兇** 是：在捲軸隱藏的瞬間（`ClientSizeChanged`）呼叫了 `BeginUpdate()`。這個元兇已經被我們徹底消滅了。
而 WinForms `ListView` 原生的繪圖機制，如果沒有開啟底層的 `DoubleBuffered`，在滾動與批次新增時，會不斷觸發 `WM_ERASEBKGND` (清除背景) 而造成頻繁的白底閃爍（Flickering）。

## 變動事項

### [DebugForm.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/DebugForm.vb)

#### [MODIFY] `DebugForm_Load`
恢復透過反射啟用 `ListView.DoubleBuffered` 的代碼，從本源上防止 WinForms 預設的背景擦除閃爍。

```vb
        ' ✅ 啟用 ListView 雙緩衝，減少閃爍
        SendMessage(lvwDebug.Handle, LVM_SETEXTENDEDLISTVIEWSTYLE, New IntPtr(LVS_EX_DOUBLEBUFFER), New IntPtr(LVS_EX_DOUBLEBUFFER))
        Me.DoubleBuffered = True    ' 表單本身也啟用 DoubleBuffered 減少視窗閃爍

        ' 2026/04/01 by AntiGravity: 恢復 ListView 內建雙緩衝設置
        ' 先前為了排查 2000px 高度 Bug 暫時移除，現已確認該 Bug 兇手為 ClientSizeChanged 內的 BeginUpdate。
        ' 恢復此設定可徹底避免 AddMessage3 (Timer 批次新增) 時產生的背景擦除閃爍。
        Dim pi = lvwDebug.GetType().GetProperty("DoubleBuffered", BindingFlags.Instance Or BindingFlags.NonPublic)
        If pi IsNot Nothing Then pi.SetValue(lvwDebug, True, Nothing)
```

#### [MODIFY] `Timer_Tick`
將 `EnsureVisible()` 移出 `BeginUpdate()...EndUpdate()` 的範圍。在 `EndUpdate()` 後才執行視野滾動，確保繪圖生命週期正確，進一步減少滾動瞬間的抖動。

```vb
            With lvwDebug
                .BeginUpdate()
                .Items.AddRange(itemsToAdd.ToArray())
                .EndUpdate()
                .Items(.Items.Count - 1).EnsureVisible()
            End With
```

## 驗證計畫
### 手動驗證
1. 觀察 `AddMessage3` 執行並輸出資料時，是否恢復先前的滑順、不再閃屏。
2. 再次將視窗高度拉長超過項目總長（測試 2000px 邊界），確認項目**絕對不會**再次消失（因為我們已經拆除了危險的 `BeginUpdate` 炸彈）。
