# 實作紀錄 - 解決 DebugForm ListView 鬼影提示與渲染溢出

## 修改成果總結
本次開發解決了當滑鼠懸停在 `DebugForm` 中長字串項目時出現的系統 Ghost ToolTip (LabelTip)，並修復了搜尋高亮模式下自繪文字會「畫穿」欄位邊界的嚴重 Bug。

### 核心變更內容

#### 1. 持久化 Win32 樣式修復 (UI 隔離)
- **問題**：原本在 `Shown` 設定的 Win32 屬性會因為 ListView 重新建立 Handle 而消失。
- **解法**：註冊了 `lvwDebug.HandleCreated` 事件，確保每次 Handle 重建時都會執行 `ApplyListViewFixes()`。
- **新增隔離**：除了移除 `LVS_EX_LABELTIP` 樣式外，額外呼叫 `LVM_SETTOOLTIPS` 切斷 ListView 與原生 ToolTip 控制項的關聯。

#### 2. 優化 OwnerDraw 邊界防禦 (渲染導正)
- **補回旗標**：在高亮拼塊繪製中補回先前漏掉的 `TextFormatFlags.PreserveGraphicsClipping`。
- **邊界預判**：在繪製 Normal/Match/Remaining 各個分塊前，新增數學判斷。一旦偵測到目前的 `currentX + blockWidth` 會超出 SubItem 邊界，立即強制截斷。
- **截斷渲染**：使用 `flags Or TextFormatFlags.EndEllipsis` 進行最終塊的繪製，確保字串過長時優雅收尾且絕不溢出到相鄰欄位。

---

## 程式碼變更細節

### [Form1_DebugForm.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_DebugForm.vb)

#### [Win32 定義與樣式管理](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_DebugForm.vb#L44-L100)
> [!NOTE]
> 現在所有與 ListView 外觀相關的 Win32 設京都集中在 `ApplyListViewFixes` 方法中。

```vb
    Private Sub ApplyListViewFixes()
        Try
            SendMessage(lvwDebug.Handle, LVM_SETEXTENDEDLISTVIEWSTYLE, New IntPtr(LVS_EX_LABELTIP), IntPtr.Zero)
            SendMessage(lvwDebug.Handle, LVM_SETTOOLTIPS, IntPtr.Zero, IntPtr.Zero)
            SendMessage(lvwDebug.Handle, LVM_SETEXTENDEDLISTVIEWSTYLE, New IntPtr(LVS_EX_DOUBLEBUFFER), New IntPtr(LVS_EX_DOUBLEBUFFER))
        Catch ex As Exception
        End Try
    End Sub
```

#### [DrawSubItem 邊界防禦邏輯](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_DebugForm.vb#L630-L680)
> [!IMPORTANT]
> 此修改確保了即使 Windows 嘗試渲染 LabelTip 失敗，我們的自繪文字也不會因為過長而蓋掉 Timestamp (欄位 1)。

```vb
    ' 以分塊繪製中的一段為例：
    If currentX + szNormal.Width > textRect.Right Then
        TextRenderer.DrawText(e.Graphics, normalPart, e.Item.Font, New Rectangle(currentX, e.Bounds.Y, textRect.Right - currentX, e.Bounds.Height), foreColor, flags Or TextFormatFlags.EndEllipsis)
        lastPos = itemText.Length : Exit For
    End If
```

---

## 驗證結果
- **Hover 測試**：將滑鼠移至長字串項目，鬼影文字標籤已徹底消失。
- **高亮溢出測試**：搜尋關鍵字觸發長文字分段顯示，各個區塊在欄位邊界處被正確截斷，不再畫到 Timestamp 欄位。
- **動態穩定性**：調整視窗大小與欄寬後，修復效果依然維持，證實 Handle 重新建立時的事件掛載正確。
