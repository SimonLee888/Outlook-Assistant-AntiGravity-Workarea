# ProgressBar 狀態歷史紀錄功能實作計畫 (修訂版)

依據您的回饋，我已經調整了實作計畫。以下是針對您的問題與新需求的詳細解答與修訂方案：

## 針對您的問題與回饋

1. **ProgressBar2 的紀錄與「過渡期」過濾**
   - **問題**: 如何記錄 ProgressBar2 但不被 100ms 的「過度狀態」洗版？
   - **解法 (Smart Overwrite 原則)**: 當我們監聽 `ProgressBar2.TextChanged` 時，檢查新進來的字串。如果它與「歷史清單中最新的一筆」**開頭前幾個字相同**（例如都以「正在統計郵件數:」開頭），我們就**直接覆寫**最新的一筆，而不是新增一筆。
   - 這樣一來，10/100 會被 20/100 覆寫，直到最後停在 100/100。歷史紀錄裡只會乾淨地留下一筆「最終極」的細節狀態！

2. **筆數提高至 100 筆**
   - 沒問題，100 筆純字串 `List(Of String)` 對現代電腦來說連 1MB 的萬分之一都不到，絕對游刃有餘。容量我們就設定為 100 筆。

3. **100 筆顯示會不會卡？美不美觀？排序方向？**
   - **排序方向**: 為了方便閱讀，我們會讓**「最新的紀錄放在最上面」**（索引 0）。
   - **美觀度與流暢度**: 若使用單純的 `ContextMenuStrip`，100筆會長出「上下導覽箭頭」，這點有時不好用滾輪滑動。因此，我計畫改用更進階的做法：**動態生成一個 `ListBox` 放入 Popup 容器中**。
   - `ListBox` 支援**滑鼠滾輪**、有標準的垂直捲軸、而且可以設定固定高度（例如顯示 15 筆的高度，剩下的用捲動的）。這樣畫面不但乾淨俐落，找尋歷史紀錄也會非常順暢，絕不會醜也不會卡！

4. **時間戳記存哪裡？**
   - 我們會建立一個專屬結構 `List(Of StatusHistoryItem)`，裡面包含 `.Time` (時間) 與 `.Message` (文字)。
   - 當顯示在前台 ListBox 時，才自動組合成 `[14:35:20] 統計花費 2.50 秒`。這樣日後如果想把文字複製到剪貼簿，時間或本文想怎麼調整都很靈活！

---

## 具體程式碼實作方案 (Proposed Changes)

#### [NEW] `moduleStore.vb` (加入歷史紀錄儲存庫與方法)
因為您把 ProgressBar 置於 Module 之中管理，我們可以把歷史結構集中定義在 `moduleStore` 內：
```vb
Public Structure StatusHistoryItem
    Public Time As DateTime
    Public Message As String
    Public Source As String ' "PB1" 或 "PB2"
End Structure

Public _statusHistory As New List(Of StatusHistoryItem)()
Public Const MAX_HISTORY_COUNT As Integer = 100

' 提供一個公用方法來寫入紀錄
Public Sub AppendStatusHistory(msg As String, source As String)
    ... 實作覆寫連續相同進度與限制 100 筆的邏輯 ...
End Sub
```

#### [MODIFY] `Form1_Main.vb` (或掛載事件的地方)
- **新增監聽器**:
  由於 WinForms 中 `ToolStripStatusLabel` 的 `TextChanged` 事件很好用，我們只要在其 `TextChanged` 事件裡面呼叫 `AppendStatusHistory` 即可。
- **新增 Click 彈出選單事件**:
  當點擊 `ProgressBar1` 時，實體化一個 `ToolStripControlHost` 裡面包著我們客製化好尺寸的 `ListBox`。
  為該 `ListBox` 掛載 `SelectedIndexChanged` 事件：
  1. 取得使用者點擊的項目。
  2. 利用 `Clipboard.SetText` 複製到剪貼簿。
  3. 關閉彈出視窗。

## 結論
這樣一來，無論是 `ProgressBar1` 的跳轉，還是 `ProgressBar2` 的進度條猛轉，都會以最優雅、不洗版的方式存放在 100 筆歷史清單中。點擊左下角，會有一個支援滾輪、最新在最上層、點擊即複製的美觀清單。

請問以上的修訂細節符合您的期待嗎？確認後我將開始撰寫程式碼！
