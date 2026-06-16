# Tab3 (尋找附件) 終極效能優化計畫 - 究極快取架構 (最終確認版)

根據您的最新意見，我們的思路已經完全打通，並且確認了所有的底層限制與優化手段。針對您的疑問與補充，說明如下：

## 1. 關於時間本位節流 (Throttling) 的超過問題
> 您問：會不會發生有些 mail 讀附件超過 100ms? 如果有, 會發生什麼事?

**答案是：會發生，但完全安全。**
這就是這套 Stopwatch 邏輯的強大之處。假設處理某封超大信件花了 300ms：
1. 迴圈進來檢查時，`swThrottle.ElapsedMilliseconds` 會是 `300`。
2. 因為 `300 >= 100`，所以它會立刻進入 `If` 條件內部。
3. 更新進度條 -> 把 Stopwatch 歸零重新計算 -> 呼叫 `Task.Delay(1)` 讓出執行緒。
**結果**：該次的 UI 畫面更新會從預期的「每 0.1 秒動一下」，變成「這一次等了 0.3 秒才動一下」。系統不會當機，只是進度條稍微停頓了一瞬，然後緊接著就乖乖把控制權還給 UI 了。這保證了 UI 更新間的**最短間隔**，且不會吃掉原本的處理時間。

## 2. 關於 Cache Miss 時的 Redemption 終極加速
> 您問：如果未命中, 使用 Redemption 可以怎麼更有效率的讀回附件檔名?

**Redemption 的極致魔法在於「繞過 MAPI 包裝庫 (OOM)」與「直接讀取 MAPI Table」：**
* 在一般 Outlook COM (`MailItem`) 中，當您呼叫 `.Attachments(i)` 時，Outlook 必須在背後把整個附件物件完全實例化（包含準備好它的內容位元組等），這非常笨重。
* **使用 Redemption (`RDOMail`)**：Redemption 是直接跟底層 Extended MAPI 溝通的。當我們打開 `RDOMail` 時，它的 `RDOMail.Attachments` 底層是一個輕量級的資料表 (MAPI Table)。要讀取檔名，它**不需要實例化檔案內容**，它只是從資料庫欄位裡直接把 `PR_ATTACH_FILENAME` 或 `PR_ATTACH_LONG_FILENAME` 掃過去而已。
* **實作方式**：
  如果專案中有 Redemption，我們可以在 Miss 時改用：
  ```vb
  Dim rSession As New Redemption.RDOSession
  rSession.MAPIOBJECT = _olNS.MAPIOBJECT ' 共用 Outlook Session
  Dim rdoMsg = rSession.GetMessageFromID(mail.EntryID)
  For Each att In rdoMsg.Attachments
      filenames.Add(att.FileName)
  Next
  Marshal.ReleaseComObject(rdoMsg)
  ```
  這速度會比原生的 `Outlook.MailItem` 再快上幾倍到幾十倍！

## 3. 共用現成基礎設施
太棒了，既然您在 `Form1_ComL3.vb` 中已經預先準備了：
`Private Shared ReadOnly _cacheAttachFilename As New ConcurrentDictionary(Of String, List(Of String))`
我們就直接徵用它！

---

## 接下來的實作藍圖 (程式碼翻新)

1. **修正 Phase 1.5 (LINQ 過濾)**：確保它完全接管大小篩選，且不需要任何快取干預。
2. **重構 `ScanAttachmentByName` (Phase 2)** (將更名為 `ScanAttachmentDetail` 等更精確的名稱)：
   * 加入 `_cacheAttachFilename.TryGetValue` 的邏輯。
   * **Hit**：直接進行 List 內的字串迴圈比對，與 `List.Count` 數量比對。
   * **Miss**：透過 Outlook COM (或未來的 Redemption) 取得 `MailItem`，將檔名收集成 `List`，塞入 `_cacheAttachFilename`，然後再次進行比對。
3. **優化節流**：將迴圈中的 `swThrottle` 更新為我們剛剛討論過的標準時間本位節流寫法。

---

一切就緒！這就是我們針對 Tab3 的最後一張設計圖。
一旦您點頭，我馬上幫您生出改寫的程式碼，保證讓 Tab3 的搜尋體驗脫胎換骨！
