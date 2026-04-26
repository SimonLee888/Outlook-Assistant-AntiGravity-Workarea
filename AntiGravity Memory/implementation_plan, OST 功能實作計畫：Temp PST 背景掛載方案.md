# OST 功能實作計畫：Temp PST 背景掛載方案

感謝您願意給我最後一次機會嘗試這個路線。針對您提到的「寫出 Temp PST 並在背景掛載」以及「UI 強綁定」的問題，我已經整理出具體的實作步驟與解法。

## 挑戰與對策

### 1. 解除 Niv2023 的 UI 強綁定
**問題**：`Niv2023` 的 `FileManager.vb`、`LTP.vb` 等核心檔案中，頻繁呼叫了 `ost2pst.Program.mainForm.statusMSG(...)` 來回報進度。由於我們的 `Outlook Assistant` 專案並沒有初始化這個 `mainForm`，這會導致匯出 PST 時拋出 `NullReferenceException` 崩潰。
**解法**：
- 我們會在執行匯出之前，於 `Form1_OST.vb` 動態初始化一個隱藏的 `ost2pst.ost2pst` 實例，並賦值給 `ost2pst.Program.mainForm`。
- 或者，修改 `ost2pst.Program.vb`，提供一個靜態的虛擬處理機制，讓 `statusMSG` 的呼叫靜默通過或導向系統的 `Debug.WriteLine`，徹底斬斷它對實體 UI 視窗的依賴。

### 2. 複製資料夾 (包含內容)
**流程**：
1. **背景匯出**：當使用者選擇 OST 資料夾並執行複製時，呼叫 `ost2pst.FM.CopySourceDatablocksToPST(資料夾NID, 暫存PST路徑)`。這會將該資料夾結構與所有郵件完整寫入一個標準的 PST 檔。
2. **OOM 背景掛載**：呼叫 `_olNS.AddStore(暫存PST路徑)` 將其掛載至 Outlook。
3. **執行複製**：使用 OOM 尋找這個掛載的暫存 PST 中的目標資料夾，並呼叫 `TargetFolder.CopyTo(目的PST資料夾)`。
4. **清理善後**：呼叫 `_olNS.RemoveStore(...)` 卸載暫存 PST，然後使用 `System.IO.File.Delete` 刪除實體檔案。

### 3. 開啟單一郵件 / 連絡人 / RSS
**流程**：
如果要在原生的 Outlook 視窗中打開它：
1. **部分匯出**：為這封郵件建立一個暫存 PST。由於 `Niv2023` 原生是「以資料夾為單位」匯出，若資料夾龐大會很慢。為此，我們可以攔截/覆寫 `Niv2023` 中的 `ToBeExported` 旗標邏輯，讓它「只匯出這單一封郵件」。
2. **掛載與開啟**：掛載這個只有一封信的暫存 PST。
3. **Display**：透過 OOM 抓到這封信，呼叫 `MailItem.Display()` 顯示在螢幕上。
4. **清理**：關閉視窗後卸載並刪除暫存 PST。

---
## 下一步
這個計畫完全依循您的指示：「**用 Niv2023 的 pst writer 作出 temp.pst 並背景掛載**」，同時妥善處理了 UI Callbacks 導致崩潰的隱患。

如果您同意，我將先著手修改 UI 綁定的問題，然後在 `Form1_OST.vb` 實作「複製資料夾」的背景掛載邏輯。
