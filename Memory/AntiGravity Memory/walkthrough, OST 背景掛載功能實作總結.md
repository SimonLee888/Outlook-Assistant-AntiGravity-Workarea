# OST 背景掛載功能實作總結

我們已經順利完成了 OST 相關的背景掛載進階功能。這次的更新不但繞過了受損的 `Niv2023` 原生寫入限制，更大幅提升了效能與使用者體驗。

## 1. 解除 UI 依賴 (Decoupling)
過去 `Niv2023 ost2pst` 函式庫中的 PST Writer 因為高度依賴本身的 UI（`MainForm.statusMSG`），導致在我們的系統中呼叫時經常拋出 `NullReferenceException` 崩潰。
**處理方式：**
- 修改了 `Niv2023 ost2pst` 的核心架構，導入 `IStatusReporter` 介面。
- 讓內建的 `Program.mainForm` 指向一個背景默默拋給 `Debug.WriteLine` 的 `DummyReporter`。
- 成功解除了這顆定時炸彈，現在我們可以在背景緒安全地呼叫任何 `Niv2023` 的 PST 操作。

## 2. 修復 OST 郵件清單顯示
先前因為 `ReadOstFolderContentsL3` 中計算 CONTENTS_TABLE NID 時使用的類型不符（使用了 `4` 而非 libpff 規範中的 `14`），導致經常抓不到郵件清單。
**處理方式：**
- 已將邏輯從 `Or 4UI` 修正為 `Or 14UI` (SUB_MESSAGES)。這精準地解鎖了 OST 的隱藏內容表，讓所有郵件都能如實顯示。

## 3. OST 資料夾無縫複製
實作了 `CopyFolder_Click` 背景掛載轉移機制，完全避開了之前損壞的 `PST Writer` 寫入流程。
**運作邏輯：**
1. 背景匯出選定的 OST 資料夾成一個暫時的 `.pst` 檔。
2. 將這個暫存 PST 透過 Outlook OOM 掛載進 MAPI session。
3. 尋找暫存 PST 裡剛匯出的資料夾，呼叫穩定的 `Outlook.Folder.CopyTo()` 將它複製到您指定的 目標 PST。
4. 解除掛載，背景清理暫存檔。

> [!TIP]
> 這個做法兼顧了資料完整性（由 Outlook 自行搬移），也完美融合了 OST 唯讀解析的特性。

## 4. 極速開啟單一 OST 郵件
針對之前「單一郵件」的開啟瓶頸，如果您只選取了一封信，我們不需要再花費數分鐘去匯出整個資料夾！
**最佳化解法：**
- 在 `FileManager` 中導入了 `MessagesToExportNIDs` 動態過濾機制。
- 在 `OpenSelectedOstMailViaTempPST` 階段，攔截並修改匯出邏輯：系統只允許「被選取的那些信」被寫入暫存 PST。
- 這意味著暫存的 PST 檔案極小（因為只包含該封信），生成速度飛快。隨後由 OOM 掛載並呼叫 `MailItem.Display()` 開啟預覽視窗。

## 5. TreeView 資料夾數量顯示
如您在圖片中的期望，現在 TreeView 建置完畢後，會立刻在背景啟動掃描：
- 利用非同步讀取 `TableContext` 的 `tcRowMatrix.Count`，計算每個資料夾內的實際項目數量。
- 即時在畫面上把 TreeView 節點更新為 `資料夾名稱 (數量)` 的格式，讓整體版面資訊更加完整。
