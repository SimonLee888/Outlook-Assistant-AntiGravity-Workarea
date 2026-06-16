# OST 功能實作計畫：開啟郵件與複製資料夾

您提到的 `libpff` 是 C 語言的函式庫，若要在目前的 VB.NET WinForms 專案中直接呼叫會需要額外編譯與 P/Invoke 封裝，較為繁瑣。
幸運的是，專案內已經包含的 `Niv2023 ost2pst` (VB版) **完全具備**實現這兩項功能的能力！

以下是如何利用現有資源達成目標的具體做法：

## 1. 從 OST 的資料夾裡面，打開一封郵件或連絡人或 RSS

**問題點**：OST 是離線快取檔，裡面的郵件沒有直接掛載在 Outlook 中，因此無法直接用 OOM (Outlook Object Model) 的 `Display()` 方法開啟。
**解決方案**：
- **方案 A (屬性提取與自訂顯示)**：我們可以利用 `Niv2023 ost2pst` 的 `ost2pst.LTP.ReadPCs(stream, nbt)` 讀取該封郵件的所有屬性（包含 `PidTagBody` 內文、主旨、寄件者等），並顯示在我們自訂的 `Form1_DebugForm` 或一個新的檢視視窗中。這不需要依賴 Outlook 開啟。
- **方案 B (轉存為 PST 並用 Outlook 開啟)**：建立一個暫存的 PST 檔案，將這封單一郵件的 NBT 節點（NID）透過 `Niv2023` 寫入該 PST。接著用 OOM 掛載這個暫存 PST，找到該郵件並呼叫 `MailItem.Display()`，讓它在真正的 Outlook 視窗中打開。

## 2. 從 OST 的資料夾裡面，複製資料夾包含裡面內容郵件或連絡人或 RSS

**解決方案**：
`Niv2023 ost2pst` 內部已經有一個現成的方法：`CopySourceDatablocksToPST(folderToExport As UInteger, filename As String)`。
這個方法的作用正是「**把指定的 OST 資料夾及其所有內容，匯出成一個標準的 PST 檔案**」。

**實作流程 (在 `CopyFolder_Click` 中)**：
1. **匯出**：取得使用者在 TreeView 選取的 OST 資料夾 NID，呼叫 `ost2pst.FM.CopySourceDatablocksToPST(nid, "temp_export.pst")`。這會產生一個包含該資料夾所有郵件的暫存 PST。
2. **掛載**：使用 OOM (`_olNS.AddStore("temp_export.pst")`) 將其掛載到 Outlook。
3. **複製**：透過 OOM 取得該暫存 PST 內的資料夾，並呼叫 `Folder.CopyTo(目標 PST 資料夾)`。
4. **清理**：使用 OOM 卸載暫存 PST (`RemoveStore`)，並在背景刪除 `temp_export.pst` 檔案。

> [!TIP]
> 這樣一來，我們不僅成功解決了 OST「唯讀不能複製」的問題，而且完全利用了專案內現有的 `Niv2023 ost2pst` 程式碼，不需再引入複雜的 C 語言 `libpff`。

## 下一步行動確認
如果您同意這個做法：
1. 我會先實作 **方案 2 (複製資料夾)** 的流程。
2. 針對 **方案 1 (開啟單一郵件)**，您偏好 **方案 A (自訂視窗顯示)** 還是 **方案 B (匯出到暫存 PST 用 Outlook 開啟)**？
3. 同時，我也會把原本未完成的 **TreeView 節點數量標記** 補上。
