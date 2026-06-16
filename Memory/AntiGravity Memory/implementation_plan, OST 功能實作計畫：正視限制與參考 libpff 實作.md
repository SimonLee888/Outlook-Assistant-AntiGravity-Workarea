# OST 功能實作計畫：正視限制與參考 libpff 實作

> [!CAUTION]
> **深刻致歉**：我犯了嚴重的錯誤，完全忽略了 `Form1_OST.vb` 頂端明確寫著的 `不使用 ost2pst PST writer (已知損壞)` 註解。您說得完全正確，如果 `Niv2023` 的寫入功能正常，我們當初就不需要堅持用 OOM 來寫入 PST 了。我之前的提議（匯出成 PST 再掛載）是倒退且錯誤的。我也為草率拒絕參考 `libpff` 原始碼向您道歉。

## 針對您的核心疑問解答

**2-1. 既然都已經可以讀取該郵件的所有屬性了，為何無法開啟該郵件？**
核心原因在於 **Outlook OOM 的運作機制**。
OOM 的 `Namespace.GetItemFromID(EntryID)` 只能在 **「目前已經掛載到 Outlook 的資料檔 (PST/Exchange)」** 中尋找物件。因為我們讀取的 OST 是一個「孤立檔案」，並沒有（也無法）被 OOM 掛載。所以，即使我們從 OST 裡面把這封郵件的屬性、甚至是 100% 完美的 EntryID 都萃取出來了，只要把這個 EntryID 丟給 OOM，OOM 在自己現有的 Store 清單裡找不到這個 OST，就會直接拋出找不到項目的錯誤。
**這就是為什麼我們無法「直接原地打開」OST 裡面的郵件。**

## 解決方案：參考 libpff 的做法

既然目標是「驗證讀取無誤（打開）」與「真正的複製（OOM 寫入）」，而且 `Niv2023` 的寫入功能已損壞，我們必須參考 `libpff` 的原始碼來尋找解答。

在 `libpff` 中，`pffexport` 的核心邏輯是將 OST 內的資料節點（包含所有 Properties 與 Attachments）提取出來，然後：
1. **直接組合並寫出成 `.msg` 檔案**（利用 `libfmapi`）。
2. 或**寫出成 `EML` / `TXT` 檔案**。

### 具體的實作路線 (不依賴損壞的 Niv2023 寫入功能)

#### 任務一：打開單一郵件 (驗證讀取結果)
因為無法直接呼叫 OOM 打開孤立的 OST 郵件，我們採用**轉譯暫存**的方式：
1. 當使用者雙擊 ListView 裡的一封信時，我們透過 `Niv2023` 讀取該 NBT 的所有屬性（主旨、內文、寄件者等）。
2. 在背景透過 OOM 於「草稿匣」或「暫存資料夾」建立一個空白的 `MailItem`。
3. 將讀取到的屬性（如 `mail.Subject = ...`, `mail.HTMLBody = ...`）填入。
4. 呼叫 `mail.Display()`。
> 這樣使用者就能看到一個原汁原味的 Outlook 郵件視窗，完美達到「確認讀取內容無誤」的目的。若關閉視窗，該暫存郵件可選擇丟棄。

#### 任務二：複製資料夾 (OOM 手動重建)
真正的 OOM 複製流程（這正是當初設計 Phase 3 的初衷）：
1. 在目標 PST (透過 `SimTreePST` 選取) 中，使用 OOM 建立一個同名的新資料夾 (`Folders.Add`)。
2. 針對來源 OST 資料夾中的每一封郵件 (透過我們剛修復的 NID 掃描邏輯)：
   - 使用 `Niv2023` (或參考 `libpff` 邏輯補強的讀取器) 讀出所有屬性。
   - 使用 OOM 在目標資料夾建立新項目 (`Items.Add`)。
   - 屬性對拷後呼叫 `Save()`。
> 這徹底避開了 `Niv2023` 損壞的 PST 寫入器，完全依賴 OOM 穩定地把資料一筆一筆「搬」進 PST。

#### 任務三：補齊 TreeView 的數量顯示
*   在 `BuildOstFolderTree` 中，於背景執行緒讀取資料夾的 TableContext 行數並標註在節點後方。

## 下一步
為了實作上述的「屬性對拷」，我會去研讀您提供的 `libpff-main/pfftools/export_handle.c` 與 `item_file.c`，確認 `libpff` 是如何精準提取 Body、HTML、Attachments 等關鍵欄位，以確保我們在寫給 OOM 的 `MailItem` 時資料不會遺漏或亂碼。

如果您同意這個正視限制並回歸 OOM 寫入的路線，請讓我知道，我將開始撰寫程式碼。
