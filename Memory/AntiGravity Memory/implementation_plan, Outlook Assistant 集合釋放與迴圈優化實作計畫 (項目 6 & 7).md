# Outlook Assistant 集合釋放與迴圈優化實作計畫 (項目 6 & 7)

本計畫針對先前提出的第 6 點（把 Folders 集合一次取出以利釋放）與第 7 點（將迴圈內的函數呼叫提至區域變數）進行詳細的修改規劃。**在您核准之前，我不會進行任何程式碼更動。**

## 💡 原理與背景說明

### 關於第 6 點 (COM 集合釋放)
在 VB.NET 中，當我們寫出 `For Each f In folder.Folders` 時，底層其實會呼叫 COM 屬性 `.Folders` 產生一個 RCW (Runtime Callable Wrapper) 集合物件，然後取得它的列舉器。
如果我們沒有把 `.Folders` 指定給一個明確的區域變數，迴圈結束後，我們將**無法**精準地使用 `TryMarshalRelease` 來釋放這個隱形的集合物件，這在遞迴或 BFS 掃描上千個資料夾時，會導致記憶體中殘留大量未釋放的 COM 參考，進而引發 OOM (Out Of Memory) 或 COM Exception。
**解法**：先 `Dim subs As Outlook.Folders = folder.Folders`，再跑迴圈，最後在 `Finally` 區塊中 `TryMarshalRelease(subs)`。

### 關於第 7 點 (迴圈內的函數提取)
例如 `For Each subF In GetSortedSubFolders(...)`。
雖然在 VB.NET 編譯器實作中，`In` 後面的運算式只會在迴圈開始前被**計算一次**（不會每次迭代都呼叫），但將其獨立提取為 `Dim sortedSubs = GetSortedSubFolders(...)` 仍是最佳實踐：
1. 方便 Debug 觀察集合數量。
2. 提高程式碼可讀性，明確區分「資料取得」與「資料迭代」的階段。

---

## 🛠 預計修改範圍

### [MODIFY] [Form1_Outlook.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/Form1_Outlook.vb)
針對核心底層掃描邏輯補上安全的釋放機制。

*   **`GetSubtreeToListL3` (約 L1605)**:
    *   **原代碼**: `For Each subF As Outlook.Folder In current.Folder.Folders`
    *   **修改**: 提取 `Dim subs As Outlook.Folders = current.Folder.Folders` 並包裝 `Try ... Finally TryMarshalRelease(subs)`。
*   **`GetMailCountRecursiveL3` (約 L2035)**:
    *   **原代碼**: `For Each f As Outlook.Folder In rootFolder.Folders`
    *   **修改**: 同上提取與釋放。

### [MODIFY] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/Form1_MainTabs.vb)
落實第 7 點的變數提取。

*   **`BuildBfsFolderTree` (約 L714)**:
    *   **原代碼**: `For Each subFolder As Outlook.Folder In GetSortedSubFolders(curr.folderObj, fPath)`
    *   **修改**: 改為 `Dim sortedSubs = GetSortedSubFolders(curr.folderObj, fPath)`，再執行迴圈。

### [MODIFY] [Form1_OST.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/Form1_OST.vb)
OST 匯出模組中的大量 COM 迭代也存在相同風險。

*   **L378, L550, L986, L1163**: 四處 `For Each f In xxx.Folders` 提取並釋放 `Folders`。
*   **L1175**: `For Each item As Object In exportedFolder.Items`，提取並釋放 `Items` 集合。

### [MODIFY] [moduleStore.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/moduleStore.vb)
輔助模組內的 COM 迭代修正。

*   **L45, L570, L594**: 提取並釋放 `Folders`。
*   **L52**: 提取並釋放 `Items`。

---

## ✅ 驗證計畫
1.  **語法與結構檢查**：實作完成後，主動讀取修改片段，確認 `Try...Finally` 結構沒有打亂原有的迴圈邏輯。
2.  **安全性確認**：確保 `TryMarshalRelease` 傳入的確實是 COM 集合變數，而非 `List` 等 .NET 託管物件。

> [!IMPORTANT]
> **開放問題等待您的決策：**
> 以上共計 12 處修改，可以有效防止長期運作下的記憶體洩漏問題（RCW 殘留）。
> 如果您同意這個修改方向，請回覆核准，我將為您建立對應的 `task.md` 並分批開始執行修改。
