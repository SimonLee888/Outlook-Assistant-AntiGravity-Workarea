# Outlook 資料夾屬性讀取效能優化計畫

本計畫旨在透過減少 Outlook COM 屬性（`.FolderPath` 與 `.Name`）的存取次數，大幅提升掃描與統計大型資料夾樹的效能。

## 使用者評論與決策參考

> [!IMPORTANT]
> **優化核心邏輯：**
> 1. **拼接模式 (Construct)**: 在已知父路徑的情況下（如 BFS/DFS），使用 `parentPath & "\" & subFolder.Name` 得到子路徑，避開昂貴的 `.FolderPath` 呼叫。
> 2. **切分模式 (Extract)**: 在需要路徑且已有 `FolderPath` 的情況下，使用 .NET 字串運算切出最後一段作為名稱，避開次要的 `.Name` 呼叫。

> [!WARNING]
> **對 FolderBfsEntry 的改動：**
> 現有的 `FolderBfsEntry` 缺乏路徑欄位，導致後續各層級（Layer2/Layer2.5）仍須頻繁回頭讀取 COM 屬性。計畫將此欄位補上，實現一次讀取，多層受惠。

## 擬議改動內容

### 1. [基礎設施] 抽離路徑處理工具
在 `Form1_Outlook.vb` 或 `moduleStore.vb` 中新增高效能靜態工具：
- **`GetSubFolderPath(parentPath, childName)`**: 處理路徑拼接邏輯（含 Root 資料夾開頭為 `\\` 的特殊處置）。
- **`ExtractFolderName(fPath)`**: 高效從路徑末端切分出名稱。

---

### 2. [核心通訊層] Form1_Outlook.vb

#### 2.1 [MODIFY] `GetSortedSubFolders(folder)`
- 修改簽章，允許傳入選擇性的 `parentPath`。
- 內部迴圈：
    - 讀取 `subF.Name` (1 次 COM)。
    - 使用 `parentPath` 拼接出 `subFPath`。
    - **效益**: 迴圈內不再需要呼叫 `.FolderPath`。

#### 2.2 [MODIFY] `FlattenSubtreeToList` (BFS 核心)
- 將 `Queue(Of Outlook.Folder)` 提升為 `Queue(Of (Folder, String))`。
- 內部迴圈：
    - 讀取 `subF.Name`。
    - 拼接出路徑後，將存入 `result` 並入隊 `queue`。
    - **效益**: 此為系統最密集讀取點，優化後預計減少數千次 COM 呼叫。

#### 2.3 [MODIFY] `GetCachedMailCount` 等 Layer 2.5 代理層
- 檢查所有 `_dbg` 記錄，將其中的 `folder.Name` 替換為 `ExtractFolderName(fPath)`。

---

### 3. [流程協調層] Form1_MainTabs.vb

#### 3.1 [MODIFY] `FolderBfsEntry`
- **新增成員**: `Public FolderPath As String`。

#### 3.2 [MODIFY] `BuildBfsFolderTree` (Tab 1 掃描器)
- 修改 BFS 邏輯，在建立 `FolderBfsEntry` 時直接填入路徑（從父層傳遞）。
- 之後所有層級（Step 2~5）直接使用 `entry.FolderPath`，完全杜絕後續對路徑的 COM 呼叫。

---

### 4. [持久化層] Form1_SQLite2.vb

#### 4.1 [MODIFY] `RenewAttachMailListAsync`
- 移除 L606/L625 針對 `folder.Name` 的讀取，改用引數傳入的 `fPath` 進行切分。

## 開放問題

> [!NOTE]
> 1. **Root 路徑處理**: `\\StoreName` 這種開頭的路徑，拼接第一層子路徑時（如 `Inbox`）需處理為 `\\StoreName\Inbox`。我將建立單一 Helper 函式處理此語法，確保一致性。
> 2. **RDO 平行路徑**: `FlattenSubtreeToList_RDO` 是否也需要同步處理？（建議一併處理，雖然 RDO 讀取較快，但省下呼叫仍有助於極限效能）。

## 驗證計畫

### 自動化測試 (透過工具驗證)
- 使用 `_dbg` 輸出兩組實驗數據：
    - 優化前後，在相同資料夾結構（如 500 個資料夾）下的總掃描耗時。
    - 驗證 `ExtractFolderName` 處理 Root/深層路徑的正確性。

### 手動驗證
- 點選 Tab 1/Tab 2 掃描大型 PST，確認統計數字（郵件總數、層級路徑）與舊版完全一致。
- 觀察 `ProgressBar` 的跳動是否變得更順暢。
