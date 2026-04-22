# Outlook 資料夾屬性讀取效能優化計畫：執行進度查核報告 (更新版)

這份報告針對幾天前規劃的優化計畫進行現狀查核。

## 執行現狀查核 (2026-04-19)

### 1. 核心優化：路徑處理邏輯
| 優化項目 | 狀態 | 實作檔案與細節 |
| :--- | :---: | :--- |
| **拼接模式 (Construct)** | ✅ **已完成** | `GetSortedSubFolders` 與 `BuildBfsFolderTree` 均已支援傳入 `fPath` 並在迴圈內手動拼接子路徑。 |
| **切分模式 (Extract)** | ✅ **已完成** | 全域實作 `ExtractFolderName(fPath)`，並在多個 `_dbg` 與 UI 顯示點取代了原本的 `.Name` 呼叫。 |
| **抽離路徑處理工具** | ✅ **已完成** | 核心工具已整合進 `Form1_Win32API.vb`，提供一致的路徑處理。 |

### 2. 核心通訊層 (Form1_Outlook.vb)
| 優化項目 | 狀態 | 實作檔案與細節 |
| :--- | :---: | :--- |
| **GetSortedSubFolders 優化** | ✅ **已完成** | 簽章已改為包含路徑的參數。 |
| **GetSubtreeToList (BFS核心)** | ✅ **已完成** | 簽章已改為包含路徑的 Tuple，Queue 處理也已改為預先拼接。 |
| **Layer 2.5 代理層優化** | ✅ **已完成** | `GetMailCount` 等所有快取代理函數均已支援 `fPath` 參數。 |

### 3. 流程協調層 (Form1_MainTabs.vb)
| 優化項目 | 狀態 | 實作檔案與細節 |
| :--- | :---: | :--- |
| **FolderBfsEntry 改動** | ✅ **已完成** | 已新增 `FolderPath` 成員，並在 `ComputeFolderStatsAsync` 流程中全面使用。 |

---

## 使用者疑問詳解

### Q1: `RenewAttachMailList` 的「微量優化空間」在哪裡？
- **細節**：在 `RenewAttachMailListAsync` (Form1_SQLite2.vb: L628) 中，呼叫 `GetAttachMailListL3(folder, Nothing)` 時，目前尚未傳遞 `fPath`。這導致 `GetAttachMailListL3` 內部為了顯示 Debug Log，會再次呼叫一次 `folder.Name` (Form1_Outlook.vb: L1864)。
- **改進點**：將 `GetAttachMailListL3` 也加入 `Optional fPath As String` 參數，利用現成的 `fPath` 進行 `ExtractFolderName`，省下這一次的 COM 呼叫。

### Q2: `RenewCacheAsync` Phase 3 的「微量重複處理」在哪裡？
- **細節**：目前的 Phase 3 (Form1_SQLite2.vb: L521-L537) 針對每個 Dirty 資料夾會連續呼叫 4~5 個 L3 函數。雖然每個函數都傳入了 `fPath`，但在 L3 內部為了確保穩定性，每個函數都各自呼叫了自己的 `folder.PropertyAccessor` 或 `EntryID/StoreID`。
- **改進點**：在 Phase 3 迴圈開頭，一次性讀取該資料夾的 `(EntryID, StoreID, Snap, Name)` 封裝成一個 `FolderBasicInfo` 結構，並傳遞給後續所有 L3 函數，達成「單次 COM 讀取，多次邏輯使用」的目標。

### Q3: `GetSubtreeToListL3_Rdo` (Redemption 路徑) 如何優化？
- **優化後的實作細節**：
  ```vb
  ' 原版：
  For Each subFolder As Redemption.RDOFolder In current.Folders
      resultBag.Add(subFolder) ' 只存物件，下次要路徑時又要 call COM
  Next

  ' 優化計畫：
  For Each subFolder As Redemption.RDOFolder In current.Folders
      Dim childPath As String = current.Path & "\" & subFolder.Name ' 遍歷時就順手拼接
      resultBag.Add((subFolder, childPath)) ' 存入 Tuple
  Next
  ```
- **效益**：這能讓 RDO 掃描器在面對大型 PST (上萬資料夾) 時，完全杜絕後續對 `.FolderPath` 的 COM Round-trip。

## 結論
依據您的指示，**「死碼（Dead Code）」將予以保留不刪除**。我將會針對上述三項「微量優化空間」進行精鍊更新。
