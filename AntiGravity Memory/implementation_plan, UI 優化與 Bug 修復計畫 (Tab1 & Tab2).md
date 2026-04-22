# UI 優化與 Bug 修復計畫 (Tab1 & Tab2)

針對使用者提出的 Tab2 統計視圖同步與收合問題，以及 Tab1 分群顯示與資料夾大小統計同步問題，進行以下重構與修復。

## 使用者評論與決策

- **Tab2 問題 1 (年度點選無效)**: 目前點選年度群組標題時，僅內部切換圖表，但 ListView2 維持展開狀態，與使用者預期「點年度看年度趨勢」不符。
- **Tab2 問題 2 (分組無法收合)**: ListView 群組預設不可收合，需加入 `LVM_SETGROUPINFO` 相關 Win32 API 調用以啟用收合功能。
- **Tab1 問題 1 (標題重複)**: 當選取單一 PST 根目錄時，「合計欄」與「PST 分組行」內容高度重複。計畫在單一分組且為根目錄時隱藏合計列。
- **Tab1 問題 2 (節點標題)**: 若選取的 node 不是 PST 本體而是子資料夾，目前的「分組名稱」讀取 `.Store.DisplayName` 可能不夠直觀。
- **Tab1 問題 3 (大小統計同步)**: 更新 FolderSize 後，`ListViewGroup` 的標題文字（含合計大小）沒有重新計算並更新。
- **Tab1 問題 4 (大小單位)**: 根據使用者要求，將 Tab1 的資料夾大小顯示單位從 KB 改為 MB。

---

## 擬議變更

### 1. Tab2 統計視圖與分組功能強化

#### [MODIFY] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_MainTabs.vb)

- **`RenderLvYearView`**: 
    - 建立群組時，加入 Win32 `LVM_SETGROUPINFO` 設定 `LVGS_COLLAPSIBLE` 狀態。
    - 調整群組標題，使其更簡潔。
- **`ListView2_MouseDown`**: 
    - 強化群組標題點擊偵測。當使用者點擊群組標題時，除了切換圖表，應切換圖表至該年度的「月份分佈」。
- **`ListView2_SelectedIndexChanged`**:
    - 修正連動邏輯。當選中月份時，圖表切換至該年月份分佈並高亮；當選中處於收合狀態的年份代表項（或標題）時，切換至該年月份。

#### [MODIFY] [Form1_Win32API.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_Win32API.vb)

- 加入 `LVGROUP` 結構、`LVM_SETGROUPINFO` 定義，以及 `LVGS_COLLAPSIBLE` 等常數，用以實現 ListView 群組收合功能。

---

### 2. Tab1 分群顯示與統計同步優化

#### [MODIFY] [Form1_MainTabs.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_MainTabs.vb)

- **`SimTree1_AfterSelect`**:
    - **分組邏輯優化**: 如果選算的 node 其父層不是 Store 根節點，分組名稱應包含父路徑或自訂標籤，避免所有非根目錄選取都顯示同一個 PST 名稱導致困惑。
    - **合計列隱藏策略**: 若偵測到只有一個分組且該分組代表一個完整的 Store 根節點，則隱藏「合計列」，避免重疊。
- **`ComputeFolderSize`**:
    - **單位調整**: 將計算結果轉化為 MB 顯示。
    - 在每個資料夾計算完畢後，即時累加該資料夾所屬群組的總大小。
    - 計算結束後，呼叫新設的 `UpdateTab1GroupHeaders()` 重新渲染 `ListViewGroup.Header` 文字，確保「分組行」顯示的大小與項目同步更新。
- **`BuildListViewItem_Tab1`**:
    - 修改顯示邏輯，將 KB 改為 MB (若有快取資料)。

---

## 驗證計畫

### 自動測試
- 無（主要涉及 WinForms UI 元件與 Win32 API 交互）。

### 手動驗證
1. **Tab2 驗證**:
    - 點選年度群組標題，確認 Chart2 是否正確切換為該年的 1-12 月分佈。
    - 確認年度群組是否出現「展開/收合」圖示，且點擊後可正常動作。
2. **Tab1 驗證**:
    - 選取單一 PST 根節點，確認合計列是否消失。
    - 選取非根節點（如 Inbox 下的子資料夾），確認分組名稱是否正確辨識。
    - 執行「統計資料夾大小」，觀察分組行括號內的 KB/MB 數值是否隨著子項目的計算完成而動態增加。
