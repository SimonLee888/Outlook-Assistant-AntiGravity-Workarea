我已經透過 PowerShell 腳本掃描了專案中所有的 `.vb` 檔案（已排除自動產生的 `Designer.vb` 以及預備刪除的 `modToBeDelete.vb`）。

比對的方式是：先用正規表示式萃取出所有的模組/類別層級的變數宣告（包含 `Private`, `Public`, `Protected`, `Friend` 等），接著在全專案程式碼中（排除註解後）搜尋該變數名稱。如果該變數名稱在整個專案中**只出現過 1 次**（也就是只出現在宣告的那一行），就判定為「完全沒用到的變數」。
這樣可以用最少的 Token 達成精確的比對。

以下是掃描出來**目前在專案中完全沒有被引用到的全域/類別層級變數**清單：

### 🎨 Form1.vb (未使用的顏色常數)
* `Steel_Blue` (Line 1979)
* `Polo_Blue` (Line 1981)
* `CoralRed` (Line 1983)
* `RustRed` (Line 1985)
* `DeepAmber` (Line 1987)

### 🐛 Form1_DebugForm.vb (未使用的 Win32 API 常數)
* `LVM_GETEXTENDEDLISTVIEWSTYLE` (Line 69)

### 📁 Form1_MainTab12.vb
* `_fontHeader` (Line 11) - *看起來是原本準備用來自訂標題字型，但後來沒套用*

### 📑 Form1_MainTab345.vb
* `_currentTabIdx` (Line 14)
* `_lv4LastHoverItem` (Line 20) - *註解寫著: 預留給 ListView4 ToolTip 邏輯使用*
* `_lv5LastSortColumn` (Line 33) - *Tab5 排序用的記錄*
* `_lv5SortOrder` (Line 36) - *Tab5 排序用的記錄*

### ⚙️ Form1_Win32API.vb (未使用的 Win32 API 常數)
* `GWL_STYLE` (Line 89)
* `WS_TABSTOP` (Line 90)
* `WM_COMMAND` (Line 92)
* `BM_CLICK` (Line 95)
* `SWP_NOREDRAW` (Line 107)
* `RDW_FRAME` (Line 112)
* `TVM_SELECTITEM` (Line 115)
* `TVGN_CARET` (Line 116)
* `LVM_SETITEMCOUNT` (Line 117)

---
> **💡 建議**
> Win32 API 相關的常數和 `ThemeColors` 中的顏色雖然目前沒用到，但保留著通常無害且未來可能需要。不過 `Form1_MainTab345.vb` 和 `Form1_MainTab12.vb` 中的那幾個 `Private` 變數如果確認後續邏輯已經改寫或不再使用，可以直接安全移除。

請確認是否有您想要我幫忙清理的項目？或者是您要親自處理？