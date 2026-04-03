# 實作計畫：Outlook 資料夾展開效能極限優化 (詳細版)

## 1. 現狀目標與預期現果
- **目標**：解決當子資料夾超過 100 個時，點開 [+] 符號會產生的 1~3 秒轉圈圈現象。
- **現果**：利用 Redemption MAPITable，將展開時間優化到 100ms 以內。

## 2. 代碼清理 (修復剛才的嘗試)
針對 `Form1.vb` 目前產生的代碼殘留，首先要進行手術：
- **[MODIFY]** 移除 L812-L816 間重複的 `GetSortedSubFolders` 簽名與舊邏輯殘留。
- **[RESTORE]** 確保 `GetSortedSubFolders` 函數只有一個完整的入口。

## 3. 核心技術實作細節

### A. Redemption 軌道 (高速路徑)
1. **獲取 RDOFolder**：`_rdo.GetFolderFromID(folder.EntryID, folder.StoreID)`。
2. **存取 Hierarchy Table**：透過 `MAPITable` 取得階層數據。
3. **定義屬性標籤 (Columns)**：
   - `0x3001001F` (`PR_DISPLAY_NAME_W`) -> 獲取名稱。
   - `0x0E090102` (`PR_ENTRYID`) -> 獲取二進位 ID。
   - `0x3613001F` (`PR_CONTAINER_CLASS_W`) -> 用於識別是否為郵件。
4. **批次過濾 (Restrict)**：
   - 僅讀取 `ContainerClass = 'IPF.Note'` 或 `'IPF.Post'`。
5. **數據提取**：
   - 使用 `GetRows` 獲取 2D 陣列。
   - 使用 `_rdo.EntryIDToString(eid)` 將二進位 ID 轉為 Hex 字串，供 OOM 的 `GetFolderFromID` 使用。

### B. 安全釋放工具 (輔助函數)
實作 `TryMarshalRelease` 處理所有的 COM 釋放，避免 try-finally 寫得過於冗長。
```vb
Private Sub TryMarshalRelease(ByRef obj As Object)
    Try
        If obj IsNot Nothing AndAlso Marshal.IsComObject(obj) Then
            Marshal.ReleaseComObject(obj)
        End If
    Catch : Finally : obj = Nothing : End Try
End Sub
```

## 4. 變更文件清單

### [Component] 底層資料夾獲取 (Form1.vb)

#### [MODIFY] [GetSortedSubFolders](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/Form1.vb)
- 重寫函數，整合「Redemption Table 軌道」與「原生 OOM Fallback 軌道」。
- 確保所有 COM 讀取動作（如 `.Name`）在 A 軌道被 Table 讀取取代，在 B 軌道被快取後的 `FolderSortInfo` 取代。

#### [NEW] [TryMarshalRelease](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/Form1.vb)
- 在輔助函數區段新增此安全釋放方法。

---

## 5. 驗證與測試計畫
1. **日誌比對**：檢查 `DebugForm` 中軌道 A 與軌道 B 的耗時輸出。
2. **過濾檢查**：勾選與取消勾選「顯示所有資料夾」，驗證 Table Restrict 是否精準過濾。
3. **穩定性**：在沒有引用 Redemption 或 `_rdo Is Nothing` 的情況下，驗證程式是否能正確回退到原生模式運行。

## 6. 開發者筆記 (by AntiGravity)
> [!TIP]
> 這種優化在處理大型 Outlook 設定檔（Exchange Online 且本地緩存未全開）時尤為重要，因為每一個同步的屬性讀取都可能引發服務器往返。MAPITable 能將這些細碎請求「打包」一併解決。
