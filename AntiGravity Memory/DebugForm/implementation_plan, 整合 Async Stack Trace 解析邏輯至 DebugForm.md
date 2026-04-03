# 整合 Async Stack Trace 解析邏輯至 DebugForm

解決非同步函數在 `WhoCallsMe` 中無法識別的問題，並將該邏輯統整為 `DebugForm` 的靜態成員。

## User Review Required

> [!IMPORTANT]
> 此變更會修改核心 Debug 邏輯。在 Debug 模式下，因 `st.GetFrame` 不需要捕捉行號 (False)，速度會比目前的原始碼稍快。對 Release 版本效能影響為 0 (由 Conditional("DEBUG") 保護)。

## Proposed Changes

### [DebugForm]

#### [MODIFY] [DebugForm.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/DebugForm.vb)
1. 將現有的 `WhoCallsMe` 函數權限提升為 `Public Shared`。
2. 加入對 Async 狀態機的解析邏輯：
   - 偵測類別名稱是否包含 `<` 與 `>d__`。
   - 解析出原始函數名稱 (Method Name)。
   - 解析出原始類別名稱 (透過 `DeclaringType.DeclaringType`)。
3. 更新 `AddMessage3` 內部呼叫。

---

### [Form1]

#### [MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/Form1.vb)
1. 刪除局部的 `WhoCallsMe` 輔助函數。
2. 更新 `Dbg()` 輔助函數，呼叫 `DebugForm.WhoCallsMe(2)`。

## Open Questions

> [!NOTE]
> 目前決定將此邏輯放在 `DebugForm.vb` 作為 `Public Shared` 函數。這樣不需要額外的 `Module` 檔案即可讓 `Form1` 呼叫。

## Verification Plan

### Manual Verification
1. 啟動程式並切換至 Tab2。
2. 選取資料夾並等待統計完成，觀察 `DebugForm` 視窗。
3. 確認呼叫者欄位是否出現 `Form1.ComputeYearCounts` 或 `Form1.ShowYearView` 等名稱（目前為 Unknown）。
4. 檢查 `Form1` 內的其他 `Dbg` 呼叫是否依然正常運作。
