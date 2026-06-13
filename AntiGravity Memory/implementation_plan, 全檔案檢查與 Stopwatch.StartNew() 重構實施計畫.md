# 全檔案檢查與 Stopwatch.StartNew() 重構實施計畫

本計畫在檢查全專案所有 `.vb` 檔案後，確認僅 `Form1_SQLite2.vb` 尚有 3 處執行中的區域變數需要重構。其餘檔案中的 Stopwatch 宣告要麼已在先前完成修改，要麼屬於類別欄位（Class Fields）或已註解的舊程式碼，不適合或不需要修改。

## 全專案檢查結果說明

1. **已修改完成的檔案**：
   - [Form1.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1.vb)
   - [Form1_MainTab12.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_MainTab12.vb)
   - [Form1_MainTab345.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_MainTab345.vb)
   - [Form1_Outlook.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_Outlook.vb)

2. **不需修改的類別欄位 (Class Fields)**：
   - `Form1_DebugForm.vb` L76: `Private sw0, sw1, sw2, sw3, sw4, sw5, sw6 As New Stopwatch` (多個類別變數宣告，不適合直接改為單一 StartNew)
   - `Form1_OST.vb` L42: `Private _tab7StatusSw As New Stopwatch()` (類別狀態變數，非區域變數)

3. **不需修改的已註解程式碼與待刪除模組**：
   - `Form1_Win32API.vb` (僅存註解掉的 Stopwatch 程式碼)
   - `modToBeDelete.vb` (待刪除模組，皆為註解代碼)

4. **待修改的檔案**：
   - [Form1_SQLite2.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_SQLite2.vb) (共 3 處)

---

## 預計變更

### Form1_SQLite2.vb

#### [MODIFY] [Form1_SQLite2.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/Form1_SQLite2.vb)

##### 修改點 1 (L375)
將：
```vb
        Dim sw As New Diagnostics.Stopwatch : sw.Start()
```
改為：
```vb
        Dim sw As Diagnostics.Stopwatch = Diagnostics.Stopwatch.StartNew()  ' by Gemini 3.5 Flash, 2026/06/07
```

##### 修改點 2 (L442)
將：
```vb
        Dim sw As New Diagnostics.Stopwatch : sw.Start()
```
改為：
```vb
        Dim sw As Diagnostics.Stopwatch = Diagnostics.Stopwatch.StartNew()  ' by Gemini 3.5 Flash, 2026/06/07
```

##### 修改點 3 (L555)
將：
```vb
            Dim swThrottle As New Stopwatch : swThrottle.Start()
```
改為：
```vb
            Dim swThrottle As Stopwatch = Stopwatch.StartNew()  ' by Gemini 3.5 Flash, 2026/06/07
```

---

## 驗證計畫

### 手動驗證
1. 確保專案建置成功，無語法錯誤。
2. 使用 `view_file` 複檢修改點，確保程式碼邏輯對齊，無多餘殘留。
