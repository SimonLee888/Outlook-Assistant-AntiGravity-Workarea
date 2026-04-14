# 附件讀取邏輯優化 Walkthrough

針對 Debug 視窗中出現的大量「無法對這種類型的附件執行作業」錯誤（未知類型作業），我已在 Layer 3 讀取端加入了類型篩選機制。

## 變更內容

### [Form1_Outlook.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_Outlook.vb)

#### 1. `GetAttachFilename` 函數優化
在讀取 `FileName` 屬性前，先檢查附件類型是否為 `olByValue`。這可以跳過 OLE 分送對象或內嵌項目，防止 Outlook 拋出錯誤。
- **RDO 路徑**：檢查 `att.Type = 1`。
- **OOM 路徑**：檢查 `att.Type = Outlook.OlAttachmentType.olByValue`。

#### 2. RDO 預載函數同步更新
針對 `PreloadAttachByRDOAsync1` 與 `PreloadAttachByRDOAsync2` 同步加入相同的篩選邏輯，確保併行預載時也能正確處理不同類型的附件。

---

## 驗證結果
- **穩定性提升**：不再因為讀取特殊附件（如嵌入圖表、郵件對象）而觸發 `TargetInvocationException`。
- **Debug 視窗改善**：大幅減少 `OOM 失敗` 的 Log 噪音。

<ctrl94> [!NOTE]
> 此修改僅針對「附件檔名讀取」部分。關於你提到的「月份統計快取」部分，如你所言暫不更動。
