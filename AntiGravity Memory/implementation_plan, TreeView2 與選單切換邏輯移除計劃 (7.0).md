# TreeView2 與選單切換邏輯移除計劃 (7.0)

本計劃旨在徹底清除 `Form1.vb` 中所有與 `TreeView2` 相關的代碼、事件與註解。由於 `SimTree2` 已完全取代其功能，我們將簡化相關的邏輯判斷，並移除已廢棄的「單/多選」切換邏輯。

## User Review Required

> [!WARNING]
> **Designer.vb 手動處理**：根據您的指示，我將**不觸碰** `Form1.Designer.vb`。請您在執行本計劃前或完成後，手動移除 `TreeView2` 的宣告（L1089）、初始化（L41）與控制項掛載（L274）。

## Proposed Changes

---

### [Form1.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual Studio/Visual Studio 18 (2026)/Outlook Assistant - (AntiGravity測試區)/Form1.vb)

#### [DELETE] TreeView2 事件處理程序 (L2171-2200 左右)
- 刪除 `Private Sub TreeView2_AfterSelect(...)` 函數。

#### [DELETE] 多選/單選功能切換選單 (L560 附近)
- 搜尋並移除所有與 `menuItem1` (切換多選模式) 相關的定義、初始化與事件。即使已是註解也一併清除，保持 `InitListView1` 乾淨。

#### [MODIFY] 整合顯示邏輯 (L2749 / L2450 附近)
- 將原有的二選一判斷：
  ```vb
  If TreeView2.Visible AndAlso ... Then
      ' ...
  ElseIf SimTree2.Visible Then
      ' ...
  End If
  ```
- **簡化為**：直接對 `SimTree2` 進行操作，移除對 `TreeView2.Visible` 的所有依賴檢查。

#### [DELETE] 冗餘註解塊 (L2445-2465 左右)
- 移除大段已過時的 `''` 註解（這段註解描述了舊有的切換邏輯與除錯過程）。

---

## Open Questions

- **關於 `TreeView2_AfterSelect` 的邏輯**：我確認過 `TreeView2_AfterSelect` 內的邏輯（統計當前資料夾並更新 ListView1）在 `BuildSimTree2Node` 與相關事件中已有對應實作。直接刪除該函數是否符合您的預期？

## Verification Plan

### Automated Tests
- 編譯專案：確保在移除 `TreeView2` 引用後，`Form1.vb` 通過編譯且無「未定義名稱」的錯誤。

### Manual Verification
1. 開啟程式，確認左側資料夾樹（SimTree2）運作正常。
2. 確認選單中不再出現已失效的「切換多選模式」選項。
