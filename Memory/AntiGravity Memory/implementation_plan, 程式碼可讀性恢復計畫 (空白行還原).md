# Outlook Assistant 程式碼可讀性恢復計畫 (空白行還原)

本計畫旨在恢復 `Form1.vb`, `DebugForm.vb`, `Form1_Main.vb`, 及 `Form1_ComL3.vb` 四個檔案的邏輯空白行，以提升程式碼的可讀性。這是在之前的自動格式化過程中可能意外壓縮了空白行後的修復動作。

## 使用者評論與要求 (User Review Required)

> [!IMPORTANT]
> **規則遵守：**
> 1. **保留註解：** 絕對不刪除任何既有的註解，特別是包含 debug 紀錄或思考過程的部分。
> 2. **AntiGravity 標記：** 我添加的註解將標記為 `by AntiGravity, 2026/04/03`。
> 3. **格式限制：** 
>    - 不使用連續兩個空白行。
>    - 變數宣告與其緊隨的迴圈/邏輯區塊保持在一起。
>    - `#Region` 與內容之間保持適當空行。
>    - 方法 (Sub/Function) 之間保持一個空行。

## 擬議變更 (Proposed Changes)

我將依序處理以下檔案，手動還原邏輯區塊之間的空行：

### [Outlook Assistant 專案]

#### [MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1.vb)
- 在 `Imports` 與 `Partial Class` 之間、`#Region` 標籤前後、以及各個 `Sub/Function` 之間插入單一空行。

#### [MODIFY] [DebugForm.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/DebugForm.vb)
- 修復 `Form1_Load` 與 `Form1_Shown` 邏輯區塊的分離，確保 UI 佈局程式碼與事件定義之間有清晰的視覺間距。

#### [MODIFY] [Form1_Main.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_Main.vb)
- 此檔案包含大量複雜邏輯 (BFS, Tab2 統計)，將重點還原演算法步驟註解上方的空行。

#### [MODIFY] [Form1_ComL3.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_ComL3.vb)
- 調整 L3, L2.5 分層定義之間的間距，讓各層級的 `Function` 定義更加突出。

## 開放性問題 (Open Questions)

> [!NOTE]
> 目前沒發現重大技術問題。我將採用「小塊寫入 (Chunked Edits)」以確保安全。

## 驗證計畫 (Verification Plan)

### 自動測試
- 我將檢視修改後的程式碼片段，確保沒有違反「連續雙空行」的規定。

### 手動驗證
- 請使用者確認程式碼視覺上是否恢復到舒適的閱讀間距。
- 確認所有原有註解 (包含 debug 演進) 都完整保留。
