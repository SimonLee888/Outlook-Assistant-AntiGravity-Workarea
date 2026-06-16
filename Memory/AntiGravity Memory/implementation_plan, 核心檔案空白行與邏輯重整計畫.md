# 核心檔案空白行與邏輯重整計畫

這是一個針對 `Form1.vb`、`DebugForm.vb`、`Form1_Main.vb` 與 `Form1_ComL3.vb` 四個檔案的格式美化與邏輯區塊化任務。由於這些檔案曾被其他程式壓縮掉空行，目前的目標是恢復其可讀性，並確保「開發思路」被清楚保留。

## 核心重整原則

1.  **邏輯斷點留白**：
    *   在每個 `Sub` / `Function` 的 `End Sub` 下方補足 1 行空行（**但若長度只有 1 行指令的短函數則不留空行**）。
    *   **不要變動** `#Region` 與 `#End Region` 的位置，使用者已手動調整完畢。
    *   在關鍵邏輯轉折處插入空行。
    *   **變數宣告與其專屬迴圈緊貼**，中間不留空行。
2.  **導覽型註解強化**：
    *   在描述「思路、歷程、避坑、優化」的註解塊上方插入空行，使其具備視覺導航功能。
    *   若註解是緊貼描述下一行技術操作，則不留空行。
3.  **註解內容保留**：
    *   **不要進行日期紀錄的整合或結構化**（如 `lvwDebug_ItemSelectionChanged` 部分），維持原樣或由使用者自行調整。
    *   對齊每個函數開頭的說明註解，確保「Fallback 鏈」等邏輯說明清晰。
4.  **安全執行**：
    *   僅操作空白行與註解對齊，**絕對不更動**任何實體程式碼邏輯。
    *   **不需要**加上作者標記（如 ' 重整 by AntiGravity）。

---

## 預計改動檔案

### [MODIFY] [DebugForm.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/DebugForm.vb)
*   **改動細節**：
    *   重整事件處理器內的空行，讓 Guard Clauses 與後續邏輯區隔。

### [MODIFY] [Form1_ComL3.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/Form1_ComL3.vb)
*   **改動細節**：
    *   對齊每個函數開頭的說明註解，確保「Fallback 鏈」等邏輯說明清晰。

### [MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/Form1.vb)
*   **重點區域**：`InitTabUI` 系列函數、`Form_Load`、`Form_Shown`。
*   **改動細節**：
    *   重整動態配置 UI 的程式碼塊，按頁面或按功能分類並補上空行。

### [MODIFY] [Form1_Main.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80)/Form1_Main.vb)
*   **重點區域**：主邏輯區、主要事件處理。
*   **改動細節**：由於此檔案最為龐大，將採取最嚴謹的分段重整。

---

## 執行流程與驗證

1.  **第一步**：先對 `DebugForm.vb` 進行示範性重整（包含您剛才提到的配對邏輯），這將作為基準風格。
2.  **第二步**：獲得確認後，依序處理其餘三個檔案。
3.  **驗證**：確保檔案仍能正確編譯（雖然只是改空行）。

---

## 開放問題 (已確認)
1. **註解標記**：本次重整不加入額外標記。
2. **處理區域**：由 AntiGravity 分小塊嚴謹執行，確保不損壞程式碼。
