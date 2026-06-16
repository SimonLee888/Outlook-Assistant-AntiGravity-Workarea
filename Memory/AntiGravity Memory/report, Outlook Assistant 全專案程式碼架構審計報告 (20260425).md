# Outlook Assistant 全專案程式碼架構審計報告 (2026/04/25)

本報告針對現有四個主要模組（Form1, Outlook, MainTabs, OST）的組織結構進行全面稽核，旨在優化 Layer 1~3 的職責邊界。

## 1. 審計發現與優化建議

### A. Form1_Outlook.vb (邏輯核心)
*   **現況**：目前將 Layer 2.5 (Cache Proxy) 與 Layer 3 (COM Data) 混合在一個名為 `■ 10 底層 COM 函數群 (新設計，現役主力)` 的巨大 Region 中。
*   **問題**：
    *   Region 名稱過於口語化（"現役主力"），無法體現分層職責。
    *   `GetUniqueFolderList` 本身具有流程性質（Layer 2），卻夾在 L2.5 函數之間。
    *   `FillFolderCacheFromDbRow` 是 L2.5 的核心邏輯，卻被放在檔案末尾的「輔助函數」區塊。
*   **優化方案**：建議仿照 `Form1_SQLite2.vb` 的模式，明確區分為 `Layer 2.5 快取代理層` 與 `Layer 3 底層數據層`。

### B. Form1_MainTabs.vb (UI 事件與分頁邏輯)
*   **現況**：包含了 Tab1 的核心計算函數 `ComputeFolderStatsAsync` (Layer 2)。
*   **問題**：
    *   `ComputeFolderStatsAsync` 是全專案最複雜的邏輯之一，放在 `MainTabs` 中使其顯得異常龐大。
    *   雖然它與 Tab1 強相關，但其「層層遞進、彙總運算」的性質更接近於「邏輯處理」而非「UI 事件」。
*   **優化方案**：考慮將其移至 `Form1_Outlook.vb` 的新區域 `Layer 2: 流程協調層`，或維持現狀但建立更清晰的 Region 標記。

### C. Form1_OST.vb (Tab7 專用模組)
*   **現況**：結構相對獨立（L1, L2, L3 都有），但部分函數與 `Form1_Outlook.vb` 有功能上的「平行感」。
*   **問題**：
    *   `ShowOstItems` 與 `ShowPstItems` 的 UI 渲染邏輯散落在 L1 事件中。
*   **優化方案**：強化其內部的 Layer 分層名稱，與主專案對齊。

### D. Region 名稱與內容不符
*   **Form1.vb**：`■ 02 Form 生命週期 & 外觀初始化` 內部的 `OkayNowYouHaveToken` 與 `ThrottledYieldAsync` 是**全域通用的並發工具函數**，不應僅歸類於「生命週期」。
*   **Form1_Outlook.vb**：`■ 10 底層 COM 函數群` 的層次感不足，建議改為 `■ 20 Layer 2.5: 快取代理層` 與 `■ 30 Layer 3: 底層數據層`。

---

## 2. 函數定位稽核表 (Misplacement Check)

| 函數名稱 | 目前位置 | 建議調整方向 | 原因 |
| :--- | :--- | :--- | :--- |
| `ComputeFolderStatsAsync` | MainTabs.vb | 建議移至 Outlook.vb 或標記為 L2 | 這是純邏輯運算，非單純 UI 事件。 |
| `GetUniqueFolderList` | Outlook.vb | 標記為 Layer 2 | 涉及跨資料夾遍歷與去重，屬於流程協調。 |
| `FillFolderCacheFromDbRow` | Outlook.vb (尾端) | 移至 Layer 2.5 區域 | 這是快取機制的關鍵內部函數。 |
| `OkayNowYouHaveToken` | Form1.vb | 移至獨立 Region：全域並發工具 | 這是非同步任務的核心設施。 |
| `ThrottledYieldAsync` | Form1.vb | 移至獨立 Region：全域並發工具 | 這是非同步任務的核心設施。 |

---

## 3. 命名合適性建議 (Naming Review)

1.  **`GetMailCountL3` / `GetMailCount` (L2.5)**: 命名非常清晰，符合現在的架構。
2.  **`SummarizeSubTreeBottomUp`**: 命名極佳，準確描述了演算法特徵。
3.  **`RenewCache`**: 在 SQLite2.vb 中，此名稱可能需要更具體（例如 `ClearFolderStatsCache`），因為它目前只處理統計數字的失效。

---

## 下一步行動建議 (Next Steps)

如果您同意上述分析，我建議：
1.  **第一階段**：重新劃分 `Form1_Outlook.vb` 的 Region，將 L2.5 與 L3 徹底切開。
2.  **第二階段**：整理 `Form1.vb`，將並發工具（Token, Yield）獨立出來。
3.  **第三階段**：優化 `Form1_MainTabs.vb`，強化 L1 與 L2 的界線。

您想先針對哪一部分開始動手？
