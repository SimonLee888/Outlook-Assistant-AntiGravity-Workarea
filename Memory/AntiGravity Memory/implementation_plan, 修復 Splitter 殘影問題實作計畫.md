# 修復 Splitter 殘影問題實作計畫

當 SplitContainer 收合（縮小至 10px）或開啟（恢復原寬度）時，右側 Panel2 內的 OwnerDraw ListView 可能會留下舊畫面的雜訊。這是因為凍結重繪期間區域變動，解鎖後 Windows 未能完全清除舊有的繪製內容。

## 使用者評論與需求
- 雜訊殘留在 Splitter 開啟後的位置。
- 懷疑 Windows 物件管理或 OwnerDraw 繪製未正確重繪。
- 維持原有程式碼註解。

## 擬定修改
### Form1.vb
在 `SplitterToggle` 方法中：
1. 在修改 `SplitterDistance` **之前**，對 `sc.Panel2` 執行一次 `Invalidate`。
2. 在修改 **之後**，將 `sc.Panel2.Invalidate(True)` 提升為更強力的重繪組合，確保背景被擦除。

#### [MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1.vb)
- 優化 `SplitterToggle` 的重繪流程。

## 驗證計畫
### 手動驗證
- 雙擊 Tab5 的 Splitter 邊緣進行收合。
- 雙擊 10px 邊緣恢復寬度。
- 確認中間原本 ListView 的內容是否被乾淨清除。
