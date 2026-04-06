# NumericUpDown 動態增幅邏輯重構計畫

此計畫旨在優化附件搜尋功能中的大小過濾控制項 (`NumberMin`, `NumberMax`)。我們將實作動態的 `Increment`（單次點擊增量）以及 `Accelerations`（長按加速）邏輯，讓 KB 單位下的微調更符合使用者習慣。

## 使用者評論與決策 (User Review Required)

> [!IMPORTANT]
> **動態增幅規則應用範圍**：
> 計畫將此邏輯同時套用於 `NumberMin` 與 `NumberMax`，以維持操作邏輯的一致性。
>
> **長按（Acceleration）數值建議**：
> 標準 `NumericUpDown` 在長按時會自動套用預設加速。我建議在長按超過 2 秒後，將增幅額外提升 5~10 倍（視當前單位而定），這部分已包含在配置清單中。

## 擬議變更

### 表單與邏輯控制項 [Component]

#### [MODIFY] [Form1.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1.vb)

1.  **移除現有的簡易邏輯** (L491-L492)：
    移除原本簡單的 `If(NumberMin.Value < 100, 1, 10)` 匿名函式。

2.  **實作核心邏輯方法 `UpdateNumericIncrement`**：
    建立一個副程式，傳入 NumericUpDown 及其對應的 ComboBox 單位，執行以下判斷：
    - 若單位為 `MB` 或 `GB`：`Increment = 1`。
    - 若單位為 `KB`：
        - `Value <= 20`: `1`
        - `Value <= 50`: `5`
        - `Value <= 200`: `10`
        - 否則 (`> 200`): `50`

3.  **掛載事件點 (Event Hub)**：
    - 在 `InitTab3UI` 中，將 `NumberMin.ValueChanged` 與 `UnitMin.SelectedIndexChanged` 均導向此更新方法。
    - 同樣處理 `NumberMax` 部分。

4.  **設定加速 (Accelerations)**：
    在初始化時為 `NumberMin` / `NumberMax` 加入加速配置：
    - 2 秒後：增量變為當前 `Increment` 的 5 倍。
    - 5 秒後：增量變為 50 (或更高)。

## 開放性問題 (Open Questions)

- **長按數值調整**：目前的加速方案是：長按 2 秒後進入下一階加速。您是否有特定的長按增量比例需求？或是使用我們建議的「當前增量 x 5」？

## 驗證計畫

### 手動測試 (Manual Verification)
1.  切換單位至 **KB**，手動由 1 按到 250，觀察增幅是否在 21、51、201 時發生變化。
2.  切換單位至 **MB**，觀察增幅是否恢復為固定 1。
3.  測試**長按住**上下箭頭，觀察跳動速度是否在 2 秒後明顯變快。
