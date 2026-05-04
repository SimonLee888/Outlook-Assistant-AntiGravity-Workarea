# Lv3 / Lv4 / Lv5 欄位寬度影響點全覽

> 分析目標：找出所有**會改變或影響** ListView3、ListView4、ListView5 欄位寬度的程式碼位置。
> 分析時間：2026-05-06 by Claude Sonnet 4.6

---

## 欄位結構對照表

| 欄位索引 | Lv3 (Tab3 搜尋) | Lv4 (Tab4 系列) | Lv5 (Tab5 重複) |
|---------|----------------|----------------|----------------|
| Columns(0) | 主旨 | 主旨 | 主旨 |
| Columns(1) | 郵件大小 | 郵件大小 | 郵件大小 |
| Columns(2) | **收到日期** | **收到日期** | **收到日期** |
| Columns(3) | 寄件者 | 寄件者 | 寄件者 |
| Columns(4) | 附件個數 | 相似 | 群組 |
| Columns(5) | EntryID | EntryID | 相似 |
| Columns(6) | — | — | EntryID |

---

## 影響寬度的程式碼位置（兩大來源）

### 來源 A：初始化時寫死比例（Form1.vb）

#### A1 — Lv3 初始欄位建立
**檔案：** [Form1.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-(AntiGravity測試區)/Form1.vb#L546-L554) — **Line 546~554**（`InitTab3UI` 函數內）

```vb
' ── ListView3: 搜尋結果欄位定義 ──
With ListView3
    .Columns.Clear()
    .Columns.Add("主旨",    "主旨",    CInt(ListView3.Width * 0.45))
    .Columns.Add("郵件大小","郵件大小",CInt(ListView3.Width * 0.13)) : .Columns(1).TextAlign = Right
    .Columns.Add("收到日期","收到日期",CInt(ListView3.Width * 0.18)) : .Columns(2).TextAlign = Center  ← 收到日期
    .Columns.Add("寄件者",  "寄件者",  CInt(ListView3.Width * 0.22))
    .OwnerDraw = True
End With
```

> [!WARNING]
> 注意：Lv3 初始化**沒有加入 `附件個數` 和 `EntryID` 兩欄**（只有4欄），但 `AutoResizeLvColumns` 在 `lv.Columns.Count >= 6` 才執行 Lv3 邏輯，所以 Resize 時的比例公式實際上**永遠不會被觸發**，除非在 `InitTab3UI` 之後另外加欄。請確認 Lv3 的附件欄是否在別的地方動態加入。

---

#### A2 — Lv4 初始欄位建立
**檔案：** [Form1.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-(AntiGravity測試區)/Form1.vb#L626-L637) — **Line 626~637**（`InitTab4UI` 函數內）

```vb
With ListView4
    .Columns.Clear()
    Dim lv4Names As String() = {"主旨","郵件大小","收到日期","寄件者","相似","EntryID"}
    For Each n In lv4Names : .Columns.Add(n, n) : Next
    .Columns("主旨").Width    = .Width * 0.4
    .Columns("郵件大小").Width = CInt(.Width * 0.13)  : .Columns("郵件大小").TextAlign = Right
    .Columns("收到日期").Width = CInt(.Width * 0.18)  : .Columns("收到日期").TextAlign = Center  ← 收到日期 = 18%
    .Columns("寄件者").Width   = .Width * 0.18
    .Columns("相似").Width     = .Width * 0.08  : .Columns("相似").TextAlign = Center
    .Columns("EntryID").Width  = 0
End With
```

---

#### A3 — Lv5 初始欄位建立
**檔案：** [Form1.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-(AntiGravity測試區)/Form1.vb#L723-L735) — **Line 723~735**（`InitTab5UI` 函數內）

```vb
With ListView5
    .Columns.Clear()
    Dim lv5Names As String() = {"主旨","郵件大小","收到日期","寄件者","群組","相似","EntryID"}
    For Each n In lv5Names : .Columns.Add(n, n) : Next
    .Columns("主旨").Width    = CInt(.Width * 0.34)
    .Columns("郵件大小").Width = CInt(.Width * 0.12)  : .Columns("郵件大小").TextAlign = Right
    .Columns("收到日期").Width = CInt(.Width * 0.17)  : .Columns("收到日期").TextAlign = Center  ← 收到日期 = 17%
    .Columns("寄件者").Width   = .Width * 0.17
    .Columns("群組").Width     = CInt(.Width * 0.08)  : .Columns("群組").TextAlign = Right
    .Columns("相似").Width     = CInt(.Width * 0.08)  : .Columns("相似").TextAlign = Center
    .Columns("EntryID").Width  = 0
End With
```

---

### 來源 B：視窗 Resize 時動態調整（Form1.vb）

**函數：** `AutoResizeLvColumns(lv As ListView)`
**觸發時機：** `HandleLvResize` 事件（每次 ListView 大小改變）以及 `CheckAttCount.CheckedChanged`（Lv3 附件欄切換）

**檔案：** [Form1.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-(AntiGravity測試區)/Form1.vb#L1735-L1806) — **Line 1735~1806**

#### B1 — Lv3 Resize 比例（Line 1762~1770）

```vb
ElseIf lv Is ListView3 Then  ' Columns.Count >= 6 才執行
    lv.Columns(1).Width = CInt(w * 0.15)    ' 郵件大小
    lv.Columns(2).Width = CInt(w * 0.2)     ' 收到日期  ← 20%
    lv.Columns(3).Width = CInt(w * 0.2)     ' 寄件者
    lv.Columns(5).Width = CInt(w * 0.01)    ' EntryID
    lv.Columns(4).Width = If(CheckAttCount.Checked, CInt(w * 0.1), 0.03)  ' 附件個數
    lv.Columns(0).Width = w - (sum of 1~5) - 5  ' 主旨 吸收剩餘
```

#### B2 — Lv4 Resize 比例（Line 1772~1780）

```vb
ElseIf lv Is ListView4 Then  ' Columns.Count >= 5 才執行
    lv.Columns(1).Width = CInt(w * 0.13)    ' 大小
    lv.Columns(2).Width = CInt(w * 0.18)    ' 收到時間  ← 18%
    lv.Columns(3).Width = CInt(w * 0.18)    ' 寄件者
    lv.Columns(4).Width = CInt(w * 0.08)    ' 相似度
    lv.Columns(5).Width = CInt(w * 0.01)    ' EntryID
    lv.Columns(0).Width = w - (sum of 1~5) - 5  ' 主旨 吸收剩餘
```

#### B3 — Lv5 Resize 比例（Line 1781~1790）

```vb
ElseIf lv Is ListView5 Then  ' Columns.Count >= 7 才執行
    lv.Columns(1).Width = CInt(w * 0.12)    ' 郵件大小
    lv.Columns(2).Width = CInt(w * 0.17)    ' 收到日期  ← 17%
    lv.Columns(3).Width = CInt(w * 0.17)    ' 寄件者
    lv.Columns(4).Width = CInt(w * 0.05)    ' 群組
    lv.Columns(5).Width = CInt(w * 0.08)    ' 相似
    lv.Columns(6).Width = CInt(w * 0.01)    ' EntryID
    lv.Columns(0).Width = w - (sum of 1~6) - 5  ' 主旨 吸收剩餘
```

---

## 收到日期欄寬度對比摘要

| | 初始化寬度 | Resize 後寬度 | 一致性 |
|---|---|---|---|
| **Lv3** | `0.18` (18%) | `0.2` (20%) | ❌ 不一致（初始 vs Resize） |
| **Lv4** | `0.18` (18%) | `0.18` (18%) | ✅ 一致 |
| **Lv5** | `0.17` (17%) | `0.17` (17%) | ✅ 一致（但與 Lv4 不同） |

> [!IMPORTANT]
> **Lv3 的 `收到日期` 欄存在初始化(18%)與 Resize(20%) 不一致的問題。**
> 只要使用者調整一次視窗大小，Lv3 的日期欄就會從 18% 跳到 20%。

---

## 如要手動調整的建議操作位置

若要統一三個 ListView 的「收到日期」欄寬度，需同時修改以下 **6 個數值**：

| 修改對象 | 檔案 | 行號 | 目前值 |
|---------|------|------|--------|
| Lv3 初始化 | Form1.vb | L551 | `0.18` |
| Lv4 初始化 | Form1.vb | L633 | `0.18` |
| Lv5 初始化 | Form1.vb | L730 | `0.17` |
| Lv3 Resize | Form1.vb | L1765 | `0.2`  |
| Lv4 Resize | Form1.vb | L1775 | `0.18` |
| Lv5 Resize | Form1.vb | L1784 | `0.17` |

