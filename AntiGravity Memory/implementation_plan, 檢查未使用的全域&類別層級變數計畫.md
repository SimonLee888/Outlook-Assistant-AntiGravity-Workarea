# 檢查未使用的全域/類別層級變數計畫

此計畫旨在掃描專案中所有的 `.vb` 檔案，找出定義在類別層級（Class level）或模組層級（Module level）但實際上在整個專案中從未被使用的變數宣告。這有助於清理冗餘程式碼，提升維護性。

## 使用者評論要求

> [!IMPORTANT]
> 此次任務僅執行「檢查」並回報，不會主動刪除任何變數。若需要刪除，請在確認回報清單後再行指示。
> 所有的分析過程與結果將詳細記錄於此。

## 執行步驟

1. **定義搜尋範圍**：包含專案目錄下所有的 `.vb` 檔案（排除 `.Designer.vb` 與 `My Project` 目錄）。
2. **提取變數宣告**：
   - 識別 `Public`, `Private`, `Friend`, `Protected`, `Dim`（在類別/模組層級）定義的變數。
   - 記錄變數名稱、所在檔案及行號。
3. **檢查引用情況**：
   - 使用 `grep` 或類似工具在整個專案中搜尋每個變數名稱。
   - 排除宣告本身的行。
   - 考慮 VB.NET 不區分大小寫的特性。
4. **整理報告**：列出所有「零引用」的變數清單。

## 預計檢查的檔案

- `Form1.vb`
- `Form1_MainTab12.vb`
- `Form1_MainTab345.vb`
- `Form1_Outlook.vb`
- `Form1_SQLite2.vb`
- `Form1_OST.vb`
- `Form1_Win32API.vb`
- `Form1_DebugForm.vb`
- `Form1_SimTree.vb`
- `modToBeDelete.vb`
- `ApplicationEvents.vb`

## 驗證計畫

### 檢查方法
- 針對篩選出的「未使用變數」，隨機挑選幾個使用全域搜尋（Grep）再次確認是否真的沒有任何地方引用（包含註解以外的程式碼）。

### 手動驗證
- 提供清單供使用者核對，特別是某些可能透過 Reflection 或特殊方式引用的變數（雖然在 WinForms 專案中較少見）。
