# 解決 Form1 Partial Class 資源檔衝突問題

在 VB.NET 中，當一個大型的 Form 被拆分為多個 `Partial Class` 檔案時，如果有多個檔案同時擁有 `.resx` 資源檔，編譯器會嘗試為每個資源檔生成相同的輸出名稱（例如 `Outlook_Assistant.Form1.resources`），從而導致「兩個輸出檔名解析成相同的輸出路徑」的錯誤。

## 使用者待確認事項

- [ ] `Form1_ComL3.resx` 檔案中目前沒有任何實際資源（僅有標頭），是否確認可以刪除？
- [ ] 拆分後的 `Form1_Main.vb` 是否有在 Visual Studio 中被誤認為是一個新的 Form？

## 提出的修改方案

### 1. 資源檔清理

#### [DELETE] [Form1_ComL3.resx](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_ComL3.resx)
這個檔案目前是空的（5.8KB 僅包含 XML 標頭），且因為它與 `Partial Class Form1` 關聯，會造成輸出衝突。刪除它後，所有的資源將統一由 `Form1.resx` 管理。

### 2. 專案檔優化 (若刪除資源檔後仍報錯)

#### [MODIFY] [Outlook Assistant.vbproj](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Outlook%20Assistant.vbproj)
明確告知編譯器 `Form1_Main.vb` 和 `Form1_ComL3.vb` 是程式碼檔案，而非獨立的 Form，避免 SDK 自動對應資源。

```xml
<ItemGroup>
  <None Remove="Form1_ComL3.resx" />
  <Compile Update="Form1_Main.vb">
    <SubType>Code</SubType>
  </Compile>
  <Compile Update="Form1_ComL3.vb">
    <SubType>Code</SubType>
  </Compile>
</ItemGroup>
```

## 驗證計畫

### 手動驗證
1. 刪除 `Form1_ComL3.resx`。
2. 進行「清理方案」(Clean Solution) 並「重新建置」(Rebuild)。
3. 確認錯誤訊息消失。
4. 若仍有衝突，檢查 `Form1_Main.vb` 是否也被建立了隱藏的 `.resx` 檔案。
