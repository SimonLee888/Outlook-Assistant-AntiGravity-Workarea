# Tab3 TreeView 替換為多選 SimTree 實施計劃

此計劃旨在將 Tab3 (附件搜尋) 的原生 `TreeView3` 控制項替換為自訂的 `SimTree` (支援多選、Ctrl/Shift 鍵)。這將提升 Tab3 的操作彈性，使其與 Tab2 的多選行為一致。

## 使用者評論與回饋要求

> [!IMPORTANT]
> 1. **資料夾去重**: 多選模式下，使用者可能選中父子層級資料夾。搜尋邏輯必須進行去重，避免重複掃描。
> 2. **效能考量**: 多選會顯著增加掃描量，Phase 2 (附件明細) 的快取預熱顯得更為重要。

## 擬議變更

### [UI 控制項層]

#### [MODIFY] [Form1.Designer.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1.Designer.vb)
*   手動將 `TreeView3` 的型別定義由 `System.Windows.Forms.TreeView` 改為 `SimTree`。
*   (選用) 將變數名稱改為 `SimTree3` 以維持命名一致性，但若為了減少程式碼變動量，維持 `TreeView3` 亦可。

### [L1 UI 事件層]

#### [MODIFY] [Form1_Main.vb](file:///d:/Users/Simon/Dropbox/%E7%A7%81%E4%BA%BA%E6%96%87%E4%BB%B6/Visual%20Studio/Visual%20Studio%2018%20%282026%29/Outlook%20Assistant%20-%20%28AntiGravity%E6%B8%AC%E8%A9%A6%E5%8D%80%29/Form1_Main.vb)

##### 1. 修改 `Button3_Click` 搜尋邏輯
*   **變更選取來源**: 從單一 `TreeView3.SelectedNode` 改為遍歷 `TreeView3.SelectedNodes` 集合。
*   **資料夾匯集與去重**:
    ```vb
    ' [實作邏輯剖析]
    Dim targetFolders As New List(Of Outlook.Folder)
    Dim pathSet As New HashSet(Of String)
    ' 遍歷所有選定節點
    For Each node As TreeNode In TreeView3.SelectedNodes
        Dim rootF = TryCast(node.Tag, Outlook.Folder)
        If rootF IsNot Nothing Then
            ' 根據 CheckSubFolder3 展開子資料夾並去重
            For Each subF In GetSubFolderList(rootF, CheckSubFolder3.Checked)
                If pathSet.Add(subF.FolderPath) Then targetFolders.Add(subF)
            Next
        End If
    Next
    ```
*   **更新進度條母數**: 使用 `targetFolders.Count` 作為掃描總數。

##### 2. 連動事件調整
*   若有 `TreeView3_AfterSelect` 事件，需確認其內容是否需相容多選（例如是否要在點擊時就清空搜尋結果）。

---

## 潛在副作用與注意事項

*   **Shadows 屬性**: `SimTree` 的 `SelectedNode` 是 Shadows 屬性，回傳最後點擊的節點。現有程式碼若只讀取一個節點，不會報錯但會「漏掉其他選取項」。
*   **Ctrl+Click 誤觸**: `SimTree` 在 `MouseUp` 才觸發 `AfterSelect`，與原生行為略有差異。
*   **右鍵選單**: `SimTree` 保留了 Windows 檔案總管風格的「右鍵不改變選取狀態」，這與目前 `Form1` 的右鍵選單邏輯應能無縫銜接。

## 開放性問題

> [!QUESTION]
> 1. 您希望將變數名稱從 `TreeView3` 正式改名為 `SimTree3` 嗎？(改名會影響更多地方，但不改名會造成型別誤導)。
> 2. Tab3 是否需要像 Tab2 一樣，在切換資料夾時「自動」清空先前的搜尋結果？

## 驗證計劃

### 手動測試 (Visual Studio)
1. **單選測試**: 選中單一資料夾，點擊搜尋，確認結果與原先一致。
2. **多選測試**: 按住 Ctrl 選中多個資料夾，點擊搜尋，確認結果包含所有選定範圍。
3. **去重測試**: 選中「收件匣」及其子資料夾「工作」，開啟「含子資料夾」，確認「工作」內的郵件不會重複出現在 List 中。
4. **中斷測試**: 大範圍搜尋時按下 ESC，確認搜尋能正確停止且 UI 恢復。
