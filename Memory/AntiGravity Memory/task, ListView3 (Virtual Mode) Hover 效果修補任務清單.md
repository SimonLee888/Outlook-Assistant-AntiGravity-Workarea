# ListView3 (Virtual Mode) Hover 效果修補任務清單

- [ ] 修改 `Form1.vb`
    - [ ] 移除 `HandleLvMouseHover` 中的 `If listView.VirtualMode Then Return` 限制。
    - [ ] 確保在 `listView.OwnerDraw` 為 True 時，虛擬模式也能觸發 `Invalidate`。
- [ ] 修改 `Form1_MainTabs.vb`
    - [ ] 在 `InitTab3UI` 中啟用 `ListView3.OwnerDraw = True`。
    - [ ] 實作 `Lv3_DrawColumnHeader` (DrawDefault = True)。
    - [ ] 實作 `Lv3_DrawItem` 與 `Lv3_DrawSubItem` (比照 ListView4 的高效繪製邏輯)。
- [ ] 複檢所有修改點確認正確、複檢修改點前後是否遺留多餘程式碼。
