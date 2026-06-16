# Task: 修正 AutoResizeLvColumns 中 EntryID 欄寬覆蓋問題

## 問題根源
- ListView5 沒有專屬 branch，掉入 Else 均分邏輯，把 EntryID(index 6) 設成 avgWidth
- ListView3、ListView4 的 EntryID 欄被設成 w*0.03，每次 Resize 都蓋掉 Width=0 的設定
- ListView4 的主旨欄計算式沒有把 EntryID(5) 的寬度扣除

## 修改清單

- [x] L1768（ListView3）: EntryID 從 `w * 0.03` 改為 `w * 0.01`
- [x] L1779（ListView4）: EntryID 從 `w * 0.03` 改為 `w * 0.01`
- [x] L1780（ListView4）: 主旨欄計算式補上 `+ lv.Columns(5).Width` 的扣除
- [x] L1782 前新增 ListView5 專屬 ElseIf branch（7欄，EntryID = w*0.01）
- [x] 複檢所有修改點確認正確、複檢修改點前後是否遺留多餘程式碼
