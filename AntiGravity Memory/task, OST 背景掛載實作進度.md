# OST 背景掛載實作進度

- `[x]` 1. 解除 `Niv2023 ost2pst` UI 綁定
  - `[x]` 1-1. 在 `Form1_OST.vb` 或 `Program.vb` 中建立 Dummy `mainForm`
  - `[x]` 1-2. 攔截/覆寫 `statusMSG` 避免拋出 NullReferenceException
- `[x]` 2. 修正 `Form1_OST.vb` 中的 `contentNid` 計算邏輯 (Type 4 -> 14)
- `[x]` 3. 實作 OST 資料夾複製 (Copy Folder)
  - `[x]` 3-1. 撰寫 `ost2pst.FM.CopySourceDatablocksToPST` 的呼叫與暫存檔生成
  - `[x]` 3-2. 透過 OOM 掛載暫存檔 (`_olNS.AddStore`)
  - `[x]` 3-3. 尋找掛載的資料夾並執行 `CopyTo` 到目標 PST
  - `[x]` 3-4. 卸載暫存檔並刪除實體檔案
- `[x]` 4. 實作 OST 單一郵件開啟
  - `[x]` 4-1. 測試單一郵件匯出 / 最小化資料夾匯出
  - `[x]` 4-2. OOM 背景掛載該暫存 PST
  - `[x]` 4-3. 呼叫 `MailItem.Display()` 開啟郵件視窗
- `[x]` 5. 實作 TreeView 資料夾數量顯示
  - `[x]` 5-1. 在背景執行緒讀取 TableContext 取得郵件數
  - `[x]` 5-2. 動態更新 TreeView 節點文字 `(數量)`
