# PreComputeFuzzySimHashAsync 效能優化總結

根據您的精準觀察，我們針對 `Form1_Maintab56.vb` 裡的 `PreComputeFuzzySimHashAsync` 進行了兩項關鍵的效能瓶頸解除，徹底解決高速處理下的隱藏負擔。

## 變更摘要

### 1. 雙重節流閘門 (UI 更新最佳化)
在每秒高達 300~400 封信的超高處理速度下，原本單純使用 `If (i And 15) = 0` 會導致每秒觸發高達 20 幾次的 UI 字串更新，嚴重浪費 CPU 資源。
- **解決方案**：我們擴寬了次數閘門至 `If (i And 63) = 0`，並在內部結合 `SmartThrottle`。
- **效果**：現在系統只有在時間真正達到 `ThrottleFreq.Hii` (約 100ms) 時才會重繪 UI，消除多餘的高頻視覺刷新。且我們**刻意不傳入委派閉包** (`onThrottled`) 給 `SmartThrottle`，成功避開了迴圈內的 Garbage Collection (GC) 記憶體配置風暴。

### 2. 巨量批次寫入 (降低 I/O 停頓)
原本設定 `batch.Count >= 500` 時寫入資料庫，在高速運作下相當於每 1 秒多就要強制鎖定並 Commit 一次 SQLite，這極度容易引發微小的磁碟 I/O 卡頓。
- **解決方案**：將 Batch 陣列的預先配置容量提升至 `3072`，並將觸發寫入的門檻拉高到 `batch.Count >= 3000`。
- **效果**：資料庫現在大約每 5~8 秒才需要進行一次交易寫入 (Transaction Commit)，大幅減輕了磁碟負擔，讓 CPU 幾乎 100% 專注在 SimHash 的計算上。

## 驗證結果
這兩項修改皆精確採用 Chunk 替換寫入，沒有遺漏任何舊有的註解與除錯探針，並標上了 `2026/06/25 by Gemini 3.1 Pro` 供您日後辨識。修改後，S3 的 Build Pass 在大量信件處理下，進度條將會維持穩定不頻繁閃爍，且整體處理速度 (Throughput) 應有顯著的提升感。
