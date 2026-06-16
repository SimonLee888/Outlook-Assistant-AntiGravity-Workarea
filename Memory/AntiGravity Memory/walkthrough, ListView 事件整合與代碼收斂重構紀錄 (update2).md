# ListView 事件整合與代碼收斂重構紀錄

為了解決 ListView3（Tab3）及 ListView4（Tab4）的事件處理重複、過度分散以及維護困難的問題，我們順利完成了一次深度重構。以下是修改細節的總結與架構梳理（已套用使用者最新的 `HandleLv3Lv4_` 命名約定）。

## 統一呼叫層級設計

我們遵照指示，確立了乾淨直觀的樹狀呼叫層次，所有通用的搜尋結果行為（複製、預覽、展開郵件等）被徹底提升（Lift Up）到了底層方法。

```mermaid
flowchart TD
    A[Form1.InitListView] -->|AddHandler| B(HandleLv3Lv4_MouseClick)
    A -->|AddHandler| C(HandleLv3Lv4_DoubleClick)
    A -->|AddHandler| D(HandleLv3Lv4_KeyPress)
    A -->|AddHandler| E(HandleLv3Lv4_KeyDown)
    
    %% 加入這行隱形連結，強迫 K 降階到與 B, C, D, E 同一層
    A ~~~ K

    B --> F{ShowFolderPathToProgressBar}
    B -.-> |Left Click| G[Clipboard.SetText]
    C --> H[OpenMailByEntryID]
    D -.-> |Enter| H
    E -.-> |Ctrl+A| I([Select All Items])
    E -.-> |ESC| J([Focus to Left Side SimTree])
    
    K[ListView4_KeyDown] -->|Delete| L[HandleListView4Delete]
    K -->|F5| M[RefreshListView4MailsAsync]
```

## 實作細節亮點

### 1. 通用邏輯層提取 (`Form1_MainTabs.vb`)
我們將原本散亂在不同地方的邏輯，抽離成一組 `HandleLv3Lv4_XXX` 的子程式，並且利用 `VirtualMode` 屬性靈活地相容了兩種模式：

- **`GetSelectedEntryIDs`**: 能夠聰明地判斷呼叫者是 `ListView3` (虛擬模式) 或 `ListView4` (實體模式)，擷取正確的 EntryID。
- **`HandleLv3Lv4_MouseClick`**: 單擊左鍵即可複製主旨（Tab3 現在也受惠於原本 Tab4 的便利功能），同時透過呼叫既有的 `ShowFolderPathToProgressBar` 同步下方狀態列。
- **`HandleLv3Lv4_KeyDown`**: 結合了 `Ctrl+A` (全選) 以及共通的 `ESC` 焦點歸位；如果是 `ListView3` 按下 `ESC` 則退回 `SimTree3`，若為 `ListView4` 按下則退回 `SimTree4`。

### 2. 動態事件綁定 (`Form1.vb`)
原本的靜態 `Handles` 被替換，改為在 UI 初次建立時的 `InitListView` 中統一動態掛載：
```vb
' Form1.vb / InitListView
If lv.Name = "ListView3" OrElse lv.Name = "ListView4" Then
    AddHandler lv.SelectedIndexChanged, AddressOf ShowFolderPathToProgressBar
    AddHandler lv.MouseClick, AddressOf HandleLv3Lv4_MouseClick
    AddHandler lv.MouseDoubleClick, AddressOf HandleLv3Lv4_DoubleClick
    AddHandler lv.KeyPress, AddressOf HandleLv3Lv4_KeyPress
    AddHandler lv.KeyDown, AddressOf HandleLv3Lv4_KeyDown
End If
```

### 3. 大幅刪減冗餘代碼
在 `Form1_MainTabs.vb` 中，下列專為 `ListView4` 硬刻的方法已被整併並刪除，讓整份類別少了快 100 行重複的程式碼：
- `[DELETE]` `ListView4_SelectedIndexChanged`
- `[DELETE]` `ListView4_MouseClick`
- `[DELETE]` `ListView4_MouseDoubleClick`
- `[DELETE]` `ListView4_KeyPress`
- `[MODIFY]` `ListView4_KeyDown` (僅殘留專屬於 Tab4 的動作：Delete 與 F5)

## 歷史註解保留
這批修改在對應的地方都加上了相關註解與操作緣由，隨後經過手動優化命名為 `HandleLv3Lv4_` 系列，讓事件的用途與關聯性變得更加一目了然。
