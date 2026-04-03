# UI 色彩整理與建議

目前在 `Form1.vb`、`Form1_Main.vb`、`DebugForm.vb` 中，用到了好幾種 `Color.FromArgb(...)`，尤其是有好幾種非常接近的灰色。

## 📍 目前使用到的色彩比較

| RGB 代碼 | 預覽 (HEX) | 用途觀察 | 建議統一命名 |
| :--- | :--- | :--- | :--- |
| `242, 242, 242` | <span style="display:inline-block; width:40px; height:20px; background-color:#F2F2F2; border:1px solid #ccc;"></span> (`#F2F2F2`) | 視窗背景色 (BackColor)、Tab/Panel 底色 | **Bg_Window** |
| `240, 240, 240` | <span style="display:inline-block; width:40px; height:20px; background-color:#F0F0F0; border:1px solid #ccc;"></span> (`#F0F0F0`) | TreeView/ListView 的 MouseHover (滑鼠懸停) 顏色 | **Bg_Hover** |
| `224, 224, 224` | <span style="display:inline-block; width:40px; height:20px; background-color:#E0E0E0; border:1px solid #ccc;"></span> (`#E0E0E0`) | 圖表邊線 (BorderlineColor, MajorGrid) | **Border_Light** |
| `0, 120, 212` | <span style="display:inline-block; width:40px; height:20px; background-color:#0078D4; border:1px solid #ccc;"></span> (`#0078D4`) | 主按鈕、重要焦點字體 (經典微軟藍) | **Brand_Blue** |

> [!NOTE]  
> 仔細看 `#F2F2F2` (這是一般視窗底色) 與 `#F0F0F0` (Hover 底色)，兩者視覺上的差異微乎其微。通常 Hover (懸停) 的顏色會與背景有更強一點的對比：例如背景給 `#F2F2F2`，Hover 可以配稍微深一點點的 `#E5E5E5`，您可以在設定統一風格後再做微調觀察！

---

## 🛠️ 程式碼重構建議

為了不要每次都重新打 `Color.FromArgb(...)` 讓程式碼變得難以閱讀或不小心打錯，我建議我們建立一個集中的**色彩常數類別**。

以您的專案架構，您可以選擇在 `Form1_Main.vb` (或其他存放全域結構的地方) 建立一個類似下面的語法：

```vb
Public Class ThemeColors
    ' --- 背景色 ---
    ''' <summary>主要視窗或Panel背景色 (#F2F2F2)</summary>
    Public Shared ReadOnly Bg_Window As Color = Color.FromArgb(242, 242, 242)
    
    ''' <summary>滑鼠懸停(Hover)的背景色 (#E5E5E5) - 建議可稍微加深以增加對比</summary>
    Public Shared ReadOnly Bg_Hover As Color = Color.FromArgb(229, 229, 229)
    
    ' --- 邊框與文字色 ---
    ''' <summary>輕微的格線或邊框色 (#E0E0E0)</summary>
    Public Shared ReadOnly Border_Light As Color = Color.FromArgb(224, 224, 224)
    
    ''' <summary>主視覺品牌藍色 (如按鈕、連結文字)</summary>
    Public Shared ReadOnly Brand_Blue As Color = Color.FromArgb(0, 120, 212)
End Class
```

未來在各個 `Form.vb` 或控制項佈局宣告的程式碼裡，就只需要改成：
```vb
Me.BackColor = ThemeColors.Bg_Window
node.BackColor = ThemeColors.Bg_Hover
Chart2.BorderlineColor = ThemeColors.Border_Light
Button3.ForeColor = ThemeColors.Brand_Blue
```

這麼做不僅可以讓所有元件都保持相同的基礎風格，如果未來的某一天您想要整體換風格 (例如加入「深色模式 Dark Mode」)，您就只要將 `ThemeColors` 裡面的顏色透過判斷式切換即可，再也不需要跑遍所有的 `*.vb` 檔案大海撈針去取代 `FromArgb` 的色碼了！
