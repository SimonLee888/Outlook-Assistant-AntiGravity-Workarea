# 建立按下 F5 自動編譯並執行專案的環境配置

使用者希望在 AntiGravity 介面中按下 F5 鍵時，能自動觸發專案編譯並在編譯成功後執行產出的程式。由於專案包含 COM 引用，必須使用完整版 Visual Studio 2026 的 MSBuild 進行編譯。

## 使用者評論與回饋要求
> [!IMPORTANT]
> 1. 本配置將硬編碼 Visual Studio 2026 的 MSBuild 路徑 (`C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe`)。如果您未安裝 Community 版本或是安裝於不同路徑，請告知我修正。
> 2. 編譯模式預設為 `Debug` 且目標平台為 `Any CPU`。

## 擬議變更

### 開發環境配置 (.vscode)

#### [NEW] .vscode/tasks.json
定義一個編譯任務來呼叫 MSBuild 2026。

#### [NEW] .vscode/launch.json
配置啟動偵錯設定，並設定 `preLaunchTask` 預先執行上述的編譯任務。

---

## 預計產出的內容參考

### tasks.json
```json
{
    "version": "2.0.0",
    "tasks": [
        {
            "label": "Build Outlook Assistant",
            "type": "shell",
            "command": "C:\\Program Files\\Microsoft Visual Studio\\18\\Community\\MSBuild\\Current\\Bin\\MSBuild.exe",
            "args": [
                "${workspaceFolder}/Outlook Assistant.sln",
                "/p:Configuration=Debug",
                "/p:Platform=Any CPU",
                "/t:Build"
            ],
            "group": {
                "kind": "build",
                "isDefault": true
            },
            "problemMatcher": "$msCompile"
        }
    ]
}
```

### launch.json
```json
{
    "version": "0.2.0",
    "configurations": [
        {
            "name": "Build and Run (Debug)",
            "type": "coreclr",
            "request": "launch",
            "preLaunchTask": "Build Outlook Assistant",
            "program": "${workspaceFolder}/bin/Debug/net10.0-windows10.0.17763.0/Outlook Assistant.exe",
            "args": [],
            "cwd": "${workspaceFolder}",
            "stopAtEntry": false,
            "console": "internalConsole"
        }
    ]
}
```

## 開放問題
1. **編譯模式**: 除了 Debug 模式外，是否需要另外建立一個 Release 模式的啟動配置？
2. **執行參數**: 程式啟動時是否需要帶入任何命令列參數？

## 驗證計畫

### 自動化測試 / 驗證
- 我已手動執行過編譯指令，確認 `C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe` 有效。
- 確認產出的執行檔路徑正確。

### 手動驗證
- 設定完成後，請使用者在 AntiGravity 環境中按下 **F5** 鍵，觀察是否會自動編譯並啟動程式。
