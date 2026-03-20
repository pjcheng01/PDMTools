# PDMTools

SolidWorks PDM 組合件 BOM 匯出工具（WPF）：遞迴參考樹、讀取資料卡變數、匯出 Excel（ClosedXML）。

## 分支說明

| 分支 | 目標框架 | 說明 |
|------|-----------|------|
| **`main`** | .NET 8（`net8.0-windows`） | 目前主力開發與預設發行線。需安裝 [.NET 8 Desktop Runtime](https://dotnet.microsoft.com/download/dotnet/8.0)。 |
| **`net48`** | .NET Framework 4.8（`net48`） | **本分支專案已設為 `net48`。** 執行環境需已安裝 **.NET Framework 4.8**；建置需 **.NET Framework 4.8 Developer Pack**（或含 4.8 開發工具的 Visual Studio）。 |

切換分支範例：

```bash
git checkout main
git checkout net48
```

## 建置（net48 分支）

```bash
dotnet build PDMTools\PDMTools.csproj -c Release
```

輸出目錄：`PDMTools\bin\Release\net48\`

或雙擊 `build_release_and_open.bat`（一鍵 Release 並開啟上述資料夾）。

## 建置（main 分支 / .NET 8）

若切回 `main`，目標為 `net8.0-windows`，請使用該分支的 `csproj` 與對應 Runtime。

## net48 分支：建議的 Commit 訊息範本

在 **`net48`** 上提交時，建議訊息格式如下，方便日後對照與 cherry-pick：

```
類型(範圍): 簡短說明（50 字內）

- 變更要點 1
- 變更要點 2

類型: feat | fix | docs | refactor | chore
範圍: net48 | pdm | excel | ui（可選）
```

**範例：**

```
feat(net48): 將 TargetFramework 改為 net48

- 調整 csproj 與 Nullable 設定
- 驗證 ClosedXML 與 EPDM Interop 於 4.8 建置通過
```

```
fix(pdm): 修正 GetReferenceTree 於 net48 的 COM 釋放時機
```

## 授權與相依

- PDM API：SolidWorks PDM 客戶端與 `SolidWorks.EPDM.Interop.epdm`（NuGet）。
- Excel：`ClosedXML`（不使用 Excel Interop）。

## 本機 Vault 路徑

預設視圖根目錄為 `C:\CP-PDM`（見 `PdmBomExportService` 常數，可依環境調整）。
