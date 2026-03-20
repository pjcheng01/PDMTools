# PDMTools

SolidWorks PDM 組合件 BOM 匯出工具（WPF）：遞迴參考樹、讀取資料卡變數、匯出 Excel（ClosedXML）。

## 分支說明

| 分支 | 目標框架 | 說明 |
|------|-----------|------|
| **`main`** | .NET 8（`net8.0-windows`） | 目前主力開發與預設發行線。需安裝 [.NET 8 Desktop Runtime](https://dotnet.microsoft.com/download/dotnet/8.0)。 |
| **`net48`** | .NET Framework 4.8 | 與企業僅允許 4.x 或舊式部署環境對齊的開發線；程式碼將逐步調整為 `net48` 相容。需安裝 **.NET Framework 4.8 Developer Pack** 以建置。 |

切換分支範例：

```bash
git checkout main
git checkout net48
```

## 建置（main / .NET 8）

```bash
dotnet build PDMTools\PDMTools.csproj -c Release
```

或雙擊專案根目錄的 `build_release_and_open.bat`（會建置 Release 並開啟輸出資料夾）。

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
