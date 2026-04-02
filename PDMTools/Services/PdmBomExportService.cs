using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ClosedXML.Excel;
using EPDM.Interop.epdm;
using PDMTools.Models;
using PDMTools.Utils;

namespace PDMTools.Services
{
    public sealed class PdmBomExportService
    {
        /// <summary>Vault 本機根目錄路徑。可由 UI 在執行期動態設定（預設為安裝時的路徑）。</summary>
        public static string VaultRootPath { get; set; } = @"C:\CP-PDM";

        /// <summary>
        /// 指定登入時優先嘗試的 Vault 名稱（例如從 UI 下拉選單選取的名稱）。
        /// 設定後，EnsureVaultLogin 會將此名稱排在候選清單的第一位。
        /// </summary>
        public string OverrideVaultName { get; set; }
        // 依 PDM 變數名稱抓值（優先使用實際變數代號，而非畫面顯示標籤）
        private static readonly CardVariableSpec[] CardVariableSpecs =
        {
            new CardVariableSpec("CP_中文品名", "CP_中文品名"),
            new CardVariableSpec("CP_功能單元", "CP_功能單元"),
            new CardVariableSpec("CP_市購品供應商", "CP_市購品供應商"),
            new CardVariableSpec("CP_材質", "CP_材質"),
            new CardVariableSpec("CP_表面處理", "CP_表面處理"),
            new CardVariableSpec("CP_英文品名", "CP_英文品名"),
            new CardVariableSpec("CP_產品", "CP_產品"),
            new CardVariableSpec("CP_規格", "CP_規格"),
            new CardVariableSpec("CP_備註", "CP_備註"),
            new CardVariableSpec("CP_單位", "CP_單位"),
            new CardVariableSpec("CP_順次", "CP_順次"),
            new CardVariableSpec("CP_圖號", "CP_圖號"),
            new CardVariableSpec("CP_機種", "CP_機種"),
            new CardVariableSpec("CP_歸檔目錄", "CP_歸檔目錄"),
            new CardVariableSpec("iGP_批號管理", "iGP_批號管理"),
            new CardVariableSpec("iGP_材料型態", "iGP_材料型態"),
            new CardVariableSpec("iGP_品號類別", "iGP_品號類別"),
            new CardVariableSpec("iGP_品號屬性", "iGP_品號屬性"),
            new CardVariableSpec("iGP_補貨政策", "iGP_補貨政策"),
            new CardVariableSpec("iGP_領料碼", "iGP_領料碼"),
            new CardVariableSpec("iGP_檢驗方式", "iGP_檢驗方式"),
            new CardVariableSpec("本機修訂版", "本機修訂版"),
            new CardVariableSpec("類別", "類別"),
            new CardVariableSpec("工作流程", "工作流程"),
            new CardVariableSpec("本機版本", "本機版本")
        };

        /// <summary>與 Excel 匯出「Card:」欄位順序一致，供 UI DataGrid 分欄綁定。</summary>
        private static readonly IReadOnlyList<string> OrderedCardVariableLabelList =
            CardVariableSpecs.Select(s => s.Label).ToList();

        public static IReadOnlyList<string> GetOrderedCardVariableLabels() => OrderedCardVariableLabelList;

        /// <summary>LoginAuto 逾時秒數（預設 30 秒）。Server 無回應時超過此時間將拋出 TimeoutException。</summary>
        public int LoginTimeoutSeconds { get; set; } = 30;

        private readonly IEdmVault5 _vault;

        public PdmBomExportService()
        {
            _vault = new EdmVault5();
        }

        // ════════════════════════════════════════════════════════════════════
        // PDM 環境偵測（靜態，不需 Vault 登入）
        // ════════════════════════════════════════════════════════════════════

        /// <summary>
        /// 偵測本機是否已安裝 SolidWorks PDM 用戶端（嘗試建立 COM 物件）。
        /// </summary>
        public static bool IsPdmClientInstalled()
        {
            try
            {
                var t = new EdmVault5();
                Marshal.ReleaseComObject(t);
                return true;
            }
            catch (COMException) { return false; }
            catch { return false; }
        }

        /// <summary>
        /// 列舉本機已安裝的所有 Vault 本機視圖（讀取 PDM 用戶端 Registry，不需登入）。
        /// 回傳值為 (VaultName, LocalRootPath) 清單；若偵測失敗則回傳空清單。
        /// </summary>
        public static IReadOnlyList<(string VaultName, string LocalPath)> GetLocalVaultViews()
            => GetLocalVaultViews(out _);

        /// <summary>
        /// 列舉本機 Vault 本機視圖，並在 <paramref name="diagnosticMessage"/> 輸出除錯記錄。
        /// </summary>
        public static IReadOnlyList<(string VaultName, string LocalPath)> GetLocalVaultViews(
            out string diagnosticMessage)
        {
            var result = new List<(string, string)>();
            var diag   = new StringBuilder();

            // ── 方法一：搜尋 RCW 實作的所有介面，找 GetVaultViews ──────────────
            // 注意：GetVaultViews 不在 IEdmVault5，在更高版本介面（IEdmVault8+），
            // 所以不能直接轉型 IEdmVault5 呼叫，必須透過介面反射搜尋。
            try
            {
                var tempVault = new EdmVault5();
                try
                {
                    var allIfaces = tempVault.GetType().GetInterfaces();
                    diag.AppendLine($"[方法一] RCW 實作 {allIfaces.Length} 個介面");

                    foreach (var iface in allIfaces)
                    {
                        var mi = iface.GetMethod("GetVaultViews",
                            BindingFlags.Instance | BindingFlags.Public);
                        if (mi == null) continue;

                        diag.AppendLine($"  找到 {iface.Name}.GetVaultViews，嘗試呼叫...");
                        try
                        {
                            var args = new object[] { null, false };
                            mi.Invoke(tempVault, args);
                            if (args[0] is EdmViewInfo[] views)
                            {
                                diag.AppendLine($"  成功，{views.Length} 個視圖");
                                foreach (var v in views)
                                    result.Add((v.mbsVaultName ?? string.Empty, v.mbsPath ?? string.Empty));
                            }
                            else
                            {
                                diag.AppendLine($"  args[0] 型別：{args[0]?.GetType()?.FullName ?? "null"}");
                            }
                        }
                        catch (Exception ex2)
                        {
                            diag.AppendLine($"  呼叫失敗：{ex2.Message}");
                        }
                        break;  // 只嘗試第一個找到的介面方法
                    }

                    if (result.Count == 0 && !allIfaces.Any(i =>
                        i.GetMethod("GetVaultViews", BindingFlags.Instance | BindingFlags.Public) != null))
                    {
                        diag.AppendLine("  所有介面均未找到 GetVaultViews");
                    }
                }
                finally { Marshal.ReleaseComObject(tempVault); }
            }
            catch (Exception ex)
            {
                diag.AppendLine($"[方法一] 例外：{ex.GetType().Name}: {ex.Message}");
            }

            // ── 方法三：直接讀 Registry（備援）──────────────────────────────────
            if (result.Count == 0)
            {
                diag.AppendLine("[方法三] COM 未取得結果，改從 Registry 讀取...");
                TryGetVaultViewsFromRegistry(result, diag);
            }

            diagnosticMessage = diag.ToString();
            return result;
        }

        /// <summary>從 Windows Registry 讀取 PDM 本機視圖（備援方案）。</summary>
        private static void TryGetVaultViewsFromRegistry(
            List<(string, string)> result, StringBuilder diag)
        {
            // PDM 在不同版本與安裝語系下可能使用不同的 Registry 路徑。
            // 每個路徑以 (Hive, SubPath) 表示，同時掃描 HKLM 與 HKCU。
            var subPaths = new[]
            {
                @"SOFTWARE\SolidWorks\Applications\PDMWorks Enterprise\Databases",
                @"SOFTWARE\WOW6432Node\SolidWorks\Applications\PDMWorks Enterprise\Databases",
                @"SOFTWARE\SolidWorks\SOLIDWORKS PDM\Databases",
                @"SOFTWARE\WOW6432Node\SolidWorks\SOLIDWORKS PDM\Databases",
                @"SOFTWARE\SolidWorks\Applications\PDMWorks Enterprise\Settings\Databases",
                @"SOFTWARE\WOW6432Node\SolidWorks\Applications\PDMWorks Enterprise\Settings\Databases",
            };

            // 同時掃描 HKLM（系統級安裝）與 HKCU（使用者級設定）
            var hives = new[]
            {
                (Microsoft.Win32.Registry.LocalMachine, "HKLM"),
                (Microsoft.Win32.Registry.CurrentUser,  "HKCU"),
            };

            foreach (var (hive, hiveName) in hives)
            {
                foreach (var regPath in subPaths)
                {
                    try
                    {
                        using (var key = hive.OpenSubKey(regPath))
                        {
                            if (key == null) { diag.AppendLine($"  {hiveName}\\{regPath} 不存在"); continue; }

                            var subNames = key.GetSubKeyNames();
                            diag.AppendLine($"  {hiveName}\\{regPath} 找到 {subNames.Length} 個子機碼");

                            foreach (var sub in subNames)
                            {
                                using (var subKey = key.OpenSubKey(sub))
                                {
                                    if (subKey == null) continue;
                                    var vaultName = subKey.GetValue("VaultName") as string
                                                 ?? subKey.GetValue("Name")      as string
                                                 ?? sub;
                                    var localPath = subKey.GetValue("LocalPath")  as string
                                                 ?? subKey.GetValue("RootPath")   as string
                                                 ?? subKey.GetValue("ViewPath")   as string
                                                 ?? string.Empty;
                                    diag.AppendLine($"    {sub}: VaultName={vaultName}, Path={localPath}");
                                    if (!string.IsNullOrWhiteSpace(vaultName)
                                        && !result.Any(r => string.Equals(r.Item1, vaultName,
                                            StringComparison.OrdinalIgnoreCase)))
                                        result.Add((vaultName, localPath));
                                }
                            }
                            if (result.Count > 0) return;
                        }
                    }
                    catch (Exception ex)
                    {
                        diag.AppendLine($"  {hiveName}\\{regPath} 讀取例外：{ex.Message}");
                    }
                }
            }
        }

        // ════════════════════════════════════════════════════════════════════
        // Vault 變數列舉（動態偵測所有已定義的資料卡變數）
        // ════════════════════════════════════════════════════════════════════

        /// <summary>
        /// 列舉 Vault 中所有已定義的資料卡變數名稱，排序後回傳。
        /// 若列舉失敗，則回傳內建的 CardVariableSpecs 清單作為備援。
        /// </summary>
        /// <param name="sampleFilePath">
        /// 可選：提供一個 Vault 內的檔案路徑，作為列舉卡片變數的樣本。
        /// 若 Vault 層級列舉失敗，將改從此檔案的變數枚舉器列舉。
        /// </param>
        public async Task<IReadOnlyList<string>> EnumerateVaultVariablesAsync(string sampleFilePath = null)
        {
            return await Task.Run(() =>
            {
                EnsureVaultLogin();
                return EnumerateVaultVariablesInternal(sampleFilePath);
            });
        }

        /// <summary>上次執行列舉的診斷記錄（供 UI 顯示）。</summary>
        public string LastEnumerationDiag { get; private set; } = string.Empty;

        /// <summary>上次抓取時實際使用的 BOM 版面名稱（經計算 BOM）。</summary>
        public string LastBomLayoutNameUsed { get; private set; } = string.Empty;

        /// <summary>上次抓取時解析後的 SolidWorks 組態名稱。</summary>
        public string LastConfigurationResolved { get; private set; } = string.Empty;

        /// <summary>上次嘗試列舉組態時的反射／呼叫記錄（列舉失敗時供除錯）。</summary>
        public string LastConfigurationEnumerationDiag { get; private set; } = string.Empty;

        /// <summary>
        /// 上次從 PDM 檔案物件讀到的「文件作用中組態」名稱（供 UI 預設選取）；若 API 無此屬性或讀取失敗則為空字串。
        /// </summary>
        public string LastDocumentActiveConfiguration { get; private set; } = string.Empty;

        /// <summary>
        /// 從指定檔案的 GetEnumeratorVariable() 枚舉所有可用的卡片變數名稱。
        /// 這是 Vault 層級列舉失敗時的備援方式。
        /// </summary>
        private void TryEnumerateFromFile(string filePath, List<string> names, System.Text.StringBuilder diag)
        {
            IEdmFolder5 folder = null;
            IEdmFile5   file   = null;
            try
            {
                file = _vault.GetFileFromPath(filePath, out folder);
                if (file == null) { diag.AppendLine("TryEnumerateFromFile：找不到檔案。"); return; }

                // 轉型為強型別介面，使用 vtable 而非 IDispatch
                var enumVarTyped = file.GetEnumeratorVariable() as IEdmEnumeratorVariable10;
                if (enumVarTyped == null) { diag.AppendLine("TryEnumerateFromFile：GetEnumeratorVariable() 無法轉型。"); return; }

                // 透過介面型別反射找方法（vtable dispatch，不走 IDispatch）
                var ifaceType   = typeof(IEdmEnumeratorVariable10);
                var relMethods  = ifaceType
                    .GetMethods(BindingFlags.Instance | BindingFlags.Public)
                    .Where(m => m.Name.IndexOf("Var",      StringComparison.OrdinalIgnoreCase) >= 0
                             || m.Name.IndexOf("Variable", StringComparison.OrdinalIgnoreCase) >= 0
                             || m.Name.IndexOf("Position", StringComparison.OrdinalIgnoreCase) >= 0)
                    .Select(m => m.Name).Distinct().OrderBy(n => n).ToList();
                diag.AppendLine($"IEdmEnumeratorVariable10 相關方法：{(relMethods.Count > 0 ? string.Join(", ", relMethods) : "（無）")}");

                var getFirstPos = ifaceType.GetMethod("GetFirstVariablePosition",
                    BindingFlags.Instance | BindingFlags.Public);
                var getNextVar  = ifaceType.GetMethod("GetNextVariable",
                    BindingFlags.Instance | BindingFlags.Public);

                diag.AppendLine($"枚舉器.GetFirstVariablePosition：{(getFirstPos != null ? "找到" : "找不到")}");
                diag.AppendLine($"枚舉器.GetNextVariable：{(getNextVar != null ? "找到" : "找不到")}");

                if (getFirstPos == null || getNextVar == null)
                {
                    diag.AppendLine("介面上找不到位置列舉方法，停止。");
                    return;
                }

                // GetFirstVariablePosition 需要 configName 參數（傳 "@" 代表檔案層級）
                var firstPosParams = getFirstPos.GetParameters();
                diag.AppendLine($"GetFirstVariablePosition 參數數：{firstPosParams.Length}");

                object posObj;
                try
                {
                    posObj = firstPosParams.Length == 0
                        ? getFirstPos.Invoke(enumVarTyped, null)
                        : getFirstPos.Invoke(enumVarTyped, new object[] { "@" });
                }
                catch (Exception ex)
                {
                    diag.AppendLine($"GetFirstVariablePosition 呼叫失敗：{ex.Message}");
                    return;
                }

                // 轉型為 IEdmPos5 使用 vtable
                var typedPos = posObj as IEdmPos5;
                if (typedPos == null)
                {
                    diag.AppendLine($"pos 無法轉換為 IEdmPos5（型別：{posObj?.GetType()?.FullName ?? "null"}），停止。");
                    return;
                }

                var safetyLimit = 1000;
                while (!typedPos.IsNull && safetyLimit-- > 0)
                {
                    object variable;
                    try
                    {
                        var args = new object[] { typedPos };
                        variable = getNextVar.Invoke(enumVarTyped, args);
                    }
                    catch (System.Reflection.TargetInvocationException tie)
                    {
                        diag.AppendLine($"枚舉器 GetNextVariable 真實錯誤：{tie.InnerException?.GetType()?.Name}: {tie.InnerException?.Message}");
                        break;
                    }
                    catch (Exception ex)
                    {
                        diag.AppendLine($"枚舉器 GetNextVariable 失敗：{ex.Message}");
                        break;
                    }
                    if (variable == null) break;

                    // 轉型為強型別取得名稱（vtable）
                    var typedVar = variable as IEdmVariable5;
                    if (typedVar != null)
                    {
                        var name = typedVar.Name;
                        if (!string.IsNullOrWhiteSpace(name))
                            names.Add(name);
                    }
                }
                diag.AppendLine($"檔案層級列舉結果：{names.Count} 個變數");
            }
            catch (Exception ex)
            {
                diag.AppendLine($"TryEnumerateFromFile 例外：{ex.GetType().Name}: {ex.Message}");
            }
            finally
            {
                ComHelper.Release(file);
                ComHelper.Release(folder);
            }
        }

        /// <summary>同步版本，供已在背景執行緒的呼叫端使用。</summary>
        private IReadOnlyList<string> EnumerateVaultVariablesInternal(string sampleFilePath = null)
        {
            var names = new List<string>();
            var diag  = new System.Text.StringBuilder();
            var vaultType = _vault.GetType();

            // ── 1. 列出 Vault 上所有含 "Var" 或 "Variable" 的方法（診斷用）──
            var varMethods = vaultType
                .GetMethods(BindingFlags.Instance | BindingFlags.Public)
                .Where(m => m.Name.IndexOf("Var", StringComparison.OrdinalIgnoreCase) >= 0
                         || m.Name.IndexOf("Variable", StringComparison.OrdinalIgnoreCase) >= 0)
                .Select(m => m.Name)
                .Distinct()
                .OrderBy(n => n)
                .ToList();

            diag.AppendLine($"Vault 型別：{vaultType.FullName}");
            diag.AppendLine($"含 Var/Variable 的方法（共 {varMethods.Count}）：");
            diag.AppendLine(varMethods.Count > 0
                ? string.Join(", ", varMethods)
                : "（無）");

            // ── 2. 嘗試 GetFirstVariablePosition / GetNextVariable ──────────
            var getFirstPos = vaultType.GetMethod("GetFirstVariablePosition",
                BindingFlags.Instance | BindingFlags.Public);
            var getNextVar  = vaultType.GetMethod("GetNextVariable",
                BindingFlags.Instance | BindingFlags.Public);

            diag.AppendLine($"GetFirstVariablePosition：{(getFirstPos != null ? "找到" : "找不到")}");
            diag.AppendLine($"GetNextVariable：{(getNextVar != null ? "找到" : "找不到")}");

            // 記錄 GetNextVariable 的參數型別（判斷是否為 ref 參數）
            if (getNextVar != null)
            {
                foreach (var p in getNextVar.GetParameters())
                    diag.AppendLine($"  參數 {p.Name}：{p.ParameterType.FullName}  IsByRef={p.ParameterType.IsByRef}  IsOut={p.IsOut}");
            }

            if (getFirstPos != null && getNextVar != null)
            {
                try
                {
                    var pos = getFirstPos.Invoke(_vault, null);
                    diag.AppendLine($"GetFirstVariablePosition() 回傳型別：{pos?.GetType()?.FullName ?? "null"}");

                    // 將 pos 轉型為強型別介面，使用 vtable 而非 IDispatch（避免 TYPE_E_LIBNOTREGISTERED）
                    var typedPos = pos as IEdmPos5;
                    if (typedPos == null)
                    {
                        diag.AppendLine("pos 無法轉換為 IEdmPos5，停止 Vault 層級列舉。");
                    }
                    else
                    {
                        var safetyLimit = 1000;
                        while (!typedPos.IsNull && safetyLimit-- > 0)
                        {
                            object variable;
                            try
                            {
                                // 以強型別 IEdmPos5 傳入，走 vtable dispatch
                                var args = new object[] { typedPos };
                                variable = getNextVar.Invoke(_vault, args);
                            }
                            catch (System.Reflection.TargetInvocationException tie)
                            {
                                diag.AppendLine($"GetNextVariable 真實錯誤：{tie.InnerException?.GetType()?.Name}: {tie.InnerException?.Message}");
                                break;
                            }
                            catch (Exception ex)
                            {
                                diag.AppendLine($"GetNextVariable 失敗：{ex.Message}");
                                break;
                            }
                            if (variable == null) break;

                            // 轉型為強型別取得 Name（vtable）
                            var typedVar = variable as IEdmVariable5;
                            if (typedVar != null)
                            {
                                var name = typedVar.Name;
                                if (!string.IsNullOrWhiteSpace(name))
                                    names.Add(name);
                            }
                        }
                    }
                    diag.AppendLine($"Vault 層級列舉結果：{names.Count} 個變數");
                }
                catch (Exception ex)
                {
                    diag.AppendLine($"列舉過程例外：{ex.GetType().Name}: {ex.Message}");
                }
            }

            // ── 3. Vault 層級失敗 → 改從樣本檔案的枚舉器列舉 ──────────────
            if (names.Count == 0 && !string.IsNullOrWhiteSpace(sampleFilePath))
            {
                diag.AppendLine($"→ Vault 層級無結果，改從樣本檔案列舉：{Path.GetFileName(sampleFilePath)}");
                TryEnumerateFromFile(sampleFilePath, names, diag);
            }

            // ── 4. 備援 ─────────────────────────────────────────────────────
            if (names.Count == 0)
            {
                diag.AppendLine("→ 回退至內建 CardVariableSpecs（25 個）。");
                names.AddRange(CardVariableSpecs.Select(s => s.Label));
            }

            LastEnumerationDiag = diag.ToString();

            return names
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(n => n, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        /// <summary>
        /// 以 PDM「經計算的 BOM」（IEdmFile7.GetComputedBOM）取得階層與用量；讀取資料卡變數仍用 PDM API。
        /// </summary>
        /// <param name="configurationName">SolidWorks 組態名稱；空白時優先使用文件作用中組態，其次「預設」／Default／清單第一筆。</param>
        public async Task<IReadOnlyList<BomItem>> CollectBomAsync(
            string assemblyPath,
            IReadOnlyList<string> cardVarNames,
            bool includeRootBomItem,
            int? maxBomLayerDepth,
            string configurationName,
            IProgress<ProgressInfo> progress,
            CancellationToken cancellationToken = default)
        {
            return await Task.Run(() =>
            {
                if (string.IsNullOrWhiteSpace(assemblyPath))
                {
                    throw new ArgumentException("請先選擇組合件檔案。", nameof(assemblyPath));
                }

                if (!File.Exists(assemblyPath))
                {
                    throw new FileNotFoundException("找不到指定的組合件檔案。", assemblyPath);
                }

                if (!assemblyPath.StartsWith(VaultRootPath, StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException($"檔案必須位於 {VaultRootPath} 內。");
                }

                EnsureVaultLogin();
                // 第一次抓取：進度條獨立使用 0～100（與「顯示工程圖」分開，互不沿用區段）。
                progress?.Report(new ProgressInfo(0, "已登入 Vault，正在開啟經計算的 BOM…"));

                IEdmFolder5 folder = null;
                IEdmFile5 file = null;
                IEdmBomMgr bomMgr = null;
                IEdmBomView bomView = null;

                try
                {
                    file = _vault.GetFileFromPath(assemblyPath, out folder);
                    if (file == null || folder == null)
                    {
                        throw new InvalidOperationException("無法從 PDM 取得檔案資訊。");
                    }

                    var file7 = file as IEdmFile7;
                    if (file7 == null)
                    {
                        throw new InvalidOperationException(
                            "目前 PDM Interop 不支援 IEdmFile7，無法取得「經計算的 BOM」。請確認已安裝 SolidWorks PDM Professional 並使用對應版本的 EPDM Interop。");
                    }

                    var vault7 = _vault as IEdmVault7;
                    if (vault7 == null)
                    {
                        throw new InvalidOperationException("無法將 Vault 轉型為 IEdmVault7，無法建立 BOM 管理員。");
                    }

                    bomMgr = vault7.CreateUtility(EdmUtility.EdmUtil_BomMgr) as IEdmBomMgr;
                    if (bomMgr == null)
                    {
                        throw new InvalidOperationException("無法建立 IEdmBomMgr（BOM 管理員）。");
                    }

                    EdmBomLayout[] layouts = null;
                    bomMgr.GetBomLayouts(out layouts);
                    if (layouts == null || layouts.Length == 0)
                    {
                        throw new InvalidOperationException(
                            "Vault 未回傳任何 BOM 版面。請在 PDM 管理工具確認已設定 BOM 類型／版面。");
                    }

                    var layout = SelectPrimaryBomLayout(layouts);
                    LastBomLayoutNameUsed = layout.mbsLayoutName ?? string.Empty;
                    var configResolved = ResolveConfigurationName(file7, file, folder, configurationName);
                    LastConfigurationResolved = configResolved;

                    bomView = TryOpenComputedBomView(file7, file, layout, configResolved);
                    if (bomView == null)
                    {
                        throw new InvalidOperationException(
                            $"無法取得經計算的 BOM（版面「{LastBomLayoutNameUsed}」、組態「{configResolved}」）。請在 PDM 中確認「經計算的 BOM」與組態名稱是否一致。");
                    }

                    EdmBomColumn[] columns = null;
                    bomView.GetColumns(out columns);

                    object[] rows = null;
                    bomView.GetRows(out rows);
                    if (rows == null)
                    {
                        rows = Array.Empty<object>();
                    }

                    progress?.Report(new ProgressInfo(5, "已載入 BOM 列，正在解析（0～100%）…"));

                    var items = BuildBomItemsFromComputedRows(
                        assemblyPath,
                        rows,
                        columns,
                        cardVarNames,
                        includeRootBomItem,
                        maxBomLayerDepth,
                        configResolved,
                        progress,
                        progressWhileParsingMin: 5,
                        progressWhileParsingMax: 100,
                        cancellationToken);

                    items = EnsureRootRowForIncludeMode(
                        items,
                        assemblyPath,
                        file,
                        cardVarNames,
                        includeRootBomItem,
                        configResolved);

                    NormalizeBomLevelPresentation(items, assemblyPath, includeRootBomItem);

                    progress?.Report(new ProgressInfo(100, $"解析完成，共 {items.Count} 筆。"));
                    return (IReadOnlyList<BomItem>)items;
                }
                finally
                {
                    ComHelper.Release(bomView);
                    ComHelper.Release(bomMgr);
                    ComHelper.Release(folder);
                    ComHelper.Release(file);
                }
            }, cancellationToken);
        }

        /// <summary>
        /// 列舉指定組合件檔案中的 SolidWorks 組態名稱（供 UI 下拉）；失敗時回傳空清單。
        /// </summary>
        public async Task<IReadOnlyList<string>> GetAssemblyConfigurationsAsync(string assemblyPath)
        {
            return await Task.Run(() =>
            {
                if (string.IsNullOrWhiteSpace(assemblyPath) || !File.Exists(assemblyPath))
                {
                    return (IReadOnlyList<string>)Array.Empty<string>();
                }

                if (!assemblyPath.StartsWith(VaultRootPath, StringComparison.OrdinalIgnoreCase))
                {
                    return (IReadOnlyList<string>)Array.Empty<string>();
                }

                EnsureVaultLogin();
                IEdmFolder5 folder = null;
                IEdmFile5 file = null;
                try
                {
                    file = _vault.GetFileFromPath(assemblyPath, out folder);
                    if (file == null)
                    {
                        LastDocumentActiveConfiguration = string.Empty;
                        return (IReadOnlyList<string>)Array.Empty<string>();
                    }

                    LastDocumentActiveConfiguration = TryGetDocumentActiveConfigurationName(file) ?? string.Empty;

                    var file7 = file as IEdmFile7;
                    if (file7 == null)
                    {
                        LastConfigurationEnumerationDiag = "無法將檔案物件轉型為 IEdmFile7，PDM Interop 版本可能過舊。";
                        return (IReadOnlyList<string>)Array.Empty<string>();
                    }

                    var diag = new StringBuilder();
                    var latestVer = GetLatestVaultVersionNumber(file);
                    var names = InvokeGetConfigurations(file7, file, folder, latestVer, diag);
                    LastConfigurationEnumerationDiag = diag.ToString();
                    return names.Count > 0
                        ? (IReadOnlyList<string>)names
                        : (IReadOnlyList<string>)Array.Empty<string>();
                }
                finally
                {
                    ComHelper.Release(folder);
                    ComHelper.Release(file);
                }
            });
        }

        private static EdmBomLayout SelectPrimaryBomLayout(EdmBomLayout[] layouts)
        {
            foreach (var lo in layouts)
            {
                var n = lo.mbsLayoutName ?? string.Empty;
                if (n.IndexOf("solidworks", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    n.IndexOf("經計算", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    n.IndexOf("計算", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return lo;
                }
            }

            return layouts[0];
        }

        /// <summary>
        /// GetComputedBOM 第一參數為 layout id 或版面名稱；部分環境需明確以 object 傳遞或改用名稱字串。
        /// </summary>
        private static IEdmBomView TryOpenComputedBomView(
            IEdmFile7 file7,
            IEdmFile5 file,
            EdmBomLayout layout,
            string configurationName)
        {
            var ver = GetLatestVaultVersionNumber(file);
            const int bomFlags = 0;

            IEdmBomView view = null;
            try
            {
                view = file7.GetComputedBOM((object)layout.mlLayoutID, ver, configurationName, bomFlags) as IEdmBomView;
            }
            catch
            {
                view = null;
            }

            if (view != null)
            {
                return view;
            }

            var name = layout.mbsLayoutName?.Trim();
            if (string.IsNullOrEmpty(name))
            {
                return null;
            }

            try
            {
                return file7.GetComputedBOM(name, ver, configurationName, bomFlags) as IEdmBomView;
            }
            catch
            {
                return null;
            }
        }

        private static string ResolveConfigurationName(
            IEdmFile7 file7,
            IEdmFile5 file,
            IEdmFolder5 folder,
            string requested)
        {
            var latestVer = GetLatestVaultVersionNumber(file);
            var available = InvokeGetConfigurations(file7, file, folder, latestVer, null);
            var req = requested?.Trim() ?? string.Empty;

            if (!string.IsNullOrEmpty(req) &&
                available.Any(a => string.Equals(a, req, StringComparison.OrdinalIgnoreCase)))
            {
                return available.First(a => string.Equals(a, req, StringComparison.OrdinalIgnoreCase));
            }

            var active = TryGetDocumentActiveConfigurationName(file);
            if (!string.IsNullOrWhiteSpace(active))
            {
                var activeHit = available.FirstOrDefault(a =>
                    string.Equals(a, active.Trim(), StringComparison.OrdinalIgnoreCase));
                if (!string.IsNullOrEmpty(activeHit))
                {
                    return activeHit;
                }
            }

            foreach (var cand in new[] { "預設", "Default", "默认" })
            {
                var hit = available.FirstOrDefault(a => string.Equals(a, cand, StringComparison.OrdinalIgnoreCase));
                if (!string.IsNullOrEmpty(hit))
                {
                    return hit;
                }
            }

            if (available.Count > 0)
            {
                return available[0];
            }

            return string.IsNullOrEmpty(req) ? "Default" : req;
        }

        /// <summary>Vault 資料庫中該檔案的目前版號（最新版），供 GetComputedBOM／GetConfigurations 使用。</summary>
        private static int GetLatestVaultVersionNumber(IEdmFile5 file)
        {
            if (file == null)
            {
                throw new ArgumentNullException(nameof(file));
            }

            return file.CurrentVersion;
        }

        /// <summary>
        /// 嘗試從 PDM 檔案 COM 物件讀取 SolidWorks「目前作用中組態」名稱（各版 Interop 屬性／方法名稱不一，故以反射嘗試）。
        /// </summary>
        private static string TryGetDocumentActiveConfigurationName(IEdmFile5 file)
        {
            if (file == null)
            {
                return null;
            }

            var asm = typeof(IEdmFile5).Assembly;
            var ifaceOrder = new[] { "IEdmFile9", "IEdmFile8", "IEdmFile7", "IEdmFile6", "IEdmFile5" };
            foreach (var ifaceName in ifaceOrder)
            {
                var t = asm.GetType("EPDM.Interop.epdm." + ifaceName, false, false);
                if (t == null)
                {
                    continue;
                }

                foreach (var propName in new[]
                         {
                             "ActiveConfiguration",
                             "ActiveConfigurationName",
                             "ActiveConfigName",
                             "DocumentActiveConfiguration"
                         })
                {
                    try
                    {
                        var p = t.GetProperty(propName, BindingFlags.Public | BindingFlags.Instance);
                        if (p == null)
                        {
                            continue;
                        }

                        var v = p.GetValue(file);
                        var s = v?.ToString()?.Trim();
                        if (!string.IsNullOrEmpty(s))
                        {
                            return s;
                        }
                    }
                    catch
                    {
                        // 下一個候選
                    }
                }

                foreach (var methodName in new[] { "GetActiveConfigurationName", "GetActiveConfiguration" })
                {
                    try
                    {
                        var methods = t.GetMethods(BindingFlags.Public | BindingFlags.Instance)
                            .Where(m => string.Equals(m.Name, methodName, StringComparison.Ordinal)
                                        && m.GetParameters().Length == 0);
                        foreach (var m in methods)
                        {
                            var v = m.Invoke(file, null);
                            var s = v?.ToString()?.Trim();
                            if (!string.IsNullOrEmpty(s))
                            {
                                return s;
                            }
                        }
                    }
                    catch
                    {
                        // 下一個候選
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// PDM 各版 interop 中 GetConfigurations 簽章不一。COM RCW 多為 __ComObject，必須從 <see cref="IEdmFile7"/> 等介面取得 MethodInfo 再 Invoke。
        /// </summary>
        private static List<string> InvokeGetConfigurations(
            IEdmFile7 file7,
            IEdmFile5 file,
            IEdmFolder5 folder,
            int latestVaultVersion,
            StringBuilder diag)
        {
            var result = new List<string>();
            var folderId = 0;
            try
            {
                folderId = folder?.ID ?? 0;
            }
            catch
            {
                folderId = 0;
            }

            var log = diag ?? new StringBuilder();

            if (TryGetConfigurationsDirect(file7, file, latestVaultVersion, result, log))
            {
                return result;
            }

            foreach (var comTarget in new object[] { file7, file })
            {
                if (comTarget == null)
                {
                    continue;
                }

                TryInvokeGetConfigurationsReflection(comTarget, folderId, latestVaultVersion, result, log);
                if (result.Count > 0)
                {
                    return result;
                }
            }

            return result;
        }

        /// <summary>
        /// 以編譯期介面呼叫 GetConfigurations，避免 MethodInfo.Invoke 對 COM ref object 封送失敗（常見為 TargetInvocationException）。
        /// </summary>
        private static bool TryGetConfigurationsDirect(
            IEdmFile7 file7,
            IEdmFile5 file,
            int latestVaultVersion,
            List<string> result,
            StringBuilder log)
        {
            if (file7 == null)
            {
                return false;
            }

            log.AppendLine("── IEdmFile7.GetConfigurations 直接呼叫（ref object），版次＝Vault 最新（" + latestVaultVersion + "）──");
            string rev = null;
            try
            {
                rev = file?.CurrentRevision;
            }
            catch
            {
                // ignore
            }

            foreach (var cand in EnumerateGetConfigurationsRefObjectCandidates(latestVaultVersion, rev))
            {
                try
                {
                    object arg = cand;
                    var ret = file7.GetConfigurations(ref arg);
                    var before = result.Count;
                    // 部分 PDM 版次：清單在回傳值；部分則寫回 ref 參數（或兩者皆有）— 須併採。
                    TryConsumeReturnValueForConfigurationList(result, ret);
                    TryAddConfigNames(result, arg);
                    if (result.Count > before)
                    {
                        log.AppendLine("  → 成功，ref 輸入：" + FormatArgForDiag(cand));
                        return true;
                    }

                    if (ret != null || arg != cand)
                    {
                        log.AppendLine(
                            "  → ref=" + FormatArgForDiag(cand) + " 呼叫未丟例外但解析為 0 筆（回傳型別："
                            + (ret?.GetType().FullName ?? "null") + "；ref 型別：" + (arg?.GetType().FullName ?? "null") + "）");
                    }
                }
                catch (Exception ex)
                {
                    log.AppendLine("  → 失敗 ref=" + FormatArgForDiag(cand) + "：" + FormatExceptionChain(ex));
                }
            }

            return false;
        }

        /// <summary>
        /// ref 參數候選順序：先 null（多數環境代表「目前／預設版次語意」），再 Vault 最新版號、修訂字串。
        /// 不傳本機路徑，避免誤當修訂而 HRESULT 0x80040224。
        /// </summary>
        private static IEnumerable<object> EnumerateGetConfigurationsRefObjectCandidates(
            int latestVaultVersion,
            string currentRevision)
        {
            yield return null;
            yield return (object)latestVaultVersion;
            if (!string.IsNullOrWhiteSpace(currentRevision))
            {
                yield return (object)currentRevision.Trim();
            }
        }

        private static string FormatArgForDiag(object o) => o == null ? "null" : o.ToString();

        private static string FormatExceptionChain(Exception ex)
        {
            if (ex == null)
            {
                return string.Empty;
            }

            var sb = new StringBuilder();
            var depth = 0;
            for (var e = ex; e != null && depth < 8; e = e.InnerException, depth++)
            {
                sb.AppendLine($"    [{depth}] {e.GetType().Name}: {e.Message}");
                if (e is System.Runtime.InteropServices.COMException comEx)
                {
                    sb.AppendLine($"         HRESULT=0x{(uint)comEx.ErrorCode:X8}");
                }
            }

            return sb.ToString().TrimEnd();
        }

        private static void TryInvokeGetConfigurationsReflection(
            object comTarget,
            int folderId,
            int latestVaultVersion,
            List<string> result,
            StringBuilder diag)
        {
            var log = diag ?? new StringBuilder();
            log.AppendLine("── GetConfigurations 列舉診斷 ──");

            var ifaceTypes = new List<Type> { typeof(IEdmFile7), typeof(IEdmFile5) };
            try
            {
                var asm = typeof(IEdmFile5).Assembly;
                foreach (var extra in new[] { "IEdmFile6", "IEdmFile8", "IEdmFile9" })
                {
                    var tExtra = asm.GetType("EPDM.Interop.epdm." + extra, throwOnError: false, ignoreCase: false);
                    if (tExtra != null)
                    {
                        ifaceTypes.Add(tExtra);
                    }
                }
            }
            catch
            {
                // ignore
            }

            var seenSig = new HashSet<string>(StringComparer.Ordinal);
            foreach (var it in ifaceTypes.Distinct())
            {
                MethodInfo[] methods;
                try
                {
                    methods = it.GetMethods(BindingFlags.Instance | BindingFlags.Public);
                }
                catch
                {
                    continue;
                }

                foreach (var m in methods.Where(x => string.Equals(x.Name, "GetConfigurations", StringComparison.Ordinal)))
                {
                    var sig = (m.DeclaringType?.FullName ?? "?") + "." + m.Name + "("
                              + string.Join(", ", m.GetParameters().Select(p => p.ParameterType.Name + " " + p.Name)) + ")"
                              + " -> " + m.ReturnType.Name;
                    if (!seenSig.Add(sig))
                    {
                        continue;
                    }

                    log.AppendLine(sig);
                    TryInvokeSingleGetConfigurationsMethod(m, comTarget, folderId, latestVaultVersion, result, log);
                    if (result.Count > 0)
                    {
                        return;
                    }
                }
            }

            // 備援：執行個體 CLR 型別（非 __ComObject 時）
            try
            {
                var rt = comTarget.GetType();
                if (rt.FullName != "System.__ComObject")
                {
                    foreach (var m in rt.GetMethods(BindingFlags.Instance | BindingFlags.Public)
                        .Where(x => string.Equals(x.Name, "GetConfigurations", StringComparison.Ordinal)))
                    {
                        log.AppendLine("(runtime) " + m);
                        TryInvokeSingleGetConfigurationsMethod(m, comTarget, folderId, latestVaultVersion, result, log);
                        if (result.Count > 0)
                        {
                            return;
                        }
                    }
                }
            }
            catch
            {
                // ignore
            }

            log.AppendLine("（結束：未取得任何組態名稱）");
        }

        private static void TryInvokeSingleGetConfigurationsMethod(
            MethodInfo m,
            object comTarget,
            int folderId,
            int latestVaultVersion,
            List<string> result,
            StringBuilder diag)
        {
            string revisionName = null;
            try
            {
                revisionName = (comTarget as IEdmFile5)?.CurrentRevision;
            }
            catch
            {
                // ignore
            }

            var before = result.Count;
            foreach (var args in BuildGetConfigurationsArgumentVariants(
                m.GetParameters(),
                folderId,
                latestVaultVersion,
                revisionName))
            {
                if (args.Length != m.GetParameters().Length)
                {
                    continue;
                }

                try
                {
                    var ret = m.Invoke(comTarget, args);

                    // 回傳值與 ref/out 皆可能帶 EdmStrLst5／字串（不可只採一邊）。
                    if (m.ReturnType != typeof(void) && ret != null)
                    {
                        TryConsumeReturnValueForConfigurationList(result, ret);
                    }

                    foreach (var p in m.GetParameters().Select((param, idx) => (param, idx)))
                    {
                        if (p.param.ParameterType.IsByRef || p.param.IsOut)
                        {
                            TryAddConfigNames(result, args[p.idx]);
                        }
                    }

                    if (result.Count > before)
                    {
                        diag.AppendLine("  → 成功，參數組：" + FormatArgsForDiag(args));
                        return;
                    }
                }
                catch (Exception ex)
                {
                    diag.AppendLine("  → 失敗 " + FormatArgsForDiag(args) + "：" + FormatExceptionChain(ex));
                }
            }
        }

        private static string FormatArgsForDiag(object[] args)
        {
            if (args == null || args.Length == 0)
            {
                return "()";
            }

            return "(" + string.Join(", ", args.Select(a => a == null ? "null" : a.ToString())) + ")";
        }

        private static IEnumerable<object[]> BuildGetConfigurationsArgumentVariants(
            ParameterInfo[] ps,
            int folderId,
            int latestVaultVersion,
            string currentRevision)
        {
            var n = ps.Length;
            if (n == 0)
            {
                yield return Array.Empty<object>();
                yield break;
            }

            var isRef = ps.Select(p => p.ParameterType.IsByRef).ToArray();

            bool IsRefObj(int i)
            {
                if (!isRef[i])
                {
                    return false;
                }

                var et = ps[i].ParameterType.GetElementType();
                return et == null || et == typeof(object);
            }

            bool IsRefInt(int i) =>
                isRef[i] && ps[i].ParameterType.GetElementType() == typeof(int);

            var cv = latestVaultVersion;

            // (out/ref object) — 以 Vault 最新版號／修訂／null 嘗試，回傳 EdmStrLst5
            if (n == 1 && isRef[0] && IsRefObj(0))
            {
                foreach (var boxed in EnumerateGetConfigurationsRefObjectCandidates(cv, currentRevision))
                {
                    yield return new object[] { boxed };
                }

                yield break;
            }

            // (int, out/ref object) — 第一參數為版次時僅試最新版
            if (n == 2 && !isRef[0] && ps[0].ParameterType == typeof(int) && isRef[1] && IsRefObj(1))
            {
                yield return new object[] { cv, null };
                yield break;
            }

            // (int folder, int ver, out/ref object)
            if (n == 3 && !isRef[0] && !isRef[1] && isRef[2] && IsRefObj(2)
                && ps[0].ParameterType == typeof(int) && ps[1].ParameterType == typeof(int))
            {
                yield return new object[] { folderId, cv, null };
                yield break;
            }

            // (int) 回傳值
            if (n == 1 && !isRef[0] && ps[0].ParameterType == typeof(int))
            {
                yield return new object[] { cv };
                yield break;
            }

            // (object) 回傳值（少見）
            if (n == 1 && !isRef[0] && ps[0].ParameterType == typeof(object))
            {
                yield return new object[] { cv };
                yield break;
            }

            // (int, int) 回傳值
            if (n == 2 && !isRef[0] && !isRef[1] && ps[0].ParameterType == typeof(int) && ps[1].ParameterType == typeof(int))
            {
                yield return new object[] { folderId, cv };
                yield break;
            }

            // 泛用：僅 ref/out，object 填 null、int 填 0
            if (isRef.All(x => x))
            {
                var a = new object[n];
                for (var i = 0; i < n; i++)
                {
                    if (IsRefInt(i))
                    {
                        a[i] = 0;
                    }
                    else
                    {
                        a[i] = null;
                    }
                }

                yield return a;
            }
        }

        private static void TryConsumeReturnValueForConfigurationList(List<string> result, object ret)
        {
            if (ret == null)
            {
                return;
            }

            if (LooksLikeEdmStringList(ret))
            {
                TryAddStringsFromEdmStringList(result, ret);
                return;
            }

            TryAddConfigNames(result, ret);
        }

        /// <summary>PDM 的 EdmStrLst5 在執行期可能是 __ComObject，不能只看型別名稱。</summary>
        private static bool LooksLikeEdmStringList(object o)
        {
            if (o == null)
            {
                return false;
            }

            var t = o.GetType();
            var fn = t.FullName ?? string.Empty;
            var nm = t.Name ?? string.Empty;
            if (fn.IndexOf("EdmStrLst", StringComparison.OrdinalIgnoreCase) >= 0 ||
                nm.IndexOf("EdmStrLst", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return true;
            }

            return t.GetMethod("GetHeadPosition", BindingFlags.Instance | BindingFlags.Public) != null
                   && t.GetMethods(BindingFlags.Instance | BindingFlags.Public).Any(x => x.Name == "GetNext");
        }

        /// <summary>
        /// EdmStrLst5Class 實作 <see cref="IEdmStrLst5"/> 等介面；公開方法在介面上，對 <c>GetType()</c> 取得的類別做
        /// <c>GetMethod("GetHeadPosition")</c> 常為 null，必須以介面型別反射再 <c>Invoke(strLst, …)</c>。
        /// </summary>
        private static void TryAddStringsFromEdmStringListViaStrLstInterfaces(List<string> result, object strLst)
        {
            var asm = typeof(IEdmFile5).Assembly;
            foreach (var ifaceName in new[]
                     {
                         "IEdmStrLst9", "IEdmStrLst8", "IEdmStrLst7", "IEdmStrLst6", "IEdmStrLst5"
                     })
            {
                Type iface;
                try
                {
                    iface = asm.GetType("EPDM.Interop.epdm." + ifaceName, false, false);
                }
                catch
                {
                    continue;
                }

                if (iface == null || !iface.IsInstanceOfType(strLst))
                {
                    continue;
                }

                if (TryEnumerateIEdmStrLstHeadNext(result, strLst, iface))
                {
                    return;
                }

                if (TryEnumerateIEdmStrLstByCountAndItem(result, strLst, iface))
                {
                    return;
                }
            }
        }

        private static bool TryEnumerateIEdmStrLstHeadNext(List<string> result, object strLst, Type iface)
        {
            try
            {
                var getHead = iface.GetMethod(
                    "GetHeadPosition",
                    BindingFlags.Instance | BindingFlags.Public,
                    binder: null,
                    types: Type.EmptyTypes,
                    modifiers: null);
                if (getHead == null)
                {
                    return false;
                }

                var getNextMethods = iface.GetMethods(BindingFlags.Instance | BindingFlags.Public)
                    .Where(m => string.Equals(m.Name, "GetNext", StringComparison.Ordinal))
                    .ToList();
                foreach (var getNext in getNextMethods)
                {
                    if (TryInvokeStrLstGetNextLoop(result, strLst, getHead, getNext, getNext.GetParameters()))
                    {
                        return true;
                    }
                }
            }
            catch
            {
                // ignore
            }

            return false;
        }

        private static bool TryInvokeStrLstGetNextLoop(
            List<string> result,
            object strLst,
            MethodInfo getHead,
            MethodInfo getNext,
            ParameterInfo[] ps)
        {
            IEdmPos5 posArg = null;
            try
            {
                posArg = getHead.Invoke(strLst, null) as IEdmPos5;
                if (posArg == null || posArg.IsNull)
                {
                    return false;
                }

                // string GetNext(ref IEdmPos5 pos) — Interop 有時未標 IsByRef，單參數即嘗試
                if (ps.Length == 1)
                {
                    var safety = 4096;
                    while (posArg != null && !posArg.IsNull && safety-- > 0)
                    {
                        var argsN = new object[] { posArg };
                        object sObj;
                        try
                        {
                            sObj = getNext.Invoke(strLst, argsN);
                        }
                        catch
                        {
                            break;
                        }

                        var nextPos = argsN[0] as IEdmPos5;
                        var s = sObj?.ToString()?.Trim();
                        if (string.IsNullOrWhiteSpace(s))
                        {
                            if (nextPos == null || ReferenceEquals(posArg, nextPos) ||
                                (nextPos is IEdmPos5 np && np.IsNull))
                            {
                                break;
                            }
                        }
                        else
                        {
                            result.Add(s);
                        }

                        if (!ReferenceEquals(posArg, nextPos))
                        {
                            ComHelper.Release(posArg);
                        }

                        posArg = nextPos;
                    }

                    return result.Count > 0;
                }

                // void GetNext(ref IEdmPos5 pos, out string psString) 等
                if (ps.Length >= 2 && ps[1].IsOut)
                {
                    var safety = 4096;
                    while (posArg != null && !posArg.IsNull && safety-- > 0)
                    {
                        var argsN = new object[] { posArg, null };
                        try
                        {
                            getNext.Invoke(strLst, argsN);
                        }
                        catch
                        {
                            break;
                        }

                        var nextPos = argsN[0] as IEdmPos5;
                        var s = argsN[1]?.ToString()?.Trim();
                        if (string.IsNullOrWhiteSpace(s))
                        {
                            if (nextPos == null || ReferenceEquals(posArg, nextPos) ||
                                (nextPos is IEdmPos5 np && np.IsNull))
                            {
                                break;
                            }
                        }
                        else
                        {
                            result.Add(s);
                        }

                        if (!ReferenceEquals(posArg, nextPos))
                        {
                            ComHelper.Release(posArg);
                        }

                        posArg = nextPos;
                    }

                    return result.Count > 0;
                }
            }
            finally
            {
                ComHelper.Release(posArg);
            }

            return false;
        }

        private static bool TryEnumerateIEdmStrLstByCountAndItem(List<string> result, object strLst, Type iface)
        {
            try
            {
                var countProp = iface.GetProperty("Count", BindingFlags.Instance | BindingFlags.Public);
                if (countProp == null)
                {
                    return false;
                }

                var cntObj = countProp.GetValue(strLst, null);
                if (cntObj == null || !int.TryParse(cntObj.ToString(), out var cnt) || cnt <= 0)
                {
                    return false;
                }

                foreach (var propName in new[] { "Item", "Str", "String" })
                {
                    var indexer = iface.GetProperties(BindingFlags.Instance | BindingFlags.Public)
                        .FirstOrDefault(p =>
                            p.Name == propName
                            && p.GetIndexParameters().Length == 1
                            && p.GetIndexParameters()[0].ParameterType == typeof(int));
                    if (indexer == null)
                    {
                        continue;
                    }

                    for (var i = 0; i < cnt; i++)
                    {
                        var v = indexer.GetValue(strLst, new object[] { i });
                        var s = v?.ToString()?.Trim();
                        if (!string.IsNullOrWhiteSpace(s))
                        {
                            result.Add(s);
                        }
                    }

                    if (result.Count > 0)
                    {
                        return true;
                    }
                }

                foreach (var methodName in new[] { "GetAt", "GetStr" })
                {
                    var gm = iface.GetMethods(BindingFlags.Instance | BindingFlags.Public)
                        .FirstOrDefault(m => m.Name == methodName
                                             && m.GetParameters().Length == 1
                                             && m.GetParameters()[0].ParameterType == typeof(int));
                    if (gm == null)
                    {
                        continue;
                    }

                    for (var i = 0; i < cnt; i++)
                    {
                        var v = gm.Invoke(strLst, new object[] { i });
                        var s = v?.ToString()?.Trim();
                        if (!string.IsNullOrWhiteSpace(s))
                        {
                            result.Add(s);
                        }
                    }

                    if (result.Count > 0)
                    {
                        return true;
                    }
                }
            }
            catch
            {
                // ignore
            }

            return false;
        }

        /// <summary>COM 上部分成員僅能透過執行期繫結呼叫。</summary>
        private static void TryAddStringsFromEdmStringListDynamic(List<string> result, object strLst)
        {
            try
            {
                dynamic d = strLst;
                dynamic pos = d.GetHeadPosition();
                if (pos == null)
                {
                    return;
                }

                try
                {
                    if ((bool)pos.IsNull)
                    {
                        return;
                    }
                }
                catch
                {
                    // 無 IsNull
                }

                var safety = 4096;
                while (pos != null && safety-- > 0)
                {
                    try
                    {
                        if ((bool)pos.IsNull)
                        {
                            break;
                        }
                    }
                    catch
                    {
                        // ignore
                    }

                    string s;
                    try
                    {
                        s = d.GetNext(ref pos);
                    }
                    catch
                    {
                        break;
                    }

                    s = s?.Trim();
                    if (string.IsNullOrWhiteSpace(s))
                    {
                        break;
                    }

                    result.Add(s);
                }
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
            {
                // 動態派發不支援此 COM 簽章
            }
            catch (System.Reflection.TargetInvocationException)
            {
                // ignore
            }
            catch
            {
                // ignore
            }
        }

        /// <summary>自 PDM EdmStrLst5 / IEdmStrLst* 取出字串（反射，避免 interop 版次差異）。</summary>
        private static void TryAddStringsFromEdmStringList(List<string> result, object strLst)
        {
            if (strLst == null)
            {
                return;
            }

            // EdmStrLst5Class 等方法定義在 IEdmStrLst* 上，對執行個體 CLR 型別 GetMethod 常拿不到，須先走介面反射。
            TryAddStringsFromEdmStringListViaStrLstInterfaces(result, strLst);
            if (result.Count > 0)
            {
                return;
            }

            TryAddStringsFromEdmStringListDynamic(result, strLst);
            if (result.Count > 0)
            {
                return;
            }

            var t = strLst.GetType();

            // 1) GetHeadPosition + GetNext(ref pos) — 執行個體型別上若可見則直接呼叫
            try
            {
                var getHead = t.GetMethod(
                    "GetHeadPosition",
                    BindingFlags.Instance | BindingFlags.Public,
                    binder: null,
                    types: Type.EmptyTypes,
                    modifiers: null);
                var getNextMethods = t.GetMethods(BindingFlags.Instance | BindingFlags.Public)
                    .Where(m => m.Name == "GetNext")
                    .ToList();
                if (getHead != null && getNextMethods.Count > 0)
                {
                    foreach (var getNext in getNextMethods)
                    {
                        var ps = getNext.GetParameters();
                        if (ps.Length != 1 || !ps[0].ParameterType.IsByRef)
                        {
                            continue;
                        }

                        IEdmPos5 posArg = null;
                        try
                        {
                            posArg = getHead.Invoke(strLst, null) as IEdmPos5;
                            if (posArg == null || posArg.IsNull)
                            {
                                continue;
                            }

                            var safety = 4096;
                            while (posArg != null && !posArg.IsNull && safety-- > 0)
                            {
                                var argsN = new object[] { posArg };
                                object sObj;
                                try
                                {
                                    sObj = getNext.Invoke(strLst, argsN);
                                }
                                catch
                                {
                                    break;
                                }

                                var nextPos = argsN[0] as IEdmPos5;
                                var s = sObj?.ToString()?.Trim();
                                if (string.IsNullOrWhiteSpace(s))
                                {
                                    if (nextPos == null || ReferenceEquals(posArg, nextPos) ||
                                        (nextPos is IEdmPos5 np && np.IsNull))
                                    {
                                        break;
                                    }
                                }
                                else
                                {
                                    result.Add(s);
                                }

                                if (!ReferenceEquals(posArg, nextPos))
                                {
                                    ComHelper.Release(posArg);
                                }

                                posArg = nextPos;
                            }
                        }
                        finally
                        {
                            ComHelper.Release(posArg);
                        }

                        if (result.Count > 0)
                        {
                            return;
                        }
                    }
                }
            }
            catch
            {
                // try next pattern
            }

            // 2) Count／MlCount／GetCount + 索引子或 GetAt
            try
            {
                int cnt = -1;
                foreach (var propName in new[] { "Count", "MlCount", "mbsCount" })
                {
                    var countProp = t.GetProperty(propName, BindingFlags.Instance | BindingFlags.Public);
                    if (countProp == null)
                    {
                        continue;
                    }

                    var cntObj = countProp.GetValue(strLst, null);
                    if (cntObj != null && int.TryParse(cntObj.ToString(), out var c) && c > 0)
                    {
                        cnt = c;
                        break;
                    }
                }

                if (cnt < 0)
                {
                    var getCount = t.GetMethods(BindingFlags.Instance | BindingFlags.Public)
                        .FirstOrDefault(m => m.Name == "GetCount" && m.GetParameters().Length == 0);
                    if (getCount != null)
                    {
                        var cntObj = getCount.Invoke(strLst, null);
                        if (cntObj != null && int.TryParse(cntObj.ToString(), out var c) && c > 0)
                        {
                            cnt = c;
                        }
                    }
                }

                if (cnt > 0)
                {
                    foreach (var indexer in t.GetProperties(BindingFlags.Instance | BindingFlags.Public)
                        .Where(p => p.Name == "Item" || p.Name == "Str" || p.Name == "String"))
                    {
                        var ip = indexer.GetIndexParameters();
                        if (ip.Length != 1 || ip[0].ParameterType != typeof(int))
                        {
                            continue;
                        }

                        for (var i = 0; i < cnt; i++)
                        {
                            var v = indexer.GetValue(strLst, new object[] { i });
                            var s = v?.ToString()?.Trim();
                            if (!string.IsNullOrWhiteSpace(s))
                            {
                                result.Add(s);
                            }
                        }

                        if (result.Count > 0)
                        {
                            return;
                        }
                    }

                    var getAt = t.GetMethods(BindingFlags.Instance | BindingFlags.Public)
                        .Where(m => (m.Name == "GetAt" || m.Name == "GetStr") && m.GetParameters().Length == 1
                                    && m.GetParameters()[0].ParameterType == typeof(int))
                        .FirstOrDefault();
                    if (getAt != null)
                    {
                        for (var i = 0; i < cnt; i++)
                        {
                            var v = getAt.Invoke(strLst, new object[] { i });
                            var s = v?.ToString()?.Trim();
                            if (!string.IsNullOrWhiteSpace(s))
                            {
                                result.Add(s);
                            }
                        }
                    }
                }
            }
            catch
            {
                // ignore
            }
        }

        private static void TryAddConfigNames(List<string> result, object raw, int depth = 0)
        {
            if (raw == null || depth > 8)
            {
                return;
            }

            if (LooksLikeEdmStringList(raw))
            {
                TryAddStringsFromEdmStringList(result, raw);
                if (result.Count > 0)
                {
                    return;
                }
            }

            if (raw is string single && !string.IsNullOrWhiteSpace(single))
            {
                result.Add(single.Trim());
                return;
            }

            if (raw is string[] sa)
            {
                foreach (var s in sa)
                {
                    if (!string.IsNullOrWhiteSpace(s))
                    {
                        result.Add(s.Trim());
                    }
                }

                return;
            }

            if (raw is object[] oa)
            {
                if (oa.Length == 1 && oa[0] != null)
                {
                    TryAddConfigNames(result, oa[0], depth + 1);
                    if (result.Count > 0)
                    {
                        return;
                    }
                }

                foreach (var o in oa)
                {
                    var s = o?.ToString();
                    if (!string.IsNullOrWhiteSpace(s))
                    {
                        result.Add(s.Trim());
                    }
                }

                return;
            }

            if (raw is System.Collections.IEnumerable en && raw is not string)
            {
                foreach (var o in en)
                {
                    if (o is string ss && !string.IsNullOrWhiteSpace(ss))
                    {
                        result.Add(ss.Trim());
                    }
                    else if (o != null)
                    {
                        var s = o.ToString();
                        if (!string.IsNullOrWhiteSpace(s))
                        {
                            result.Add(s.Trim());
                        }
                    }
                }
            }
        }

        private static string ReferencedAsForConfiguration(string fullPath, string configurationName)
        {
            var name = Path.GetFileName(fullPath ?? string.Empty);
            if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(configurationName))
            {
                return string.Empty;
            }

            return name + "<" + configurationName + ">";
        }

        /// <param name="progressWhileParsingMin">
        /// 解析迴圈進度下限（含）。與 <paramref name="progressWhileParsingMax"/> 構成「本次抓取」內的 0～100 對應區段。
        /// </param>
        /// <param name="progressWhileParsingMax">
        /// 解析迴圈進度上限（含）。最後一列會對應到此值，使整段解析在單次操作中為單調遞增的 0～100 體感。
        /// </param>
        private List<BomItem> BuildBomItemsFromComputedRows(
            string assemblyPath,
            object[] rows,
            EdmBomColumn[] columns,
            IReadOnlyList<string> cardVarNames,
            bool includeRootBomItem,
            int? maxBomLayerDepth,
            string configurationName,
            IProgress<ProgressInfo> progress,
            int progressWhileParsingMin,
            int progressWhileParsingMax,
            CancellationToken cancellationToken)
        {
            var items = new List<BomItem>();
            var levelCounters = new List<int>();
            var rootSkipped = false;
            var rootSkipTreeLevel = 0;
            int? skipDescendantsWhileTRawGreaterThan = null;

            for (var i = 0; i < rows.Length; i++)
            {
                try
                {
                cancellationToken.ThrowIfCancellationRequested();

                var cellObj = rows[i];
                var cell = cellObj as IEdmBomCell;
                if (cell == null)
                {
                    ComHelper.Release(cellObj);
                    continue;
                }

                try
                {
                    var path = SafeGetPathFromBomCell(cell, columns);
                    if (string.IsNullOrWhiteSpace(path))
                    {
                        continue;
                    }

                    var tRaw = SafeGetTreeLevel(cell);

                    if (!includeRootBomItem &&
                        !rootSkipped &&
                        IsSameAssemblyPath(path, assemblyPath))
                    {
                        rootSkipped = true;
                        rootSkipTreeLevel = tRaw;
                        continue;
                    }

                    var tAdj = rootSkipped ? tRaw - (rootSkipTreeLevel + 1) : tRaw;
                    if (tAdj < 0)
                    {
                        continue;
                    }

                    var levelStr = AdvanceBomLevelString(levelCounters, tAdj);

                    if (skipDescendantsWhileTRawGreaterThan != null &&
                        tRaw > skipDescendantsWhileTRawGreaterThan.Value)
                    {
                        continue;
                    }

                    skipDescendantsWhileTRawGreaterThan = null;

                    if (maxBomLayerDepth != null &&
                        CountLevelSegments(levelStr) > maxBomLayerDepth.Value)
                    {
                        skipDescendantsWhileTRawGreaterThan = tRaw;
                        continue;
                    }

                    IEdmFolder5 rowFolder = null;
                    IEdmFile5 rowFile = null;
                    try
                    {
                        rowFile = _vault.GetFileFromPath(path, out rowFolder);
                        if (rowFile == null || rowFolder == null)
                        {
                            continue;
                        }

                        var referencedAs = ReferencedAsForConfiguration(path, configurationName);
                        var item = BuildBomItem(levelStr, rowFile, path, referencedAs, cardVarNames);
                        item.UsageCount = TryReadBomQuantity(cell, columns);

                        if (includeRootBomItem &&
                            string.Equals(levelStr, "1", StringComparison.Ordinal) &&
                            IsSameAssemblyPath(path, assemblyPath))
                        {
                            item.UsageCount = null;
                        }

                        items.Add(item);
                    }
                    finally
                    {
                        ComHelper.Release(rowFile);
                        ComHelper.Release(rowFolder);
                    }
                }
                finally
                {
                    ComHelper.Release(cell);
                }
                }
                finally
                {
                    // 勿用 items.Count % N：每 N 筆會從上限跳回低百分比，進度條會「縮回去」。
                    // 以「列索引／總列數」線性對應到 [progressWhileParsingMin, progressWhileParsingMax]，單調遞增。
                    var denom = Math.Max(1, rows.Length);
                    var lo = Math.Max(0, Math.Min(100, progressWhileParsingMin));
                    var hi = Math.Max(lo, Math.Min(100, progressWhileParsingMax));
                    var span = hi - lo;
                    var frac = (double)(i + 1) / denom;
                    var scanPct = lo + (int)Math.Round(span * frac, MidpointRounding.AwayFromZero);
                    if (scanPct < lo)
                    {
                        scanPct = lo;
                    }

                    if (scanPct > hi)
                    {
                        scanPct = hi;
                    }

                    var msg = items.Count > 0
                        ? $"解析中（{i + 1}/{rows.Length}，{scanPct}%）：{items[items.Count - 1].FileName}"
                        : $"掃描 BOM…（{i + 1}/{rows.Length}，{scanPct}%）";
                    progress?.Report(new ProgressInfo(scanPct, msg));
                }
            }

            return items;
        }

        private static string SafeGetPathFromBomCell(IEdmBomCell cell, EdmBomColumn[] columns)
        {
            try
            {
                var p = cell.GetPathName();
                if (!string.IsNullOrWhiteSpace(p))
                {
                    return p.Trim();
                }
            }
            catch
            {
                // fall through
            }

            if (columns == null)
            {
                return string.Empty;
            }

            foreach (var col in columns)
            {
                if (col.meType != EdmBomColumnType.EdmBomCol_Path)
                {
                    continue;
                }

                try
                {
                    object val, compVal;
                    string cfg;
                    bool ro;
                    cell.GetVar(col.mlVariableID, col.meType, out val, out compVal, out cfg, out ro);
                    var s = val?.ToString() ?? compVal?.ToString() ?? string.Empty;
                    if (!string.IsNullOrWhiteSpace(s))
                    {
                        return s.Trim();
                    }
                }
                catch
                {
                    // try next
                }
            }

            return string.Empty;
        }

        private static int SafeGetTreeLevel(IEdmBomCell cell)
        {
            try
            {
                return cell.GetTreeLevel();
            }
            catch
            {
                return 0;
            }
        }

        private static bool IsSameAssemblyPath(string rowPath, string assemblyPath)
        {
            if (string.IsNullOrWhiteSpace(rowPath) || string.IsNullOrWhiteSpace(assemblyPath))
            {
                return false;
            }

            if (!rowPath.EndsWith(".sldasm", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            var a = rowPath.Trim().TrimEnd('\\');
            var b = assemblyPath.Trim().TrimEnd('\\');
            return string.Equals(a, b, StringComparison.OrdinalIgnoreCase);
        }

        private static string AdvanceBomLevelString(List<int> counters, int treeLevel)
        {
            while (counters.Count > treeLevel + 1)
            {
                counters.RemoveAt(counters.Count - 1);
            }

            while (counters.Count < treeLevel + 1)
            {
                counters.Add(0);
            }

            counters[treeLevel]++;
            return string.Join(".", counters.Take(treeLevel + 1));
        }

        private static int CountLevelSegments(string level)
        {
            if (string.IsNullOrWhiteSpace(level))
            {
                return 0;
            }

            return level.Split('.').Length;
        }

        /// <summary>
        /// 某些 PDM 經計算 BOM 版面不會回傳根列；在「定義A（含根）」時補上一列根組件，
        /// 並將既有列階層整體下移一層（1.x...），以維持層數語意一致。
        /// </summary>
        private List<BomItem> EnsureRootRowForIncludeMode(
            List<BomItem> items,
            string assemblyPath,
            IEdmFile5 rootFile,
            IReadOnlyList<string> cardVarNames,
            bool includeRootBomItem,
            string configurationName)
        {
            if (!includeRootBomItem)
            {
                return items;
            }

            var hasRoot = items.Any(i =>
                i != null &&
                IsSameAssemblyPath(i.FullPath, assemblyPath));
            if (hasRoot)
            {
                return items;
            }

            var normalized = new List<BomItem>(items.Count + 1);
            BomItem rootItem = null;

            if (rootFile != null)
            {
                var referencedAs = ReferencedAsForConfiguration(assemblyPath, configurationName);
                rootItem = BuildBomItem("1", rootFile, assemblyPath, referencedAs, cardVarNames);
            }
            else
            {
                rootItem = new BomItem
                {
                    Level = "1",
                    FileName = Path.GetFileName(assemblyPath),
                    FullPath = assemblyPath
                };
            }

            rootItem.Level = "0";
            rootItem.UsageCount = null;
            normalized.Add(rootItem);

            foreach (var item in items)
            {
                if (item == null || string.IsNullOrWhiteSpace(item.Level))
                {
                    continue;
                }
                normalized.Add(item);
            }

            return normalized;
        }

        /// <summary>
        /// 將 Level 正規化為展示規則：
        /// 定義A（含根）：0, 0.x, 0.x.x ...
        /// 定義B（不含根）：x, x.x, x.x.x ...（即 A 去掉最前面的「0.」並移除根列）。
        /// </summary>
        private static void NormalizeBomLevelPresentation(IList<BomItem> items, string assemblyPath, bool includeRootBomItem)
        {
            if (items == null || items.Count == 0)
            {
                return;
            }

            foreach (var item in items)
            {
                if (item == null)
                {
                    continue;
                }

                if (IsSameAssemblyPath(item.FullPath, assemblyPath))
                {
                    item.Level = "0";
                    item.UsageCount = null;
                    continue;
                }

                var lv = (item.Level ?? string.Empty).Trim();
                if (string.IsNullOrWhiteSpace(lv))
                {
                    continue;
                }

                if (!lv.Contains("."))
                {
                    item.Level = "0." + lv;
                    continue;
                }

                if (!lv.StartsWith("0.", StringComparison.Ordinal))
                {
                    item.Level = "0." + lv;
                }
            }

            if (includeRootBomItem)
            {
                return;
            }

            for (var i = items.Count - 1; i >= 0; i--)
            {
                var item = items[i];
                if (item == null)
                {
                    items.RemoveAt(i);
                    continue;
                }

                if (string.Equals(item.Level, "0", StringComparison.Ordinal))
                {
                    items.RemoveAt(i);
                    continue;
                }

                if (!string.IsNullOrWhiteSpace(item.Level) &&
                    item.Level.StartsWith("0.", StringComparison.Ordinal))
                {
                    item.Level = item.Level.Substring(2);
                }
            }
        }

        private static int? TryReadBomQuantity(IEdmBomCell cell, EdmBomColumn[] columns)
        {
            if (columns == null)
            {
                return null;
            }

            foreach (var colType in new[]
                     {
                         EdmBomColumnType.EdmBomCol_RefCount,
                         EdmBomColumnType.EdmBomCol_RefCountNoBomQty
                     })
            {
                foreach (var col in columns)
                {
                    if (col.meType != colType)
                    {
                        continue;
                    }

                    try
                    {
                        object val, compVal;
                        string cfg;
                        bool ro;
                        cell.GetVar(col.mlVariableID, col.meType, out val, out compVal, out cfg, out ro);
                        var s = val?.ToString() ?? compVal?.ToString() ?? string.Empty;
                        if (TryParseQuantityString(s, out var q))
                        {
                            return q;
                        }
                    }
                    catch
                    {
                        // try next column
                    }
                }
            }

            return null;
        }

        private static bool TryParseQuantityString(string text, out int quantity)
        {
            quantity = 0;
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            text = text.Trim();
            if (int.TryParse(text, System.Globalization.NumberStyles.Integer,
                    System.Globalization.CultureInfo.InvariantCulture, out quantity))
            {
                return true;
            }

            if (double.TryParse(text, System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture, out var d))
            {
                quantity = (int)Math.Round(d);
                return true;
            }

            return false;
        }

        private static bool IsAsmOrPart(string fullPath)
        {
            if (string.IsNullOrWhiteSpace(fullPath)) return false;
            var ext = Path.GetExtension(fullPath);
            return ext != null &&
                   (ext.Equals(".sldasm", StringComparison.OrdinalIgnoreCase) ||
                    ext.Equals(".sldprt", StringComparison.OrdinalIgnoreCase));
        }

        public async Task ExportToExcelAsync(
            IReadOnlyList<BomItem> items,
            IReadOnlyList<string> selectedFixedColumns,
            IReadOnlyList<string> cardVarNames,
            string outputPath,
            IProgress<ProgressInfo> progress,
            CancellationToken cancellationToken = default)
        {
            await Task.Run(() =>
            {
                if (items == null || items.Count == 0)
                    throw new InvalidOperationException("沒有可匯出的資料。");

                progress?.Report(new ProgressInfo(80, "開始輸出 Excel..."));

                // 固定欄：null = 全部顯示
                var fixedSet = selectedFixedColumns != null
                    ? new HashSet<string>(selectedFixedColumns, StringComparer.OrdinalIgnoreCase)
                    : null;
                bool ShowFixed(string n) => fixedSet == null || fixedSet.Contains(n);

                // (欄位標題, 取值 Func) 清單，依序建立
                var columns = new List<(string Header, Func<BomItem, string> Get)>();

                columns.Add(("Level", b => b.Level)); // Level 永遠顯示
                columns.Add(("Use Count", b => b.UsageCount.HasValue ? b.UsageCount.Value.ToString() : string.Empty));

                if (ShowFixed("File Name"))               columns.Add(("File Name",               b => b.FileName));
                if (ShowFixed("State"))                   columns.Add(("State",                   b => b.State));
                if (ShowFixed("Workflow State"))           columns.Add(("Workflow State",           b => b.WorkflowState));
                if (ShowFixed("Description"))             columns.Add(("Description",             b => b.Description));
                if (ShowFixed("Part Number"))             columns.Add(("Part Number",             b => b.PartNumber));
                if (ShowFixed("Referenced As"))           columns.Add(("Referenced As",           b => b.ReferencedAs));
                if (ShowFixed("Full Path"))               columns.Add(("Full Path",               b => b.FullPath));
                if (ShowFixed("Description Var Used"))    columns.Add(("Description Var Used",    b => b.DescriptionVarUsed));
                if (ShowFixed("Description Config Used")) columns.Add(("Description Config Used", b => b.DescriptionConfigUsed));
                if (ShowFixed("Part Number Var Used"))    columns.Add(("Part Number Var Used",    b => b.PartNumberVarUsed));
                if (ShowFixed("Part Number Config Used")) columns.Add(("Part Number Config Used", b => b.PartNumberConfigUsed));

                // 卡片變數欄（備援：內建清單）
                var exportVarNames = (cardVarNames != null && cardVarNames.Count > 0)
                    ? cardVarNames
                    : (IReadOnlyList<string>)CardVariableSpecs.Select(s => s.Label).ToList();

                foreach (var varName in exportVarNames)
                {
                    var captured = varName;
                    columns.Add(($"Card:{captured}", b =>
                    {
                        b.CardVariables.TryGetValue(captured, out var v);
                        return v ?? string.Empty;
                    }));
                }

                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("BOM");

                    // 標題列
                    for (var c = 0; c < columns.Count; c++)
                        ws.Cell(1, c + 1).Value = columns[c].Header;

                    // 資料列
                    for (var r = 0; r < items.Count; r++)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        var item = items[r];
                        for (var c = 0; c < columns.Count; c++)
                            ws.Cell(r + 2, c + 1).Value = columns[c].Get(item);

                        // 工程圖列以淡藍色底色區分
                        if (item.IsDrawing)
                            ws.Row(r + 2).Style.Fill.BackgroundColor = XLColor.FromHtml("#D6EEFF");
                    }

                    ws.Row(1).Style.Font.Bold = true;
                    ws.Columns().AdjustToContents();
                    workbook.SaveAs(outputPath);
                }

                progress?.Report(new ProgressInfo(100, $"匯出完成：{outputPath}"));
            }, cancellationToken);
        }

        // ════════════════════════════════════════════════════════════════════
        // 工程圖搜尋：在每個零組件/組合件之後插入同名 .SLDDRW 列
        // ════════════════════════════════════════════════════════════════════

        /// <summary>
        /// 針對 BOM 中每個 .SLDASM/.SLDPRT，以兩段式搜尋其同名工程圖，
        /// 並在其後插入工程圖 BomItem。同一工程圖檔若對應多個出現位置，
        /// 每個位置都會插入一列（Level 格式為「原Level-DRW」）。
        /// </summary>
        public async Task<IReadOnlyList<BomItem>> AppendDrawingItemsAsync(
            IReadOnlyList<BomItem> bomItems,
            IReadOnlyList<string> cardVarNames,
            IProgress<ProgressInfo> progress,
            CancellationToken cancellationToken = default)
        {
            return await Task.Run(() =>
            {
                EnsureVaultLogin();

                var result  = new List<BomItem>(bomItems.Count * 2);
                var total   = bomItems.Count;

                // 「顯示工程圖」：進度條獨立 0～100，與第一次抓取無關（不沿用 75～95 等舊區段）。
                progress?.Report(new ProgressInfo(0, "正在搜尋工程圖…"));

                if (total == 0)
                {
                    progress?.Report(new ProgressInfo(100, "無 BOM 列可處理。"));
                    return (IReadOnlyList<BomItem>)result;
                }

                // 快取：baseName → 找到的工程圖原型（Level 空白，後續 Clone 時填入）
                var foundCache    = new Dictionary<string, BomItem>(StringComparer.OrdinalIgnoreCase);
                var notFoundCache = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                for (var i = 0; i < total; i++)
                {
                    try
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        var item = bomItems[i];
                        result.Add(item);

                        if (item.IsDrawing)
                        {
                            continue;
                        }

                        var ext = Path.GetExtension(item.FileName);
                        if (!ext.Equals(".sldasm", StringComparison.OrdinalIgnoreCase) &&
                            !ext.Equals(".sldprt", StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        var baseName      = Path.GetFileNameWithoutExtension(item.FileName);
                        var drawingLevel  = item.Level + "-DRW";

                        if (notFoundCache.Contains(baseName))
                        {
                            continue;
                        }

                        BomItem drawingItem;
                        if (foundCache.TryGetValue(baseName, out var proto))
                        {
                            drawingItem = CloneDrawingItem(proto, drawingLevel);
                        }
                        else
                        {
                            var proto2 = FindDrawingProto(item, baseName + ".SLDDRW", cardVarNames);
                            if (proto2 != null)
                            {
                                foundCache[baseName] = proto2;
                                drawingItem = CloneDrawingItem(proto2, drawingLevel);
                            }
                            else
                            {
                                notFoundCache.Add(baseName);
                                continue;
                            }
                        }

                        result.Add(drawingItem);
                    }
                    finally
                    {
                        var pct = (int)((100L * (i + 1) + total - 1) / total);
                        if (pct > 100)
                        {
                            pct = 100;
                        }

                        var name = i < bomItems.Count ? bomItems[i].FileName : string.Empty;
                        progress?.Report(new ProgressInfo(
                            pct,
                            $"搜尋工程圖（{i + 1}/{total}，{pct}%）：{name}"));
                    }
                }

                progress?.Report(new ProgressInfo(100, "工程圖搜尋完成。"));
                return (IReadOnlyList<BomItem>)result;
            }, cancellationToken);
        }

        /// <summary>
        /// 兩段式搜尋工程圖，回傳不含 Level 的原型 BomItem（Level 為空）。
        /// Stage 1：同資料夾；Stage 2：IEdmSearch5 全庫搜尋。
        /// </summary>
        private BomItem FindDrawingProto(BomItem parentItem, string drawingFileName, IReadOnlyList<string> cardVarNames)
        {
            // ── Stage 1：同資料夾 ──────────────────────────────────────────
            if (!string.IsNullOrWhiteSpace(parentItem.FullPath))
            {
                var dir = Path.GetDirectoryName(parentItem.FullPath);
                if (!string.IsNullOrWhiteSpace(dir))
                {
                    var drawingPath = Path.Combine(dir, drawingFileName);
                    IEdmFolder5 folder = null;
                    IEdmFile5   file   = null;
                    try
                    {
                        file = _vault.GetFileFromPath(drawingPath, out folder);
                        if (file != null && folder != null)
                        {
                            var proto = BuildBomItem(string.Empty, file, drawingPath, string.Empty, cardVarNames);
                            proto.IsDrawing = true;
                            return proto;
                        }
                    }
                    catch { /* 找不到則往 Stage 2 */ }
                    finally
                    {
                        ComHelper.Release(file);
                        ComHelper.Release(folder);
                    }
                }
            }

            // ── Stage 2：全庫搜尋（IEdmSearch5）──────────────────────────
            return TryFindDrawingVaultWide(drawingFileName, cardVarNames);
        }

        /// <summary>
        /// 使用 IEdmSearch5 在整個 Vault 搜尋指定工程圖檔名。
        /// 找到後取第一個結果。若搜尋 API 不可用，回傳 null。
        /// </summary>
        private BomItem TryFindDrawingVaultWide(string drawingFileName, IReadOnlyList<string> cardVarNames)
        {
            object searchRaw = null;
            IEdmFile5 file = null;
            IEdmFolder5 folder = null;

            try
            {
                // 由於不同 EPDM Interop 版本中搜尋相關 API 名稱/簽名可能不同，
                // 這裡改用 reflection 嘗試建立「檔案搜尋」utility，並觸發搜尋。
                var vaultObj = (object)_vault;
                var vaultType = vaultObj.GetType();

                // 1) 找 CreateUtility(...)（不同版本可能只有具體類別才有）
                var createUtilMethods = vaultType
                    .GetMethods(BindingFlags.Instance | BindingFlags.Public)
                    .Where(m => m.Name.Equals("CreateUtility", StringComparison.OrdinalIgnoreCase))
                    .Where(m => m.GetParameters().Length == 1)
                    .ToList();

                if (createUtilMethods.Count == 0) return null;

                // 2) 從 EdmUtility enum 挑一個最像「檔案搜尋」的值
                var utilFields = typeof(EdmUtility).GetFields(BindingFlags.Public | BindingFlags.Static);
                var utilField =
                    utilFields.FirstOrDefault(f =>
                        f.Name.IndexOf("File", StringComparison.OrdinalIgnoreCase) >= 0 &&
                        f.Name.IndexOf("Search", StringComparison.OrdinalIgnoreCase) >= 0)
                    ?? utilFields.FirstOrDefault(f =>
                        f.Name.IndexOf("Search", StringComparison.OrdinalIgnoreCase) >= 0);

                if (utilField == null) return null;

                var utilValue = utilField.GetValue(null);

                // 3) 逐一嘗試呼叫，找到可用的 searchRaw
                foreach (var m in createUtilMethods)
                {
                    try
                    {
                        var p0 = m.GetParameters()[0].ParameterType;
                        var arg = utilValue;
                        if (utilValue != null && p0 != utilValue.GetType())
                        {
                            // 若需要 int/short 等轉型，盡可能轉成功
                            try { arg = Convert.ChangeType(utilValue, p0); } catch { arg = utilValue; }
                        }

                        searchRaw = m.Invoke(vaultObj, new object[] { arg });
                        if (searchRaw != null) break;
                    }
                    catch
                    {
                        // 嘗試下一個重載
                    }
                }

                if (searchRaw == null) return null;

                // 4) 取得 IEdmSearch5（若該版本不支援，就不強行）
                var search = searchRaw as IEdmSearch5;
                if (search == null) return null;

                // 檔名條件
                search.FileName = drawingFileName;

                // 5) 觸發 FindFiles：它在介面裡可能是方法，也可能是屬性（版本差異）
                try
                {
                    var st = searchRaw.GetType();
                    var mFind = st.GetMethod("FindFiles", BindingFlags.Instance | BindingFlags.Public);
                    if (mFind != null)
                    {
                        mFind.Invoke(search, null);
                    }
                    else
                    {
                        var pFind = st.GetProperty("FindFiles", BindingFlags.Instance | BindingFlags.Public);
                        if (pFind != null)
                        {
                            // 若是 bool/flag 類型，嘗試寫入 True；否則取值觸發副作用
                            if (pFind.CanWrite && pFind.PropertyType == typeof(bool))
                            {
                                pFind.SetValue(search, true, null);
                            }
                            else
                            {
                                pFind.GetValue(search, null);
                            }
                        }
                    }
                }
                catch
                {
                    // 找不到 FindFiles 也直接回傳 null
                    return null;
                }

                // 6) 走 GetFirstFilePosition / GetNextFile
                var searchType = searchRaw.GetType();
                var mGetFirstFilePos = searchType.GetMethod("GetFirstFilePosition",
                    BindingFlags.Instance | BindingFlags.Public);
                var mGetNextFile = searchType.GetMethod("GetNextFile",
                    BindingFlags.Instance | BindingFlags.Public);

                if (mGetFirstFilePos == null || mGetNextFile == null) return null;

                var posObj = mGetFirstFilePos.Invoke(search, null);
                var pos = posObj as IEdmPos5;
                if (pos == null || pos.IsNull) return null;

                var fileObj = mGetNextFile.Invoke(search, new object[] { pos });
                file = fileObj as IEdmFile5;
                if (file == null) return null;

                // 7) 取得第一個資料夾以建構本機路徑（避免不同 IEdmFile5 版本差異）
                var fileType = typeof(IEdmFile5);
                var mGetFirstFolderPos = fileType.GetMethod("GetFirstFolderPosition",
                    BindingFlags.Instance | BindingFlags.Public);
                var mGetNextFolder = fileType.GetMethod("GetNextFolder",
                    BindingFlags.Instance | BindingFlags.Public);

                if (mGetFirstFolderPos == null || mGetNextFolder == null) return null;

                var folderPosObj = mGetFirstFolderPos.Invoke(file, null);
                var folderPos = folderPosObj as IEdmPos5;
                if (folderPos == null || folderPos.IsNull) return null;

                var folderObj = mGetNextFolder.Invoke(file, new object[] { folderPos });
                folder = folderObj as IEdmFolder5;
                if (folder == null) return null;

                var localPath = file.GetLocalPath(folder.ID);
                if (string.IsNullOrWhiteSpace(localPath)) return null;

                var proto = BuildBomItem(string.Empty, file, localPath, string.Empty, cardVarNames);
                proto.IsDrawing = true;
                return proto;
            }
            catch
            {
                return null;
            }
            finally
            {
                ComHelper.Release(folder);
                ComHelper.Release(file);
                ComHelper.Release(searchRaw);
            }
        }

        /// <summary>根據原型複製一個工程圖 BomItem，並指定新的 Level。</summary>
        private static BomItem CloneDrawingItem(BomItem proto, string newLevel) =>
            new BomItem
            {
                IsDrawing             = true,
                Level                 = newLevel,
                FileName              = proto.FileName,
                FullPath              = proto.FullPath,
                State                 = proto.State,
                WorkflowState         = proto.WorkflowState,
                Description           = proto.Description,
                PartNumber            = proto.PartNumber,
                ReferencedAs          = proto.ReferencedAs,
                DescriptionVarUsed    = proto.DescriptionVarUsed,
                DescriptionConfigUsed = proto.DescriptionConfigUsed,
                PartNumberVarUsed     = proto.PartNumberVarUsed,
                PartNumberConfigUsed  = proto.PartNumberConfigUsed,
                CardVariables         = new Dictionary<string, string>(proto.CardVariables,
                                            StringComparer.OrdinalIgnoreCase)
            };

        /// <summary>
        /// 確保已登入 Vault，供其他服務（例如批次轉狀態）共用同一 COM 連線。
        /// </summary>
        public IEdmVault5 GetVault()
        {
            EnsureVaultLogin();
            return _vault;
        }

        private void EnsureVaultLogin()
        {
            if (_vault.IsLoggedIn)
                return;

            // ── 建立候選 Vault 名稱清單 ──────────────────────────────────────
            var candidates = new List<string>();

            // 優先使用 UI 明確選取的 Vault 名稱（若有設定）
            if (!string.IsNullOrWhiteSpace(OverrideVaultName))
                candidates.Add(OverrideVaultName);

            var resolved = ResolveVaultNameFromPathSafe(VaultRootPath);
            if (!string.IsNullOrWhiteSpace(resolved)
                && !candidates.Contains(resolved, StringComparer.OrdinalIgnoreCase))
                candidates.Add(resolved);

            var folderName = Path.GetFileName(VaultRootPath.TrimEnd('\\', '/'));
            if (!string.IsNullOrWhiteSpace(folderName)
                && !candidates.Contains(folderName, StringComparer.OrdinalIgnoreCase))
                candidates.Add(folderName);

            if (candidates.Count == 0)
                throw new InvalidOperationException(
                    $"找不到可用的 Vault 名稱，請確認本機視圖路徑是否有效：{VaultRootPath}");

            // ── 逐一嘗試登入（含 Timeout 保護）────────────────────────────────
            Exception lastError = null;
            foreach (var candidate in candidates)
            {
                try
                {
                    // LoginAuto 在 Server 無回應時會 hang；以 Task.Run + Wait 加 Timeout 保護
                    var vaultRef = _vault;
                    var loginTask = Task.Run(() => vaultRef.LoginAuto(candidate, 0));
                    bool finished = loginTask.Wait(TimeSpan.FromSeconds(LoginTimeoutSeconds));

                    if (!finished)
                    {
                        throw new TimeoutException(
                            $"連線到 Vault「{candidate}」逾時（{LoginTimeoutSeconds} 秒）。\n"
                            + "PDM Server 目前可能無法從此網路環境連線，請確認網路後重試。");
                    }

                    // 完成但 LoginAuto 內部拋出例外（例如 COM 錯誤）
                    if (loginTask.IsFaulted)
                        throw loginTask.Exception?.InnerException ?? loginTask.Exception
                              ?? new Exception("LoginAuto 發生未知錯誤");

                    if (_vault.IsLoggedIn)
                        return;   // ✅ 成功
                }
                catch (TimeoutException) { throw; }   // 直接往上傳，不被下一個 candidate 吞掉
                catch (Exception ex)     { lastError = ex; }
            }

            // ── 所有候選均失敗 ────────────────────────────────────────────────
            var joined = string.Join("、", candidates);

            if (lastError != null)
                throw new InvalidOperationException(
                    $"PDM 登入失敗（已嘗試 Vault：{joined}）。\n"
                    + $"錯誤詳情：{lastError.Message}",
                    lastError);

            // IsLoggedIn == false，但沒有例外 → 使用者取消或帳號無權限
            throw new InvalidOperationException(
                "PDM 登入未完成。可能原因：\n"
                + "• 使用者關閉了 PDM 登入視窗\n"
                + "• 帳號或密碼錯誤\n"
                + "• 此帳號在 Vault 中尚未建立使用者\n"
                + $"（已嘗試 Vault：{joined}）");
        }

        private string ResolveVaultNameFromPathSafe(string localPath)
        {
            try
            {
                var methods = _vault.GetType().GetMethods(BindingFlags.Instance | BindingFlags.Public);
                foreach (var m in methods)
                {
                    if (!string.Equals(m.Name, "GetVaultNameFromPath", StringComparison.Ordinal))
                    {
                        continue;
                    }

                    try
                    {
                        var parameters = m.GetParameters();
                        if (parameters.Length == 1 && parameters[0].ParameterType == typeof(string))
                        {
                            var name = m.Invoke(_vault, new object[] { localPath }) as string;
                            if (!string.IsNullOrWhiteSpace(name))
                            {
                                return name;
                            }
                        }

                        if (parameters.Length == 2 &&
                            parameters[0].ParameterType == typeof(string) &&
                            parameters[1].ParameterType == typeof(string).MakeByRefType())
                        {
                            object[] args = { localPath, string.Empty };
                            m.Invoke(_vault, args);
                            var name = args[1]?.ToString();
                            if (!string.IsNullOrWhiteSpace(name))
                            {
                                return name;
                            }
                        }
                    }
                    catch
                    {
                        // try next overload
                    }
                }
            }
            catch
            {
                // ignore and return empty
            }

            return string.Empty;
        }

        private BomItem BuildBomItem(string level, IEdmFile5 file, string fullPath, string referencedAs = "", IReadOnlyList<string> cardVarNames = null)
        {
            var state = string.Empty;
            var workflowState = string.Empty;
            var description = string.Empty;
            var partNumber = string.Empty;
            var descriptionVarUsed = string.Empty;
            var descriptionConfigUsed = string.Empty;
            var partNumberVarUsed = string.Empty;
            var partNumberConfigUsed = string.Empty;
            var cardVariables = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            try
            {
                var currentState = file.CurrentState;
                if (currentState != null)
                {
                    workflowState = currentState.Name ?? string.Empty;
                    ComHelper.Release(currentState);
                }
            }
            catch
            {
                workflowState = string.Empty;
            }

            // 某些 interop 版本不提供 CurrentStateID / GetStateFromID，這裡僅使用 CurrentState。

            IEdmEnumeratorVariable10 enumVar = null;
            try
            {
                enumVar = file.GetEnumeratorVariable() as IEdmEnumeratorVariable10;
                if (enumVar != null)
                {
                    var configs = BuildConfigCandidates(referencedAs);
                    state = GetVarValue(
                        enumVar,
                        new[] { "工作流程", "CP_基本資料狀態", "基本資料狀態", "狀態", "State", "Workflow State" },
                        configs,
                        out _,
                        out _);
                    description = GetVarValue(
                        enumVar,
                        new[] { "Description", "描述", "說明" },
                        configs,
                        out descriptionVarUsed,
                        out descriptionConfigUsed);
                    partNumber = GetVarValue(
                        enumVar,
                        new[] { "Part Number", "PartNumber", "料號", "品號", "零件編號" },
                        configs,
                        out partNumberVarUsed,
                        out partNumberConfigUsed);

                    // 動態模式：直接用使用者選定的變數名稱；備援：使用內建 CardVariableSpecs。
                    if (cardVarNames != null && cardVarNames.Count > 0)
                    {
                        foreach (var varName in cardVarNames)
                        {
                            var value = GetVarValue(enumVar, new[] { varName }, configs, out _, out _);
                            cardVariables[varName] = value;
                        }
                    }
                    else
                    {
                        foreach (var cardVar in CardVariableSpecs)
                        {
                            var value = GetVarValue(enumVar, cardVar.VariableNames, configs, out _, out _);
                            cardVariables[cardVar.Label] = value;
                        }
                    }
                }
            }
            finally
            {
                ComHelper.Release(enumVar);
            }

            var fn = Path.GetFileName(fullPath);
            var isDrw = string.Equals(Path.GetExtension(fn ?? string.Empty), ".slddrw",
                StringComparison.OrdinalIgnoreCase);

            return new BomItem
            {
                Level = level,
                FileName = fn,
                State = string.IsNullOrWhiteSpace(state) ? workflowState : state,
                WorkflowState = workflowState,
                Description = description,
                PartNumber = partNumber,
                FullPath = fullPath,
                ReferencedAs = referencedAs ?? string.Empty,
                DescriptionVarUsed = descriptionVarUsed,
                DescriptionConfigUsed = descriptionConfigUsed,
                PartNumberVarUsed = partNumberVarUsed,
                PartNumberConfigUsed = partNumberConfigUsed,
                CardVariables = cardVariables,
                // 經計算 BOM 可能直接帶出 .SLDDRW 列（階層常帶「-DRW」）；與事後插入的工程圖列相同，須標記供 UI 藍底。
                IsDrawing = isDrw
            };
        }

        private static string GetVarValue(
            IEdmEnumeratorVariable10 enumVar,
            IEnumerable<string> variableNames,
            IEnumerable<string> configCandidates,
            out string usedVariableName,
            out string usedConfig)
        {
            usedVariableName = string.Empty;
            usedConfig = string.Empty;

            foreach (var varName in variableNames)
            {
                foreach (var config in configCandidates)
                {
                    try
                    {
                        object value;
                        enumVar.GetVar(varName, config, out value);
                        var text = value?.ToString() ?? string.Empty;
                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            usedVariableName = varName;
                            usedConfig = config;
                            return text.Trim();
                        }
                    }
                    catch
                    {
                        // try next candidate
                    }
                }
            }

            return string.Empty;
        }

        private static IReadOnlyList<string> BuildConfigCandidates(string referencedAs)
        {
            var candidates = new List<string>();
            // 資料卡多數存於檔案層，優先讀 @
            candidates.Add("@");

            if (!string.IsNullOrWhiteSpace(referencedAs))
            {
                var refText = referencedAs.Trim();
                candidates.Add(refText);

                // 常見格式: filename.sldprt<ConfigName>
                var start = refText.IndexOf('<');
                var end = refText.LastIndexOf('>');
                if (start >= 0 && end > start + 1)
                {
                    var cfg = refText.Substring(start + 1, end - start - 1).Trim();
                    if (!string.IsNullOrWhiteSpace(cfg))
                    {
                        candidates.Add(cfg);
                    }
                }
            }

            return candidates
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        /// <summary>
        /// 將 Vault 檔案取至本機視圖（GetFileCopy／GetFileCopy2），成功回傳本機完整路徑。
        /// </summary>
        public string EnsureLocalFileRetrieved(string vaultFullPath, out string errorMessage)
        {
            errorMessage = string.Empty;
            if (string.IsNullOrWhiteSpace(vaultFullPath))
            {
                errorMessage = "路徑為空。";
                return null;
            }

            try
            {
                var norm = Path.GetFullPath(vaultFullPath.Trim().Trim('"'));
                if (!norm.StartsWith(VaultRootPath, StringComparison.OrdinalIgnoreCase))
                {
                    errorMessage = $"檔案必須位於 {VaultRootPath} 內。";
                    return null;
                }

                EnsureVaultLogin();
                IEdmFolder5 folder = null;
                IEdmFile5 file = null;
                try
                {
                    file = _vault.GetFileFromPath(norm, out folder);
                    if (file == null || folder == null)
                    {
                        errorMessage = "Vault 找不到檔案。";
                        return null;
                    }

                    // 先嘗試直接取得本機路徑（若使用者本機視圖已存在檔案，GetLocalPath 可能就能直接命中）
                    var local0 = file.GetLocalPath(folder.ID);
                    if (!string.IsNullOrWhiteSpace(local0) && File.Exists(local0))
                    {
                        return local0;
                    }

                    var okFileCopy = TryInvokeGetFileCopy(file, folder, out var getCopyDiagFile);
                    var okVaultCopy = false;
                    string getCopyDiagVault = string.Empty;
                    if (!okFileCopy)
                    {
                        okVaultCopy = TryInvokeGetFileCopyOnVault(_vault, file, folder, out getCopyDiagVault);
                    }

                    if (!okFileCopy && !okVaultCopy)
                    {
                        errorMessage = "無法呼叫 GetFileCopy／GetFileCopy2（EPDM API 不相容）。"
                                         + (string.IsNullOrWhiteSpace(getCopyDiagFile) ? string.Empty : " File診斷：" + getCopyDiagFile)
                                         + (string.IsNullOrWhiteSpace(getCopyDiagVault) ? string.Empty : " Vault診斷：" + getCopyDiagVault);
                        return null;
                    }

                    var local = file.GetLocalPath(folder.ID);
                    if (string.IsNullOrWhiteSpace(local) || !File.Exists(local))
                    {
                        errorMessage = "取檔後本機仍無檔案。";
                        return null;
                    }

                    return local;
                }
                finally
                {
                    ComHelper.Release(file);
                    ComHelper.Release(folder);
                }
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                return null;
            }
        }

        /// <summary>
        /// 取檔後嘗試在 Vault 執行「取出（check out / lock）」；不做 check in。
        /// </summary>
        public bool EnsureLocalFileRetrievedAndCheckedOut(
            string vaultFullPath,
            out string localPath,
            out string errorMessage)
        {
            IEdmFolder5 folder = null;
            IEdmFile5 file = null;
            try
            {
                EnsureVaultLogin();
                file = _vault.GetFileFromPath(vaultFullPath, out folder);
                if (file == null || folder == null)
                {
                    localPath = null;
                    errorMessage = "Vault 找不到檔案，無法進行取檔/check out。";
                    return false;
                }

                // 先嘗試取檔（GetFileCopy）
                localPath = EnsureLocalFileRetrieved(vaultFullPath, out var getErr);

                // ── Lock（check out）──
                var lockOk = false;
                var lockDiag = string.Empty;

                // 策略1：直接呼叫 IEdmFile5.LockFile(folderID, hWnd=0)
                if (!lockOk)
                {
                    try
                    {
                        file.LockFile(folder.ID, 0);
                        lockOk = true;
                    }
                    catch (Exception ex1)
                    {
                        lockDiag = $"直接 LockFile(folderID,0) 失敗：{ex1.Message}";
                    }
                }

                // 策略2：直接呼叫 LockFile(folderID, hWnd=0, flags=0)（IEdmFile7+）
                if (!lockOk)
                {
                    try
                    {
                        ((dynamic)file).LockFile(folder.ID, 0, 0);
                        lockOk = true;
                    }
                    catch (Exception ex2)
                    {
                        lockDiag += $" | dynamic LockFile(folderID,0,0) 失敗：{ex2.Message}";
                    }
                }

                // 策略3：反射暴力搜尋 Lock* 方法
                if (!lockOk)
                {
                    lockOk = TryInvokeLockFile(file, folder, out var reflDiag);
                    if (!lockOk)
                        lockDiag += " | " + reflDiag;
                }

                // 判斷 lock 後狀態
                if (!lockOk)
                {
                    // 或許已被自己鎖定
                    if (TryReadLockedState(file, out var isLocked) && isLocked)
                        lockOk = true;
                }

                // 取得本機路徑（lock 有時會連帶拉回檔案）
                if (string.IsNullOrWhiteSpace(localPath))
                {
                    try { localPath = file.GetLocalPath(folder.ID); } catch { /* ignore */ }
                }

                var fileExists = !string.IsNullOrWhiteSpace(localPath) && File.Exists(localPath);

                if (lockOk && fileExists)
                {
                    errorMessage = string.Empty;
                    return true;
                }

                if (lockOk && !fileExists)
                {
                    errorMessage = "check out 成功，但本機仍找不到檔案。";
                    return false;
                }

                // lock 失敗但本機有檔案 → 不需 check in，重建才是重點，視為可繼續
                if (!lockOk && fileExists)
                {
                    errorMessage = "WARNING: lock 未成功，但本機檔案已存在，繼續。 " + lockDiag;
                    return true;
                }

                errorMessage = "check out 失敗且本機無檔案。"
                               + (string.IsNullOrWhiteSpace(getErr) ? string.Empty : " 取檔診斷：" + getErr)
                               + " Lock診斷：" + lockDiag;
                return false;
            }
            catch (Exception ex)
            {
                localPath = null;
                errorMessage = "check out 發生例外：" + ex.Message;
                return false;
            }
            finally
            {
                ComHelper.Release(file);
                ComHelper.Release(folder);
            }
        }

        private static bool TryInvokeGetFileCopy(IEdmFile5 file, IEdmFolder5 folder, out string diag)
        {
            diag = string.Empty;
            if (file == null || folder == null)
            {
                return false;
            }

            var folderId = folder.ID;
            var allMethods = new List<MethodInfo>();
            foreach (var itf in GetKnownEdmFileInterfaceTypes())
            {
                try
                {
                    allMethods.AddRange(itf.GetMethods(BindingFlags.Instance | BindingFlags.Public));
                }
                catch
                {
                    // ignore
                }
            }

            var methods = allMethods
                .Where(m => m != null && m.Name != null && m.Name.StartsWith("GetFileCopy", StringComparison.OrdinalIgnoreCase))
                .GroupBy(m => m.ToString(), StringComparer.OrdinalIgnoreCase)
                .Select(g => g.First())
                .ToList();

            if (methods.Count == 0)
            {
                diag = "找不到 GetFileCopy* 重載（IEdmFile* 介面）。";
                return false;
            }

            foreach (var m in methods)
            {
                var p = m.GetParameters();
                if (p.Length == 0)
                {
                    continue;
                }

                // 組裝每個參數的候選值，並限制嘗試次數避免爆炸。
                // 目標：提高跨 EPDM 版本的重載簽名匹配成功率。
                var candidates = new List<List<object>>(p.Length);
                for (var i = 0; i < p.Length; i++)
                {
                    var pi = p[i];
                    var pt = pi.ParameterType;
                    var pn = (pi.Name ?? string.Empty);

                    if (pt.IsByRef)
                    {
                        var elem = pt.GetElementType();
                        if (elem == typeof(int))
                        {
                            candidates.Add(new List<object> { 0 });
                        }
                        else if (elem == typeof(short))
                        {
                            candidates.Add(new List<object> { (short)0 });
                        }
                        else if (elem == typeof(bool))
                        {
                            candidates.Add(new List<object> { false });
                        }
                        else if (elem != null && elem.IsValueType)
                        {
                            try
                            {
                                candidates.Add(new List<object> { Activator.CreateInstance(elem) });
                            }
                            catch
                            {
                                candidates.Add(new List<object> { null });
                            }
                        }
                        else
                        {
                            candidates.Add(new List<object> { null });
                        }
                        continue;
                    }

                    // Folder 介面（通常需要傳 IEdmFolder5）
                    if (pt.Name.IndexOf("Folder", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        candidates.Add(new List<object> { folder });
                        continue;
                    }

                    if (pt == typeof(string))
                    {
                        // 有些版本可能把此參數當作選項字串或暫存路徑
                        candidates.Add(new List<object> { string.Empty, null });
                        continue;
                    }

                    if (pt == typeof(bool))
                    {
                        candidates.Add(new List<object> { true, false });
                        continue;
                    }

                    if (pt == typeof(int) || pt == typeof(short) || pt == typeof(long) || pt == typeof(uint))
                    {
                        var ints = new List<long> { folderId, 1, 0 };
                        if (pt == typeof(int) || pt == typeof(short) || pt == typeof(long))
                        {
                            ints.Add(-1);
                            ints.Add(2);
                        }

                        var asObjects = new List<object>();
                        foreach (var v in ints)
                        {
                            try
                            {
                                if (pt == typeof(uint))
                                {
                                    if (v < 0) continue;
                                    asObjects.Add(Convert.ToUInt32(v));
                                }
                                else
                                {
                                    asObjects.Add(Convert.ChangeType(v, pt));
                                }
                            }
                            catch
                            {
                                // ignore
                            }
                        }

                        if (asObjects.Count == 0)
                        {
                            asObjects.Add(Convert.ChangeType(1, pt));
                        }

                        candidates.Add(asObjects);
                        continue;
                    }

                    if (!pt.IsValueType)
                    {
                        candidates.Add(new List<object> { null });
                        continue;
                    }

                    try
                    {
                        candidates.Add(new List<object> { Activator.CreateInstance(pt) });
                    }
                    catch
                    {
                        candidates.Add(new List<object> { null });
                    }
                }

                try
                {
                    // 限制最多嘗試 N 次（避免候選組合爆炸）
                    var maxAttempts = 48;
                    var attempt = 0;

                    var args = new object[p.Length];
                    void Recurse(int idx)
                    {
                        if (attempt >= maxAttempts)
                        {
                            return;
                        }

                        if (idx == p.Length)
                        {
                            if (attempt >= maxAttempts) return;
                            attempt++;
                            try
                            {
                                m.Invoke(file, args);

                                var local = file.GetLocalPath(folderId);
                                if (!string.IsNullOrWhiteSpace(local) && File.Exists(local))
                                {
                                    throw new _LocalRetrievedSignal();
                                }
                            }
                            catch (_LocalRetrievedSignal)
                            {
                                throw;
                            }
                            catch
                            {
                                // 嘗試下一組候選參數
                            }

                            return;
                        }

                        foreach (var cv in candidates[idx])
                        {
                            args[idx] = cv;
                            Recurse(idx + 1);
                            if (attempt >= maxAttempts)
                            {
                                return;
                            }
                        }
                    }

                    try
                    {
                        Recurse(0);
                    }
                    catch (_LocalRetrievedSignal)
                    {
                        // 已取到檔
                        return true;
                    }
                }
                catch (Exception ex)
                {
                    diag = $"嘗試方法 {m.Name}（{p.Length}參數）時出現例外：{ex.Message}";
                    // 嘗試下一個重載
                }
            }

            return false;
        }

        private static bool TryInvokeGetFileCopyOnVault(IEdmVault5 vault, IEdmFile5 file, IEdmFolder5 folder, out string diag)
        {
            diag = string.Empty;
            if (vault == null || file == null || folder == null)
            {
                return false;
            }

            var allMethods = new List<MethodInfo>();
            foreach (var itf in GetKnownEdmVaultInterfaceTypes())
            {
                try
                {
                    allMethods.AddRange(itf.GetMethods(BindingFlags.Instance | BindingFlags.Public));
                }
                catch
                {
                    // ignore
                }
            }

            var methods = allMethods
                .Where(m => m != null && m.Name != null && m.Name.StartsWith("GetFileCopy", StringComparison.OrdinalIgnoreCase))
                .GroupBy(m => m.ToString(), StringComparer.OrdinalIgnoreCase)
                .Select(g => g.First())
                .ToList();

            if (methods.Count == 0)
            {
                diag = "找不到 Vault.GetFileCopy* 重載（IEdmVault* 介面）。";
                return false;
            }

            // 檔案 ID：有些簽名可能用檔案 ID 而非 IEdmFile 物件。
            var fileId = 0;
            try { fileId = file.ID; } catch { /* ignore */ }

            var folderId = folder.ID;
            var fileType = file.GetType();
            var folderType = folder.GetType();

            foreach (var m in methods)
            {
                var p = m.GetParameters();
                if (p.Length == 0) continue;

                var args = new object[p.Length];
                var okArgBuild = true;
                for (var i = 0; i < p.Length; i++)
                {
                    var pt = p[i].ParameterType;
                    var pn = (p[i].Name ?? string.Empty);
                    var byref = pt.IsByRef;
                    var elem = byref ? pt.GetElementType() : pt;

                    try
                    {
                        if (byref && elem == typeof(int))
                        {
                            args[i] = 0;
                            continue;
                        }

                        if (elem == typeof(bool))
                        {
                            args[i] = true;
                            continue;
                        }

                        if (elem == typeof(string))
                        {
                            args[i] = string.Empty;
                            continue;
                        }

                        if (elem == typeof(int) || elem == typeof(short) || elem == typeof(long) || elem == typeof(uint))
                        {
                            // 根據參數名稱猜測是檔案 ID 或資料夾 ID 或選項旗標。
                            if (pn.IndexOf("File", StringComparison.OrdinalIgnoreCase) >= 0 || pn.IndexOf("Obj", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                args[i] = Convert.ChangeType(fileId, elem);
                                continue;
                            }

                            if (pn.IndexOf("Folder", StringComparison.OrdinalIgnoreCase) >= 0 || pn.IndexOf("Dest", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                args[i] = Convert.ChangeType(folderId, elem);
                                continue;
                            }

                            // 預設選項/旗標值
                            args[i] = Convert.ChangeType(1, elem);
                            continue;
                        }

                        // 物件型別：猜 IEdmFile / IEdmFolder
                        if (pt.Name.IndexOf("File", StringComparison.OrdinalIgnoreCase) >= 0 ||
                            pt.FullName != null && pt.FullName.IndexOf("IEdmFile", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            if (pt.IsAssignableFrom(fileType))
                            {
                                args[i] = file;
                            }
                            else
                            {
                                okArgBuild = false;
                            }
                            continue;
                        }

                        if (pt.Name.IndexOf("Folder", StringComparison.OrdinalIgnoreCase) >= 0 ||
                            pt.FullName != null && pt.FullName.IndexOf("IEdmFolder", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            if (pt.IsAssignableFrom(folderType))
                            {
                                args[i] = folder;
                            }
                            else
                            {
                                okArgBuild = false;
                            }
                            continue;
                        }

                        if (!elem.IsValueType)
                        {
                            args[i] = null;
                            continue;
                        }

                        // 其他值型別用預設值
                        args[i] = Activator.CreateInstance(elem);
                    }
                    catch
                    {
                        okArgBuild = false;
                    }
                }

                if (!okArgBuild) continue;

                try
                {
                    m.Invoke(vault, args);

                    var local = file.GetLocalPath(folder.ID);
                    if (!string.IsNullOrWhiteSpace(local) && File.Exists(local))
                    {
                        return true;
                    }
                }
                catch (Exception ex)
                {
                    diag = $"呼叫 {m.Name} 時例外：{ex.Message}";
                    // 嘗試下一個重載
                }
            }

            return false;
        }

        private static bool TryInvokeLockFile(IEdmFile5 file, IEdmFolder5 folder, out string diag)
        {
            diag = string.Empty;
            if (file == null || folder == null)
            {
                return false;
            }

            var allMethods = new List<MethodInfo>();
            // COM 執行期型別通常是 __ComObject，方法簽名不一定掛在執行期型別上；
            // 改由已知 IEdmFile* 介面清單抓 Lock* 方法。
            foreach (var itf in GetKnownEdmFileInterfaceTypes())
            {
                try
                {
                    allMethods.AddRange(itf.GetMethods(BindingFlags.Instance | BindingFlags.Public));
                }
                catch { /* ignore */ }
            }

            var methods = allMethods
                .Where(m => m != null && m.Name != null && m.Name.StartsWith("Lock", StringComparison.OrdinalIgnoreCase))
                .GroupBy(m => m.ToString(), StringComparer.OrdinalIgnoreCase)
                .Select(g => g.First())
                .ToList();

            if (methods.Count == 0)
            {
                diag = "找不到 Lock* 重載（IEdmFile* 介面）。";
                return false;
            }

            foreach (var m in methods)
            {
                var p = m.GetParameters();
                var argSets = BuildLockArgumentSets(p, folder.ID, folder);
                foreach (var args in argSets)
                {
                    try
                    {
                        m.Invoke(file, args);
                        if (TryReadLockedState(file, out var locked) && locked)
                        {
                            return true;
                        }

                        // 若無法讀取狀態，保守視為成功（部分 API 會丟 UI/狀態更新延遲）
                        if (!TryReadLockedState(file, out _))
                        {
                            return true;
                        }
                    }
                    catch (TargetInvocationException tie)
                    {
                        var inner = tie.InnerException;
                        var hr = inner is COMException cex ? $" (HRESULT=0x{cex.ErrorCode:X8})" : string.Empty;
                        diag = $"嘗試 {m} 失敗：{inner?.Message ?? tie.Message}{hr}";
                    }
                    catch (Exception ex)
                    {
                        diag = $"嘗試 {m} 失敗：{ex.Message}";
                    }
                }
            }

            if (string.IsNullOrWhiteSpace(diag))
            {
                diag = "Lock* 已嘗試但未能判定為已鎖定。可用方法：" +
                       string.Join(" | ", methods.Select(x => x.Name).Distinct(StringComparer.OrdinalIgnoreCase));
            }

            return false;
        }

        private static IReadOnlyList<object[]> BuildLockArgumentSets(ParameterInfo[] parameters, int folderId, IEdmFolder5 folder)
        {
            var sets = new List<object[]>();
            var bases = new[]
            {
                new[] { 0, folderId, 0, 0 },   // hwnd=0, parent=folderId, flags=0
                new[] { 0, folderId, 1, 0 },   // flags=1
                new[] { folderId, 0, 0, 0 },   // parent first
                new[] { folderId, 0, 1, 0 },   // parent first + flags
            };

            foreach (var b in bases)
            {
                var args = new object[parameters.Length];
                var ok = true;
                var intIndex = 0;
                for (var i = 0; i < parameters.Length; i++)
                {
                    var pt = parameters[i].ParameterType;
                    var pn = parameters[i].Name ?? string.Empty;
                    var elem = pt.IsByRef ? pt.GetElementType() : pt;
                    try
                    {
                        if (pt.IsByRef)
                        {
                            args[i] = elem == typeof(int) ? 0 :
                                      elem == typeof(bool) ? (object)false :
                                      (elem != null && elem.IsValueType ? Activator.CreateInstance(elem) : null);
                            continue;
                        }

                        if (pt.Name.IndexOf("Folder", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            args[i] = folder;
                            continue;
                        }

                        if (elem == typeof(int) || elem == typeof(short) || elem == typeof(long) || elem == typeof(uint))
                        {
                            var v = b[Math.Min(intIndex, b.Length - 1)];
                            // 強化命名判斷
                            if (pn.IndexOf("Hwnd", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                pn.IndexOf("Wnd", StringComparison.OrdinalIgnoreCase) >= 0)
                                v = 0;
                            else if (pn.IndexOf("Folder", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                     pn.IndexOf("Parent", StringComparison.OrdinalIgnoreCase) >= 0)
                                v = folderId;
                            intIndex++;

                            if (elem == typeof(uint) && v < 0) v = 0;
                            args[i] = Convert.ChangeType(v, elem);
                            continue;
                        }

                        if (elem == typeof(bool))
                        {
                            args[i] = true;
                            continue;
                        }

                        if (elem == typeof(string))
                        {
                            args[i] = string.Empty;
                            continue;
                        }

                        args[i] = elem != null && elem.IsValueType ? Activator.CreateInstance(elem) : null;
                    }
                    catch
                    {
                        ok = false;
                        break;
                    }
                }

                if (ok) sets.Add(args);
            }

            return sets;
        }

        private static bool TryReadLockedState(IEdmFile5 file, out bool locked)
        {
            locked = false;
            if (file == null) return false;

            var names = new[] { "IsLocked", "Locked", "IsLockedByMe", "LockedByMe" };
            foreach (var itf in GetKnownEdmFileInterfaceTypes())
            {
                foreach (var n in names)
                {
                    try
                    {
                        var p = itf.GetProperty(n, BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
                        if (p != null && p.PropertyType == typeof(bool))
                        {
                            locked = (bool)p.GetValue(file, null);
                            return true;
                        }
                    }
                    catch { /* ignore */ }
                }
            }

            return false;
        }

        private static IReadOnlyList<Type> GetKnownEdmFileInterfaceTypes()
        {
            var asm = typeof(IEdmFile5).Assembly;
            return asm.GetTypes()
                .Where(t =>
                    t != null &&
                    t.IsInterface &&
                    t.Name.StartsWith("IEdmFile", StringComparison.OrdinalIgnoreCase))
                .OrderBy(t => t.Name, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static IReadOnlyList<Type> GetKnownEdmVaultInterfaceTypes()
        {
            var asm = typeof(IEdmVault5).Assembly;
            return asm.GetTypes()
                .Where(t =>
                    t != null &&
                    t.IsInterface &&
                    t.Name.StartsWith("IEdmVault", StringComparison.OrdinalIgnoreCase))
                .OrderBy(t => t.Name, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private sealed class _LocalRetrievedSignal : Exception
        {
            // 用於在候選參數暴力嘗試時快速跳出
        }

        private sealed class CardVariableSpec
        {
            public string Label { get; }
            public IReadOnlyList<string> VariableNames { get; }

            public CardVariableSpec(string label, params string[] variableNames)
            {
                Label = label;
                VariableNames = variableNames
                    .Where(x => !string.IsNullOrWhiteSpace(x))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();
            }
        }
    }
}
