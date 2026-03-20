using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
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
        private const string VaultRootPath = @"C:\CP-PDM";
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

        private readonly IEdmVault5 _vault;

        public PdmBomExportService()
        {
            _vault = new EdmVault5();
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

        public async Task<IReadOnlyList<BomItem>> CollectBomAsync(
            string assemblyPath,
            IReadOnlyList<string> cardVarNames,
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
                progress?.Report(new ProgressInfo(10, "已登入 PDM Vault，開始解析參考樹..."));

                var items = new List<BomItem>();
                IEdmFolder5 folder = null;
                IEdmFile5 file = null;
                IEdmReference5 refTree = null;

                try
                {
                    file = _vault.GetFileFromPath(assemblyPath, out folder);
                    if (file == null || folder == null)
                    {
                        throw new InvalidOperationException("無法從 PDM 取得檔案資訊。");
                    }

                    refTree = GetReferenceTree(file, folder);
                    if (refTree == null)
                    {
                        throw new InvalidOperationException("無法取得參考樹，請確認檔案版本或 PDM 權限。");
                    }

                    var rootItem = BuildBomItem("1", file, assemblyPath, string.Empty, cardVarNames);
                    items.Add(rootItem);

                    TraverseReferenceNodes(
                        parentNode: refTree,
                        parentLevel: "1",
                        output: items,
                        ancestryPaths: new HashSet<string>(StringComparer.OrdinalIgnoreCase) { assemblyPath },
                        cardVarNames: cardVarNames,
                        progress: progress,
                        cancellationToken: cancellationToken);
                }
                finally
                {
                    ComHelper.Release(refTree);
                    ComHelper.Release(folder);
                    ComHelper.Release(file);
                }

                progress?.Report(new ProgressInfo(75, $"解析完成，共 {items.Count} 筆。"));
                return (IReadOnlyList<BomItem>)items;
            }, cancellationToken);
        }

        public async Task ExportToExcelAsync(
            IReadOnlyList<BomItem> items,
            IReadOnlyList<string> cardVarNames,
            string outputPath,
            IProgress<ProgressInfo> progress,
            CancellationToken cancellationToken = default)
        {
            await Task.Run(() =>
            {
                if (items == null || items.Count == 0)
                {
                    throw new InvalidOperationException("沒有可匯出的資料。");
                }

                progress?.Report(new ProgressInfo(80, "開始輸出 Excel..."));

                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("BOM");
                    ws.Cell(1, 1).Value = "Level";
                    ws.Cell(1, 2).Value = "File Name";
                    ws.Cell(1, 3).Value = "State";
                    ws.Cell(1, 4).Value = "Workflow State";
                    ws.Cell(1, 5).Value = "Description";
                    ws.Cell(1, 6).Value = "Part Number";
                    ws.Cell(1, 7).Value = "Referenced As";
                    ws.Cell(1, 8).Value = "Full Path";
                    ws.Cell(1, 9).Value = "Description Var Used";
                    ws.Cell(1, 10).Value = "Description Config Used";
                    ws.Cell(1, 11).Value = "Part Number Var Used";
                    ws.Cell(1, 12).Value = "Part Number Config Used";
                    // 動態欄位標題
                    var exportVarNames = (cardVarNames != null && cardVarNames.Count > 0)
                        ? cardVarNames
                        : (IReadOnlyList<string>)CardVariableSpecs.Select(s => s.Label).ToList();

                    for (var i = 0; i < exportVarNames.Count; i++)
                    {
                        ws.Cell(1, 13 + i).Value = $"Card:{exportVarNames[i]}";
                    }

                    for (var i = 0; i < items.Count; i++)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        var row = i + 2;
                        var item = items[i];
                        ws.Cell(row, 1).Value = item.Level;
                        ws.Cell(row, 2).Value = item.FileName;
                        ws.Cell(row, 3).Value = item.State;
                        ws.Cell(row, 4).Value = item.WorkflowState;
                        ws.Cell(row, 5).Value = item.Description;
                        ws.Cell(row, 6).Value = item.PartNumber;
                        ws.Cell(row, 7).Value = item.ReferencedAs;
                        ws.Cell(row, 8).Value = item.FullPath;
                        ws.Cell(row, 9).Value = item.DescriptionVarUsed;
                        ws.Cell(row, 10).Value = item.DescriptionConfigUsed;
                        ws.Cell(row, 11).Value = item.PartNumberVarUsed;
                        ws.Cell(row, 12).Value = item.PartNumberConfigUsed;
                        for (var cardIndex = 0; cardIndex < exportVarNames.Count; cardIndex++)
                        {
                            var key = exportVarNames[cardIndex];
                            item.CardVariables.TryGetValue(key, out var val);
                            ws.Cell(row, 13 + cardIndex).Value = val ?? string.Empty;
                        }
                    }

                    ws.Row(1).Style.Font.Bold = true;
                    ws.Columns().AdjustToContents();
                    workbook.SaveAs(outputPath);
                }

                progress?.Report(new ProgressInfo(100, $"匯出完成：{outputPath}"));
            }, cancellationToken);
        }

        private void EnsureVaultLogin()
        {
            if (_vault.IsLoggedIn)
            {
                return;
            }

            var candidates = new List<string>();
            var resolved = ResolveVaultNameFromPathSafe(VaultRootPath);
            if (!string.IsNullOrWhiteSpace(resolved))
            {
                candidates.Add(resolved);
            }

            var folderName = Path.GetFileName(VaultRootPath.TrimEnd('\\', '/'));
            if (!string.IsNullOrWhiteSpace(folderName) && !candidates.Contains(folderName, StringComparer.OrdinalIgnoreCase))
            {
                candidates.Add(folderName);
            }

            if (candidates.Count == 0)
            {
                throw new InvalidOperationException($"找不到可用的 Vault 名稱，請確認本機視圖路徑是否有效：{VaultRootPath}");
            }

            Exception lastError = null;
            foreach (var candidate in candidates)
            {
                try
                {
                    _vault.LoginAuto(candidate, 0);
                    if (_vault.IsLoggedIn)
                    {
                        return;
                    }
                }
                catch (Exception ex)
                {
                    lastError = ex;
                }
            }

            var joined = string.Join(", ", candidates);
            throw new InvalidOperationException(
                $"PDM 登入失敗。已嘗試 Vault 名稱：{joined}。請確認 Vault 實際名稱與本機視圖對應是否一致。",
                lastError);
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

        private static IEdmReference5 GetReferenceTree(IEdmFile5 file, IEdmFolder5 folder)
        {
            try
            {
                return file.GetReferenceTree(folder.ID, file.CurrentVersion);
            }
            catch
            {
                return null;
            }
        }

        private void TraverseReferenceNodes(
            IEdmReference5 parentNode,
            string parentLevel,
            ICollection<BomItem> output,
            ISet<string> ancestryPaths,
            IReadOnlyList<string> cardVarNames,
            IProgress<ProgressInfo> progress,
            CancellationToken cancellationToken)
        {
            var index = 1;
            foreach (var node in EnumerateChildren(parentNode))
            {
                cancellationToken.ThrowIfCancellationRequested();

                var currentLevel = $"{parentLevel}.{index}";
                index++;

                var path = TryGetPathFromReferenceNode(node);
                if (string.IsNullOrWhiteSpace(path))
                {
                    ComHelper.Release(node);
                    continue;
                }

                // 只防止循環參照，不做全域去重，保留每個節點在 BOM 的實際出現。
                if (ancestryPaths.Contains(path))
                {
                    ComHelper.Release(node);
                    continue;
                }

                IEdmFolder5 folder = null;
                IEdmFile5 file = null;
                IEdmReference5 childTree = null;
                try
                {
                    file = _vault.GetFileFromPath(path, out folder);
                    if (file == null || folder == null)
                    {
                        continue;
                    }

                    var item = BuildBomItem(currentLevel, file, path, node.ReferencedAs, cardVarNames);
                    output.Add(item);
                    progress?.Report(new ProgressInfo(
                        percentage: Math.Min(70, 10 + (output.Count % 60)),
                        message: $"解析中：{item.FileName}"));

                    ancestryPaths.Add(path);
                    try
                    {
                        // 以子檔案重新取得參考樹，避免某些版本 node 子節點無法完整展開。
                        childTree = GetReferenceTree(file, folder);
                        if (childTree != null)
                        {
                            TraverseReferenceNodes(childTree, currentLevel, output, ancestryPaths, cardVarNames, progress, cancellationToken);
                        }
                    }
                    finally
                    {
                        ancestryPaths.Remove(path);
                    }
                }
                finally
                {
                    ComHelper.Release(childTree);
                    ComHelper.Release(file);
                    ComHelper.Release(folder);
                    ComHelper.Release(node);
                }
            }
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

            return new BomItem
            {
                Level = level,
                FileName = Path.GetFileName(fullPath),
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
                CardVariables = cardVariables
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

        private static IEnumerable<IEdmReference5> EnumerateChildren(IEdmReference5 parentNode)
        {
            string projectName = string.Empty;
            IEdmPos5 pos;
            try
            {
                pos = parentNode.GetFirstChildPosition(ref projectName, true, true, 0);
            }
            catch
            {
                yield break;
            }

            while (pos != null)
            {
                IEdmReference5 child;
                try
                {
                    child = parentNode.GetNextChild(pos);
                }
                catch
                {
                    yield break;
                }

                if (child == null)
                {
                    break;
                }

                yield return child;
            }
        }

        private static string TryGetPathFromReferenceNode(IEdmReference5 node)
        {
            try
            {
                var path = node.FoundPath;
                return path ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }
    }
}
