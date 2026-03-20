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
            bool includeRootBomItem,
            int? maxBomLayerDepth,
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

                    if (includeRootBomItem)
                    {
                        var rootItem = BuildBomItem("1", file, assemblyPath, string.Empty, cardVarNames);
                        items.Add(rootItem);

                        // root 在 BOM 顯示層 => root 視為 layer 1
                        TraverseReferenceNodes(
                            parentNode: refTree,
                            parentLevelPrefix: "1",
                            parentLayer: 1,
                            maxBomLayerDepth: maxBomLayerDepth,
                            output: items,
                            ancestryPaths: new HashSet<string>(StringComparer.OrdinalIgnoreCase) { assemblyPath },
                            cardVarNames: cardVarNames,
                            progress: progress,
                            cancellationToken: cancellationToken);
                    }
                    else
                    {
                        // root 不顯示 => root 視為 layer 0，第一層子件顯示為 layer 1
                        TraverseReferenceNodes(
                            parentNode: refTree,
                            parentLevelPrefix: string.Empty,
                            parentLayer: 0,
                            maxBomLayerDepth: maxBomLayerDepth,
                            output: items,
                            ancestryPaths: new HashSet<string>(StringComparer.OrdinalIgnoreCase) { assemblyPath },
                            cardVarNames: cardVarNames,
                            progress: progress,
                            cancellationToken: cancellationToken);
                    }
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

                // 快取：baseName → 找到的工程圖原型（Level 空白，後續 Clone 時填入）
                var foundCache    = new Dictionary<string, BomItem>(StringComparer.OrdinalIgnoreCase);
                var notFoundCache = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                for (var i = 0; i < total; i++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var item = bomItems[i];
                    result.Add(item);

                    if (item.IsDrawing) continue;

                    var ext = Path.GetExtension(item.FileName);
                    if (!ext.Equals(".sldasm", StringComparison.OrdinalIgnoreCase) &&
                        !ext.Equals(".sldprt", StringComparison.OrdinalIgnoreCase))
                        continue;

                    var baseName      = Path.GetFileNameWithoutExtension(item.FileName);
                    var drawingLevel  = item.Level + "-DRW";

                    progress?.Report(new ProgressInfo(
                        Math.Min(95, 75 + (int)(20.0 * i / total)),
                        $"搜尋工程圖：{item.FileName}"));

                    if (notFoundCache.Contains(baseName)) continue;

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

                progress?.Report(new ProgressInfo(96, $"工程圖搜尋完成。"));
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
            string parentLevelPrefix,
            int parentLayer,
            int? maxBomLayerDepth,
            ICollection<BomItem> output,
            ISet<string> ancestryPaths,
            IReadOnlyList<string> cardVarNames,
            IProgress<ProgressInfo> progress,
            CancellationToken cancellationToken)
        {
            // parentLayer 已達最大層數 => 不再展開其子節點
            if (maxBomLayerDepth != null && parentLayer >= maxBomLayerDepth.Value)
                return;

            var index = 1;
            foreach (var node in EnumerateChildren(parentNode))
            {
                cancellationToken.ThrowIfCancellationRequested();

                var currentLayer = parentLayer + 1;
                if (maxBomLayerDepth != null && currentLayer > maxBomLayerDepth.Value)
                    return;

                // parentLevelPrefix 為空時，第一層以「index」表示（避免多出一層點號）
                var currentLevel = string.IsNullOrWhiteSpace(parentLevelPrefix)
                    ? index.ToString()
                    : $"{parentLevelPrefix}.{index}";
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
                            TraverseReferenceNodes(
                                childTree,
                                currentLevel,
                                currentLayer,
                                maxBomLayerDepth,
                                output,
                                ancestryPaths,
                                cardVarNames,
                                progress,
                                cancellationToken);
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
