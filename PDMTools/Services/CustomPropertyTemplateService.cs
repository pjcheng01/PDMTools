using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading;
using PDMTools.Models;
using SolidWorks.Interop.sldworks;

namespace PDMTools.Services
{
    public sealed class CustomPropertyTemplateResult
    {
        public List<CustomPropertyTemplateRow> Rows { get; set; } = new List<CustomPropertyTemplateRow>();
        public int SuccessCount { get; set; }
        public int FailCount { get; set; }
    }

    public sealed class CustomPropertyTemplateRow : INotifyPropertyChanged
    {
        private bool _isSelectedForApply;
        private string _lastApplyMessage = string.Empty;

        public string FileName { get; set; } = string.Empty;
        public string Extension { get; set; } = string.Empty;
        public string VaultFullPath { get; set; } = string.Empty;
        /// <summary>一般自訂屬性範本（CustomPropertyBuilderTemplate[WeldmentTemplate=false]）。</summary>
        public string TemplatePath { get; set; } = string.Empty;
        public string TemplateName { get; set; } = string.Empty;
        /// <summary>焊件用範本（CustomPropertyBuilderTemplate[WeldmentTemplate=true]，.wldprp）。</summary>
        public string WeldmentTemplatePath { get; set; } = string.Empty;
        public string WeldmentTemplateName { get; set; } = string.Empty;
        public string Status { get; set; } = string.Empty;
        public string Message { get; set; } = string.Empty;

        /// <summary>是否參與「套用範本」批次處理。</summary>
        public bool IsSelectedForApply
        {
            get => _isSelectedForApply;
            set
            {
                if (_isSelectedForApply == value)
                    return;
                _isSelectedForApply = value;
                OnPropertyChanged();
            }
        }

        /// <summary>最近一次套用範本操作的結果說明。</summary>
        public string LastApplyMessage
        {
            get => _lastApplyMessage;
            set
            {
                if (_lastApplyMessage == value)
                    return;
                _lastApplyMessage = value ?? string.Empty;
                OnPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public sealed class CustomPropertyTemplateProgressInfo
    {
        public int Current { get; }
        public int Total { get; }
        public string Message { get; }

        public CustomPropertyTemplateProgressInfo(int current, int total, string message)
        {
            Current = current;
            Total = total;
            Message = message ?? string.Empty;
        }
    }

    public sealed class CustomPropertyTemplateApplyResult
    {
        public int SuccessCount { get; set; }
        public int FailCount { get; set; }
    }

    public sealed class CustomPropertyTemplateService
    {
        private const int SwDocPart = 1;
        private const int SwDocAssembly = 2;
        private const int SwDocDrawing = 3;
        private const int SwOpenSilent = 1;

        public CustomPropertyTemplateResult Run(
            IReadOnlyList<BomItem> bomItems,
            PdmBomExportService pdm,
            IProgress<CustomPropertyTemplateProgressInfo> progress,
            CancellationToken cancellationToken)
        {
            if (pdm == null)
                throw new ArgumentNullException(nameof(pdm));

            var result = new CustomPropertyTemplateResult();

            var list = (bomItems ?? Array.Empty<BomItem>())
                .Where(i => i != null && !string.IsNullOrWhiteSpace(i.FullPath))
                .Where(i =>
                {
                    var ext = Path.GetExtension(i.FullPath).ToLowerInvariant();
                    return ext == ".sldprt" || ext == ".sldasm" || ext == ".slddrw";
                })
                .GroupBy(i => Path.GetFullPath(i.FullPath.Trim()), StringComparer.OrdinalIgnoreCase)
                .Select(g => g.First())
                .ToList();

            var total = list.Count;
            if (total == 0)
                return result;

            var swApp = ConnectSolidWorks(out var createdByThisRun);

            try
            {
                try
                {
                    if (createdByThisRun)
                        swApp.Visible = false;
                }
                catch { }

                for (var idx = 0; idx < total; idx++)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    var item = list[idx];
                    var vaultPath = item.FullPath.Trim();
                    var cur = idx + 1;
                    progress?.Report(new CustomPropertyTemplateProgressInfo(
                        cur, total, $"讀取中 ({cur}/{total})：{item.FileName}"));

                    var row = new CustomPropertyTemplateRow
                    {
                        FileName = item.FileName,
                        Extension = Path.GetExtension(vaultPath).ToLowerInvariant(),
                        VaultFullPath = vaultPath,
                        Status = "失敗"
                    };

                    ModelDoc2 modelDoc = null;
                    var weOpened = false;

                    try
                    {
                        var localPath = pdm.EnsureLocalFileRetrieved(vaultPath, out var getErr);
                        if (string.IsNullOrWhiteSpace(localPath) || !File.Exists(localPath))
                        {
                            row.Message = "取檔失敗：" + (getErr ?? "未知");
                            result.FailCount++;
                            result.Rows.Add(row);
                            continue;
                        }

                        var docType = GetDocTypeFromExtension(Path.GetExtension(localPath));
                        if (docType == 0)
                        {
                            row.Message = "不支援的副檔名。";
                            result.FailCount++;
                            result.Rows.Add(row);
                            continue;
                        }

                        modelDoc = TryFindAlreadyOpenDocument(swApp, localPath);
                        if (modelDoc == null)
                        {
                            var err = 0;
                            var warn = 0;
                            modelDoc = (ModelDoc2)swApp.OpenDoc6(
                                localPath, docType, SwOpenSilent, "", ref err, ref warn);
                            if (modelDoc == null)
                            {
                                row.Message = $"SolidWorks 無法開啟（Err={err}, Warn={warn}）。";
                                result.FailCount++;
                                result.Rows.Add(row);
                                continue;
                            }
                            weOpened = true;
                        }

                        var ext = modelDoc.Extension as ModelDocExtension;
                        if (ext == null)
                        {
                            row.Message = "無法取得 ModelDocExtension。";
                            result.FailCount++;
                            result.Rows.Add(row);
                            continue;
                        }

                        // SolidWorks 2023 interop：IModelDocExtension.CustomPropertyBuilderTemplate 為「以 bool 索引」的屬性：
                        // get_CustomPropertyBuilderTemplate(WeldmentTemplate)：false = 一般 .prtprp/.asmprp/.drwprp；true = 焊件 .wldprp
                        string templatePath;
                        string weldPath;
                        try
                        {
                            templatePath = ext.CustomPropertyBuilderTemplate[false];
                            weldPath = ext.CustomPropertyBuilderTemplate[true];
                        }
                        catch (Exception ex)
                        {
                            row.Message = "讀取 CustomPropertyBuilderTemplate 失敗：" + ex.Message;
                            result.FailCount++;
                            result.Rows.Add(row);
                            continue;
                        }

                        row.TemplatePath = templatePath ?? string.Empty;
                        row.TemplateName = string.IsNullOrWhiteSpace(templatePath)
                            ? "（無）"
                            : Path.GetFileName(templatePath);
                        row.WeldmentTemplatePath = weldPath ?? string.Empty;
                        row.WeldmentTemplateName = string.IsNullOrWhiteSpace(weldPath)
                            ? "（無）"
                            : Path.GetFileName(weldPath);
                        row.Status = "成功";
                        row.Message = BuildTemplateSummaryMessage(templatePath, weldPath);
                        result.SuccessCount++;
                        result.Rows.Add(row);
                    }
                    catch (OperationCanceledException) { throw; }
                    catch (Exception ex)
                    {
                        row.Message = ex.Message;
                        result.FailCount++;
                        result.Rows.Add(row);
                    }
                    finally
                    {
                        if (weOpened && modelDoc != null)
                        {
                            try
                            {
                                var title = modelDoc.GetTitle();
                                if (!string.IsNullOrEmpty(title))
                                    swApp.CloseDoc(title);
                            }
                            catch { }
                        }
                    }
                }
            }
            finally
            {
                if (createdByThisRun)
                {
                    try { swApp.ExitApp(); } catch { }
                }
            }

            return result;
        }

        /// <summary>
        /// 將範本路徑寫入勾選列的 SolidWorks 檔（一般與焊件索引皆寫入）；取檔並 check out、存檔，不 check in。
        /// 勾選列必須為單一副檔名類型（僅 .sldprt、僅 .sldasm 或僅 .slddrw）。
        /// </summary>
        public CustomPropertyTemplateApplyResult ApplyTemplatesToRows(
            IReadOnlyList<CustomPropertyTemplateRow> allRows,
            string generalTemplatePath,
            string weldmentTemplatePathOptional,
            PdmBomExportService pdm,
            IProgress<CustomPropertyTemplateProgressInfo> progress,
            CancellationToken cancellationToken,
            out string validationError)
        {
            validationError = null;
            var result = new CustomPropertyTemplateApplyResult();

            if (pdm == null)
                throw new ArgumentNullException(nameof(pdm));

            var selected = (allRows ?? Array.Empty<CustomPropertyTemplateRow>())
                .Where(r => r != null && r.IsSelectedForApply)
                .ToList();

            if (selected.Count == 0)
            {
                validationError = "請至少勾選一列。";
                return result;
            }

            var exts = selected
                .Select(r => (r.Extension ?? string.Empty).ToLowerInvariant())
                .Where(e => !string.IsNullOrEmpty(e))
                .Distinct()
                .ToList();

            if (exts.Count != 1)
            {
                validationError =
                    "套用對象必須為單一 SolidWorks 類型：請勿混勾 .sldprt、.sldasm、.slddrw，請只勾選其中一種副檔名。";
                return result;
            }

            var docExt = exts[0];
            if (docExt != ".sldprt" && docExt != ".sldasm" && docExt != ".slddrw")
            {
                validationError = "不支援的檔案類型。";
                return result;
            }

            var gen = (generalTemplatePath ?? string.Empty).Trim();
            if (string.IsNullOrEmpty(gen) || !File.Exists(gen))
            {
                validationError = "請指定存在的一般範本檔案。";
                return result;
            }

            var wldRaw = (weldmentTemplatePathOptional ?? string.Empty).Trim();
            var wld = string.IsNullOrEmpty(wldRaw) ? gen : wldRaw;
            if (!File.Exists(wld))
            {
                validationError = "焊件範本路徑無效或檔案不存在。";
                return result;
            }

            gen = Path.GetFullPath(gen);
            wld = Path.GetFullPath(wld);

            if (!ValidateTemplateExtensionsMatchDocType(docExt, gen, wld, out var ve))
            {
                validationError = ve;
                return result;
            }

            var swApp = ConnectSolidWorks(out var createdByThisRun);

            try
            {
                try
                {
                    if (createdByThisRun)
                        swApp.Visible = false;
                }
                catch { }

                var total = selected.Count;
                for (var idx = 0; idx < total; idx++)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    var row = selected[idx];
                    var cur = idx + 1;
                    progress?.Report(new CustomPropertyTemplateProgressInfo(
                        cur, total, $"套用中 ({cur}/{total})：{row.FileName}"));

                    var vaultPath = (row.VaultFullPath ?? string.Empty).Trim();
                    row.LastApplyMessage = string.Empty;

                    ModelDoc2 modelDoc = null;
                    var weOpened = false;

                    try
                    {
                        if (string.IsNullOrEmpty(vaultPath))
                        {
                            row.LastApplyMessage = "失敗：無 Vault 路徑。";
                            result.FailCount++;
                            continue;
                        }

                        var okCo = pdm.EnsureLocalFileRetrievedAndCheckedOut(vaultPath, out var localPath, out var coErr);
                        if (!okCo || string.IsNullOrWhiteSpace(localPath) || !File.Exists(localPath))
                        {
                            row.LastApplyMessage = "失敗：取檔／check out：" + (coErr ?? "未知");
                            result.FailCount++;
                            continue;
                        }

                        var openType = GetDocTypeFromExtension(Path.GetExtension(localPath));
                        if (openType == 0)
                        {
                            row.LastApplyMessage = "失敗：不支援的副檔名。";
                            result.FailCount++;
                            continue;
                        }

                        modelDoc = TryFindAlreadyOpenDocument(swApp, localPath);
                        if (modelDoc == null)
                        {
                            var err = 0;
                            var warn = 0;
                            modelDoc = (ModelDoc2)swApp.OpenDoc6(
                                localPath, openType, SwOpenSilent, "", ref err, ref warn);
                            if (modelDoc == null)
                            {
                                row.LastApplyMessage = $"失敗：無法開啟（Err={err}, Warn={warn}）。";
                                result.FailCount++;
                                continue;
                            }
                            weOpened = true;
                        }

                        var ext = modelDoc.Extension as ModelDocExtension;
                        if (ext == null)
                        {
                            row.LastApplyMessage = "失敗：無法取得 ModelDocExtension。";
                            result.FailCount++;
                            continue;
                        }

                        if (!TrySetCustomPropertyBuilderTemplates(ext, gen, wld, out var setErr))
                        {
                            row.LastApplyMessage = "失敗：寫入範本：" + (setErr ?? string.Empty);
                            result.FailCount++;
                            continue;
                        }

                        if (!TrySaveNative(modelDoc, out var saveErr))
                        {
                            row.LastApplyMessage = "失敗：存檔：" + (saveErr ?? string.Empty);
                            result.FailCount++;
                            continue;
                        }

                        row.TemplatePath = gen;
                        row.TemplateName = Path.GetFileName(gen);
                        row.WeldmentTemplatePath = wld;
                        row.WeldmentTemplateName = Path.GetFileName(wld);
                        row.Message = BuildTemplateSummaryMessage(gen, wld);
                        row.Status = "成功";
                        row.LastApplyMessage =
                            "已寫入並存檔（一般與焊件索引皆已設定；原可能為未套用範本者亦已覆寫）。";
                        result.SuccessCount++;
                    }
                    catch (OperationCanceledException) { throw; }
                    catch (Exception ex)
                    {
                        row.LastApplyMessage = "失敗：" + ex.Message;
                        result.FailCount++;
                    }
                    finally
                    {
                        if (weOpened && modelDoc != null)
                        {
                            try
                            {
                                var title = modelDoc.GetTitle();
                                if (!string.IsNullOrEmpty(title))
                                    swApp.CloseDoc(title);
                            }
                            catch { }
                        }
                    }
                }
            }
            finally
            {
                if (createdByThisRun)
                {
                    try { swApp.ExitApp(); } catch { }
                }
            }

            return result;
        }

        private static SldWorks ConnectSolidWorks(out bool createdByThisRun)
        {
            createdByThisRun = false;
            var swType = Type.GetTypeFromProgID("SldWorks.Application");
            if (swType == null)
                throw new InvalidOperationException("找不到 SolidWorks（ProgID SldWorks.Application）。請確認已安裝。");

            SldWorks swApp = null;
            try
            {
                swApp = Marshal.GetActiveObject("SldWorks.Application") as SldWorks;
            }
            catch (COMException)
            {
                swApp = null;
            }

            if (swApp == null)
            {
                swApp = (SldWorks)Activator.CreateInstance(swType);
                createdByThisRun = true;
            }

            if (swApp == null)
                throw new InvalidOperationException("無法啟動 SolidWorks。");

            return swApp;
        }

        private static bool ValidateTemplateExtensionsMatchDocType(
            string docExt,
            string generalFullPath,
            string weldFullPath,
            out string error)
        {
            error = null;
            var g = generalFullPath.ToLowerInvariant();
            var w = weldFullPath.ToLowerInvariant();
            var sameFile = string.Equals(
                Path.GetFullPath(generalFullPath),
                Path.GetFullPath(weldFullPath),
                StringComparison.OrdinalIgnoreCase);

            switch (docExt.ToLowerInvariant())
            {
                case ".sldprt":
                    if (!g.EndsWith(".prtprp", StringComparison.OrdinalIgnoreCase))
                    {
                        error = "零件檔須搭配 .prtprp 一般範本。";
                        return false;
                    }
                    if (!sameFile && !w.EndsWith(".wldprp", StringComparison.OrdinalIgnoreCase))
                    {
                        error = "焊件索引須使用 .wldprp，或與一般範本指定同一檔案。";
                        return false;
                    }
                    break;
                case ".sldasm":
                    if (!g.EndsWith(".asmprp", StringComparison.OrdinalIgnoreCase))
                    {
                        error = "組合件須搭配 .asmprp 一般範本。";
                        return false;
                    }
                    if (!sameFile && !w.EndsWith(".asmprp", StringComparison.OrdinalIgnoreCase))
                    {
                        error = "第二路徑若非與一般範本相同，亦須為 .asmprp。";
                        return false;
                    }
                    break;
                case ".slddrw":
                    if (!g.EndsWith(".drwprp", StringComparison.OrdinalIgnoreCase))
                    {
                        error = "工程圖須搭配 .drwprp 一般範本。";
                        return false;
                    }
                    if (!sameFile && !w.EndsWith(".drwprp", StringComparison.OrdinalIgnoreCase))
                    {
                        error = "第二路徑若非與一般範本相同，亦須為 .drwprp。";
                        return false;
                    }
                    break;
                default:
                    error = "不支援的檔案類型。";
                    return false;
            }

            return true;
        }

        private static bool TrySetCustomPropertyBuilderTemplates(
            ModelDocExtension ext,
            string pathGeneral,
            string pathWeldment,
            out string errorMessage)
        {
            errorMessage = null;
            try
            {
                dynamic d = ext;
                d.CustomPropertyBuilderTemplate[false] = pathGeneral ?? string.Empty;
                d.CustomPropertyBuilderTemplate[true] = pathWeldment ?? string.Empty;
                return true;
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                return false;
            }
        }

        private static bool TrySaveNative(ModelDoc2 model, out string errorMessage)
        {
            errorMessage = string.Empty;
            if (model == null)
            {
                errorMessage = "ModelDoc2 為 null。";
                return false;
            }

            try
            {
                var errs = 0;
                var warns = 0;
                var ok = model.Save3(1, ref errs, ref warns);
                if (ok)
                    return true;

                errorMessage = $"Save3 回傳 false（Errors={errs}, Warnings={warns}）。";
                return false;
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                return false;
            }
        }

        private static int GetDocTypeFromExtension(string ext)
        {
            if (string.IsNullOrWhiteSpace(ext)) return 0;
            switch (ext.ToLowerInvariant())
            {
                case ".sldprt": return SwDocPart;
                case ".sldasm": return SwDocAssembly;
                case ".slddrw": return SwDocDrawing;
                default: return 0;
            }
        }

        private static string BuildTemplateSummaryMessage(string standardPath, string weldmentPath)
        {
            var stdEmpty = string.IsNullOrWhiteSpace(standardPath);
            var wldEmpty = string.IsNullOrWhiteSpace(weldmentPath);
            if (stdEmpty && wldEmpty)
                return "未套用範本（一般／焊件皆無）";
            if (stdEmpty)
                return "焊件範本：" + weldmentPath;
            if (wldEmpty)
                return "一般範本：" + standardPath;
            if (string.Equals(standardPath, weldmentPath, StringComparison.OrdinalIgnoreCase))
                return standardPath;
            return "一般：" + standardPath + "；焊件：" + weldmentPath;
        }

        private static ModelDoc2 TryFindAlreadyOpenDocument(SldWorks swApp, string fullPath)
        {
            try
            {
                var target = Path.GetFullPath(fullPath.Trim());
                var docs = swApp.GetDocuments() as object[];
                if (docs == null) return null;
                foreach (var obj in docs)
                {
                    var doc = obj as ModelDoc2;
                    if (doc == null) continue;
                    try
                    {
                        var p = doc.GetPathName();
                        if (!string.IsNullOrEmpty(p) &&
                            string.Equals(Path.GetFullPath(p), target, StringComparison.OrdinalIgnoreCase))
                            return doc;
                    }
                    catch { }
                }
            }
            catch { }
            return null;
        }
    }
}
