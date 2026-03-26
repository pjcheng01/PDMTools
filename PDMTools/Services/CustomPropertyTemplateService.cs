using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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

    public sealed class CustomPropertyTemplateRow
    {
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

            var swType = Type.GetTypeFromProgID("SldWorks.Application");
            if (swType == null)
                throw new InvalidOperationException("找不到 SolidWorks（ProgID SldWorks.Application）。請確認已安裝。");

            SldWorks swApp = null;
            var createdByThisRun = false;
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
