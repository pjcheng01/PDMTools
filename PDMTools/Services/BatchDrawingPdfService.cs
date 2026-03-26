using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using PDMTools.Models;
using SolidWorks.Interop.sldworks;

namespace PDMTools.Services
{
    /// <summary>
    /// 以目前 BOM 篩選後可見之工程圖列：取檔、SolidWorks 開啟／重建、匯出 PDF。
    /// COM 呼叫須在 STA 執行緒（建議由 WPF UI 執行緒直接呼叫）。
    /// </summary>
    public sealed class BatchDrawingPdfService
    {
        // swDocumentTypes_e.swDocDRAWING
        private const int SwDocDrawing = 3;

        // swOpenDocOptions_e：Silent
        private const int SwOpenSilent = 1;

        public BatchDrawingPdfResult Run(
            IReadOnlyList<BomItem> drawingItems,
            string outputFolder,
            PdmBomExportService pdm,
            IProgress<BatchDrawingPdfProgressInfo> progress,
            CancellationToken cancellationToken)
        {
            if (pdm == null)
            {
                throw new ArgumentNullException(nameof(pdm));
            }

            var result = new BatchDrawingPdfResult
            {
                Rows = new List<BatchDrawingPdfReportRow>()
            };

            var outDir = outputFolder?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(outDir))
            {
                throw new ArgumentException("請指定輸出資料夾。", nameof(outputFolder));
            }

            try
            {
                outDir = Path.GetFullPath(outDir);
            }
            catch (Exception ex)
            {
                throw new ArgumentException("輸出資料夾路徑無效。", nameof(outputFolder), ex);
            }

            if (!Directory.Exists(outDir))
            {
                Directory.CreateDirectory(outDir);
            }

            var list = (drawingItems ?? Array.Empty<BomItem>())
                .Where(i => i != null && i.IsDrawing && !string.IsNullOrWhiteSpace(i.FullPath))
                .GroupBy(i => Path.GetFullPath(i.FullPath.Trim()), StringComparer.OrdinalIgnoreCase)
                .Select(g => g.First())
                .ToList();

            var total = list.Count;
            if (total == 0)
            {
                result.ReportPath = WriteReportCsv(outDir, result.Rows);
                return result;
            }

            var swType = Type.GetTypeFromProgID("SldWorks.Application");
            if (swType == null)
            {
                throw new InvalidOperationException("找不到 SolidWorks（ProgID SldWorks.Application）。請確認已安裝 SolidWorks。");
            }

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
            {
                throw new InvalidOperationException("無法啟動 SolidWorks。");
            }

            try
            {
                try
                {
                    if (createdByThisRun)
                    {
                        swApp.Visible = false;
                    }
                }
                catch
                {
                    // 部分版本仍允許繼續
                }

                var usedPdfNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                for (var idx = 0; idx < list.Count; idx++)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    var item = list[idx];
                    var vaultPath = item.FullPath.Trim();
                    var cur = idx + 1;
                    progress?.Report(new BatchDrawingPdfProgressInfo(cur, total, $"處理中 ({cur}/{total})：{item.FileName}"));

                    var row = new BatchDrawingPdfReportRow
                    {
                        FileName = item.FileName,
                        VaultFullPath = vaultPath,
                        Status = "失敗"
                    };

                    ModelDoc2 modelDoc = null;
                    var weOpened = false;

                    try
                    {
                        var localDrw = pdm.EnsureLocalFileRetrieved(vaultPath, out var getErr);
                        if (string.IsNullOrWhiteSpace(localDrw))
                        {
                            row.Message = "取檔失敗：" + (getErr ?? string.Empty);
                            result.FailCount++;
                            result.Rows.Add(row);
                            continue;
                        }

                        if (!string.Equals(Path.GetExtension(localDrw), ".slddrw", StringComparison.OrdinalIgnoreCase))
                        {
                            row.Message = "取檔後副檔名不是 .slddrw。";
                            result.FailCount++;
                            result.Rows.Add(row);
                            continue;
                        }

                        modelDoc = TryFindAlreadyOpenDocument(swApp, localDrw);
                        if (modelDoc == null)
                        {
                            var err = 0;
                            var warn = 0;
                            modelDoc = (ModelDoc2)swApp.OpenDoc6(
                                localDrw,
                                SwDocDrawing,
                                SwOpenSilent,
                                "",
                                ref err,
                                ref warn);
                            if (modelDoc == null)
                            {
                                row.Message = $"SolidWorks 無法開啟工程圖（錯誤碼 {err}，警告碼 {warn}）。";
                                result.FailCount++;
                                result.Rows.Add(row);
                                continue;
                            }

                            weOpened = true;
                        }

                        foreach (var dep in GetDependencyPaths(modelDoc))
                        {
                            cancellationToken.ThrowIfCancellationRequested();
                            if (!TryNormalizeVaultPath(dep, PdmBomExportService.VaultRootPath, out var depPath))
                            {
                                continue;
                            }

                            var depLocal = pdm.EnsureLocalFileRetrieved(depPath, out var depErr);
                            if (string.IsNullOrWhiteSpace(depLocal))
                            {
                                row.Message = $"相依取檔失敗：{depPath} → {depErr}";
                                result.FailCount++;
                                result.Rows.Add(row);
                                if (weOpened && modelDoc != null)
                                {
                                    TryCloseDocument(swApp, modelDoc);
                                    weOpened = false;
                                    modelDoc = null;
                                }

                                continue;
                            }
                        }

                        if (weOpened && modelDoc != null)
                        {
                            TryCloseDocument(swApp, modelDoc);
                            weOpened = false;
                            modelDoc = null;
                        }

                        if (modelDoc == null)
                        {
                            var err2 = 0;
                            var warn2 = 0;
                            modelDoc = (ModelDoc2)swApp.OpenDoc6(
                                localDrw,
                                SwDocDrawing,
                                SwOpenSilent,
                                "",
                                ref err2,
                                ref warn2);
                            if (modelDoc == null)
                            {
                                row.Message = $"相依取檔後無法重新開啟工程圖（錯誤碼 {err2}，警告碼 {warn2}）。";
                                result.FailCount++;
                                result.Rows.Add(row);
                                continue;
                            }

                            weOpened = true;
                        }

                        TryRebuild(modelDoc);

                        var baseName = Path.GetFileNameWithoutExtension(localDrw);
                        var pdfName = BuildUniquePdfFileName(baseName, usedPdfNames);
                        var pdfPath = Path.Combine(outDir, pdfName);

                        if (!TrySaveAsPdf(modelDoc, pdfPath, out var saveErr))
                        {
                            row.Message = "匯出 PDF 失敗：" + (saveErr ?? string.Empty);
                            result.FailCount++;
                            result.Rows.Add(row);
                            continue;
                        }

                        row.Status = "成功";
                        row.Message = pdfPath;
                        row.OutputPdfPath = pdfPath;
                        result.SuccessCount++;
                        result.Rows.Add(row);
                    }
                    catch (OperationCanceledException)
                    {
                        throw;
                    }
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
                            TryCloseDocument(swApp, modelDoc);
                        }
                    }
                }
            }
            finally
            {
                if (createdByThisRun)
                {
                    try
                    {
                        swApp.ExitApp();
                    }
                    catch
                    {
                        // 忽略
                    }
                }
            }

            result.ReportPath = WriteReportCsv(outDir, result.Rows);
            return result;
        }

        private static string BuildUniquePdfFileName(string baseName, HashSet<string> used)
        {
            var name = $"{baseName}.pdf";
            if (!used.Contains(name))
            {
                used.Add(name);
                return name;
            }

            for (var n = 2; n < 10000; n++)
            {
                var candidate = $"{baseName}_{n}.pdf";
                if (!used.Contains(candidate))
                {
                    used.Add(candidate);
                    return candidate;
                }
            }

            return $"{baseName}_{Guid.NewGuid():N}.pdf";
        }

        private static void TryRebuild(ModelDoc2 model)
        {
            if (model == null)
            {
                return;
            }

            try
            {
                model.EditRebuild3();
            }
            catch
            {
                // 忽略
            }

            try
            {
                var extObj = model.Extension;
                if (extObj != null)
                {
                    var m = extObj.GetType().GetMethod(
                        "ForceRebuild3",
                        BindingFlags.Instance | BindingFlags.Public,
                        null,
                        new[] { typeof(bool) },
                        null);
                    m?.Invoke(extObj, new object[] { true });
                }
            }
            catch
            {
                // 忽略（此版 Interop 可能未宣告 ForceRebuild3）
            }
        }

        private static bool TrySaveAsPdf(ModelDoc2 model, string pdfPath, out string errorMessage)
        {
            errorMessage = string.Empty;
            try
            {
                var ext = (ModelDocExtension)model.Extension;
                var errs = 0;
                var warns = 0;
                // swSaveAsVersion_e.swSaveAsCurrentVersion = 0；swSaveAsOptions_e.swSaveAsOptions_Silent = 1
                // 此版 Interop 簽名為 SaveAs(string, int, int, object, ref int, ref int)
                ext.SaveAs(pdfPath, 0, 1, Missing.Value, ref errs, ref warns);
                if (File.Exists(pdfPath))
                {
                    return true;
                }

                errorMessage = $"SaveAs 回傳後仍無 PDF 檔（Errors={errs}，Warnings={warns}）。";
                return false;
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                return false;
            }
        }

        private static void TryCloseDocument(SldWorks swApp, ModelDoc2 modelDoc)
        {
            try
            {
                var title = modelDoc.GetTitle();
                if (!string.IsNullOrEmpty(title))
                {
                    swApp.CloseDoc(title);
                }
            }
            catch
            {
                // 忽略
            }
        }

        private static ModelDoc2 TryFindAlreadyOpenDocument(SldWorks swApp, string fullPath)
        {
            try
            {
                var target = Path.GetFullPath(fullPath.Trim());
                var docs = swApp.GetDocuments() as object[];
                if (docs == null)
                {
                    return null;
                }

                foreach (var d in docs)
                {
                    if (d is ModelDoc2 md)
                    {
                        try
                        {
                            var p = md.GetPathName();
                            if (!string.IsNullOrWhiteSpace(p) &&
                                string.Equals(Path.GetFullPath(p.Trim()), target, StringComparison.OrdinalIgnoreCase))
                            {
                                return md;
                            }
                        }
                        catch
                        {
                            // 忽略
                        }
                    }
                }
            }
            catch
            {
                // 忽略
            }

            return null;
        }

        private static IReadOnlyList<string> GetDependencyPaths(ModelDoc2 model)
        {
            var list = new List<string>();
            if (model == null)
            {
                return list;
            }

            try
            {
                var ext = model.Extension;
                var t = ext.GetType();
                foreach (var m in t.GetMethods(BindingFlags.Instance | BindingFlags.Public))
                {
                    if (!string.Equals(m.Name, "GetDependencies", StringComparison.Ordinal))
                    {
                        continue;
                    }

                    var p = m.GetParameters();
                    if (p.Length != 3 || p[0].ParameterType != typeof(bool))
                    {
                        continue;
                    }

                    for (var perm = 0; perm < 8; perm++)
                    {
                        var a = (perm & 1) != 0;
                        var b = (perm & 2) != 0;
                        var c = (perm & 4) != 0;
                        object r;
                        try
                        {
                            r = m.Invoke(ext, new object[] { a, b, c });
                        }
                        catch
                        {
                            continue;
                        }

                        if (r is string[] sa)
                        {
                            foreach (var s in sa)
                            {
                                if (!string.IsNullOrWhiteSpace(s))
                                {
                                    list.Add(s.Trim());
                                }
                            }

                            return list;
                        }

                        if (r is object[] ob)
                        {
                            foreach (var o in ob)
                            {
                                var s = o?.ToString();
                                if (!string.IsNullOrWhiteSpace(s))
                                {
                                    list.Add(s.Trim());
                                }
                            }

                            return list;
                        }
                    }
                }
            }
            catch
            {
                // 忽略
            }

            return list;
        }

        /// <summary>
        /// 將相依路徑正規化為 Vault 下完整路徑（不要求檔案已存在）。
        /// </summary>
        private static bool TryNormalizeVaultPath(string path, string vaultRoot, out string fullPath)
        {
            fullPath = string.Empty;
            if (string.IsNullOrWhiteSpace(path))
            {
                return false;
            }

            try
            {
                var t = path.Trim().Trim('"');
                if (t.Length >= 2 && t[0] == '"' && t[t.Length - 1] == '"')
                {
                    t = t.Substring(1, t.Length - 2);
                }

                var pipe = t.IndexOf('|');
                if (pipe >= 0)
                {
                    t = t.Substring(0, pipe).Trim();
                }

                if (!t.StartsWith(vaultRoot, StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }

                fullPath = Path.GetFullPath(t);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static string WriteReportCsv(string outputFolder, IReadOnlyList<BatchDrawingPdfReportRow> rows)
        {
            var name = $"BatchDrawingPdf_{DateTime.Now:yyyyMMdd_HHmmss}.csv";
            var path = Path.Combine(outputFolder, name);
            var sb = new StringBuilder();
            sb.AppendLine("FileName,VaultFullPath,Status,Message,OutputPdfPath");
            foreach (var r in rows ?? Array.Empty<BatchDrawingPdfReportRow>())
            {
                sb.AppendLine(string.Join(",",
                    CsvEscape(r?.FileName),
                    CsvEscape(r?.VaultFullPath),
                    CsvEscape(r?.Status),
                    CsvEscape(r?.Message),
                    CsvEscape(r?.OutputPdfPath)));
            }

            File.WriteAllText(path, sb.ToString(), new UTF8Encoding(true));
            return path;
        }

        private static string CsvEscape(string s)
        {
            if (string.IsNullOrEmpty(s))
            {
                return "\"\"";
            }

            var esc = s.Replace("\"", "\"\"");
            if (esc.IndexOfAny(new[] { ',', '"', '\r', '\n' }) >= 0)
            {
                return "\"" + esc + "\"";
            }

            return "\"" + esc + "\"";
        }
    }

    public sealed class BatchDrawingPdfResult
    {
        public int SuccessCount { get; set; }
        public int FailCount { get; set; }
        public string ReportPath { get; set; }
        public List<BatchDrawingPdfReportRow> Rows { get; set; }
    }

    public sealed class BatchDrawingPdfReportRow
    {
        public string FileName { get; set; }
        public string VaultFullPath { get; set; }
        public string Status { get; set; }
        public string Message { get; set; }
        public string OutputPdfPath { get; set; }
    }

    public sealed class BatchDrawingPdfProgressInfo
    {
        public BatchDrawingPdfProgressInfo(int current, int total, string message)
        {
            Current = current;
            Total = total;
            Message = message ?? string.Empty;
        }

        public int Current { get; }
        public int Total { get; }
        public string Message { get; }
    }
}
