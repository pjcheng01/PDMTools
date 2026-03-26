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
            CancellationToken cancellationToken,
            bool viewPdfAfterSaving = false)
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
                    var stepLog = new StringBuilder();

                    ModelDoc2 modelDoc = null;
                    var weOpened = false;

                    try
                    {
                        stepLog.AppendLine("步驟1：工程圖取檔 + check out");
                        var okDrw = pdm.EnsureLocalFileRetrievedAndCheckedOut(vaultPath, out var localDrw, out var getErr);
                        if (okDrw)
                        {
                            stepLog.AppendLine("  - 工程圖已取檔並完成 check out：" + localDrw);
                            if (!string.IsNullOrWhiteSpace(getErr))
                                stepLog.AppendLine("  - 注意：" + getErr);
                        }
                        if (!okDrw || string.IsNullOrWhiteSpace(localDrw))
                        {
                            row.Message = stepLog + "失敗：工程圖取檔/check out 失敗：" + (getErr ?? string.Empty);
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

                        // ── 步驟2：開啟工程圖（僅為取得相依列表），取完即關閉 ──
                        stepLog.AppendLine("步驟2：開啟工程圖取得相依列表");
                        {
                            var err = 0;
                            var warn = 0;
                            modelDoc = TryFindAlreadyOpenDocument(swApp, localDrw);
                            if (modelDoc == null)
                            {
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
                        }

                        var allDeps = GetDependencyPaths(modelDoc);
                        stepLog.AppendLine($"  - 找到 {allDeps.Count} 個相依路徑");

                        // 關閉工程圖（及其獨佔的相依檔），以便 PDM 可 lock 相依
                        if (weOpened && modelDoc != null)
                        {
                            TryCloseDocument(swApp, modelDoc);
                            weOpened = false;
                            modelDoc = null;
                        }

                        // ── 步驟2b：lock 所有相依 ──
                        stepLog.AppendLine("步驟2b：相依檔 check out（工程圖已先關閉）");
                        var depLocalPaths = new List<string>();
                        var depCheckOutFailed = false;
                        foreach (var dep in allDeps)
                        {
                            cancellationToken.ThrowIfCancellationRequested();
                            if (!TryNormalizeVaultPath(dep, PdmBomExportService.VaultRootPath, out var depPath))
                            {
                                stepLog.AppendLine("  - 略過非 Vault 路徑：" + dep);
                                continue;
                            }

                            stepLog.AppendLine("  - 相依 check out：" + depPath);
                            var depOk = pdm.EnsureLocalFileRetrievedAndCheckedOut(depPath, out var depLocal, out var depErr);
                            if (!depOk || string.IsNullOrWhiteSpace(depLocal))
                            {
                                stepLog.AppendLine("    -> 失敗：" + (depErr ?? "未知"));
                                depCheckOutFailed = true;
                                continue;
                            }

                            if (!string.IsNullOrWhiteSpace(depErr))
                                stepLog.AppendLine("    -> 注意：" + depErr);

                            stepLog.AppendLine("    -> check out 完成：" + depLocal);
                            depLocalPaths.Add(depLocal);
                        }

                        if (depCheckOutFailed && depLocalPaths.Count == 0 && allDeps.Count > 0)
                        {
                            row.Message = stepLog + "失敗：所有相依 check out 均失敗。";
                            result.FailCount++;
                            result.Rows.Add(row);
                            continue;
                        }

                        // ── 步驟3：重新開啟工程圖（相依已 lock，檔案為可寫入狀態） ──
                        stepLog.AppendLine("步驟3：重新開啟工程圖（相依已 check out）");
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
                                row.Message = stepLog + $"失敗：相依 check out 後無法重新開啟工程圖（錯誤碼 {err2}，警告碼 {warn2}）。";
                                result.FailCount++;
                                result.Rows.Add(row);
                                continue;
                            }

                            weOpened = true;
                        }

                        stepLog.AppendLine("步驟4：相依檔（零件/組合件）逐一重建 + 原檔存檔");
                        foreach (var depLocal in depLocalPaths
                                     .Where(x => !string.IsNullOrWhiteSpace(x))
                                     .Distinct(StringComparer.OrdinalIgnoreCase))
                        {
                            cancellationToken.ThrowIfCancellationRequested();

                            var depExt = Path.GetExtension(depLocal);
                            var depDocType = GetDocTypeFromExtension(depExt);
                            if (depDocType == 0)
                            {
                                stepLog.AppendLine("  - 略過不支援副檔名：" + depLocal);
                                continue;
                            }

                            ModelDoc2 depDoc = null;
                            var depWeOpened = false;
                            try
                            {
                                stepLog.AppendLine("  - 開啟相依檔：" + Path.GetFileName(depLocal));
                                depDoc = TryFindAlreadyOpenDocument(swApp, depLocal);
                                if (depDoc == null)
                                {
                                    var depErrOpen = 0;
                                    var depWarnOpen = 0;
                                    depDoc = (ModelDoc2)swApp.OpenDoc6(
                                        depLocal,
                                        depDocType,
                                        SwOpenSilent,
                                        "",
                                        ref depErrOpen,
                                        ref depWarnOpen);
                                    if (depDoc == null)
                                    {
                                        stepLog.AppendLine("    -> 開啟失敗（Err=" + depErrOpen + ", Warn=" + depWarnOpen + "）");
                                        row.Message = stepLog.ToString().TrimEnd();
                                        result.FailCount++;
                                        result.Rows.Add(row);
                                        goto NextRow;
                                    }

                                    depWeOpened = true;
                                }

                                stepLog.AppendLine("    -> 已開啟，開始重建…");
                                TryRebuild(depDoc);
                                stepLog.AppendLine("    -> 重建完成，存檔中…");
                                if (!TrySaveNative(depDoc, out var depSaveErr))
                                {
                                    stepLog.AppendLine("    -> 存檔失敗：" + (depSaveErr ?? "未知"));
                                    row.Message = stepLog.ToString().TrimEnd();
                                    result.FailCount++;
                                    result.Rows.Add(row);
                                    goto NextRow;
                                }

                                stepLog.AppendLine("    -> 相依完成重建+存檔 OK");
                            }
                            finally
                            {
                                if (depWeOpened && depDoc != null)
                                {
                                    TryCloseDocument(swApp, depDoc);
                                }
                            }
                        }

                        stepLog.AppendLine("步驟5：工程圖重建 + 原檔存檔");
                        TryRebuild(modelDoc);
                        if (!TrySaveNative(modelDoc, out var drwSaveErr))
                        {
                            row.Message = stepLog + "失敗：工程圖重建後存檔失敗：" + (drwSaveErr ?? string.Empty);
                            result.FailCount++;
                            result.Rows.Add(row);
                            continue;
                        }

                        var baseName = Path.GetFileNameWithoutExtension(localDrw);
                        var pdfName = BuildUniquePdfFileName(baseName, usedPdfNames);
                        var pdfPath = Path.Combine(outDir, pdfName);

                        stepLog.AppendLine("步驟6：輸出 PDF：" + pdfPath);
                        if (!TrySaveAsPdf(swApp, modelDoc, pdfPath, viewPdfAfterSaving, out var saveErr))
                        {
                            row.Message = stepLog + "失敗：匯出 PDF 失敗：" + (saveErr ?? string.Empty);
                            result.FailCount++;
                            result.Rows.Add(row);
                            continue;
                        }

                        row.Status = "成功";
                        stepLog.AppendLine("完成：工程圖與相依已 check out、已重建、PDF 已輸出。");
                        row.Message = stepLog.ToString().TrimEnd();
                        row.OutputPdfPath = pdfPath;
                        result.SuccessCount++;
                        result.Rows.Add(row);
                    NextRow:
                        ;
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

        private static int GetDocTypeFromExtension(string ext)
        {
            if (string.IsNullOrWhiteSpace(ext))
            {
                return 0;
            }

            if (string.Equals(ext, ".sldprt", StringComparison.OrdinalIgnoreCase))
            {
                return 1; // swDocPART
            }

            if (string.Equals(ext, ".sldasm", StringComparison.OrdinalIgnoreCase))
            {
                return 2; // swDocASSEMBLY
            }

            if (string.Equals(ext, ".slddrw", StringComparison.OrdinalIgnoreCase))
            {
                return 3; // swDocDRAWING
            }

            return 0;
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
                // swSaveAsOptions_Silent = 1
                var ok = model.Save3(1, ref errs, ref warns);
                if (ok)
                {
                    return true;
                }

                errorMessage = $"Save3 回傳 false（Errors={errs}, Warnings={warns}）。";
                return false;
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                return false;
            }
        }

        private static bool TrySaveAsPdf(SldWorks swApp, ModelDoc2 model, string pdfPath,
            bool viewPdfAfterSaving, out string errorMessage)
        {
            errorMessage = string.Empty;
            try
            {
                var ext = (ModelDocExtension)model.Extension;
                var errs = 0;
                var warns = 0;

                object exportData = Missing.Value;
                try
                {
                    // swExportDataFileType_e.swExportPdfData = 1
                    var pdfData = (IExportPdfData)swApp.GetExportFileData(1);
                    if (pdfData != null)
                    {
                        pdfData.ViewPdfAfterSaving = viewPdfAfterSaving;
                        exportData = pdfData;
                    }
                }
                catch
                {
                    // 若 API 不支援則 fallback 為 Missing.Value
                }

                // swSaveAsVersion_e.swSaveAsCurrentVersion = 0；swSaveAsOptions_e.swSaveAsOptions_Silent = 1
                ext.SaveAs(pdfPath, 0, 1, exportData, ref errs, ref warns);
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

        /// <summary>
        /// 從工程圖取得所有參考的模型路徑（零件/組合件）。
        /// 使用三種方式依序嘗試：
        /// 1) DrawingDoc.GetSheetNames → Sheet.GetViews → View.GetReferencedModelName
        /// 2) ModelDoc2.GetDocumentDependencies2（EPDM 常見 API）
        /// 3) ModelDocExtension.GetDependencies（反射）
        /// </summary>
        private static IReadOnlyList<string> GetDependencyPaths(ModelDoc2 model)
        {
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            if (model == null)
            {
                return new List<string>();
            }

            // ── 方法1：DrawingDoc → Sheet → View → ReferencedModelName ──
            try
            {
                var drwDoc = (DrawingDoc)model;
                var sheetNames = drwDoc.GetSheetNames() as string[];
                if (sheetNames != null)
                {
                    foreach (var sheetName in sheetNames)
                    {
                        if (string.IsNullOrWhiteSpace(sheetName))
                        {
                            continue;
                        }

                        try
                        {
                            drwDoc.ActivateSheet(sheetName);
                        }
                        catch
                        {
                            // ignore
                        }

                        var sheet = (Sheet)drwDoc.GetCurrentSheet();
                        if (sheet == null)
                        {
                            continue;
                        }

                        var views = sheet.GetViews() as object[];
                        if (views == null)
                        {
                            continue;
                        }

                        foreach (var vObj in views)
                        {
                            if (vObj is View view)
                            {
                                try
                                {
                                    var refModel = view.GetReferencedModelName();
                                    if (!string.IsNullOrWhiteSpace(refModel))
                                    {
                                        set.Add(refModel.Trim());
                                    }
                                }
                                catch
                                {
                                    // ignore
                                }
                            }
                        }
                    }
                }
            }
            catch
            {
                // 不是 DrawingDoc 或介面差異
            }

            // ── 方法2：ModelDoc2 動態呼叫 GetDocumentDependencies2 ──
            if (set.Count == 0)
            {
                try
                {
                    dynamic dyn = model;
                    var deps = dyn.GetDocumentDependencies2(true, true, false);
                    if (deps is string[] sa2)
                    {
                        // 回傳格式為交錯：[name0, path0, name1, path1, ...]
                        for (var i = 1; i < sa2.Length; i += 2)
                        {
                            if (!string.IsNullOrWhiteSpace(sa2[i]))
                            {
                                set.Add(sa2[i].Trim());
                            }
                        }
                    }
                }
                catch
                {
                    // ignore
                }
            }

            // ── 方法3：ModelDocExtension.GetDependencies（反射）──
            if (set.Count == 0)
            {
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
                                        set.Add(s.Trim());
                                    }
                                }

                                if (set.Count > 0)
                                {
                                    break;
                                }
                            }

                            if (r is object[] ob)
                            {
                                foreach (var o in ob)
                                {
                                    var s = o?.ToString();
                                    if (!string.IsNullOrWhiteSpace(s))
                                    {
                                        set.Add(s.Trim());
                                    }
                                }

                                if (set.Count > 0)
                                {
                                    break;
                                }
                            }
                        }

                        if (set.Count > 0)
                        {
                            break;
                        }
                    }
                }
                catch
                {
                    // ignore
                }
            }

            return set.ToList();
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
