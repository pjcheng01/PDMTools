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
    public sealed class DrawingSheetFormatQueryResult
    {
        public List<DrawingSheetFormatRow> Rows { get; set; } = new List<DrawingSheetFormatRow>();
        public int SuccessCount { get; set; }
        public int FailCount { get; set; }
    }

    /// <summary>
    /// 單一工程圖中的單一圖頁列（ISheet.GetTemplateName → .slddrt 路徑）。
    /// </summary>
    public sealed class DrawingSheetFormatRow
    {
        public string FileName { get; set; } = string.Empty;
        public string VaultFullPath { get; set; } = string.Empty;
        public string SheetName { get; set; } = string.Empty;
        public string FormatPath { get; set; } = string.Empty;
        public string FormatFileName { get; set; } = string.Empty;
        public string Status { get; set; } = string.Empty;
        public string Message { get; set; } = string.Empty;
    }

    public sealed class DrawingSheetFormatProgressInfo
    {
        public int Current { get; }
        public int Total { get; }
        public string Message { get; }

        public DrawingSheetFormatProgressInfo(int current, int total, string message)
        {
            Current = current;
            Total = total;
            Message = message ?? string.Empty;
        }
    }

    public sealed class DrawingSheetFormatQueryService
    {
        private const int SwDocDrawing = 3;
        private const int SwOpenSilent = 1;

        public DrawingSheetFormatQueryResult Run(
            IReadOnlyList<BomItem> drawingItems,
            PdmBomExportService pdm,
            IProgress<DrawingSheetFormatProgressInfo> progress,
            CancellationToken cancellationToken)
        {
            if (pdm == null)
                throw new ArgumentNullException(nameof(pdm));

            var result = new DrawingSheetFormatQueryResult();

            var list = (drawingItems ?? Array.Empty<BomItem>())
                .Where(i => i != null && i.IsDrawing && !string.IsNullOrWhiteSpace(i.FullPath))
                .Where(i => string.Equals(Path.GetExtension(i.FullPath), ".slddrw", StringComparison.OrdinalIgnoreCase))
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
                    progress?.Report(new DrawingSheetFormatProgressInfo(
                        cur, total, $"讀取中 ({cur}/{total})：{item.FileName}"));

                    ModelDoc2 modelDoc = null;
                    var weOpened = false;

                    try
                    {
                        var localPath = pdm.EnsureLocalFileRetrieved(vaultPath, out var getErr);
                        if (string.IsNullOrWhiteSpace(localPath) || !File.Exists(localPath))
                        {
                            result.Rows.Add(new DrawingSheetFormatRow
                            {
                                FileName = item.FileName,
                                VaultFullPath = vaultPath,
                                SheetName = "—",
                                Status = "失敗",
                                Message = "取檔失敗：" + (getErr ?? "未知")
                            });
                            result.FailCount++;
                            continue;
                        }

                        if (!string.Equals(Path.GetExtension(localPath), ".slddrw", StringComparison.OrdinalIgnoreCase))
                        {
                            result.Rows.Add(new DrawingSheetFormatRow
                            {
                                FileName = item.FileName,
                                VaultFullPath = vaultPath,
                                SheetName = "—",
                                Status = "失敗",
                                Message = "本機路徑副檔名不是 .slddrw。"
                            });
                            result.FailCount++;
                            continue;
                        }

                        modelDoc = TryFindAlreadyOpenDocument(swApp, localPath);
                        if (modelDoc == null)
                        {
                            var err = 0;
                            var warn = 0;
                            modelDoc = (ModelDoc2)swApp.OpenDoc6(
                                localPath, SwDocDrawing, SwOpenSilent, "", ref err, ref warn);
                            if (modelDoc == null)
                            {
                                result.Rows.Add(new DrawingSheetFormatRow
                                {
                                    FileName = item.FileName,
                                    VaultFullPath = vaultPath,
                                    SheetName = "—",
                                    Status = "失敗",
                                    Message = $"SolidWorks 無法開啟（Err={err}, Warn={warn}）。"
                                });
                                result.FailCount++;
                                continue;
                            }

                            weOpened = true;
                        }

                        var drw = modelDoc as DrawingDoc;
                        if (drw == null)
                        {
                            result.Rows.Add(new DrawingSheetFormatRow
                            {
                                FileName = item.FileName,
                                VaultFullPath = vaultPath,
                                SheetName = "—",
                                Status = "失敗",
                                Message = "無法轉成 DrawingDoc。"
                            });
                            result.FailCount++;
                            continue;
                        }

                        var sheetNames = drw.GetSheetNames() as string[];
                        if (sheetNames == null || sheetNames.Length == 0)
                        {
                            result.Rows.Add(new DrawingSheetFormatRow
                            {
                                FileName = item.FileName,
                                VaultFullPath = vaultPath,
                                SheetName = "—",
                                Status = "失敗",
                                Message = "未取得任何圖頁名稱（GetSheetNames）。"
                            });
                            result.FailCount++;
                            continue;
                        }

                        foreach (var sheetName in sheetNames)
                        {
                            cancellationToken.ThrowIfCancellationRequested();

                            if (string.IsNullOrWhiteSpace(sheetName))
                            {
                                result.Rows.Add(new DrawingSheetFormatRow
                                {
                                    FileName = item.FileName,
                                    VaultFullPath = vaultPath,
                                    SheetName = "（空白）",
                                    Status = "失敗",
                                    Message = "圖頁名稱為空，已略過。"
                                });
                                result.FailCount++;
                                continue;
                            }

                            var row = new DrawingSheetFormatRow
                            {
                                FileName = item.FileName,
                                VaultFullPath = vaultPath,
                                SheetName = sheetName.Trim(),
                                Status = "失敗"
                            };

                            try
                            {
                                try
                                {
                                    drw.ActivateSheet(sheetName);
                                }
                                catch
                                {
                                    // 仍嘗試讀取目前圖頁
                                }

                                var sheet = drw.GetCurrentSheet() as Sheet;
                                if (sheet == null)
                                {
                                    row.Message = "GetCurrentSheet 為 null。";
                                    result.FailCount++;
                                    result.Rows.Add(row);
                                    continue;
                                }

                                string formatPath;
                                try
                                {
                                    formatPath = sheet.GetTemplateName();
                                }
                                catch (Exception ex)
                                {
                                    row.Message = "GetTemplateName 失敗：" + ex.Message;
                                    result.FailCount++;
                                    result.Rows.Add(row);
                                    continue;
                                }

                                row.FormatPath = formatPath ?? string.Empty;
                                row.FormatFileName = string.IsNullOrWhiteSpace(formatPath)
                                    ? "（無）"
                                    : Path.GetFileName(formatPath);
                                row.Status = "成功";
                                row.Message = string.IsNullOrWhiteSpace(formatPath)
                                    ? "未指定圖頁格式路徑"
                                    : formatPath;
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
                        }
                    }
                    catch (OperationCanceledException)
                    {
                        throw;
                    }
                    catch (Exception ex)
                    {
                        result.Rows.Add(new DrawingSheetFormatRow
                        {
                            FileName = item.FileName,
                            VaultFullPath = vaultPath,
                            SheetName = "—",
                            Status = "失敗",
                            Message = ex.Message
                        });
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
