using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using ClosedXML.Excel;
using PDMTools.Models;

namespace PDMTools.Services
{
    public sealed class IgpBomComparer
    {
        private static readonly Regex BracketRegex = new Regex(@"\[(.*?)\]", RegexOptions.Compiled);

        public async Task<IgpBomCompareResult> CompareAsync(
            IReadOnlyList<BomItem> pdmBomItems,
            string igpBomFilePath,
            int maxDepth,
            bool excludeDrawings,
            IProgress<ProgressInfo> progress,
            CancellationToken cancellationToken)
        {
            return await Task.Run(() =>
            {
                if (string.IsNullOrWhiteSpace(igpBomFilePath) || !File.Exists(igpBomFilePath))
                {
                    throw new FileNotFoundException("找不到 iGP BOM 檔案（.xlsx / .csv）。", igpBomFilePath);
                }

                if (pdmBomItems == null || pdmBomItems.Count == 0)
                {
                    throw new InvalidOperationException("PDM BOM 無資料可比對。");
                }

                if (maxDepth <= 0)
                {
                    throw new ArgumentOutOfRangeException(nameof(maxDepth));
                }

                progress?.Report(new ProgressInfo(0, "解析 iGP BOM…"));
                var igpNodes = ParseIgpBomFileToNodes(igpBomFilePath, maxDepth, progress, cancellationToken, out var parseNotes);

                progress?.Report(new ProgressInfo(25, "建立 PDM BOM 節點…"));
                var pdmNodes = BuildPdmNodes(pdmBomItems, maxDepth, excludeDrawings);

                progress?.Report(new ProgressInfo(45, "Stage1：彙總比對…"));
                var stage1 = CompareStage1(pdmNodes, igpNodes);

                progress?.Report(new ProgressInfo(60, "Stage2：結構邊比對…"));
                var pdmEdges = BuildEdges(pdmNodes);
                var igpEdges = BuildEdges(igpNodes);
                var stage2 = CompareStage2(pdmEdges, igpEdges);

                progress?.Report(new ProgressInfo(100, "比對完成。"));

                return new IgpBomCompareResult
                {
                    IgpSourcePath = igpBomFilePath,
                    PdmSourcePath = string.Empty,
                    MaxDepth = maxDepth,
                    ComparedAtLocal = DateTime.Now,
                    IgpNodeCount = igpNodes.Count,
                    PdmNodeCount = pdmNodes.Count,
                    IgpEdgeCount = igpEdges.Count,
                    PdmEdgeCount = pdmEdges.Count,
                    Notes = parseNotes ?? string.Empty,
                    Stage1Rows = stage1,
                    Stage2Rows = stage2
                };
            }, cancellationToken);
        }

        /// <summary>
        /// 依副檔名選擇：.xlsx/.xlsm 用 ClosedXML 讀 Unicode；.csv 則維持 UTF-8／Big5 回退。
        /// </summary>
        private static List<IgpBomNode> ParseIgpBomFileToNodes(
            string path,
            int maxDepth,
            IProgress<ProgressInfo> progress,
            CancellationToken cancellationToken,
            out string notes)
        {
            var ext = Path.GetExtension(path ?? string.Empty);
            if (IsExcelWorkbookExtension(ext))
            {
                var nodes = ParseIgpBomXlsxToNodes(path, maxDepth, progress, cancellationToken);
                notes = nodes.Count > 0
                    ? "iGP 來源：Excel（.xlsx／.xlsm），工作表以 Unicode 讀取，無 CSV 編碼猜測問題。"
                    : "警告：iGP Excel 未解析到任何節點。請確認第一個工作表含樹狀欄位（含「[」與階層符號 ◎□■◇◆○●△）。";
                if (nodes.Count > 0)
                {
                    AssignRootLabels(nodes);
                }

                return nodes;
            }

            return ParseIgpBomCsvToNodesWithFallbackEncodings(path, maxDepth, progress, cancellationToken, out notes);
        }

        private static bool IsExcelWorkbookExtension(string ext)
        {
            return string.Equals(ext, ".xlsx", StringComparison.OrdinalIgnoreCase)
                   || string.Equals(ext, ".xlsm", StringComparison.OrdinalIgnoreCase);
        }

        private static List<IgpBomNode> ParseIgpBomXlsxToNodes(
            string path,
            int maxDepth,
            IProgress<ProgressInfo> progress,
            CancellationToken cancellationToken)
        {
            var nodes = new List<IgpBomNode>(256);
            using (var wb = new XLWorkbook(path))
            {
                var ws = wb.Worksheet(1);
                var lastCell = ws.LastCellUsed();
                if (lastCell == null)
                {
                    return nodes;
                }

                var lastRow = lastCell.Address.RowNumber;
                var maxCol = lastCell.Address.ColumnNumber;
                if (lastRow <= 0 || maxCol <= 0)
                {
                    return nodes;
                }

                var total = Math.Max(1, lastRow);
                for (var r = 1; r <= lastRow; r++)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    var row = ws.Row(r);
                    var treeCell = FindTreeCellFromExcelRow(row, maxCol);
                    if (string.IsNullOrWhiteSpace(treeCell))
                    {
                        continue;
                    }

                    TryAppendNodeFromTreeCell(treeCell, maxDepth, nodes);

                    if ((r % 60) == 0)
                    {
                        var pct = (int)Math.Round(20.0 * r / total, MidpointRounding.AwayFromZero);
                        progress?.Report(new ProgressInfo(pct, $"解析 iGP BOM（Excel {r}/{lastRow}）…"));
                    }
                }
            }

            return nodes;
        }

        /// <summary>同 CSV：該列中第一個含「[」的儲存格視為樹狀儲存格。</summary>
        private static string FindTreeCellFromExcelRow(IXLRow row, int maxColumn)
        {
            if (row == null || maxColumn <= 0)
            {
                return string.Empty;
            }

            for (var c = 1; c <= maxColumn; c++)
            {
                var cell = row.Cell(c);
                var text = GetExcelCellText(cell);
                if (!string.IsNullOrEmpty(text) && text.IndexOf('[') >= 0)
                {
                    return text.Trim();
                }
            }

            return string.Empty;
        }

        private static string GetExcelCellText(IXLCell cell)
        {
            if (cell == null)
            {
                return string.Empty;
            }

            try
            {
                if (cell.HasFormula)
                {
                    var v = cell.Value;
                    if (!v.IsBlank)
                    {
                        return v.ToString()?.Trim() ?? string.Empty;
                    }
                }

                var s = cell.GetString();
                if (!string.IsNullOrWhiteSpace(s))
                {
                    return s.Trim();
                }

                var val = cell.Value;
                return val.IsBlank ? string.Empty : (val.ToString()?.Trim() ?? string.Empty);
            }
            catch
            {
                try
                {
                    return cell.Value.ToString()?.Trim() ?? string.Empty;
                }
                catch
                {
                    return string.Empty;
                }
            }
        }

        private static List<IgpBomNode> ParseIgpBomCsvToNodesWithFallbackEncodings(
            string igpCsvPath,
            int maxDepth,
            IProgress<ProgressInfo> progress,
            CancellationToken cancellationToken,
            out string notes)
        {
            // 常見：iGP CSV 可能是 Big5/ANSI；先用 UTF-8 解析，若完全抓不到節點再回退。
            notes = string.Empty;
            var utf8 = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: false);
            var nodes = ParseIgpBomCsvToNodes(igpCsvPath, maxDepth, utf8, progress, cancellationToken);
            if (nodes.Count > 0)
            {
                notes = "iGP CSV 編碼：UTF-8（或可被 UTF-8 容錯讀取）";
                AssignRootLabels(nodes);
                return nodes;
            }

            try
            {
                var big5 = Encoding.GetEncoding(950); // Big5
                nodes = ParseIgpBomCsvToNodes(igpCsvPath, maxDepth, big5, progress, cancellationToken);
                if (nodes.Count > 0)
                {
                    notes = "iGP CSV 編碼：Big5(950)（UTF-8 解析不到節點，已自動回退）";
                    AssignRootLabels(nodes);
                    return nodes;
                }
            }
            catch
            {
                // ignore and fall through
            }

            notes = "警告：iGP CSV 未解析到任何節點。可能原因：檔案編碼、符號（◎□■◇◆○●△）不一致、或格式不符。";
            return nodes;
        }

        private static void AssignRootLabels(List<IgpBomNode> nodes)
        {
            if (nodes == null || nodes.Count == 0)
            {
                return;
            }

            string currentRoot = string.Empty;
            foreach (var n in nodes.OrderBy(x => x.RowOrder))
            {
                if (n == null) continue;
                if (n.Depth == 1)
                {
                    currentRoot = (n.PartNumber ?? string.Empty).Trim();
                    n.RootPartNumber = currentRoot;
                }
                else
                {
                    n.RootPartNumber = currentRoot;
                }
            }
        }

        private static List<IgpBomNode> ParseIgpBomCsvToNodes(
            string igpCsvPath,
            int maxDepth,
            Encoding encoding,
            IProgress<ProgressInfo> progress,
            CancellationToken cancellationToken)
        {
            var lines = File.ReadAllLines(igpCsvPath, encoding ?? Encoding.UTF8);
            var nodes = new List<IgpBomNode>(Math.Max(64, lines.Length));
            var total = Math.Max(1, lines.Length);

            for (var i = 0; i < lines.Length; i++)
            {
                cancellationToken.ThrowIfCancellationRequested();

                var line = lines[i] ?? string.Empty;
                if (string.IsNullOrWhiteSpace(line))
                {
                    continue;
                }

                var treeCell = FindTreeCell(line);
                if (string.IsNullOrWhiteSpace(treeCell))
                {
                    continue;
                }

                TryAppendNodeFromTreeCell(treeCell, maxDepth, nodes);

                if ((i % 60) == 0)
                {
                    var pct = (int)Math.Round(20.0 * i / total, MidpointRounding.AwayFromZero);
                    progress?.Report(new ProgressInfo(pct, $"解析 iGP BOM（{i + 1}/{lines.Length}）…"));
                }
            }

            return nodes;
        }

        private static string FindTreeCell(string csvLine)
        {
            // iGP 的 CSV 常在前面有多個空欄位，我們找第一個包含 '[' 的欄位即為 tree cell
            var parts = csvLine.Split(',');
            foreach (var p in parts)
            {
                if (p != null && p.IndexOf('[') >= 0)
                {
                    return p.Trim();
                }
            }

            return string.Empty;
        }

        /// <summary>若 tree cell 可辨識為有效階層列則加入 <paramref name="nodes"/>。</summary>
        private static void TryAppendNodeFromTreeCell(string treeCell, int maxDepth, List<IgpBomNode> nodes)
        {
            var depth = TryMapDepthFromSymbol(treeCell);
            if (depth == null || depth.Value <= 0 || depth.Value > maxDepth)
            {
                return;
            }

            var partNumber = ExtractPartNumber(treeCell);
            if (string.IsNullOrWhiteSpace(partNumber))
            {
                return;
            }

            var qty = ExtractIgpQty(treeCell);
            nodes.Add(new IgpBomNode
            {
                RowOrder = nodes.Count,
                Depth = depth.Value,
                PartNumber = partNumber,
                Qty = qty
            });
        }

        private static int? TryMapDepthFromSymbol(string treeCell)
        {
            var s = (treeCell ?? string.Empty).TrimStart();
            if (s.Length < 2)
            {
                return null;
            }

            var symbol2 = s.Substring(0, 2);
            return symbol2 switch
            {
                "◎□" => 1,
                "■◇" => 2,
                "◆○" => 3,
                "●△" => 4,
                _ => null
            };
        }

        private static string ExtractPartNumber(string treeCell)
        {
            var s = (treeCell ?? string.Empty).Trim();
            var bracketIdx = s.IndexOf('[');
            var prefix = bracketIdx >= 0 ? s.Substring(0, bracketIdx) : s;
            prefix = prefix.Trim();
            if (prefix.Length == 0)
            {
                return string.Empty;
            }

            // prefix 形式："{符號}{可能空白}{料號}"，取最後一段 token
            var tokens = prefix.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
            if (tokens.Length == 0)
            {
                return string.Empty;
            }

            return tokens[tokens.Length - 1].Trim();
        }

        private static decimal ExtractIgpQty(string treeCell)
        {
            // 欄位順序：[品名][規格][庫存數][組成用量][底數]
            var matches = BracketRegex.Matches(treeCell ?? string.Empty);
            if (matches.Count < 4)
            {
                return 0m;
            }

            var raw = matches[3].Groups[1].Value?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(raw))
            {
                return 0m;
            }

            // iGP 數量多為整數；保守起見也支援小數
            if (decimal.TryParse(raw, NumberStyles.Number, CultureInfo.InvariantCulture, out var d))
            {
                return d;
            }
            if (decimal.TryParse(raw, NumberStyles.Number, CultureInfo.CurrentCulture, out d))
            {
                return d;
            }

            return 0m;
        }

        private static List<IgpBomNode> BuildPdmNodes(IReadOnlyList<BomItem> items, int maxDepth, bool excludeDrawings)
        {
            var nodes = new List<IgpBomNode>(items.Count);
            foreach (var it in items)
            {
                if (it == null)
                {
                    continue;
                }

                if (excludeDrawings && it.IsDrawing)
                {
                    continue;
                }

                var pn = ResolvePdmKey(it);
                if (string.IsNullOrWhiteSpace(pn))
                {
                    continue;
                }

                var depth = DeriveDepthFromLevel(it.Level);
                if (depth <= 0 || depth > maxDepth)
                {
                    continue;
                }

                var qty = it.UsageCount.HasValue ? (decimal)it.UsageCount.Value : 0m;
                nodes.Add(new IgpBomNode
                {
                    RowOrder = nodes.Count,
                    Depth = depth,
                    PartNumber = pn,
                    Qty = qty
                });
            }

            AssignRootLabels(nodes);
            return nodes;
        }

        private static string ResolvePdmKey(BomItem item)
        {
            if (item == null)
            {
                return string.Empty;
            }

            var pn = (item.PartNumber ?? string.Empty).Trim();
            if (!string.IsNullOrWhiteSpace(pn))
            {
                return pn;
            }

            // Vault 端若未設定「Part Number/料號」變數，改用常見卡片欄位作為備援 key。
            if (item.CardVariables != null)
            {
                foreach (var k in new[]
                         {
                             "CP_圖號",
                             "圖號",
                             "料號",
                             "品號",
                             "零件編號",
                             "Part Number",
                             "PartNumber"
                         })
                {
                    if (item.CardVariables.TryGetValue(k, out var v))
                    {
                        var s = (v ?? string.Empty).Trim();
                        if (!string.IsNullOrWhiteSpace(s))
                        {
                            return s;
                        }
                    }
                }
            }

            // 最後備援：用檔名（去副檔名），至少能讓樹結構比對跑起來。
            var fn = (item.FileName ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(fn))
            {
                return string.Empty;
            }

            try
            {
                return Path.GetFileNameWithoutExtension(fn)?.Trim() ?? string.Empty;
            }
            catch
            {
                return fn;
            }
        }

        private static int DeriveDepthFromLevel(string level)
        {
            var s = (level ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(s))
            {
                return 0;
            }

            // Level 目前為 0 / 0.1 / 0.1.1 ...
            return s.Split('.').Length;
        }

        private static Dictionary<string, decimal> AggregateByPartNumber(IEnumerable<IgpBomNode> nodes)
        {
            var dict = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
            foreach (var n in nodes)
            {
                if (n == null || string.IsNullOrWhiteSpace(n.PartNumber))
                {
                    continue;
                }

                dict.TryGetValue(n.PartNumber, out var v);
                dict[n.PartNumber] = v + n.Qty;
            }

            return dict;
        }

        private static List<IgpBomStage1DiffRow> CompareStage1(List<IgpBomNode> pdmNodes, List<IgpBomNode> igpNodes)
        {
            // Stage1：只比對第二階（Depth=2）的料號用量；根階不參與 matching。
            var pdm2 = pdmNodes?.Where(n => n != null && n.Depth == 2).ToList() ?? new List<IgpBomNode>();
            var igp2 = igpNodes?.Where(n => n != null && n.Depth == 2).ToList() ?? new List<IgpBomNode>();

            var pdmAgg = AggregateByPartNumber(pdm2);
            var igpAgg = AggregateByPartNumber(igp2);

            var keys = new HashSet<string>(pdmAgg.Keys, StringComparer.OrdinalIgnoreCase);
            keys.UnionWith(igpAgg.Keys);

            var rows = new List<IgpBomStage1DiffRow>();
            foreach (var key in keys.OrderBy(k => k, StringComparer.OrdinalIgnoreCase))
            {
                pdmAgg.TryGetValue(key, out var p);
                igpAgg.TryGetValue(key, out var i);
                if (p == 0m && i == 0m)
                {
                    continue;
                }

                var rootPdm = string.Join(";", pdm2
                    .Where(n => string.Equals(n.PartNumber, key, StringComparison.OrdinalIgnoreCase)
                                && !string.IsNullOrWhiteSpace(n.RootPartNumber))
                    .Select(n => n.RootPartNumber)
                    .Distinct(StringComparer.OrdinalIgnoreCase));
                var rootIgp = string.Join(";", igp2
                    .Where(n => string.Equals(n.PartNumber, key, StringComparison.OrdinalIgnoreCase)
                                && !string.IsNullOrWhiteSpace(n.RootPartNumber))
                    .Select(n => n.RootPartNumber)
                    .Distinct(StringComparer.OrdinalIgnoreCase));

                if (p == i)
                {
                    rows.Add(new IgpBomStage1DiffRow
                    {
                        RootPdm = rootPdm,
                        RootIgp = rootIgp,
                        PartNumber = key,
                        PdmQty = p,
                        IgpQty = i,
                        DiffType = "Match"
                    });
                    continue;
                }

                var type = p == 0m ? "OnlyiGP" : (i == 0m ? "OnlyPDM" : "Mismatch");
                rows.Add(new IgpBomStage1DiffRow
                {
                    RootPdm = rootPdm,
                    RootIgp = rootIgp,
                    PartNumber = key,
                    PdmQty = p,
                    IgpQty = i,
                    DiffType = type
                });
            }

            return rows;
        }

        private sealed class EdgeAgg
        {
            public decimal QtySum { get; set; }
            public HashSet<string> RootParts { get; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        }

        private static Dictionary<IgpBomEdgeKey, EdgeAgg> BuildEdges(List<IgpBomNode> nodes)
        {
            // stack: depth -> last seen partnumber at that depth
            var lastAtDepth = new Dictionary<int, string>();
            var dict = new Dictionary<IgpBomEdgeKey, EdgeAgg>();

            foreach (var n in nodes.OrderBy(x => x.RowOrder))
            {
                if (n.Depth <= 1)
                {
                    lastAtDepth[n.Depth] = n.PartNumber;
                    continue;
                }

                if (!lastAtDepth.TryGetValue(n.Depth - 1, out var parent) || string.IsNullOrWhiteSpace(parent))
                {
                    // 找不到 parent，略過（資料不完整或深度跳階）
                    lastAtDepth[n.Depth] = n.PartNumber;
                    continue;
                }

                // 根階（Depth=1 -> Depth=2 的邊）不參與比較：
                // 也就是忽略 Child=Depth=2 的那一層邊。
                // 這樣即使兩邊根的料號不同，仍能比對 Depth>=2 的子樹結構。
                if (n.Depth == 2)
                {
                    lastAtDepth[n.Depth] = n.PartNumber; // 仍要更新，用於 Depth=3 的 parent 推導
                    continue;
                }

                var k = new IgpBomEdgeKey(parent, n.PartNumber);
                if (!dict.TryGetValue(k, out var agg))
                {
                    agg = new EdgeAgg();
                    dict[k] = agg;
                }
                agg.QtySum += n.Qty; // 選項 A：用子節點自身數量
                if (!string.IsNullOrWhiteSpace(n.RootPartNumber))
                {
                    agg.RootParts.Add(n.RootPartNumber);
                }

                lastAtDepth[n.Depth] = n.PartNumber;
            }

            return dict;
        }

        private static List<IgpBomStage2DiffRow> CompareStage2(
            Dictionary<IgpBomEdgeKey, EdgeAgg> pdmEdges,
            Dictionary<IgpBomEdgeKey, EdgeAgg> igpEdges)
        {
            var keys = new HashSet<IgpBomEdgeKey>(pdmEdges.Keys);
            keys.UnionWith(igpEdges.Keys);

            var rows = new List<IgpBomStage2DiffRow>();
            foreach (var k in keys)
            {
                pdmEdges.TryGetValue(k, out var p);
                igpEdges.TryGetValue(k, out var i);
                var pQty = p?.QtySum ?? 0m;
                var iQty = i?.QtySum ?? 0m;

                if (pQty == iQty)
                {
                    rows.Add(new IgpBomStage2DiffRow
                    {
                        RootPdm = string.Join(";", p?.RootParts?.Distinct(StringComparer.OrdinalIgnoreCase)
                                                   ?? Array.Empty<string>()),
                        RootIgp = string.Join(";", i?.RootParts?.Distinct(StringComparer.OrdinalIgnoreCase)
                                                    ?? Array.Empty<string>()),
                        ParentPartNumber = k.Parent,
                        ChildPartNumber = k.Child,
                        PdmQty = pQty,
                        IgpQty = iQty,
                        DiffType = "Match"
                    });
                    continue;
                }

                var type = pQty == 0m ? "OnlyiGPEdge" : (iQty == 0m ? "OnlyPDMEdge" : "Mismatch");
                rows.Add(new IgpBomStage2DiffRow
                {
                    RootPdm = string.Join(";", p?.RootParts?.Distinct(StringComparer.OrdinalIgnoreCase)
                                           ?? Array.Empty<string>()),
                    RootIgp = string.Join(";", i?.RootParts?.Distinct(StringComparer.OrdinalIgnoreCase)
                                                ?? Array.Empty<string>()),
                    ParentPartNumber = k.Parent,
                    ChildPartNumber = k.Child,
                    PdmQty = pQty,
                    IgpQty = iQty,
                    DiffType = type
                });
            }

            return rows
                .OrderBy(r => r.ParentPartNumber, StringComparer.OrdinalIgnoreCase)
                .ThenBy(r => r.ChildPartNumber, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }
    }
}

