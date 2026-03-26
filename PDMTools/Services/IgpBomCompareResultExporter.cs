using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using ClosedXML.Excel;
using PDMTools.Models;

namespace PDMTools.Services
{
    /// <summary>將 <see cref="IgpBomCompareResult"/> 匯出為 .xlsx 或 .json（完整比對列，不含畫面篩選）。</summary>
    public static class IgpBomCompareResultExporter
    {
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            WriteIndented = true,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
        };

        /// <summary>Windows 上「微軟正黑體」的字型名稱。</summary>
        private const string ExcelFontName = "Microsoft JhengHei";
        private const double ExcelFontSize = 12;
        private const double ExcelRowHeightPoints = 18;

        // 與 IgpBomCompareWindow 列前景一致（ARGB）
        private static readonly XLColor ColorDefault = XLColor.FromArgb(unchecked((int)0xFF212121));
        private static readonly XLColor ColorOnlyPdm = XLColor.FromArgb(unchecked((int)0xFFC62828));
        private static readonly XLColor ColorOnlyIgp = XLColor.FromArgb(unchecked((int)0xFF1565C0));
        private static readonly XLColor ColorMismatch = XLColor.FromArgb(unchecked((int)0xFF2E7D32));

        public static void Save(IgpBomCompareResult result, string path)
        {
            if (result == null)
            {
                throw new ArgumentNullException(nameof(result));
            }

            if (string.IsNullOrWhiteSpace(path))
            {
                throw new ArgumentException("路徑不可為空。", nameof(path));
            }

            var ext = Path.GetExtension(path);
            if (string.Equals(ext, ".json", StringComparison.OrdinalIgnoreCase))
            {
                SaveJson(result, path);
            }
            else
            {
                SaveExcel(result, path);
            }
        }

        public static void SaveJson(IgpBomCompareResult result, string path)
        {
            var json = JsonSerializer.Serialize(result, JsonOptions);
            File.WriteAllText(path, json, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        }

        public static void SaveExcel(IgpBomCompareResult result, string path)
        {
            using var wb = new XLWorkbook();

            BuildSummarySheet(wb, result);
            BuildStage1Sheet(wb, result);
            BuildStage2Sheet(wb, result);

            wb.SaveAs(path);
        }

        private static void BuildSummarySheet(XLWorkbook wb, IgpBomCompareResult result)
        {
            var summary = wb.Worksheets.Add("摘要");
            summary.Cell(1, 1).Value = "項目";
            summary.Cell(1, 2).Value = "內容";
            summary.Row(1).Style.Font.Bold = true;

            var row = 2;
            void Pair(string label, string value)
            {
                summary.Cell(row, 1).Value = label;
                summary.Cell(row, 2).Value = value ?? string.Empty;
                row++;
            }

            Pair("iGP 來源", result.IgpSourcePath);
            Pair("PDM 來源", result.PdmSourcePath);
            Pair("深度上限", result.MaxDepth.ToString(CultureInfo.InvariantCulture));
            Pair("比對時間（本機）", result.ComparedAtLocal.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.CurrentCulture));
            Pair("iGP Nodes", result.IgpNodeCount.ToString(CultureInfo.InvariantCulture));
            Pair("PDM Nodes", result.PdmNodeCount.ToString(CultureInfo.InvariantCulture));
            Pair("iGP Edges", result.IgpEdgeCount.ToString(CultureInfo.InvariantCulture));
            Pair("PDM Edges", result.PdmEdgeCount.ToString(CultureInfo.InvariantCulture));
            Pair("備註", result.Notes ?? string.Empty);

            var lastRow = row - 1;
            if (lastRow >= 2)
            {
                var tbl = summary.Range(1, 1, lastRow, 2).CreateTable("TblSummary");
                tbl.Theme = XLTableTheme.TableStyleMedium2;
                tbl.ShowTotalsRow = false;
            }

            ApplySheetLayout(summary);
        }

        private static void BuildStage1Sheet(XLWorkbook wb, IgpBomCompareResult result)
        {
            const int colCount = 7;
            var s1 = wb.Worksheets.Add("Stage1_料號彙總");
            s1.Cell(1, 1).Value = "料號";
            s1.Cell(1, 2).Value = "根(PDM)";
            s1.Cell(1, 3).Value = "根(iGP)";
            s1.Cell(1, 4).Value = "PDM 總量";
            s1.Cell(1, 5).Value = "iGP 總量";
            s1.Cell(1, 6).Value = "差異";
            s1.Cell(1, 7).Value = "類型";
            s1.Row(1).Style.Font.Bold = true;

            var r1 = 2;
            foreach (var x in result.Stage1Rows ?? new List<IgpBomStage1DiffRow>())
            {
                if (x == null)
                {
                    continue;
                }

                s1.Cell(r1, 1).Value = x.PartNumber;
                s1.Cell(r1, 2).Value = x.RootPdm;
                s1.Cell(r1, 3).Value = x.RootIgp;
                s1.Cell(r1, 4).Value = (double)x.PdmQty;
                s1.Cell(r1, 5).Value = (double)x.IgpQty;
                s1.Cell(r1, 6).Value = (double)x.Diff;
                s1.Cell(r1, 7).Value = x.DiffType;
                r1++;
            }

            var lastRow = r1 - 1;
            if (lastRow >= 2)
            {
                var tbl = s1.Range(1, 1, lastRow, colCount).CreateTable("TblStage1Bom");
                tbl.Theme = XLTableTheme.TableStyleMedium2;
                tbl.ShowTotalsRow = false;
            }

            ApplySheetLayout(s1);

            if (lastRow >= 2)
            {
                for (var r = 2; r <= lastRow; r++)
                {
                    var dt = s1.Cell(r, 7).GetString();
                    ApplyRowFontColor(s1, r, 1, colCount, GetFontColorForDiffTypeStage1(dt));
                }
            }
        }

        private static void BuildStage2Sheet(XLWorkbook wb, IgpBomCompareResult result)
        {
            const int colCount = 8;
            var s2 = wb.Worksheets.Add("Stage2_父子結構");
            s2.Cell(1, 1).Value = "根(PDM)";
            s2.Cell(1, 2).Value = "根(iGP)";
            s2.Cell(1, 3).Value = "父料號";
            s2.Cell(1, 4).Value = "子料號";
            s2.Cell(1, 5).Value = "PDM 用量";
            s2.Cell(1, 6).Value = "iGP 用量";
            s2.Cell(1, 7).Value = "差異";
            s2.Cell(1, 8).Value = "類型";
            s2.Row(1).Style.Font.Bold = true;

            var r2 = 2;
            foreach (var x in result.Stage2Rows ?? new List<IgpBomStage2DiffRow>())
            {
                if (x == null)
                {
                    continue;
                }

                s2.Cell(r2, 1).Value = x.RootPdm;
                s2.Cell(r2, 2).Value = x.RootIgp;
                s2.Cell(r2, 3).Value = x.ParentPartNumber;
                s2.Cell(r2, 4).Value = x.ChildPartNumber;
                s2.Cell(r2, 5).Value = (double)x.PdmQty;
                s2.Cell(r2, 6).Value = (double)x.IgpQty;
                s2.Cell(r2, 7).Value = (double)x.Diff;
                s2.Cell(r2, 8).Value = x.DiffType;
                r2++;
            }

            var lastRow = r2 - 1;
            if (lastRow >= 2)
            {
                var tbl = s2.Range(1, 1, lastRow, colCount).CreateTable("TblStage2Edges");
                tbl.Theme = XLTableTheme.TableStyleMedium2;
                tbl.ShowTotalsRow = false;
            }

            ApplySheetLayout(s2);

            if (lastRow >= 2)
            {
                for (var r = 2; r <= lastRow; r++)
                {
                    var dt = s2.Cell(r, 8).GetString();
                    ApplyRowFontColor(s2, r, 1, colCount, GetFontColorForDiffTypeStage2(dt));
                }
            }
        }

        /// <summary>微軟正黑體 12pt、列高 18、欄寬自動；第 1 列維持粗體（標題列）。</summary>
        private static void ApplySheetLayout(IXLWorksheet ws)
        {
            var used = ws.RangeUsed();
            if (used == null)
            {
                return;
            }

            used.Style.Font.FontName = ExcelFontName;
            used.Style.Font.FontSize = ExcelFontSize;

            var firstRow = used.RangeAddress.FirstAddress.RowNumber;
            var lastRow = used.RangeAddress.LastAddress.RowNumber;
            for (var r = firstRow; r <= lastRow; r++)
            {
                ws.Row(r).Height = ExcelRowHeightPoints;
            }

            ws.Columns().AdjustToContents();
            ws.Row(1).Style.Font.Bold = true;
        }

        private static void ApplyRowFontColor(IXLWorksheet ws, int row, int c1, int c2, XLColor color)
        {
            for (var c = c1; c <= c2; c++)
            {
                ws.Cell(row, c).Style.Font.FontColor = color;
            }
        }

        private static XLColor GetFontColorForDiffTypeStage1(string diffType)
        {
            var t = diffType ?? string.Empty;
            if (string.Equals(t, "OnlyPDM", StringComparison.OrdinalIgnoreCase))
            {
                return ColorOnlyPdm;
            }

            if (string.Equals(t, "OnlyiGP", StringComparison.OrdinalIgnoreCase))
            {
                return ColorOnlyIgp;
            }

            if (string.Equals(t, "Mismatch", StringComparison.OrdinalIgnoreCase))
            {
                return ColorMismatch;
            }

            return ColorDefault;
        }

        private static XLColor GetFontColorForDiffTypeStage2(string diffType)
        {
            var t = diffType ?? string.Empty;
            if (string.Equals(t, "OnlyPDMEdge", StringComparison.OrdinalIgnoreCase))
            {
                return ColorOnlyPdm;
            }

            if (string.Equals(t, "OnlyiGPEdge", StringComparison.OrdinalIgnoreCase))
            {
                return ColorOnlyIgp;
            }

            if (string.Equals(t, "Mismatch", StringComparison.OrdinalIgnoreCase))
            {
                return ColorMismatch;
            }

            return ColorDefault;
        }
    }
}
