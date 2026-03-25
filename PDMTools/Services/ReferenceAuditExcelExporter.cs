using System.Collections.Generic;
using ClosedXML.Excel;
using PDMTools.Models;

namespace PDMTools.Services
{
    public static class ReferenceAuditExcelExporter
    {
        public static void Export(IReadOnlyList<ReferenceAuditRow> rows, string outputPath)
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("ReferenceAudit");

            ws.Cell(1, 1).Value = "頂層組合件";
            ws.Cell(1, 2).Value = "Level";
            ws.Cell(1, 3).Value = "上一層組合件路徑";
            ws.Cell(1, 4).Value = "完整路徑";
            ws.Cell(1, 5).Value = "檔名";
            ws.Cell(1, 6).Value = "副檔名";
            ws.Cell(1, 7).Value = "工程圖列";
            ws.Cell(1, 8).Value = "對應模型路徑";
            ws.Cell(1, 9).Value = "在 Vault 內";
            ws.Cell(1, 10).Value = "檔案存在";

            for (var i = 0; i < rows.Count; i++)
            {
                var r = rows[i];
                var row = i + 2;
                ws.Cell(row, 1).Value = r.TopLevelAssemblyPath;
                ws.Cell(row, 2).Value = r.Level;
                ws.Cell(row, 3).Value = r.ParentAssemblyPath;
                ws.Cell(row, 4).Value = r.FullPath;
                ws.Cell(row, 5).Value = r.FileName;
                ws.Cell(row, 6).Value = r.Extension;
                ws.Cell(row, 7).Value = r.IsDrawing ? "是" : "否";
                ws.Cell(row, 8).Value = r.RelatedModelPath ?? string.Empty;
                ws.Cell(row, 9).Value = r.IsUnderVaultRoot ? "是" : "否";
                ws.Cell(row, 10).Value = r.FileExists ? "是" : "否";
            }

            ws.Row(1).Style.Font.Bold = true;
            ws.Columns().AdjustToContents();
            wb.SaveAs(outputPath);
        }
    }
}
