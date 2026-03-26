using System;
using System.Collections.Generic;

namespace PDMTools.Models
{
    /// <summary>
    /// 結果 DataGrid 用的欄篩選狀態（與主視窗 Vault BOM 相同邏輯，ValueGetter 改為 object）。</summary>
    public sealed class ReportColumnFilterState
    {
        public const string BlankDisplayText = "(空白)";

        public string Key { get; set; } = string.Empty;
        public Func<object, string> ValueGetter { get; set; } = _ => string.Empty;
        public List<string> AvailableValues { get; set; } = new List<string>();
        public HashSet<string> SelectedValues { get; set; } = new HashSet<string>(StringComparer.Ordinal);
        public string SearchText { get; set; } = string.Empty;
    }
}
