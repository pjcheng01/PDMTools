using System;
using System.Collections.Generic;

namespace PDMTools.Models
{
    internal sealed class ColumnFilterState
    {
        public const string BlankDisplayText = "(空白)";

        public string Key { get; set; } = string.Empty;
        public Func<BomItem, string> ValueGetter { get; set; } = _ => string.Empty;
        public List<string> AvailableValues { get; set; } = new List<string>();
        public HashSet<string> SelectedValues { get; set; } = new HashSet<string>(StringComparer.Ordinal);
        public string SearchText { get; set; } = string.Empty;
    }
}
