using System.Collections.Generic;

namespace PDMTools.Models
{
    public sealed class ColumnPresetProfile
    {
        public string Name { get; set; } = string.Empty;
        public List<string> SelectedItems { get; set; } = new List<string>();
    }
}
