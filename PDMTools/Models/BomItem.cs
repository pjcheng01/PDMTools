using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PDMTools.Models
{
    public sealed class BomItem
    {
        public string Level { get; set; } = string.Empty;
        public string FileName { get; set; } = string.Empty;
        public string State { get; set; } = string.Empty;
        public string WorkflowState { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public string PartNumber { get; set; } = string.Empty;
        public string FullPath { get; set; } = string.Empty;
        public string ReferencedAs { get; set; } = string.Empty;
        public string DescriptionVarUsed { get; set; } = string.Empty;
        public string DescriptionConfigUsed { get; set; } = string.Empty;
        public string PartNumberVarUsed { get; set; } = string.Empty;
        public string PartNumberConfigUsed { get; set; } = string.Empty;
        public Dictionary<string, string> CardVariables { get; set; } = new Dictionary<string, string>();

        /// <summary>供預覽用：將非空白卡片變數濃縮為單一字串。</summary>
        public string CardVariablesDisplay
        {
            get
            {
                if (CardVariables == null || CardVariables.Count == 0)
                {
                    return string.Empty;
                }

                var sb = new StringBuilder();
                foreach (var kv in CardVariables.OrderBy(x => x.Key, System.StringComparer.OrdinalIgnoreCase))
                {
                    if (string.IsNullOrWhiteSpace(kv.Value))
                    {
                        continue;
                    }

                    if (sb.Length > 0)
                    {
                        sb.Append("; ");
                    }

                    sb.Append(kv.Key);
                    sb.Append('=');
                    sb.Append(kv.Value);
                }

                return sb.ToString();
            }
        }
    }
}
