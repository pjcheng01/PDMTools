using System.Collections.Generic;

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
    }
}
