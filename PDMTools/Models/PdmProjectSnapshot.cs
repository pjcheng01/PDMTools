using System;
using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace PDMTools.Models
{
    /// <summary>根物件：PDMTools 專案／快照檔（JSON）。</summary>
    public sealed class PdmProjectSnapshotDocument
    {
        public int FormatVersion { get; set; } = 1;

        /// <summary>快照儲存時間（使用者本機區域時間）。</summary>
        public DateTime SavedAtLocal { get; set; }

        /// <summary>舊版快照 JSON 欄位，僅供反序列化相容；新檔不寫入。</summary>
        [JsonPropertyName("savedAtUtc")]
        public DateTime? SavedAtUtcLegacy { get; set; }
        /// <summary>組合件 .sldasm 完整路徑（兩種模式共用）。</summary>
        public string AssemblyPath { get; set; } = string.Empty;
        /// <summary>"Bom" 或 "ReferenceAudit"</summary>
        public string WorkMode { get; set; } = "Bom";

        public BomSnapshotData Bom { get; set; }
        public ReferenceAuditSnapshotData ReferenceAudit { get; set; }
    }

    public sealed class BomSnapshotData
    {
        public int BomLevelDefinitionIndex { get; set; }
        public bool AllBomDepth { get; set; }
        public string MaxBomDepthText { get; set; } = "3";
        public string ConfigurationName { get; set; } = string.Empty;
        public bool? ShowDrawingsChecked { get; set; }

        public List<string> ActiveCardVarNames { get; set; } = new List<string>();
        public List<string> ActiveFixedColumns { get; set; }

        public List<BomItem> Items { get; set; } = new List<BomItem>();
        public List<BomItem> RawBomItems { get; set; } = new List<BomItem>();
        public List<BomItem> BomItemsWithDrawings { get; set; }

        public List<BomColumnFilterSnapshot> ColumnFilters { get; set; } = new List<BomColumnFilterSnapshot>();
    }

    public sealed class BomColumnFilterSnapshot
    {
        public string Key { get; set; } = string.Empty;
        public List<string> SelectedValues { get; set; } = new List<string>();
        public string SearchText { get; set; } = string.Empty;
    }

    public sealed class ReferenceAuditSnapshotData
    {
        public string ReferenceAuditMaxDepthText { get; set; } = "3";
        public bool ReferenceAuditAllDepth { get; set; }
        public bool ReferenceAuditIncludeDrawings { get; set; }
        public int PreviewFilterComboIndex { get; set; }

        public List<ReferenceAuditRow> Rows { get; set; } = new List<ReferenceAuditRow>();
    }
}
