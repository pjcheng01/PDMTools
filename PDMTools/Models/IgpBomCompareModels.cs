using System;
using System.Collections.Generic;

namespace PDMTools.Models
{
    public sealed class IgpBomCompareResult
    {
        public string IgpSourcePath { get; set; } = string.Empty;
        public string PdmSourcePath { get; set; } = string.Empty;
        public int MaxDepth { get; set; }
        public DateTime ComparedAtLocal { get; set; }

        public int IgpNodeCount { get; set; }
        public int PdmNodeCount { get; set; }
        public int IgpEdgeCount { get; set; }
        public int PdmEdgeCount { get; set; }
        public string Notes { get; set; } = string.Empty;

        public List<IgpBomStage1DiffRow> Stage1Rows { get; set; } = new List<IgpBomStage1DiffRow>();
        public List<IgpBomStage2DiffRow> Stage2Rows { get; set; } = new List<IgpBomStage2DiffRow>();
    }

    public sealed class IgpBomStage1DiffRow
    {
        // 用於顯示階層上下文：這個 PartNumber 出現在哪些根階（PDM / iGP）
        public string RootPdm { get; set; } = string.Empty;
        public string RootIgp { get; set; } = string.Empty;

        public string PartNumber { get; set; } = string.Empty;
        public decimal PdmQty { get; set; }
        public decimal IgpQty { get; set; }
        public decimal Diff => PdmQty - IgpQty;
        public string DiffType { get; set; } = string.Empty; // OnlyPDM / OnlyiGP / Mismatch / Match
    }

    public sealed class IgpBomStage2DiffRow
    {
        // 用於顯示階層上下文：這條邊（Parent->Child）出現在哪些根階（PDM / iGP）
        public string RootPdm { get; set; } = string.Empty;
        public string RootIgp { get; set; } = string.Empty;

        public string ParentPartNumber { get; set; } = string.Empty;
        public string ChildPartNumber { get; set; } = string.Empty;
        public decimal PdmQty { get; set; }
        public decimal IgpQty { get; set; }
        public decimal Diff => PdmQty - IgpQty;
        public string DiffType { get; set; } = string.Empty; // OnlyPDMEdge / OnlyiGPEdge / Mismatch / Match
    }

    internal sealed class IgpBomNode
    {
        public int RowOrder { get; set; }
        public int Depth { get; set; }
        public string RootPartNumber { get; set; } = string.Empty;
        public string PartNumber { get; set; } = string.Empty;
        public decimal Qty { get; set; }
    }

    internal readonly struct IgpBomEdgeKey : IEquatable<IgpBomEdgeKey>
    {
        public readonly string Parent;
        public readonly string Child;

        public IgpBomEdgeKey(string parent, string child)
        {
            Parent = parent ?? string.Empty;
            Child = child ?? string.Empty;
        }

        public bool Equals(IgpBomEdgeKey other) =>
            string.Equals(Parent, other.Parent, StringComparison.OrdinalIgnoreCase) &&
            string.Equals(Child, other.Child, StringComparison.OrdinalIgnoreCase);

        public override bool Equals(object obj) => obj is IgpBomEdgeKey other && Equals(other);

        public override int GetHashCode()
        {
            return StringComparer.OrdinalIgnoreCase.GetHashCode(Parent) * 397 ^
                   StringComparer.OrdinalIgnoreCase.GetHashCode(Child);
        }
    }
}

