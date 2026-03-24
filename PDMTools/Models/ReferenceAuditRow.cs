namespace PDMTools.Models
{
    /// <summary>
    /// 單一頂層組合件經 SolidWorks API 遞迴展開後的一列引用資訊（供稽核／匯出）。
    /// </summary>
    public sealed class ReferenceAuditRow
    {
        public string TopLevelAssemblyPath { get; set; } = string.Empty;
        /// <summary>此引用在樹狀結構中的深度（頂層組合件為 0）。</summary>
        public int Level { get; set; }
        /// <summary>直接上層組合件（.sldasm）的完整路徑；若由多處引用合併則以「; 」分隔。</summary>
        public string ParentAssemblyPath { get; set; } = string.Empty;
        public string FullPath { get; set; } = string.Empty;
        public string FileName { get; set; } = string.Empty;
        public string Extension { get; set; } = string.Empty;
        public bool IsUnderVaultRoot { get; set; }
        public bool FileExists { get; set; }
    }
}
