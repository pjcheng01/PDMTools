using System.Collections.Generic;
using System.Text;

namespace PDMTools.Models
{
    /// <summary>單一轉換選項（來自 IEdmBatchChangeState*.GetAvailableTransitionList）。</summary>
    public sealed class PdmTransitionOption
    {
        public int TransitionId { get; set; }
        /// <summary>顯示用（優先 moName，見 Analyze）。</summary>
        public string TransitionName { get; set; } = string.Empty;
        /// <summary>結構欄位 <c>mbsTransitionName</c> 原文；<c>CreateTree</c>／<c>ChangeState2</c> 宜優先使用。</summary>
        public string MbsTransitionNameRaw { get; set; } = string.Empty;
        public string TargetStateName { get; set; } = string.Empty;

        /// <summary>
        /// 供 ComboBox 顯示：僅使用與 PDM Interop 結構（如 EdmChangeStateTransitionInfo）一致的原始欄位名稱。
        /// </summary>
        public string DisplayText
        {
            get
            {
                var sb = new StringBuilder();
                sb.Append("mlTransitionID=").Append(TransitionId);
                sb.Append("; mbsTransitionName=").Append(TransitionName ?? string.Empty);
                sb.Append("; mbsTargetStateName=").Append(TargetStateName ?? string.Empty);
                return sb.ToString();
            }
        }
    }

    /// <summary>同一狀態（Vault 內狀態 ID）之檔案群組與可選轉換。</summary>
    public sealed class BatchWorkflowTransitionGroup
    {
        public int StateId { get; set; }
        public string StateName { get; set; } = string.Empty;
        public string WorkflowName { get; set; } = string.Empty;
        public List<string> FilePaths { get; } = new List<string>();
        public List<PdmTransitionOption> AvailableTransitions { get; } = new List<PdmTransitionOption>();
        public string AnalyzeError { get; set; } = string.Empty;
    }

    public sealed class BatchWorkflowTransitionAnalyzeResult
    {
        public List<BatchWorkflowTransitionGroup> Groups { get; } = new List<BatchWorkflowTransitionGroup>();
        public List<string> SkippedPaths { get; } = new List<string>();
    }

    public sealed class BatchWorkflowTransitionExecuteRow
    {
        public string FullPath { get; set; } = string.Empty;
        public bool Success { get; set; }
        public string Message { get; set; } = string.Empty;
    }

    public sealed class BatchWorkflowTransitionExecuteResult
    {
        public List<BatchWorkflowTransitionExecuteRow> Rows { get; } = new List<BatchWorkflowTransitionExecuteRow>();
    }
}
