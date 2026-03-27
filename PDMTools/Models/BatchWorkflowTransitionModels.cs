using System.Collections.Generic;

namespace PDMTools.Models
{
    /// <summary>單一轉換選項（來自 IEdmBatchChangeState*.GetAvailableTransitionList）。</summary>
    public sealed class PdmTransitionOption
    {
        public int TransitionId { get; set; }
        public string TransitionName { get; set; } = string.Empty;
        public string TargetStateName { get; set; } = string.Empty;

        /// <summary>
        /// 供 ComboBox 顯示：以「轉換後的目標狀態」為主（與 Explorer「State／本機狀態」一致）；
        /// 「動作」為 PDM 轉換名稱（常與流程或步驟同名，勿與「Workflow」欄混淆）。
        /// </summary>
        public string DisplayText
        {
            get
            {
                var idSeg = TransitionId != 0 ? $"[轉換 ID {TransitionId}] " : string.Empty;
                var action = (TransitionName ?? string.Empty).Trim();
                var target = (TargetStateName ?? string.Empty).Trim();

                if (!string.IsNullOrEmpty(target) && !string.IsNullOrEmpty(action))
                {
                    return idSeg + "目標狀態「" + target + "」｜動作：" + action;
                }

                if (!string.IsNullOrEmpty(target))
                {
                    return idSeg + "目標狀態「" + target + "」";
                }

                if (!string.IsNullOrEmpty(action))
                {
                    return idSeg + "動作：" + action + "（僅轉換名；未解析到目標狀態，易與流程名混淆）";
                }

                return TransitionId != 0 ? $"[轉換 ID {TransitionId}]" : "(無法讀取轉換／目標狀態)";
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
