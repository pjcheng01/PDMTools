using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using EPDM.Interop.epdm;
using PDMTools.Models;
using PDMTools.Utils;

namespace PDMTools.Services
{
    /// <summary>
    /// 依 Vault BOM 篩選後路徑，以 PDM <see cref="EdmUtility.EdmUtil_BatchChangeState"/> 分析群組並批次轉換狀態。
    /// COM 呼叫須在 STA 執行緒（WPF UI 執行緒）。
    /// </summary>
    public sealed class PdmBatchWorkflowTransitionService
    {
        public BatchWorkflowTransitionAnalyzeResult Analyze(
            IReadOnlyList<string> fullPaths,
            IEdmVault5 vault)
        {
            var result = new BatchWorkflowTransitionAnalyzeResult();
            if (vault == null)
            {
                throw new ArgumentNullException(nameof(vault));
            }

            var vault7 = vault as IEdmVault7;
            if (vault7 == null)
            {
                throw new InvalidOperationException("無法將 Vault 轉型為 IEdmVault7，無法建立批次轉狀態工具。");
            }

            var distinct = DeduplicatePaths(fullPaths);
            var fileRows = new List<FileStateRow>();

            foreach (var path in distinct)
            {
                if (!TryResolveFileRow(vault, path, out var row, out var err))
                {
                    result.SkippedPaths.Add(string.IsNullOrWhiteSpace(err) ? path : $"{path}（{err}）");
                    continue;
                }

                fileRows.Add(row);
            }

            foreach (var g in fileRows.GroupBy(r => r.StateId))
            {
                var first = g.First();
                var group = new BatchWorkflowTransitionGroup
                {
                    StateId = first.StateId,
                    StateName = first.StateName,
                    WorkflowName = first.WorkflowName
                };
                foreach (var r in g)
                {
                    group.FilePaths.Add(r.FullPath);
                }

                try
                {
                    FillAvailableTransitions(vault7, group);
                }
                catch (Exception ex)
                {
                    group.AnalyzeError = ex.Message;
                }

                result.Groups.Add(group);
            }

            result.Groups.Sort((a, b) => string.Compare(
                a.WorkflowName + "|" + a.StateName,
                b.WorkflowName + "|" + b.StateName,
                StringComparison.OrdinalIgnoreCase));

            return result;
        }

        public BatchWorkflowTransitionExecuteResult ExecuteGroup(
            IEdmVault5 vault,
            IReadOnlyList<string> filePathsInGroup,
            PdmTransitionOption transition,
            int parentWindowHandle)
        {
            if (transition == null)
            {
                throw new ArgumentNullException(nameof(transition));
            }

            var result = new BatchWorkflowTransitionExecuteResult();
            var vault7 = vault as IEdmVault7;
            if (vault7 == null)
            {
                throw new InvalidOperationException("無法將 Vault 轉型為 IEdmVault7。");
            }

            object batchObj = null;
            try
            {
                batchObj = vault7.CreateUtility(EdmUtility.EdmUtil_BatchChangeState);
                if (batchObj == null)
                {
                    throw new InvalidOperationException("CreateUtility(EdmUtil_BatchChangeState) 回傳 null。");
                }

                if (!TryAddFilesToBatch(batchObj, vault, filePathsInGroup, result))
                {
                    return result;
                }

                var transitionName = PickTransitionNameForChangeState2(batchObj, transition);
                if (string.IsNullOrWhiteSpace(transitionName))
                {
                    foreach (var p in filePathsInGroup)
                    {
                        result.Rows.Add(new BatchWorkflowTransitionExecuteRow
                        {
                            FullPath = p,
                            Success = false,
                            Message = "無法解析轉換名稱（ChangeState2 需要名稱字串）。"
                        });
                    }

                    return result;
                }

                InvokeChangeState2(batchObj, parentWindowHandle, transitionName);

                foreach (var p in filePathsInGroup)
                {
                    result.Rows.Add(new BatchWorkflowTransitionExecuteRow
                    {
                        FullPath = p,
                        Success = true,
                        Message = "已送出轉換。"
                    });
                }
            }
            catch (Exception ex)
            {
                foreach (var p in filePathsInGroup)
                {
                    if (result.Rows.Any(r => string.Equals(r.FullPath, p, StringComparison.OrdinalIgnoreCase)))
                    {
                        continue;
                    }

                    result.Rows.Add(new BatchWorkflowTransitionExecuteRow
                    {
                        FullPath = p,
                        Success = false,
                        Message = ex.Message
                    });
                }

                if (result.Rows.Count == 0)
                {
                    foreach (var p in filePathsInGroup)
                    {
                        result.Rows.Add(new BatchWorkflowTransitionExecuteRow
                        {
                            FullPath = p,
                            Success = false,
                            Message = ex.Message
                        });
                    }
                }
            }
            finally
            {
                ComHelper.Release(batchObj);
            }

            return result;
        }

        private static string PickTransitionNameForChangeState2(object batchObj, PdmTransitionOption transition)
        {
            if (!string.IsNullOrWhiteSpace(transition.TransitionName))
            {
                return transition.TransitionName.Trim();
            }

            // 後備：從目前批次再取清單對應 ID
            EdmChangeStateTransitionInfo[] infos = null;
            try
            {
                InvokeGetAvailableTransitionList(batchObj, out infos);
            }
            catch
            {
                return string.Empty;
            }

            if (infos == null)
            {
                return string.Empty;
            }

            foreach (var info in infos)
            {
                TryReadTransitionFields(info, out var id, out var name, out _);
                if (id == transition.TransitionId && !string.IsNullOrWhiteSpace(name))
                {
                    return name.Trim();
                }
            }

            return string.Empty;
        }

        private static bool TryAddFilesToBatch(
            object batchObj,
            IEdmVault5 vault,
            IReadOnlyList<string> filePathsInGroup,
            BatchWorkflowTransitionExecuteResult result)
        {
            foreach (var path in filePathsInGroup)
            {
                IEdmFolder5 folder = null;
                IEdmFile5 file = null;
                try
                {
                    file = vault.GetFileFromPath(path, out folder);
                    if (file == null || folder == null)
                    {
                        result.Rows.Add(new BatchWorkflowTransitionExecuteRow
                        {
                            FullPath = path,
                            Success = false,
                            Message = "GetFileFromPath 失敗。"
                        });
                        return false;
                    }

                    InvokeAddFile(batchObj, file.ID, folder.ID);
                }
                catch (Exception ex)
                {
                    result.Rows.Add(new BatchWorkflowTransitionExecuteRow
                    {
                        FullPath = path,
                        Success = false,
                        Message = ex.Message
                    });
                    return false;
                }
                finally
                {
                    ComHelper.Release(file);
                    ComHelper.Release(folder);
                }
            }

            return true;
        }

        private static void FillAvailableTransitions(IEdmVault7 vault7, BatchWorkflowTransitionGroup group)
        {
            object batchObj = null;
            try
            {
                batchObj = vault7.CreateUtility(EdmUtility.EdmUtil_BatchChangeState);
                if (batchObj == null)
                {
                    group.AnalyzeError = "無法建立批次轉狀態工具。";
                    return;
                }

                foreach (var path in group.FilePaths)
                {
                    IEdmFolder5 folder = null;
                    IEdmFile5 file = null;
                    try
                    {
                        file = vault7.GetFileFromPath(path, out folder);
                        if (file == null || folder == null)
                        {
                            group.AnalyzeError = "部分路徑無法解析為 Vault 檔案：" + path;
                            return;
                        }

                        InvokeAddFile(batchObj, file.ID, folder.ID);
                    }
                    finally
                    {
                        ComHelper.Release(file);
                        ComHelper.Release(folder);
                    }
                }

                EdmChangeStateTransitionInfo[] infos = null;
                try
                {
                    InvokeGetAvailableTransitionList(batchObj, out infos);
                }
                catch
                {
                    infos = null;
                }

                var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                if (infos != null && infos.Length > 0)
                {
                    foreach (var info in infos)
                    {
                        TryReadTransitionFields(info, out var id, out var name, out var target);
                        var key = id + "\u001f" + (name ?? string.Empty);
                        if (!seen.Add(key))
                        {
                            continue;
                        }

                        group.AvailableTransitions.Add(new PdmTransitionOption
                        {
                            TransitionId = id,
                            TransitionName = name ?? string.Empty,
                            TargetStateName = target ?? string.Empty
                        });
                    }
                }

                // 部分 Interop／Vault 回傳的 EdmChangeStateTransitionInfo 無法從反射讀到欄位，會變成 ID=0 且無名稱；
                // 改由目前狀態列舉 IEdmTransition*（與 PDM 主程式一致來源）。
                if (TransitionsListLooksUnusable(group.AvailableTransitions) &&
                    group.FilePaths.Count > 0)
                {
                    group.AvailableTransitions.Clear();
                    seen.Clear();
                    TryFillTransitionsFromCurrentState(vault7, group.FilePaths[0], group.AvailableTransitions, seen);
                }

                if (group.AvailableTransitions.Count == 0 && string.IsNullOrWhiteSpace(group.AnalyzeError))
                {
                    group.AnalyzeError = "此群組找不到任何可用轉換（批次 API 與狀態列舉皆無結果）；請確認權限與工作流程設定。";
                }
            }
            finally
            {
                ComHelper.Release(batchObj);
            }
        }

        private static bool TransitionsListLooksUnusable(List<PdmTransitionOption> list)
        {
            if (list == null || list.Count == 0)
            {
                return true;
            }

            return list.TrueForAll(o =>
                o.TransitionId == 0 &&
                string.IsNullOrWhiteSpace(o.TransitionName) &&
                string.IsNullOrWhiteSpace(o.TargetStateName));
        }

        /// <summary>
        /// 以檔案目前狀態列舉可執行之轉換（後備，補強 GetAvailableTransitionList 結構欄位讀不到的情況）。
        /// </summary>
        private static void TryFillTransitionsFromCurrentState(
            IEdmVault5 vault,
            string samplePath,
            List<PdmTransitionOption> target,
            HashSet<string> seenKeys)
        {
            IEdmFolder5 folder = null;
            IEdmFile5 file = null;
            IEdmState5 state = null;
            IEdmPos5 pos = null;

            try
            {
                file = vault.GetFileFromPath(samplePath, out folder);
                if (file == null)
                {
                    return;
                }

                state = file.CurrentState;
                if (state == null)
                {
                    return;
                }

                var state6 = state as IEdmState6;
                if (state6 == null)
                {
                    return;
                }

                try
                {
                    pos = state6.GetFirstTransitionPosition(false);
                }
                catch
                {
                    pos = null;
                }

                if (pos == null)
                {
                    try
                    {
                        pos = state6.GetFirstTransitionPosition(true);
                    }
                    catch
                    {
                        return;
                    }
                }

                if (pos == null || pos.IsNull)
                {
                    return;
                }

                var guard = 512;
                while (pos != null && !pos.IsNull && guard-- > 0)
                {
                    IEdmTransition5 tr = null;
                    try
                    {
                        tr = state6.GetNextTransition(pos);
                        if (tr == null)
                        {
                            break;
                        }

                        TryReadTransitionIdAndName(tr, out var tid, out var tname, out var dstName);
                        var key = tid + "\u001f" + (tname ?? string.Empty);
                        if (!seenKeys.Add(key))
                        {
                            continue;
                        }

                        target.Add(new PdmTransitionOption
                        {
                            TransitionId = tid,
                            TransitionName = tname ?? string.Empty,
                            TargetStateName = dstName ?? string.Empty
                        });
                    }
                    finally
                    {
                        ComHelper.Release(tr);
                    }
                }
            }
            catch
            {
                // 後備失敗不擲出，由呼叫端依 AvailableTransitions 為空判斷
            }
            finally
            {
                ComHelper.Release(pos);
                ComHelper.Release(state);
                ComHelper.Release(file);
                ComHelper.Release(folder);
            }
        }

        private static void TryReadTransitionIdAndName(
            IEdmTransition5 tr,
            out int transitionId,
            out string transitionName,
            out string targetStateName)
        {
            transitionId = 0;
            transitionName = string.Empty;
            targetStateName = string.Empty;
            if (tr == null)
            {
                return;
            }

            if (tr is IEdmObject5 eo)
            {
                try
                {
                    transitionId = eo.ID;
                }
                catch
                {
                    // ignore
                }

                try
                {
                    transitionName = (eo.Name ?? string.Empty).Trim();
                }
                catch
                {
                    // ignore
                }
            }

            if (string.IsNullOrWhiteSpace(targetStateName))
            {
                TryReadTargetStateNameFromTransitionCom(tr, out targetStateName);
            }
        }

        private static void TryReadTargetStateNameFromTransitionCom(IEdmTransition5 tr, out string targetStateName)
        {
            targetStateName = string.Empty;
            if (tr == null)
            {
                return;
            }

            // 晚繫結：常見為 GetDstState(out IEdmState5)。
            try
            {
                dynamic d = tr;
                object dstObj = null;
                try
                {
                    d.GetDstState(out dstObj);
                }
                catch
                {
                    dstObj = null;
                }

                if (TryCopyStateNameFromObject(dstObj, out targetStateName))
                {
                    return;
                }
            }
            catch
            {
                // ignore
            }

            try
            {
                var t = tr.GetType();
                foreach (var m in t.GetMethods(BindingFlags.Instance | BindingFlags.Public))
                {
                    var ps = m.GetParameters();
                    if (ps.Length == 1 && ps[0].IsOut &&
                        m.Name.IndexOf("Dst", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        object[] args = { null };
                        try
                        {
                            m.Invoke(tr, args);
                        }
                        catch
                        {
                            continue;
                        }

                        if (TryCopyStateNameFromObject(args[0], out targetStateName))
                        {
                            return;
                        }
                    }
                }

                foreach (var m in t.GetMethods(BindingFlags.Instance | BindingFlags.Public))
                {
                    if (m.GetParameters().Length != 0)
                    {
                        continue;
                    }

                    if (m.Name.IndexOf("Dst", StringComparison.OrdinalIgnoreCase) < 0 &&
                        m.Name.IndexOf("Target", StringComparison.OrdinalIgnoreCase) < 0 &&
                        m.Name.IndexOf("ToState", StringComparison.OrdinalIgnoreCase) < 0)
                    {
                        continue;
                    }

                    object r;
                    try
                    {
                        r = m.Invoke(tr, null);
                    }
                    catch
                    {
                        continue;
                    }

                    if (TryCopyStateNameFromObject(r, out targetStateName))
                    {
                        return;
                    }
                }
            }
            catch
            {
                // ignore
            }
        }

        private static bool TryCopyStateNameFromObject(object comObj, out string stateName)
        {
            stateName = string.Empty;
            if (comObj == null)
            {
                return false;
            }

            try
            {
                if (comObj is IEdmState5 st)
                {
                    stateName = (st.Name ?? string.Empty).Trim();
                    ComHelper.Release(st);
                    return !string.IsNullOrWhiteSpace(stateName);
                }

                ComHelper.Release(comObj);
            }
            catch
            {
                ComHelper.Release(comObj);
            }

            return false;
        }

        private sealed class FileStateRow
        {
            public string FullPath { get; set; } = string.Empty;
            public int StateId { get; set; }
            public string StateName { get; set; } = string.Empty;
            public string WorkflowName { get; set; } = string.Empty;
        }

        private static List<string> DeduplicatePaths(IReadOnlyList<string> fullPaths)
        {
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var list = new List<string>();
            if (fullPaths == null)
            {
                return list;
            }

            foreach (var raw in fullPaths)
            {
                var p = raw?.Trim() ?? string.Empty;
                if (string.IsNullOrWhiteSpace(p))
                {
                    continue;
                }

                try
                {
                    p = Path.GetFullPath(p);
                }
                catch
                {
                    continue;
                }

                if (set.Add(p))
                {
                    list.Add(p);
                }
            }

            return list;
        }

        private static bool TryResolveFileRow(IEdmVault5 vault, string fullPath, out FileStateRow row, out string error)
        {
            row = null;
            error = string.Empty;

            if (!fullPath.StartsWith(PdmBomExportService.VaultRootPath, StringComparison.OrdinalIgnoreCase))
            {
                error = $"檔案必須位於 {PdmBomExportService.VaultRootPath} 內。";
                return false;
            }

            if (!File.Exists(fullPath))
            {
                error = "本機找不到檔案。";
                return false;
            }

            IEdmFolder5 folder = null;
            IEdmFile5 file = null;
            IEdmState5 state = null;
            try
            {
                file = vault.GetFileFromPath(fullPath, out folder);
                if (file == null || folder == null)
                {
                    error = "PDM 無法解析路徑。";
                    return false;
                }

                state = file.CurrentState;
                if (state == null)
                {
                    error = "無法讀取目前狀態。";
                    return false;
                }

                row = new FileStateRow
                {
                    FullPath = fullPath,
                    StateId = state.ID,
                    StateName = state.Name?.Trim() ?? string.Empty,
                    WorkflowName = TryGetWorkflowName(state)
                };
                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
            finally
            {
                ComHelper.Release(state);
                ComHelper.Release(file);
                ComHelper.Release(folder);
            }
        }

        private static string TryGetWorkflowName(IEdmState5 state)
        {
            if (state == null)
            {
                return string.Empty;
            }

            // __ComObject 上 GetType().GetMethod 常失敗；晚繫結通常可走 COM vtable。
            try
            {
                dynamic d = state;
                object wfObj = d.GetWorkflow();
                if (wfObj != null)
                {
                    try
                    {
                        dynamic w = wfObj;
                        var n = w.Name;
                        if (n != null)
                        {
                            var s = n.ToString()?.Trim() ?? string.Empty;
                            if (!string.IsNullOrEmpty(s))
                            {
                                return s;
                            }
                        }
                    }
                    finally
                    {
                        ComHelper.Release(wfObj);
                    }
                }
            }
            catch
            {
                // 改試介面／反射
            }

            try
            {
                var t = typeof(IEdmState5);
                var m = t.GetMethod("GetWorkflow", BindingFlags.Instance | BindingFlags.Public);
                if (m != null && m.GetParameters().Length == 0)
                {
                    var wfObj = m.Invoke(state, null);
                    if (wfObj != null)
                    {
                        try
                        {
                            if (wfObj is IEdmObject5 wfO)
                            {
                                return (wfO.Name ?? string.Empty).Trim();
                            }
                        }
                        finally
                        {
                            ComHelper.Release(wfObj);
                        }
                    }
                }
            }
            catch
            {
                // ignore
            }

            return string.Empty;
        }

        /// <summary>
        /// COM RCW 執行期型別為 __ComObject，不可用 GetType().GetMethod 反射呼叫；須轉成 IEdmBatchChangeState* 再呼叫。
        /// </summary>
        private static void InvokeAddFile(object batchObj, int fileId, int folderId)
        {
            switch (batchObj)
            {
                case IEdmBatchChangeState6 b6:
                    b6.AddFile(fileId, folderId);
                    return;
                case IEdmBatchChangeState5 b5:
                    b5.AddFile(fileId, folderId);
                    return;
                case IEdmBatchChangeState4 b4:
                    b4.AddFile(fileId, folderId);
                    return;
                default:
                    throw new InvalidOperationException(
                        "批次轉狀態工具無法轉型為 IEdmBatchChangeState4～6，請確認 SolidWorks PDM Professional 與 EPDM.Interop.epdm 版本一致。");
            }
        }

        private static void InvokeGetAvailableTransitionList(object batchObj, out EdmChangeStateTransitionInfo[] infos)
        {
            infos = null;
            switch (batchObj)
            {
                case IEdmBatchChangeState6 b6:
                    b6.GetAvailableTransitionList(out infos);
                    return;
                case IEdmBatchChangeState5 b5:
                    b5.GetAvailableTransitionList(out infos);
                    return;
                case IEdmBatchChangeState4 b4:
                    b4.GetAvailableTransitionList(out infos);
                    return;
                default:
                    throw new InvalidOperationException(
                        "批次轉狀態工具無法轉型為 IEdmBatchChangeState4～6（GetAvailableTransitionList）。");
            }
        }

        private static void InvokeChangeState2(object batchObj, int parentHwnd, string transitionName)
        {
            switch (batchObj)
            {
                case IEdmBatchChangeState6 b6:
                    b6.ChangeState2(parentHwnd, transitionName, null);
                    return;
                case IEdmBatchChangeState5 b5:
                    b5.ChangeState2(parentHwnd, transitionName, null);
                    return;
                case IEdmBatchChangeState4 b4:
                    b4.ChangeState2(parentHwnd, transitionName, null);
                    return;
                default:
                    throw new InvalidOperationException(
                        "批次轉狀態工具無法轉型為 IEdmBatchChangeState4～6（ChangeState2）。");
            }
        }

        private static void TryReadTransitionFields(
            EdmChangeStateTransitionInfo info,
            out int transitionId,
            out string transitionName,
            out string targetStateName)
        {
            transitionId = 0;
            transitionName = string.Empty;
            targetStateName = string.Empty;

            var t = info.GetType();
            const BindingFlags fieldFlags =
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;

            // 常見 Interop 欄位名（不同年份／版次可能略有差異，先精確比對再退回啟發式）。
            foreach (var f in t.GetFields(fieldFlags))
            {
                var n = f.Name ?? string.Empty;
                object v;
                try
                {
                    v = f.GetValue(info);
                }
                catch
                {
                    continue;
                }

                switch (n)
                {
                    case "mlTransitionID":
                    case "mlSimpleID":
                        if (TryConvertToInt32(v, out var tid))
                        {
                            transitionId = tid;
                        }

                        break;
                    case "mbsTransitionName":
                    case "mbsName":
                        if (v is string s1 && !string.IsNullOrWhiteSpace(s1))
                        {
                            transitionName = s1.Trim();
                        }

                        break;
                    case "mbsTargetStateName":
                    case "mbsToStateName":
                    case "mbsToState":
                        if (v is string s2 && !string.IsNullOrWhiteSpace(s2))
                        {
                            targetStateName = s2.Trim();
                        }

                        break;
                }
            }

            // 後備：僅補齊仍為預設值的欄位（欄位名與 Interop 版本可能略有出入）。
            foreach (var f in t.GetFields(fieldFlags))
            {
                var n = f.Name ?? string.Empty;
                object v;
                try
                {
                    v = f.GetValue(info);
                }
                catch
                {
                    continue;
                }

                if (v == null)
                {
                    continue;
                }

                if (transitionId == 0 && IsIntegralType(f.FieldType) && TryConvertToInt32(v, out var iv))
                {
                    if (n.IndexOf("Transition", StringComparison.OrdinalIgnoreCase) >= 0 &&
                        n.IndexOf("State", StringComparison.OrdinalIgnoreCase) < 0 &&
                        (n.IndexOf("ID", StringComparison.OrdinalIgnoreCase) >= 0 ||
                         n.EndsWith("ID", StringComparison.OrdinalIgnoreCase)))
                    {
                        transitionId = iv;
                    }
                }
                else if (f.FieldType == typeof(string) && v is string str && !string.IsNullOrWhiteSpace(str))
                {
                    str = str.Trim();
                    if (string.IsNullOrWhiteSpace(targetStateName) &&
                        (n.IndexOf("Target", StringComparison.OrdinalIgnoreCase) >= 0 ||
                         n.IndexOf("ToState", StringComparison.OrdinalIgnoreCase) >= 0 ||
                         (n.IndexOf("State", StringComparison.OrdinalIgnoreCase) >= 0 &&
                          n.IndexOf("Transition", StringComparison.OrdinalIgnoreCase) < 0 &&
                          n.IndexOf("From", StringComparison.OrdinalIgnoreCase) < 0)))
                    {
                        targetStateName = str;
                    }
                    else if (string.IsNullOrWhiteSpace(transitionName) &&
                             n.IndexOf("Transition", StringComparison.OrdinalIgnoreCase) >= 0 &&
                             n.IndexOf("Name", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        transitionName = str;
                    }
                    else if (string.IsNullOrWhiteSpace(transitionName) &&
                             n.IndexOf("mbs", StringComparison.OrdinalIgnoreCase) == 0 &&
                             n.IndexOf("Name", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        transitionName = str;
                    }
                }
            }
        }

        private static bool IsIntegralType(Type type)
        {
            if (type == null)
            {
                return false;
            }

            return type == typeof(int) || type == typeof(short) || type == typeof(long) ||
                   type == typeof(uint) || type == typeof(ushort) || type == typeof(byte) ||
                   type == typeof(sbyte);
        }

        private static bool TryConvertToInt32(object v, out int x)
        {
            x = 0;
            if (v == null)
            {
                return false;
            }

            try
            {
                x = Convert.ToInt32(v);
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
