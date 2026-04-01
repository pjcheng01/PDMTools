using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
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
        /// <summary>
        /// PDM EdmChangeStateFlags.EdmChg_ShowDialog（顯示原生變更狀態對話框）。Interop NuGet 未匯出該列舉，改以 API 對應整數。
        /// </summary>
        private const int EdmChg_ShowDialogFlag = 1;

        /// <summary>
        /// PDM EdmChangeStateFlags.EdmChg_ShowErrors（失敗時於對話框顯示錯誤）。Interop 未匯出時之對應值。
        /// </summary>
        private const int EdmChg_ShowErrorsFlag = 2;

        /// <summary>
        /// PDM COM 須在 STA 執行緒呼叫；MTA（如 Task.Run／預設 Thread）會導致對話框無法正確回傳結果。
        /// </summary>
        private static void EnsureStaThreadForPdmCom()
        {
            if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
            {
                throw new InvalidOperationException(
                    "PDM COM 必須在 STA（單執行緒公寓）執行緒上呼叫。請勿在 Task.Run、ThreadPool 或預設為 MTA 的執行緒執行；請於 WPF UI 執行緒呼叫（必要時對控制項使用 Dispatcher）。");
            }
        }

        public BatchWorkflowTransitionAnalyzeResult Analyze(
            IReadOnlyList<string> fullPaths,
            IEdmVault5 vault)
        {
            var result = new BatchWorkflowTransitionAnalyzeResult();
            if (vault == null)
            {
                throw new ArgumentNullException(nameof(vault));
            }

            EnsureStaThreadForPdmCom();

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

        /// <summary>
        /// 除錯／驗證：與主視窗相同邏輯——優先批次 <see cref="InvokeGetAvailableTransitionList"/>（單檔時與 Explorer「變更狀態」一致），再後備列舉。
        /// </summary>
        public IReadOnlyList<PdmTransitionOption> ListNextTransitionableStatesForPath(IEdmVault5 vault, string fullPath)
        {
            EnsureStaThreadForPdmCom();

            var vault7 = vault as IEdmVault7;
            if (vault7 == null)
            {
                return Array.Empty<PdmTransitionOption>();
            }

            var g = new BatchWorkflowTransitionGroup();
            g.FilePaths.Add(fullPath);
            FillAvailableTransitions(vault7, g);
            return g.AvailableTransitions.ToList();
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

            EnsureStaThreadForPdmCom();

            if (parentWindowHandle == 0)
            {
                throw new ArgumentOutOfRangeException(
                    nameof(parentWindowHandle),
                    "父視窗 hwnd 不可為 0。PDM「Change State」對話框需要有效視窗控制代碼，否則可能無法正確執行或無回傳結果。");
            }

            var result = new BatchWorkflowTransitionExecuteResult();
            var vault7 = vault as IEdmVault7;
            if (vault7 == null)
            {
                throw new InvalidOperationException("無法將 Vault 轉型為 IEdmVault7。");
            }

            // 有轉換 ID 時優先 IEdmVault5.ChangeState(…, EdmSelItem[], mlTransitionID, …, EdmChg_ShowDialog, hwnd)。
            // 與 Explorer／工作流程設定一致，可顯示「Change State」原生對話框（含多檔同一 State 之一併轉換）。
            if (transition.TransitionId != 0)
            {
                var vaultCs = TryVaultChangeStateWithNativeDialog(
                    vault,
                    filePathsInGroup,
                    transition.TransitionId,
                    parentWindowHandle,
                    out var vaultCsDetail);

                switch (vaultCs)
                {
                    case VaultChangeStateOutcome.Completed:
                        foreach (var p in filePathsInGroup)
                        {
                            result.Rows.Add(new BatchWorkflowTransitionExecuteRow
                            {
                                FullPath = p,
                                Success = true,
                                Message = string.IsNullOrEmpty(vaultCsDetail)
                                    ? "已透過 ChangeState（原生對話框）完成轉換。"
                                    : vaultCsDetail
                            });
                        }

                        return result;

                    case VaultChangeStateOutcome.UserCancelled:
                        foreach (var p in filePathsInGroup)
                        {
                            result.Rows.Add(new BatchWorkflowTransitionExecuteRow
                            {
                                FullPath = p,
                                Success = false,
                                Message = "使用者已取消，狀態未變更。"
                            });
                        }

                        return result;

                    case VaultChangeStateOutcome.Failed:
                        foreach (var p in filePathsInGroup)
                        {
                            result.Rows.Add(new BatchWorkflowTransitionExecuteRow
                            {
                                FullPath = p,
                                Success = false,
                                Message = string.IsNullOrEmpty(vaultCsDetail) ? "變更狀態失敗。" : vaultCsDetail
                            });
                        }

                        return result;

                    case VaultChangeStateOutcome.UseBatchFallback:
                        break;
                }
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

                // 官方 API 正確序列：AddFile → CreateTree → ShowDlg → ChangeState2(hwnd, password)
                // 參考：SolidWorks PDM API BatchChangeFileStates 官方範例
                var createTreeOk = TryInvokeCreateTree(batchObj, transitionName);
                if (!createTreeOk)
                {
                    foreach (var p in filePathsInGroup)
                    {
                        result.Rows.Add(new BatchWorkflowTransitionExecuteRow
                        {
                            FullPath = p,
                            Success = false,
                            Message = $"CreateTree 失敗，無法初始化批次轉換物件（transition={transitionName}）。"
                        });
                    }

                    return result;
                }

                var showDlgResult = TryInvokeShowDlg(batchObj, parentWindowHandle);
                if (showDlgResult == false)
                {
                    // 使用者在「Change State」對話框按下取消
                    foreach (var p in filePathsInGroup)
                    {
                        result.Rows.Add(new BatchWorkflowTransitionExecuteRow
                        {
                            FullPath = p,
                            Success = false,
                            Message = "使用者已取消，狀態未變更。"
                        });
                    }

                    return result;
                }

                // ChangeState2 第二個參數為密碼（非轉換名稱）；工作流程不需密碼時傳空字串
                InvokeChangeState2(batchObj, parentWindowHandle);

                foreach (var p in filePathsInGroup)
                {
                    var successMsg = $"已透過 CreateTree + ShowDlg + ChangeState2 完成轉換。[transition={transitionName}]";

                    result.Rows.Add(new BatchWorkflowTransitionExecuteRow
                    {
                        FullPath = p,
                        Success = true,
                        Message = successMsg
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
                        Message = FormatPdmExecuteException(vault, ex)
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
                            Message = FormatPdmExecuteException(vault, ex)
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

        /// <summary>
        /// 將 COM 錯誤碼轉成可讀說明（若 Vault 提供 GetErrorMessage）。
        /// </summary>
        private static string FormatPdmExecuteException(IEdmVault5 vault, Exception ex)
        {
            if (ex == null)
            {
                return string.Empty;
            }

            if (ex is not COMException com)
            {
                return ex.Message;
            }

            var hr = unchecked((uint)com.ErrorCode);
            var hex = "HRESULT 0x" + hr.ToString("X8");
            var vaultMsg = TryGetVaultErrorMessage(vault, com.ErrorCode);
            var baseMsg = !string.IsNullOrWhiteSpace(vaultMsg)
                ? ex.Message + " — " + vaultMsg + "（" + hex + "）"
                : ex.Message + "（" + hex + "）";

            return AppendPdmComErrorHints(hr, baseMsg, ex.Message);
        }

        /// <summary>
        /// hwnd／STA 正常時仍可能失敗：補充 0x8004028F（密碼／工作階段）與 XML 附帶訊息的說明。
        /// </summary>
        private static string AppendPdmComErrorHints(uint hr, string formattedMessage, string originalComMessage)
        {
            var msg = formattedMessage ?? string.Empty;
            var om = originalComMessage ?? string.Empty;

            if (hr == 0x8004028F)
            {
                msg +=
                    " 【說明】若您在檔案總管手動轉換亦不需輸入密碼，此 HRESULT 常非「密碼打錯」，而可能是：API 重複提交、ShowDlg 與 ChangeState2 組合與版次不相容、或 COM 工作階段與 Explorer 不同步。建議重新登入 Vault、更新 PDM 用戶端與 EPDM.Interop 版次一致，並請管理員查看伺服器日誌；仍無法排除時可改以僅 IEdmVault5.ChangeState（原生對話框）單一路徑測試。";
            }

            if (om.IndexOf("XML", StringComparison.OrdinalIgnoreCase) >= 0
                || om.IndexOf("最上層", StringComparison.Ordinal) >= 0)
            {
                msg +=
                    " 【說明】XML 相關字樣常為 PDM 在認證或轉換失敗時一併回傳，未必代表檔案內容損毀。";
            }

            return msg;
        }

        private static string TryGetVaultErrorMessage(IEdmVault5 vault, int errorCode)
        {
            if (vault == null)
            {
                return string.Empty;
            }

            try
            {
                dynamic v = vault;
                var s = v.GetErrorMessage(errorCode);
                if (s != null)
                {
                    var t = s.ToString()?.Trim() ?? string.Empty;
                    if (!string.IsNullOrEmpty(t))
                    {
                        return t;
                    }
                }
            }
            catch
            {
                // 改試介面反射
            }

            try
            {
                var asm = typeof(IEdmVault5).Assembly;
                var ns = typeof(IEdmVault5).Namespace;
                foreach (var ifaceName in new[] { "IEdmVault11", "IEdmVault10", "IEdmVault9", "IEdmVault8", "IEdmVault7", "IEdmVault5" })
                {
                    var iface = asm.GetType(ns + "." + ifaceName);
                    if (iface == null)
                    {
                        continue;
                    }

                    var m = iface.GetMethod("GetErrorMessage", new[] { typeof(int) });
                    if (m == null)
                    {
                        continue;
                    }

                    try
                    {
                        var r = m.Invoke(vault, new object[] { errorCode });
                        var t = r?.ToString()?.Trim() ?? string.Empty;
                        if (!string.IsNullOrEmpty(t))
                        {
                            return t;
                        }
                    }
                    catch
                    {
                        // try next
                    }
                }
            }
            catch
            {
                // ignore
            }

            return string.Empty;
        }

        private enum VaultChangeStateOutcome
        {
            /// <summary>使用者於對話框確認且轉換成功。</summary>
            Completed,
            /// <summary>使用者於對話框按取消（COM HRESULT 0x80040005）。</summary>
            UserCancelled,
            /// <summary>權限、Checkout 等可辨識錯誤，或一般失敗。</summary>
            Failed,
            /// <summary>無 ChangeState 或無法解析檔案時改走 EdmUtil_BatchChangeState。</summary>
            UseBatchFallback
        }

        /// <summary>
        /// 組裝 <see cref="EdmSelItem"/>；若 Interop 含 <c>mlVersion</c> 則設為 <see cref="IEdmFile5.CurrentVersion"/>，
        /// 避免 ChangeState 對話框無法套用至正確版次。
        /// </summary>
        private static EdmSelItem CreateEdmSelItemForVaultChangeState(IEdmFile5 file, IEdmFolder5 folder)
        {
            var item = new EdmSelItem
            {
                mlDocID = file.ID,
                mlProjID = folder.ID
            };

            if (file == null)
            {
                return item;
            }

            var t = typeof(EdmSelItem);
            var fv = t.GetField("mlVersion", BindingFlags.Public | BindingFlags.Instance);
            if (fv != null && fv.FieldType == typeof(int))
            {
                try
                {
                    object boxed = item;
                    fv.SetValue(boxed, file.CurrentVersion);
                    return (EdmSelItem)boxed;
                }
                catch
                {
                    // 忽略：舊版 Interop 欄位型別不同
                }
            }

            return item;
        }

        /// <summary>
        /// 對應文件：IEdmVault5.ChangeState(EdmObjectType.EdmObject_File, EdmSelItem[], transitionId, comment, flags, hwnd)。
        /// 使用 EdmChg_ShowDialog|EdmChg_ShowErrors；以 COMException 區分取消、失敗與改走批次後備（如 0x8004028F）。
        /// </summary>
        private static VaultChangeStateOutcome TryVaultChangeStateWithNativeDialog(
            IEdmVault5 vault,
            IReadOnlyList<string> filePathsInGroup,
            int transitionId,
            int parentWindowHandle,
            out string detailMessage)
        {
            detailMessage = null;

            if (transitionId == 0 || filePathsInGroup == null || filePathsInGroup.Count == 0)
            {
                return VaultChangeStateOutcome.UseBatchFallback;
            }

            var items = new List<EdmSelItem>();
            foreach (var path in filePathsInGroup)
            {
                IEdmFolder5 folder = null;
                IEdmFile5 file = null;
                try
                {
                    file = vault.GetFileFromPath(path, out folder);
                    if (file == null || folder == null)
                    {
                        return VaultChangeStateOutcome.UseBatchFallback;
                    }

                    items.Add(CreateEdmSelItemForVaultChangeState(file, folder));
                }
                finally
                {
                    ComHelper.Release(file);
                    ComHelper.Release(folder);
                }
            }

            // EdmChg_ShowDialog | EdmChg_ShowErrors（與技術文件一致）；Interop 未匯出列舉時以常數 OR。
            var flags = EdmChg_ShowDialogFlag | EdmChg_ShowErrorsFlag;

            // NuGet Interop 的 IEdmVault5 常未宣告 ChangeState，執行期仍可能存在；僅能晚繫結呼叫。
            try
            {
                dynamic v = vault;
                v.ChangeState(
                    EdmObjectType.EdmObject_File,
                    items.ToArray(),
                    transitionId,
                    string.Empty,
                    flags,
                    parentWindowHandle);
                return VaultChangeStateOutcome.Completed;
            }
            catch (COMException ex)
            {
                var hr = unchecked((uint)ex.ErrorCode);
                switch (hr)
                {
                    case 0x80040005:
                        return VaultChangeStateOutcome.UserCancelled;
                    case 0x80040008:
                        detailMessage = "權限不足，無法執行此轉換。";
                        return VaultChangeStateOutcome.Failed;
                    case 0x80040009:
                        detailMessage = "檔案已 Check Out，請先 Check In 後再試。";
                        return VaultChangeStateOutcome.Failed;
                    case 0x8004028F:
                        // 密碼／認證等：改試批次 API（ShowDlg + ChangeState2）
                        return VaultChangeStateOutcome.UseBatchFallback;
                    default:
                        detailMessage = $"未預期的錯誤：0x{hr:X8} — {ex.Message}";
                        return VaultChangeStateOutcome.Failed;
                }
            }
            catch (Exception ex)
            {
                // 晚繫結無 ChangeState 時會擲出 RuntimeBinderException；部分專案無法直接參考該型別（CS0246），改以 FullName 判斷（.NET Framework 與新版 CLR 命名空間可能不同）。
                var exFull = ex.GetType().FullName;
                if (string.Equals(exFull, "Microsoft.CSharp.RuntimeBinder.RuntimeBinderException", StringComparison.Ordinal)
                    || string.Equals(exFull, "Microsoft.CSharp.RuntimeBinderException", StringComparison.Ordinal))
                {
                    return VaultChangeStateOutcome.UseBatchFallback;
                }

                if (ex is MissingMethodException)
                {
                    return VaultChangeStateOutcome.UseBatchFallback;
                }

                detailMessage = $"變更狀態失敗：{ex.Message}";
                return VaultChangeStateOutcome.Failed;
            }
        }

        private static string PickTransitionNameForChangeState2(object batchObj, PdmTransitionOption transition)
        {
            if (!string.IsNullOrWhiteSpace(transition.MbsTransitionNameRaw))
            {
                return transition.MbsTransitionNameRaw.Trim();
            }

            if (!string.IsNullOrWhiteSpace(transition.TransitionName))
            {
                return transition.TransitionName.Trim();
            }

            // 後備：從目前批次再取清單對應 ID（優先 mbsTransitionName 原文）
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
                TryReadTransitionFields(info, out var id, out var name, out _, out var rawMbs);
                if (id == transition.TransitionId)
                {
                    if (!string.IsNullOrWhiteSpace(rawMbs))
                    {
                        return rawMbs.Trim();
                    }

                    if (!string.IsNullOrWhiteSpace(name))
                    {
                        return name.Trim();
                    }
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
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // ── 主路徑：IEdmState6.GetFirstTransitionPosition(true) ────────────────────────────────
            // GetFirstTransitionPosition(true)  = exit（離開當前狀態）transitions，與 Explorer「變更狀態」一致。
            // GetFirstTransitionPosition(false) = entry（進入當前狀態）transitions，不應顯示。
            // 批次 API GetAvailableTransitionList 在部分 PDM 版本會同時回傳 exit 與 entry transitions，
            // 導致顯示錯誤的「取消…」類反向轉換；故改以 IEdmState6 列舉作為主路徑。
            if (group.FilePaths.Count > 0)
            {
                TryFillTransitionsFromCurrentState(vault7, group.FilePaths[0], group.AvailableTransitions, seen);
            }

            if (!TransitionsListLooksUnusable(group.AvailableTransitions))
            {
                return;
            }

            group.AvailableTransitions.Clear();
            seen.Clear();

            // ── 後備 1：批次 GetAvailableTransitionList ───────────────────────────────────────────
            // 僅當 IEdmState6 路徑完全無結果時才使用。
            object batchObj = null;
            try
            {
                batchObj = vault7.CreateUtility(EdmUtility.EdmUtil_BatchChangeState);
                if (batchObj == null)
                {
                    group.AnalyzeError = "無法建立批次轉狀態工具。";
                }
                else
                {
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

                    if (infos != null && infos.Length > 0)
                    {
                        foreach (var info in infos)
                        {
                            TryReadTransitionFields(info, out var id, out var name, out var target, out var rawMbs);
                            var key = id + "\u001f" + (name ?? string.Empty);
                            if (!seen.Add(key))
                            {
                                continue;
                            }

                            group.AvailableTransitions.Add(new PdmTransitionOption
                            {
                                TransitionId = id,
                                TransitionName = name ?? string.Empty,
                                MbsTransitionNameRaw = rawMbs ?? string.Empty,
                                TargetStateName = target ?? string.Empty
                            });
                        }

                        if (group.FilePaths.Count > 0)
                        {
                            TryEnrichZeroTransitionIdsFromStateEnumeration(
                                vault7,
                                group.FilePaths[0],
                                group.AvailableTransitions);
                        }
                    }
                }
            }
            finally
            {
                ComHelper.Release(batchObj);
            }

            if (!TransitionsListLooksUnusable(group.AvailableTransitions))
            {
                return;
            }

            group.AvailableTransitions.Clear();
            seen.Clear();

            // ── 後備 2：列舉器（GetEnumeratorTransition）────────────────────────────────────────
            if (group.FilePaths.Count > 0)
            {
                TryFillTransitionsViaEnumeratorTransition(vault7, group.FilePaths[0], group.AvailableTransitions, seen);
            }

            if (group.AvailableTransitions.Count == 0 && string.IsNullOrWhiteSpace(group.AnalyzeError))
            {
                group.AnalyzeError = "此群組找不到任何可用轉換（IEdmState6、批次 GetAvailableTransitionList、列舉器皆無結果）；請確認權限與工作流程設定。";
            }
        }

        /// <summary>
        /// PDF：GetFileState → GetEnumeratorTransition → MoveNext(out 轉換結構)。
        /// NuGet 版 Interop 常未產生 IEdmEnumeratorTransition／EdmTransition 型別，改以晚繫結＋反射讀欄位。
        /// </summary>
        private static void TryFillTransitionsViaEnumeratorTransition(
            IEdmVault5 vault,
            string samplePath,
            List<PdmTransitionOption> target,
            HashSet<string> seenKeys)
        {
            IEdmFolder5 folder = null;
            IEdmFile5 file = null;
            IEdmState5 state = null;
            object enumerator = null;

            try
            {
                file = vault.GetFileFromPath(samplePath, out folder);
                if (file == null || folder == null)
                {
                    return;
                }

                state = TryGetFileStateFromFile(file, folder);
                if (state == null)
                {
                    return;
                }

                try
                {
                    dynamic st = state;
                    enumerator = st.GetEnumeratorTransition();
                }
                catch
                {
                    return;
                }

                if (enumerator == null)
                {
                    return;
                }

                var guard = 512;
                while (guard-- > 0)
                {
                    if (!TryEnumeratorTransitionMoveNext(enumerator, out var transObj))
                    {
                        break;
                    }

                    ReadMlTransitionFieldsFromBoxedStruct(transObj, out var id, out var name, out var tgt, out var rawMbs);
                    var key = id + "\u001f" + name;
                    if (!seenKeys.Add(key))
                    {
                        continue;
                    }

                    target.Add(new PdmTransitionOption
                    {
                        TransitionId = id,
                        TransitionName = name,
                        MbsTransitionNameRaw = rawMbs ?? string.Empty,
                        TargetStateName = tgt ?? string.Empty
                    });
                }
            }
            catch
            {
                // 由後續批次／IEdmState6 後備
            }
            finally
            {
                ComHelper.Release(enumerator);
                ComHelper.Release(state);
                ComHelper.Release(file);
                ComHelper.Release(folder);
            }
        }

        private static bool TryEnumeratorTransitionMoveNext(object enumerator, out object transitionOut)
        {
            transitionOut = null;
            if (enumerator == null)
            {
                return false;
            }

            var t = enumerator.GetType();
            foreach (var cand in t.GetMethods(BindingFlags.Instance | BindingFlags.Public))
            {
                if (!string.Equals(cand.Name, "MoveNext", StringComparison.Ordinal))
                {
                    continue;
                }

                var ps = cand.GetParameters();
                if (ps.Length != 1 || !ps[0].ParameterType.IsByRef)
                {
                    continue;
                }

                var elem = ps[0].ParameterType.GetElementType();
                if (elem == null)
                {
                    continue;
                }

                var args = new object[1];
                args[0] = elem.IsValueType ? Activator.CreateInstance(elem) : null;

                bool moved;
                try
                {
                    var ret = cand.Invoke(enumerator, args);
                    moved = ret is bool mb && mb;
                }
                catch
                {
                    return false;
                }

                transitionOut = args[0];
                return moved;
            }

            return false;
        }

        private static void ReadMlTransitionFieldsFromBoxedStruct(
            object transObj,
            out int transitionId,
            out string transitionName,
            out string targetStateName,
            out string rawMbsTransitionName)
        {
            transitionId = 0;
            transitionName = string.Empty;
            targetStateName = string.Empty;
            rawMbsTransitionName = string.Empty;
            if (transObj == null)
            {
                return;
            }

            string moName = null;
            string mbsTn = null;
            var tt = transObj.GetType();
            foreach (var f in tt.GetFields(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic))
            {
                var n = f.Name ?? string.Empty;
                object v;
                try
                {
                    v = f.GetValue(transObj);
                }
                catch
                {
                    continue;
                }

                if (string.Equals(n, "mlTransitionID", StringComparison.Ordinal) && v != null)
                {
                    try
                    {
                        transitionId = Convert.ToInt32(v);
                    }
                    catch
                    {
                        // ignore
                    }
                }
                else if (string.Equals(n, "moName", StringComparison.Ordinal) && v != null)
                {
                    moName = NormalizeTransitionNameString(v);
                }
                else if (string.Equals(n, "mbsTransitionName", StringComparison.Ordinal) && v != null)
                {
                    mbsTn = NormalizeTransitionNameString(v);
                }
                else if ((string.Equals(n, "mbsTargetStateName", StringComparison.Ordinal) ||
                          string.Equals(n, "mbsToStateName", StringComparison.Ordinal)) && v != null)
                {
                    targetStateName = NormalizeTransitionNameString(v);
                }
            }

            transitionName = PickPreferredTransitionDisplayName(moName, mbsTn, null);
            rawMbsTransitionName = (mbsTn ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(rawMbsTransitionName))
            {
                rawMbsTransitionName = transitionName;
            }
        }

        /// <summary>
        /// Explorer 只顯示目前使用者有權執行之轉換；IEdmTransition6.CheckPermission 與之一致。
        /// </summary>
        private static bool TransitionAllowedForCurrentUser(IEdmTransition5 tr)
        {
            if (tr == null)
            {
                return false;
            }

            if (tr is IEdmTransition6 t6)
            {
                try
                {
                    return t6.CheckPermission();
                }
                catch
                {
                    return true;
                }
            }

            try
            {
                dynamic d = tr;
                return d.CheckPermission();
            }
            catch
            {
                return true;
            }
        }

        private static string NormalizeTransitionNameString(object v)
        {
            if (v == null)
            {
                return string.Empty;
            }

            if (v is string s)
            {
                return s.Trim();
            }

            return Convert.ToString(v, System.Globalization.CultureInfo.InvariantCulture)?.Trim() ?? string.Empty;
        }

        /// <summary>PDF Step 4：優先 GetFileState(folderId)；若無則不強制（列舉器會再試）。分析分組仍可用 CurrentState 後備。</summary>
        private static IEdmState5 TryGetFileStateFromFile(IEdmFile5 file, IEdmFolder5 folder)
        {
            if (file == null || folder == null)
            {
                return null;
            }

            try
            {
                dynamic f = file;
                object s = f.GetFileState(folder.ID);
                return s as IEdmState5;
            }
            catch
            {
                return null;
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
                if (file == null || folder == null)
                {
                    return;
                }

                state = TryGetFileStateFromFile(file, folder);
                if (state == null)
                {
                    state = file.CurrentState;
                }

                if (state == null)
                {
                    return;
                }

                var state6 = state as IEdmState6;
                if (state6 == null)
                {
                    return;
                }

                // GetFirstTransitionPosition(true)  = exit transitions（離開此狀態，即使用者想執行的轉換）。
                // GetFirstTransitionPosition(false) = entry transitions（進入此狀態的反向轉換，不應列出）。
                // 先試 true（exit），若例外再試 false 作為後備（理論上不應走到）。
                try
                {
                    pos = state6.GetFirstTransitionPosition(true);
                }
                catch
                {
                    pos = null;
                }

                if (pos == null)
                {
                    try
                    {
                        pos = state6.GetFirstTransitionPosition(false);
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

                        if (!TransitionAllowedForCurrentUser(tr))
                        {
                            continue;
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
                            MbsTransitionNameRaw = (tname ?? string.Empty).Trim(),
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

            // IEdmTransition5 提供 ToState 屬性（getter = get_ToState），直接取得目標狀態。
            // 注意：ToState 回傳的 IEdmState5 為 COM 物件，使用後須 Release。
            IEdmState5 dstState = null;
            try
            {
                dstState = tr.ToState;
                if (dstState != null)
                {
                    targetStateName = (dstState.Name ?? string.Empty).Trim();
                }
            }
            catch
            {
                // ignore
            }
            finally
            {
                ComHelper.Release(dstState);
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

                state = TryGetFileStateFromFile(file, folder);
                if (state == null)
                {
                    state = file.CurrentState;
                }

                if (state == null)
                {
                    error = "無法讀取目前狀態（GetFileState／CurrentState 皆無）。";
                    return false;
                }

                row = new FileStateRow
                {
                    FullPath = fullPath,
                    StateId = state.ID,
                    StateName = state.Name?.Trim() ?? string.Empty,
                    WorkflowName = TryGetWorkflowName(vault, file, state)
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

        /// <summary>
        /// 先由狀態／檔案 <c>GetWorkflow()</c>（PDM 2023 仍為主路徑）；若名稱仍空，後備以 workflow ID + Vault <c>GetObject(EdmObject_Workflow, …)</c> 再讀名稱。
        /// </summary>
        private static string TryGetWorkflowName(IEdmVault5 vault, IEdmFile5 file, IEdmState5 state)
        {
            object wfObj = null;
            try
            {
                wfObj = TryInvokeGetWorkflow(state);
                var name = TryReadWorkflowComObjectName(wfObj);
                if (!string.IsNullOrWhiteSpace(name))
                {
                    return name;
                }
            }
            finally
            {
                ComHelper.Release(wfObj);
                wfObj = null;
            }

            try
            {
                wfObj = TryInvokeGetWorkflow(file);
                var name = TryReadWorkflowComObjectName(wfObj);
                if (!string.IsNullOrWhiteSpace(name))
                {
                    return name;
                }
            }
            finally
            {
                ComHelper.Release(wfObj);
                wfObj = null;
            }

            return TryGetWorkflowNameViaVaultGetObject(vault, file, state);
        }

        /// <summary>
        /// PDM 2023 後備：由 state／file 讀取 workflow ID（常見欄位 mlWorkflowID），再以 Vault.GetObject(EdmObject_Workflow, id) 取得工作流程物件。
        /// </summary>
        private static string TryGetWorkflowNameViaVaultGetObject(IEdmVault5 vault, IEdmFile5 file, IEdmState5 state)
        {
            if (vault == null)
            {
                return string.Empty;
            }

            var wfId = TryGetWorkflowIdFromStateOrFile(state, file);
            if (wfId == 0)
            {
                return string.Empty;
            }

            object wfObj = null;
            try
            {
                wfObj = TryVaultGetWorkflowObject(vault, wfId);
                return TryReadWorkflowComObjectName(wfObj);
            }
            finally
            {
                ComHelper.Release(wfObj);
            }
        }

        private static int TryGetWorkflowIdFromStateOrFile(IEdmState5 state, IEdmFile5 file)
        {
            var id = TryGetWorkflowIdFromComObject(state);
            if (id != 0)
            {
                return id;
            }

            return TryGetWorkflowIdFromComObject(file);
        }

        /// <summary>
        /// 嘗試從 Interop 公開屬性讀取 workflow ID（含 mlWorkflowID／WorkflowID 等；各版 IEdmState6+／IEdmFile6+ 命名可能不同）。
        /// </summary>
        private static int TryGetWorkflowIdFromComObject(object comTarget)
        {
            if (comTarget == null)
            {
                return 0;
            }

            foreach (var methodName in new[] { "GetWorkflowID", "GetWorkflowId", "GetMLWorkflowID" })
            {
                try
                {
                    dynamic d = comTarget;
                    object r = null;
                    switch (methodName)
                    {
                        case "GetWorkflowID":
                            r = d.GetWorkflowID();
                            break;
                        case "GetWorkflowId":
                            r = d.GetWorkflowId();
                            break;
                        case "GetMLWorkflowID":
                            r = d.GetMLWorkflowID();
                            break;
                    }

                    if (TryConvertToPositiveInt32(r, out var mid))
                    {
                        return mid;
                    }
                }
                catch
                {
                    // try next method name
                }
            }

            var asm = typeof(IEdmState5).Assembly;
            var ns = typeof(IEdmState5).Namespace;
            var ifaceNames = comTarget is IEdmFile5
                ? new[] { "IEdmFile9", "IEdmFile8", "IEdmFile7", "IEdmFile6", "IEdmFile5" }
                : new[] { "IEdmState9", "IEdmState8", "IEdmState7", "IEdmState6", "IEdmState5" };

            foreach (var ifaceName in ifaceNames)
            {
                var iface = asm.GetType(ns + "." + ifaceName);
                if (iface == null)
                {
                    continue;
                }

                foreach (var mn in new[] { "GetWorkflowID", "GetWorkflowId", "GetMLWorkflowID" })
                {
                    var m = iface.GetMethod(mn, Type.EmptyTypes);
                    if (m == null)
                    {
                        continue;
                    }

                    try
                    {
                        var r = m.Invoke(comTarget, null);
                        if (TryConvertToPositiveInt32(r, out var mid))
                        {
                            return mid;
                        }
                    }
                    catch
                    {
                        // try next
                    }
                }
            }

            foreach (var propName in new[] { "mlWorkflowID", "WorkflowID", "mlWorkflowId", "WorkflowId" })
            {
                try
                {
                    var t = comTarget.GetType();
                    var p = t.GetProperty(propName, BindingFlags.Instance | BindingFlags.Public);
                    if (p == null)
                    {
                        continue;
                    }

                    var v = p.GetValue(comTarget);
                    if (TryConvertToPositiveInt32(v, out var id))
                    {
                        return id;
                    }
                }
                catch
                {
                    // try next
                }
            }

            foreach (var ifaceName in ifaceNames)
            {
                var iface = asm.GetType(ns + "." + ifaceName);
                if (iface == null)
                {
                    continue;
                }

                foreach (var p in iface.GetProperties(BindingFlags.Instance | BindingFlags.Public))
                {
                    if (p.GetIndexParameters().Length != 0)
                    {
                        continue;
                    }

                    var n = p.Name;
                    if (n.IndexOf("Workflow", StringComparison.OrdinalIgnoreCase) < 0)
                    {
                        continue;
                    }

                    if (n.IndexOf("ID", StringComparison.OrdinalIgnoreCase) < 0 &&
                        !string.Equals(n, "WorkflowID", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    try
                    {
                        var v = p.GetValue(comTarget);
                        if (TryConvertToPositiveInt32(v, out var id))
                        {
                            return id;
                        }
                    }
                    catch
                    {
                        // try next property
                    }
                }
            }

            return 0;
        }

        private static bool TryConvertToPositiveInt32(object v, out int id)
        {
            id = 0;
            if (v == null)
            {
                return false;
            }

            try
            {
                switch (v)
                {
                    case int i:
                        id = i;
                        break;
                    case short s:
                        id = s;
                        break;
                    case long l:
                        id = checked((int)l);
                        break;
                    case uint ui:
                        id = checked((int)ui);
                        break;
                    default:
                        if (!int.TryParse(Convert.ToString(v, System.Globalization.CultureInfo.InvariantCulture), out id))
                        {
                            return false;
                        }

                        break;
                }
            }
            catch
            {
                return false;
            }

            return id > 0;
        }

        private static object TryResolveEdmObjectWorkflowEnumBoxed()
        {
            var t = typeof(EdmObjectType);
            foreach (var name in Enum.GetNames(t))
            {
                if (string.Equals(name, "EdmObject_Workflow", StringComparison.OrdinalIgnoreCase))
                {
                    return Enum.Parse(t, name);
                }
            }

            foreach (var name in Enum.GetNames(t))
            {
                if (name.IndexOf("Workflow", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return Enum.Parse(t, name);
                }
            }

            return null;
        }

        /// <summary>
        /// 呼叫 Vault.GetObject(workflowType, id)；支援回傳值或 out 參數等簽章差異（PDM 2023 Interop）。
        /// </summary>
        private static object TryVaultGetWorkflowObject(IEdmVault5 vault, int workflowId)
        {
            if (vault == null || workflowId <= 0)
            {
                return null;
            }

            var enumVal = TryResolveEdmObjectWorkflowEnumBoxed();
            if (enumVal == null)
            {
                return null;
            }

            try
            {
                dynamic v = vault;
                var o = v.GetObject(enumVal, workflowId);
                if (o != null)
                {
                    return o;
                }
            }
            catch
            {
                // try overloads
            }

            var vaultIfaceTypes = new List<Type> { typeof(IEdmVault5), typeof(IEdmVault7) };
            var epdmAsm = typeof(IEdmVault5).Assembly;
            var epdmNs = typeof(IEdmVault5).Namespace;
            foreach (var extra in new[] { "IEdmVault8", "IEdmVault9" })
            {
                var t = epdmAsm.GetType(epdmNs + "." + extra);
                if (t != null)
                {
                    vaultIfaceTypes.Add(t);
                }
            }

            foreach (var vaultIface in vaultIfaceTypes)
            {
                foreach (var m in vaultIface.GetMethods(BindingFlags.Instance | BindingFlags.Public))
                {
                    if (!string.Equals(m.Name, "GetObject", StringComparison.Ordinal))
                    {
                        continue;
                    }

                    var ps = m.GetParameters();
                    if (ps.Length == 2 &&
                        ps[0].ParameterType.IsEnum &&
                        ps[1].ParameterType == typeof(int))
                    {
                        try
                        {
                            return m.Invoke(vault, new[] { enumVal, workflowId });
                        }
                        catch
                        {
                            // try next
                        }
                    }

                    if (ps.Length == 3 && ps[2].ParameterType.IsByRef)
                    {
                        try
                        {
                            var args = new object[] { enumVal, workflowId, null };
                            m.Invoke(vault, args);
                            return args[2];
                        }
                        catch
                        {
                            // try next
                        }
                    }

                    // 部分 Interop 以整數傳入 EdmObjectType 鑄成的值
                    if (ps.Length == 2 &&
                        ps[0].ParameterType == typeof(int) &&
                        ps[1].ParameterType == typeof(int))
                    {
                        try
                        {
                            var typeInt = Convert.ToInt32(enumVal);
                            return m.Invoke(vault, new object[] { typeInt, workflowId });
                        }
                        catch
                        {
                            // try next
                        }
                    }

                    if (ps.Length == 3 &&
                        ps[0].ParameterType == typeof(int) &&
                        ps[1].ParameterType == typeof(int) &&
                        ps[2].ParameterType.IsByRef)
                    {
                        try
                        {
                            var typeInt = Convert.ToInt32(enumVal);
                            var args = new object[] { typeInt, workflowId, null };
                            m.Invoke(vault, args);
                            return args[2];
                        }
                        catch
                        {
                            // try next
                        }
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// 取得 IEdmWorkflow*。COM RCW 執行期為 __ComObject，不可用 GetType().GetMethods() 列舉介面方法；
        /// 必須在 Interop 介面型別上取得 GetWorkflow() 再 Invoke（與 PdmBomExportService 處理 IEdmFile7 相同）。
        /// </summary>
        private static object TryInvokeGetWorkflow(object comTarget)
        {
            if (comTarget == null)
            {
                return null;
            }

            try
            {
                dynamic d = comTarget;
                var o = d.GetWorkflow();
                if (o != null)
                {
                    return o;
                }
            }
            catch
            {
                // 改試介面型別反射
            }

            var asm = typeof(IEdmState5).Assembly;
            var ns = typeof(IEdmState5).Namespace;
            var ifaceNames = comTarget is IEdmFile5
                ? new[] { "IEdmFile9", "IEdmFile8", "IEdmFile7", "IEdmFile6", "IEdmFile5" }
                : new[] { "IEdmState9", "IEdmState8", "IEdmState7", "IEdmState6", "IEdmState5" };

            foreach (var ifaceName in ifaceNames)
            {
                var iface = asm.GetType(ns + "." + ifaceName);
                if (iface == null)
                {
                    continue;
                }

                var m = iface.GetMethod("GetWorkflow", Type.EmptyTypes);
                if (m == null)
                {
                    continue;
                }

                try
                {
                    var o = m.Invoke(comTarget, null);
                    if (o != null)
                    {
                        return o;
                    }
                }
                catch
                {
                    // RCW 可能未實作此介面層級，試下一版
                }
            }

            return null;
        }

        private static string TryReadWorkflowComObjectName(object wfObj)
        {
            if (wfObj == null)
            {
                return string.Empty;
            }

            try
            {
                if (wfObj is IEdmObject5 wfO)
                {
                    var s = (wfO.Name ?? string.Empty).Trim();
                    if (!string.IsNullOrEmpty(s))
                    {
                        return s;
                    }
                }
            }
            catch
            {
                // ignore
            }

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
            catch
            {
                // ignore
            }

            try
            {
                var t = wfObj.GetType();
                foreach (var propName in new[] { "Name", "mbsName" })
                {
                    var p = t.GetProperty(propName, BindingFlags.Instance | BindingFlags.Public);
                    if (p == null)
                    {
                        continue;
                    }

                    var v = p.GetValue(wfObj);
                    if (v == null)
                    {
                        continue;
                    }

                    var s = v.ToString()?.Trim() ?? string.Empty;
                    if (!string.IsNullOrEmpty(s))
                    {
                        return s;
                    }
                }
            }
            catch
            {
                // ignore
            }

            try
            {
                var asm = typeof(IEdmObject5).Assembly;
                var ns = typeof(IEdmObject5).Namespace;
                foreach (var ifaceName in new[] { "IEdmWorkflow9", "IEdmWorkflow8", "IEdmWorkflow7", "IEdmWorkflow6", "IEdmWorkflow5" })
                {
                    var iface = asm.GetType(ns + "." + ifaceName);
                    if (iface == null)
                    {
                        continue;
                    }

                    var p = iface.GetProperty("Name", BindingFlags.Instance | BindingFlags.Public);
                    if (p == null)
                    {
                        continue;
                    }

                    try
                    {
                        var v = p.GetValue(wfObj);
                        if (v == null)
                        {
                            continue;
                        }

                        var s = v.ToString()?.Trim() ?? string.Empty;
                        if (!string.IsNullOrEmpty(s))
                        {
                            return s;
                        }
                    }
                    catch
                    {
                        // try next interface
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

        /// <summary>
        /// ChangeState2 的第二參數為「使用者密碼」（官方 API 範例：bsPasswd），並非轉換名稱。
        /// 轉換已由 <see cref="TryInvokeCreateTree"/> 設定；此處傳入空字串表示不需密碼。
        /// </summary>
        private static void InvokeChangeState2(object batchObj, int parentHwnd)
        {
            const string password = "";   // 工作流程不需密碼時傳空字串
            EdmCallback callback = null;
            switch (batchObj)
            {
                case IEdmBatchChangeState6 b6:
                    b6.ChangeState2(parentHwnd, password, callback);
                    return;
                case IEdmBatchChangeState5 b5:
                    b5.ChangeState2(parentHwnd, password, callback);
                    return;
                case IEdmBatchChangeState4 b4:
                    b4.ChangeState2(parentHwnd, password, callback);
                    return;
                default:
                    throw new InvalidOperationException(
                        "批次轉狀態工具無法轉型為 IEdmBatchChangeState4～6（ChangeState2）。");
            }
        }

        /// <summary>
        /// CreateTree 定義於基底 <c>IEdmBatchChangeState</c>，須在 AddFile 後、ShowDlg 前呼叫，以設定本次批次轉換目標。
        /// 回傳 true 表示成功呼叫；false 表示 dynamic 與反射均未能執行（可供呼叫端記錄診斷）。
        /// </summary>
        private static bool TryInvokeCreateTree(object batchObj, string transitionName)
        {
            if (batchObj == null || string.IsNullOrWhiteSpace(transitionName))
            {
                return false;
            }

            try
            {
                dynamic b = batchObj;
                b.CreateTree(transitionName);
                return true;
            }
            catch
            {
                // 改試介面反射
            }

            // CreateTree 定義於基底 IEdmBatchChangeState；反射需涵蓋至低版次介面，不可只列 v4+。
            var asm = typeof(IEdmBatchChangeState4).Assembly;
            var ns = typeof(IEdmBatchChangeState4).Namespace;
            var ifaceTypes = new List<Type>
            {
                typeof(IEdmBatchChangeState6),
                typeof(IEdmBatchChangeState5),
                typeof(IEdmBatchChangeState4)
            };

            foreach (var name in new[] { "IEdmBatchChangeState3", "IEdmBatchChangeState2", "IEdmBatchChangeState" })
            {
                var t = asm.GetType(ns + "." + name);
                if (t != null)
                {
                    ifaceTypes.Add(t);
                }
            }

            foreach (var iface in ifaceTypes)
            {
                var m = iface.GetMethod("CreateTree", new[] { typeof(string) });
                if (m == null)
                {
                    continue;
                }

                try
                {
                    m.Invoke(batchObj, new object[] { transitionName });
                    return true;
                }
                catch
                {
                    // try next
                }
            }

            return false;
        }

        /// <summary>
        /// COM／VARIANT_BOOL（常為 short：0 假、-1 真）轉成 bool。
        /// </summary>
        private static bool CoerceComVariantBool(object r)
        {
            if (r == null || r is DBNull)
            {
                return false;
            }

            if (r is bool b)
            {
                return b;
            }

            try
            {
                return Convert.ToBoolean(r);
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 顯示批次變更狀態對話方塊。若介面無此方法則回傳 null。
        /// </summary>
        private static bool? TryInvokeShowDlg(object batchObj, int parentHwnd)
        {
            if (batchObj == null)
            {
                return null;
            }

            try
            {
                dynamic b = batchObj;
                var ret = b.ShowDlg(parentHwnd);
                return CoerceComVariantBool(ret);
            }
            catch
            {
                // 改試介面反射
            }

            foreach (var iface in new[] { typeof(IEdmBatchChangeState6), typeof(IEdmBatchChangeState5), typeof(IEdmBatchChangeState4) })
            {
                var m = iface.GetMethod("ShowDlg", new[] { typeof(int) });
                if (m == null)
                {
                    continue;
                }

                try
                {
                    var ret = m.Invoke(batchObj, new object[] { parentHwnd });
                    return CoerceComVariantBool(ret);
                }
                catch
                {
                    // try next
                }
            }

            return null;
        }

        /// <summary>
        /// 批次結構 <see cref="EdmChangeStateTransitionInfo"/> 可能同時含 <c>moName</c> 與 <c>mbsTransitionName</c>：
        /// SOLIDWORKS 官方範例以 <c>moName</c> 作為顯示字串；Explorer「變更狀態」亦較接近 <c>moName</c>。
        /// <c>mbsTransitionName</c> 有時為另一組命名（例如與圖示標籤「01_…」／<c>moName</c> 不同）。
        /// </summary>
        private static string PickPreferredTransitionDisplayName(
            string moName,
            string mbsTransitionName,
            string mbsNameFallback)
        {
            if (!string.IsNullOrWhiteSpace(moName))
            {
                return moName.Trim();
            }

            if (!string.IsNullOrWhiteSpace(mbsTransitionName))
            {
                return mbsTransitionName.Trim();
            }

            return (mbsNameFallback ?? string.Empty).Trim();
        }

        private static void TryReadTransitionFields(
            EdmChangeStateTransitionInfo info,
            out int transitionId,
            out string transitionName,
            out string targetStateName,
            out string rawMbsTransitionName)
        {
            transitionId = 0;
            transitionName = string.Empty;
            targetStateName = string.Empty;
            rawMbsTransitionName = string.Empty;

            var t = info.GetType();
            const BindingFlags fieldFlags =
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;

            string moName = null;
            string mbsTransitionNameStr = null;
            string mbsNameStr = null;
            var mlTransitionIdRaw = 0;
            var mlSimpleIdRaw = 0;

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
                        if (TryConvertToInt32(v, out var tid))
                        {
                            mlTransitionIdRaw = tid;
                        }

                        break;
                    case "mlSimpleID":
                        if (TryConvertToInt32(v, out var sid))
                        {
                            mlSimpleIdRaw = sid;
                        }

                        break;
                    case "moName":
                        if (v is string s0 && !string.IsNullOrWhiteSpace(s0))
                        {
                            moName = s0.Trim();
                        }

                        break;
                    case "mbsTransitionName":
                        if (v is string s1 && !string.IsNullOrWhiteSpace(s1))
                        {
                            mbsTransitionNameStr = s1.Trim();
                        }

                        break;
                    case "mbsName":
                        if (v is string s1b && !string.IsNullOrWhiteSpace(s1b))
                        {
                            mbsNameStr = s1b.Trim();
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

            // mlTransitionID 與 mlSimpleID 須分開讀：GetFields 順序不固定，若二者同列於結構且其一為 0，合併寫入同一欄會把有效 ID 蓋成 0。
            transitionId = mlTransitionIdRaw != 0 ? mlTransitionIdRaw : mlSimpleIdRaw;

            transitionName = PickPreferredTransitionDisplayName(moName, mbsTransitionNameStr, mbsNameStr);
            rawMbsTransitionName = (mbsTransitionNameStr ?? string.Empty).Trim();

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
                             string.Equals(n, "moName", StringComparison.OrdinalIgnoreCase))
                    {
                        transitionName = str;
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

            if (transitionId == 0)
            {
                transitionId = TryHarvestTransitionIdFromStructIntegralFields(t, info, fieldFlags);
            }

            if (string.IsNullOrWhiteSpace(rawMbsTransitionName))
            {
                rawMbsTransitionName = transitionName;
            }
        }

        /// <summary>
        /// 部分 Interop／批次回傳中 <c>mlTransitionID</c> 與 <c>mlSimpleID</c> 皆為 0，但其它整數欄位仍帶有效轉換 ID。
        /// </summary>
        private static int TryHarvestTransitionIdFromStructIntegralFields(
            Type structType,
            EdmChangeStateTransitionInfo info,
            BindingFlags fieldFlags)
        {
            var bestWeak = 0;
            foreach (var f in structType.GetFields(fieldFlags))
            {
                if (!IsIntegralType(f.FieldType))
                {
                    continue;
                }

                object v;
                try
                {
                    v = f.GetValue(info);
                }
                catch
                {
                    continue;
                }

                if (!TryConvertToInt32(v, out var iv) || iv <= 0)
                {
                    continue;
                }

                var n = f.Name ?? string.Empty;
                if (IsProbablyNotTransitionIdStructField(n))
                {
                    continue;
                }

                if (n.IndexOf("Transition", StringComparison.OrdinalIgnoreCase) >= 0 &&
                    (n.IndexOf("ID", StringComparison.OrdinalIgnoreCase) >= 0 ||
                     n.EndsWith("ID", StringComparison.OrdinalIgnoreCase)))
                {
                    return iv;
                }

                if (bestWeak == 0 &&
                    n.IndexOf("Simple", StringComparison.OrdinalIgnoreCase) >= 0 &&
                    (n.IndexOf("ID", StringComparison.OrdinalIgnoreCase) >= 0 ||
                     n.EndsWith("ID", StringComparison.OrdinalIgnoreCase)))
                {
                    bestWeak = iv;
                }
            }

            return bestWeak;
        }

        private static bool IsProbablyNotTransitionIdStructField(string fieldName)
        {
            if (string.IsNullOrEmpty(fieldName))
            {
                return true;
            }

            var n = fieldName.ToUpperInvariant();
            if (n.Contains("STATE") && n.IndexOf("TRANSITION", StringComparison.Ordinal) < 0)
            {
                return true;
            }

            if (n.Contains("WORKFLOW"))
            {
                return true;
            }

            if (n.Contains("FOLDER"))
            {
                return true;
            }

            if (n.Contains("FILE") && n.IndexOf("TRANSITION", StringComparison.Ordinal) < 0)
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// 當批次結構未帶轉換 ID（皆為 0）時，自目前檔案狀態列舉 <see cref="IEdmState6"/> 轉換，依名稱（含寬鬆比對）或列舉順序補齊 ID。
        /// </summary>
        private static void TryEnrichZeroTransitionIdsFromStateEnumeration(
            IEdmVault5 vault,
            string samplePath,
            List<PdmTransitionOption> options)
        {
            if (vault == null || options == null || options.Count == 0)
            {
                return;
            }

            if (options.TrueForAll(o => o.TransitionId != 0))
            {
                return;
            }

            var enumerated = TryEnumerateAllowedTransitionsFromSamplePath(vault, samplePath);
            if (enumerated == null || enumerated.Count == 0)
            {
                return;
            }

            var usedEnumIds = new HashSet<int>();
            foreach (var o in options)
            {
                if (o.TransitionId > 0)
                {
                    usedEnumIds.Add(o.TransitionId);
                }
            }

            foreach (var opt in options)
            {
                if (opt.TransitionId != 0 || string.IsNullOrWhiteSpace(opt.TransitionName))
                {
                    continue;
                }

                foreach (var e in enumerated)
                {
                    if (e.Id <= 0 || usedEnumIds.Contains(e.Id))
                    {
                        continue;
                    }

                    if (string.Equals(opt.TransitionName, e.Name, StringComparison.OrdinalIgnoreCase))
                    {
                        opt.TransitionId = e.Id;
                        usedEnumIds.Add(e.Id);
                        break;
                    }

                    if (TransitionDisplayNamesLikelyMatch(opt.TransitionName, e.Name))
                    {
                        opt.TransitionId = e.Id;
                        usedEnumIds.Add(e.Id);
                        break;
                    }
                }
            }

            // 批次與狀態列舉順序通常一致：已補上部分 ID 時，其餘仍為 0 者依索引對齊（不須「全部皆為 0」）。
            if (enumerated.Count == options.Count)
            {
                for (var i = 0; i < options.Count; i++)
                {
                    if (options[i].TransitionId != 0)
                    {
                        continue;
                    }

                    var eid = enumerated[i].Id;
                    if (eid <= 0)
                    {
                        continue;
                    }

                    if (options.Any(o => !ReferenceEquals(o, options[i]) && o.TransitionId == eid))
                    {
                        continue;
                    }

                    options[i].TransitionId = eid;
                }
            }
        }

        private static bool TransitionDisplayNamesLikelyMatch(string fromBatchDisplay, string fromTransitionCom)
        {
            if (string.IsNullOrWhiteSpace(fromBatchDisplay) || string.IsNullOrWhiteSpace(fromTransitionCom))
            {
                return false;
            }

            var a = fromBatchDisplay.Trim();
            var b = fromTransitionCom.Trim();
            if (a.Length >= 4 && b.IndexOf(a, StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return true;
            }

            if (b.Length >= 4 && a.IndexOf(b, StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return true;
            }

            var coreA = NormalizeTransitionNameForLooseCompare(a);
            var coreB = NormalizeTransitionNameForLooseCompare(b);
            if (!string.IsNullOrEmpty(coreA) &&
                !string.IsNullOrEmpty(coreB) &&
                string.Equals(coreA, coreB, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            // 例如「02_新圖面估價」vs「取消圖面估價」→ 核心「新圖面估價」與「圖面估價」需靠包含關係比對。
            if (!string.IsNullOrEmpty(coreA) && !string.IsNullOrEmpty(coreB) &&
                coreA.Length >= 3 && coreB.Length >= 3)
            {
                if (coreA.IndexOf(coreB, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }

                if (coreB.IndexOf(coreA, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }

            return false;
        }

        private static string NormalizeTransitionNameForLooseCompare(string s)
        {
            var t = s.Trim();
            while (t.Length > 0 && (char.IsDigit(t[0]) || t[0] == '_' || t[0] == '-'))
            {
                t = t.Substring(1);
            }

            if (t.StartsWith("取消", StringComparison.Ordinal))
            {
                t = t.Substring(2);
            }

            return t.Trim();
        }

        private sealed class EnumeratedTransition
        {
            public int Id { get; set; }
            public string Name { get; set; } = string.Empty;
        }

        private static List<EnumeratedTransition> TryEnumerateAllowedTransitionsFromSamplePath(
            IEdmVault5 vault,
            string samplePath)
        {
            var result = new List<EnumeratedTransition>();
            IEdmFolder5 folder = null;
            IEdmFile5 file = null;
            IEdmState5 state = null;
            IEdmPos5 pos = null;

            try
            {
                file = vault.GetFileFromPath(samplePath, out folder);
                if (file == null || folder == null)
                {
                    return result;
                }

                state = TryGetFileStateFromFile(file, folder) ?? file.CurrentState;
                if (state == null)
                {
                    return result;
                }

                var state6 = state as IEdmState6;
                if (state6 == null)
                {
                    return result;
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
                        return result;
                    }
                }

                if (pos == null || pos.IsNull)
                {
                    return result;
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

                        if (!TransitionAllowedForCurrentUser(tr))
                        {
                            continue;
                        }

                        TryReadTransitionIdAndName(tr, out var tid, out var tname, out _);
                        if (tid > 0)
                        {
                            result.Add(new EnumeratedTransition
                            {
                                Id = tid,
                                Name = (tname ?? string.Empty).Trim()
                            });
                        }
                    }
                    finally
                    {
                        ComHelper.Release(tr);
                    }
                }
            }
            catch
            {
                result.Clear();
            }
            finally
            {
                ComHelper.Release(pos);
                ComHelper.Release(state);
                ComHelper.Release(file);
                ComHelper.Release(folder);
            }

            return result;
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
