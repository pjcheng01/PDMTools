using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using PDMTools.Models;
using SolidWorks.Interop.sldworks;

namespace PDMTools.Services
{
    /// <summary>
    /// 以 SolidWorks API（執行期 COM）遞迴展開單一頂層組合件之零組件引用，供「未入庫」稽核。
    /// 型別定義來自 NuGet <c>SolidWorks.Interop.sldworks</c>（建置不需本機安裝 SW）；執行仍須已註冊之 SolidWorks。
    /// </summary>
    public sealed class SolidWorksReferenceAuditService
    {
        // swDocumentTypes_e.swDocASSEMBLY
        private const int SwDocAssembly = 2;

        // swOpenDocOptions_e：Silent | ReadOnly
        private const int SwOpenSilentReadOnly = 1 | 2;

        /// <summary>
        /// 在呼叫執行緒上同步執行 COM（建議由 UI STA 執行緒呼叫）。
        /// 相同完整路徑僅保留一列；若多處引用會合併「上一層組合件路徑」並取較淺的 <see cref="ReferenceAuditRow.Level"/>。
        /// 略過組合件中已「使為虛擬」的零組件（不列入表、不遞迴其子階層）。
        /// </summary>
        public IReadOnlyList<ReferenceAuditRow> AuditAssembly(string assemblyFullPath, string vaultRootPath)
        {
            if (string.IsNullOrWhiteSpace(assemblyFullPath))
                throw new ArgumentException("請指定組合件路徑。", nameof(assemblyFullPath));

            var normalizedAsm = Path.GetFullPath(assemblyFullPath.Trim().Trim('"'));
            if (!File.Exists(normalizedAsm))
                throw new FileNotFoundException("找不到指定的組合件檔案。", normalizedAsm);

            if (!string.Equals(Path.GetExtension(normalizedAsm), ".sldasm", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("僅支援 .sldasm 組合件。");

            var vaultPrefix = NormalizeVaultPrefix(vaultRootPath);

            var swType = Type.GetTypeFromProgID("SldWorks.Application");
            if (swType == null)
                throw new InvalidOperationException(
                    "無法建立 SolidWorks 應用程式（找不到 ProgID SldWorks.Application）。請確認已安裝 SolidWorks。");

            var swApp = (SldWorks)Activator.CreateInstance(swType);
            if (swApp == null)
                throw new InvalidOperationException("無法啟動 SolidWorks 應用程式。");

            ModelDoc2 modelDoc = null;
            var weOpened = false;

            try
            {
                try
                {
                    swApp.Visible = false;
                }
                catch
                {
                    // 部分版本仍允許繼續
                }

                modelDoc = TryFindAlreadyOpenDocument(swApp, normalizedAsm);
                if (modelDoc == null)
                {
                    var err = 0;
                    var warn = 0;
                    modelDoc = (ModelDoc2)swApp.OpenDoc6(
                        normalizedAsm,
                        SwDocAssembly,
                        SwOpenSilentReadOnly,
                        "",
                        ref err,
                        ref warn);
                    if (modelDoc == null)
                        throw new InvalidOperationException(
                            $"SolidWorks 無法開啟組合件（錯誤碼 {err}，警告碼 {warn}）。請確認檔案可正常在 SolidWorks 中開啟。");
                    weOpened = true;
                }

                AssemblyDoc assyDoc;
                try
                {
                    assyDoc = (AssemblyDoc)modelDoc;
                }
                catch (InvalidCastException ex)
                {
                    throw new InvalidOperationException(
                        "開啟的文件不是組合件（無法取得 AssemblyDoc）。請確認檔案為有效的 .sldasm。", ex);
                }

                var byPath = new Dictionary<string, ReferenceAuditRow>(StringComparer.OrdinalIgnoreCase);
                UpsertRow(byPath, normalizedAsm, vaultPrefix, normalizedAsm, level: 0, parentAssemblyPath: string.Empty);

                var topRaw = assyDoc.GetComponents(true);
                foreach (var comp in ToComponentEnumerable(topRaw))
                {
                    if (comp != null && !IsVirtualizedComponent(comp))
                        VisitComponent(comp, normalizedAsm, normalizedAsm, vaultPrefix, byPath, level: 1);
                }

                return byPath.Values
                    .OrderBy(r => r.FullPath, StringComparer.OrdinalIgnoreCase)
                    .ThenBy(r => r.Level)
                    .ToList();
            }
            finally
            {
                if (weOpened && modelDoc != null)
                {
                    try
                    {
                        var title = modelDoc.GetTitle();
                        if (!string.IsNullOrEmpty(title))
                            swApp.CloseDoc(title);
                    }
                    catch
                    {
                        // 忽略關閉失敗
                    }
                }
            }
        }

        private static string NormalizeVaultPrefix(string vaultRootPath)
        {
            if (string.IsNullOrWhiteSpace(vaultRootPath))
                return string.Empty;
            try
            {
                var full = Path.GetFullPath(vaultRootPath.Trim())
                    .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                return full + Path.DirectorySeparatorChar;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string DedupeKeyForPath(string fullNorm)
        {
            if (string.IsNullOrWhiteSpace(fullNorm))
                return "\0__NO_PATH__";
            return fullNorm;
        }

        private static void UpsertRow(
            IDictionary<string, ReferenceAuditRow> byPath,
            string topAsm,
            string vaultPrefix,
            string path,
            int level,
            string parentAssemblyPath)
        {
            var p = path?.Trim() ?? string.Empty;
            var exists = !string.IsNullOrWhiteSpace(p) && File.Exists(p);
            var fullNorm = string.Empty;
            try
            {
                if (!string.IsNullOrWhiteSpace(p))
                    fullNorm = Path.GetFullPath(p);
            }
            catch
            {
                fullNorm = p;
            }

            var inVault = !string.IsNullOrEmpty(vaultPrefix) &&
                          !string.IsNullOrEmpty(fullNorm) &&
                          fullNorm.StartsWith(vaultPrefix, StringComparison.OrdinalIgnoreCase);

            var key = DedupeKeyForPath(fullNorm);

            if (!byPath.TryGetValue(key, out var row))
            {
                byPath[key] = new ReferenceAuditRow
                {
                    TopLevelAssemblyPath = topAsm,
                    Level = level,
                    ParentAssemblyPath = parentAssemblyPath?.Trim() ?? string.Empty,
                    FullPath = fullNorm,
                    FileName = string.IsNullOrWhiteSpace(p) ? string.Empty : Path.GetFileName(p),
                    Extension = string.IsNullOrWhiteSpace(p) ? string.Empty : Path.GetExtension(p),
                    IsUnderVaultRoot = inVault,
                    FileExists = exists
                };
                return;
            }

            row.Level = Math.Min(row.Level, level);
            MergeParentAssemblyPaths(row, parentAssemblyPath);
        }

        private static void MergeParentAssemblyPaths(ReferenceAuditRow row, string newParent)
        {
            if (string.IsNullOrWhiteSpace(newParent))
                return;
            newParent = newParent.Trim();
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var part in (row.ParentAssemblyPath ?? string.Empty).Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
            {
                var t = part.Trim();
                if (t.Length > 0)
                    set.Add(t);
            }

            set.Add(newParent);
            row.ParentAssemblyPath = string.Join("; ", set.OrderBy(s => s, StringComparer.OrdinalIgnoreCase));
        }

        /// <summary>
        /// 遞迴走訪零組件；<paramref name="immediateParentAsmPath"/> 為直接上層組合件 .sldasm 的完整路徑。
        /// </summary>
        private static void VisitComponent(
            Component2 comp,
            string immediateParentAsmPath,
            string topAsm,
            string vaultPrefix,
            IDictionary<string, ReferenceAuditRow> byPath,
            int level)
        {
            if (IsVirtualizedComponent(comp))
                return;

            string path = string.Empty;
            try
            {
                path = comp.GetPathName() ?? string.Empty;
            }
            catch (COMException)
            {
                path = string.Empty;
            }

            UpsertRow(byPath, topAsm, vaultPrefix, path, level, immediateParentAsmPath);

            string fullNorm = string.Empty;
            try
            {
                if (!string.IsNullOrWhiteSpace(path))
                    fullNorm = Path.GetFullPath(path.Trim());
            }
            catch
            {
                fullNorm = path?.Trim() ?? string.Empty;
            }

            // 子件的「上一層組合件」：若本件為子組合件則為本件路徑，否則沿用目前組合件脈絡
            var parentForChildren = IsAssemblyFilePath(fullNorm) ? fullNorm : immediateParentAsmPath;

            object childrenRaw;
            try
            {
                childrenRaw = comp.GetChildren();
            }
            catch (COMException)
            {
                return;
            }

            foreach (var child in ToComponentEnumerable(childrenRaw))
            {
                if (child != null && !IsVirtualizedComponent(child))
                    VisitComponent(child, parentForChildren, topAsm, vaultPrefix, byPath, level + 1);
            }
        }

        private static bool IsAssemblyFilePath(string fullNorm)
        {
            return !string.IsNullOrWhiteSpace(fullNorm) &&
                   fullNorm.EndsWith(".sldasm", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// 是否為組合件中「使為虛擬」的零組件（無獨立磁碟檔、不應納入引用稽核列）。
        /// </summary>
        private static bool IsVirtualizedComponent(Component2 comp)
        {
            if (comp == null)
                return true;
            try
            {
                // SolidWorks IComponent2::IsVirtual（Interop 為屬性，非方法）
                return comp.IsVirtual;
            }
            catch (Exception ex) when (ex is COMException or MissingMethodException or InvalidCastException
                                         or TargetInvocationException)
            {
                return TryReadIsVirtualViaReflection(comp);
            }
        }

        /// <summary>
        /// 部分執行期／Interop 差異下無法直接讀取屬性時，改以反射讀取 <c>IsVirtual</c> 屬性或同名方法。
        /// 若仍無法判斷則回傳 <c>false</c>（視為非虛擬，以免誤刪實體引用）。
        /// </summary>
        private static bool TryReadIsVirtualViaReflection(object comp)
        {
            if (comp == null)
                return false;
            try
            {
                var t = comp.GetType();
                var p = t.GetProperty("IsVirtual", BindingFlags.Instance | BindingFlags.Public);
                if (p != null)
                {
                    var r = p.GetValue(comp, null);
                    return CoerceToBool(r);
                }

                var m = t.GetMethod("IsVirtual", BindingFlags.Instance | BindingFlags.Public, null, Type.EmptyTypes, null);
                if (m != null)
                {
                    var r = m.Invoke(comp, null);
                    return CoerceToBool(r);
                }
            }
            catch
            {
                // ignore
            }

            return false;
        }

        private static bool CoerceToBool(object r)
        {
            return r switch
            {
                bool b => b,
                short s => s != 0,
                int i => i != 0,
                long l => l != 0,
                _ => false
            };
        }

        private static IEnumerable<Component2> ToComponentEnumerable(object v)
        {
            if (v == null)
                yield break;

            if (v is object[] arr)
            {
                foreach (var o in arr)
                {
                    if (o is Component2 c)
                        yield return c;
                }

                yield break;
            }

            if (v is Array a)
            {
                foreach (var o in a)
                {
                    if (o is Component2 c)
                        yield return c;
                }
            }
        }

        private static ModelDoc2 TryFindAlreadyOpenDocument(SldWorks swApp, string fullPath)
        {
            string normalizedTarget;
            try
            {
                normalizedTarget = Path.GetFullPath(fullPath);
            }
            catch
            {
                return null;
            }

            try
            {
                var doc = (ModelDoc2)swApp.GetFirstDocument();
                while (doc != null)
                {
                    string p = string.Empty;
                    try
                    {
                        p = doc.GetPathName() ?? string.Empty;
                    }
                    catch (COMException)
                    {
                        p = string.Empty;
                    }

                    if (!string.IsNullOrWhiteSpace(p))
                    {
                        try
                        {
                            if (string.Equals(Path.GetFullPath(p), normalizedTarget, StringComparison.OrdinalIgnoreCase))
                                return doc;
                        }
                        catch
                        {
                            // 略過無法正規化的路徑
                        }
                    }

                    try
                    {
                        doc = (ModelDoc2)doc.GetNext();
                    }
                    catch (COMException)
                    {
                        break;
                    }
                    catch (InvalidCastException)
                    {
                        break;
                    }
                }
            }
            catch (COMException)
            {
                return null;
            }
            catch (InvalidCastException)
            {
                return null;
            }

            return null;
        }
    }
}
