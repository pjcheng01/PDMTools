using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using EPDM.Interop.epdm;

namespace PdmVariableQuery
{
    /// <summary>PDM 變數 ↔ 自定屬性 對應查詢（獨立 WinForms 工具）。</summary>
    public sealed class MainForm : Form
    {
        private Label _lblVaultPath;
        private Button _btnLogin;
        private Button _btnExport;
        private TextBox _txtSearch;
        private Label _lblStatus;
        private DataGridView _grid;
        private Label _lblVault;
        private Label _lblSearch;
        private Panel _topPanel;
        private Label _lblTitle;
        private Label _lblCount;

        /// <summary>本機 PDM 視圖根目錄（與 Explorer 中 Vault 資料夾一致）。</summary>
        private const string VaultLocalRootPath = @"C:\CP-PDM";

        private readonly List<VariableMapping> _allMappings = new List<VariableMapping>();
        private EdmVault5 _vault;

        private sealed class VariableMapping
        {
            /// <summary>PDM 系統內建變數，或管理員在 Vault 中建立的自訂變數（依 <see cref="IEdmVariable5.VariableFlags"/> 的 EdmVar_Custom）。</summary>
            public string VarOrigin { get; set; }
            public string VarName { get; set; }
            public string VarType { get; set; }
            public string BlockName { get; set; }
            public string AttrName { get; set; }
            public string Remark { get; set; }
        }

        public MainForm()
        {
            BuildUi();
            InitializeVaultDisplay();
        }

        private void BuildUi()
        {
            Text = "PDM 變數 ↔ 自定屬性 對應查詢工具";
            Size = new Size(1000, 640);
            MinimumSize = new Size(800, 500);
            StartPosition = FormStartPosition.CenterScreen;
            Font = new Font("Microsoft JhengHei UI", 9.5f);
            BackColor = Color.FromArgb(245, 247, 250);

            _topPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 60,
                BackColor = Color.FromArgb(30, 80, 160)
            };
            _lblTitle = new Label
            {
                Text = "PDM 變數對應查詢工具",
                ForeColor = Color.White,
                Font = new Font("Microsoft JhengHei UI", 14f, FontStyle.Bold),
                AutoSize = true,
                Location = new Point(16, 16)
            };
            _topPanel.Controls.Add(_lblTitle);

            var ctrlPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 58,
                Padding = new Padding(12, 10, 12, 6),
                BackColor = Color.FromArgb(235, 240, 250)
            };

            _lblVault = new Label
            {
                Text = "本機視圖：",
                AutoSize = true,
                Location = new Point(12, 18),
                Font = new Font("Microsoft JhengHei UI", 9.5f, FontStyle.Bold)
            };

            _lblVaultPath = new Label
            {
                Location = new Point(92, 14),
                Width = 420,
                Height = 22,
                Text = VaultLocalRootPath,
                AutoEllipsis = true,
                ForeColor = Color.FromArgb(30, 60, 120),
                Font = new Font("Microsoft JhengHei UI", 9.5f, FontStyle.Bold)
            };

            _btnLogin = new Button
            {
                Text = "連線查詢",
                Location = new Point(520, 12),
                Width = 110,
                Height = 30,
                BackColor = Color.FromArgb(30, 80, 160),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            _btnLogin.FlatAppearance.BorderSize = 0;
            _btnLogin.Click += BtnLogin_Click;

            _lblSearch = new Label
            {
                Text = "搜尋：",
                AutoSize = true,
                Location = new Point(640, 18),
                Font = new Font("Microsoft JhengHei UI", 9.5f, FontStyle.Bold)
            };

            _txtSearch = new TextBox
            {
                Location = new Point(688, 14),
                Width = 180
            };
            _txtSearch.TextChanged += TxtSearch_TextChanged;
            var searchTip = new ToolTip { ToolTipTitle = "搜尋" };
            searchTip.SetToolTip(_txtSearch, "輸入變數名稱或自定屬性名稱…");

            _btnExport = new Button
            {
                Text = "匯出 CSV",
                Location = new Point(878, 12),
                Width = 110,
                Height = 30,
                BackColor = Color.FromArgb(46, 139, 87),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                Enabled = false
            };
            _btnExport.FlatAppearance.BorderSize = 0;
            _btnExport.Click += BtnExport_Click;

            ctrlPanel.Controls.AddRange(new Control[] { _lblVault, _lblVaultPath, _btnLogin, _lblSearch, _txtSearch, _btnExport });

            _grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                RowHeadersVisible = false,
                GridColor = Color.FromArgb(220, 225, 235),
                Font = new Font("Microsoft JhengHei UI", 9.5f),
                ColumnHeadersHeight = 34
            };

            _grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(30, 80, 160);
            _grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            _grid.ColumnHeadersDefaultCellStyle.Font = new Font("Microsoft JhengHei UI", 9.5f, FontStyle.Bold);
            _grid.ColumnHeadersDefaultCellStyle.Padding = new Padding(6, 0, 0, 0);
            _grid.EnableHeadersVisualStyles = false;

            _grid.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(245, 248, 255);
            _grid.DefaultCellStyle.Padding = new Padding(6, 2, 6, 2);
            _grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(180, 210, 255);
            _grid.DefaultCellStyle.SelectionForeColor = Color.Black;

            _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "VarName", HeaderText = "PDM 變數名稱", FillWeight = 20 });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "VarOrigin", HeaderText = "變數來源", FillWeight = 14 });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "VarType", HeaderText = "型別", FillWeight = 9 });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "BlockName", HeaderText = "Block Name", FillWeight = 16 });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "AttrName", HeaderText = "對應自定屬性名稱", FillWeight = 22 });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Remark", HeaderText = "備註", FillWeight = 19 });

            var bottomPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 28,
                BackColor = Color.FromArgb(220, 228, 245),
                Padding = new Padding(10, 4, 10, 4)
            };

            _lblCount = new Label
            {
                AutoSize = true,
                ForeColor = Color.FromArgb(30, 80, 160),
                Font = new Font("Microsoft JhengHei UI", 9f, FontStyle.Bold),
                Dock = DockStyle.Right,
                TextAlign = ContentAlignment.MiddleRight
            };
            _lblStatus = new Label
            {
                Text = "尚未連線",
                AutoSize = true,
                ForeColor = Color.FromArgb(80, 80, 100),
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft
            };
            bottomPanel.Controls.Add(_lblCount);
            bottomPanel.Controls.Add(_lblStatus);

            Controls.Add(_grid);
            Controls.Add(bottomPanel);
            Controls.Add(ctrlPanel);
            Controls.Add(_topPanel);
        }

        private void InitializeVaultDisplay()
        {
            SetStatus($"已固定本機視圖 {VaultLocalRootPath}，請按「連線查詢」", Color.Gray);
        }

        /// <summary>與 PDMTools <c>PdmBomExportService.EnsureVaultLogin</c> 相同：由路徑解析 Vault 名稱後再 <c>LoginAuto</c>。</summary>
        private static string TryLoginWithLocalPath(EdmVault5 vault, string localRootPath)
        {
            if (vault == null)
            {
                throw new ArgumentNullException(nameof(vault));
            }

            if (string.IsNullOrWhiteSpace(localRootPath))
            {
                throw new ArgumentException("本機視圖路徑不可為空。", nameof(localRootPath));
            }

            var candidates = new List<string>();
            var resolved = ResolveVaultNameFromPathSafe(vault, localRootPath);
            if (!string.IsNullOrWhiteSpace(resolved))
            {
                candidates.Add(resolved);
            }

            var folderName = Path.GetFileName(localRootPath.TrimEnd('\\', '/'));
            if (!string.IsNullOrWhiteSpace(folderName)
                && !candidates.Contains(folderName, StringComparer.OrdinalIgnoreCase))
            {
                candidates.Add(folderName);
            }

            if (candidates.Count == 0)
            {
                throw new InvalidOperationException($"找不到可用的 Vault 名稱，請確認本機視圖路徑是否有效：{localRootPath}");
            }

            Exception lastError = null;
            foreach (var candidate in candidates)
            {
                try
                {
                    vault.LoginAuto(candidate, 0);
                    if (vault.IsLoggedIn)
                    {
                        return candidate;
                    }
                }
                catch (Exception ex)
                {
                    lastError = ex;
                }
            }

            var joined = string.Join(", ", candidates);
            throw new InvalidOperationException(
                $"PDM 登入失敗。已嘗試 Vault 名稱：{joined}。請確認路徑 {localRootPath} 與本機視圖是否一致。",
                lastError);
        }

        private static string ResolveVaultNameFromPathSafe(EdmVault5 vault, string localPath)
        {
            try
            {
                var methods = vault.GetType().GetMethods(BindingFlags.Instance | BindingFlags.Public);
                foreach (var m in methods)
                {
                    if (!string.Equals(m.Name, "GetVaultNameFromPath", StringComparison.Ordinal))
                    {
                        continue;
                    }

                    try
                    {
                        var parameters = m.GetParameters();
                        if (parameters.Length == 1 && parameters[0].ParameterType == typeof(string))
                        {
                            var name = m.Invoke(vault, new object[] { localPath }) as string;
                            if (!string.IsNullOrWhiteSpace(name))
                            {
                                return name;
                            }
                        }

                        if (parameters.Length == 2
                            && parameters[0].ParameterType == typeof(string)
                            && parameters[1].ParameterType == typeof(string).MakeByRefType())
                        {
                            object[] args = { localPath, string.Empty };
                            m.Invoke(vault, args);
                            var name = args[1]?.ToString();
                            if (!string.IsNullOrWhiteSpace(name))
                            {
                                return name;
                            }
                        }
                    }
                    catch
                    {
                        // try next overload
                    }
                }
            }
            catch
            {
                // ignore
            }

            return string.Empty;
        }

        private void BtnLogin_Click(object sender, EventArgs e)
        {
            _btnLogin.Enabled = false;
            SetStatus("連線中...", Color.DodgerBlue);

            try
            {
                _vault = new EdmVault5();
                var vaultNameUsed = TryLoginWithLocalPath(_vault, VaultLocalRootPath);

                var userName = TryGetLoggedInUserName(_vault);
                _allMappings.Clear();
                _allMappings.AddRange(ReadMappings(_vault));
                RefreshGrid(_allMappings);

                _btnExport.Enabled = true;
                SetStatus($"已登入 Vault [{vaultNameUsed}]（{VaultLocalRootPath}），使用者：{userName}，共 {_allMappings.Count} 筆對應關係", Color.DarkGreen);
            }
            catch (Exception ex)
            {
                SetStatus($"連線失敗：{ex.Message}", Color.OrangeRed);
                MessageBox.Show($"連線失敗：\n\n{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _btnLogin.Enabled = true;
            }
        }

        /// <summary>
        /// 與 PDMTools PdmBomExportService 相同：以 GetFirstVariablePosition / GetNextVariable 列舉 Vault 變數
        /// （NuGet Interop 未提供 GetVariables / IEdmVariableEnumerator）。
        /// 屬性對應以 dynamic 呼叫 GetAttributes，避開介面上可能缺漏的成員。
        /// </summary>
        private static List<VariableMapping> ReadMappings(IEdmVault5 vault)
        {
            var result = new List<VariableMapping>();
            if (vault == null)
            {
                return result;
            }

            var vaultType = vault.GetType();
            var getFirstPos = vaultType.GetMethod("GetFirstVariablePosition",
                BindingFlags.Instance | BindingFlags.Public);
            var getNextVar = vaultType.GetMethod("GetNextVariable",
                BindingFlags.Instance | BindingFlags.Public);
            if (getFirstPos == null || getNextVar == null)
            {
                return result;
            }

            object pos;
            try
            {
                pos = getFirstPos.Invoke(vault, null);
            }
            catch
            {
                return result;
            }

            var typedPos = pos as IEdmPos5;
            if (typedPos == null)
            {
                return result;
            }

            var safetyLimit = 10000;
            while (!typedPos.IsNull && safetyLimit-- > 0)
            {
                object variable;
                try
                {
                    var args = new object[] { typedPos };
                    variable = getNextVar.Invoke(vault, args);
                }
                catch
                {
                    break;
                }

                if (variable == null)
                {
                    break;
                }

                var typedVar = variable as IEdmVariable5;
                if (typedVar == null || string.IsNullOrWhiteSpace(typedVar.Name))
                {
                    continue;
                }

                var typeName = FormatVariableType(GetVariableTypeBoxed(typedVar));
                var origin = ClassifyVariableOrigin(typedVar);
                AppendMappingsForVariable(result, typedVar, typeName, origin);
            }

            return result;
        }

        private static object GetVariableTypeBoxed(IEdmVariable5 variable)
        {
            if (variable == null)
            {
                return null;
            }

            try
            {
                return variable.VariableType;
            }
            catch
            {
                return null;
            }
        }

        private static void AppendMappingsForVariable(List<VariableMapping> result, IEdmVariable5 variable, string typeName, string varOrigin)
        {
            var hasAttr = false;
            try
            {
                dynamic dv = variable;
                dynamic attrEnum = dv.GetAttributes();
                while (attrEnum.MoveNext())
                {
                    dynamic attr = attrEnum.Current;
                    string blockName = attr.BlockName ?? "";
                    string attrName = attr.AttributeName ?? "";
                    var remark = "";
                    if (blockName == "CustomProperty")
                    {
                        remark = "對應到 SW 自定屬性（通用）";
                    }
                    else if (blockName == "$PRP")
                    {
                        remark = "對應到 SW 組態屬性";
                    }
                    else if (!string.IsNullOrWhiteSpace(blockName))
                    {
                        remark = $"Block: {blockName}";
                    }

                    result.Add(new VariableMapping
                    {
                        VarOrigin = varOrigin,
                        VarName = variable.Name,
                        VarType = typeName,
                        BlockName = blockName,
                        AttrName = attrName,
                        Remark = remark
                    });
                    hasAttr = true;
                }
            }
            catch
            {
                // GetAttributes 不存在或 COM 不支援時，改為單列提示
            }

            if (!hasAttr)
            {
                result.Add(new VariableMapping
                {
                    VarOrigin = varOrigin,
                    VarName = variable.Name,
                    VarType = typeName,
                    BlockName = "",
                    AttrName = "",
                    Remark = "此 PDM 變數無對應自定屬性（或無法列舉屬性對應）"
                });
            }
        }

        /// <summary>
        /// SolidWorks PDM：<c>IEdmVariable5.VariableFlags</c> 含 <c>EdmVar_Custom</c> 時為管理員在 Vault 建立的變數；否則為產品內建／系統保留變數。
        /// 以反射讀取旗標與列舉，相容不同 Interop 版本。
        /// </summary>
        private static string ClassifyVariableOrigin(IEdmVariable5 variable)
        {
            var flags = TryReadVariableFlags(variable);
            if (!flags.HasValue)
            {
                return "無法判斷";
            }

            var customMask = ResolveEdmVariableFlagMask(typeof(IEdmVariable5).Assembly, "EdmVar_Custom");
            if (customMask.HasValue)
            {
                return (flags.Value & customMask.Value) != 0 ? "管理員自訂" : "PDM 系統內建";
            }

            // 文件慣例：EdmVar_Custom = 0x1（若列舉未內嵌於組件則僅作備援）
            const long fallbackCustom = 1L;
            return (flags.Value & fallbackCustom) != 0 ? "管理員自訂（推測）" : "PDM 系統內建（推測）";
        }

        private static long? TryReadVariableFlags(object variable)
        {
            if (variable == null)
            {
                return null;
            }

            foreach (var t in EnumerateVariableApiTypes(variable))
            {
                foreach (var propName in new[] { "VariableFlags", "Flags", "EdmVariableFlags" })
                {
                    var p = t.GetProperty(propName, BindingFlags.Public | BindingFlags.Instance);
                    if (p == null)
                    {
                        continue;
                    }

                    try
                    {
                        var val = p.GetValue(variable, null);
                        if (val != null)
                        {
                            return Convert.ToInt64(val);
                        }
                    }
                    catch
                    {
                        // 下一個屬性
                    }
                }
            }

            return null;
        }

        private static IEnumerable<Type> EnumerateVariableApiTypes(object variable)
        {
            yield return typeof(IEdmVariable5);
            var rt = variable.GetType();
            yield return rt;
            foreach (var i in rt.GetInterfaces())
            {
                if (i.Name.IndexOf("Variable", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    yield return i;
                }
            }
        }

        private static long? ResolveEdmVariableFlagMask(Assembly interopAssembly, string enumMemberName)
        {
            if (interopAssembly == null || string.IsNullOrEmpty(enumMemberName))
            {
                return null;
            }

            foreach (var enumType in interopAssembly.GetTypes())
            {
                if (!enumType.IsEnum || !string.Equals(enumType.Name, "EdmVariableFlag", StringComparison.Ordinal))
                {
                    continue;
                }

                try
                {
                    foreach (var n in Enum.GetNames(enumType))
                    {
                        if (string.Equals(n, enumMemberName, StringComparison.OrdinalIgnoreCase))
                        {
                            return Convert.ToInt64(Enum.Parse(enumType, n));
                        }
                    }
                }
                catch
                {
                    // 下一個型別
                }
            }

            return null;
        }

        private static string TryGetLoggedInUserName(object vault)
        {
            if (vault == null)
            {
                return "（未知）";
            }

            try
            {
                dynamic d = vault;
                var u = d.CurrentUser;
                if (u != null)
                {
                    var name = u.Name;
                    if (name != null)
                    {
                        return name.ToString();
                    }
                }
            }
            catch
            {
                // 略過，改試反射
            }

            try
            {
                var t = vault.GetType();
                foreach (var p in t.GetProperties(BindingFlags.Instance | BindingFlags.Public))
                {
                    if (p.Name.IndexOf("User", StringComparison.OrdinalIgnoreCase) < 0)
                    {
                        continue;
                    }

                    var o = p.GetValue(vault);
                    if (o == null)
                    {
                        continue;
                    }

                    var ot = o.GetType();
                    var np = ot.GetProperty("Name", BindingFlags.Instance | BindingFlags.Public);
                    if (np != null)
                    {
                        var n = np.GetValue(o);
                        if (n != null)
                        {
                            return n.ToString();
                        }
                    }

                    return o.ToString();
                }
            }
            catch
            {
                // ignore
            }

            return "（未知）";
        }

        private void TxtSearch_TextChanged(object sender, EventArgs e)
        {
            var keyword = _txtSearch.Text.Trim().ToLowerInvariant();

            if (string.IsNullOrEmpty(keyword))
            {
                RefreshGrid(_allMappings);
                return;
            }

            var filtered = _allMappings.Where(m =>
                (m.VarName ?? "").ToLowerInvariant().Contains(keyword) ||
                (m.VarOrigin ?? "").ToLowerInvariant().Contains(keyword) ||
                (m.AttrName ?? "").ToLowerInvariant().Contains(keyword) ||
                (m.BlockName ?? "").ToLowerInvariant().Contains(keyword) ||
                (m.VarType ?? "").ToLowerInvariant().Contains(keyword)).ToList();

            RefreshGrid(filtered);
        }

        private void RefreshGrid(List<VariableMapping> data)
        {
            _grid.Rows.Clear();

            foreach (var m in data)
            {
                var idx = _grid.Rows.Add(m.VarName, m.VarOrigin, m.VarType, m.BlockName, m.AttrName, m.Remark);

                if (string.IsNullOrEmpty(m.AttrName))
                {
                    _grid.Rows[idx].DefaultCellStyle.ForeColor = Color.Gray;
                }

                if (m.BlockName == "CustomProperty")
                {
                    _grid.Rows[idx].Cells["BlockName"].Style.ForeColor = Color.FromArgb(30, 80, 160);
                }

                var originCell = _grid.Rows[idx].Cells["VarOrigin"];
                if (m.VarOrigin != null && m.VarOrigin.StartsWith("管理員", StringComparison.Ordinal))
                {
                    originCell.Style.ForeColor = Color.FromArgb(0, 100, 60);
                    originCell.Style.Font = new Font(_grid.Font, FontStyle.Bold);
                }
                else if (m.VarOrigin != null && m.VarOrigin.StartsWith("PDM 系統內建", StringComparison.Ordinal))
                {
                    originCell.Style.ForeColor = Color.FromArgb(40, 60, 120);
                }
            }

            _lblCount.Text = $"顯示 {data.Count} 筆";
        }

        private void BtnExport_Click(object sender, EventArgs e)
        {
            using (var dlg = new SaveFileDialog())
            {
                dlg.Filter = "CSV 檔案 (*.csv)|*.csv";
                dlg.FileName = $"PDM_Variables_{DateTime.Now:yyyyMMdd_HHmm}.csv";

                if (dlg.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                try
                {
                    var lines = new List<string> { "PDM變數名稱,變數來源,型別,Block Name,對應自定屬性名稱,備註" };
                    lines.AddRange(_allMappings.Select(m =>
                        $"\"{m.VarName}\",\"{m.VarOrigin}\",\"{m.VarType}\",\"{m.BlockName}\",\"{m.AttrName}\",\"{m.Remark}\""));

                    File.WriteAllLines(dlg.FileName, lines, new UTF8Encoding(true));

                    SetStatus($"已匯出至：{dlg.FileName}", Color.DarkGreen);
                    MessageBox.Show("匯出成功！", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"匯出失敗：{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void SetStatus(string msg, Color color)
        {
            _lblStatus.Text = msg;
            _lblStatus.ForeColor = color;
        }

        /// <summary>Interop 各版 EdmVariableType 成員名稱不一，改以字串比對顯示中文。</summary>
        private static string FormatVariableType(object vt)
        {
            if (vt == null)
            {
                return "";
            }

            var s = vt.ToString();
            if (string.IsNullOrEmpty(s))
            {
                return "";
            }

            if (s.IndexOf("Text", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "文字";
            }

            if (s.IndexOf("Integer", StringComparison.OrdinalIgnoreCase) >= 0
                || s.IndexOf("Int", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "整數";
            }

            if (s.IndexOf("Float", StringComparison.OrdinalIgnoreCase) >= 0
                || s.IndexOf("Double", StringComparison.OrdinalIgnoreCase) >= 0
                || s.IndexOf("Real", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "浮點數";
            }

            if (s.IndexOf("Date", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "日期";
            }

            if (s.IndexOf("Bool", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "布林";
            }

            return s;
        }
    }
}
