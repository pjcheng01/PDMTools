using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Threading;
using Microsoft.Win32;
using PDMTools.Converters;
using PDMTools.Models;
using PDMTools.Services;

namespace PDMTools
{
    public partial class MainWindow : Window
    {
        private const string VaultRootPath = @"C:\CP-PDM";
        private PdmBomExportService _exportService;
        private readonly CancellationTokenSource _cancellationTokenSource = new CancellationTokenSource();
        private readonly ObservableCollection<BomItem> _bomItems = new ObservableCollection<BomItem>();
        private readonly CardVariableLookupConverter _cardVariableConverter = new CardVariableLookupConverter();

        // ── 固定欄位定義（Level 不在此列，永遠顯示）──────────────────────
        private static readonly IReadOnlyList<string> AllFixedColumnNames = new List<string>
        {
            "File Name", "State", "Workflow State", "Description", "Part Number",
            "Referenced As", "Full Path",
            "Description Var Used", "Description Config Used",
            "Part Number Var Used", "Part Number Config Used"
        };

        // 目前作用中的固定欄位（null = 全顯示）
        private List<string> _activeFixedColumns = null;

        // 目前作用中的資料卡欄位清單（由設定視窗管理）
        private List<string> _activeCardVarNames = new List<string>();

        // ── 工程圖列相關狀態 ──────────────────────────────────────────────
        // 純零組件 BOM（不含工程圖列）；抓取後快取，用於 Toggle OFF 還原
        private List<BomItem> _rawBomItems = new List<BomItem>();
        // 含工程圖列的 BOM；Toggle ON 時由 AppendDrawingItemsAsync 建立並快取
        private List<BomItem> _bomItemsWithDrawings = null;

        public MainWindow()
        {
            InitializeComponent();
            SetProgressPercent(0);
            BomDataGrid.ItemsSource = _bomItems;

            // 載入已存的欄位設定
            var settings = ColumnSettings.Load();
            _activeCardVarNames  = settings.SelectedVariables ?? new List<string>();
            _activeFixedColumns  = settings.HasFixedColumnsSetting
                ? settings.SelectedFixedColumns
                : null;   // null = 尚未設定，預設全顯示

            // 若尚未設定卡片變數，使用內建清單作為起始預設
            if (_activeCardVarNames.Count == 0)
            {
                _activeCardVarNames = PdmBomExportService.GetOrderedCardVariableLabels().ToList();
            }

            SetupBomDataGridColumns();

            ResetConfigurationComboNoFile();

            StatusTextBlock.Text = ColumnSettings.FileExists
                ? $"就緒（已載入欄位設定：{_activeCardVarNames.Count} 個）"
                : "就緒（使用內建欄位設定，可按「設定欄位...」自訂）";

            // 視窗完全載入後，背景偵測 Vault 是否有新變數
            Loaded += async (s, e) => await CheckForNewVaultVariablesAsync();
        }

        // ── DataGrid 欄位建立 ──────────────────────────────────────────────

        /// <summary>依 _activeFixedColumns + _activeCardVarNames 建立 DataGrid 欄位。</summary>
        private void SetupBomDataGridColumns()
        {
            BomDataGrid.Columns.Clear();

            void Add(string header, Binding binding, double minWidth = 60, double maxWidth = double.PositiveInfinity)
            {
                var col = new DataGridTextColumn
                {
                    Header   = header,
                    Binding  = binding,
                    MinWidth = minWidth
                };
                if (!double.IsPositiveInfinity(maxWidth))
                    col.MaxWidth = maxWidth;
                BomDataGrid.Columns.Add(col);
            }

            // Level 永遠顯示
            Add("Level", new Binding("Level") { Mode = BindingMode.OneWay }, 50);

            // 用量統計：在父階底下，相同檔案路徑出現幾次
            Add("Use Count", new Binding("UsageCount") { Mode = BindingMode.OneWay }, 80);

            // 固定欄：依選擇決定是否加入（null = 全顯示）
            var fixedSet = _activeFixedColumns != null
                ? new HashSet<string>(_activeFixedColumns, StringComparer.OrdinalIgnoreCase)
                : null;   // null 代表全顯示

            bool ShowFixed(string name) => fixedSet == null || fixedSet.Contains(name);

            if (ShowFixed("File Name"))
                Add("File Name",            new Binding("FileName")              { Mode = BindingMode.OneWay }, 90);
            if (ShowFixed("State"))
                Add("State",                new Binding("State")                 { Mode = BindingMode.OneWay }, 70);
            if (ShowFixed("Workflow State"))
                Add("Workflow State",       new Binding("WorkflowState")         { Mode = BindingMode.OneWay }, 90);
            if (ShowFixed("Description"))
                Add("Description",          new Binding("Description")           { Mode = BindingMode.OneWay }, 80);
            if (ShowFixed("Part Number"))
                Add("Part Number",          new Binding("PartNumber")            { Mode = BindingMode.OneWay }, 80);
            if (ShowFixed("Referenced As"))
                Add("Referenced As",        new Binding("ReferencedAs")          { Mode = BindingMode.OneWay }, 90);
            if (ShowFixed("Full Path"))
                Add("Full Path",            new Binding("FullPath")              { Mode = BindingMode.OneWay }, 120, 520);
            if (ShowFixed("Description Var Used"))
                Add("Description Var Used", new Binding("DescriptionVarUsed")    { Mode = BindingMode.OneWay }, 80);
            if (ShowFixed("Description Config Used"))
                Add("Description Config Used", new Binding("DescriptionConfigUsed") { Mode = BindingMode.OneWay }, 80);
            if (ShowFixed("Part Number Var Used"))
                Add("Part Number Var Used", new Binding("PartNumberVarUsed")     { Mode = BindingMode.OneWay }, 80);
            if (ShowFixed("Part Number Config Used"))
                Add("Part Number Config Used", new Binding("PartNumberConfigUsed") { Mode = BindingMode.OneWay }, 80);

            // 資料卡變數欄
            foreach (var varName in _activeCardVarNames)
            {
                var binding = new Binding(".")
                {
                    Mode               = BindingMode.OneWay,
                    Converter          = _cardVariableConverter,
                    ConverterParameter = varName
                };
                Add("Card:" + varName, binding, 72);
            }
        }

        private void AutoSizeDataGridColumns()
        {
            BomDataGrid.UpdateLayout();
            foreach (var col in BomDataGrid.Columns)
                col.Width = new DataGridLength(1, DataGridLengthUnitType.SizeToCells);
        }

        // ── 按鈕事件 ──────────────────────────────────────────────────────

        private async void BrowseButton_OnClick(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Title = "選擇 SolidWorks 組合件",
                Filter = "SolidWorks Assembly (*.sldasm)|*.sldasm",
                InitialDirectory = Directory.Exists(VaultRootPath)
                    ? VaultRootPath
                    : Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
                CheckFileExists = true,
                Multiselect = false
            };

            if (dialog.ShowDialog(this) != true)
            {
                return;
            }

            AssemblyPathTextBox.Text = dialog.FileName;
            StatusTextBlock.Text = "已選擇組合件。";
            _bomItems.Clear();

            await PopulateConfigurationComboForPathAsync(dialog.FileName);
        }

        private async void StartGrabButton_OnClick(object sender, RoutedEventArgs e)
        {
            var assemblyPath = AssemblyPathTextBox.Text?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(assemblyPath))
            {
                MessageBox.Show(this, "請先選擇 .sldasm 檔案。", "提醒", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var includeRootBomItem = BomLevelDefinitionComboBox.SelectedIndex == 0; // A：含根

            int? maxBomLayerDepth = null;
            if (AllBomDepthCheckBox.IsChecked == true)
            {
                maxBomLayerDepth = null; // 取所有層
            }
            else
            {
                var txt = MaxBomDepthTextBox.Text?.Trim() ?? string.Empty;
                if (!int.TryParse(txt, out var maxDepth) || maxDepth <= 0)
                {
                    MessageBox.Show(this, "請輸入正確的「最大層數」。例如：3", "提醒", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                maxBomLayerDepth = maxDepth;
            }

            if (ConfigComboBox.SelectedItem == null || ConfigComboBox.Items.Count == 0)
            {
                MessageBox.Show(
                    this,
                    "請先以「瀏覽」選擇組合件，並等待組態清單載入完成後，再從「組態」下拉選單選取組態。",
                    "提醒",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            var configurationName = (ConfigComboBox.SelectedItem as string ?? ConfigComboBox.SelectedItem?.ToString() ?? string.Empty).Trim();
            if (string.IsNullOrEmpty(configurationName))
            {
                MessageBox.Show(this, "請在「組態」下拉選單選取一個組態。", "提醒", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            await RunGrabFlowAsync(assemblyPath, includeRootBomItem, maxBomLayerDepth, configurationName);
        }

        private async void ExportButton_OnClick(object sender, RoutedEventArgs e)
        {
            if (_bomItems.Count == 0)
            {
                MessageBox.Show(this, "請先按「開始抓取」取得資料後，再匯出 xlsx。", "提醒", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var assemblyPath = AssemblyPathTextBox.Text?.Trim() ?? string.Empty;
            var saveDialog = new SaveFileDialog
            {
                Title = "儲存 BOM Excel",
                Filter = "Excel Workbook (*.xlsx)|*.xlsx",
                FileName = string.IsNullOrWhiteSpace(assemblyPath)
                    ? $"BOM_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
                    : $"BOM_{Path.GetFileNameWithoutExtension(assemblyPath)}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx",
                AddExtension = true,
                OverwritePrompt = true
            };

            if (saveDialog.ShowDialog(this) != true)
                return;

            await RunExportExcelOnlyAsync(saveDialog.FileName);
        }

        private async void ShowDrawingsToggle_Click(object sender, RoutedEventArgs e)
        {
            var isOn = ShowDrawingsToggle.IsChecked == true;

            if (!isOn)
            {
                // Toggle OFF：還原純零組件 BOM
                _bomItems.Clear();
                foreach (var item in _rawBomItems) _bomItems.Add(item);
                StatusTextBlock.Text = $"已隱藏工程圖列，共 {_bomItems.Count} 筆。";
                return;
            }

            // Toggle ON：若已有快取直接使用，否則背景搜尋
            if (_bomItemsWithDrawings != null)
            {
                _bomItems.Clear();
                foreach (var item in _bomItemsWithDrawings) _bomItems.Add(item);
                var drwCount = _bomItemsWithDrawings.Count(x => x.IsDrawing);
                StatusTextBlock.Text = $"已顯示工程圖列（找到 {drwCount} 個工程圖）。";
                return;
            }

            SetUiBusy(true);
            SetProgressPercent(0);
            IProgress<ProgressInfo> progress = new Progress<ProgressInfo>(p =>
            {
                SetProgressPercent(p.Percentage);
                StatusTextBlock.Text = p.Message;
            });

            try
            {
                _exportService = _exportService ?? new PdmBomExportService();
                var withDrawings = await _exportService.AppendDrawingItemsAsync(
                    _rawBomItems,
                    _activeCardVarNames,
                    progress,
                    _cancellationTokenSource.Token);

                _bomItemsWithDrawings = withDrawings.ToList();

                _bomItems.Clear();
                foreach (var item in _bomItemsWithDrawings) _bomItems.Add(item);

                var drwCount = _bomItemsWithDrawings.Count(x => x.IsDrawing);
                StatusTextBlock.Text = drwCount > 0
                    ? $"已顯示工程圖列（找到 {drwCount} 個工程圖，共 {_bomItems.Count} 列）。"
                    : $"未找到任何工程圖，共 {_bomItems.Count} 筆零組件。";
                SetProgressPercent(0);

                await Dispatcher.InvokeAsync(AutoSizeDataGridColumns, DispatcherPriority.Loaded);
            }
            catch (OperationCanceledException)
            {
                ShowDrawingsToggle.IsChecked = false;
                MessageBox.Show(this, "作業已取消。", "取消", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                ShowDrawingsToggle.IsChecked = false;
                MessageBox.Show(this, ex.Message, "錯誤", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                SetUiBusy(false);
            }
        }

        private async void ColumnSettingsButton_OnClick(object sender, RoutedEventArgs e)
        {
            SetUiBusy(true);
            StatusTextBlock.Text = "正在從 Vault 讀取可用變數清單...";
            SetProgressPercent(0);

            IReadOnlyList<string> allVars = null;
            try
            {
                _exportService = _exportService ?? new PdmBomExportService();
                // 傳入目前選取的組合件路徑作為樣本，供 Vault 層級失敗時改從檔案列舉
                var samplePath = AssemblyPathTextBox.Text?.Trim();
                allVars = await _exportService.EnumerateVaultVariablesAsync(samplePath);
                SetProgressPercent(0);
            }
            catch (Exception ex)
            {
                StatusTextBlock.Text = $"無法讀取 Vault 變數（{ex.Message}），改用本機已知清單。";
                allVars = _activeCardVarNames.Count > 0
                    ? _activeCardVarNames
                    : (IReadOnlyList<string>)PdmBomExportService.GetOrderedCardVariableLabels();
            }
            finally
            {
                SetUiBusy(false);
            }

            // 若結果等於備援數量，顯示診斷資訊供除錯
            var fallbackCount = PdmBomExportService.GetOrderedCardVariableLabels().Count;
            if (allVars.Count == fallbackCount && !string.IsNullOrWhiteSpace(_exportService?.LastEnumerationDiag))
            {
                var diagMsg = "⚠ 動態列舉未取得結果，使用內建備援清單。\n\n"
                            + "── 診斷資訊 ────────────────────────────\n"
                            + _exportService.LastEnumerationDiag
                            + "\n請將以上內容回報給開發者以修正 API 呼叫。";
                MessageBox.Show(this, diagMsg, "變數列舉診斷",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }

            await OpenColumnSettingsWindowAsync(allVars);
        }

        /// <summary>
        /// 以指定的卡片變數清單開啟設定視窗，儲存結果後更新 DataGrid 欄位。
        /// </summary>
        private async Task OpenColumnSettingsWindowAsync(IReadOnlyList<string> allCardVars)
        {
            // 合併所有可選欄位：固定欄在前，卡片變數在後
            // 若卡片變數名稱與固定欄名稱重複（如 "Description"），不重複加入，避免設定視窗出現重複項目
            var fixedNamesSet = new HashSet<string>(AllFixedColumnNames, StringComparer.OrdinalIgnoreCase);
            var allItems = AllFixedColumnNames.ToList();
            allItems.AddRange(allCardVars.Where(v => !fixedNamesSet.Contains(v)));

            // 合併目前已選欄位（維持順序：固定欄在前，卡片欄在後）
            var currentFixed = _activeFixedColumns ?? AllFixedColumnNames.ToList();
            var currentSelected = currentFixed
                .Concat(_activeCardVarNames)
                .Where(x => allItems.Contains(x, StringComparer.OrdinalIgnoreCase))
                .ToList();

            var win = new ColumnSettingsWindow(allItems, currentSelected) { Owner = this };

            if (win.ShowDialog() != true)
            {
                StatusTextBlock.Text = _bomItems.Count > 0
                    ? $"已有 {_bomItems.Count} 筆資料，可匯出 xlsx。"
                    : "就緒";
                return;
            }

            // 分隔回固定欄位與卡片變數
            _activeFixedColumns = win.OrderedSelectedItems
                .Where(x => fixedNamesSet.Contains(x)).ToList();
            _activeCardVarNames = win.OrderedSelectedItems
                .Where(x => !fixedNamesSet.Contains(x)).ToList();

            var colSettings = new ColumnSettings
            {
                SelectedVariables    = _activeCardVarNames,
                SelectedFixedColumns = _activeFixedColumns,
                // KnownVariables 直接用 Vault 完整列舉清單，確保下次啟動比對時不會誤報
                // （不依賴 AllShownAvailableItems，避免固定欄名稱與卡片變數名稱重疊時漏記）
                KnownVariables       = allCardVars.ToList()
            };
            colSettings.Save();

            SetupBomDataGridColumns();

            if (_bomItems.Count > 0)
                await Dispatcher.InvokeAsync(AutoSizeDataGridColumns, DispatcherPriority.Loaded);

            var cardCnt  = _activeCardVarNames.Count;
            var fixedCnt = _activeFixedColumns.Count + 1; // +1 for Level
            StatusTextBlock.Text = _bomItems.Count > 0
                ? $"欄位已更新（固定 {fixedCnt} 欄 + 資料卡 {cardCnt} 欄）。請重新抓取以套用至資料。"
                : $"欄位設定已儲存（固定 {fixedCnt} 欄 + 資料卡 {cardCnt} 欄）。";
        }

        /// <summary>
        /// 啟動時背景靜默偵測：若 Vault 有新增變數（不在上次已知清單中），跳出提示。
        /// 任何例外均靜默忽略，不影響正常啟動。
        /// </summary>
        private async Task CheckForNewVaultVariablesAsync()
        {
            var storedKnown = ColumnSettings.Load().KnownVariables;

            // 尚未設定過（沒有 known list），略過偵測
            if (storedKnown == null || storedKnown.Count == 0) return;

            IReadOnlyList<string> vaultVars;
            try
            {
                _exportService = _exportService ?? new PdmBomExportService();
                vaultVars = await _exportService.EnumerateVaultVariablesAsync(
                    AssemblyPathTextBox.Text?.Trim());
            }
            catch
            {
                return; // Vault 不可用時靜默略過
            }

            var knownSet = new HashSet<string>(storedKnown, StringComparer.OrdinalIgnoreCase);
            var newVars  = new List<string>();
            foreach (var v in vaultVars)
            {
                if (!knownSet.Contains(v)) newVars.Add(v);
            }

            if (newVars.Count == 0) return;

            // 有新變數：跳出提示
            StatusTextBlock.Text = $"偵測到 Vault 新增了 {newVars.Count} 個資料卡變數。";

            var msg = $"偵測到 Vault 新增了 {newVars.Count} 個資料卡變數：\n\n"
                    + string.Join("\n", newVars)
                    + "\n\n是否立即開啟「設定欄位」視窗？";

            var answer = MessageBox.Show(this, msg, "偵測到新變數",
                MessageBoxButton.YesNo, MessageBoxImage.Information);

            if (answer == MessageBoxResult.Yes)
                await OpenColumnSettingsWindowAsync(vaultVars);
            else
                StatusTextBlock.Text = $"就緒（{newVars.Count} 個新變數尚未加入，可按「設定欄位...」設定）";
        }

        // ── 核心流程 ──────────────────────────────────────────────────────

        /// <summary>尚未選檔或無法載入組態時：清空下拉並停用，避免使用者誤以為可手動輸入。</summary>
        private void ResetConfigurationComboNoFile()
        {
            ConfigComboBox.Items.Clear();
            ConfigComboBox.SelectedItem = null;
            ConfigComboBox.IsEnabled = false;
        }

        /// <summary>
        /// 優先選取 PDM 回報之「文件作用中組態」（若存在於清單）；否則「預設」「Default」「默认」；再否則第一筆。
        /// </summary>
        private static string PickPreferredConfiguration(IReadOnlyList<string> cfgs, string documentActiveConfiguration)
        {
            if (cfgs == null || cfgs.Count == 0)
            {
                return null;
            }

            string Find(string name) =>
                cfgs.FirstOrDefault(x => string.Equals(x, name, StringComparison.OrdinalIgnoreCase));

            var active = documentActiveConfiguration?.Trim();
            if (!string.IsNullOrEmpty(active))
            {
                var hit = Find(active);
                if (!string.IsNullOrEmpty(hit))
                {
                    return hit;
                }
            }

            return Find("預設") ?? Find("Default") ?? Find("默认") ?? cfgs[0];
        }

        private static bool ComboBoxHasItemIgnoreCase(ComboBox box, string value)
        {
            foreach (var it in box.Items)
            {
                var s = it as string ?? it?.ToString();
                if (string.Equals(s, value, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        private async Task PopulateConfigurationComboForPathAsync(string assemblyPath)
        {
            if (string.IsNullOrWhiteSpace(assemblyPath) || !File.Exists(assemblyPath))
            {
                return;
            }

            try
            {
                _exportService = _exportService ?? new PdmBomExportService();
                var cfgs = await _exportService.GetAssemblyConfigurationsAsync(assemblyPath);
                var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var ordered = new List<string>();
                foreach (var c in cfgs)
                {
                    if (string.IsNullOrWhiteSpace(c))
                    {
                        continue;
                    }

                    var t = c.Trim();
                    if (seen.Add(t))
                    {
                        ordered.Add(t);
                    }
                }

                await Dispatcher.InvokeAsync(() =>
                {
                    ConfigComboBox.Items.Clear();
                    ConfigComboBox.SelectedItem = null;

                    if (ordered.Count == 0)
                    {
                        ConfigComboBox.IsEnabled = false;
                        StatusTextBlock.Text = "已選擇組合件，但無法從 PDM 讀取組態清單。";
                        var detail = _exportService?.LastConfigurationEnumerationDiag ?? string.Empty;
                        if (detail.Length > 2800)
                        {
                            detail = detail.Substring(0, 2800) + "\n…（以下略）";
                        }

                        MessageBox.Show(
                            this,
                            "無法從 PDM 讀取此組合件的組態清單，因此無法選擇組態。\n\n"
                            + "請確認：\n"
                            + "• 檔案已同步為本機最新版且可正常在 PDM 中開啟\n"
                            + "• 本機 PDM 用戶端與程式參考的 EPDM Interop 版本一致\n"
                            + "• 必要時請先以 SolidWorks／PDM 開啟該組合件一次後再試\n\n"
                            + (string.IsNullOrWhiteSpace(detail)
                                ? string.Empty
                                : "── 程式診斷（可複製給開發者）──\n" + detail),
                            "無法載入組態",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning);
                        return;
                    }

                    foreach (var c in ordered)
                    {
                        if (!ComboBoxHasItemIgnoreCase(ConfigComboBox, c))
                        {
                            ConfigComboBox.Items.Add(c);
                        }
                    }

                    var preferred = PickPreferredConfiguration(ordered, _exportService?.LastDocumentActiveConfiguration);
                    foreach (var it in ConfigComboBox.Items)
                    {
                        var s = it as string ?? it?.ToString();
                        if (string.Equals(s, preferred, StringComparison.OrdinalIgnoreCase))
                        {
                            ConfigComboBox.SelectedItem = it;
                            break;
                        }
                    }

                    ConfigComboBox.IsEnabled = true;
                    StatusTextBlock.Text =
                        $"已選擇組合件，已載入 {ordered.Count} 個組態；預設選取「{preferred}」（優先為文件作用中組態，可改選）。";
                });
            }
            catch (Exception ex)
            {
                await Dispatcher.InvokeAsync(() =>
                {
                    ConfigComboBox.Items.Clear();
                    ConfigComboBox.SelectedItem = null;
                    ConfigComboBox.IsEnabled = false;
                    StatusTextBlock.Text = "已選擇組合件，但讀取組態時發生錯誤。";
                    MessageBox.Show(
                        this,
                        "讀取組態清單時發生錯誤：\n" + ex.Message,
                        "錯誤",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                });
            }
        }

        private async Task RunGrabFlowAsync(
            string assemblyPath,
            bool includeRootBomItem,
            int? maxBomLayerDepth,
            string configurationName)
        {
            SetUiBusy(true);
            SetProgressPercent(0);

            IProgress<ProgressInfo> progress = new Progress<ProgressInfo>(p =>
            {
                SetProgressPercent(p.Percentage);
                StatusTextBlock.Text = p.Message;
            });

            try
            {
                _exportService = _exportService ?? new PdmBomExportService();
                progress.Report(new ProgressInfo(0, "開始抓取 BOM 與資料卡…"));
                _bomItems.Clear();

                var items = await _exportService.CollectBomAsync(
                    assemblyPath,
                    _activeCardVarNames,
                    includeRootBomItem,
                    maxBomLayerDepth,
                    configurationName,
                    progress,
                    _cancellationTokenSource.Token);

                // 快取純零組件 BOM，重置工程圖快取與 Toggle 狀態
                _rawBomItems           = items.ToList();
                _bomItemsWithDrawings  = null;
                ShowDrawingsToggle.IsChecked = false;
                ShowDrawingsToggle.IsEnabled = items.Count > 0;

                foreach (var item in items)
                    _bomItems.Add(item);

                var cfg = _exportService.LastConfigurationResolved;
                var layout = _exportService.LastBomLayoutNameUsed;
                StatusTextBlock.Text =
                    string.IsNullOrWhiteSpace(cfg) && string.IsNullOrWhiteSpace(layout)
                        ? $"抓取完成，共 {_bomItems.Count} 筆。"
                        : $"抓取完成，共 {_bomItems.Count} 筆。（組態：{cfg}；BOM 版面：{layout}）";
                SetProgressPercent(0);

                await Dispatcher.InvokeAsync(new Action(AutoSizeDataGridColumns), DispatcherPriority.Loaded);
            }
            catch (OperationCanceledException)
            {
                MessageBox.Show(this, "作業已取消。", "取消", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "錯誤", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                SetUiBusy(false);
            }
        }

        private async Task RunExportExcelOnlyAsync(string outputPath)
        {
            SetUiBusy(true);
            SetProgressPercent(0);

            IProgress<ProgressInfo> progress = new Progress<ProgressInfo>(p =>
            {
                SetProgressPercent(p.Percentage);
                StatusTextBlock.Text = p.Message;
            });

            try
            {
                _exportService = _exportService ?? new PdmBomExportService();
                progress.Report(new ProgressInfo(0, "正在匯出 Excel..."));

                await _exportService.ExportToExcelAsync(
                    _bomItems.ToList(),
                    _activeFixedColumns,   // null = 全部固定欄；有值 = 僅匯出使用者選取的
                    _activeCardVarNames,
                    outputPath,
                    progress,
                    _cancellationTokenSource.Token);

                MessageBox.Show(this, $"匯出成功！\n{outputPath}", "完成", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (OperationCanceledException)
            {
                MessageBox.Show(this, "作業已取消。", "取消", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (IOException ioEx)
            {
                MessageBox.Show(
                    this,
                    $"Excel 檔案可能被佔用或無法寫入：\n{ioEx.Message}",
                    "匯出失敗",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "錯誤", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                SetUiBusy(false);
                SetProgressPercent(0);
                StatusTextBlock.Text = _bomItems.Count > 0
                    ? $"抓取完成，共 {_bomItems.Count} 筆。（可匯出 xlsx）"
                    : "就緒";
            }
        }

        private void SetUiBusy(bool isBusy)
        {
            BrowseButton.IsEnabled         = !isBusy;
            StartGrabButton.IsEnabled      = !isBusy;
            ExportButton.IsEnabled         = !isBusy;
            ColumnSettingsButton.IsEnabled = !isBusy;
            // Toggle 只在有資料時才可操作，忙碌中一律禁用
            ShowDrawingsToggle.IsEnabled   = !isBusy && _rawBomItems.Count > 0;
            Mouse.OverrideCursor           = isBusy ? Cursors.Wait : null;
        }

        /// <summary>同步更新進度條與右側固定位置之百分比文字。</summary>
        private void SetProgressPercent(double value)
        {
            var v = Math.Max(0, Math.Min(100, value));
            ProgressBar.Value = v;
            ProgressPercentTextBlock.Text = $"{(int)Math.Round(v, MidpointRounding.AwayFromZero)}%";
        }
    }
}
