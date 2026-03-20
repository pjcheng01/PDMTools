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

        public MainWindow()
        {
            InitializeComponent();
            ProgressBar.Value = 0;
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

        private void BrowseButton_OnClick(object sender, RoutedEventArgs e)
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

            if (dialog.ShowDialog(this) == true)
            {
                AssemblyPathTextBox.Text = dialog.FileName;
                StatusTextBlock.Text = "已選擇組合件。";
                _bomItems.Clear();
            }
        }

        private async void StartGrabButton_OnClick(object sender, RoutedEventArgs e)
        {
            var assemblyPath = AssemblyPathTextBox.Text?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(assemblyPath))
            {
                MessageBox.Show(this, "請先選擇 .sldasm 檔案。", "提醒", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            await RunGrabFlowAsync(assemblyPath);
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

        private async void ColumnSettingsButton_OnClick(object sender, RoutedEventArgs e)
        {
            SetUiBusy(true);
            StatusTextBlock.Text = "正在從 Vault 讀取可用變數清單...";
            ProgressBar.Value = 0;

            IReadOnlyList<string> allVars = null;
            try
            {
                _exportService = _exportService ?? new PdmBomExportService();
                // 傳入目前選取的組合件路徑作為樣本，供 Vault 層級失敗時改從檔案列舉
                var samplePath = AssemblyPathTextBox.Text?.Trim();
                allVars = await _exportService.EnumerateVaultVariablesAsync(samplePath);
                ProgressBar.Value = 0;
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
            var allItems = AllFixedColumnNames.ToList();
            allItems.AddRange(allCardVars);

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
            var fixedSet = new HashSet<string>(AllFixedColumnNames, StringComparer.OrdinalIgnoreCase);
            _activeFixedColumns = win.OrderedSelectedItems
                .Where(x => fixedSet.Contains(x)).ToList();
            _activeCardVarNames = win.OrderedSelectedItems
                .Where(x => !fixedSet.Contains(x)).ToList();

            var colSettings = new ColumnSettings
            {
                SelectedVariables    = _activeCardVarNames,
                SelectedFixedColumns = _activeFixedColumns,
                KnownVariables       = win.AllShownAvailableItems
                    .Where(x => !fixedSet.Contains(x)).ToList()  // 只記卡片變數的 known list
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

        private async Task RunGrabFlowAsync(string assemblyPath)
        {
            SetUiBusy(true);

            IProgress<ProgressInfo> progress = new Progress<ProgressInfo>(p =>
            {
                ProgressBar.Value = Math.Max(0, Math.Min(100, p.Percentage));
                StatusTextBlock.Text = p.Message;
            });

            try
            {
                _exportService = _exportService ?? new PdmBomExportService();
                progress.Report(new ProgressInfo(0, "開始抓取 BOM 與資料卡..."));
                _bomItems.Clear();

                var items = await _exportService.CollectBomAsync(
                    assemblyPath,
                    _activeCardVarNames,
                    progress,
                    _cancellationTokenSource.Token);

                foreach (var item in items)
                    _bomItems.Add(item);

                StatusTextBlock.Text = $"抓取完成，共 {_bomItems.Count} 筆。";
                ProgressBar.Value = 0;

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

            IProgress<ProgressInfo> progress = new Progress<ProgressInfo>(p =>
            {
                ProgressBar.Value = Math.Max(0, Math.Min(100, p.Percentage));
                StatusTextBlock.Text = p.Message;
            });

            try
            {
                _exportService = _exportService ?? new PdmBomExportService();
                progress.Report(new ProgressInfo(0, "正在匯出 Excel..."));

                await _exportService.ExportToExcelAsync(
                    _bomItems.ToList(),
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
                ProgressBar.Value = 0;
                StatusTextBlock.Text = _bomItems.Count > 0
                    ? $"抓取完成，共 {_bomItems.Count} 筆。（可匯出 xlsx）"
                    : "就緒";
            }
        }

        private void SetUiBusy(bool isBusy)
        {
            BrowseButton.IsEnabled = !isBusy;
            StartGrabButton.IsEnabled = !isBusy;
            ExportButton.IsEnabled = !isBusy;
            ColumnSettingsButton.IsEnabled = !isBusy;
            Mouse.OverrideCursor = isBusy ? Cursors.Wait : null;
        }
    }
}
