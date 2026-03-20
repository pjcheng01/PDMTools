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

        // 目前作用中的資料卡欄位清單（由設定視窗管理）
        private List<string> _activeCardVarNames = new List<string>();

        public MainWindow()
        {
            InitializeComponent();
            ProgressBar.Value = 0;
            BomDataGrid.ItemsSource = _bomItems;

            // 載入已存的欄位設定
            var settings = ColumnSettings.Load();
            _activeCardVarNames = settings.SelectedVariables ?? new List<string>();

            // 若尚未設定，使用內建清單作為起始預設
            if (_activeCardVarNames.Count == 0)
            {
                _activeCardVarNames = PdmBomExportService.GetOrderedCardVariableLabels().ToList();
            }

            SetupBomDataGridColumns();

            StatusTextBlock.Text = ColumnSettings.FileExists
                ? $"就緒（已載入欄位設定：{_activeCardVarNames.Count} 個）"
                : "就緒（使用內建欄位設定，可按「設定欄位...」自訂）";
        }

        // ── DataGrid 欄位建立 ──────────────────────────────────────────────

        /// <summary>依 _activeCardVarNames 建立 DataGrid 欄位。</summary>
        private void SetupBomDataGridColumns()
        {
            BomDataGrid.Columns.Clear();

            void Add(string header, Binding binding, double minWidth = 60, double maxWidth = double.PositiveInfinity)
            {
                var col = new DataGridTextColumn
                {
                    Header = header,
                    Binding = binding,
                    MinWidth = minWidth
                };
                if (!double.IsPositiveInfinity(maxWidth))
                    col.MaxWidth = maxWidth;
                BomDataGrid.Columns.Add(col);
            }

            Add("Level",                  new Binding("Level")              { Mode = BindingMode.OneWay }, 50);
            Add("File Name",              new Binding("FileName")           { Mode = BindingMode.OneWay }, 90);
            Add("State",                  new Binding("State")              { Mode = BindingMode.OneWay }, 70);
            Add("Workflow State",         new Binding("WorkflowState")      { Mode = BindingMode.OneWay }, 90);
            Add("Description",            new Binding("Description")        { Mode = BindingMode.OneWay }, 80);
            Add("Part Number",            new Binding("PartNumber")         { Mode = BindingMode.OneWay }, 80);
            Add("Referenced As",          new Binding("ReferencedAs")       { Mode = BindingMode.OneWay }, 90);
            Add("Full Path",              new Binding("FullPath")           { Mode = BindingMode.OneWay }, 120, 520);
            Add("Description Var Used",   new Binding("DescriptionVarUsed") { Mode = BindingMode.OneWay }, 80);
            Add("Description Config",     new Binding("DescriptionConfigUsed") { Mode = BindingMode.OneWay }, 80);
            Add("Part Number Var Used",   new Binding("PartNumberVarUsed")  { Mode = BindingMode.OneWay }, 80);
            Add("Part Number Config",     new Binding("PartNumberConfigUsed") { Mode = BindingMode.OneWay }, 80);

            foreach (var varName in _activeCardVarNames)
            {
                var binding = new Binding(".")
                {
                    Mode = BindingMode.OneWay,
                    Converter = _cardVariableConverter,
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
                allVars = await _exportService.EnumerateVaultVariablesAsync();
                ProgressBar.Value = 0;
            }
            catch (Exception ex)
            {
                // 即使讀取失敗，仍用現有清單或內建清單開啟視窗
                StatusTextBlock.Text = $"無法讀取 Vault 變數（{ex.Message}），改用本機已知清單。";
                allVars = _activeCardVarNames.Count > 0
                    ? _activeCardVarNames
                    : (IReadOnlyList<string>)PdmBomExportService.GetOrderedCardVariableLabels();
            }
            finally
            {
                SetUiBusy(false);
            }

            // 開啟設定視窗
            var win = new ColumnSettingsWindow(allVars, _activeCardVarNames)
            {
                Owner = this
            };

            if (win.ShowDialog() != true)
            {
                StatusTextBlock.Text = _bomItems.Count > 0
                    ? $"已有 {_bomItems.Count} 筆資料，可匯出 xlsx。"
                    : "就緒";
                return;
            }

            // 套用新設定
            _activeCardVarNames = win.SelectedVariables ?? new List<string>();

            var colSettings = new ColumnSettings { SelectedVariables = _activeCardVarNames };
            colSettings.Save();

            // 重建 DataGrid 欄位
            SetupBomDataGridColumns();

            if (_bomItems.Count > 0)
                await Dispatcher.InvokeAsync(AutoSizeDataGridColumns, DispatcherPriority.Loaded);

            StatusTextBlock.Text = _bomItems.Count > 0
                ? $"欄位已更新（{_activeCardVarNames.Count} 個資料卡欄）。請重新抓取以套用至資料。"
                : $"欄位設定已儲存（{_activeCardVarNames.Count} 個資料卡欄）。";
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
