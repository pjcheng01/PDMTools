using System;
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

        public MainWindow()
        {
            InitializeComponent();
            ProgressBar.Value = 0;
            StatusTextBlock.Text = "就緒";
            BomDataGrid.ItemsSource = _bomItems;
            SetupBomDataGridColumns();
        }

        /// <summary>建立與 Excel 匯出相同順序的欄位（含各 Card: 變數獨立欄）。</summary>
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
                {
                    col.MaxWidth = maxWidth;
                }

                BomDataGrid.Columns.Add(col);
            }

            Add("Level", new Binding("Level") { Mode = BindingMode.OneWay }, 50);
            Add("File Name", new Binding("FileName") { Mode = BindingMode.OneWay }, 90);
            Add("State", new Binding("State") { Mode = BindingMode.OneWay }, 70);
            Add("Workflow State", new Binding("WorkflowState") { Mode = BindingMode.OneWay }, 90);
            Add("Description", new Binding("Description") { Mode = BindingMode.OneWay }, 80);
            Add("Part Number", new Binding("PartNumber") { Mode = BindingMode.OneWay }, 80);
            Add("Referenced As", new Binding("ReferencedAs") { Mode = BindingMode.OneWay }, 90);
            Add("Full Path", new Binding("FullPath") { Mode = BindingMode.OneWay }, 120, 520);
            Add("Description Var Used", new Binding("DescriptionVarUsed") { Mode = BindingMode.OneWay }, 80);
            Add("Description Config Used", new Binding("DescriptionConfigUsed") { Mode = BindingMode.OneWay }, 80);
            Add("Part Number Var Used", new Binding("PartNumberVarUsed") { Mode = BindingMode.OneWay }, 80);
            Add("Part Number Config Used", new Binding("PartNumberConfigUsed") { Mode = BindingMode.OneWay }, 80);

            foreach (var label in PdmBomExportService.GetOrderedCardVariableLabels())
            {
                var binding = new Binding(".")
                {
                    Mode = BindingMode.OneWay,
                    Converter = _cardVariableConverter,
                    ConverterParameter = label
                };
                Add("Card:" + label, binding, 72);
            }
        }

        /// <summary>依儲存格內容自動調整欄寬（與 Excel AdjustToContents 類似）。</summary>
        private void AutoSizeDataGridColumns()
        {
            BomDataGrid.UpdateLayout();
            foreach (var col in BomDataGrid.Columns)
            {
                col.Width = new DataGridLength(1, DataGridLengthUnitType.SizeToCells);
            }
        }

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
            {
                return;
            }

            await RunExportExcelOnlyAsync(saveDialog.FileName);
        }

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
                    progress,
                    _cancellationTokenSource.Token);

                foreach (var item in items)
                {
                    _bomItems.Add(item);
                }

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
            Mouse.OverrideCursor = isBusy ? System.Windows.Input.Cursors.Wait : null;
        }
    }
}
