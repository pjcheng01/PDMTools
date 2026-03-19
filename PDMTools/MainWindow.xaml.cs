using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using Microsoft.Win32;
using PDMTools.Models;
using PDMTools.Services;

namespace PDMTools
{
    public partial class MainWindow : Window
    {
        private const string VaultRootPath = @"C:\CP-PDM";
        private PdmBomExportService? _exportService;
        private readonly CancellationTokenSource _cancellationTokenSource = new CancellationTokenSource();

        public MainWindow()
        {
            InitializeComponent();
            ProgressBar.Value = 0;
            StatusTextBlock.Text = "就緒";
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
            }
        }

        private async void ExportButton_OnClick(object sender, RoutedEventArgs e)
        {
            var assemblyPath = AssemblyPathTextBox.Text?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(assemblyPath))
            {
                MessageBox.Show(this, "請先選擇 .sldasm 檔案。", "提醒", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var saveDialog = new SaveFileDialog
            {
                Title = "儲存 BOM Excel",
                Filter = "Excel Workbook (*.xlsx)|*.xlsx",
                FileName = $"BOM_{Path.GetFileNameWithoutExtension(assemblyPath)}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx",
                AddExtension = true,
                OverwritePrompt = true
            };

            if (saveDialog.ShowDialog(this) != true)
            {
                return;
            }

            await RunExportFlowAsync(assemblyPath, saveDialog.FileName);
        }

        private async Task RunExportFlowAsync(string assemblyPath, string outputPath)
        {
            SetUiBusy(true);

            IProgress<ProgressInfo> progress = new Progress<ProgressInfo>(p =>
            {
                ProgressBar.Value = Math.Max(0, Math.Min(100, p.Percentage));
                StatusTextBlock.Text = p.Message;
            });

            try
            {
                _exportService ??= new PdmBomExportService();
                progress.Report(new ProgressInfo(0, "開始執行匯出..."));
                var items = await _exportService.CollectBomAsync(
                    assemblyPath,
                    progress,
                    _cancellationTokenSource.Token);

                await _exportService.ExportToExcelAsync(
                    items,
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
                if (ProgressBar.Value >= 100)
                {
                    StatusTextBlock.Text = "就緒";
                    ProgressBar.Value = 0;
                }
            }
        }

        private void SetUiBusy(bool isBusy)
        {
            BrowseButton.IsEnabled = !isBusy;
            ExportButton.IsEnabled = !isBusy;
            Mouse.OverrideCursor = isBusy ? System.Windows.Input.Cursors.Wait : null;
        }
    }
}
