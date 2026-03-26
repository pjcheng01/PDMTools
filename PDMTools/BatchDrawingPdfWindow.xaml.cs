using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Windows;
using System.Windows.Forms;
using PDMTools.Models;
using PDMTools.Services;
using WinFormsDialog = System.Windows.Forms.FolderBrowserDialog;

namespace PDMTools
{
    public partial class BatchDrawingPdfWindow : Window
    {
        private readonly IReadOnlyList<BomItem> _drawingItems;
        private readonly PdmBomExportService _exportService;
        private CancellationTokenSource _cts;

        public BatchDrawingPdfWindow(IReadOnlyList<BomItem> drawingItems, PdmBomExportService exportService)
        {
            InitializeComponent();
            _drawingItems = drawingItems ?? Array.Empty<BomItem>();
            _exportService = exportService ?? throw new ArgumentNullException(nameof(exportService));

            var n = _drawingItems.Count;
            SummaryTextBlock.Text =
                $"將處理目前 DataGrid 篩選後可見之工程圖列（IsDrawing），共 {n} 筆（已依路徑去重）。\n"
                + "請選擇輸出資料夾；PDF 檔名為原工程圖主檔名（重名時自動加尾碼）。失敗列會寫入同資料夾之 CSV 報告。";
        }

        private void BrowseOutputFolder_OnClick(object sender, RoutedEventArgs e)
        {
            using (var dlg = new WinFormsDialog())
            {
                dlg.Description = "選擇 PDF 輸出資料夾";
                dlg.ShowNewFolderButton = true;
                if (!string.IsNullOrWhiteSpace(OutputFolderTextBox.Text))
                {
                    try
                    {
                        if (Directory.Exists(OutputFolderTextBox.Text.Trim()))
                        {
                            dlg.SelectedPath = Path.GetFullPath(OutputFolderTextBox.Text.Trim());
                        }
                    }
                    catch
                    {
                        // 忽略
                    }
                }

                if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    OutputFolderTextBox.Text = dlg.SelectedPath;
                }
            }
        }

        private void Start_OnClick(object sender, RoutedEventArgs e)
        {
            if (_drawingItems.Count == 0)
            {
                System.Windows.MessageBox.Show(this, "沒有可匯出的工程圖列（請確認 BOM 已抓取且篩選後可見之列含 .slddrw）。", "提醒", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var folder = OutputFolderTextBox.Text?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(folder))
            {
                System.Windows.MessageBox.Show(this, "請先選擇輸出資料夾。", "提醒", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            _cts = new CancellationTokenSource();
            StartButton.IsEnabled = false;
            CancelButton.IsEnabled = true;
            LogTextBox.Clear();
            BatchProgressBar.Value = 0;
            FooterTextBlock.Text = string.Empty;

            AppendLog("開始批次…");

            try
            {
                var svc = new BatchDrawingPdfService();
                var progress = new Progress<BatchDrawingPdfProgressInfo>(p =>
                {
                    if (p.Total > 0)
                    {
                        BatchProgressBar.Value = 100.0 * p.Current / p.Total;
                    }

                    AppendLog(p.Message);
                });

                var result = svc.Run(_drawingItems, folder, _exportService, progress, _cts.Token);

                FooterTextBlock.Text =
                    $"完成：成功 {result.SuccessCount}，失敗 {result.FailCount}。報告：{result.ReportPath ?? "（無）"}";
                AppendLog($"完成：成功 {result.SuccessCount}，失敗 {result.FailCount}。");
                if (!string.IsNullOrWhiteSpace(result.ReportPath))
                {
                    AppendLog("報告：" + result.ReportPath);
                }
            }
            catch (OperationCanceledException)
            {
                AppendLog("已取消。");
                FooterTextBlock.Text = "已取消。";
            }
            catch (Exception ex)
            {
                AppendLog("錯誤：" + ex.Message);
                FooterTextBlock.Text = "發生錯誤：" + ex.Message;
                System.Windows.MessageBox.Show(this, ex.Message, "錯誤", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                StartButton.IsEnabled = true;
                CancelButton.IsEnabled = false;
                _cts?.Dispose();
                _cts = null;
            }
        }

        private void Cancel_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                _cts?.Cancel();
            }
            catch
            {
                // 忽略
            }
        }

        private void AppendLog(string line)
        {
            if (string.IsNullOrEmpty(line))
            {
                return;
            }

            LogTextBox.AppendText(line + Environment.NewLine);
            LogTextBox.ScrollToEnd();
        }
    }
}
