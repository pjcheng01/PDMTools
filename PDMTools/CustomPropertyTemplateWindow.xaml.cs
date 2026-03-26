using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Windows;
using PDMTools.Models;
using PDMTools.Services;

namespace PDMTools
{
    public partial class CustomPropertyTemplateWindow : Window
    {
        private readonly IReadOnlyList<BomItem> _bomItems;
        private readonly PdmBomExportService _exportService;
        private CancellationTokenSource _cts;

        public CustomPropertyTemplateWindow(IReadOnlyList<BomItem> bomItems, PdmBomExportService exportService)
        {
            InitializeComponent();
            _bomItems = bomItems ?? Array.Empty<BomItem>();
            _exportService = exportService ?? throw new ArgumentNullException(nameof(exportService));

            var extGroups = _bomItems
                .Where(i => !string.IsNullOrWhiteSpace(i.FullPath))
                .Select(i => System.IO.Path.GetExtension(i.FullPath).ToLowerInvariant())
                .Where(e => e == ".sldprt" || e == ".sldasm" || e == ".slddrw")
                .GroupBy(e => e)
                .OrderBy(g => g.Key)
                .Select(g => $"{g.Key}：{g.Count()} 筆")
                .ToList();

            SummaryTextBlock.Text =
                $"將查詢目前 DataGrid 篩選後可見之 SolidWorks 檔案所套用的「自訂屬性範本」。\n"
                + $"共 {_bomItems.Count} 筆（{string.Join("、", extGroups)}）。";
        }

        private void Start_OnClick(object sender, RoutedEventArgs e)
        {
            if (_bomItems.Count == 0)
            {
                MessageBox.Show(this,
                    "沒有可查詢的 SolidWorks 檔案。請確認 BOM 已抓取且篩選後有可見的 .sldprt / .sldasm / .slddrw。",
                    "提醒", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            _cts = new CancellationTokenSource();
            StartButton.IsEnabled = false;
            CancelButton.IsEnabled = true;
            ResultDataGrid.ItemsSource = null;
            BatchProgressBar.Value = 0;
            ProgressText.Text = string.Empty;
            FooterTextBlock.Text = string.Empty;

            try
            {
                var svc = new CustomPropertyTemplateService();
                var progress = new Progress<CustomPropertyTemplateProgressInfo>(p =>
                {
                    if (p.Total > 0)
                        BatchProgressBar.Value = 100.0 * p.Current / p.Total;
                    ProgressText.Text = p.Message;
                });

                var result = svc.Run(_bomItems, _exportService, progress, _cts.Token);

                ResultDataGrid.ItemsSource = result.Rows;
                FooterTextBlock.Text =
                    $"完成：成功 {result.SuccessCount}，失敗 {result.FailCount}。";
            }
            catch (OperationCanceledException)
            {
                FooterTextBlock.Text = "已取消。";
            }
            catch (Exception ex)
            {
                FooterTextBlock.Text = "發生錯誤：" + ex.Message;
                MessageBox.Show(this, ex.Message, "錯誤", MessageBoxButton.OK, MessageBoxImage.Error);
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
            try { _cts?.Cancel(); } catch { }
        }
    }
}
