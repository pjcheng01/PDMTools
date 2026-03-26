using System;
using System.Collections.Generic;
using System.Threading;
using System.Windows;
using PDMTools.Models;
using PDMTools.Services;

namespace PDMTools
{
    public partial class DrawingSheetFormatWindow : Window
    {
        private readonly IReadOnlyList<BomItem> _drawingItems;
        private readonly PdmBomExportService _exportService;
        private readonly VaultStyleReportGridFilterHelper _gridFilter;
        private CancellationTokenSource _cts;

        public DrawingSheetFormatWindow(IReadOnlyList<BomItem> drawingItems, PdmBomExportService exportService)
        {
            InitializeComponent();
            _drawingItems = drawingItems ?? Array.Empty<BomItem>();
            _exportService = exportService ?? throw new ArgumentNullException(nameof(exportService));

            _gridFilter = new VaultStyleReportGridFilterHelper(this, ResultDataGrid, ResultFilterCountTextBlock);
            _gridFilter.SetupDrawingSheetFormatColumns();

            var n = _drawingItems.Count;
            SummaryTextBlock.Text =
                "僅處理目前 DataGrid 篩選後可見之工程圖列（IsDrawing、.slddrw），依路徑去重。\n"
                + $"共 {n} 個工程圖檔；每張圖可有多列（每圖頁一列），透過 SolidWorks API ISheet.GetTemplateName() 讀取 .slddrt 路徑。\n"
                + "各欄標題右側漏斗可開啟篩選（全選、清除選取、搜尋、勾選值），與主視窗 Vault BOM 相同。";
        }

        private void ColumnFilterButton_OnClick(object sender, RoutedEventArgs e)
        {
            _gridFilter.OnColumnFilterButtonClick(sender, e);
        }

        private void ClearColumnFilters_OnClick(object sender, RoutedEventArgs e)
        {
            _gridFilter.ClearAllFilters();
        }

        private void Start_OnClick(object sender, RoutedEventArgs e)
        {
            if (_drawingItems.Count == 0)
            {
                MessageBox.Show(this,
                    "沒有可查詢的工程圖。請先抓取 BOM，並確認篩選後可見之列含工程圖。",
                    "提醒", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            _cts = new CancellationTokenSource();
            StartButton.IsEnabled = false;
            CancelButton.IsEnabled = true;
            _gridFilter.ClearItemsSource();
            BatchProgressBar.Value = 0;
            ProgressText.Text = string.Empty;
            FooterTextBlock.Text = string.Empty;

            try
            {
                var svc = new DrawingSheetFormatQueryService();
                var progress = new Progress<DrawingSheetFormatProgressInfo>(p =>
                {
                    if (p.Total > 0)
                        BatchProgressBar.Value = 100.0 * p.Current / p.Total;
                    ProgressText.Text = p.Message;
                });

                var result = svc.Run(_drawingItems, _exportService, progress, _cts.Token);

                _gridFilter.SetItemsSource(result.Rows);
                FooterTextBlock.Text =
                    $"完成：成功 {result.SuccessCount} 列（圖頁），失敗 {result.FailCount} 列。";
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
