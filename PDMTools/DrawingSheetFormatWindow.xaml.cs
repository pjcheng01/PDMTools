using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Windows;
using System.Windows.Data;
using Microsoft.Win32;
using PDMTools.Models;
using PDMTools.Services;

namespace PDMTools
{
    public partial class DrawingSheetFormatWindow : Window
    {
        private readonly IReadOnlyList<BomItem> _drawingItems;
        private readonly PdmBomExportService _exportService;
        private readonly VaultStyleReportGridFilterHelper _gridFilter;
        private CancellationTokenSource _queryCts;
        private CancellationTokenSource _applyCts;

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
                + "各欄標題右側漏斗可開啟篩選；勾選「套用」欄後可指定圖頁格式範本並批次套用（取檔並 check out、SetTemplateName、存檔，不 check in）。\n"
                + "同一工程圖多列勾選時僅開檔／存檔一次。";

            ApplySheetFormatButton.IsEnabled = false;
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

            _queryCts = new CancellationTokenSource();
            StartButton.IsEnabled = false;
            ApplySheetFormatButton.IsEnabled = false;
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

                var result = svc.Run(_drawingItems, _exportService, progress, _queryCts.Token);

                _gridFilter.SetItemsSource(result.Rows);
                ApplySheetFormatButton.IsEnabled = result.Rows.Count > 0;
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
                _queryCts?.Dispose();
                _queryCts = null;
            }
        }

        private void Cancel_OnClick(object sender, RoutedEventArgs e)
        {
            try { _queryCts?.Cancel(); } catch { }
            try { _applyCts?.Cancel(); } catch { }
        }

        private static List<DrawingSheetFormatRow> MaterializeRows(IEnumerable source)
        {
            var list = new List<DrawingSheetFormatRow>();
            if (source == null)
                return list;
            foreach (var o in source)
            {
                if (o is DrawingSheetFormatRow r)
                    list.Add(r);
            }
            return list;
        }

        private void BrowseSheetFormatTemplate_OnClick(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                CheckFileExists = true,
                Filter = "圖頁格式範本 (*.slddrt)|*.slddrt|所有檔案|*.*",
                DefaultExt = ".slddrt"
            };
            if (dlg.ShowDialog(this) == true)
                SheetFormatTemplatePathTextBox.Text = dlg.FileName;
        }

        private void SelectAllVisibleRows_OnClick(object sender, RoutedEventArgs e)
        {
            var view = CollectionViewSource.GetDefaultView(ResultDataGrid.ItemsSource);
            if (view == null)
                return;
            foreach (var o in view)
            {
                if (o is DrawingSheetFormatRow r)
                    r.IsSelectedForApply = true;
            }
        }

        private void ClearVisibleSelection_OnClick(object sender, RoutedEventArgs e)
        {
            var view = CollectionViewSource.GetDefaultView(ResultDataGrid.ItemsSource);
            if (view == null)
                return;
            foreach (var o in view)
            {
                if (o is DrawingSheetFormatRow r)
                    r.IsSelectedForApply = false;
            }
        }

        private void ApplySheetFormat_OnClick(object sender, RoutedEventArgs e)
        {
            var rows = MaterializeRows(ResultDataGrid.ItemsSource);
            if (rows.Count == 0)
            {
                MessageBox.Show(this, "請先執行查詢。", "提醒", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var path = SheetFormatTemplatePathTextBox.Text?.Trim() ?? string.Empty;
            if (string.IsNullOrEmpty(path) || !File.Exists(path))
            {
                MessageBox.Show(this, "請先指定存在的圖頁格式範本（.slddrt）。", "提醒", MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            _applyCts = new CancellationTokenSource();
            ApplySheetFormatButton.IsEnabled = false;
            StartButton.IsEnabled = false;
            CancelButton.IsEnabled = true;
            BatchProgressBar.Value = 0;
            ProgressText.Text = string.Empty;

            try
            {
                var svc = new DrawingSheetFormatQueryService();
                var progress = new Progress<DrawingSheetFormatProgressInfo>(p =>
                {
                    if (p.Total > 0)
                        BatchProgressBar.Value = 100.0 * p.Current / p.Total;
                    ProgressText.Text = p.Message;
                });

                var applyResult = svc.ApplySheetFormatsToRows(
                    rows,
                    path,
                    _exportService,
                    progress,
                    _applyCts.Token,
                    out var validationError);

                if (!string.IsNullOrEmpty(validationError))
                {
                    MessageBox.Show(this, validationError, "無法套用", MessageBoxButton.OK, MessageBoxImage.Warning);
                    FooterTextBlock.Text = validationError;
                    return;
                }

                FooterTextBlock.Text =
                    $"套用完成：成功 {applyResult.SuccessCount} 列，失敗 {applyResult.FailCount} 列。";
            }
            catch (OperationCanceledException)
            {
                FooterTextBlock.Text = "套用已取消。";
            }
            catch (Exception ex)
            {
                FooterTextBlock.Text = "套用發生錯誤：" + ex.Message;
                MessageBox.Show(this, ex.Message, "錯誤", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                StartButton.IsEnabled = true;
                ApplySheetFormatButton.IsEnabled = rows.Count > 0;
                CancelButton.IsEnabled = false;
                _applyCts?.Dispose();
                _applyCts = null;
            }
        }
    }
}
