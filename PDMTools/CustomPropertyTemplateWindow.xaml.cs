using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Data;
using Microsoft.Win32;
using PDMTools.Models;
using PDMTools.Services;

namespace PDMTools
{
    public partial class CustomPropertyTemplateWindow : Window
    {
        private readonly IReadOnlyList<BomItem> _bomItems;
        private readonly PdmBomExportService _exportService;
        private readonly VaultStyleReportGridFilterHelper _gridFilter;
        private CancellationTokenSource _queryCts;
        private CancellationTokenSource _applyCts;

        public CustomPropertyTemplateWindow(IReadOnlyList<BomItem> bomItems, PdmBomExportService exportService)
        {
            InitializeComponent();
            _bomItems = bomItems ?? Array.Empty<BomItem>();
            _exportService = exportService ?? throw new ArgumentNullException(nameof(exportService));

            _gridFilter = new VaultStyleReportGridFilterHelper(this, ResultDataGrid, ResultFilterCountTextBlock);
            _gridFilter.SetupCustomPropertyTemplateColumns();

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
                + $"共 {_bomItems.Count} 筆（{string.Join("、", extGroups)}）。\n"
                + "各欄標題右側漏斗可開啟篩選（全選、清除選取、搜尋、勾選值），與主視窗 Vault BOM 相同。\n"
                + "查詢完成後可勾選列、指定一般／焊件範本檔並「套用」：會同時寫入一般與焊件兩個索引（含原先未套用範本者），且僅允許勾選單一類型（不可混勾零件／組合件／工程圖）。";

            ApplyTemplateButton.IsEnabled = false;
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
            if (_bomItems.Count == 0)
            {
                MessageBox.Show(this,
                    "沒有可查詢的 SolidWorks 檔案。請確認 BOM 已抓取且篩選後有可見的 .sldprt / .sldasm / .slddrw。",
                    "提醒", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            _queryCts = new CancellationTokenSource();
            StartButton.IsEnabled = false;
            ApplyTemplateButton.IsEnabled = false;
            CancelButton.IsEnabled = true;
            _gridFilter.ClearItemsSource();
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

                var result = svc.Run(_bomItems, _exportService, progress, _queryCts.Token);

                _gridFilter.SetItemsSource(result.Rows);
                ApplyTemplateButton.IsEnabled = result.Rows.Count > 0;
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
                _queryCts?.Dispose();
                _queryCts = null;
            }
        }

        private void Cancel_OnClick(object sender, RoutedEventArgs e)
        {
            try { _queryCts?.Cancel(); } catch { }
            try { _applyCts?.Cancel(); } catch { }
        }

        private static List<CustomPropertyTemplateRow> MaterializeRows(IEnumerable source)
        {
            var list = new List<CustomPropertyTemplateRow>();
            if (source == null)
                return list;
            foreach (var o in source)
            {
                if (o is CustomPropertyTemplateRow r)
                    list.Add(r);
            }
            return list;
        }

        private string GetSingleCheckedDocExtensionOrNull()
        {
            var rows = MaterializeRows(ResultDataGrid.ItemsSource);
            var exts = rows
                .Where(r => r.IsSelectedForApply)
                .Select(r => (r.Extension ?? string.Empty).ToLowerInvariant())
                .Where(e => !string.IsNullOrEmpty(e))
                .Distinct()
                .ToList();
            return exts.Count == 1 ? exts[0] : null;
        }

        private void BrowseGeneralTemplate_OnClick(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog { CheckFileExists = true };
            var ext = GetSingleCheckedDocExtensionOrNull();
            switch (ext)
            {
                case ".sldprt":
                    dlg.Filter = "零件自訂屬性範本 (*.prtprp)|*.prtprp|所有檔案|*.*";
                    dlg.DefaultExt = ".prtprp";
                    break;
                case ".sldasm":
                    dlg.Filter = "組合件自訂屬性範本 (*.asmprp)|*.asmprp|所有檔案|*.*";
                    dlg.DefaultExt = ".asmprp";
                    break;
                case ".slddrw":
                    dlg.Filter = "工程圖自訂屬性範本 (*.drwprp)|*.drwprp|所有檔案|*.*";
                    dlg.DefaultExt = ".drwprp";
                    break;
                default:
                    dlg.Filter =
                        "SolidWorks 範本|*.prtprp;*.asmprp;*.drwprp|零件 (*.prtprp)|*.prtprp|組合件 (*.asmprp)|*.asmprp|工程圖 (*.drwprp)|*.drwprp|所有檔案|*.*";
                    break;
            }

            if (dlg.ShowDialog(this) == true)
                GeneralTemplatePathTextBox.Text = dlg.FileName;
        }

        private void BrowseWeldmentTemplate_OnClick(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog { CheckFileExists = true };
            var ext = GetSingleCheckedDocExtensionOrNull();
            switch (ext)
            {
                case ".sldprt":
                    dlg.Filter = "焊件範本 (*.wldprp)|*.wldprp|零件範本 (*.prtprp)|*.prtprp|所有檔案|*.*";
                    dlg.DefaultExt = ".wldprp";
                    break;
                case ".sldasm":
                    dlg.Filter = "組合件自訂屬性範本 (*.asmprp)|*.asmprp|所有檔案|*.*";
                    dlg.DefaultExt = ".asmprp";
                    break;
                case ".slddrw":
                    dlg.Filter = "工程圖自訂屬性範本 (*.drwprp)|*.drwprp|所有檔案|*.*";
                    dlg.DefaultExt = ".drwprp";
                    break;
                default:
                    dlg.Filter = "焊件 (*.wldprp)|*.wldprp|SolidWorks 範本|*.prtprp;*.asmprp;*.drwprp|所有檔案|*.*";
                    break;
            }

            if (dlg.ShowDialog(this) == true)
                WeldmentTemplatePathTextBox.Text = dlg.FileName;
        }

        private void SelectAllVisibleRows_OnClick(object sender, RoutedEventArgs e)
        {
            var view = CollectionViewSource.GetDefaultView(ResultDataGrid.ItemsSource);
            if (view == null)
                return;
            foreach (var o in view)
            {
                if (o is CustomPropertyTemplateRow r)
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
                if (o is CustomPropertyTemplateRow r)
                    r.IsSelectedForApply = false;
            }
        }

        private void ApplyTemplate_OnClick(object sender, RoutedEventArgs e)
        {
            var rows = MaterializeRows(ResultDataGrid.ItemsSource);
            if (rows.Count == 0)
            {
                MessageBox.Show(this, "請先執行查詢。", "提醒", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var general = GeneralTemplatePathTextBox.Text?.Trim() ?? string.Empty;
            if (string.IsNullOrEmpty(general) || !File.Exists(general))
            {
                MessageBox.Show(this, "請先指定存在的一般範本檔案。", "提醒", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            _applyCts = new CancellationTokenSource();
            ApplyTemplateButton.IsEnabled = false;
            StartButton.IsEnabled = false;
            CancelButton.IsEnabled = true;
            BatchProgressBar.Value = 0;
            ProgressText.Text = string.Empty;

            try
            {
                var svc = new CustomPropertyTemplateService();
                var progress = new Progress<CustomPropertyTemplateProgressInfo>(p =>
                {
                    if (p.Total > 0)
                        BatchProgressBar.Value = 100.0 * p.Current / p.Total;
                    ProgressText.Text = p.Message;
                });

                var applyResult = svc.ApplyTemplatesToRows(
                    rows,
                    general,
                    WeldmentTemplatePathTextBox.Text,
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
                    $"套用完成：成功 {applyResult.SuccessCount}，失敗 {applyResult.FailCount}。";
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
                ApplyTemplateButton.IsEnabled = rows.Count > 0;
                CancelButton.IsEnabled = false;
                _applyCts?.Dispose();
                _applyCts = null;
            }
        }
    }
}
