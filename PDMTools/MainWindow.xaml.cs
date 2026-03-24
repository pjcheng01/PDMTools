using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
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
        /// <summary>參考稽核預覽篩選：路徑是否落在 S 槽（不區分大小寫）。</summary>
        private const string ReferenceAuditSDriveFolderPrefix = @"S:\";
        private PdmBomExportService _exportService;
        private readonly CancellationTokenSource _cancellationTokenSource = new CancellationTokenSource();
        private readonly ObservableCollection<BomItem> _bomItems = new ObservableCollection<BomItem>();
        private readonly CardVariableLookupConverter _cardVariableConverter = new CardVariableLookupConverter();
        private readonly Dictionary<string, ColumnFilterState> _columnFilters =
            new Dictionary<string, ColumnFilterState>(StringComparer.OrdinalIgnoreCase);
        private ICollectionView _bomItemsView;
        private bool _suppressAssemblyPathTextChanged;
        private Popup _inlineTextSelectionPopup;
        private TextBox _inlineTextSelectionTextBox;
        private DataGridCell _inlineTextSelectionCell;

        private readonly ObservableCollection<ReferenceAuditRow> _auditRows = new ObservableCollection<ReferenceAuditRow>();
        private readonly SolidWorksReferenceAuditService _referenceAuditService = new SolidWorksReferenceAuditService();
        private ICollectionView _auditRowsView;
        private ReferenceAuditPreviewFilterMode _auditPreviewFilterMode = ReferenceAuditPreviewFilterMode.All;
        private MainWorkMode _currentWorkMode = MainWorkMode.Bom;

        private enum ReferenceAuditPreviewFilterMode
        {
            All = 0,
            InVaultOnly = 1,
            NotInVaultOnly = 2,
            FileMissingOnly = 3,
            UnderSDriveOnly = 4,
            NotUnderSDriveOnly = 5
        }

        private enum MainWorkMode
        {
            Bom = 0,
            ReferenceAudit = 1
        }

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
        // 使用者儲存的欄位組合清單
        private List<ColumnPresetProfile> _columnProfiles = new List<ColumnPresetProfile>();
        private string _lastUsedColumnProfileName = string.Empty;

        // ── 工程圖列相關狀態 ──────────────────────────────────────────────
        // 純零組件 BOM（不含工程圖列）；抓取後快取，用於 Toggle OFF 還原
        private List<BomItem> _rawBomItems = new List<BomItem>();
        // 含工程圖列的 BOM；Toggle ON 時由 AppendDrawingItemsAsync 建立並快取
        private List<BomItem> _bomItemsWithDrawings = null;

        public MainWindow()
        {
            InitializeComponent();
            SetProgressPercent(0);
            _bomItemsView = CollectionViewSource.GetDefaultView(_bomItems);
            _bomItemsView.Filter = FilterBomItem;
            _bomItems.CollectionChanged += BomItems_CollectionChanged;
            BomDataGrid.ItemsSource = _bomItemsView;
            _auditRowsView = CollectionViewSource.GetDefaultView(_auditRows);
            _auditRowsView.Filter = FilterReferenceAuditRow;
            ReferenceAuditDataGrid.ItemsSource = _auditRowsView;
            _auditRows.CollectionChanged += AuditRows_CollectionChanged;
            UpdateAuditPreviewCountLabel();
            ApplyMainModeUi(MainWorkMode.Bom);

            // 載入已存的欄位設定
            var settings = ColumnSettings.Load();
            _activeCardVarNames  = settings.SelectedVariables ?? new List<string>();
            _activeFixedColumns  = settings.HasFixedColumnsSetting
                ? settings.SelectedFixedColumns
                : null;   // null = 尚未設定，預設全顯示
            _columnProfiles = (settings.ColumnProfiles ?? new List<ColumnPresetProfile>())
                .Where(p => p != null && !string.IsNullOrWhiteSpace(p.Name))
                .Select(p => new ColumnPresetProfile
                {
                    Name = p.Name,
                    SelectedItems = (p.SelectedItems ?? new List<string>()).ToList()
                })
                .ToList();
            _lastUsedColumnProfileName = settings.LastUsedProfileName ?? string.Empty;

            var hasAppliedLastProfile = ApplyLastUsedColumnProfileIfAvailable();

            // 若尚未設定卡片變數，使用內建清單作為起始預設
            if (_activeCardVarNames.Count == 0 && !hasAppliedLastProfile)
            {
                _activeCardVarNames = PdmBomExportService.GetOrderedCardVariableLabels().ToList();
            }

            SetupBomDataGridColumns();

            ResetConfigurationComboNoFile();

            StatusTextBlock.Text = ColumnSettings.FileExists
                ? $"就緒（已載入欄位設定：{_activeCardVarNames.Count} 個）"
                : "就緒（使用內建欄位設定，可按「設定欄位...」自訂）";
            SetAssemblyPathPendingVisual(false);

            // 視窗完全載入後，延後執行背景偵測，避免影響啟動體感
            Loaded += MainWindow_OnLoaded;
        }
        private async void MainWindow_OnLoaded(object sender, RoutedEventArgs e)
        {
            // 先讓視窗完成第一輪繪製，避免啟動畫面卡頓
            await Dispatcher.InvokeAsync(() => { }, DispatcherPriority.ApplicationIdle);
            await Task.Delay(800);
            await CheckForNewVaultVariablesAsync();
        }


        private sealed class ColumnHeaderInfo : INotifyPropertyChanged
        {
            private bool _isFiltered;

            public string Key { get; set; } = string.Empty;
            public string Title { get; set; } = string.Empty;

            public bool IsFiltered
            {
                get => _isFiltered;
                set
                {
                    if (_isFiltered == value) return;
                    _isFiltered = value;
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsFiltered)));
                }
            }

            public event PropertyChangedEventHandler PropertyChanged;
        }

        // ── DataGrid 欄位建立 ──────────────────────────────────────────────

        /// <summary>依 _activeFixedColumns + _activeCardVarNames 建立 DataGrid 欄位。</summary>
        private void SetupBomDataGridColumns()
        {
            BomDataGrid.Columns.Clear();
            _columnFilters.Clear();

            void Add(
                string key,
                string header,
                Binding binding,
                Func<BomItem, string> valueGetter,
                double minWidth = 60,
                double maxWidth = double.PositiveInfinity)
            {
                var headerInfo = new ColumnHeaderInfo
                {
                    Key = key,
                    Title = header
                };

                var col = new DataGridTextColumn
                {
                    Header = headerInfo,
                    Binding  = binding,
                    MinWidth = minWidth
                };
                if (!double.IsPositiveInfinity(maxWidth))
                    col.MaxWidth = maxWidth;
                BomDataGrid.Columns.Add(col);

                _columnFilters[key] = new ColumnFilterState
                {
                    Key = key,
                    ValueGetter = valueGetter
                };
            }

            // Level 永遠顯示
            Add("Level", "Level", new Binding("Level") { Mode = BindingMode.OneWay }, i => i.Level, 50);

            // 用量統計：在父階底下，相同檔案路徑出現幾次
            Add("Use Count", "Use Count", new Binding("UsageCount") { Mode = BindingMode.OneWay }, i => i.UsageCount?.ToString() ?? string.Empty, 80);

            // 固定欄：依選擇決定是否加入（null = 全顯示）
            var fixedSet = _activeFixedColumns != null
                ? new HashSet<string>(_activeFixedColumns, StringComparer.OrdinalIgnoreCase)
                : null;   // null 代表全顯示

            bool ShowFixed(string name) => fixedSet == null || fixedSet.Contains(name);

            if (ShowFixed("File Name"))
                Add("File Name", "File Name", new Binding("FileName") { Mode = BindingMode.OneWay }, i => i.FileName, 90);
            if (ShowFixed("State"))
                Add("State", "State", new Binding("State") { Mode = BindingMode.OneWay }, i => i.State, 70);
            if (ShowFixed("Workflow State"))
                Add("Workflow State", "Workflow State", new Binding("WorkflowState") { Mode = BindingMode.OneWay }, i => i.WorkflowState, 90);
            if (ShowFixed("Description"))
                Add("Description", "Description", new Binding("Description") { Mode = BindingMode.OneWay }, i => i.Description, 80);
            if (ShowFixed("Part Number"))
                Add("Part Number", "Part Number", new Binding("PartNumber") { Mode = BindingMode.OneWay }, i => i.PartNumber, 80);
            if (ShowFixed("Referenced As"))
                Add("Referenced As", "Referenced As", new Binding("ReferencedAs") { Mode = BindingMode.OneWay }, i => i.ReferencedAs, 90);
            if (ShowFixed("Full Path"))
                Add("Full Path", "Full Path", new Binding("FullPath") { Mode = BindingMode.OneWay }, i => i.FullPath, 120, 520);
            if (ShowFixed("Description Var Used"))
                Add("Description Var Used", "Description Var Used", new Binding("DescriptionVarUsed") { Mode = BindingMode.OneWay }, i => i.DescriptionVarUsed, 80);
            if (ShowFixed("Description Config Used"))
                Add("Description Config Used", "Description Config Used", new Binding("DescriptionConfigUsed") { Mode = BindingMode.OneWay }, i => i.DescriptionConfigUsed, 80);
            if (ShowFixed("Part Number Var Used"))
                Add("Part Number Var Used", "Part Number Var Used", new Binding("PartNumberVarUsed") { Mode = BindingMode.OneWay }, i => i.PartNumberVarUsed, 80);
            if (ShowFixed("Part Number Config Used"))
                Add("Part Number Config Used", "Part Number Config Used", new Binding("PartNumberConfigUsed") { Mode = BindingMode.OneWay }, i => i.PartNumberConfigUsed, 80);

            // 資料卡變數欄
            foreach (var varName in _activeCardVarNames)
            {
                var binding = new Binding(".")
                {
                    Mode               = BindingMode.OneWay,
                    Converter          = _cardVariableConverter,
                    ConverterParameter = varName
                };
                var cardVarKey = "Card:" + varName;
                Add(cardVarKey, cardVarKey, binding, i =>
                {
                    if (i.CardVariables == null) return string.Empty;
                    return i.CardVariables.TryGetValue(varName, out var value) ? value : string.Empty;
                }, 72);
            }

            RebuildAllFilterValueOptions();
            UpdateFilterUiState();
            _bomItemsView?.Refresh();
        }

        private void AutoSizeDataGridColumns()
        {
            BomDataGrid.UpdateLayout();
            var cellWidths = new Dictionary<DataGridColumn, double>();
            foreach (var col in BomDataGrid.Columns)
            {
                col.Width = new DataGridLength(1, DataGridLengthUnitType.SizeToCells);
            }

            BomDataGrid.UpdateLayout();
            foreach (var col in BomDataGrid.Columns)
            {
                cellWidths[col] = col.ActualWidth;
                col.Width = new DataGridLength(1, DataGridLengthUnitType.SizeToHeader);
            }

            BomDataGrid.UpdateLayout();
            foreach (var col in BomDataGrid.Columns)
            {
                var finalWidth = Math.Max(cellWidths.TryGetValue(col, out var cellWidth) ? cellWidth : 0, col.ActualWidth);
                col.Width = new DataGridLength(finalWidth + 4);
            }
        }

        // ── 按鈕事件 ──────────────────────────────────────────────────────

        private void ColumnFilterButton_OnClick(object sender, RoutedEventArgs e)
        {
            if (!(sender is Button button) || !(button.Tag is string key) || string.IsNullOrWhiteSpace(key))
                return;
            if (!_columnFilters.TryGetValue(key, out var state))
                return;

            EnsureSelectionInitialized(state);
            RebuildFilterValueOptions(state);

            var menu = BuildFilterContextMenu(state);
            button.ContextMenu = menu;
            menu.PlacementTarget = button;
            menu.Placement = PlacementMode.Bottom;
            menu.IsOpen = true;
            e.Handled = true;
        }

        private void ClearFiltersButton_OnClick(object sender, RoutedEventArgs e)
        {
            foreach (var state in _columnFilters.Values)
            {
                state.SelectedValues.Clear();
            }

            _bomItemsView?.Refresh();
            UpdateFilterUiState();
            StatusTextBlock.Text = _bomItems.Count > 0
                ? $"已清除篩選，顯示 {_bomItems.Count} / {_bomItems.Count} 筆。"
                : "已清除篩選。";
        }

        private void BomDataGrid_OnPreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            var original = e.OriginalSource as DependencyObject;
            var cell = FindVisualParent<DataGridCell>(original);
            if (cell == null)
                return;

            if (!cell.IsFocused)
                cell.Focus();

            var row = FindVisualParent<DataGridRow>(cell);
            if (row?.Item != null)
            {
                BomDataGrid.SelectedItem = row.Item; // 保留列選取行為
            }

            BomDataGrid.CurrentCell = new DataGridCellInfo(cell.DataContext, cell.Column);
        }

        private void CopyCurrentCellMenuItem_OnClick(object sender, RoutedEventArgs e)
        {
            if (!TryGetCurrentCellText(out var cellText, out _))
            {
                MessageBox.Show(this, "請先在表格中選取一個儲存格。", "提醒", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            Clipboard.SetText(cellText ?? string.Empty);
            StatusTextBlock.Text = "已複製儲存格內容。";
        }

        private void OpenTextSelectionModeMenuItem_OnClick(object sender, RoutedEventArgs e)
        {
            if (!TryGetCurrentCellText(out var cellText, out var cell))
            {
                MessageBox.Show(this, "請先在表格中選取一個儲存格。", "提醒", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            OpenTextSelectionMode(cellText ?? string.Empty, cell);
        }

        private async void BrowseButton_OnClick(object sender, RoutedEventArgs e)
        {
            if (_currentWorkMode == MainWorkMode.ReferenceAudit)
            {
                var auditDialog = new OpenFileDialog
                {
                    Title = "選擇要稽核的組合件（可為 Vault 外路徑）",
                    Filter = "SolidWorks Assembly (*.sldasm)|*.sldasm",
                    CheckFileExists = true,
                    Multiselect = false,
                    InitialDirectory = Directory.Exists(VaultRootPath)
                        ? VaultRootPath
                        : Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
                };

                if (auditDialog.ShowDialog(this) == true)
                {
                    SetAssemblyPathText(auditDialog.FileName, isImported: true);
                    StatusTextBlock.Text = "已選擇稽核組合件。";
                }

                return;
            }

            var candidatePath = (AssemblyPathTextBox.Text ?? string.Empty).Trim();
            if (!string.IsNullOrWhiteSpace(candidatePath))
            {
                if (!TryValidateImportPath(candidatePath, out var normalizedPath, out var errorMessage))
                {
                    MessageBox.Show(this, errorMessage, "路徑無效", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                SetAssemblyPathText(normalizedPath, isImported: true);
                StatusTextBlock.Text = "已導入組合件。";
                _bomItems.Clear();
                ClearAllFiltersSilently();
                await PopulateConfigurationComboForPathAsync(normalizedPath);
                return;
            }

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

            SetAssemblyPathText(dialog.FileName, isImported: true);
            StatusTextBlock.Text = "已選擇組合件。";
            _bomItems.Clear();
            ClearAllFiltersSilently();

            await PopulateConfigurationComboForPathAsync(dialog.FileName);
        }

        private async void RunReferenceAuditButton_OnClick(object sender, RoutedEventArgs e)
        {
            var raw = AssemblyPathTextBox.Text?.Trim().Trim('"') ?? string.Empty;
            if (string.IsNullOrWhiteSpace(raw))
            {
                MessageBox.Show(this, "請指定稽核用之 .sldasm 路徑。", "提醒", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string path;
            try
            {
                path = Path.GetFullPath(raw);
            }
            catch
            {
                MessageBox.Show(this, "路徑格式不正確。", "提醒", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!File.Exists(path))
            {
                MessageBox.Show(this, "找不到檔案。", "提醒", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!string.Equals(Path.GetExtension(path), ".sldasm", StringComparison.OrdinalIgnoreCase))
            {
                MessageBox.Show(this, "僅支援 .sldasm。", "提醒", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            SetUiBusy(true);
            StatusTextBlock.Text = "正在以 SolidWorks 稽核引用…";
            _auditRows.Clear();
            ExportReferenceAuditButton.IsEnabled = false;
            UpdateAuditPreviewCountLabel();

            try
            {
                IReadOnlyList<ReferenceAuditRow> rows = null;
                await Dispatcher.InvokeAsync(() =>
                {
                    rows = _referenceAuditService.AuditAssembly(path, VaultRootPath);
                }, DispatcherPriority.Normal);

                rows ??= Array.Empty<ReferenceAuditRow>();

                foreach (var r in rows)
                    _auditRows.Add(r);

                var notInVault = rows.Count(r =>
                    !r.IsUnderVaultRoot && !string.IsNullOrWhiteSpace(r.FullPath));
                StatusTextBlock.Text =
                    $"稽核完成，共 {rows.Count} 筆引用（其中 {notInVault} 筆路徑不在 Vault 根目錄 {VaultRootPath} 下）。";
                ExportReferenceAuditButton.IsEnabled = rows.Count > 0;
                _auditRowsView?.Refresh();
                UpdateAuditPreviewCountLabel();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "稽核失敗", MessageBoxButton.OK, MessageBoxImage.Error);
                StatusTextBlock.Text = "稽核失敗。";
            }
            finally
            {
                SetUiBusy(false);
            }
        }

        private void MainModeRadioButton_OnChecked(object sender, RoutedEventArgs e)
        {
            var mode = ReferenceAuditModeRadioButton.IsChecked == true
                ? MainWorkMode.ReferenceAudit
                : MainWorkMode.Bom;
            ApplyMainModeUi(mode);
        }

        private void ApplyMainModeUi(MainWorkMode mode)
        {
            _currentWorkMode = mode;
            var isBomMode = mode == MainWorkMode.Bom;

            if (BomOperationsPanel != null)
                BomOperationsPanel.Visibility = isBomMode ? Visibility.Visible : Visibility.Collapsed;
            if (BomModeContentPanel != null)
                BomModeContentPanel.Visibility = isBomMode ? Visibility.Visible : Visibility.Collapsed;
            if (ReferenceAuditModeContentPanel != null)
                ReferenceAuditModeContentPanel.Visibility = isBomMode ? Visibility.Collapsed : Visibility.Visible;

            if (BrowseButton != null)
                BrowseButton.Content = isBomMode ? "導入文件" : "選擇檔案";

            if (StatusTextBlock != null)
                StatusTextBlock.Text = isBomMode
                    ? (_bomItems.Count > 0 ? $"目前為 Vault BOM 模式（{_bomItems.Count} 筆）。" : "目前為 Vault BOM 模式。")
                    : (_auditRows.Count > 0 ? $"目前為參考稽核模式（{_auditRows.Count} 筆）。" : "目前為參考稽核模式。");
        }

        private void ReferenceAuditPreviewFilterComboBox_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is not ComboBox cb)
                return;
            var idx = cb.SelectedIndex;
            if (idx < 0)
                idx = 0;
            var maxIdx = (int)ReferenceAuditPreviewFilterMode.NotUnderSDriveOnly;
            _auditPreviewFilterMode = (ReferenceAuditPreviewFilterMode)Math.Min(idx, maxIdx);
            _auditRowsView?.Refresh();
            UpdateAuditPreviewCountLabel();
        }

        private void AuditRows_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            Dispatcher.BeginInvoke(DispatcherPriority.Background, new Action(UpdateAuditPreviewCountLabel));
        }

        private bool FilterReferenceAuditRow(object item)
        {
            if (item is not ReferenceAuditRow row)
                return false;
            return _auditPreviewFilterMode switch
            {
                ReferenceAuditPreviewFilterMode.InVaultOnly => row.IsUnderVaultRoot,
                ReferenceAuditPreviewFilterMode.NotInVaultOnly => !row.IsUnderVaultRoot,
                ReferenceAuditPreviewFilterMode.FileMissingOnly => !row.FileExists,
                ReferenceAuditPreviewFilterMode.UnderSDriveOnly => IsFullPathUnderPrefix(row.FullPath, ReferenceAuditSDriveFolderPrefix),
                ReferenceAuditPreviewFilterMode.NotUnderSDriveOnly => !IsFullPathUnderPrefix(row.FullPath, ReferenceAuditSDriveFolderPrefix),
                _ => true
            };
        }

        /// <summary>完整路徑是否在指定根目錄下（不區分大小寫）；無路徑視為不在其下。</summary>
        private static bool IsFullPathUnderPrefix(string fullPath, string rootPrefix)
        {
            if (string.IsNullOrWhiteSpace(fullPath) || string.IsNullOrWhiteSpace(rootPrefix))
                return false;
            var p = fullPath.Trim();
            var root = rootPrefix.TrimEnd('\\', '/') + "\\";
            return p.StartsWith(root, StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(p.TrimEnd('\\', '/'), rootPrefix.TrimEnd('\\', '/'), StringComparison.OrdinalIgnoreCase);
        }

        private void UpdateAuditPreviewCountLabel()
        {
            if (ReferenceAuditPreviewCountTextBlock == null)
                return;
            var total = _auditRows.Count;
            if (total == 0)
            {
                ReferenceAuditPreviewCountTextBlock.Text = "預覽：—";
                return;
            }

            if (_auditRowsView == null)
            {
                ReferenceAuditPreviewCountTextBlock.Text = $"預覽：{total} 筆";
                return;
            }

            var shown = _auditRowsView.Cast<object>().Count();
            ReferenceAuditPreviewCountTextBlock.Text =
                shown == total
                    ? $"預覽：共 {total} 筆"
                    : $"預覽：顯示 {shown} / 共 {total} 筆";
        }

        private void ExportReferenceAuditButton_OnClick(object sender, RoutedEventArgs e)
        {
            if (_auditRows.Count == 0)
            {
                MessageBox.Show(this, "請先執行稽核。", "提醒", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var toExport = GetFilteredReferenceAuditRowsForExport();
            if (toExport.Count == 0)
            {
                MessageBox.Show(this,
                    "目前預覽篩選下沒有任何列可匯出。請調整篩選條件，或改選「全部顯示」後再匯出。",
                    "提醒",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            var saveDialog = new SaveFileDialog
            {
                Title = "儲存參考稽核報表（與目前預覽篩選一致）",
                Filter = "Excel Workbook (*.xlsx)|*.xlsx",
                FileName = $"ReferenceAudit_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx",
                AddExtension = true,
                OverwritePrompt = true
            };

            if (saveDialog.ShowDialog(this) != true)
                return;

            try
            {
                ReferenceAuditExcelExporter.Export(toExport, saveDialog.FileName);
                MessageBox.Show(this,
                    $"已匯出 {toExport.Count} 筆（與目前預覽篩選一致）：\n{saveDialog.FileName}",
                    "完成",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "匯出失敗", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>取得目前預覽篩選後的稽核列（與 DataGrid 可見列一致）。</summary>
        private List<ReferenceAuditRow> GetFilteredReferenceAuditRowsForExport()
        {
            if (_auditRowsView == null)
                return _auditRows.ToList();
            return _auditRowsView.Cast<ReferenceAuditRow>().ToList();
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
                    "請先以「導入文件」選擇或貼上組合件路徑，並等待組態清單載入完成後，再從「組態」下拉選單選取組態。",
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
                RebuildAllFilterValueOptions();
                _bomItemsView?.Refresh();
                UpdateFilterUiState();
                StatusTextBlock.Text = $"已隱藏工程圖列，共 {GetFilteredCount()} / {_bomItems.Count} 筆。";
                return;
            }

            // Toggle ON：若已有快取直接使用，否則背景搜尋
            if (_bomItemsWithDrawings != null)
            {
                _bomItems.Clear();
                foreach (var item in _bomItemsWithDrawings) _bomItems.Add(item);
                var drwCount = _bomItemsWithDrawings.Count(x => x.IsDrawing);
                RebuildAllFilterValueOptions();
                _bomItemsView?.Refresh();
                UpdateFilterUiState();
                StatusTextBlock.Text = $"已顯示工程圖列（找到 {drwCount} 個工程圖，顯示 {GetFilteredCount()} / {_bomItems.Count} 筆）。";
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
                RebuildAllFilterValueOptions();
                _bomItemsView?.Refresh();
                UpdateFilterUiState();

                var drwCount = _bomItemsWithDrawings.Count(x => x.IsDrawing);
                StatusTextBlock.Text = drwCount > 0
                    ? $"已顯示工程圖列（找到 {drwCount} 個工程圖，顯示 {GetFilteredCount()} / {_bomItems.Count} 筆）。"
                    : $"未找到任何工程圖，顯示 {GetFilteredCount()} / {_bomItems.Count} 筆。";
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

            var win = new ColumnSettingsWindow(allItems, currentSelected, _columnProfiles, _lastUsedColumnProfileName) { Owner = this };

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
            _columnProfiles = (win.UpdatedProfiles ?? new List<ColumnPresetProfile>())
                .Where(p => p != null && !string.IsNullOrWhiteSpace(p.Name))
                .Select(p => new ColumnPresetProfile
                {
                    Name = p.Name,
                    SelectedItems = (p.SelectedItems ?? new List<string>()).ToList()
                })
                .ToList();
            _lastUsedColumnProfileName = win.LastUsedProfileName ?? string.Empty;

            var colSettings = new ColumnSettings
            {
                SelectedVariables    = _activeCardVarNames,
                SelectedFixedColumns = _activeFixedColumns,
                // KnownVariables 直接用 Vault 完整列舉清單，確保下次啟動比對時不會誤報
                // （不依賴 AllShownAvailableItems，避免固定欄名稱與卡片變數名稱重疊時漏記）
                KnownVariables       = allCardVars.ToList(),
                ColumnProfiles       = _columnProfiles,
                LastUsedProfileName  = _lastUsedColumnProfileName
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
                var enumTask = _exportService.EnumerateVaultVariablesAsync(
                    AssemblyPathTextBox.Text?.Trim());
                var completed = await Task.WhenAny(enumTask, Task.Delay(5000));
                if (completed != enumTask)
                {
                    return; // 啟動時若 Vault 回應過慢，直接略過本次偵測
                }
                vaultVars = await enumTask;
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
                ClearAllFiltersSilently();

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

                RebuildAllFilterValueOptions();
                _bomItemsView?.Refresh();
                UpdateFilterUiState();

                var cfg = _exportService.LastConfigurationResolved;
                var layout = _exportService.LastBomLayoutNameUsed;
                StatusTextBlock.Text =
                    string.IsNullOrWhiteSpace(cfg) && string.IsNullOrWhiteSpace(layout)
                        ? $"抓取完成，顯示 {GetFilteredCount()} / {_bomItems.Count} 筆。"
                        : $"抓取完成，顯示 {GetFilteredCount()} / {_bomItems.Count} 筆。（組態：{cfg}；BOM 版面：{layout}）";
                SetProgressPercent(0);

                await Dispatcher.InvokeAsync(new Action(AutoSizeDataGridColumns), DispatcherPriority.Loaded);
                await Dispatcher.InvokeAsync(new Action(AutoSizeDataGridColumns), DispatcherPriority.ContextIdle);
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
            var isBomMode = _currentWorkMode == MainWorkMode.Bom;
            StartGrabButton.IsEnabled      = !isBusy && isBomMode;
            ExportButton.IsEnabled         = !isBusy && isBomMode;
            ColumnSettingsButton.IsEnabled = !isBusy && isBomMode;
            ClearFiltersButton.IsEnabled   = !isBusy && isBomMode && HasActiveFilters();
            // Toggle 只在有資料時才可操作，忙碌中一律禁用
            ShowDrawingsToggle.IsEnabled   = !isBusy && isBomMode && _rawBomItems.Count > 0;
            RunReferenceAuditButton.IsEnabled = !isBusy && !isBomMode;
            ExportReferenceAuditButton.IsEnabled = !isBusy && !isBomMode && _auditRows.Count > 0;
            ReferenceAuditPreviewFilterComboBox.IsEnabled = !isBusy && !isBomMode;
            BomModeRadioButton.IsEnabled = !isBusy;
            ReferenceAuditModeRadioButton.IsEnabled = !isBusy;
            Mouse.OverrideCursor           = isBusy ? Cursors.Wait : null;
        }

        /// <summary>同步更新進度條與右側固定位置之百分比文字。</summary>
        private void SetProgressPercent(double value)
        {
            var v = Math.Max(0, Math.Min(100, value));
            ProgressBar.Value = v;
            ProgressPercentTextBlock.Text = $"{(int)Math.Round(v, MidpointRounding.AwayFromZero)}%";
        }

        private void AssemblyPathTextBox_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            if (_suppressAssemblyPathTextChanged)
                return;

            SetAssemblyPathPendingVisual(!string.IsNullOrWhiteSpace(AssemblyPathTextBox.Text));
        }

        private void SetAssemblyPathText(string path, bool isImported)
        {
            _suppressAssemblyPathTextChanged = true;
            try
            {
                AssemblyPathTextBox.Text = path ?? string.Empty;
            }
            finally
            {
                _suppressAssemblyPathTextChanged = false;
            }

            SetAssemblyPathPendingVisual(!isImported && !string.IsNullOrWhiteSpace(path));
        }

        private void SetAssemblyPathPendingVisual(bool isPending)
        {
            AssemblyPathTextBox.Foreground = isPending ? System.Windows.Media.Brushes.Gray : System.Windows.Media.Brushes.Black;
        }

        private static bool TryValidateImportPath(string inputPath, out string normalizedPath, out string errorMessage)
        {
            normalizedPath = string.Empty;
            errorMessage = string.Empty;

            if (string.IsNullOrWhiteSpace(inputPath))
            {
                errorMessage = "請先輸入或貼上組合件路徑。";
                return false;
            }

            var trimmed = inputPath.Trim().Trim('"');
            try
            {
                normalizedPath = Path.GetFullPath(trimmed);
            }
            catch
            {
                errorMessage = "路徑格式不正確，請確認後再試。";
                return false;
            }

            if (!string.Equals(Path.GetExtension(normalizedPath), ".sldasm", StringComparison.OrdinalIgnoreCase))
            {
                errorMessage = "僅支援導入 .sldasm 組合件檔案。";
                return false;
            }

            if (!File.Exists(normalizedPath))
            {
                errorMessage = "找不到指定檔案，請確認路徑是否正確。";
                return false;
            }

            var fullVaultRoot = Path.GetFullPath(VaultRootPath)
                .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            var rootWithSep = fullVaultRoot + Path.DirectorySeparatorChar;
            if (!normalizedPath.StartsWith(rootWithSep, StringComparison.OrdinalIgnoreCase))
            {
                errorMessage = $"檔案必須位於 Vault 路徑內：{VaultRootPath}";
                return false;
            }

            return true;
        }

        private bool ApplyLastUsedColumnProfileIfAvailable()
        {
            if (string.IsNullOrWhiteSpace(_lastUsedColumnProfileName) || _columnProfiles == null || _columnProfiles.Count == 0)
                return false;

            var profile = _columnProfiles.FirstOrDefault(p =>
                string.Equals(p.Name, _lastUsedColumnProfileName, StringComparison.OrdinalIgnoreCase));
            if (profile == null || profile.SelectedItems == null || profile.SelectedItems.Count == 0)
                return false;

            var fixedSet = new HashSet<string>(AllFixedColumnNames, StringComparer.OrdinalIgnoreCase);
            var selectedItems = profile.SelectedItems
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .ToList();

            var selectedFixed = selectedItems.Where(x => fixedSet.Contains(x)).ToList();
            var selectedCard = selectedItems.Where(x => !fixedSet.Contains(x)).ToList();
            if (selectedFixed.Count == 0 && selectedCard.Count == 0)
                return false;

            _activeFixedColumns = selectedFixed;
            _activeCardVarNames = selectedCard;
            return true;
        }

        private void BomItems_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            RebuildAllFilterValueOptions();
            _bomItemsView?.Refresh();
            UpdateFilterUiState();
        }

        private bool FilterBomItem(object obj)
        {
            if (!(obj is BomItem item))
                return false;

            foreach (var state in _columnFilters.Values)
            {
                if (state.AvailableValues.Count == 0)
                    continue;

                var selectedCount = state.SelectedValues.Count;
                if (selectedCount == 0 || selectedCount == state.AvailableValues.Count)
                    continue;

                var currentValue = NormalizeFilterValue(state.ValueGetter(item));
                if (!state.SelectedValues.Contains(currentValue))
                    return false;
            }

            return true;
        }

        private bool TryGetCurrentCellText(out string cellText, out DataGridCell cell)
        {
            cellText = string.Empty;
            cell = null;
            var currentCell = BomDataGrid.CurrentCell;
            if (currentCell.Item == null || currentCell.Column == null)
                return false;
            if (!(currentCell.Item is BomItem bomItem))
                return false;
            if (!(currentCell.Column.Header is ColumnHeaderInfo header))
                return false;
            if (!_columnFilters.TryGetValue(header.Key, out var filterState))
                return false;

            cellText = filterState.ValueGetter(bomItem) ?? string.Empty;
            cell = GetCellFromCurrentCellInfo(currentCell);
            return true;
        }

        private void OpenTextSelectionMode(string fullText, DataGridCell targetCell)
        {
            if (targetCell == null)
            {
                MessageBox.Show(this, "目前儲存格不可用，請再點選一次後重試。", "提醒", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            CloseInlineTextSelectionMode();

            _inlineTextSelectionTextBox = new TextBox
            {
                Text = fullText ?? string.Empty,
                IsReadOnly = true,
                AcceptsReturn = false,
                BorderThickness = new Thickness(0),
                Background = Brushes.White,
                VerticalContentAlignment = VerticalAlignment.Center,
                Padding = new Thickness(3, 1, 3, 1),
                MinWidth = Math.Max(120, targetCell.ActualWidth - 6),
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto
            };

            var hostBorder = new Border
            {
                Background = Brushes.White,
                BorderBrush = new SolidColorBrush(Color.FromRgb(33, 150, 243)),
                BorderThickness = new Thickness(1),
                Child = _inlineTextSelectionTextBox
            };

            _inlineTextSelectionPopup = new Popup
            {
                PlacementTarget = targetCell,
                Placement = PlacementMode.Center,
                StaysOpen = false,
                AllowsTransparency = true,
                Child = hostBorder
            };
            _inlineTextSelectionCell = targetCell;

            _inlineTextSelectionTextBox.PreviewKeyDown += (_, e) =>
            {
                if (e.Key == Key.Escape)
                {
                    CloseInlineTextSelectionMode();
                    e.Handled = true;
                }
            };
            _inlineTextSelectionPopup.Closed += (_, __) =>
            {
                CloseInlineTextSelectionMode();
            };

            _inlineTextSelectionPopup.IsOpen = true;
            _inlineTextSelectionTextBox.Focus();
            _inlineTextSelectionTextBox.SelectAll();
        }

        private void CloseInlineTextSelectionMode()
        {
            if (_inlineTextSelectionPopup != null)
            {
                var popup = _inlineTextSelectionPopup;
                _inlineTextSelectionPopup = null;
                popup.IsOpen = false;
            }

            _inlineTextSelectionTextBox = null;

            if (_inlineTextSelectionCell != null)
            {
                _inlineTextSelectionCell.Focus();
                _inlineTextSelectionCell = null;
            }
        }

        private DataGridCell GetCellFromCurrentCellInfo(DataGridCellInfo currentCell)
        {
            if (currentCell.Item == null || currentCell.Column == null)
                return null;

            var rowContainer = BomDataGrid.ItemContainerGenerator.ContainerFromItem(currentCell.Item) as DataGridRow;
            if (rowContainer == null)
            {
                BomDataGrid.ScrollIntoView(currentCell.Item);
                BomDataGrid.UpdateLayout();
                rowContainer = BomDataGrid.ItemContainerGenerator.ContainerFromItem(currentCell.Item) as DataGridRow;
                if (rowContainer == null)
                    return null;
            }

            var presenter = FindVisualChild<DataGridCellsPresenter>(rowContainer);
            if (presenter == null)
            {
                rowContainer.ApplyTemplate();
                presenter = FindVisualChild<DataGridCellsPresenter>(rowContainer);
            }
            if (presenter == null)
                return null;

            var columnIndex = currentCell.Column.DisplayIndex;
            var cell = presenter.ItemContainerGenerator.ContainerFromIndex(columnIndex) as DataGridCell;
            if (cell == null)
            {
                BomDataGrid.ScrollIntoView(currentCell.Item, currentCell.Column);
                BomDataGrid.UpdateLayout();
                cell = presenter.ItemContainerGenerator.ContainerFromIndex(columnIndex) as DataGridCell;
            }

            return cell;
        }

        private static T FindVisualParent<T>(DependencyObject child) where T : DependencyObject
        {
            var current = child;
            while (current != null)
            {
                if (current is T hit)
                    return hit;
                current = VisualTreeHelper.GetParent(current);
            }

            return null;
        }

        private static T FindVisualChild<T>(DependencyObject parent) where T : DependencyObject
        {
            if (parent == null)
                return null;

            var childCount = VisualTreeHelper.GetChildrenCount(parent);
            for (var i = 0; i < childCount; i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child is T hit)
                    return hit;

                var descendant = FindVisualChild<T>(child);
                if (descendant != null)
                    return descendant;
            }

            return null;
        }

        private ContextMenu BuildFilterContextMenu(ColumnFilterState state)
        {
            var menu = new ContextMenu();
            var valueMenuItems = new List<(string Value, MenuItem Item)>();
            TextBox searchBox = null;

            void RefocusSearchBox()
            {
                if (searchBox == null) return;
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    searchBox.Focus();
                    searchBox.Select(searchBox.Text.Length, 0);
                }), DispatcherPriority.Background);
            }

            var selectAll = new MenuItem
            {
                Header = "全選",
                StaysOpenOnClick = true
            };
            selectAll.Click += (_, __) =>
            {
                state.SelectedValues = new HashSet<string>(state.AvailableValues, StringComparer.Ordinal);
                foreach (var (_, item) in valueMenuItems)
                {
                    item.IsChecked = true;
                }
                ApplyFiltersAndRefreshStatus();
                RefocusSearchBox();
            };
            menu.Items.Add(selectAll);

            var clearSelection = new MenuItem
            {
                Header = "清除選取",
                StaysOpenOnClick = true
            };
            clearSelection.Click += (_, __) =>
            {
                state.SelectedValues.Clear();
                state.SearchText = string.Empty;
                if (searchBox != null)
                {
                    searchBox.Text = string.Empty;
                }
                foreach (var (_, item) in valueMenuItems)
                {
                    item.IsChecked = false;
                }
                ApplyFiltersAndRefreshStatus();
                RefocusSearchBox();
            };
            menu.Items.Add(clearSelection);

            menu.Items.Add(new Separator());

            searchBox = new TextBox
            {
                MinWidth = 180,
                Margin = new Thickness(2),
                Text = state.SearchText ?? string.Empty,
                ToolTip = "支援萬用字元：*、?"
            };

            var searchHost = new MenuItem
            {
                StaysOpenOnClick = true,
                Focusable = false,
                Header = searchBox
            };
            menu.Items.Add(searchHost);
            menu.Items.Add(new Separator());

            foreach (var value in state.AvailableValues)
            {
                var item = new MenuItem
                {
                    Header = value,
                    IsCheckable = true,
                    IsChecked = state.SelectedValues.Contains(value),
                    StaysOpenOnClick = true
                };

                item.Click += (_, __) =>
                {
                    if (item.IsChecked)
                        state.SelectedValues.Add(value);
                    else
                        state.SelectedValues.Remove(value);

                    ApplyFiltersAndRefreshStatus();
                };
                menu.Items.Add(item);
                valueMenuItems.Add((value, item));
            }

            void ApplySearchFilter(bool syncSelectionToSearch)
            {
                var keyword = (searchBox.Text ?? string.Empty).Trim();
                state.SearchText = keyword;
                var matchedValues = new List<string>();

                foreach (var (value, item) in valueMenuItems)
                {
                    var isMatch = IsSearchMatch(value, keyword);
                    item.Visibility = isMatch
                        ? Visibility.Visible
                        : Visibility.Collapsed;
                    if (isMatch)
                    {
                        matchedValues.Add(value);
                    }
                }

                if (syncSelectionToSearch && !string.IsNullOrEmpty(keyword))
                {
                    state.SelectedValues = new HashSet<string>(matchedValues, StringComparer.Ordinal);
                    foreach (var (value, item) in valueMenuItems)
                    {
                        item.IsChecked = state.SelectedValues.Contains(value);
                    }
                    ApplyFiltersAndRefreshStatus();
                }
            }

            searchBox.TextChanged += (_, __) => ApplySearchFilter(syncSelectionToSearch: true);
            menu.Opened += (_, __) =>
            {
                searchBox.Focus();
                searchBox.Select(searchBox.Text.Length, 0);
                ApplySearchFilter(syncSelectionToSearch: false);
            };

            return menu;
        }

        private void ApplyFiltersAndRefreshStatus()
        {
            _bomItemsView?.Refresh();
            UpdateFilterUiState();
            StatusTextBlock.Text = _bomItems.Count > 0
                ? $"已套用篩選，顯示 {GetFilteredCount()} / {_bomItems.Count} 筆。"
                : "已套用篩選。";
        }

        private void ClearAllFiltersSilently()
        {
            foreach (var state in _columnFilters.Values)
            {
                state.SelectedValues.Clear();
            }
            _bomItemsView?.Refresh();
            UpdateFilterUiState();
        }

        private static string NormalizeFilterValue(string raw)
        {
            return string.IsNullOrWhiteSpace(raw) ? ColumnFilterState.BlankDisplayText : raw.Trim();
        }

        private static bool IsSearchMatch(string value, string keyword)
        {
            if (string.IsNullOrWhiteSpace(keyword))
                return true;

            var pattern = keyword.Trim();
            var hasWildcard = pattern.IndexOf('*') >= 0 || pattern.IndexOf('?') >= 0;
            if (!hasWildcard)
                return value.IndexOf(pattern, StringComparison.OrdinalIgnoreCase) >= 0;

            var regexPattern = "^" + Regex.Escape(pattern)
                .Replace(@"\*", ".*")
                .Replace(@"\?", ".") + "$";

            return Regex.IsMatch(value ?? string.Empty, regexPattern, RegexOptions.IgnoreCase);
        }

        private void EnsureSelectionInitialized(ColumnFilterState state)
        {
            if (state.SelectedValues.Count > 0) return;
            state.SelectedValues = new HashSet<string>(state.AvailableValues, StringComparer.Ordinal);
        }

        private void RebuildAllFilterValueOptions()
        {
            foreach (var state in _columnFilters.Values)
            {
                RebuildFilterValueOptions(state);
            }
        }

        private void RebuildFilterValueOptions(ColumnFilterState state)
        {
            var oldValues = state.AvailableValues ?? new List<string>();
            var wasUnfiltered =
                oldValues.Count == 0 ||
                state.SelectedValues.Count == 0 ||
                state.SelectedValues.Count == oldValues.Count;

            var values = _bomItems
                .Select(x => NormalizeFilterValue(state.ValueGetter(x)))
                .Distinct(StringComparer.Ordinal)
                .ToList();

            var hasBlank = values.Any(x => string.Equals(x, ColumnFilterState.BlankDisplayText, StringComparison.Ordinal));
            var orderedValues = values
                .Where(x => !string.Equals(x, ColumnFilterState.BlankDisplayText, StringComparison.Ordinal))
                .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (hasBlank)
            {
                orderedValues.Insert(0, ColumnFilterState.BlankDisplayText);
            }

            values = orderedValues;
            state.AvailableValues = values;

            if (wasUnfiltered)
            {
                state.SelectedValues = new HashSet<string>(values, StringComparer.Ordinal);
                return;
            }

            state.SelectedValues.IntersectWith(values);
            if (state.SelectedValues.Count == 0)
            {
                state.SelectedValues = new HashSet<string>(values, StringComparer.Ordinal);
            }
        }

        private void UpdateFilterUiState()
        {
            foreach (var col in BomDataGrid.Columns)
            {
                if (!(col.Header is ColumnHeaderInfo header)) continue;
                if (!_columnFilters.TryGetValue(header.Key, out var state))
                {
                    header.IsFiltered = false;
                    continue;
                }

                header.IsFiltered =
                    state.AvailableValues.Count > 0 &&
                    state.SelectedValues.Count > 0 &&
                    state.SelectedValues.Count < state.AvailableValues.Count;
            }

            ClearFiltersButton.IsEnabled = HasActiveFilters() && !IsUiBusy();
        }

        private bool HasActiveFilters()
        {
            return _columnFilters.Values.Any(state =>
                state.AvailableValues.Count > 0 &&
                state.SelectedValues.Count > 0 &&
                state.SelectedValues.Count < state.AvailableValues.Count);
        }

        private int GetFilteredCount()
        {
            return _bomItemsView?.Cast<object>().Count() ?? 0;
        }

        private bool IsUiBusy()
        {
            return Mouse.OverrideCursor != null;
        }
    }
}
