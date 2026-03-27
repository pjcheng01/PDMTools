using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Collections;
using System.IO;
using System.Globalization;
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
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;

namespace PDMTools
{
    public partial class MainWindow : Window
    {
        private const string VaultRootPath = @"C:\CP-PDM";
        /// <summary>與程式執行檔同層之下存放 JSON 快照的資料夾名稱。</summary>
        private const string SnapshotLibraryFolderName = "PdmSnapshots";
        private const int MaxRecentAssemblyPaths = 12;
        private static readonly string RecentAssemblyPathsFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "PDMTools",
            "recent-assemblies.txt");
        /// <summary>參考稽核預覽篩選：路徑是否落在 S 槽（不區分大小寫）。</summary>
        private const string ReferenceAuditSDriveFolderPrefix = @"S:\";
        private PdmBomExportService _exportService;
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
        private SolidWorksReferenceAuditService _referenceAuditService;
        private ICollectionView _auditRowsView;
        private ReferenceAuditPreviewFilterMode _auditPreviewFilterMode = ReferenceAuditPreviewFilterMode.All;
        private MainWorkMode _currentWorkMode = MainWorkMode.Bom;
        private bool _referenceAuditEnvironmentReady = true;
        private string _referenceAuditEnvironmentMessage = string.Empty;
        private bool _configurationLoadInProgress;

        private static readonly Brush BomConfigStatusReadyBrush = new SolidColorBrush(Color.FromRgb(46, 125, 50));
        private static readonly Brush BomConfigStatusWarnBrush = new SolidColorBrush(Color.FromRgb(198, 40, 40));
        private static readonly Brush BomConfigStatusInfoBrush = new SolidColorBrush(Color.FromRgb(21, 101, 192));

        private enum ReferenceAuditPreviewFilterMode
        {
            All = 0,
            InVaultOnly = 1,
            NotInVaultOnly = 2,
            FileMissingOnly = 3,
            UnderSDriveOnly = 4,
            NotUnderSDriveOnly = 5,
            DrawingsOnly = 6,
            ExcludeDrawings = 7
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
        private readonly List<string> _recentAssemblyPaths = new List<string>();

        public MainWindow()
        {
            InitializeComponent();
            ConfigComboBox.AddHandler(TextBoxBase.TextChangedEvent, new TextChangedEventHandler((_, __) => UpdateBomConfigurationUiState()));
            LoadRecentAssemblyPaths();
            UpdateDepthInputStates();
            ApplyBomDepthDefaultByDefinition();
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
            UpdateBomConfigurationUiState();
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

        private sealed class BomItemLevelComparer : IComparer
        {
            private readonly ListSortDirection _direction;

            public BomItemLevelComparer(ListSortDirection direction)
            {
                _direction = direction;
            }

            public int Compare(object x, object y)
            {
                var a = x as BomItem;
                var b = y as BomItem;
                var result = CompareLevel(a?.Level, b?.Level);
                return _direction == ListSortDirection.Ascending ? result : -result;
            }

            private static int CompareLevel(string left, string right)
            {
                var la = ParseLevel(left);
                var rb = ParseLevel(right);
                var n = Math.Min(la.Count, rb.Count);
                for (var i = 0; i < n; i++)
                {
                    var cmp = la[i].CompareTo(rb[i]);
                    if (cmp != 0)
                    {
                        return cmp;
                    }
                }

                return la.Count.CompareTo(rb.Count);
            }

            private static List<int> ParseLevel(string level)
            {
                var text = (level ?? string.Empty).Trim();
                if (text.Length == 0)
                {
                    return new List<int> { int.MaxValue };
                }

                var parts = text.Split('.');
                var result = new List<int>(parts.Length);
                foreach (var p in parts)
                {
                    if (int.TryParse(p, out var v))
                    {
                        result.Add(v);
                    }
                    else
                    {
                        result.Add(int.MaxValue - 1);
                    }
                }

                return result;
            }
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

        private void BomDataGrid_OnSorting(object sender, DataGridSortingEventArgs e)
        {
            var headerInfo = e.Column?.Header as ColumnHeaderInfo;
            if (!string.Equals(headerInfo?.Key, "Level", StringComparison.OrdinalIgnoreCase))
            {
                if (_bomItemsView is ListCollectionView lcvDefault)
                {
                    lcvDefault.CustomSort = null;
                }

                return;
            }

            e.Handled = true;
            var direction = e.Column.SortDirection != ListSortDirection.Ascending
                ? ListSortDirection.Ascending
                : ListSortDirection.Descending;

            foreach (var col in BomDataGrid.Columns)
            {
                col.SortDirection = null;
            }

            e.Column.SortDirection = direction;
            if (_bomItemsView is ListCollectionView lcv)
            {
                lcv.CustomSort = new BomItemLevelComparer(direction);
            }
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

        private void RecentPathsButton_OnClick(object sender, RoutedEventArgs e)
        {
            if (_recentAssemblyPaths.Count == 0)
            {
                MessageBox.Show(this, "目前沒有最近使用的組合件。", "提醒", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (sender is not Button button)
            {
                return;
            }

            var menu = new ContextMenu();
            foreach (var p in _recentAssemblyPaths)
            {
                var header = p;
                if (header.Length > 100)
                {
                    header = "..." + header.Substring(header.Length - 97);
                }

                var mi = new MenuItem
                {
                    Header = header,
                    ToolTip = p,
                    Tag = p
                };
                mi.Click += (_, __) =>
                {
                    SetAssemblyPathText(p, isImported: true);
                    StatusTextBlock.Text = "已套用最近使用路徑。";
                    _ = PopulateConfigurationComboForPathAsync(p, showFailureDialogs: false);
                };
                menu.Items.Add(mi);
            }

            menu.PlacementTarget = button;
            menu.Placement = PlacementMode.Bottom;
            menu.IsOpen = true;
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
                    _ = PopulateConfigurationComboForPathAsync(auditDialog.FileName, showFailureDialogs: false);
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

        private void BomLevelDefinitionComboBox_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ApplyBomDepthDefaultByDefinition();
        }

        private async void RunReferenceAuditButton_OnClick(object sender, RoutedEventArgs e)
        {
            if (!EvaluateReferenceAuditEnvironment(showMessageBoxWhenUnavailable: true))
            {
                return;
            }

            int? maxAuditDepth = null;
            if (ReferenceAuditAllDepthCheckBox?.IsChecked == true)
            {
                maxAuditDepth = null;
            }
            else
            {
                var txt = ReferenceAuditMaxDepthTextBox?.Text?.Trim() ?? string.Empty;
                if (!int.TryParse(txt, out var parsed) || parsed <= 0)
                {
                    MessageBox.Show(this, "請輸入正確的「最大層數」。例如：3", "提醒", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                maxAuditDepth = parsed;
            }

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
                var includeDrw = ReferenceAuditIncludeDrawingsCheckBox?.IsChecked == true;
                await Dispatcher.InvokeAsync(() =>
                {
                    rows = (_referenceAuditService ??= new SolidWorksReferenceAuditService())
                        .AuditAssembly(path, VaultRootPath, maxAuditDepth, includeDrw);
                }, DispatcherPriority.Normal);

                rows ??= Array.Empty<ReferenceAuditRow>();

                foreach (var r in rows)
                    _auditRows.Add(r);

                var notInVault = rows.Count(r =>
                    !r.IsUnderVaultRoot && !string.IsNullOrWhiteSpace(r.FullPath));
                var drwCount = rows.Count(r => r.IsDrawing);
                StatusTextBlock.Text = includeDrw
                    ? $"稽核完成，共 {rows.Count} 筆（含工程圖 {drwCount} 筆；其中 {notInVault} 筆路徑不在 Vault 根目錄 {VaultRootPath} 下）。"
                    : $"稽核完成，共 {rows.Count} 筆引用（其中 {notInVault} 筆路徑不在 Vault 根目錄 {VaultRootPath} 下）。";
                ExportReferenceAuditButton.IsEnabled = rows.Count > 0;
                _auditRowsView?.Refresh();
                UpdateAuditPreviewCountLabel();
            }
            catch (Exception ex)
            {
                var detail = BuildReferenceAuditFailureMessage(ex);
                MessageBox.Show(this, detail, "稽核失敗", MessageBoxButton.OK, MessageBoxImage.Error);
                StatusTextBlock.Text = "稽核失敗。";
            }
            finally
            {
                SetUiBusy(false);
            }
        }

        private void MainModeRadioButton_OnChecked(object sender, RoutedEventArgs e)
        {
            // XAML 初始化期間可能先觸發 Checked，此時另一顆 RadioButton 尚未建立。
            var mode = MainWorkMode.Bom;
            if (ReferenceAuditModeRadioButton != null)
            {
                mode = ReferenceAuditModeRadioButton.IsChecked == true
                    ? MainWorkMode.ReferenceAudit
                    : MainWorkMode.Bom;
            }
            else if (sender is RadioButton rb)
            {
                mode = string.Equals(rb.Name, nameof(ReferenceAuditModeRadioButton), StringComparison.Ordinal)
                    ? MainWorkMode.ReferenceAudit
                    : MainWorkMode.Bom;
            }

            ApplyMainModeUi(mode);
        }

        private void DepthAllCheckBox_OnCheckedChanged(object sender, RoutedEventArgs e)
        {
            UpdateDepthInputStates();
        }

        private void UpdateDepthInputStates()
        {
            if (MaxBomDepthTextBox != null && AllBomDepthCheckBox != null)
            {
                MaxBomDepthTextBox.IsEnabled = AllBomDepthCheckBox.IsChecked != true;
            }

            if (ReferenceAuditMaxDepthTextBox != null && ReferenceAuditAllDepthCheckBox != null)
            {
                ReferenceAuditMaxDepthTextBox.IsEnabled = ReferenceAuditAllDepthCheckBox.IsChecked != true;
            }
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

            if (!isBomMode)
            {
                EvaluateReferenceAuditEnvironment(showMessageBoxWhenUnavailable: false);
            }

            if (StatusTextBlock != null)
                StatusTextBlock.Text = isBomMode
                    ? (_bomItems.Count > 0 ? $"目前為 Vault BOM 模式（{_bomItems.Count} 筆）。" : "目前為 Vault BOM 模式。")
                    : (_referenceAuditEnvironmentReady
                        ? (_auditRows.Count > 0 ? $"目前為參考稽核模式（{_auditRows.Count} 筆）。" : "目前為參考稽核模式。")
                        : $"目前為參考稽核模式（環境未就緒：{_referenceAuditEnvironmentMessage}）");

            if (isBomMode)
            {
                UpdateBomConfigurationUiState();
                Dispatcher.BeginInvoke(DispatcherPriority.Loaded, new Action(() => _ = EnsureBomConfigurationsReadyAsync()));
            }
            else if (BomConfigStatusTextBlock != null)
            {
                BomConfigStatusTextBlock.Text = string.Empty;
            }

            SetUiBusy(IsUiBusy());
        }

        private bool EvaluateReferenceAuditEnvironment(bool showMessageBoxWhenUnavailable)
        {
            Type swType = null;
            try
            {
                swType = Type.GetTypeFromProgID("SldWorks.Application");
            }
            catch
            {
                swType = null;
            }

            if (swType == null)
            {
                _referenceAuditEnvironmentReady = false;
                _referenceAuditEnvironmentMessage = "偵測不到 SolidWorks COM（SldWorks.Application）";
            }
            else
            {
                _referenceAuditEnvironmentReady = true;
                _referenceAuditEnvironmentMessage = string.Empty;
            }

            if (RunReferenceAuditButton != null)
            {
                RunReferenceAuditButton.IsEnabled = _referenceAuditEnvironmentReady;
            }

            if (!_referenceAuditEnvironmentReady && showMessageBoxWhenUnavailable)
            {
                MessageBox.Show(
                    this,
                    "目前無法執行「參考稽核」。\n\n" +
                    "原因：偵測不到 SolidWorks 的 COM 元件（SldWorks.Application）。\n\n" +
                    "請確認：\n" +
                    "1) 本機已安裝 SolidWorks\n" +
                    "2) SolidWorks 安裝/修復完成，COM 註冊正常\n" +
                    "3) 以目前使用者可正常啟動 SolidWorks",
                    "環境未就緒",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
            }

            return _referenceAuditEnvironmentReady;
        }

        private static string BuildReferenceAuditFailureMessage(Exception ex)
        {
            var msg = ex?.Message ?? "未知錯誤";
            if (msg.IndexOf("SldWorks.Application", StringComparison.OrdinalIgnoreCase) >= 0 ||
                msg.IndexOf("無法建立 SolidWorks", StringComparison.OrdinalIgnoreCase) >= 0 ||
                msg.IndexOf("ProgID", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "無法啟動 SolidWorks 進行稽核。\n\n" +
                       "請先確認 SolidWorks 已安裝且可由目前使用者正常啟動，" +
                       "再重新執行稽核。\n\n" +
                       $"原始訊息：{msg}";
            }

            return msg;
        }

        private void ReferenceAuditPreviewFilterComboBox_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is not ComboBox cb)
                return;
            var idx = cb.SelectedIndex;
            if (idx < 0)
                idx = 0;
            var maxIdx = (int)ReferenceAuditPreviewFilterMode.ExcludeDrawings;
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
                ReferenceAuditPreviewFilterMode.DrawingsOnly => row.IsDrawing,
                ReferenceAuditPreviewFilterMode.ExcludeDrawings => !row.IsDrawing,
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

            var configurationName = GetEffectiveConfigurationName();
            if (string.IsNullOrWhiteSpace(configurationName))
            {
                MessageBox.Show(
                    this,
                    "請指定要使用的組態：\n\n"
                    + "• 若已載入清單：請從「組態」下拉選單選取一項\n"
                    + "• 若清單載入失敗：請直接在「組態」欄位輸入或貼上組態名稱\n"
                    + "• 亦可按「重新載入組態」向 PDM 再試一次",
                    "尚未選組態",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                UpdateBomConfigurationUiState();
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

            var prevProgressPanelVis = MainProgressPanel != null
                ? MainProgressPanel.Visibility
                : Visibility.Visible;
            if (MainProgressPanel != null)
                MainProgressPanel.Visibility = Visibility.Collapsed;

            var mainWasEnabled = IsEnabled;
            IsEnabled = false;

            using var cts = new CancellationTokenSource();
            BomGrabProgressWindow progressWin = null;
            IProgress<ProgressInfo> progress = new Progress<ProgressInfo>(p =>
            {
                if (progressWin != null)
                    Dispatcher.Invoke(() => progressWin.UpdateProgress(p));
            });

            try
            {
                progressWin = new BomGrabProgressWindow(cts)
                {
                    Title = "顯示工程圖 · 作業中",
                    WindowStartupLocation = WindowStartupLocation.CenterScreen
                };
                progressWin.Show();
                progressWin.Activate();

                _exportService = _exportService ?? new PdmBomExportService();
                var withDrawings = await _exportService.AppendDrawingItemsAsync(
                    _rawBomItems,
                    _activeCardVarNames,
                    progress,
                    cts.Token);

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
                MessageBox.Show(
                    progressWin != null ? (Window)progressWin : this,
                    "作業已取消。",
                    "取消",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                ShowDrawingsToggle.IsChecked = false;
                MessageBox.Show(
                    progressWin != null ? (Window)progressWin : this,
                    ex.Message,
                    "錯誤",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            finally
            {
                progressWin?.CloseFromOwner();
                IsEnabled = mainWasEnabled;
                if (MainProgressPanel != null)
                    MainProgressPanel.Visibility = prevProgressPanelVis;
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
            ConfigComboBox.Text = string.Empty;
            ConfigComboBox.IsEnabled = false;
            UpdateBomConfigurationUiState();
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

        private string GetEffectiveConfigurationName()
        {
            if (ConfigComboBox == null)
            {
                return string.Empty;
            }

            if (ConfigComboBox.SelectedItem != null)
            {
                var s = ConfigComboBox.SelectedItem as string ?? ConfigComboBox.SelectedItem.ToString();
                if (!string.IsNullOrWhiteSpace(s))
                {
                    return s.Trim();
                }
            }

            return (ConfigComboBox.Text ?? string.Empty).Trim();
        }

        private bool TryGetNormalizedAssemblyPathFromUi(out string normalizedPath)
        {
            normalizedPath = null;
            var raw = AssemblyPathTextBox?.Text ?? string.Empty;
            if (string.IsNullOrWhiteSpace(raw))
            {
                return false;
            }

            if (!TryValidateImportPath(raw, out var np, out _))
            {
                return false;
            }

            if (!File.Exists(np))
            {
                return false;
            }

            normalizedPath = np;
            return true;
        }

        private void UpdateBomConfigurationUiState()
        {
            if (BomConfigStatusTextBlock == null)
            {
                return;
            }

            if (_currentWorkMode != MainWorkMode.Bom)
            {
                BomConfigStatusTextBlock.Text = string.Empty;
                return;
            }

            if (_configurationLoadInProgress)
            {
                BomConfigStatusTextBlock.Text = "組態：載入中…";
                BomConfigStatusTextBlock.Foreground = Brushes.DimGray;
                return;
            }

            if (!TryGetNormalizedAssemblyPathFromUi(out _))
            {
                BomConfigStatusTextBlock.Text = "組態：請先指定有效的 .sldasm 路徑";
                BomConfigStatusTextBlock.Foreground = Brushes.DimGray;
                return;
            }

            if (ConfigComboBox == null)
            {
                return;
            }

            if (!ConfigComboBox.IsEnabled)
            {
                BomConfigStatusTextBlock.Text = "組態：請先導入組合件以載入清單";
                BomConfigStatusTextBlock.Foreground = Brushes.DimGray;
                return;
            }

            var name = GetEffectiveConfigurationName();
            if (ConfigComboBox.Items.Count > 0)
            {
                if (!string.IsNullOrWhiteSpace(name))
                {
                    BomConfigStatusTextBlock.Text = $"組態：已就緒（{name}）";
                    BomConfigStatusTextBlock.Foreground = BomConfigStatusReadyBrush;
                }
                else
                {
                    BomConfigStatusTextBlock.Text = "組態：尚未選組態—請從下拉選取或手動輸入名稱";
                    BomConfigStatusTextBlock.Foreground = BomConfigStatusWarnBrush;
                }

                return;
            }

            if (!string.IsNullOrWhiteSpace(name))
            {
                BomConfigStatusTextBlock.Text = $"組態：將使用手動輸入「{name}」（建議按「重新載入組態」確認）";
                BomConfigStatusTextBlock.Foreground = BomConfigStatusInfoBrush;
            }
            else
            {
                BomConfigStatusTextBlock.Text = "組態：清單未載入—請手動輸入名稱，或按「重新載入組態」";
                BomConfigStatusTextBlock.Foreground = BomConfigStatusWarnBrush;
            }
        }

        private void ConfigComboBox_SelectionOrTextChanged(object sender, RoutedEventArgs e)
        {
            UpdateBomConfigurationUiState();
        }

        private async void ReloadConfigurationsButton_OnClick(object sender, RoutedEventArgs e)
        {
            if (!TryGetNormalizedAssemblyPathFromUi(out var path))
            {
                MessageBox.Show(this, "請先指定有效的 .sldasm 路徑。", "提醒", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            await PopulateConfigurationComboForPathAsync(path, showFailureDialogs: true);
        }

        private async Task EnsureBomConfigurationsReadyAsync()
        {
            if (_currentWorkMode != MainWorkMode.Bom)
            {
                return;
            }

            if (!TryGetNormalizedAssemblyPathFromUi(out var path))
            {
                await Dispatcher.InvokeAsync(UpdateBomConfigurationUiState);
                return;
            }

            var needLoad = await Dispatcher.InvokeAsync(() => ConfigComboBox.Items.Count == 0);
            if (!needLoad)
            {
                await Dispatcher.InvokeAsync(UpdateBomConfigurationUiState);
                return;
            }

            await PopulateConfigurationComboForPathAsync(path, showFailureDialogs: false);
        }

        /// <param name="showFailureDialogs">false 時僅更新狀態列，不跳出阻斷式對話框（例如由稽核模式切回或背景預載）。</param>
        private async Task PopulateConfigurationComboForPathAsync(string assemblyPath, bool showFailureDialogs = true)
        {
            if (string.IsNullOrWhiteSpace(assemblyPath) || !File.Exists(assemblyPath))
            {
                await Dispatcher.InvokeAsync(UpdateBomConfigurationUiState);
                return;
            }

            _configurationLoadInProgress = true;
            await Dispatcher.InvokeAsync(() =>
            {
                SetUiBusy(IsUiBusy());
                if (BomConfigStatusTextBlock != null)
                {
                    BomConfigStatusTextBlock.Text = "組態：載入中…";
                    BomConfigStatusTextBlock.Foreground = Brushes.DimGray;
                }
            });

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
                        ConfigComboBox.IsEnabled = true;
                        StatusTextBlock.Text = showFailureDialogs
                            ? "已選擇組合件，但無法從 PDM 讀取組態清單。"
                            : "已選擇組合件，尚未取得組態清單；可手動輸入組態或按「重新載入組態」。";
                        var detail = _exportService?.LastConfigurationEnumerationDiag ?? string.Empty;
                        if (detail.Length > 2800)
                        {
                            detail = detail.Substring(0, 2800) + "\n…（以下略）";
                        }

                        if (showFailureDialogs)
                        {
                            MessageBox.Show(
                                this,
                                "無法從 PDM 讀取此組合件的組態清單。\n\n"
                                + "您仍可嘗試：\n"
                                + "• 在「組態」欄位手動輸入或貼上組態名稱後按「開始抓取」\n"
                                + "• 按「重新載入組態」再試一次\n\n"
                                + "若持續失敗請確認：\n"
                                + "• 檔案已同步為本機最新版且可正常在 PDM 中開啟\n"
                                + "• 本機 PDM 用戶端與程式參考的 EPDM Interop 版本一致\n"
                                + "• 必要時請先以 SolidWorks／PDM 開啟該組合件一次後再試\n\n"
                                + (string.IsNullOrWhiteSpace(detail)
                                    ? string.Empty
                                    : "── 程式診斷（可複製給開發者）──\n" + detail),
                                "無法載入組態",
                                MessageBoxButton.OK,
                                MessageBoxImage.Warning);
                        }

                        UpdateBomConfigurationUiState();
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
                    UpdateBomConfigurationUiState();
                });
            }
            catch (Exception ex)
            {
                await Dispatcher.InvokeAsync(() =>
                {
                    ConfigComboBox.Items.Clear();
                    ConfigComboBox.SelectedItem = null;
                    ConfigComboBox.IsEnabled = true;
                    StatusTextBlock.Text = showFailureDialogs
                        ? "已選擇組合件，但讀取組態時發生錯誤。"
                        : "已選擇組合件，讀取組態失敗；可直接輸入組態名稱，或按「重新載入組態」。";

                    if (showFailureDialogs)
                    {
                        MessageBox.Show(
                            this,
                            "讀取組態清單時發生錯誤：\n" + ex.Message + "\n\n您仍可手動輸入組態名稱，或按「重新載入組態」再試。",
                            "錯誤",
                            MessageBoxButton.OK,
                            MessageBoxImage.Error);
                    }

                    UpdateBomConfigurationUiState();
                });
            }
            finally
            {
                _configurationLoadInProgress = false;
                await Dispatcher.InvokeAsync(() =>
                {
                    SetUiBusy(IsUiBusy());
                    UpdateBomConfigurationUiState();
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

            var prevProgressPanelVis = MainProgressPanel != null
                ? MainProgressPanel.Visibility
                : Visibility.Visible;
            if (MainProgressPanel != null)
                MainProgressPanel.Visibility = Visibility.Collapsed;

            var mainWasEnabled = IsEnabled;
            IsEnabled = false;

            using var cts = new CancellationTokenSource();
            BomGrabProgressWindow progressWin = null;
            IProgress<ProgressInfo> progress = new Progress<ProgressInfo>(p =>
            {
                if (progressWin != null)
                    Dispatcher.Invoke(() => progressWin.UpdateProgress(p));
            });

            try
            {
                progressWin = new BomGrabProgressWindow(cts)
                {
                    WindowStartupLocation = WindowStartupLocation.CenterScreen
                };
                progressWin.Show();
                progressWin.Activate();

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
                    cts.Token);

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
                MessageBox.Show(
                    progressWin != null ? (Window)progressWin : this,
                    "作業已取消。",
                    "取消",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    progressWin != null ? (Window)progressWin : this,
                    ex.Message,
                    "錯誤",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            finally
            {
                progressWin?.CloseFromOwner();
                IsEnabled = mainWasEnabled;
                if (MainProgressPanel != null)
                    MainProgressPanel.Visibility = prevProgressPanelVis;
                SetUiBusy(false);
            }
        }

        private async Task RunExportExcelOnlyAsync(string outputPath)
        {
            SetUiBusy(true);
            SetProgressPercent(0);

            using var cts = new CancellationTokenSource();
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
                    cts.Token);

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
            // InitializeComponent 期間，作業模式 Radio 的 Checked 可能早於下方面板建立，須逐項判空。
            var isBomMode = _currentWorkMode == MainWorkMode.Bom;

            if (BrowseButton != null)
            {
                BrowseButton.IsEnabled = !isBusy;
            }
            if (RecentPathsButton != null)
            {
                RecentPathsButton.IsEnabled = !isBusy && _recentAssemblyPaths.Count > 0;
            }

            if (StartGrabButton != null)
            {
                StartGrabButton.IsEnabled = !isBusy && isBomMode;
            }

            if (ExportButton != null)
            {
                ExportButton.IsEnabled = !isBusy && isBomMode;
            }

            if (ColumnSettingsButton != null)
            {
                ColumnSettingsButton.IsEnabled = !isBusy && isBomMode;
            }

            if (CompareIgpBomButton != null)
            {
                CompareIgpBomButton.IsEnabled = !isBusy && isBomMode && _bomItems.Count > 0;
            }

            if (BatchDrawingPdfButton != null)
            {
                BatchDrawingPdfButton.IsEnabled = !isBusy && isBomMode && _bomItems.Count > 0;
            }

            if (CustomPropTemplateButton != null)
            {
                CustomPropTemplateButton.IsEnabled = !isBusy && isBomMode && _bomItems.Count > 0;
            }

            if (DrawingSheetFormatButton != null)
            {
                DrawingSheetFormatButton.IsEnabled = !isBusy && isBomMode && _bomItems.Count > 0;
            }

            if (IgpBomCompareDepthComboBox != null)
            {
                IgpBomCompareDepthComboBox.IsEnabled = !isBusy && isBomMode && _bomItems.Count > 0;
            }

            if (ClearFiltersButton != null)
            {
                ClearFiltersButton.IsEnabled = !isBusy && isBomMode && HasActiveFilters();
            }

            if (ShowDrawingsToggle != null)
            {
                ShowDrawingsToggle.IsEnabled = !isBusy && isBomMode && _rawBomItems.Count > 0;
            }

            if (RunReferenceAuditButton != null)
            {
                RunReferenceAuditButton.IsEnabled = !isBusy && !isBomMode && _referenceAuditEnvironmentReady;
            }

            if (ExportReferenceAuditButton != null)
            {
                ExportReferenceAuditButton.IsEnabled = !isBusy && !isBomMode && _auditRows.Count > 0;
            }

            if (ReferenceAuditPreviewFilterComboBox != null)
            {
                ReferenceAuditPreviewFilterComboBox.IsEnabled = !isBusy && !isBomMode;
            }

            if (BomModeRadioButton != null)
            {
                BomModeRadioButton.IsEnabled = !isBusy;
            }

            if (ReferenceAuditModeRadioButton != null)
            {
                ReferenceAuditModeRadioButton.IsEnabled = !isBusy;
            }

            if (SaveProjectButton != null)
            {
                SaveProjectButton.IsEnabled = !isBusy;
            }

            if (LoadProjectButton != null)
            {
                LoadProjectButton.IsEnabled = !isBusy;
            }

            if (ReloadConfigurationsButton != null)
            {
                ReloadConfigurationsButton.IsEnabled = !isBusy && isBomMode && !_configurationLoadInProgress &&
                                                       TryGetNormalizedAssemblyPathFromUi(out _);
            }

            if (BomDataGrid != null)
            {
                BomDataGrid.IsEnabled = !isBusy && isBomMode;
            }

            Mouse.OverrideCursor = isBusy ? Cursors.Wait : null;
        }

        /// <summary>讀取「比對 iGP BOM」所用之深度上限（2～4），異常時預設為 3。</summary>
        private int GetIgpBomCompareMaxDepthFromUi()
        {
            if (IgpBomCompareDepthComboBox?.SelectedItem is System.Windows.Controls.ComboBoxItem item &&
                item.Tag is string tag &&
                int.TryParse(tag, NumberStyles.Integer, CultureInfo.InvariantCulture, out var d) &&
                d >= 2 && d <= 4)
            {
                return d;
            }

            return 3;
        }

        /// <summary>目前 DataGrid 篩選後可見之列中，標記為工程圖（IsDrawing）的項目。</summary>
        private IEnumerable<BomItem> GetVisibleFilteredDrawingBomItems()
        {
            if (_bomItemsView == null)
            {
                return Enumerable.Empty<BomItem>();
            }

            return _bomItemsView
                .Cast<BomItem>()
                .Where(i => i != null && i.IsDrawing);
        }

        private void BatchWorkflowTransitionButton_OnClick(object sender, RoutedEventArgs e)
        {
            if (_currentWorkMode != MainWorkMode.Bom)
            {
                return;
            }

            var items = GetVisibleFilteredBomItems().ToList();
            if (items.Count == 0)
            {
                MessageBox.Show(
                    this,
                    "目前篩選後可見之列中沒有檔案路徑，或尚未抓取 BOM。",
                    "提醒",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            var paths = items
                .Select(i => i.FullPath?.Trim() ?? string.Empty)
                .Where(p => !string.IsNullOrWhiteSpace(p))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            _exportService = _exportService ?? new PdmBomExportService();
            var win = new BatchWorkflowTransitionWindow(paths, _exportService)
            {
                Owner = this
            };
            win.Show();
        }

        private void BatchDrawingPdfButton_OnClick(object sender, RoutedEventArgs e)
        {
            if (_currentWorkMode != MainWorkMode.Bom)
            {
                return;
            }

            var drawings = GetVisibleFilteredDrawingBomItems().ToList();
            if (drawings.Count == 0)
            {
                MessageBox.Show(
                    this,
                    "目前篩選後可見之列中沒有工程圖，或尚未抓取 BOM。請先抓取 BOM，並確認工程圖列已顯示且未被篩選排除。",
                    "提醒",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            _exportService = _exportService ?? new PdmBomExportService();
            var win = new BatchDrawingPdfWindow(drawings, _exportService)
            {
                Owner = this
            };
            win.Show();
        }

        private void DrawingSheetFormatButton_OnClick(object sender, RoutedEventArgs e)
        {
            if (_currentWorkMode != MainWorkMode.Bom)
                return;

            var drawings = GetVisibleFilteredDrawingBomItems().ToList();
            if (drawings.Count == 0)
            {
                MessageBox.Show(
                    this,
                    "目前篩選後可見之列中沒有工程圖，或尚未抓取 BOM。",
                    "提醒",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            _exportService = _exportService ?? new PdmBomExportService();
            var win = new DrawingSheetFormatWindow(drawings, _exportService)
            {
                Owner = this
            };
            win.Show();
        }

        private void CustomPropTemplateButton_OnClick(object sender, RoutedEventArgs e)
        {
            if (_currentWorkMode != MainWorkMode.Bom)
                return;

            var items = GetVisibleFilteredBomItems().ToList();
            if (items.Count == 0)
            {
                MessageBox.Show(
                    this,
                    "目前篩選後可見之列中沒有 SolidWorks 檔案（.sldprt / .sldasm / .slddrw），或尚未抓取 BOM。",
                    "提醒",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            _exportService = _exportService ?? new PdmBomExportService();
            var win = new CustomPropertyTemplateWindow(items, _exportService)
            {
                Owner = this
            };
            win.Show();
        }

        private IEnumerable<BomItem> GetVisibleFilteredBomItems()
        {
            if (_bomItemsView == null)
                return Enumerable.Empty<BomItem>();

            return _bomItemsView
                .Cast<BomItem>()
                .Where(i => i != null && !string.IsNullOrWhiteSpace(i.FullPath));
        }

        private async void CompareIgpBomButton_OnClick(object sender, RoutedEventArgs e)
        {
            if (_currentWorkMode != MainWorkMode.Bom)
            {
                return;
            }

            if (_bomItems.Count == 0)
            {
                MessageBox.Show(this, "請先在 Vault BOM 模式按「開始抓取」取得資料後，再進行比對。", "提醒", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var dialog = new OpenFileDialog
            {
                Title = "選擇 iGP BOM（BOMMI02.xlsx 或 .csv）",
                Filter = "Excel / CSV|*.xlsx;*.xlsm;*.csv|Excel (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|CSV (*.csv)|*.csv|所有檔案 (*.*)|*.*"
            };
            if (dialog.ShowDialog(this) != true)
            {
                return;
            }

            SetUiBusy(true);
            SetProgressPercent(0);
            StatusTextBlock.Text = "正在比對 iGP BOM…";

            using var compareCts = new CancellationTokenSource();
            try
            {
                IProgress<ProgressInfo> progress = new Progress<ProgressInfo>(p =>
                {
                    SetProgressPercent(p.Percentage);
                    StatusTextBlock.Text = p.Message;
                });

                var compareMaxDepth = GetIgpBomCompareMaxDepthFromUi();

                var comparer = new IgpBomComparer();
                var result = await comparer.CompareAsync(
                    pdmBomItems: _bomItems.ToList(),
                    igpBomFilePath: dialog.FileName,
                    maxDepth: compareMaxDepth,
                    excludeDrawings: true,
                    progress: progress,
                    cancellationToken: compareCts.Token);

                var win = new IgpBomCompareWindow(result)
                {
                    Owner = this
                };
                win.Show();

                StatusTextBlock.Text = $"比對完成：Stage1 {result.Stage1Rows.Count} 筆、Stage2 {result.Stage2Rows.Count} 筆（結果視窗可篩選各欄）。";
                SetProgressPercent(0);
            }
            catch (OperationCanceledException)
            {
                MessageBox.Show(this, "作業已取消。", "取消", MessageBoxButton.OK, MessageBoxImage.Information);
                StatusTextBlock.Text = "已取消比對。";
                SetProgressPercent(0);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "比對失敗", MessageBoxButton.OK, MessageBoxImage.Error);
                StatusTextBlock.Text = "比對失敗。";
                SetProgressPercent(0);
            }
            finally
            {
                SetUiBusy(false);
            }
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

        private void ApplyBomDepthDefaultByDefinition()
        {
            if (BomLevelDefinitionComboBox == null || MaxBomDepthTextBox == null)
            {
                return;
            }

            // 定義A（含根）與定義B（不含根）皆預設 3，仍可手動修改。
            MaxBomDepthTextBox.Text = "3";
        }

        private void LoadRecentAssemblyPaths()
        {
            _recentAssemblyPaths.Clear();
            try
            {
                if (!File.Exists(RecentAssemblyPathsFilePath))
                {
                    return;
                }

                foreach (var line in File.ReadAllLines(RecentAssemblyPathsFilePath))
                {
                    var p = line?.Trim() ?? string.Empty;
                    if (p.Length == 0)
                    {
                        continue;
                    }

                    if (_recentAssemblyPaths.Any(x => string.Equals(x, p, StringComparison.OrdinalIgnoreCase)))
                    {
                        continue;
                    }

                    _recentAssemblyPaths.Add(p);
                    if (_recentAssemblyPaths.Count >= MaxRecentAssemblyPaths)
                    {
                        break;
                    }
                }
            }
            catch
            {
                // ignore and keep empty list
            }
        }

        private void SaveRecentAssemblyPaths()
        {
            try
            {
                var dir = Path.GetDirectoryName(RecentAssemblyPathsFilePath);
                if (!string.IsNullOrWhiteSpace(dir))
                {
                    Directory.CreateDirectory(dir);
                }

                File.WriteAllLines(RecentAssemblyPathsFilePath, _recentAssemblyPaths);
            }
            catch
            {
                // ignore
            }
        }

        private void AddRecentAssemblyPath(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return;
            }

            string normalized;
            try
            {
                normalized = Path.GetFullPath(path.Trim().Trim('"'));
            }
            catch
            {
                return;
            }

            _recentAssemblyPaths.RemoveAll(x => string.Equals(x, normalized, StringComparison.OrdinalIgnoreCase));
            _recentAssemblyPaths.Insert(0, normalized);
            if (_recentAssemblyPaths.Count > MaxRecentAssemblyPaths)
            {
                _recentAssemblyPaths.RemoveRange(MaxRecentAssemblyPaths, _recentAssemblyPaths.Count - MaxRecentAssemblyPaths);
            }

            SaveRecentAssemblyPaths();
            SetUiBusy(IsUiBusy());
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
            if (isImported)
            {
                AddRecentAssemblyPath(path);
            }
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

        // ── 專案／快照檔（JSON）──────────────────────────────────────────

        private static string GetSnapshotLibraryDirectoryPath() =>
            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, SnapshotLibraryFolderName);

        private static void EnsureSnapshotLibraryDirectoryExists(string directoryPath)
        {
            if (string.IsNullOrWhiteSpace(directoryPath))
            {
                return;
            }

            Directory.CreateDirectory(directoryPath);
        }

        /// <summary>預設快照檔名：組合件主檔名 + 模式標籤（VaultBom／RefAudit）+ 本機時間戳。</summary>
        private string BuildDefaultSnapshotFileName()
        {
            var raw = AssemblyPathTextBox?.Text?.Trim().Trim('"') ?? string.Empty;
            var baseName = "組合件";
            if (!string.IsNullOrWhiteSpace(raw))
            {
                try
                {
                    baseName = Path.GetFileNameWithoutExtension(Path.GetFullPath(raw));
                }
                catch
                {
                    baseName = "組合件";
                }
            }

            foreach (var c in Path.GetInvalidFileNameChars())
            {
                baseName = baseName.Replace(c, '_');
            }

            baseName = baseName.Trim();
            if (string.IsNullOrWhiteSpace(baseName))
            {
                baseName = "組合件";
            }

            if (baseName.Length > 120)
            {
                baseName = baseName.Substring(0, 120);
            }

            var modeTag = _currentWorkMode == MainWorkMode.Bom ? "VaultBom" : "RefAudit";
            return $"{baseName}_{modeTag}_{DateTime.Now:yyyyMMdd_HHmmss}.json";
        }

        private static string FormatSnapshotLocalTime(DateTime savedAt)
        {
            if (savedAt == default)
            {
                return "（未知）";
            }

            var dt = savedAt.Kind == DateTimeKind.Utc ? savedAt.ToLocalTime() : savedAt;
            return dt.ToString("F", CultureInfo.CurrentCulture);
        }

        private void SaveProjectButton_OnClick(object sender, RoutedEventArgs e)
        {
            var snapDir = GetSnapshotLibraryDirectoryPath();
            EnsureSnapshotLibraryDirectoryExists(snapDir);

            var dlg = new SaveFileDialog
            {
                Title = "儲存專案快照",
                Filter = "PDMTools 快照 (*.json)|*.json|所有檔案 (*.*)|*.*",
                InitialDirectory = snapDir,
                FileName = BuildDefaultSnapshotFileName(),
                AddExtension = true,
                OverwritePrompt = true
            };

            if (dlg.ShowDialog(this) != true)
            {
                return;
            }

            try
            {
                var doc = BuildProjectSnapshotDocument();
                PdmProjectSnapshotSerializer.Save(doc, dlg.FileName);
                MessageBox.Show(this, $"已儲存快照：\n{dlg.FileName}", "完成", MessageBoxButton.OK, MessageBoxImage.Information);
                StatusTextBlock.Text = $"已儲存專案快照：{Path.GetFileName(dlg.FileName)}";
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "儲存失敗", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async void LoadProjectButton_OnClick(object sender, RoutedEventArgs e)
        {
            var snapDir = GetSnapshotLibraryDirectoryPath();
            EnsureSnapshotLibraryDirectoryExists(snapDir);

            var dlg = new OpenFileDialog
            {
                Title = "開啟專案快照",
                Filter = "PDMTools 快照 (*.json)|*.json|所有檔案 (*.*)|*.*",
                InitialDirectory = snapDir
            };

            if (dlg.ShowDialog(this) != true)
            {
                return;
            }

            SetUiBusy(true);
            try
            {
                var doc = PdmProjectSnapshotSerializer.Load(dlg.FileName);
                await ApplyProjectSnapshotDocumentAsync(doc);
                MessageBox.Show(this,
                    $"已載入快照（儲存時間本機：{FormatSnapshotLocalTime(doc.SavedAtLocal)}）。",
                    "完成",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "載入失敗", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                SetUiBusy(false);
            }
        }

        private PdmProjectSnapshotDocument BuildProjectSnapshotDocument()
        {
            var doc = new PdmProjectSnapshotDocument
            {
                FormatVersion = 1,
                SavedAtLocal = DateTime.Now,
                AssemblyPath = AssemblyPathTextBox.Text?.Trim() ?? string.Empty,
                WorkMode = _currentWorkMode == MainWorkMode.Bom ? "Bom" : "ReferenceAudit"
            };

            if (_currentWorkMode == MainWorkMode.Bom)
            {
                doc.Bom = new BomSnapshotData
                {
                    BomLevelDefinitionIndex = BomLevelDefinitionComboBox?.SelectedIndex ?? 1,
                    AllBomDepth = AllBomDepthCheckBox?.IsChecked == true,
                    MaxBomDepthText = MaxBomDepthTextBox?.Text?.Trim() ?? "3",
                    ConfigurationName = GetEffectiveConfigurationName(),
                    ShowDrawingsChecked = ShowDrawingsToggle?.IsChecked,
                    ActiveCardVarNames = _activeCardVarNames != null ? new List<string>(_activeCardVarNames) : new List<string>(),
                    ActiveFixedColumns = _activeFixedColumns != null ? new List<string>(_activeFixedColumns) : null,
                    Items = PdmProjectSnapshotSerializer.CloneBomItems(_bomItems),
                    RawBomItems = PdmProjectSnapshotSerializer.CloneBomItems(_rawBomItems),
                    BomItemsWithDrawings = _bomItemsWithDrawings != null
                        ? PdmProjectSnapshotSerializer.CloneBomItems(_bomItemsWithDrawings)
                        : null,
                    ColumnFilters = BuildBomColumnFilterSnapshots()
                };
            }
            else
            {
                doc.ReferenceAudit = new ReferenceAuditSnapshotData
                {
                    ReferenceAuditMaxDepthText = ReferenceAuditMaxDepthTextBox?.Text?.Trim() ?? "3",
                    ReferenceAuditAllDepth = ReferenceAuditAllDepthCheckBox?.IsChecked == true,
                    ReferenceAuditIncludeDrawings = ReferenceAuditIncludeDrawingsCheckBox?.IsChecked == true,
                    PreviewFilterComboIndex = ReferenceAuditPreviewFilterComboBox?.SelectedIndex ?? 0,
                    Rows = PdmProjectSnapshotSerializer.CloneReferenceAuditRows(_auditRows)
                };
            }

            return doc;
        }

        private List<BomColumnFilterSnapshot> BuildBomColumnFilterSnapshots()
        {
            var list = new List<BomColumnFilterSnapshot>();
            foreach (var kv in _columnFilters)
            {
                var st = kv.Value;
                list.Add(new BomColumnFilterSnapshot
                {
                    Key = st.Key,
                    SelectedValues = st.SelectedValues.OrderBy(x => x, StringComparer.OrdinalIgnoreCase).ToList(),
                    SearchText = st.SearchText ?? string.Empty
                });
            }

            return list;
        }

        private async Task ApplyProjectSnapshotDocumentAsync(PdmProjectSnapshotDocument doc)
        {
            if (doc == null)
            {
                throw new ArgumentNullException(nameof(doc));
            }

            var bomMode = !string.Equals(doc.WorkMode, "ReferenceAudit", StringComparison.OrdinalIgnoreCase);

            SetAssemblyPathText(doc.AssemblyPath ?? string.Empty, isImported: false);

            if (bomMode)
            {
                BomModeRadioButton.IsChecked = true;
            }
            else
            {
                ReferenceAuditModeRadioButton.IsChecked = true;
            }

            ApplyMainModeUi(bomMode ? MainWorkMode.Bom : MainWorkMode.ReferenceAudit);

            if (doc.Bom != null)
            {
                ApplyBomSnapshotData(doc.Bom);
            }
            else
            {
                _bomItems.Clear();
                _rawBomItems.Clear();
                _bomItemsWithDrawings = null;
                if (ShowDrawingsToggle != null)
                {
                    ShowDrawingsToggle.IsEnabled = false;
                    ShowDrawingsToggle.IsChecked = false;
                }

                SetupBomDataGridColumns();
            }

            if (doc.ReferenceAudit != null)
            {
                ApplyReferenceAuditSnapshotData(doc.ReferenceAudit);
            }
            else
            {
                _auditRows.Clear();
                if (ExportReferenceAuditButton != null)
                {
                    ExportReferenceAuditButton.IsEnabled = false;
                }

                _auditRowsView?.Refresh();
                UpdateAuditPreviewCountLabel();
            }

            if (bomMode && doc.Bom != null && TryGetNormalizedAssemblyPathFromUi(out var np))
            {
                await PopulateConfigurationComboForPathAsync(np, showFailureDialogs: false);
                await Dispatcher.InvokeAsync(() =>
                {
                    var cfg = doc.Bom.ConfigurationName ?? string.Empty;
                    if (!string.IsNullOrWhiteSpace(cfg))
                    {
                        ConfigComboBox.Text = cfg;
                        foreach (var it in ConfigComboBox.Items)
                        {
                            var s = it as string ?? it?.ToString();
                            if (string.Equals(s, cfg, StringComparison.OrdinalIgnoreCase))
                            {
                                ConfigComboBox.SelectedItem = it;
                                break;
                            }
                        }
                    }

                    UpdateBomConfigurationUiState();
                });
            }

            await Dispatcher.InvokeAsync(AutoSizeDataGridColumns, DispatcherPriority.Loaded);
            StatusTextBlock.Text = $"已載入專案快照（儲存時間本機：{FormatSnapshotLocalTime(doc.SavedAtLocal)}）。";
        }

        private void ApplyBomSnapshotData(BomSnapshotData b)
        {
            if (b == null)
            {
                return;
            }

            _activeCardVarNames = b.ActiveCardVarNames != null ? new List<string>(b.ActiveCardVarNames) : new List<string>();
            _activeFixedColumns = b.ActiveFixedColumns != null ? new List<string>(b.ActiveFixedColumns) : null;
            if (_activeFixedColumns != null && _activeFixedColumns.Count == 0)
            {
                _activeFixedColumns = null;
            }

            if (BomLevelDefinitionComboBox != null)
            {
                BomLevelDefinitionComboBox.SelectedIndex = Math.Max(0, Math.Min(1, b.BomLevelDefinitionIndex));
            }

            if (AllBomDepthCheckBox != null)
            {
                AllBomDepthCheckBox.IsChecked = b.AllBomDepth;
            }

            if (MaxBomDepthTextBox != null && !string.IsNullOrWhiteSpace(b.MaxBomDepthText))
            {
                MaxBomDepthTextBox.Text = b.MaxBomDepthText;
            }

            UpdateDepthInputStates();

            _rawBomItems = PdmProjectSnapshotSerializer.CloneBomItems(b.RawBomItems);
            _bomItemsWithDrawings = b.BomItemsWithDrawings != null
                ? PdmProjectSnapshotSerializer.CloneBomItems(b.BomItemsWithDrawings)
                : null;

            _bomItems.Clear();
            foreach (var it in PdmProjectSnapshotSerializer.CloneBomItems(b.Items))
            {
                _bomItems.Add(it);
            }

            if (ShowDrawingsToggle != null)
            {
                ShowDrawingsToggle.IsEnabled = _rawBomItems.Count > 0;
                ShowDrawingsToggle.IsChecked = b.ShowDrawingsChecked == true;
            }

            SetupBomDataGridColumns();
            ApplyBomColumnFiltersFromSnapshot(b.ColumnFilters);
            _bomItemsView?.Refresh();
            UpdateFilterUiState();
        }

        private void ApplyBomColumnFiltersFromSnapshot(IReadOnlyList<BomColumnFilterSnapshot> filters)
        {
            if (filters == null || filters.Count == 0)
            {
                return;
            }

            foreach (var f in filters)
            {
                if (!_columnFilters.TryGetValue(f.Key, out var st))
                {
                    continue;
                }

                var avail = st.AvailableValues;
                var set = new HashSet<string>(StringComparer.Ordinal);
                foreach (var s in f.SelectedValues ?? new List<string>())
                {
                    if (avail.Contains(s))
                    {
                        set.Add(s);
                    }
                }

                if (set.Count > 0)
                {
                    st.SelectedValues = set;
                }
                else if (avail.Count > 0)
                {
                    st.SelectedValues = new HashSet<string>(avail, StringComparer.Ordinal);
                }

                st.SearchText = f.SearchText ?? string.Empty;
            }
        }

        private void ApplyReferenceAuditSnapshotData(ReferenceAuditSnapshotData r)
        {
            if (r == null)
            {
                return;
            }

            if (ReferenceAuditMaxDepthTextBox != null)
            {
                ReferenceAuditMaxDepthTextBox.Text = r.ReferenceAuditMaxDepthText ?? "3";
            }

            if (ReferenceAuditAllDepthCheckBox != null)
            {
                ReferenceAuditAllDepthCheckBox.IsChecked = r.ReferenceAuditAllDepth;
            }

            if (ReferenceAuditIncludeDrawingsCheckBox != null)
            {
                ReferenceAuditIncludeDrawingsCheckBox.IsChecked = r.ReferenceAuditIncludeDrawings;
            }

            _auditRows.Clear();
            foreach (var row in PdmProjectSnapshotSerializer.CloneReferenceAuditRows(r.Rows))
            {
                _auditRows.Add(row);
            }

            if (ExportReferenceAuditButton != null)
            {
                ExportReferenceAuditButton.IsEnabled = _auditRows.Count > 0;
            }

            var maxIdx = (int)ReferenceAuditPreviewFilterMode.ExcludeDrawings;
            var idx = Math.Max(0, Math.Min(maxIdx, r.PreviewFilterComboIndex));
            if (ReferenceAuditPreviewFilterComboBox != null
                && ReferenceAuditPreviewFilterComboBox.Items.Count > idx)
            {
                ReferenceAuditPreviewFilterComboBox.SelectedIndex = idx;
            }

            _auditRowsView?.Refresh();
            UpdateAuditPreviewCountLabel();
        }
    }
}
