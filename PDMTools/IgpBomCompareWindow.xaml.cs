using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Threading;
using Microsoft.Win32;
using PDMTools.Models;
using PDMTools.Services;

namespace PDMTools
{
    public partial class IgpBomCompareWindow : Window
    {
        private const string Stage1Prefix = "S1|";
        private const string Stage2Prefix = "S2|";

        private readonly IgpBomCompareResult _result;
        private readonly List<IgpBomStage1DiffRow> _stage1All;
        private readonly List<IgpBomStage2DiffRow> _stage2All;
        private readonly ICollectionView _stage1View;
        private readonly ICollectionView _stage2View;

        private readonly Dictionary<string, IgpColumnFilterState> _stage1Filters =
            new Dictionary<string, IgpColumnFilterState>(StringComparer.Ordinal);

        private readonly Dictionary<string, IgpColumnFilterState> _stage2Filters =
            new Dictionary<string, IgpColumnFilterState>(StringComparer.Ordinal);

        public IgpBomCompareWindow(IgpBomCompareResult result)
        {
            InitializeComponent();
            _result = result ?? new IgpBomCompareResult();

            _stage1All = (_result.Stage1Rows ?? new List<IgpBomStage1DiffRow>()).ToList();
            _stage2All = (_result.Stage2Rows ?? new List<IgpBomStage2DiffRow>()).ToList();

            _stage1View = CollectionViewSource.GetDefaultView(_stage1All);
            _stage1View.Filter = FilterStage1;

            _stage2View = CollectionViewSource.GetDefaultView(_stage2All);
            _stage2View.Filter = FilterStage2;

            var headerStyle = (Style)FindResource("IgpCompareDataGridColumnHeaderStyle");

            SetupStage1Columns(headerStyle);
            SetupStage2Columns(headerStyle);

            Stage1Grid.ItemsSource = _stage1View;
            Stage2Grid.ItemsSource = _stage2View;

            RebuildAllFilterOptions();
            UpdateFilterUiStage1();
            UpdateFilterUiStage2();

            var s1OnlyPdm = _stage1All.Count(r => string.Equals(r.DiffType, "OnlyPDM", StringComparison.OrdinalIgnoreCase));
            var s1OnlyIgp = _stage1All.Count(r => string.Equals(r.DiffType, "OnlyiGP", StringComparison.OrdinalIgnoreCase));
            var s1Mismatch = _stage1All.Count(r => string.Equals(r.DiffType, "Mismatch", StringComparison.OrdinalIgnoreCase));
            var s1Match = _stage1All.Count(r => string.Equals(r.DiffType, "Match", StringComparison.OrdinalIgnoreCase));

            var s2OnlyPdm = _stage2All.Count(r => string.Equals(r.DiffType, "OnlyPDMEdge", StringComparison.OrdinalIgnoreCase));
            var s2OnlyIgp = _stage2All.Count(r => string.Equals(r.DiffType, "OnlyiGPEdge", StringComparison.OrdinalIgnoreCase));
            var s2Mismatch = _stage2All.Count(r => string.Equals(r.DiffType, "Mismatch", StringComparison.OrdinalIgnoreCase));
            var s2Match = _stage2All.Count(r => string.Equals(r.DiffType, "Match", StringComparison.OrdinalIgnoreCase));

            SummaryText.Text =
                $"iGP：{_result.IgpSourcePath}\n"
                + $"深度口徑：Depth<= {_result.MaxDepth}（工程圖列已排除）\n"
                + $"比對時間（本機）：{_result.ComparedAtLocal:yyyy-MM-dd HH:mm:ss}\n"
                + $"解析：iGP Nodes={_result.IgpNodeCount} / Edges={_result.IgpEdgeCount}；PDM Nodes={_result.PdmNodeCount} / Edges={_result.PdmEdgeCount}\n"
                + $"Stage1：OnlyPDM={s1OnlyPdm}、OnlyiGP={s1OnlyIgp}、Mismatch={s1Mismatch}、一致={s1Match}\n"
                + $"Stage2：OnlyPDMEdge={s2OnlyPdm}、OnlyiGPEdge={s2OnlyIgp}、Mismatch={s2Mismatch}、一致={s2Match}"
                + (string.IsNullOrWhiteSpace(_result.Notes) ? string.Empty : "\n" + _result.Notes);

            UpdateDisplayedStats();
        }

        private void SaveResultButton_OnClick(object sender, RoutedEventArgs e)
        {
            var dlg = new SaveFileDialog
            {
                Title = "儲存 iGP BOM 比對結果",
                Filter = "Excel (*.xlsx)|*.xlsx|JSON (*.json)|*.json",
                DefaultExt = ".xlsx",
                FilterIndex = 1,
                AddExtension = true,
                FileName = BuildDefaultExportFileName()
            };

            if (dlg.ShowDialog(this) != true)
            {
                return;
            }

            try
            {
                IgpBomCompareResultExporter.Save(_result, dlg.FileName);
                MessageBox.Show(this, $"已儲存：\n{dlg.FileName}", "完成", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "儲存失敗", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private string BuildDefaultExportFileName()
        {
            var baseName = "IgpBomCompare";
            try
            {
                var p = _result?.IgpSourcePath;
                if (!string.IsNullOrWhiteSpace(p))
                {
                    baseName = Path.GetFileNameWithoutExtension(p);
                    foreach (var c in Path.GetInvalidFileNameChars())
                    {
                        baseName = baseName.Replace(c, '_');
                    }

                    if (baseName.Length > 120)
                    {
                        baseName = baseName.Substring(0, 120);
                    }

                    if (string.IsNullOrWhiteSpace(baseName))
                    {
                        baseName = "IgpBomCompare";
                    }
                }
            }
            catch
            {
                // 使用預設檔名
            }

            return $"{baseName}_compare_{DateTime.Now:yyyyMMdd_HHmmss}";
        }

        private void MainTabControl_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!IsLoaded)
            {
                return;
            }

            UpdateDisplayedStats();
        }

        /// <summary>依目前分頁與欄篩選後的 <see cref="ICollectionView"/> 更新狀態列（四類 + 總筆數）。</summary>
        private void UpdateDisplayedStats()
        {
            if (DisplayStatsTextBlock == null || MainTabControl == null)
            {
                return;
            }

            if (MainTabControl.SelectedIndex == 1)
            {
                var rows = _stage2View?.Cast<object>().OfType<IgpBomStage2DiffRow>().ToList()
                           ?? new List<IgpBomStage2DiffRow>();
                CountStage2(rows, out var onlyPdm, out var onlyIgp, out var mismatch, out var match, out var other);
                var total = rows.Count;
                DisplayStatsTextBlock.Text =
                    "「Stage2」父子結構｜目前顯示（已套用欄篩選）："
                    + $"{total} 筆；僅 PDM {onlyPdm}、僅 iGP {onlyIgp}、數量不一致 {mismatch}、一致 {match}"
                    + (other > 0 ? $"、其他 {other}" : string.Empty)
                    + "。";
            }
            else
            {
                var rows = _stage1View?.Cast<object>().OfType<IgpBomStage1DiffRow>().ToList()
                           ?? new List<IgpBomStage1DiffRow>();
                CountStage1(rows, out var onlyPdm, out var onlyIgp, out var mismatch, out var match, out var other);
                var total = rows.Count;
                DisplayStatsTextBlock.Text =
                    "「Stage1」料號彙總｜目前顯示（已套用欄篩選）："
                    + $"{total} 筆；僅 PDM {onlyPdm}、僅 iGP {onlyIgp}、數量不一致 {mismatch}、一致 {match}"
                    + (other > 0 ? $"、其他 {other}" : string.Empty)
                    + "。";
            }
        }

        private static void CountStage1(
            List<IgpBomStage1DiffRow> rows,
            out int onlyPdm,
            out int onlyIgp,
            out int mismatch,
            out int match,
            out int other)
        {
            onlyPdm = 0;
            onlyIgp = 0;
            mismatch = 0;
            match = 0;
            other = 0;
            foreach (var r in rows)
            {
                var t = r.DiffType ?? string.Empty;
                if (string.Equals(t, "OnlyPDM", StringComparison.OrdinalIgnoreCase))
                {
                    onlyPdm++;
                }
                else if (string.Equals(t, "OnlyiGP", StringComparison.OrdinalIgnoreCase))
                {
                    onlyIgp++;
                }
                else if (string.Equals(t, "Mismatch", StringComparison.OrdinalIgnoreCase))
                {
                    mismatch++;
                }
                else if (string.Equals(t, "Match", StringComparison.OrdinalIgnoreCase))
                {
                    match++;
                }
                else
                {
                    other++;
                }
            }
        }

        private static void CountStage2(
            List<IgpBomStage2DiffRow> rows,
            out int onlyPdm,
            out int onlyIgp,
            out int mismatch,
            out int match,
            out int other)
        {
            onlyPdm = 0;
            onlyIgp = 0;
            mismatch = 0;
            match = 0;
            other = 0;
            foreach (var r in rows)
            {
                var t = r.DiffType ?? string.Empty;
                if (string.Equals(t, "OnlyPDMEdge", StringComparison.OrdinalIgnoreCase))
                {
                    onlyPdm++;
                }
                else if (string.Equals(t, "OnlyiGPEdge", StringComparison.OrdinalIgnoreCase))
                {
                    onlyIgp++;
                }
                else if (string.Equals(t, "Mismatch", StringComparison.OrdinalIgnoreCase))
                {
                    mismatch++;
                }
                else if (string.Equals(t, "Match", StringComparison.OrdinalIgnoreCase))
                {
                    match++;
                }
                else
                {
                    other++;
                }
            }
        }

        private void SetupStage1Columns(Style headerStyle)
        {
            Stage1Grid.Columns.Clear();
            _stage1Filters.Clear();

            void Add(string key, string title, Binding binding, double minWidth, DataGridLength width, Func<IgpBomStage1DiffRow, string> getter)
            {
                var hi = new CompareColumnHeaderInfo
                {
                    Key = Stage1Prefix + key,
                    Title = title
                };
                var col = new DataGridTextColumn
                {
                    Header = hi,
                    Binding = binding,
                    MinWidth = minWidth,
                    Width = width,
                    HeaderStyle = headerStyle
                };
                Stage1Grid.Columns.Add(col);
                _stage1Filters[key] = new IgpColumnFilterState
                {
                    Key = key,
                    ValueGetter = o => getter((IgpBomStage1DiffRow)o)
                };
            }

            Add("PartNumber", "料號", new Binding("PartNumber") { Mode = BindingMode.OneWay }, 80, new DataGridLength(1, DataGridLengthUnitType.Star), r => r.PartNumber ?? string.Empty);
            Add("RootPdm", "根(PDM)", new Binding("RootPdm") { Mode = BindingMode.OneWay }, 100, 160, r => r.RootPdm ?? string.Empty);
            Add("RootIgp", "根(iGP)", new Binding("RootIgp") { Mode = BindingMode.OneWay }, 100, 160, r => r.RootIgp ?? string.Empty);
            Add("PdmQty", "PDM 總量", new Binding("PdmQty") { Mode = BindingMode.OneWay }, 72, 110, r => FormatDecimal(r.PdmQty));
            Add("IgpQty", "iGP 總量", new Binding("IgpQty") { Mode = BindingMode.OneWay }, 72, 110, r => FormatDecimal(r.IgpQty));
            Add("Diff", "差異", new Binding("Diff") { Mode = BindingMode.OneWay }, 72, 110, r => FormatDecimal(r.Diff));
            Add("DiffType", "類型", new Binding("DiffType") { Mode = BindingMode.OneWay }, 72, 120, r => r.DiffType ?? string.Empty);
        }

        private void SetupStage2Columns(Style headerStyle)
        {
            Stage2Grid.Columns.Clear();
            _stage2Filters.Clear();

            void Add(string key, string title, Binding binding, double minWidth, DataGridLength width, Func<IgpBomStage2DiffRow, string> getter)
            {
                var hi = new CompareColumnHeaderInfo
                {
                    Key = Stage2Prefix + key,
                    Title = title
                };
                var col = new DataGridTextColumn
                {
                    Header = hi,
                    Binding = binding,
                    MinWidth = minWidth,
                    Width = width,
                    HeaderStyle = headerStyle
                };
                Stage2Grid.Columns.Add(col);
                _stage2Filters[key] = new IgpColumnFilterState
                {
                    Key = key,
                    ValueGetter = o => getter((IgpBomStage2DiffRow)o)
                };
            }

            Add("RootPdm", "根(PDM)", new Binding("RootPdm") { Mode = BindingMode.OneWay }, 100, 160, r => r.RootPdm ?? string.Empty);
            Add("RootIgp", "根(iGP)", new Binding("RootIgp") { Mode = BindingMode.OneWay }, 100, 160, r => r.RootIgp ?? string.Empty);
            Add("ParentPartNumber", "父料號", new Binding("ParentPartNumber") { Mode = BindingMode.OneWay }, 80, 160, r => r.ParentPartNumber ?? string.Empty);
            Add("ChildPartNumber", "子料號", new Binding("ChildPartNumber") { Mode = BindingMode.OneWay }, 80, 160, r => r.ChildPartNumber ?? string.Empty);
            Add("PdmQty", "PDM 用量", new Binding("PdmQty") { Mode = BindingMode.OneWay }, 72, 110, r => FormatDecimal(r.PdmQty));
            Add("IgpQty", "iGP 用量", new Binding("IgpQty") { Mode = BindingMode.OneWay }, 72, 110, r => FormatDecimal(r.IgpQty));
            Add("Diff", "差異", new Binding("Diff") { Mode = BindingMode.OneWay }, 72, 110, r => FormatDecimal(r.Diff));
            Add("DiffType", "類型", new Binding("DiffType") { Mode = BindingMode.OneWay }, 72, 140, r => r.DiffType ?? string.Empty);
        }

        private static string FormatDecimal(decimal d)
        {
            return d.ToString(CultureInfo.InvariantCulture);
        }

        private bool FilterStage1(object obj)
        {
            if (!(obj is IgpBomStage1DiffRow row))
            {
                return false;
            }

            foreach (var state in _stage1Filters.Values)
            {
                if (state.AvailableValues.Count == 0)
                {
                    continue;
                }

                var selectedCount = state.SelectedValues.Count;
                if (selectedCount == 0 || selectedCount == state.AvailableValues.Count)
                {
                    continue;
                }

                var current = NormalizeFilterValue(state.ValueGetter(row));
                if (!state.SelectedValues.Contains(current))
                {
                    return false;
                }
            }

            return true;
        }

        private bool FilterStage2(object obj)
        {
            if (!(obj is IgpBomStage2DiffRow row))
            {
                return false;
            }

            foreach (var state in _stage2Filters.Values)
            {
                if (state.AvailableValues.Count == 0)
                {
                    continue;
                }

                var selectedCount = state.SelectedValues.Count;
                if (selectedCount == 0 || selectedCount == state.AvailableValues.Count)
                {
                    continue;
                }

                var current = NormalizeFilterValue(state.ValueGetter(row));
                if (!state.SelectedValues.Contains(current))
                {
                    return false;
                }
            }

            return true;
        }

        private void RebuildAllFilterOptions()
        {
            foreach (var s in _stage1Filters.Values)
            {
                RebuildFilterValueOptions(s, _stage1All);
            }

            foreach (var s in _stage2Filters.Values)
            {
                RebuildFilterValueOptions(s, _stage2All);
            }
        }

        private static void RebuildFilterValueOptions(IgpColumnFilterState state, IEnumerable source)
        {
            var oldValues = state.AvailableValues ?? new List<string>();
            var wasUnfiltered =
                oldValues.Count == 0 ||
                state.SelectedValues.Count == 0 ||
                state.SelectedValues.Count == oldValues.Count;

            var values = source
                .Cast<object>()
                .Select(x => NormalizeFilterValue(state.ValueGetter(x)))
                .Distinct(StringComparer.Ordinal)
                .ToList();

            var hasBlank = values.Any(x => string.Equals(x, IgpColumnFilterState.BlankDisplayText, StringComparison.Ordinal));
            var orderedValues = values
                .Where(x => !string.Equals(x, IgpColumnFilterState.BlankDisplayText, StringComparison.Ordinal))
                .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (hasBlank)
            {
                orderedValues.Insert(0, IgpColumnFilterState.BlankDisplayText);
            }

            state.AvailableValues = orderedValues;

            if (wasUnfiltered)
            {
                state.SelectedValues = new HashSet<string>(orderedValues, StringComparer.Ordinal);
                return;
            }

            state.SelectedValues.IntersectWith(orderedValues);
            if (state.SelectedValues.Count == 0)
            {
                state.SelectedValues = new HashSet<string>(orderedValues, StringComparer.Ordinal);
            }
        }

        private static string GetFilterFieldKey(string fullHeaderKey)
        {
            var p = (fullHeaderKey ?? string.Empty).IndexOf('|');
            return p >= 0 && p + 1 < fullHeaderKey.Length ? fullHeaderKey.Substring(p + 1) : fullHeaderKey;
        }

        private static string NormalizeFilterValue(string raw)
        {
            return string.IsNullOrWhiteSpace(raw) ? IgpColumnFilterState.BlankDisplayText : raw.Trim();
        }

        private void IgpCompareColumnFilterButton_OnClick(object sender, RoutedEventArgs e)
        {
            if (!(sender is Button button) || !(button.Tag is string fullKey))
            {
                return;
            }

            Dictionary<string, IgpColumnFilterState> dict;
            ICollectionView view;
            Action updateUi;
            IEnumerable sourceRows;

            if (fullKey.StartsWith(Stage1Prefix, StringComparison.Ordinal))
            {
                dict = _stage1Filters;
                view = _stage1View;
                updateUi = UpdateFilterUiStage1;
                sourceRows = _stage1All;
            }
            else if (fullKey.StartsWith(Stage2Prefix, StringComparison.Ordinal))
            {
                dict = _stage2Filters;
                view = _stage2View;
                updateUi = UpdateFilterUiStage2;
                sourceRows = _stage2All;
            }
            else
            {
                return;
            }

            var shortKey = GetFilterFieldKey(fullKey);
            if (!dict.TryGetValue(shortKey, out var state))
            {
                return;
            }

            EnsureSelectionInitialized(state);
            RebuildFilterValueOptions(state, sourceRows);

            var menu = BuildFilterContextMenu(state, view, updateUi);
            button.ContextMenu = menu;
            menu.PlacementTarget = button;
            menu.Placement = PlacementMode.Bottom;
            menu.IsOpen = true;
            e.Handled = true;
        }

        private static void EnsureSelectionInitialized(IgpColumnFilterState state)
        {
            if (state.SelectedValues.Count > 0)
            {
                return;
            }

            state.SelectedValues = new HashSet<string>(state.AvailableValues, StringComparer.Ordinal);
        }

        private ContextMenu BuildFilterContextMenu(IgpColumnFilterState state, ICollectionView view, Action updateUi)
        {
            var menu = new ContextMenu();
            var valueMenuItems = new List<(string Value, MenuItem Item)>();
            TextBox searchBox = null;

            void RefocusSearchBox()
            {
                if (searchBox == null)
                {
                    return;
                }

                Dispatcher.BeginInvoke(new Action(() =>
                {
                    searchBox.Focus();
                    searchBox.Select(searchBox.Text.Length, 0);
                }), DispatcherPriority.Background);
            }

            void ApplyFiltersAndRefresh()
            {
                view?.Refresh();
                updateUi?.Invoke();
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

                ApplyFiltersAndRefresh();
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

                ApplyFiltersAndRefresh();
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
                    {
                        state.SelectedValues.Add(value);
                    }
                    else
                    {
                        state.SelectedValues.Remove(value);
                    }

                    ApplyFiltersAndRefresh();
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

                    ApplyFiltersAndRefresh();
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

        private static bool IsSearchMatch(string value, string keyword)
        {
            if (string.IsNullOrWhiteSpace(keyword))
            {
                return true;
            }

            var pattern = keyword.Trim();
            var hasWildcard = pattern.IndexOf('*') >= 0 || pattern.IndexOf('?') >= 0;
            if (!hasWildcard)
            {
                return value.IndexOf(pattern, StringComparison.OrdinalIgnoreCase) >= 0;
            }

            var regexPattern = "^" + Regex.Escape(pattern)
                .Replace(@"\*", ".*")
                .Replace(@"\?", ".") + "$";

            return Regex.IsMatch(value ?? string.Empty, regexPattern, RegexOptions.IgnoreCase);
        }

        private void UpdateFilterUiStage1()
        {
            foreach (var col in Stage1Grid.Columns)
            {
                if (!(col.Header is CompareColumnHeaderInfo header))
                {
                    continue;
                }

                if (!_stage1Filters.TryGetValue(GetFilterFieldKey(header.Key), out var state))
                {
                    header.IsFiltered = false;
                    continue;
                }

                header.IsFiltered =
                    state.AvailableValues.Count > 0 &&
                    state.SelectedValues.Count > 0 &&
                    state.SelectedValues.Count < state.AvailableValues.Count;
            }

            UpdateDisplayedStats();
        }

        private void UpdateFilterUiStage2()
        {
            foreach (var col in Stage2Grid.Columns)
            {
                if (!(col.Header is CompareColumnHeaderInfo header))
                {
                    continue;
                }

                if (!_stage2Filters.TryGetValue(GetFilterFieldKey(header.Key), out var state))
                {
                    header.IsFiltered = false;
                    continue;
                }

                header.IsFiltered =
                    state.AvailableValues.Count > 0 &&
                    state.SelectedValues.Count > 0 &&
                    state.SelectedValues.Count < state.AvailableValues.Count;
            }

            UpdateDisplayedStats();
        }

        private sealed class CompareColumnHeaderInfo : INotifyPropertyChanged
        {
            private bool _isFiltered;

            public string Key { get; set; } = string.Empty;
            public string Title { get; set; } = string.Empty;

            public bool IsFiltered
            {
                get => _isFiltered;
                set
                {
                    if (_isFiltered == value)
                    {
                        return;
                    }

                    _isFiltered = value;
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsFiltered)));
                }
            }

            public event PropertyChangedEventHandler PropertyChanged;
        }

        private sealed class IgpColumnFilterState
        {
            public const string BlankDisplayText = "(空白)";

            public string Key { get; set; } = string.Empty;
            public Func<object, string> ValueGetter { get; set; } = _ => string.Empty;
            public List<string> AvailableValues { get; set; } = new List<string>();
            public HashSet<string> SelectedValues { get; set; } = new HashSet<string>(StringComparer.Ordinal);
            public string SearchText { get; set; } = string.Empty;
        }
    }
}
