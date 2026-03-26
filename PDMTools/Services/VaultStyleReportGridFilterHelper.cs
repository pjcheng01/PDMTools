using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Threading;
using PDMTools.Models;

namespace PDMTools.Services
{
    /// <summary>
    /// 與主視窗 Vault BOM 相同的漏斗＋勾選清單欄篩選，用於查詢結果 DataGrid。
    /// </summary>
    public sealed class VaultStyleReportGridFilterHelper
    {
        private readonly Window _window;
        private readonly DataGrid _grid;
        private readonly TextBlock _countBlock;
        private readonly Dictionary<string, ReportColumnFilterState> _columnFilters =
            new Dictionary<string, ReportColumnFilterState>(StringComparer.OrdinalIgnoreCase);

        public VaultStyleReportGridFilterHelper(Window window, DataGrid grid, TextBlock countBlock)
        {
            _window = window ?? throw new ArgumentNullException(nameof(window));
            _grid = grid ?? throw new ArgumentNullException(nameof(grid));
            _countBlock = countBlock;
        }

        public IReadOnlyDictionary<string, ReportColumnFilterState> ColumnFilters => _columnFilters;

        public void SetupCustomPropertyTemplateColumns()
        {
            _grid.Columns.Clear();
            _columnFilters.Clear();
            _grid.IsReadOnly = false;

            _grid.Columns.Add(new DataGridCheckBoxColumn
            {
                Header = "套用",
                Width = 52,
                Binding = new Binding("IsSelectedForApply") { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged }
            });

            void Add(string key, string title, string path, double width, DataGridLengthUnitType unit, Func<object, string> getter, double minWidth = 56)
            {
                var hi = new ReportColumnHeaderInfo { Key = key, Title = title };
                var col = new DataGridTextColumn
                {
                    Header = hi,
                    Binding = new Binding(path) { Mode = BindingMode.OneWay },
                    SortMemberPath = path,
                    MinWidth = minWidth,
                    IsReadOnly = true
                };
                col.Width = unit == DataGridLengthUnitType.Star
                    ? new DataGridLength(width, DataGridLengthUnitType.Star)
                    : new DataGridLength(width);
                _grid.Columns.Add(col);
                _columnFilters[key] = new ReportColumnFilterState { Key = key, ValueGetter = getter };
            }

            Add("FileName", "檔名", "FileName", 180, DataGridLengthUnitType.Pixel,
                o => (o as CustomPropertyTemplateRow)?.FileName ?? string.Empty);
            Add("Extension", "副檔名", "Extension", 72, DataGridLengthUnitType.Pixel,
                o => (o as CustomPropertyTemplateRow)?.Extension ?? string.Empty);
            Add("TemplateName", "範本名稱", "TemplateName", 180, DataGridLengthUnitType.Pixel,
                o => (o as CustomPropertyTemplateRow)?.TemplateName ?? string.Empty);
            Add("TemplatePath", "範本完整路徑", "TemplatePath", 1, DataGridLengthUnitType.Star,
                o => (o as CustomPropertyTemplateRow)?.TemplatePath ?? string.Empty, 120);
            Add("WeldmentTemplateName", "焊件範本名", "WeldmentTemplateName", 140, DataGridLengthUnitType.Pixel,
                o => (o as CustomPropertyTemplateRow)?.WeldmentTemplateName ?? string.Empty);
            Add("WeldmentTemplatePath", "焊件範本路徑", "WeldmentTemplatePath", 200, DataGridLengthUnitType.Pixel,
                o => (o as CustomPropertyTemplateRow)?.WeldmentTemplatePath ?? string.Empty);
            Add("Status", "狀態", "Status", 64, DataGridLengthUnitType.Pixel,
                o => (o as CustomPropertyTemplateRow)?.Status ?? string.Empty);
            Add("Message", "訊息", "Message", 200, DataGridLengthUnitType.Pixel,
                o => (o as CustomPropertyTemplateRow)?.Message ?? string.Empty);

            var applyCol = new DataGridTextColumn
            {
                Header = "套用結果",
                Binding = new Binding("LastApplyMessage") { Mode = BindingMode.OneWay },
                SortMemberPath = "LastApplyMessage",
                Width = new DataGridLength(220),
                MinWidth = 120,
                IsReadOnly = true
            };
            _grid.Columns.Add(applyCol);
        }

        public void SetupDrawingSheetFormatColumns()
        {
            _grid.Columns.Clear();
            _columnFilters.Clear();
            _grid.IsReadOnly = false;

            _grid.Columns.Add(new DataGridCheckBoxColumn
            {
                Header = "套用",
                Width = 52,
                Binding = new Binding("IsSelectedForApply") { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged }
            });

            void Add(string key, string title, string path, double width, DataGridLengthUnitType unit, Func<object, string> getter, double minWidth = 48)
            {
                var hi = new ReportColumnHeaderInfo { Key = key, Title = title };
                var col = new DataGridTextColumn
                {
                    Header = hi,
                    Binding = new Binding(path) { Mode = BindingMode.OneWay },
                    SortMemberPath = path,
                    MinWidth = minWidth,
                    IsReadOnly = true
                };
                col.Width = unit == DataGridLengthUnitType.Star
                    ? new DataGridLength(width, DataGridLengthUnitType.Star)
                    : new DataGridLength(width);
                _grid.Columns.Add(col);
                _columnFilters[key] = new ReportColumnFilterState { Key = key, ValueGetter = getter };
            }

            Add("FileName", "工程圖檔名", "FileName", 160, DataGridLengthUnitType.Pixel,
                o => (o as DrawingSheetFormatRow)?.FileName ?? string.Empty);
            Add("SheetName", "圖頁名稱", "SheetName", 120, DataGridLengthUnitType.Pixel,
                o => (o as DrawingSheetFormatRow)?.SheetName ?? string.Empty);
            Add("FormatFileName", "格式檔名", "FormatFileName", 160, DataGridLengthUnitType.Pixel,
                o => (o as DrawingSheetFormatRow)?.FormatFileName ?? string.Empty);
            Add("FormatPath", "圖頁格式完整路徑", "FormatPath", 1, DataGridLengthUnitType.Star,
                o => (o as DrawingSheetFormatRow)?.FormatPath ?? string.Empty, 120);
            Add("Status", "狀態", "Status", 56, DataGridLengthUnitType.Pixel,
                o => (o as DrawingSheetFormatRow)?.Status ?? string.Empty);
            Add("Message", "訊息", "Message", 200, DataGridLengthUnitType.Pixel,
                o => (o as DrawingSheetFormatRow)?.Message ?? string.Empty);

            var applyCol = new DataGridTextColumn
            {
                Header = "套用結果",
                Binding = new Binding("LastApplyMessage") { Mode = BindingMode.OneWay },
                SortMemberPath = "LastApplyMessage",
                Width = new DataGridLength(220),
                MinWidth = 120,
                IsReadOnly = true
            };
            _grid.Columns.Add(applyCol);
        }

        public void SetItemsSource(IEnumerable rows)
        {
            _grid.ItemsSource = rows;
            var view = CollectionViewSource.GetDefaultView(_grid.ItemsSource);
            if (view != null)
                view.Filter = FilterRow;

            foreach (var state in _columnFilters.Values)
                RebuildFilterValueOptions(state);

            foreach (var state in _columnFilters.Values)
                EnsureSelectionInitialized(state);

            view?.Refresh();
            UpdateFilterUiState();
            UpdateCount(view);

            // 查詢結果載入後依「欄首＋儲存格」取較大寬度，讓表格內容較易完整可見
            _grid.Dispatcher.BeginInvoke((Action)AutoFitColumnsToContent, DispatcherPriority.ContextIdle);
        }

        /// <summary>
        /// 欄寬取欄首與儲存格量測的較大值（欄首含漏斗按鈕範本）。
        /// 列數不多時暫關虛擬化以量到所有列；列數過多時僅量測目前已產生的儲存格，避免卡頓。
        /// </summary>
        private void AutoFitColumnsToContent()
        {
            if (_grid == null || _grid.Columns.Count == 0)
                return;

            const int maxRowsFullMeasure = 500;

            try
            {
                var items = _grid.ItemsSource;
                var rowCount = GetEnumerableCount(items, maxRowsFullMeasure);
                var useFullMeasure = rowCount >= 0 && rowCount <= maxRowsFullMeasure;

                var savedVirtualization = _grid.EnableRowVirtualization;
                if (useFullMeasure)
                    _grid.EnableRowVirtualization = false;

                try
                {
                    _grid.UpdateLayout();

                    var n = _grid.Columns.Count;
                    var headerWidths = new double[n];
                    for (var i = 0; i < n; i++)
                        _grid.Columns[i].Width = new DataGridLength(1, DataGridLengthUnitType.SizeToHeader);
                    _grid.UpdateLayout();
                    for (var i = 0; i < n; i++)
                        headerWidths[i] = _grid.Columns[i].ActualWidth;

                    for (var i = 0; i < n; i++)
                        _grid.Columns[i].Width = new DataGridLength(1, DataGridLengthUnitType.SizeToCells);
                    _grid.UpdateLayout();

                    var finalW = new double[n];
                    for (var i = 0; i < n; i++)
                    {
                        var col = _grid.Columns[i];
                        var w = Math.Max(headerWidths[i], col.ActualWidth);
                        if (w <= 0 || double.IsNaN(w) || double.IsInfinity(w))
                            w = col.MinWidth > 0 ? col.MinWidth : 80;
                        else
                            w = Math.Max(w, col.MinWidth);
                        finalW[i] = w;
                    }

                    // 最後一欄改為 * 佔滿剩餘寬度，避免右側出現 WPF 預留的空白填補區（像多餘欄）
                    for (var i = 0; i < n - 1; i++)
                        _grid.Columns[i].Width = new DataGridLength(finalW[i]);

                    var lastCol = _grid.Columns[n - 1];
                    lastCol.MinWidth = Math.Max(lastCol.MinWidth, finalW[n - 1]);
                    lastCol.Width = new DataGridLength(1, DataGridLengthUnitType.Star);
                }
                finally
                {
                    if (useFullMeasure)
                        _grid.EnableRowVirtualization = savedVirtualization;
                }
            }
            catch
            {
                // 量測失敗時略過，不影響查詢結果
            }
        }

        /// <param name="maxFullMeasure">列數超過此值時不再為虛擬化關閉做精確計數，傳回值會大於此值。</param>
        private static int GetEnumerableCount(object itemsSource, int maxFullMeasure)
        {
            if (itemsSource == null)
                return 0;
            if (itemsSource is ICollection col)
                return col.Count;
            if (itemsSource is IEnumerable en)
            {
                var n = 0;
                foreach (var _ in en)
                {
                    n++;
                    if (n > maxFullMeasure)
                        return n;
                }
                return n;
            }
            return 0;
        }

        public void ClearItemsSource()
        {
            _grid.ItemsSource = null;
            foreach (var s in _columnFilters.Values)
            {
                s.AvailableValues.Clear();
                s.SelectedValues.Clear();
            }

            if (_countBlock != null)
                _countBlock.Text = string.Empty;
            UpdateFilterUiState();
        }

        public void ClearAllFilters()
        {
            foreach (var state in _columnFilters.Values)
                state.SelectedValues.Clear();

            var view = CollectionViewSource.GetDefaultView(_grid.ItemsSource);
            view?.Refresh();
            UpdateFilterUiState();
            UpdateCount(view);
        }

        public void OnColumnFilterButtonClick(object sender, RoutedEventArgs e)
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

        private bool FilterRow(object obj)
        {
            foreach (var state in _columnFilters.Values)
            {
                if (state.AvailableValues.Count == 0)
                    continue;

                var selectedCount = state.SelectedValues.Count;
                if (selectedCount == 0 || selectedCount == state.AvailableValues.Count)
                    continue;

                var currentValue = NormalizeFilterValue(state.ValueGetter(obj));
                if (!state.SelectedValues.Contains(currentValue))
                    return false;
            }

            return true;
        }

        private void ApplyFiltersAndRefreshStatus()
        {
            var view = CollectionViewSource.GetDefaultView(_grid.ItemsSource);
            view?.Refresh();
            UpdateFilterUiState();
            UpdateCount(view);
        }

        private void UpdateCount(ICollectionView view)
        {
            if (_countBlock == null)
                return;
            if (view?.SourceCollection == null)
            {
                _countBlock.Text = string.Empty;
                return;
            }

            var total = 0;
            foreach (var _ in view.SourceCollection)
                total++;

            var visible = 0;
            foreach (var _ in view)
                visible++;

            _countBlock.Text = $"顯示 {visible} / {total} 列";
        }

        private void UpdateFilterUiState()
        {
            foreach (var col in _grid.Columns)
            {
                if (!(col.Header is ReportColumnHeaderInfo header))
                    continue;
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
        }

        private static void EnsureSelectionInitialized(ReportColumnFilterState state)
        {
            if (state.SelectedValues.Count > 0)
                return;
            state.SelectedValues = new HashSet<string>(state.AvailableValues, StringComparer.Ordinal);
        }

        private void RebuildFilterValueOptions(ReportColumnFilterState state)
        {
            var items = _grid.ItemsSource as IEnumerable;
            if (items == null)
            {
                state.AvailableValues = new List<string>();
                return;
            }

            var oldValues = state.AvailableValues ?? new List<string>();
            var wasUnfiltered =
                oldValues.Count == 0 ||
                state.SelectedValues.Count == 0 ||
                state.SelectedValues.Count == oldValues.Count;

            var values = items.Cast<object>()
                .Select(x => NormalizeFilterValue(state.ValueGetter(x)))
                .Distinct(StringComparer.Ordinal)
                .ToList();

            var hasBlank = values.Any(x => string.Equals(x, ReportColumnFilterState.BlankDisplayText, StringComparison.Ordinal));
            var orderedValues = values
                .Where(x => !string.Equals(x, ReportColumnFilterState.BlankDisplayText, StringComparison.Ordinal))
                .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (hasBlank)
                orderedValues.Insert(0, ReportColumnFilterState.BlankDisplayText);

            state.AvailableValues = orderedValues;

            if (wasUnfiltered)
            {
                state.SelectedValues = new HashSet<string>(orderedValues, StringComparer.Ordinal);
                return;
            }

            state.SelectedValues.IntersectWith(orderedValues);
            if (state.SelectedValues.Count == 0)
                state.SelectedValues = new HashSet<string>(orderedValues, StringComparer.Ordinal);
        }

        private ContextMenu BuildFilterContextMenu(ReportColumnFilterState state)
        {
            var menu = new ContextMenu();
            var valueMenuItems = new List<(string Value, MenuItem Item)>();
            TextBox searchBox = null;

            void RefocusSearchBox()
            {
                if (searchBox == null)
                    return;
                _window.Dispatcher.BeginInvoke(new Action(() =>
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
                    item.IsChecked = true;
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
                    searchBox.Text = string.Empty;
                foreach (var (_, item) in valueMenuItems)
                    item.IsChecked = false;
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
                    item.Visibility = isMatch ? Visibility.Visible : Visibility.Collapsed;
                    if (isMatch)
                        matchedValues.Add(value);
                }

                if (syncSelectionToSearch && !string.IsNullOrEmpty(keyword))
                {
                    state.SelectedValues = new HashSet<string>(matchedValues, StringComparer.Ordinal);
                    foreach (var (value, item) in valueMenuItems)
                        item.IsChecked = state.SelectedValues.Contains(value);
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

        private static string NormalizeFilterValue(string raw)
        {
            return string.IsNullOrWhiteSpace(raw) ? ReportColumnFilterState.BlankDisplayText : raw.Trim();
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
    }
}
