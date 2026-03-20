using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace PDMTools
{
    public partial class ColumnSettingsWindow : Window
    {
        // ── 資料來源 ────────────────────────────────────────────────────
        private readonly List<string> _allAvailable = new List<string>(); // 完整可用清單
        private readonly ObservableCollection<string> _available  = new ObservableCollection<string>(); // 篩選後顯示
        private readonly ObservableCollection<string> _selected   = new ObservableCollection<string>(); // 已選（有序）

        // ── 輸出屬性 ─────────────────────────────────────────────────────

        /// <summary>確定後：使用者選定的欄位清單（依左側順序）。</summary>
        public List<string> OrderedSelectedItems { get; private set; } = new List<string>();

        /// <summary>確定後：右側清單中所有的變數名稱（供存入 KnownVariables）。</summary>
        public List<string> AllShownAvailableItems { get; private set; } = new List<string>();

        // ── 建構子 ───────────────────────────────────────────────────────

        /// <param name="allItems">所有可選欄位（固定欄 + 資料卡變數）</param>
        /// <param name="currentSelected">目前已選欄位（依順序）</param>
        public ColumnSettingsWindow(
            IReadOnlyList<string> allItems,
            IReadOnlyList<string> currentSelected)
        {
            InitializeComponent();

            var selectedSet = new HashSet<string>(
                currentSelected ?? new List<string>(),
                StringComparer.OrdinalIgnoreCase);

            // 已選清單：保留原有順序
            foreach (var item in currentSelected ?? new List<string>())
                _selected.Add(item);

            // 可用清單：從 allItems 中排除已選的，依字母排序
            foreach (var item in allItems ?? new List<string>())
            {
                if (!selectedSet.Contains(item))
                    _allAvailable.Add(item);
            }
            _allAvailable.Sort(StringComparer.OrdinalIgnoreCase);

            SelectedList.ItemsSource  = _selected;
            AvailableList.ItemsSource = _available;
            RefreshFilter();
            UpdateCount();
        }

        // ── 搜尋篩選 ─────────────────────────────────────────────────────

        private void SearchBox_TextChanged(object sender, TextChangedEventArgs e) => RefreshFilter();

        private void RefreshFilter()
        {
            var keyword = SearchBox.Text?.Trim() ?? string.Empty;
            _available.Clear();
            foreach (var item in _allAvailable)
            {
                if (string.IsNullOrEmpty(keyword) ||
                    item.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    _available.Add(item);
                }
            }
        }

        // ── 穿梭按鈕 ─────────────────────────────────────────────────────

        /// <summary>把右側選取的項目加入左側。</summary>
        private void AddSelected_Click(object sender, RoutedEventArgs e)
        {
            var items = AvailableList.SelectedItems.Cast<string>().ToList();
            if (items.Count == 0) return;
            MoveToSelected(items);
        }

        /// <summary>把右側全部項目加入左側。</summary>
        private void AddAll_Click(object sender, RoutedEventArgs e)
        {
            MoveToSelected(_available.ToList());
        }

        /// <summary>把左側選取的項目移回右側。</summary>
        private void RemoveSelected_Click(object sender, RoutedEventArgs e)
        {
            var items = SelectedList.SelectedItems.Cast<string>().ToList();
            if (items.Count == 0) return;
            MoveToAvailable(items);
        }

        /// <summary>把左側全部項目移回右側。</summary>
        private void RemoveAll_Click(object sender, RoutedEventArgs e)
        {
            MoveToAvailable(_selected.ToList());
        }

        private void MoveToSelected(IEnumerable<string> items)
        {
            foreach (var item in items)
            {
                _selected.Add(item);
                _allAvailable.Remove(item);
            }
            RefreshFilter();
            UpdateCount();
        }

        private void MoveToAvailable(IEnumerable<string> items)
        {
            foreach (var item in items)
            {
                _selected.Remove(item);
                if (!_allAvailable.Contains(item))
                {
                    // 插入到正確的排序位置
                    var idx = _allAvailable.BinarySearch(item, StringComparer.OrdinalIgnoreCase);
                    _allAvailable.Insert(idx < 0 ? ~idx : idx, item);
                }
            }
            RefreshFilter();
            UpdateCount();
        }

        // ── 雙擊快速移動 ────────────────────────────────────────────────

        private void AvailableList_DoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (AvailableList.SelectedItem is string item)
                MoveToSelected(new[] { item });
        }

        private void SelectedList_DoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (SelectedList.SelectedItem is string item)
                MoveToAvailable(new[] { item });
        }

        // ── 排序（↑ / ↓）────────────────────────────────────────────────

        private void MoveUp_Click(object sender, RoutedEventArgs e)
        {
            var indices = SelectedList.SelectedItems
                .Cast<string>()
                .Select(s => _selected.IndexOf(s))
                .OrderBy(i => i)
                .ToList();

            if (indices.Count == 0 || indices[0] == 0) return;

            foreach (var idx in indices)
                _selected.Move(idx, idx - 1);

            RestoreSelection(indices.Select(i => i - 1));
        }

        private void MoveDown_Click(object sender, RoutedEventArgs e)
        {
            var indices = SelectedList.SelectedItems
                .Cast<string>()
                .Select(s => _selected.IndexOf(s))
                .OrderByDescending(i => i)
                .ToList();

            if (indices.Count == 0 || indices[0] == _selected.Count - 1) return;

            foreach (var idx in indices)
                _selected.Move(idx, idx + 1);

            RestoreSelection(indices.Select(i => i + 1));
        }

        private void RestoreSelection(IEnumerable<int> indices)
        {
            SelectedList.SelectedItems.Clear();
            foreach (var idx in indices)
            {
                if (idx >= 0 && idx < _selected.Count)
                    SelectedList.SelectedItems.Add(_selected[idx]);
            }
        }

        // ── 計數更新 ─────────────────────────────────────────────────────

        private void UpdateCount()
        {
            var total = _selected.Count + _allAvailable.Count;
            CountText.Text = $"已選 {_selected.Count} / {total} 個欄位";
        }

        // ── 確定 / 取消 ──────────────────────────────────────────────────

        private void OK_Click(object sender, RoutedEventArgs e)
        {
            OrderedSelectedItems  = _selected.ToList();
            AllShownAvailableItems = _allAvailable.ToList();
            DialogResult = true;
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}
