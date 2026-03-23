using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using PDMTools.Models;

namespace PDMTools
{
    public partial class ColumnSettingsWindow : Window
    {
        // ── 資料來源 ────────────────────────────────────────────────────
        private readonly List<string> _allAvailable = new List<string>(); // 完整可用清單
        private readonly ObservableCollection<string> _available  = new ObservableCollection<string>(); // 篩選後顯示
        private readonly ObservableCollection<string> _selected   = new ObservableCollection<string>(); // 已選（有序）
        private readonly List<string> _allItems = new List<string>();
        private readonly List<ColumnPresetProfile> _profiles = new List<ColumnPresetProfile>();
        private string _activeProfileName = string.Empty;

        // ── 輸出屬性 ─────────────────────────────────────────────────────

        /// <summary>確定後：使用者選定的欄位清單（依左側順序）。</summary>
        public List<string> OrderedSelectedItems { get; private set; } = new List<string>();

        /// <summary>確定後：右側清單中所有的變數名稱（供存入 KnownVariables）。</summary>
        public List<string> AllShownAvailableItems { get; private set; } = new List<string>();
        public List<ColumnPresetProfile> UpdatedProfiles { get; private set; } = new List<ColumnPresetProfile>();
        public string LastUsedProfileName { get; private set; } = string.Empty;

        // ── 建構子 ───────────────────────────────────────────────────────

        /// <param name="allItems">所有可選欄位（固定欄 + 資料卡變數）</param>
        /// <param name="currentSelected">目前已選欄位（依順序）</param>
        public ColumnSettingsWindow(
            IReadOnlyList<string> allItems,
            IReadOnlyList<string> currentSelected,
            IReadOnlyList<ColumnPresetProfile> profiles,
            string lastUsedProfileName)
        {
            InitializeComponent();
            _allItems.AddRange((allItems ?? new List<string>()).ToList());
            _profiles.AddRange((profiles ?? new List<ColumnPresetProfile>())
                .Where(p => p != null && !string.IsNullOrWhiteSpace(p.Name))
                .Select(p => new ColumnPresetProfile
                {
                    Name = p.Name,
                    SelectedItems = (p.SelectedItems ?? new List<string>()).ToList()
                }));

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
            ReloadProfileCombo();
            SelectProfile(lastUsedProfileName, keepNameInput: true);
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

        private void ProfileCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ProfileCombo.SelectedItem is string name)
            {
                _activeProfileName = name;
                ProfileNameBox.Text = name;
            }
        }

        private void SaveProfile_Click(object sender, RoutedEventArgs e)
        {
            var name = (ProfileNameBox.Text ?? string.Empty).Trim();
            if (string.IsNullOrEmpty(name))
            {
                MessageBox.Show(this, "請先輸入組合名稱。", "提醒", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var existing = _profiles.FirstOrDefault(p => string.Equals(p.Name, name, StringComparison.OrdinalIgnoreCase));
            if (existing == null)
            {
                _profiles.Add(new ColumnPresetProfile
                {
                    Name = name,
                    SelectedItems = _selected.ToList()
                });
            }
            else
            {
                existing.SelectedItems = _selected.ToList();
                existing.Name = name;
            }

            _activeProfileName = name;
            ReloadProfileCombo();
            SelectProfile(name, keepNameInput: true);
            UpdateCount();
        }

        private void ApplyProfile_Click(object sender, RoutedEventArgs e)
        {
            var name = ProfileCombo.SelectedItem as string ?? (ProfileNameBox.Text ?? string.Empty).Trim();
            if (string.IsNullOrEmpty(name))
            {
                MessageBox.Show(this, "請先選擇要套用的組合。", "提醒", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var profile = _profiles.FirstOrDefault(p => string.Equals(p.Name, name, StringComparison.OrdinalIgnoreCase));
            if (profile == null)
            {
                MessageBox.Show(this, "找不到指定的欄位組合。", "提醒", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            ApplyProfileToLists(profile);
            _activeProfileName = profile.Name;
            SelectProfile(profile.Name, keepNameInput: true);
            UpdateCount();
        }

        private void DeleteProfile_Click(object sender, RoutedEventArgs e)
        {
            var name = ProfileCombo.SelectedItem as string ?? (ProfileNameBox.Text ?? string.Empty).Trim();
            if (string.IsNullOrEmpty(name))
            {
                MessageBox.Show(this, "請先選擇要刪除的組合。", "提醒", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var profile = _profiles.FirstOrDefault(p => string.Equals(p.Name, name, StringComparison.OrdinalIgnoreCase));
            if (profile == null)
            {
                MessageBox.Show(this, "找不到指定的欄位組合。", "提醒", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            _profiles.Remove(profile);
            if (string.Equals(_activeProfileName, profile.Name, StringComparison.OrdinalIgnoreCase))
            {
                _activeProfileName = string.Empty;
            }

            ReloadProfileCombo();
            ProfileNameBox.Text = string.Empty;
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
            OrderedSelectedItems = _selected.ToList();

            // KnownVariables 需要涵蓋「已選 + 可用」全部，才能在下次啟動時正確判斷是否有新增變數
            AllShownAvailableItems = _selected.Concat(_allAvailable).ToList();
            UpdatedProfiles = _profiles
                .Where(p => !string.IsNullOrWhiteSpace(p.Name))
                .Select(p => new ColumnPresetProfile
                {
                    Name = p.Name.Trim(),
                    SelectedItems = (p.SelectedItems ?? new List<string>())
                        .Where(x => !string.IsNullOrWhiteSpace(x))
                        .ToList()
                })
                .ToList();
            LastUsedProfileName = _activeProfileName ?? string.Empty;

            DialogResult = true;
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }

        private void ReloadProfileCombo()
        {
            var names = _profiles
                .Select(p => p.Name)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(n => n, StringComparer.OrdinalIgnoreCase)
                .ToList();
            ProfileCombo.ItemsSource = names;
        }

        private void SelectProfile(string profileName, bool keepNameInput)
        {
            if (!string.IsNullOrWhiteSpace(profileName))
            {
                var hit = (ProfileCombo.ItemsSource as IEnumerable<string>)
                    ?.FirstOrDefault(x => string.Equals(x, profileName, StringComparison.OrdinalIgnoreCase));
                ProfileCombo.SelectedItem = hit;
                _activeProfileName = hit ?? string.Empty;
                if (keepNameInput && !string.IsNullOrWhiteSpace(hit))
                {
                    ProfileNameBox.Text = hit;
                }
            }
            else
            {
                ProfileCombo.SelectedItem = null;
                _activeProfileName = string.Empty;
                if (!keepNameInput)
                {
                    ProfileNameBox.Text = string.Empty;
                }
            }
        }

        private void ApplyProfileToLists(ColumnPresetProfile profile)
        {
            var selected = (profile?.SelectedItems ?? new List<string>())
                .Where(x => _allItems.Contains(x, StringComparer.OrdinalIgnoreCase))
                .ToList();

            _selected.Clear();
            foreach (var item in selected)
            {
                _selected.Add(item);
            }

            _allAvailable.Clear();
            foreach (var item in _allItems)
            {
                if (!selected.Contains(item, StringComparer.OrdinalIgnoreCase))
                {
                    _allAvailable.Add(item);
                }
            }
            _allAvailable.Sort(StringComparer.OrdinalIgnoreCase);
            RefreshFilter();
        }
    }
}
