using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace PDMTools
{
    public partial class ColumnSettingsWindow : Window
    {
        // 確定後可讀取此清單（有序）
        public List<string> SelectedVariables { get; private set; } = new List<string>();

        public ColumnSettingsWindow(
            IReadOnlyList<string> allVariables,
            IReadOnlyList<string> currentSelected)
        {
            InitializeComponent();
            PopulateList(allVariables, currentSelected);
            UpdateCount();
        }

        // ── 建立 CheckBox 清單 ───────────────────────────────────────────

        private void PopulateList(
            IReadOnlyList<string> allVariables,
            IReadOnlyList<string> currentSelected)
        {
            VariablePanel.Children.Clear();
            var selectedSet = new HashSet<string>(
                currentSelected ?? new List<string>(),
                System.StringComparer.OrdinalIgnoreCase);

            foreach (var name in allVariables)
            {
                var cb = new CheckBox
                {
                    Content = name,
                    IsChecked = selectedSet.Contains(name),
                    Margin = new Thickness(2, 3, 2, 3),
                    FontSize = 12
                };
                cb.Checked   += (s, e) => UpdateCount();
                cb.Unchecked += (s, e) => UpdateCount();
                VariablePanel.Children.Add(cb);
            }
        }

        private void UpdateCount()
        {
            var total   = VariablePanel.Children.OfType<CheckBox>().Count();
            var checked_ = VariablePanel.Children.OfType<CheckBox>().Count(c => c.IsChecked == true);
            CountText.Text = $"已選 {checked_} / {total} 個";
        }

        // ── 按鈕事件 ─────────────────────────────────────────────────────

        private void SelectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var cb in VariablePanel.Children.OfType<CheckBox>())
                cb.IsChecked = true;
        }

        private void SelectNone_Click(object sender, RoutedEventArgs e)
        {
            foreach (var cb in VariablePanel.Children.OfType<CheckBox>())
                cb.IsChecked = false;
        }

        private void OK_Click(object sender, RoutedEventArgs e)
        {
            SelectedVariables = VariablePanel.Children
                .OfType<CheckBox>()
                .Where(cb => cb.IsChecked == true)
                .Select(cb => cb.Content as string)
                .Where(n => !string.IsNullOrWhiteSpace(n))
                .ToList();

            DialogResult = true;
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}
