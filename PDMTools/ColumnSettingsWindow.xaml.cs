using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace PDMTools
{
    public partial class ColumnSettingsWindow : Window
    {
        /// <summary>確定後：使用者勾選的固定欄位清單。</summary>
        public List<string> SelectedFixedColumns { get; private set; } = new List<string>();

        /// <summary>確定後：使用者勾選的資料卡變數清單。</summary>
        public List<string> SelectedVariables { get; private set; } = new List<string>();

        /// <summary>確定後：視窗中顯示的所有資料卡變數（含未勾選），供存入 KnownVariables。</summary>
        public List<string> AllShownVariables { get; private set; } = new List<string>();

        // ── 建構子 ───────────────────────────────────────────────────────

        public ColumnSettingsWindow(
            IReadOnlyList<string> allFixedColumns,
            IReadOnlyList<string> selectedFixedColumns,
            IReadOnlyList<string> allCardVariables,
            IReadOnlyList<string> selectedCardVariables)
        {
            InitializeComponent();
            PopulateSection(FixedColumnPanel,  allFixedColumns,   selectedFixedColumns);
            PopulateSection(VariablePanel,     allCardVariables,  selectedCardVariables);
            UpdateCount();
        }

        // ── 建立 CheckBox 清單 ───────────────────────────────────────────

        private void PopulateSection(
            StackPanel panel,
            IReadOnlyList<string> all,
            IReadOnlyList<string> selected)
        {
            panel.Children.Clear();
            var selectedSet = new System.Collections.Generic.HashSet<string>(
                selected ?? new List<string>(),
                System.StringComparer.OrdinalIgnoreCase);

            foreach (var name in all ?? new List<string>())
            {
                var cb = new CheckBox
                {
                    Content   = name,
                    IsChecked = selectedSet.Contains(name),
                    Margin    = new Thickness(2, 3, 2, 3),
                    FontSize  = 12
                };
                cb.Checked   += (s, e) => UpdateCount();
                cb.Unchecked += (s, e) => UpdateCount();
                panel.Children.Add(cb);
            }
        }

        private void UpdateCount()
        {
            var allCbs     = AllCheckBoxes().ToList();
            var checkedCnt = allCbs.Count(c => c.IsChecked == true);
            CountText.Text = $"已選 {checkedCnt} / {allCbs.Count} 個";
        }

        private IEnumerable<CheckBox> AllCheckBoxes() =>
            FixedColumnPanel.Children.OfType<CheckBox>()
                .Concat(VariablePanel.Children.OfType<CheckBox>());

        // ── 按鈕事件 ─────────────────────────────────────────────────────

        private void SelectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var cb in AllCheckBoxes()) cb.IsChecked = true;
        }

        private void SelectNone_Click(object sender, RoutedEventArgs e)
        {
            foreach (var cb in AllCheckBoxes()) cb.IsChecked = false;
        }

        private void OK_Click(object sender, RoutedEventArgs e)
        {
            SelectedFixedColumns = FixedColumnPanel.Children
                .OfType<CheckBox>()
                .Where(cb => cb.IsChecked == true)
                .Select(cb => cb.Content as string)
                .Where(n => !string.IsNullOrWhiteSpace(n))
                .ToList();

            var allCardCbs = VariablePanel.Children.OfType<CheckBox>().ToList();

            SelectedVariables = allCardCbs
                .Where(cb => cb.IsChecked == true)
                .Select(cb => cb.Content as string)
                .Where(n => !string.IsNullOrWhiteSpace(n))
                .ToList();

            AllShownVariables = allCardCbs
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
