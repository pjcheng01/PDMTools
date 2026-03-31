using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using EPDM.Interop.epdm;
using PDMTools.Models;
using PDMTools.Services;

namespace PDMTools
{
    public partial class BatchWorkflowTransitionWindow : Window
    {
        private readonly IReadOnlyList<string> _paths;
        private readonly PdmBomExportService _exportService;
        private readonly PdmBatchWorkflowTransitionService _service = new PdmBatchWorkflowTransitionService();
        private readonly List<(BatchWorkflowTransitionGroup Group, ComboBox Combo)> _groupCombos = new List<(BatchWorkflowTransitionGroup, ComboBox)>();
        private BatchWorkflowTransitionAnalyzeResult _lastAnalyze;

        public BatchWorkflowTransitionWindow(IReadOnlyList<string> vaultFilePaths, PdmBomExportService exportService)
        {
            InitializeComponent();
            _paths = vaultFilePaths ?? Array.Empty<string>();
            _exportService = exportService ?? throw new ArgumentNullException(nameof(exportService));
            SummaryTextBlock.Text =
                $"依目前 DataGrid 篩選後可見之列（已依路徑去重），透過 PDM 批次轉狀態 API 分組並轉換。\n"
                + $"共 {_paths.Count} 個路徑；請先按「重新分析」載入各群組可執行之轉換。";
            Loaded += (_, __) => RunAnalyze();
        }

        private void AnalyzeButton_OnClick(object sender, RoutedEventArgs e)
        {
            RunAnalyze();
        }

        private void RunAnalyze()
        {
            GroupsPanel.Children.Clear();
            _groupCombos.Clear();
            SkippedTextBlock.Text = string.Empty;
            LogTextBox.Clear();
            _lastAnalyze = null;

            try
            {
                var vault = _exportService.GetVault();
                _lastAnalyze = _service.Analyze(_paths, vault);
            }
            catch (Exception ex)
            {
                AppendLog("分析失敗：" + ex.Message);
                MessageBox.Show(this, ex.Message, "錯誤", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (_lastAnalyze.SkippedPaths.Count > 0)
            {
                SkippedTextBlock.Text = "略過或失敗的路徑（前 5 筆）：\n"
                    + string.Join("\n", _lastAnalyze.SkippedPaths.Take(5));
                if (_lastAnalyze.SkippedPaths.Count > 5)
                {
                    SkippedTextBlock.Text += $"\n… 其餘 {_lastAnalyze.SkippedPaths.Count - 5} 筆";
                }
            }

            foreach (var g in _lastAnalyze.Groups)
            {
                var border = new Border
                {
                    BorderBrush = System.Windows.Media.Brushes.LightGray,
                    BorderThickness = new Thickness(1),
                    Margin = new Thickness(0, 0, 0, 10),
                    Padding = new Thickness(10, 8, 10, 8)
                };

                var grid = new Grid();
                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(220) });
                grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

                var title = new TextBlock
                {
                    TextWrapping = TextWrapping.Wrap,
                    FontWeight = FontWeights.SemiBold,
                    Text = "mbsWorkflowName=" + (g.WorkflowName ?? string.Empty)
                        + "\nmlStateID=" + g.StateId
                        + "; mbsStateName=" + (g.StateName ?? string.Empty)
                        + "\nFilePaths.Count=" + g.FilePaths.Count
                };
                Grid.SetRow(title, 0);
                Grid.SetColumnSpan(title, 2);
                grid.Children.Add(title);

                var lbl = new TextBlock
                {
                    Text = "EdmChangeStateTransitionInfo:",
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(0, 8, 8, 0)
                };
                Grid.SetRow(lbl, 1);
                Grid.SetColumn(lbl, 0);

                var combo = new ComboBox
                {
                    MinHeight = 28,
                    Margin = new Thickness(0, 8, 0, 0),
                    DisplayMemberPath = nameof(PdmTransitionOption.DisplayText)
                };
                foreach (var t in g.AvailableTransitions)
                {
                    combo.Items.Add(t);
                }

                if (combo.Items.Count > 0)
                {
                    combo.SelectedIndex = 0;
                }

                Grid.SetRow(combo, 1);
                Grid.SetColumn(combo, 1);
                grid.Children.Add(lbl);
                grid.Children.Add(combo);

                if (!string.IsNullOrWhiteSpace(g.AnalyzeError))
                {
                    grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                    var err = new TextBlock
                    {
                        Foreground = System.Windows.Media.Brushes.DarkRed,
                        TextWrapping = TextWrapping.Wrap,
                        Margin = new Thickness(0, 6, 0, 0),
                        Text = "分析錯誤：" + g.AnalyzeError
                    };
                    Grid.SetRow(err, 2);
                    Grid.SetColumnSpan(err, 2);
                    grid.Children.Add(err);
                }

                border.Child = grid;
                GroupsPanel.Children.Add(border);
                _groupCombos.Add((g, combo));
            }

            AppendLog($"分析完成：{_lastAnalyze.Groups.Count} 個狀態群組。");
        }

        private void ExecuteButton_OnClick(object sender, RoutedEventArgs e)
        {
            if (_lastAnalyze == null || _groupCombos.Count == 0)
            {
                MessageBox.Show(this, "請先完成分析。", "提醒", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // PDM 對話框須傳入有效父視窗 hwnd（不可為 0）。EnsureHandle 強制建立 HWND，避免尚未配置時 Handle 仍為 0。
            var helper = new WindowInteropHelper(this);
            helper.EnsureHandle();
            var hwnd = helper.Handle.ToInt32();
            if (hwnd == 0)
            {
                MessageBox.Show(this, "無法取得有效視窗控制代碼（hwnd），PDM Change State 對話框可能無法正常執行。請稍候再試。", "提醒", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            IEdmVault5 vault;
            try
            {
                vault = _exportService.GetVault();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "登入失敗", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            AppendLog("── 開始執行轉換 ──");
            AppendLog($"[診斷] hwnd={hwnd}（若為 0 則父視窗無效，PDM 對話框可能無法正確執行）");
            AppendLog($"[診斷] Thread ApartmentState={Thread.CurrentThread.GetApartmentState()}（PDM COM 須為 STA；若為 MTA 請勿在背景執行緒呼叫）");

            foreach (var (group, combo) in _groupCombos)
            {
                if (group.FilePaths.Count == 0)
                {
                    continue;
                }

                if (combo.SelectedItem is not PdmTransitionOption opt)
                {
                    AppendLog($"略過群組（未選轉換）：{group.StateName}");
                    continue;
                }

                if (group.AvailableTransitions.Count == 0)
                {
                    AppendLog($"略過群組（無可用轉換）：{group.StateName}");
                    continue;
                }

                try
                {
                    var exec = _service.ExecuteGroup(vault, group.FilePaths, opt, hwnd);
                    foreach (var row in exec.Rows)
                    {
                        AppendLog((row.Success ? "✓ " : "✗ ") + row.FullPath + " — " + row.Message);
                    }
                }
                catch (Exception ex)
                {
                    AppendLog("群組執行錯誤：" + group.StateName + " — " + ex.Message);
                }
            }

            AppendLog("── 執行結束 ──");
            MessageBox.Show(this, "已處理所有群組；詳情見下方記錄。", "完成", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void AppendLog(string line)
        {
            if (LogTextBox.Text.Length > 0)
            {
                LogTextBox.AppendText(Environment.NewLine);
            }

            LogTextBox.AppendText(line);
            LogTextBox.ScrollToEnd();
        }
    }
}
