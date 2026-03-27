using System;
using System.ComponentModel;
using System.Threading;
using System.Windows;
using PDMTools.Models;

namespace PDMTools
{
    public partial class BomGrabProgressWindow : Window
    {
        private readonly CancellationTokenSource _cancellationTokenSource;
        private bool _closeFromOwner;

        public BomGrabProgressWindow(CancellationTokenSource cancellationTokenSource)
        {
            InitializeComponent();
            _cancellationTokenSource = cancellationTokenSource ?? throw new ArgumentNullException(nameof(cancellationTokenSource));
        }

        public void UpdateProgress(ProgressInfo p)
        {
            if (p == null)
                return;

            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.Invoke(() => UpdateProgress(p));
                return;
            }

            var pct = p.Percentage;
            if (pct < 0)
                pct = 0;
            if (pct > 100)
                pct = 100;
            BatchProgressBar.Value = pct;
            PercentTextBlock.Text = pct + "%";
            MessageTextBlock.Text = p.Message ?? string.Empty;
        }

        /// <summary>由主視窗在作業正常結束後關閉，避免 Closing 時再送出 Cancel。</summary>
        public void CloseFromOwner()
        {
            _closeFromOwner = true;
            try
            {
                Close();
            }
            catch
            {
                // ignore
            }
        }

        private void Interrupt_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                _cancellationTokenSource.Cancel();
            }
            catch
            {
                // ignore
            }
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            if (!_closeFromOwner)
            {
                try
                {
                    _cancellationTokenSource.Cancel();
                }
                catch
                {
                    // ignore
                }
            }

            base.OnClosing(e);
        }
    }
}
