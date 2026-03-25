using System;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Threading;

namespace PDMTools
{
    public partial class App : Application
    {
        private static readonly string StartupErrorLogPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "PDMTools",
            "startup-error.log");

        protected override void OnStartup(StartupEventArgs e)
        {
            AttachGlobalExceptionHandlers();
            base.OnStartup(e);

            try
            {
                var window = new MainWindow();
                MainWindow = window;
                window.Show();
            }
            catch (Exception ex)
            {
                ShowFatalError("啟動主視窗失敗", ex);
                Shutdown(-1);
            }
        }

        private void AttachGlobalExceptionHandlers()
        {
            DispatcherUnhandledException += (_, args) =>
            {
                ShowFatalError("UI 執行緒未處理例外", args.Exception);
                args.Handled = true;
                Shutdown(-1);
            };

            AppDomain.CurrentDomain.UnhandledException += (_, args) =>
            {
                var ex = args.ExceptionObject as Exception
                    ?? new Exception(args.ExceptionObject?.ToString() ?? "未知例外");
                ShowFatalError("AppDomain 未處理例外", ex);
            };
        }

        private static void ShowFatalError(string title, Exception ex)
        {
            var message = $"{title}\n\n{ex.Message}\n\n{ex}";
            TryAppendErrorLog(message);
            MessageBox.Show(
                message + $"\n\n錯誤記錄：{StartupErrorLogPath}",
                "PDMTools 啟動錯誤",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }

        private static void TryAppendErrorLog(string message)
        {
            try
            {
                var dir = Path.GetDirectoryName(StartupErrorLogPath);
                if (!string.IsNullOrWhiteSpace(dir))
                {
                    Directory.CreateDirectory(dir);
                }

                var content = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}\n\n";
                File.AppendAllText(StartupErrorLogPath, content, Encoding.UTF8);
            }
            catch
            {
                // ignore logging failures
            }
        }
    }
}
