using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Threading;

namespace Skill_Loop
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        [DllImport("user32.dll")]
        private static extern bool SetProcessDPIAware();
        private DispatcherTimer _cacheCleanupTimer;
        protected override void OnStartup(StartupEventArgs e)
        {

            try
            {
                // 启用每显示器 DPI 感知
                SetProcessDpiAwarenessContext((int)DpiAwarenessContext.PerMonitorAwareV2);
                DispatcherUnhandledException += Application_DispatcherUnhandledException;
                // 启用DPI感知
                SetProcessDPIAware();
            }
            catch (Exception ex)
            {
                // 如果API调用失败，记录错误但继续运行
                Debug.WriteLine($"启用DPI感知失败: {ex.Message}");
            }
            _cacheCleanupTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromMinutes(1),
                IsEnabled = true
            };

            _cacheCleanupTimer.Tick += (s, ev) => OcrCacheService.RemoveExpiredItems();
            _cacheCleanupTimer.Start();
            base.OnStartup(e);

        }
        protected override void OnExit(ExitEventArgs e)
        {
            _cacheCleanupTimer?.Stop();
            OcrCacheService.Clear();
            base.OnExit(e);
        }
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetProcessDpiAwarenessContext(int value);

        private enum DpiAwarenessContext
        {
            PerMonitorAwareV2 = 34
        }
        private void Application_DispatcherUnhandledException(object sender, System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
        {
            MessageBox.Show($"未处理异常: {e.Exception}", "错误",
                MessageBoxButton.OK, MessageBoxImage.Error);
            e.Handled = true;
        }
    }

}
