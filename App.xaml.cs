using System.Runtime.InteropServices;
using System.Windows;

namespace AutoXHGM_Skill
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            // 启用每显示器 DPI 感知
            SetProcessDpiAwarenessContext((int)DpiAwarenessContext.PerMonitorAwareV2);
            DispatcherUnhandledException += Application_DispatcherUnhandledException;
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
