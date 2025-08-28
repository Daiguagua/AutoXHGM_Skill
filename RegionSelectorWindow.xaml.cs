using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Shapes;

namespace AutoXHGM_Skill
{
    public partial class RegionSelectorWindow : Window
    {
        private readonly IntPtr _targetWindowHandle;
        private System.Windows.Point _startPoint;
        private bool _isSelecting;
        private DpiScale _dpi;
        public SelectionRegion SelectedRegion { get; private set; }

        public RegionSelectorWindow(IntPtr targetWindowHandle)
        {
            Debug.WriteLine("[RegionSelectorWindow] 构造函数调用");
            InitializeComponent();
            _targetWindowHandle = targetWindowHandle;

            // 获取主窗口的DPI缩放因子
            var mainWindow = Application.Current.MainWindow;
            if (mainWindow != null)
            {
                _dpi = VisualTreeHelper.GetDpi(mainWindow);
                Debug.WriteLine($"[RegionSelectorWindow] DPI缩放因子: X={_dpi.DpiScaleX}, Y={_dpi.DpiScaleY}");
            }
            else
            {
                _dpi = new DpiScale(1.0, 1.0);
            }

            SetForegroundWindow(targetWindowHandle);
            ShowWindow(targetWindowHandle, SW_RESTORE);
        }
        
        private void OverlayCanvas_MouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            // 右键取消选择
            DialogResult = false;
            Close();
        }
        private void OverlayCanvas_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            Debug.WriteLine("[RegionSelector] 鼠标左键按下");
            _startPoint = e.GetPosition(overlayCanvas);
            _isSelecting = true;
            selectionRect.Visibility = Visibility.Visible;

            Canvas.SetLeft(selectionRect, _startPoint.X);
            Canvas.SetTop(selectionRect, _startPoint.Y);
            selectionRect.Width = 0;
            selectionRect.Height = 0;
        }

        private void OverlayCanvas_MouseMove(object sender, MouseEventArgs e)
        {
            var currentPoint = e.GetPosition(overlayCanvas);

            if (_isSelecting)
            {
                double left = Math.Min(_startPoint.X, currentPoint.X);
                double top = Math.Min(_startPoint.Y, currentPoint.Y);
                double width = Math.Abs(currentPoint.X - _startPoint.X);
                double height = Math.Abs(currentPoint.Y - _startPoint.Y);

                Canvas.SetLeft(selectionRect, left);
                Canvas.SetTop(selectionRect, top);
                selectionRect.Width = width;
                selectionRect.Height = height;

                // 转换为屏幕坐标
                var screenStart = PointToScreen(_startPoint);
                var screenEnd = PointToScreen(currentPoint);

                infoText.Text = $"区域: {screenStart.X},{screenStart.Y} -> {screenEnd.X},{screenEnd.Y}\n" +
                               $"尺寸: {width}x{height}";
            }
            else
            {
                var screenPos = PointToScreen(currentPoint);
                infoText.Text = $"当前位置: {screenPos.X},{screenPos.Y}";
            }
        }

        private void OverlayCanvas_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {

            if (!_isSelecting) return;

            _isSelecting = false;

            var endPoint = e.GetPosition(overlayCanvas);

            // 转换为屏幕坐标
            var screenStart = PointToScreen(_startPoint);
            var screenEnd = PointToScreen(endPoint);

            // 使用与取色模式相同的客户区偏移计算方法
            var clientOffset = Win32PointHelper.GetClientTopLeft(_targetWindowHandle);

            // 获取窗口矩形
            Win32PointHelper.GetWindowRect(_targetWindowHandle, out var windowRect);

            // 计算相对于客户区的坐标（与取色模式保持一致）
            int x = (int)(screenStart.X - (windowRect.Left + clientOffset.X));
            int y = (int)(screenStart.Y - (windowRect.Top + clientOffset.Y));
            int width = (int)Math.Abs(screenEnd.X - screenStart.X);
            int height = (int)Math.Abs(screenEnd.Y - screenStart.Y);

            SelectedRegion = new SelectionRegion(x, y, width, height);
            DialogResult = true;
            // 在RegionSelectorWindow的OverlayCanvas_MouseLeftButtonUp方法中添加
            Debug.WriteLine($"窗口矩形: L={windowRect.Left}, T={windowRect.Top}, R={windowRect.Right}, B={windowRect.Bottom}");
            Debug.WriteLine($"客户区偏移: X={clientOffset.X}, Y={clientOffset.Y}");
            Debug.WriteLine($"屏幕起点: X={screenStart.X}, Y={screenStart.Y}");
            Debug.WriteLine($"屏幕终点: X={screenEnd.X}, Y={screenEnd.Y}");
            Debug.WriteLine($"计算出的区域: X={x}, Y={y}, W={width}, H={height}");
            Close();
        }
        // 辅助方法：将逻辑点转换为屏幕点（考虑DPI）
        private System.Windows.Point PointToScreen(System.Windows.Point point, DpiScale dpi)
        {
            var screenPoint = this.PointToScreen(point);
            return new System.Windows.Point(screenPoint.X / dpi.DpiScaleX, screenPoint.Y / dpi.DpiScaleY);
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                DialogResult = false;
                Close();
            }
            base.OnKeyDown(e);
        }
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        private const int SW_RESTORE = 9;
    }

    // 重命名类以避免冲突
    public class SelectionRegion
    {
        public int X { get; set; }
        public int Y { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }

        public SelectionRegion(int x, int y, int width, int height)
        {
            X = x;
            Y = y;
            Width = width;
            Height = height;
        }
    }
}