using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using Application = System.Windows.Application;
using Point = System.Drawing.Point;
using Size = System.Drawing.Size;
using WMBrushes = System.Windows.Media.Brushes;
// 修改1：在类顶部添加别名（可选但推荐）
using WMColor = System.Windows.Media.Color;
using WMImage = System.Windows.Controls.Image;
namespace AutoXHGM_Skill
{
    public class ProfessionalColorPicker
    {
        // 在类顶部添加 DPI 相关字段
        private double _dpiScaleX = 1.0;
        private double _dpiScaleY = 1.0;
        private readonly IntPtr _targetWindow;
        private Bitmap _screenCapture;
        private Window _overlayWindow;
        private WMImage _magnifierImage;
        private Point _selectedPosition;
        private System.Drawing.Color _selectedColor; // 明确使用System.Drawing
        private bool _isSelecting;

        public System.Drawing.Color SelectedColor => _selectedColor;
        public Point SelectedPosition => _selectedPosition;

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool AdjustWindowRectEx(ref RECT lpRect, uint dwStyle, bool bMenu, uint dwExStyle);

        //private Point ConvertToClientCoordinates(IntPtr hwnd, int screenX, int screenY)
        //{
        //    // 应用DPI缩放
        //    double logicalX = screenX / _dpiScaleX;
        //    double logicalY = screenY / _dpiScaleY;

        //    // 对齐到像素中心
        //    // 确保坐标是整数（对准像素中心）
        //    screenX = (int)Math.Round(logicalX);
        //    screenY = (int)Math.Round(logicalY);
        //    // 获取窗口矩形（物理像素）
        //    GetWindowRect(hwnd, out RECT windowRect);

        //    // 获取客户端矩形（物理像素）
        //    GetClientRect(hwnd, out RECT clientRect);

        //    // 计算边框大小
        //    int borderWidth = (windowRect.Right - windowRect.Left - clientRect.Right) / 2;
        //    int titleBarHeight = (windowRect.Bottom - windowRect.Top - clientRect.Bottom) - borderWidth;

        //    // 调试输出
        //    Debug.WriteLine($"窗口矩形: L{windowRect.Left} T{windowRect.Top} R{windowRect.Right} B{windowRect.Bottom}");
        //    Debug.WriteLine($"客户端矩形: W{clientRect.Right} H{clientRect.Bottom}");
        //    Debug.WriteLine($"边框: W{borderWidth} H{titleBarHeight}");

        //    // 转换为客户端坐标（物理像素）
        //    Point clientPoint = new Point(screenX, screenY);
        //    ScreenToClient(hwnd, ref clientPoint);
        //    return clientPoint;
        //}
        private Point ConvertToClientCoordinates(IntPtr hwnd, int screenX, int screenY)
        {
            Point clientPoint = new Point(screenX, screenY);
            ScreenToClient(hwnd, ref clientPoint);
            return clientPoint;
        }
        // 添加所需的Win32 API
        [DllImport("user32.dll")]
        private static extern bool ScreenToClient(IntPtr hWnd, ref Point lpPoint);

        private RECT _clientRect;
        [DllImport("user32.dll")]
        private static extern bool ClientToScreen(IntPtr hWnd, ref Point lpPoint);
        public ProfessionalColorPicker(IntPtr targetWindow)
        {
            _targetWindow = targetWindow;

            // 激活并显示目标窗口
            SetForegroundWindow(targetWindow);
            ShowWindow(targetWindow, SW_RESTORE);

            // 获取客户区位置（屏幕坐标）
            GetClientRect(targetWindow, out _clientRect);
            Point topLeft = new Point(_clientRect.Left, _clientRect.Top);
            ClientToScreen(targetWindow, ref topLeft);
            _clientRect = new RECT
            {
                Left = topLeft.X,
                Top = topLeft.Y,
                Right = topLeft.X + (_clientRect.Right - _clientRect.Left),
                Bottom = topLeft.Y + (_clientRect.Bottom - _clientRect.Top)
            };

            // 最小化主窗口
            Application.Current.Dispatcher.Invoke(() =>
            {
                if (Application.Current.MainWindow != null)
                {
                    Application.Current.MainWindow.WindowState = WindowState.Minimized;
                }
            });
        }
        // 添加 Win32 API 声明
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        private const int SW_RESTORE = 9;
        public bool? ShowDialog()
        {
            // 捕获屏幕
            CaptureScreen();

            // 创建全屏覆盖窗口
            CreateOverlayWindow();

            // 进入选择模式
            _isSelecting = true;

            // 显示模态窗口
            _overlayWindow.ShowDialog();

            return _isSelecting;
        }

        private void CaptureScreen()
        {
            // 获取整个屏幕的尺寸
            int screenWidth = Screen.PrimaryScreen.Bounds.Width;
            int screenHeight = Screen.PrimaryScreen.Bounds.Height;

            // 创建屏幕截图
            _screenCapture = new Bitmap(screenWidth, screenHeight);

            using (var g = Graphics.FromImage(_screenCapture))
            {
                g.CopyFromScreen(0, 0, 0, 0, new Size(screenWidth, screenHeight));
            }
        }

        private void CreateOverlayWindow()
        {
            _overlayWindow = new Window
            {
                WindowStyle = WindowStyle.None,
                WindowState = WindowState.Maximized,
                Topmost = true,
                Background = WMBrushes.Transparent,
                AllowsTransparency = true,
                ShowInTaskbar = false
            };

            // 使用 Canvas 作为容器
            var canvas = new Canvas();
            canvas.Background = new SolidColorBrush(WMColor.FromArgb(80, 0, 0, 0));

            // 添加鼠标事件处理
            canvas.MouseMove += OverlayMouseMove;
            canvas.MouseDown += OverlayMouseDown;
            canvas.KeyDown += OverlayKeyDown;

            // 添加放大镜
            _magnifierImage = new WMImage
            {
                Width = 200,
                Height = 200,
                Visibility = Visibility.Collapsed
            };

            canvas.Children.Add(_magnifierImage);
            _overlayWindow.Content = canvas;

            // 设置窗口位置
            var helper = new WindowInteropHelper(_overlayWindow);
            helper.EnsureHandle();

            // 启用键盘事件
            _overlayWindow.Focusable = true;
            _overlayWindow.Focus();
        }

        private void OverlayMouseMove(object sender, System.Windows.Input.MouseEventArgs e)
        {
            // 获取鼠标位置（屏幕逻辑坐标）
            var position = e.GetPosition(_overlayWindow);

            // 转换为屏幕物理坐标
            System.Windows.Point wpfScreenPoint = _overlayWindow.PointToScreen(position); // 明确使用 System.Windows.Point
            int screenX = (int)wpfScreenPoint.X;
            int screenY = (int)wpfScreenPoint.Y;

            // 显示放大镜
            UpdateMagnifier(screenX, screenY);

            // 更新放大镜位置（使用逻辑坐标）
            double displayX = position.X + 20;
            double displayY = position.Y + 20;

            // 确保放大镜不超出屏幕
            displayX = Math.Min(displayX, SystemParameters.PrimaryScreenWidth - 220);
            displayY = Math.Min(displayY, SystemParameters.PrimaryScreenHeight - 220);

            Canvas.SetLeft(_magnifierImage, displayX);
            Canvas.SetTop(_magnifierImage, displayY);
            _magnifierImage.Visibility = Visibility.Visible;
            Debug.WriteLine($"鼠标位置: 逻辑({position.X},{position.Y}) 物理({screenX},{screenY})");
        }

        private void UpdateMagnifier(int centerX, int centerY)
        {
            const int magnifierSize = 100; // 放大镜大小
            const int zoomFactor = 8; // 放大倍数

            // 创建放大镜图像
            var bitmap = new Bitmap(magnifierSize, magnifierSize);
            // 确保中心点对准像素中心
            int srcX = centerX;
            int srcY = centerY;
            // 计算源图像区域（以目标像素为中心）
            int srcStartX = srcX - magnifierSize / (2 * zoomFactor);
            int srcStartY = srcY - magnifierSize / (2 * zoomFactor);
            int srcWidth = magnifierSize / zoomFactor;
            int srcHeight = magnifierSize / zoomFactor;

            using (var g = Graphics.FromImage(bitmap))
            {
                // 绘制放大区域
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.NearestNeighbor;
                g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.Half;
                g.DrawImage(_screenCapture,
                    new Rectangle(0, 0, magnifierSize, magnifierSize),
                    new Rectangle(srcStartX, srcStartY, srcWidth, srcHeight),
                    GraphicsUnit.Pixel);

                // 绘制精确的十字线（对准像素中心）
                using (var pen = new System.Drawing.Pen(System.Drawing.Color.Red, 1))
                {
                    int center = magnifierSize / 2;
                    g.DrawLine(pen, center, 0, center, magnifierSize);
                    g.DrawLine(pen, 0, center, magnifierSize, center);
                }

                // 添加坐标和颜色信息
                using (var font = new Font("Arial", 8))
                using (var brush = new SolidBrush(System.Drawing.Color.White))
                {
                    string info = $"{srcX},{srcY}\n";
                    g.DrawString(info, font, brush, 5, 5);

                    // 获取并显示当前像素颜色
                    var pixelColor = _screenCapture.GetPixel(srcX, srcY);
                    string colorInfo = $"R:{pixelColor.R} G:{pixelColor.G} B:{pixelColor.B}";
                    g.DrawString(colorInfo, font, brush, 5, 25);
                }
            }

            // 更新WPF图像控件
            _magnifierImage.Source = BitmapToImageSource(bitmap);
        }

        private ImageSource BitmapToImageSource(Bitmap bitmap)
        {
            var hBitmap = bitmap.GetHbitmap();
            try
            {
                return System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                    hBitmap, IntPtr.Zero, Int32Rect.Empty,
                    System.Windows.Media.Imaging.BitmapSizeOptions.FromEmptyOptions());
            }
            finally
            {
                DeleteObject(hBitmap);
            }
        }

        private void OverlayMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
            {
                // 获取鼠标位置（屏幕逻辑坐标）
                var position = e.GetPosition(_overlayWindow);

                // 转换为屏幕物理坐标
                System.Windows.Point wpfScreenPoint = _overlayWindow.PointToScreen(position); // 明确使用 System.Windows.Point
                int screenX = (int)wpfScreenPoint.X;
                int screenY = (int)wpfScreenPoint.Y;
                // 应用DPI缩放
                double logicalX = screenX / _dpiScaleX;
                double logicalY = screenY / _dpiScaleY;

                // 对齐到像素中心
                // 确保坐标是整数（对准像素中心）
                screenX = (int)Math.Round(logicalX);
                screenY = (int)Math.Round(logicalY);
                // 转换为目标窗口的客户端坐标
                _selectedPosition = ConvertToClientCoordinates(_targetWindow, screenX, screenY);

                // 获取颜色（使用物理坐标）
                _selectedColor = _screenCapture.GetPixel(screenX, screenY);

                // 调试输出
                Debug.WriteLine($"选中位置: 窗口({_selectedPosition.X},{_selectedPosition.Y}) 屏幕({screenX},{screenY})");

                // 关闭窗口
                _isSelecting = true;
                _overlayWindow.DialogResult = true;
            }
            else if (e.ChangedButton == MouseButton.Right)
            {
                CancelSelection();
            }
        }

        // 添加获取客户区大小的 Win32 API
        [DllImport("user32.dll")]
        private static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);

        private int GetClientWidth(IntPtr hwnd)
        {
            GetClientRect(hwnd, out RECT rect);
            return rect.Right - rect.Left;
        }
        private int GetClientHeight(IntPtr hwnd)
        {
            GetClientRect(hwnd, out RECT rect);
            return rect.Bottom - rect.Top;
        }

        private void OverlayKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                // ESC 取消
                CancelSelection();
            }
        }
        private void CancelSelection()
        {
            _isSelecting = false;
            _overlayWindow.DialogResult = false;
        }
        #region Win32 API
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("gdi32.dll")]
        private static extern bool DeleteObject(IntPtr hObject);

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }
        #endregion
    }
}
