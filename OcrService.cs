using OpenCvSharp;
using Sdcb.PaddleInference;
using Sdcb.PaddleOCR;
using Sdcb.PaddleOCR.Models;
using Sdcb.PaddleOCR.Models.Local;
using System.Diagnostics;
using System.Drawing;
using System.Windows;
using System.Windows.Media;
using System.Windows.Controls;
using System.Windows.Shapes;
using System.Windows.Threading;

namespace AutoXHGM_Skill
{
    public static class OcrService
    {
        private static PaddleOcrAll _ocrEngine;
        private static bool _isInitialized = false;
        private static readonly object _lock = new object();

        public static void Initialize()
        {
            if (_isInitialized) return;

            lock (_lock)
            {
                if (_isInitialized) return;

                try
                {
                    // 使用中文V5模型
                    FullOcrModel model = LocalFullModels.ChineseV5;
                    //var paddleConfig = new PaddleConfig()
                    //{
                    //    EnableGpuMultiStream = true,
                    //    UseGpu=true,
                    //    CudnnEnabled=true,
                    //    OnnxEnabled=true,
                    //    MkldnnEnabled=true,
                    //};

                    _ocrEngine = new PaddleOcrAll(model, config =>
                    {
                        config.EnableGpuMultiStream = true;
                        config.UseGpu = true;
                        config.MkldnnEnabled = true;
                        config.CudnnEnabled = true;
                        config.OnnxEnabled = true;
                    })
                    {
                        AllowRotateDetection = false,
                        Enable180Classification = false
                    };

                    _isInitialized = true;
                    Debug.WriteLine("OCR引擎初始化成功");
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"OCR引擎初始化失败: {ex.Message}");
                    _isInitialized = false;
                }
            }
        }
        // OcrService.cs
        public static OcrResult RecognizeTextFromRegion(IntPtr windowHandle, SkillCondition condition)
        {
            try
            {
                // 获取窗口位置和客户区偏移
                Win32PointHelper.GetWindowRect(windowHandle, out var windowRect);
                var clientOffset = Win32PointHelper.GetClientTopLeft(windowHandle);

                // 计算OCR区域在屏幕上的绝对坐标
                int absoluteX = windowRect.Left + clientOffset.X + condition.OcrRegionX;
                int absoluteY = windowRect.Top + clientOffset.Y + condition.OcrRegionY;

                // 创建区域对象
                var region = new System.Drawing.Rectangle(
                    absoluteX, absoluteY,
                    condition.OcrRegionWidth, condition.OcrRegionHeight);

                // 绘制调试矩形
                DrawDebugRectangle(region, System.Drawing.Color.Blue, 3000);

                using (var bitmap = MainWindow.CaptureRegion(region))
                //using (var graphics = Graphics.FromImage(bitmap))
                {
                    //graphics.CopyFromScreen(region.Location, System.Drawing.Point.Empty, region.Size);

                    // 生成缓存键
                    string cacheKey = OcrCacheKeyGenerator.GenerateKey(windowHandle, condition);

                    // 调用OCR识别
                    string result = RecognizeText(bitmap, cacheKey);

                    // 计算相似度
                    double similarity = CalculateSimilarity(result, condition.OcrTextToMatch);
                    // 实时读取阈值并检查是否匹配
                    bool isMatch = similarity * 100 >= condition.OcrSimilarityThreshold;
                    return new OcrResult
                    {
                        Text = result,
                        Similarity = similarity,
                        IsMatch = isMatch,
                        Region = region
                    };
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"OCR识别失败: {ex.Message}");
                return new OcrResult { Text = "", Similarity = 0, IsMatch = false };
            }
        }

        public class OcrResult
        {
            public string Text { get; set; }
            public double Similarity { get; set; }
            public bool IsMatch { get; set; }
            public System.Drawing.Rectangle Region { get; set; }
        }

        //public static string RecognizeText(Bitmap image)
        //{
        //    if (!_isInitialized)
        //    {
        //        try
        //        {
        //            Initialize();
        //        }
        //        catch
        //        {
        //            return string.Empty;
        //        }
        //    }

        //    if (!_isInitialized)
        //        return string.Empty;
        //    //FullOcrModel model = LocalFullModels.ChineseV5;
        //    //using (PaddleOcrAll all = new PaddleOcrAll(model, PaddleDevice.Mkldnn())
        //    //{
        //    //    AllowRotateDetection = false, /* 允许识别有角度的文字 */
        //    //    Enable180Classification = false, /* 允许识别旋转角度大于90度的文字 */
        //    //})
        //    try
        //    {
        //        using (Mat mat = BitmapToMat(image))
        //        {
        //            PaddleOcrResult result = _ocrEngine.Run(mat);
        //            return result.Text;
        //        }
        //        //using (Mat src = Cv2.ImRead(@"C:\Users\Dait\Desktop\test.png"))
        //        //{
        //        //    PaddleOcrResult result = _ocrEngine.Run(src);
        //        //    Debug.WriteLine("Detected all texts: \n" + result.Text);
        //        //    foreach (PaddleOcrResultRegion region in result.Regions)
        //        //    {
        //        //        Debug.WriteLine($"Text: {region.Text}, Score: {region.Score}, RectCenter: {region.Rect.Center}, RectSize:    {region.Rect.Size}, Angle: {region.Rect.Angle}");
        //        //    }
        //        //    return result.Text;
        //        //}
        //    }
        //    catch (Exception ex)
        //    {
        //        Debug.WriteLine($"OCR识别失败: {ex.Message}");
        //        return string.Empty;
        //    }
        //}
        // 在RecognizeText方法中添加缓存支持
        public static string RecognizeText(Bitmap image, string cacheKey = null)
        {
            if (!_isInitialized)
            {
                try
                {
                    Initialize();
                }
                catch
                {
                    return string.Empty;
                }
            }

            if (!_isInitialized)
                return string.Empty;

            // 如果有缓存键，尝试从缓存获取
            if (!string.IsNullOrEmpty(cacheKey))
            {
                return OcrCacheService.GetOrAdd(cacheKey, () => RecognizeTextInternal(image));
            }

            return RecognizeTextInternal(image);
        }

        private static string RecognizeTextInternal(Bitmap image)
        {
            try
            {
                using (Mat mat = BitmapToMat(image))
                {
                    PaddleOcrResult result = _ocrEngine.Run(mat);
                    return result.Text;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"OCR识别失败: {ex.Message}");
                return string.Empty;
            }
        }
        private static Mat BitmapToMat(Bitmap bitmap)
        {
            // 将Bitmap转换为OpenCV的Mat
            var rect = new System.Drawing.Rectangle(0, 0, bitmap.Width, bitmap.Height);
            var data = bitmap.LockBits(rect, System.Drawing.Imaging.ImageLockMode.ReadOnly, bitmap.PixelFormat);

            try
            {
                // 使用新的 FromPixelData 方法
                Mat mat = Mat.FromPixelData(
                    bitmap.Height,
                    bitmap.Width,
                    MatType.CV_8UC3,
                    data.Scan0,
                    data.Stride);

                return mat.Clone(); // 创建副本以避免释放问题
            }
            finally
            {
                bitmap.UnlockBits(data);
            }
        }
        // 计算字符串相似度（使用Levenshtein距离算法）
        public static double CalculateSimilarity(string source, string target)
        {
            if (string.IsNullOrEmpty(source))
                return string.IsNullOrEmpty(target) ? 1.0 : 0.0;

            if (string.IsNullOrEmpty(target))
                return 0.0;

            double maxLength = Math.Max(source.Length, target.Length);
            if (maxLength == 0)
                return 1.0;

            int distance = ComputeLevenshteinDistance(source, target);
            return 1.0 - (distance / maxLength);
        }
        private static int ComputeLevenshteinDistance(string source, string target)
        {
            int[,] distance = new int[source.Length + 1, target.Length + 1];

            for (int i = 0; i <= source.Length; i++)
                distance[i, 0] = i;

            for (int j = 0; j <= target.Length; j++)
                distance[0, j] = j;

            for (int i = 1; i <= source.Length; i++)
            {
                for (int j = 1; j <= target.Length; j++)
                {
                    int cost = (target[j - 1] == source[i - 1]) ? 0 : 1;
                    distance[i, j] = Math.Min(
                        Math.Min(distance[i - 1, j] + 1, distance[i, j - 1] + 1),
                        distance[i - 1, j - 1] + cost);
                }
            }

            return distance[source.Length, target.Length];
        }
        // 在 OcrService.cs 中添加以下方法
        public static void DrawDebugRectangle(System.Drawing.Rectangle region, System.Drawing.Color color, int durationMs = 1000)
        {
            try
            {
                // 获取主窗口的DPI缩放因子
                var mainWindow = Application.Current.MainWindow;
                if (mainWindow == null) return;

                var dpiScale = VisualTreeHelper.GetDpi(mainWindow);
                double dpiScaleX = dpiScale.DpiScaleX;
                double dpiScaleY = dpiScale.DpiScaleY;

                // 调整区域坐标和大小以考虑DPI缩放
                var adjustedRegion = new System.Drawing.Rectangle(
                    (int)(region.X / dpiScaleX),
                    (int)(region.Y / dpiScaleY),
                    (int)(region.Width / dpiScaleX),
                    (int)(region.Height / dpiScaleY)
                );

                // 创建一个临时窗口来绘制矩形
                var debugWindow = new System.Windows.Window
                {
                    WindowStyle = System.Windows.WindowStyle.None,
                    AllowsTransparency = true,
                    Background = System.Windows.Media.Brushes.Transparent,
                    Topmost = true,
                    ShowInTaskbar = false,
                    Width = System.Windows.SystemParameters.PrimaryScreenWidth,
                    Height = System.Windows.SystemParameters.PrimaryScreenHeight,
                    Left = 0,
                    Top = 0
                };

                var canvas = new System.Windows.Controls.Canvas();
                debugWindow.Content = canvas;

                // 创建矩形框
                var rect = new System.Windows.Shapes.Rectangle
                {
                    Width = adjustedRegion.Width,
                    Height = adjustedRegion.Height,
                    Stroke = new System.Windows.Media.SolidColorBrush(
                        System.Windows.Media.Color.FromArgb(color.A, color.R, color.G, color.B)),
                    StrokeThickness = 2,
                    Fill = System.Windows.Media.Brushes.Transparent
                };

                System.Windows.Controls.Canvas.SetLeft(rect, adjustedRegion.X);
                System.Windows.Controls.Canvas.SetTop(rect, adjustedRegion.Y);
                canvas.Children.Add(rect);

                // 添加文本显示区域信息
                var textBlock = new System.Windows.Controls.TextBlock
                {
                    Text = $"({region.X},{region.Y}) {region.Width}x{region.Height}",
                    Foreground = System.Windows.Media.Brushes.Red,
                    Background = System.Windows.Media.Brushes.White,
                    FontWeight = System.Windows.FontWeights.Bold
                };

                System.Windows.Controls.Canvas.SetLeft(textBlock, adjustedRegion.X);
                System.Windows.Controls.Canvas.SetTop(textBlock, adjustedRegion.Y - 20);
                canvas.Children.Add(textBlock);

                // 显示窗口
                debugWindow.Show();

                // 设置定时器自动关闭窗口
                var timer = new System.Windows.Threading.DispatcherTimer
                {
                    Interval = TimeSpan.FromMilliseconds(durationMs)
                };
                timer.Tick += (s, e) =>
                {
                    timer.Stop();
                    debugWindow.Close();
                };
                timer.Start();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"绘制调试矩形失败: {ex.Message}");
            }
        }
        public static void Dispose()
        {
            _ocrEngine?.Dispose();
            _isInitialized = false;
        }
    }
}