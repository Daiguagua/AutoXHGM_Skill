using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using Vanara.PInvoke;
//using System.Windows.Forms;

namespace AutoXHGM_Skill
{
    //public class BooleanToVisibilityConverter : IValueConverter
    //{
    //    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    //    {
    //        return (value is bool && (bool)value) ? Visibility.Visible : Visibility.Collapsed;
    //    }

    //    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    //    {
    //        return value is Visibility && (Visibility)value == Visibility.Visible;
    //    }
    //}

    //public class InverseBooleanToVisibilityConverter : IValueConverter
    //{
    //    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    //    {
    //        return (value is bool && (bool)value) ? Visibility.Collapsed : Visibility.Visible;
    //    }

    //    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    //    {
    //        return value is Visibility && (Visibility)value == Visibility.Collapsed;
    //    }
    //}
    /// <summary>
    /// ConditionEditorWindow.xaml 的交互逻辑
    /// </summary>
    /// 
    public partial class ConditionEditorWindow : Window
    {
        //private MainWindow _mainWindow;
        private readonly SkillRule _rule;
        private IntPtr _targetWindowHandle; // 添加目标窗口句柄
        private void AddCondition_Click(object sender, RoutedEventArgs e)
        {
            // 创建选择对话框
            var dialog = new ConditionTypeDialog();
            if (dialog.ShowDialog() == true)
            {
                if (dialog.SelectedType == ConditionType.Color)
                {
                    AddColorCondition();
                }
                else if (dialog.SelectedType == ConditionType.OCR)
                {
                    AddOcrConditionWithRegionSelection();
                }
            }
        }
        private void AddColorCondition()
        {
            // 使用专业取色器
            var picker = new ProfessionalColorPicker(_targetWindowHandle);

            if (picker.ShowDialog() == true)
            {
                _rule.Conditions.Add(new SkillCondition
                {
                    OffsetX = picker.SelectedPosition.X,
                    OffsetY = picker.SelectedPosition.Y,
                    TargetColor = picker.SelectedColor,
                    Tolerance = 10
                });
            }
        }
        private void AddOcrConditionWithRegionSelection()
        {
            Application.Current.Dispatcher.Invoke(() =>
            {
                // 使用区域选择器
                var regionSelector = new RegionSelectorWindow(_targetWindowHandle);

                if (regionSelector.ShowDialog() == true)
                {
                    var selectedRegion = regionSelector.SelectedRegion;

                    _rule.Conditions.Add(new SkillCondition
                    {
                        IsOcrCondition = true,
                        OcrRegionX = selectedRegion.X,
                        OcrRegionY = selectedRegion.Y,
                        OcrRegionWidth = selectedRegion.Width,
                        OcrRegionHeight = selectedRegion.Height,
                        OcrTextToMatch = "请输入要匹配的文本",
                        OcrSimilarityThreshold = 70
                    });
                }
            });
        }
        public ConditionEditorWindow(SkillRule rule, IntPtr targetWindowHandle)
        {
            InitializeComponent();
            _rule = rule;

            // 确保有有效的窗口句柄
            if (targetWindowHandle == IntPtr.Zero)
            {
                // 尝试获取活动窗口
                targetWindowHandle = GetActiveWindow();

                // 如果仍然无效，使用主窗口
                if (targetWindowHandle == IntPtr.Zero)
                {
                    var mainWindow = Application.Current.MainWindow;
                    if (mainWindow != null)
                    {
                        var helper = new WindowInteropHelper(mainWindow);
                        targetWindowHandle = helper.Handle;
                    }
                }
            }

            // 验证窗口句柄是否有效
            if (!IsWindow(targetWindowHandle))
            {
                MessageBox.Show("无效的窗口句柄，将使用主窗口");
                var helper = new WindowInteropHelper(Application.Current.MainWindow);
                targetWindowHandle = helper.Handle;
            }

            _targetWindowHandle = targetWindowHandle;
            DataContext = rule;
        }
        // 添加验证窗口的Win32 API
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsWindow(IntPtr hWnd);
        // 添加获取活动窗口的 Win32 API
        [DllImport("user32.dll")]
        private static extern IntPtr GetActiveWindow();
        // 添加双击事件处理
        private void DgConditions_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            // 先提交任何挂起的编辑并退出编辑模式
            dgConditions.CommitEdit();   // 提交当前编辑单元
            dgConditions.CancelEdit();   // 退出编辑模式
            if (dgConditions.SelectedItem is SkillCondition selectedCondition)
    {
        if (selectedCondition.IsOcrCondition)
        {
            // 对于OCR条件，重新选择区域
            var regionSelector = new RegionSelectorWindow(_targetWindowHandle);
            
            if (regionSelector.ShowDialog() == true)
            {
                var selectedRegion = regionSelector.SelectedRegion;
                
                selectedCondition.OcrRegionX = selectedRegion.X;
                selectedCondition.OcrRegionY = selectedRegion.Y;
                selectedCondition.OcrRegionWidth = selectedRegion.Width;
                selectedCondition.OcrRegionHeight = selectedRegion.Height;
                
                // 刷新显示
                dgConditions.Items.Refresh();
            }
        }
        else
        {
            // 对于颜色条件，使用取色器
            var picker = new ProfessionalColorPicker(_targetWindowHandle);

            if (picker.ShowDialog() == true)
            {
                selectedCondition.OffsetX = picker.SelectedPosition.X;
                selectedCondition.OffsetY = picker.SelectedPosition.Y;
                selectedCondition.TargetColor = picker.SelectedColor;

                // 刷新显示
                dgConditions.Items.Refresh();
            }
        }
    }
        }
        //private void TestOcr_Click(object sender, RoutedEventArgs e)
        //{
        //    if (dgConditions.SelectedItem is SkillCondition selectedCondition && selectedCondition.IsOcrCondition)
        //    {
        //        // 获取窗口位置和客户区偏移（与取色模式相同）
        //        Win32PointHelper.GetWindowRect(_targetWindowHandle, out var windowRect);
        //        var clientOffset = Win32PointHelper.GetClientTopLeft(_targetWindowHandle);

        //        // 计算OCR区域在屏幕上的绝对坐标（与取色模式相同）
        //        int absoluteX = windowRect.Left + clientOffset.X + selectedCondition.OcrRegionX;
        //        int absoluteY = windowRect.Top + clientOffset.Y + selectedCondition.OcrRegionY;

        //        // 输出调试信息
        //        Debug.WriteLine($"OCR测试 - 窗口位置: L={windowRect.Left}, T={windowRect.Top}");
        //        Debug.WriteLine($"OCR测试 - 客户区偏移: X={clientOffset.X}, Y={clientOffset.Y}");
        //        Debug.WriteLine($"OCR测试 - 区域坐标: X={selectedCondition.OcrRegionX}, Y={selectedCondition.OcrRegionY}");
        //        Debug.WriteLine($"OCR测试 - 绝对坐标: X={absoluteX}, Y={absoluteY}");
        //        Debug.WriteLine($"OCR测试 - 区域大小: W={selectedCondition.OcrRegionWidth}, H={selectedCondition.OcrRegionHeight}");

        //        // 截取指定区域
        //        var region = new System.Drawing.Rectangle(
        //            absoluteX, absoluteY,
        //            selectedCondition.OcrRegionWidth, selectedCondition.OcrRegionHeight);
        //        // 绘制调试矩形（蓝色，显示3秒）
        //        OcrService.DrawDebugRectangle(region, System.Drawing.Color.Blue, 3000);
        //        try
        //        {
        //            //using (var bitmap = new Bitmap(region.Width, region.Height))
        //            //using (var graphics = Graphics.FromImage(bitmap))
        //            //{
        //            //    graphics.CopyFromScreen(region.Location, System.Drawing.Point.Empty, region.Size);

        //            //    // 调用OCR识别
        //            //    string result = OcrService.RecognizeText(bitmap);
        //            //    txtOcrResult.Text = result;

        //            //    // 计算相似度
        //            //    double similarity = OcrService.CalculateSimilarity(result, selectedCondition.OcrTextToMatch);
        //            //    MessageBox.Show($"OCR识别结果: {result}\n相似度: {similarity:P2}");
        //            //}

        //            using (var bitmap =MainWindow.CaptureRegion(region))
        //            {
        //                // 调用OCR识别
        //                string result = OcrService.RecognizeText(bitmap);
        //                txtOcrResult.Text = result;

        //                // 计算相似度
        //                double similarity = OcrService.CalculateSimilarity(result, selectedCondition.OcrTextToMatch);
        //                MessageBox.Show($"OCR识别结果: {result}\n相似度: {similarity:P2}");
        //                //}}
        //            }
        //        }
        //        catch (Exception ex)
        //        {
        //            MessageBox.Show($"OCR测试失败: {ex.Message}");
        //        }
        //    }
        //    else
        //    {
        //        MessageBox.Show("请先选择一个OCR条件");
        //    }
        //}
        // ConditionEditorWindow.xaml.cs
        private bool TestOcrCondition(SkillCondition condition)
        {
            if (condition == null || !condition.IsOcrCondition) return false;

            var result = OcrService.RecognizeTextFromRegion(_targetWindowHandle, condition);

            // 显示结果
            txtOcrResult.Text = $"识别结果: {result.Text}\n相似度: {result.Similarity:P2}\n匹配: {result.IsMatch}";

            return result.IsMatch;
        }

        private void TestOcr_Click(object sender, RoutedEventArgs e)
        {
            // 提交任何挂起的编辑
            dgConditions.CommitEdit(); // 提交当前编辑的单元格
            dgConditions.CancelEdit(); // 退出编辑模式，确保更改已提交

            if (dgConditions.SelectedItem is SkillCondition selectedCondition)
            {
                bool isMatch = TestOcrCondition(selectedCondition);
                MessageBox.Show($"OCR测试完成，匹配结果: {isMatch}", "测试结果");
            }
            else
            {
                MessageBox.Show("请先选择一个OCR条件");
            }
        }

        private void DeleteCondition_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.DataContext is SkillCondition condition)
            {
                _rule.Conditions.Remove(condition);
            }
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
        private void DgConditions_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //MainWindow mainWindow = new MainWindow();
            if (dgConditions.SelectedItem is SkillCondition selectedCondition)
            {
                // 获取窗口位置
                Win32PointHelper.GetWindowRect(_targetWindowHandle, out var windowRect);
                // 获取客户区偏移
                var clientOffset = Win32PointHelper.GetClientTopLeft(_targetWindowHandle);
                // 计算绝对坐标 (使用一致的算法)
                int absoluteX = windowRect.Left + clientOffset.X + selectedCondition.OffsetX;
                int absoluteY = windowRect.Top + clientOffset.Y + selectedCondition.OffsetY;
                // 获取颜色
                var pixelColor = Win32ColorHelper.GetPixelColor(absoluteX, absoluteY);

                // 更新UI显示
                tbCurrentColor.Text = $"当前颜色: R:{pixelColor.R} G:{pixelColor.G} B:{pixelColor.B}";
                borderCurrentColor.Background = new SolidColorBrush(
                    System.Windows.Media.Color.FromRgb(pixelColor.R, pixelColor.G, pixelColor.B));
            }
        }

    }
}