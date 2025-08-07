using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using Vanara.PInvoke;

namespace AutoXHGM_Skill
{
    /// <summary>
    /// ConditionEditorWindow.xaml 的交互逻辑
    /// </summary>
    public partial class ConditionEditorWindow : Window
    {
        private readonly SkillRule _rule;
        private IntPtr _targetWindowHandle; // 添加目标窗口句柄

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
                // 使用专业取色器
                var picker = new ProfessionalColorPicker(_targetWindowHandle);

                if (picker.ShowDialog() == true)
                {
                    selectedCondition.OffsetX = picker.SelectedPosition.X;
                    selectedCondition.OffsetY = picker.SelectedPosition.Y;
                    selectedCondition.TargetColor = picker.SelectedColor;

                    // 刷新显示
                    //dgConditions.Items.Refresh();
                }
            }
        }
        private void AddCondition_Click(object sender, RoutedEventArgs e)
        {
            _rule.Conditions.Add(new SkillCondition
            {
                OffsetX = 0,
                OffsetY = 0,
                TargetColor = System.Drawing.Color.Blue,
                Tolerance = 10
            });
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
            MainWindow mainWindow = new MainWindow();
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
