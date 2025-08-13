using Fischless.HotkeyCapture;
using Fischless.WindowsInput;
using Hardcodet.Wpf.TaskbarNotification;
using Microsoft.Win32;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Shapes;
using System.Windows.Threading;
using Vanara.PInvoke;
using Color = System.Drawing.Color;
using Point = System.Drawing.Point;

namespace AutoXHGM_Skill
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        // 修改1：添加热键处理锁防止重入
        private bool _isHotKeyProcessing = false;
        // 添加系统托盘图标
        private TaskbarIcon _trayIcon;
        private readonly InputSimulator _inputSimulator = new InputSimulator();
        private readonly List<WindowInfo> _gameWindows = new List<WindowInfo>();
        private DispatcherTimer _checkTimer;
        private IntPtr _selectedWindowHandle = IntPtr.Zero;
        private bool _isRunning = false;
        private string _configFilePath = "skill_config.json";
        
        private HotkeyHolder _hotkeyHolder = new HotkeyHolder();
        //// 热键注册
        //[DllImport("user32.dll")]
        //private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        //[DllImport("user32.dll")]
        //private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        // 热键相关常量
        //private const int HOTKEY_ID_START = 1;
        //private const int HOTKEY_ID_STOP = 2;
        //private const uint MOD_NONE = 0x0000;

        // 鼠标按键虚拟键码
        private const uint VK_MBUTTON = 0x04;     // 鼠标中键
        private const uint VK_XBUTTON1 = 0x05;    // 鼠标侧键1
        private const uint VK_XBUTTON2 = 0x06;    // 鼠标侧键2
        private const uint VK_OEM_3 = 0xC0;       // 波浪号(~)
        private const uint VK_OEM_1 = 0xBA;       // 分号(;)
        private const uint VK_OEM_2 = 0xBF;       // 问号(?)
        private const uint VK_OEM_4 = 0xDB;       // 正括号([)
        private const uint VK_OEM_5 = 0xDC;       // 斜线(\)
        private const uint VK_OEM_6 = 0xDD;       // 反括号(])
        private const uint VK_OEM_COMMA = 0xBC;   // 逗号(,)
        private const uint VK_OEM_PERIOD = 0xBE;  // 句号(.)
        private const uint VK_OEM_PLUS = 0xBB;    // 等号(=)
        private const uint VK_OEM_MINUS = 0xBD;   // 横线(-)

        // 小键盘键码
        private const uint VK_NUMPAD0 = 0x60;
        private const uint VK_NUMPAD1 = 0x61;
        private const uint VK_NUMPAD2 = 0x62;
        private const uint VK_NUMPAD3 = 0x63;
        private const uint VK_NUMPAD4 = 0x64;
        private const uint VK_NUMPAD5 = 0x65;
        private const uint VK_NUMPAD6 = 0x66;
        private const uint VK_NUMPAD7 = 0x67;
        private const uint VK_NUMPAD8 = 0x68;
        private const uint VK_NUMPAD9 = 0x69;
        private const uint VK_DECIMAL = 0x6E;     // 小键盘.
        private const uint VK_DIVIDE = 0x6F;      // 小键盘/
        private const uint VK_MULTIPLY = 0x6A;    // 小键盘*
        private const uint VK_SUBTRACT = 0x6D;    // 小键盘-
        private const uint VK_ADD = 0x6B;         // 小键盘+

        // 存储热键设置
        private string _startHotKey = "~";
        private string _stopHotKey = "~";


        private readonly List<Ellipse> _debugPoints = new List<Ellipse>();
        private DispatcherTimer _positionUpdateTimer;
        // 添加初始化调试点的方法
        private void InitializeDebugPoints()
        {
            debugCanvas.Children.Clear();
            _debugPoints.Clear();

            foreach (var rule in Rules.Where(r => r.IsEnabled))
            {
                foreach (var condition in rule.Conditions)
                {
                    var ellipse = new Ellipse
                    {
                        Width = 6,
                        Height = 6,
                        Fill = System.Windows.Media.Brushes.Red,
                        Stroke = System.Windows.Media.Brushes.White,
                        StrokeThickness = 1,
                        Visibility = Visibility.Visible,
                        Opacity = 0.8
                    };

                    debugCanvas.Children.Add(ellipse);
                    _debugPoints.Add(ellipse);
                }
            }
        }
        // 更新调试点位置
        private void UpdateDebugPoints()
        {
            if (_selectedWindowHandle == IntPtr.Zero) return;

            // 获取窗口位置
            GetWindowRect(_selectedWindowHandle, out RECT windowRect);

            int pointIndex = 0;
            foreach (var rule in Rules.Where(r => r.IsEnabled))
            {
                foreach (var condition in rule.Conditions)
                {
                    if (pointIndex >= _debugPoints.Count) break;

                    var ellipse = _debugPoints[pointIndex];

                    // 计算绝对坐标
                    int screenX = windowRect.Left + condition.OffsetX;
                    int screenY = windowRect.Top + condition.OffsetY;

                    // 转换为窗口内坐标
                    var pointInWindow = PointFromScreen(new System.Windows.Point(screenX, screenY));

                    // 设置红点位置
                    Canvas.SetLeft(ellipse, pointInWindow.X - 3);
                    Canvas.SetTop(ellipse, pointInWindow.Y - 3);

                    pointIndex++;
                }
            }
        }
        // 添加窗口位置监听定时器
        private void StartPositionTracking()
        {
            _positionUpdateTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(100),
                IsEnabled = true
            };

            _positionUpdateTimer.Tick += (s, e) => UpdateDebugPoints();
            _positionUpdateTimer.Start();
        }
        private void StopPositionTracking()
        {
            _positionUpdateTimer?.Stop();
            _positionUpdateTimer = null;
        }
        public ObservableCollection<SkillRule> Rules { get; } = new ObservableCollection<SkillRule>();
        private class ConfigData
        {
            public string SelectedWindowTitle { get; set; }
            public ObservableCollection<SkillRule> Rules { get; set; }
            public string StartHotKey { get; set; } = "~";
            public string StopHotKey { get; set; } = "~";
        }
        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;
            LoadConfig();
            RefreshWindows();
            // 初始化系统托盘
            InitializeTrayIcon();
            // 简化热键注册
            Loaded += (s, e) => {
                // 延迟注册确保窗口完全加载
                Dispatcher.BeginInvoke(() => RegisterHotkeys(), DispatcherPriority.ApplicationIdle);
            };

            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            _configFilePath = System.IO.Path.Combine(appData, "AutoXHGM_Skill", "skill_config.json");
            Directory.CreateDirectory(System.IO.Path.GetDirectoryName(_configFilePath));
        }
        private void InitializeTrayIcon()
        {
            _trayIcon = new TaskbarIcon
            {
                Icon = CreateIconFromVisual(this), // 从窗口创建图标
                ToolTipText = "星痕共鸣技能循环助手",
                Visibility = Visibility.Visible
            };

            // 创建托盘菜单
            _trayIcon.ContextMenu = new ContextMenu
            {
                Items =
            {
                new MenuItem { Header = "打开主界面", Command = new RelayCommand(_ => ShowFromTray()) },
                new MenuItem { Header = "退出", Command = new RelayCommand(_ => Close()) }
            }
            };

            _trayIcon.TrayMouseDoubleClick += (s, e) => ShowFromTray();
        }
        private Icon CreateIconFromVisual(Visual visual)
        {
            // 直接加载项目中的 ICO 文件
            var iconPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "app.ico");

            if (System.IO.File.Exists(iconPath))
            {
                return new System.Drawing.Icon(iconPath);
            }

            // 备用方法：创建默认图标
            return System.Drawing.Icon.ExtractAssociatedIcon(System.Reflection.Assembly.GetEntryAssembly().Location);
        }
        private void ShowFromTray()
        {
            Show();
            WindowState = WindowState.Normal;
            Activate();
        }
        private void HideToTray()
        {
            Hide();
        }
        private int _retryCount = 0;
        private const int MAX_RETRY_COUNT = 5;
        private int _registerRetryCount = 0;
        private const int MAX_REGISTER_RETRY = 3;
        private bool _isClosing = false;
        // 添加 IsClosing 属性
        public bool IsClosing => _isClosing;
        private bool IsWindowReadyForHotkeys()
        {
            return this.IsLoaded &&
                   this.IsVisible &&
                   !_isClosing &&
                   PresentationSource.FromVisual(this) != null;
        }
        protected override void OnClosing(CancelEventArgs e)
        {
            _isClosing = true;

            // 清理热键
            try
            {
                // 只清理新热键系统
                HotkeyHolder.UnregisterHotKey();
                Debug.WriteLine("[CLEANUP] 热键已注销");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"关闭时清理热键失败: {ex.Message}");
            }

            base.OnClosing(e);
        }

        private void RegisterHotkeys()
        {
            try
            {
                Debug.WriteLine("[HOTKEY] 开始重新注册热键");

                // 先清理所有热键
                HotkeyHolder.UnregisterHotKey();

                // 延迟一下确保清理完成
                Thread.Sleep(100);

                // 只注册启动热键（用于切换状态）
                if (!string.IsNullOrEmpty(_startHotKey))
                {
                    Debug.WriteLine($"[HOTKEY] 准备注册热键: '{_startHotKey}'");
                    HotkeyHolder.RegisterHotKey(_startHotKey, HotkeyPressed);
                    Debug.WriteLine($"[HOTKEY] 热键注册完成: '{_startHotKey}'");
                }
                else
                {
                    Debug.WriteLine("[HOTKEY] 启动热键为空，跳过注册");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ERROR] 热键注册异常: {ex.Message}");
                Debug.WriteLine($"[ERROR] 堆栈跟踪: {ex.StackTrace}");
                System.Windows.MessageBox.Show($"热键注册失败: {ex.Message}\n请修改热键设置",
                               "热键冲突", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        // 新的热键处理函数
        private void HotkeyPressed(object sender, KeyPressedEventArgs e)
        {
            Debug.WriteLine("=== 热键事件开始 ===");
            Debug.WriteLine($"[HOTKEY_EVENT] ✅ 收到热键事件!");
            Debug.WriteLine($"[HOTKEY_EVENT] Sender: {sender?.GetType().Name}");
            Debug.WriteLine($"[HOTKEY_EVENT] Modifier: {e.Modifier}");
            Debug.WriteLine($"[HOTKEY_EVENT] Key: {e.Key}");
            Debug.WriteLine($"[HOTKEY_EVENT] 时间: {DateTime.Now:HH:mm:ss.fff}");

            try
            {
                Dispatcher.Invoke(() =>
                {
                    string pressedHotkey = BuildPressedHotkeyString(e);
                    Debug.WriteLine($"[HOTKEY] 构建的热键字符串: '{pressedHotkey}'");
                    Debug.WriteLine($"[HOTKEY] 期望的启动热键: '{_startHotKey}'");

                    if (pressedHotkey == _startHotKey)
                    {
                        Debug.WriteLine("[HOTKEY] ✅ 热键匹配成功!");

                        if (_isRunning)
                        {
                            StopSkillLoop();
                            Debug.WriteLine("[HOTKEY] 停止技能循环");
                        }
                        else
                        {
                            StartSkillLoop();
                            Debug.WriteLine("[HOTKEY] 启动技能循环");
                        }
                    }
                    else
                    {
                        Debug.WriteLine($"[HOTKEY] ❌ 热键不匹配");
                        Debug.WriteLine($"[HOTKEY] 期望: '{_startHotKey}'");
                        Debug.WriteLine($"[HOTKEY] 实际: '{pressedHotkey}'");
                    }
                });
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ERROR] 热键处理异常: {ex.Message}");
                Debug.WriteLine($"[ERROR] 堆栈跟踪: {ex.StackTrace}");
            }

            Debug.WriteLine("=== 热键事件结束 ===");
        }
        private string BuildPressedHotkeyString(KeyPressedEventArgs e)
        {
            var builder = new StringBuilder();

            // 添加修饰键
            if (e.Modifier.HasFlag(User32.HotKeyModifiers.MOD_CONTROL))
                builder.Append("Ctrl + ");
            if (e.Modifier.HasFlag(User32.HotKeyModifiers.MOD_SHIFT))
                builder.Append("Shift + ");
            if (e.Modifier.HasFlag(User32.HotKeyModifiers.MOD_ALT))
                builder.Append("Alt + ");

            // 添加主键，需要反向映射
            string keyName = GetKeyDisplayName(e.Key);
            builder.Append(keyName);

            return builder.ToString().TrimEnd(' ', '+');
        }
        private string GetKeyDisplayName(Keys key)
        {
            return key switch
            {
                Keys.XButton1 => "鼠标侧键1",
                Keys.XButton2 => "鼠标侧键2",
                Keys.MButton => "鼠标中键",
                Keys.Oem3 => "~",
                // ... 其他映射
                _ => key.ToString()
            };
        }
        private User32.HotKeyModifiers GetCurrentModifiers(bool isStart)
        {
            var modifiers = User32.HotKeyModifiers.MOD_NONE;

            if (isStart)
            {
                if (chkstartCtrl.IsChecked == true) modifiers |= User32.HotKeyModifiers.MOD_CONTROL;
                if (chkstartShift.IsChecked == true) modifiers |= User32.HotKeyModifiers.MOD_SHIFT;
                if (chkstartAlt.IsChecked == true) modifiers |= User32.HotKeyModifiers.MOD_ALT;
            }
            else
            {
                if (chkstopCtrl.IsChecked == true) modifiers |= User32.HotKeyModifiers.MOD_CONTROL;
                if (chkstopShift.IsChecked == true) modifiers |= User32.HotKeyModifiers.MOD_SHIFT;
                if (chkstopAlt.IsChecked == true) modifiers |= User32.HotKeyModifiers.MOD_ALT;
            }

            return modifiers;
        }
        // 辅助方法：将字符串转换为Keys枚举
        private Keys GetKeyFromString(string keyStr)
        {
            return keyStr switch
            {
                // 数字键
                "0" => Keys.D0,
                "1" => Keys.D1,
                "2" => Keys.D2,
                "3" => Keys.D3,
                "4" => Keys.D4,
                "5" => Keys.D5,
                "6" => Keys.D6,
                "7" => Keys.D7,
                "8" => Keys.D8,
                "9" => Keys.D9,

                // 符号键
                ";" => Keys.Oem1,
                "?" => Keys.Oem2,
                "~" => Keys.Oem3,
                "[" => Keys.Oem4,
                "\\" => Keys.Oem5,
                "]" => Keys.Oem6,
                "," => Keys.Oemcomma,
                "." => Keys.OemPeriod,
                "=" => Keys.Oemplus,
                "-" => Keys.OemMinus,

                // 功能键
                "F9" => Keys.F9,
                "F10" => Keys.F10,
                "F11" => Keys.F11,
                "F12" => Keys.F12,

                // 鼠标键
                "鼠标中键" => Keys.MButton,
                "鼠标侧键1" => Keys.XButton1,
                "鼠标侧键2" => Keys.XButton2,
                "鼠标滚轮上" => Keys.None,  // 需要特殊处理
                "鼠标滚轮下" => Keys.None,  // 需要特殊处理

                // 小键盘
                "小键盘/" => Keys.Divide,
                "小键盘*" => Keys.Multiply,
                "小键盘-" => Keys.Subtract,
                "小键盘+" => Keys.Add,
                "小键盘." => Keys.Decimal,
                "小键盘0" => Keys.NumPad0,
                "小键盘1" => Keys.NumPad1,
                "小键盘2" => Keys.NumPad2,
                "小键盘3" => Keys.NumPad3,
                "小键盘4" => Keys.NumPad4,
                "小键盘5" => Keys.NumPad5,
                "小键盘6" => Keys.NumPad6,
                "小键盘7" => Keys.NumPad7,
                "小键盘8" => Keys.NumPad8,
                "小键盘9" => Keys.NumPad9,

                // 字母键（添加这行）
                _ => (Keys)Enum.Parse(typeof(Keys), keyStr, true)
            };
        }
        // 在窗口关闭时注销热键

        //private void RegisterHotKeyInternal(IntPtr handle)
        //{
        //    uint startVk = GetVkCode(_startHotKey);
        //    uint stopVk = GetVkCode(_stopHotKey);

        //    // 注册启动热键
        //    if (!RegisterHotKey(handle, HOTKEY_ID_START, MOD_NONE, startVk))
        //    {
        //        Debug.WriteLine($"启动热键注册失败: {_startHotKey}");
        //    }
        //    else
        //    {
        //        Debug.WriteLine($"启动热键注册成功: {_startHotKey}");
        //    }

        //    // 注册停止热键（如果不同）
        //    if (startVk != stopVk)
        //    {
        //        if (!RegisterHotKey(handle, HOTKEY_ID_STOP, MOD_NONE, stopVk))
        //        {
        //            Debug.WriteLine($"停止热键注册失败: {_stopHotKey}");
        //        }
        //        else
        //        {
        //            Debug.WriteLine($"停止热键注册成功: {_stopHotKey}");
        //        }
        //    }
        //}
        //private async Task InitializeHotkeysAsync()
        //{
        //    const int maxAttempts = 10;
        //    const int delayMs = 200;

        //    for (int attempt = 1; attempt <= maxAttempts; attempt++)
        //    {
        //        try
        //        {
        //            var helper = new WindowInteropHelper(this);
        //            helper.EnsureHandle();

        //            if (helper.Handle != IntPtr.Zero)
        //            {
        //                var source = HwndSource.FromHwnd(helper.Handle);
        //                if (source != null)
        //                {
        //                    _source = source;
        //                    _source.AddHook(HwndHook);
        //                    RegisterHotKeyInternal(helper.Handle);

        //                    Debug.WriteLine($"热键注册成功 (尝试 {attempt}/{maxAttempts})");
        //                    return;
        //                }
        //            }
        //        }
        //        catch (Exception ex)
        //        {
        //            Debug.WriteLine($"尝试 {attempt} 失败: {ex.Message}");
        //        }

        //        if (attempt < maxAttempts)
        //        {
        //            await Task.Delay(delayMs);
        //        }
        //    }

        //    Debug.WriteLine("所有热键注册尝试均失败");
        //    ShowHotkeyError();
        //}

        private void ShowHotkeyError()
        {
            tbStatus.Text = "热键注册失败，请手动控制";
            System.Windows.MessageBox.Show("热键注册失败，您需要手动点击按钮来控制程序",
                "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
        }

        //private void UnregisterHotkeys()
        //{
        //    //try
        //    //{
        //    var helper = new WindowInteropHelper(this);
        //    HotkeyHolder.UnregisterHotKey(helper.Handle, HOTKEY_ID_START);
        //    HotkeyHolder.UnregisterHotKey(helper.Handle, HOTKEY_ID_STOP);

        //    // 额外添加安全注销
        //    for (int i = 0; i < 10; i++) // 尝试注销多个可能的ID
        //    {
        //        HotkeyHolder.UnregisterHotKey(helper.Handle, i);
        //    }
        //}
        //    }
        //    catch { /* 忽略注销错误 */ }
        //}
        private uint GetVkCode(string hotkeyName)
        {
            return hotkeyName switch
            {
                // 数字键
                "0" => 0x30,
                "1" => 0x31,
                "2" => 0x32,
                "3" => 0x33,
                "4" => 0x34,
                "5" => 0x35,
                "6" => 0x36,
                "7" => 0x37,
                "8" => 0x38,
                "9" => 0x39,

                // 符号键
                ";" => VK_OEM_1,
                "?" => VK_OEM_2,
                "~" => VK_OEM_3,
                "[" => VK_OEM_4,
                "\\" => VK_OEM_5,
                "]" => VK_OEM_6,
                "," => VK_OEM_COMMA,
                "." => VK_OEM_PERIOD,
                "=" => VK_OEM_PLUS,
                "-" => VK_OEM_MINUS,

                // 功能键
                "F9" => 0x78,
                "F10" => 0x79,
                "F11" => 0x7A,
                "F12" => 0x7B,

                // 鼠标键
                "鼠标中键" => VK_MBUTTON,
                "鼠标侧键1" => VK_XBUTTON1,
                "鼠标侧键2" => VK_XBUTTON2,

                // 小键盘
                "小键盘/" => VK_DIVIDE,
                "小键盘*" => VK_MULTIPLY,
                "小键盘-" => VK_SUBTRACT,
                "小键盘+" => VK_ADD,
                "小键盘." => VK_DECIMAL,
                "小键盘0" => VK_NUMPAD0,
                "小键盘1" => VK_NUMPAD1,
                "小键盘2" => VK_NUMPAD2,
                "小键盘3" => VK_NUMPAD3,
                "小键盘4" => VK_NUMPAD4,
                "小键盘5" => VK_NUMPAD5,
                "小键盘6" => VK_NUMPAD6,
                "小键盘7" => VK_NUMPAD7,
                "小键盘8" => VK_NUMPAD8,
                "小键盘9" => VK_NUMPAD9,

                _ => VK_OEM_3 // 默认波浪号
            };
        }

        private void RefreshWindows_Click(object sender, RoutedEventArgs e)
        {
            RefreshWindows();
        }

        private void RefreshWindows()
        {
            _gameWindows.Clear();
            cbGameWindows.Items.Clear();

            // 查找所有窗口
            EnumWindows((hWnd, lParam) =>
            {
                if (IsWindowVisible(hWnd))
                {
                    const int maxLength = 256;
                    var title = new System.Text.StringBuilder(maxLength);
                    GetWindowText(hWnd, title, maxLength);

                    if (title.Length > 0)
                    {
                        var windowInfo = new WindowInfo
                        {
                            Handle = hWnd,
                            Title = title.ToString()
                        };

                        // 检查是否是目标游戏窗口
                        if (title.ToString().Contains("star") ||
                            GetProcessName(hWnd).ToLower().Contains("star"))
                        {
                            _gameWindows.Add(windowInfo);
                        }
                    }
                }
                return true;
            }, IntPtr.Zero);

            foreach (var window in _gameWindows)
            {
                cbGameWindows.Items.Add(window);
            }

            if (cbGameWindows.Items.Count > 0)
            {
                cbGameWindows.SelectedIndex = 0;
                // 关键修复：同步更新窗口句柄
                if (cbGameWindows.SelectedItem is WindowInfo selectedWindow)
                {
                    _selectedWindowHandle = selectedWindow.Handle;
                    tbStatus.Text = $"已自动选择窗口: {selectedWindow.Title}";
                }
            }
            else
            {
                // 清空状态
                _selectedWindowHandle = IntPtr.Zero;
                tbStatus.Text = "未找到符合条件的窗口";
            }
        }

        private void AddRule_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var rule = new SkillRule
                {
                    Key = (cbNewKey.SelectedItem as ComboBoxItem)?.Content.ToString(),
                    //CheckInterval = 50 // 默认值
                };

                // 添加默认条件
                rule.Conditions.Add(new SkillCondition
                {
                    OffsetX = 0,
                    OffsetY = 0,
                    TargetColor = Color.Blue,
                    Tolerance = 10
                });

                Rules.Add(rule);
                if (_isRunning)
                {
                    InitializeDebugPoints();
                    UpdateDebugPoints();
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"添加规则失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void RemoveRule_Click(object sender, RoutedEventArgs e)
        {
            if (dgRules.SelectedItem is SkillRule selectedRule)
            {
                Rules.Remove(selectedRule);
            }
            if (_isRunning)
            {
                InitializeDebugPoints();
                UpdateDebugPoints();
            }
        }

        private void MoveRuleUp_Click(object sender, RoutedEventArgs e)
        {
            if (dgRules.SelectedItem is SkillRule selectedRule)
            {
                int index = Rules.IndexOf(selectedRule);
                if (index > 0)
                {
                    Rules.RemoveAt(index);
                    Rules.Insert(index - 1, selectedRule);
                    dgRules.SelectedItem = selectedRule;
                }
            }
        }

        private void MoveRuleDown_Click(object sender, RoutedEventArgs e)
        {
            if (dgRules.SelectedItem is SkillRule selectedRule)
            {
                int index = Rules.IndexOf(selectedRule);
                if (index < Rules.Count - 1)
                {
                    Rules.RemoveAt(index);
                    Rules.Insert(index + 1, selectedRule);
                    dgRules.SelectedItem = selectedRule;
                }
            }
        }
        private void EditConditions_Click(object sender, RoutedEventArgs e)
        {
            if (sender is System.Windows.Controls.Button button && button.DataContext is SkillRule rule)
            {
                // 确保有选中的窗口
                if (_selectedWindowHandle == IntPtr.Zero)
                {
                    System.Windows.MessageBox.Show("请先选择一个游戏窗口", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // 验证窗口句柄是否仍然有效
                if (!IsWindow(_selectedWindowHandle))
                {
                    System.Windows.MessageBox.Show("当前选择的窗口已关闭，请重新选择", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                var editor = new ConditionEditorWindow(rule, _selectedWindowHandle);
                if (editor.ShowDialog() == true)
                {
                    dgRules.Items.Refresh();
                }
            }
            if (_isRunning)
            {
                InitializeDebugPoints();
                UpdateDebugPoints();
            }
        }

        private void StartStop_Click(object sender, RoutedEventArgs e)
        {
            if (!_isRunning)
            {
                StartSkillLoop();
            }
            else
            {
                StopSkillLoop();
            }
        }

        private void StartSkillLoop()
        {
            if (cbGameWindows.SelectedItem is WindowInfo selectedWindow)
            {
                // 从界面读取最新的热键设置
                //_startHotKey = (cbStartHotKey.SelectedItem as ComboBoxItem)?.Content.ToString() ?? "~";
                //_stopHotKey = (cbStopHotKey.SelectedItem as ComboBoxItem)?.Content.ToString() ?? "~";
                // 重新注册热键
                //UnregisterHotkeys();
                RegisterHotkeys();
                _selectedWindowHandle = selectedWindow.Handle;
                _isRunning = true;
                btnStart.IsEnabled = false;
                btnStop.IsEnabled = true;
                tbStatus.Text = "技能循环运行中...";
                tbActive.Text = "运行中";
                tbActive.Foreground = new SolidColorBrush(Colors.Green);

                // 最小化到托盘
                HideToTray();

                int interval = int.TryParse(txtGlobalInterval.Text, out int globalInterval) ? globalInterval : 50;

                _checkTimer = new DispatcherTimer
                {
                    Interval = TimeSpan.FromMilliseconds(interval)
                };
                _checkTimer.Tick += CheckRules;
                _checkTimer.Start();
                // 显示调试点
                debugCanvas.Visibility = Visibility.Visible;
                InitializeDebugPoints();
                StartPositionTracking();
                UpdateDebugPoints();
            }
            else
            {
                System.Windows.MessageBox.Show("请先选择一个游戏窗口", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void StopSkillLoop()
        {
            _isRunning = false;
            btnStart.IsEnabled = true;
            btnStop.IsEnabled = false;
            tbStatus.Text = "已停止技能循环";
            tbActive.Text = "未运行";
            tbActive.Foreground = new SolidColorBrush(Colors.Gray);

            // 不自动显示窗口，保持托盘状态
            //ShowFromTray(); // 注释掉这行，保持后台运行

            if (_checkTimer != null)
            {
                _checkTimer.Stop();
                _checkTimer = null;
            }
            // 隐藏调试点
            debugCanvas.Visibility = Visibility.Collapsed;
            StopPositionTracking();
            // 停止后重新注册热键
            //UnregisterHotkeys();
            RegisterHotkeys();

        }
        private const int SW_RESTORE = 9;
        [DllImport("user32.dll")]
        private static extern bool ClientToScreen(IntPtr hWnd, ref Point lpPoint);
        private void CheckRules(object sender, EventArgs e)
        {
            if (_selectedWindowHandle == IntPtr.Zero) return;
            // 确保游戏窗口激活
            SetForegroundWindow(_selectedWindowHandle);
            ShowWindow(_selectedWindowHandle, SW_RESTORE);

            // 添加短暂延迟确保窗口激活
            Thread.Sleep(50);

            // 获取窗口位置
            GetWindowRect(_selectedWindowHandle, out RECT windowRect);
            Debug.WriteLine($"窗口位置: L={windowRect.Left}, T={windowRect.Top}, R={windowRect.Right}, B={windowRect.Bottom}");
            var clientOffset = Win32PointHelper.GetClientTopLeft(_selectedWindowHandle);
            foreach (var rule in Rules.Where(r => r.IsEnabled))
            {
                Debug.WriteLine($"检查规则: {rule.Key}, 条件数={rule.Conditions.Count}");
                bool skipRule = false;
                bool allConditionsMet = true;
                int conditionIndex = 0;
                // 检查所有条件是否都满足
                foreach (var condition in rule.Conditions)
                {
                    // 计算绝对坐标

                    int absoluteX = windowRect.Left + clientOffset.X + condition.OffsetX;
                    int absoluteY = windowRect.Top + clientOffset.Y + condition.OffsetY;

                    // 获取该点颜色
                    var pixelColor = Win32ColorHelper.GetPixelColor(absoluteX, absoluteY);

                    // 检查颜色是否匹配（考虑容差）
                    bool isMatch = IsColorMatch(pixelColor, condition.TargetColor, condition.Tolerance);
                    // 在条件检查后立即记录结果
                    Debug.WriteLine($"条件#{conditionIndex}: 位置({absoluteX},{absoluteY}) " +
                                   $"实际颜色:{pixelColor} 目标颜色:{condition.TargetColor} " +
                                   $"容差:{condition.Tolerance} 匹配:{isMatch}");
                    Debug.WriteLine($"窗口位置: L{windowRect.Left} T{windowRect.Top}");
                    Debug.WriteLine($"客户区偏移: X{clientOffset.X} Y{clientOffset.Y}");
                    Debug.WriteLine($"条件偏移: X{condition.OffsetX} Y{condition.OffsetY}");
                    Debug.WriteLine($"绝对坐标: X{absoluteX} Y{absoluteY}");
                    // 检查跳过条件
                    if (condition.IsSkipCondition && isMatch)
                    {
                        skipRule = true;
                        Debug.WriteLine($"跳过规则 {rule.Key}，跳过条件满足");
                        break;
                    }
                    // 检查普通条件
                    if (!condition.IsSkipCondition && !isMatch)
                    {
                        allConditionsMet = false;
                        break;
                    }
                    conditionIndex++;
                }
                // 跳过规则或条件不满足
                if (skipRule || !allConditionsMet) continue;
                // 执行按键
                PressKey(rule.Key);

                //// 根据规则频率延迟
                //if (rule.CheckInterval > 0)
                //{
                //    Thread.Sleep(rule.CheckInterval);
                //}
                // 只执行匹配的第一个规则
                break;
            }
        }

        private void PressKey(string key)
        {
            if (string.IsNullOrEmpty(key)) return;

            try
            {
                // 处理鼠标左键
                if (key == "鼠标左键")
                {
                    _inputSimulator.Mouse.LeftButtonClick();
                    return;
                }
                // 在PressKey方法中添加
                if (key == "鼠标滚轮上")
                    _inputSimulator.Mouse.VerticalScroll(1);
                else if (key == "鼠标滚轮下")
                    _inputSimulator.Mouse.VerticalScroll(-1);
                // 模拟按键
                _inputSimulator.Keyboard.KeyPress((User32.VK)Enum.Parse(typeof(User32.VK), $"VK_{key}"));
            }
            catch
            {
                // 处理特殊键
                switch (key.ToUpper())
                {
                    case "Q":
                        _inputSimulator.Keyboard.KeyPress(User32.VK.VK_Q);
                        break;
                    case "E":
                        _inputSimulator.Keyboard.KeyPress(User32.VK.VK_E);
                        break;
                    case "Z":
                        _inputSimulator.Keyboard.KeyPress(User32.VK.VK_Z);
                        break;
                    case "X":
                        _inputSimulator.Keyboard.KeyPress(User32.VK.VK_X);
                        break;
                    case "H":
                        _inputSimulator.Keyboard.KeyPress(User32.VK.VK_H);
                        break;
                    default:
                        if (int.TryParse(key, out int num))
                        {
                            _inputSimulator.Keyboard.KeyPress((User32.VK)(0x30 + num)); // 0x30 = '0'
                        }
                        break;
                }
            }
        }

        private bool IsColorMatch(Color color1, Color color2, int tolerance)
        {
            // 计算各通道差异
            int diffR = Math.Abs(color1.R - color2.R);
            int diffG = Math.Abs(color1.G - color2.G);
            int diffB = Math.Abs(color1.B - color2.B);

            // 检查每个通道是否在容差范围内
            return diffR <= tolerance && diffG <= tolerance && diffB <= tolerance;
        }

        public Color GetPixelColor(int x, int y)
        {
            IntPtr hdc = GetDC(IntPtr.Zero);
            uint pixel = GetPixel(hdc, x, y);
            ReleaseDC(IntPtr.Zero, hdc);

            // 修复颜色转换错误
            return Color.FromArgb(
                (int)(pixel & 0x000000FF),
                (int)((pixel & 0x0000FF00) >> 8),
                (int)((pixel & 0x00FF0000) >> 16));
        }

        private Color ParseColor(string colorString)
        {
            try
            {
                if (colorString.StartsWith("#"))
                {
                    return ColorTranslator.FromHtml(colorString);
                }
                return Color.FromName(colorString);
            }
            catch
            {
                return Color.Red;
            }
        }

        private void SaveConfig_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 创建序列化选项
                var jsonOptions = new JsonSerializerOptions
                {
                    WriteIndented = true,
                    Converters = { new ColorJsonConverter() }  // 添加自定义转换器
                };

                // 序列化配置
                string json = JsonSerializer.Serialize(new ConfigData
                {
                    Rules = Rules,
                    StartHotKey = _startHotKey,
                    StopHotKey = _stopHotKey,
                    SelectedWindowTitle = (cbGameWindows.SelectedItem as WindowInfo)?.Title
                }, jsonOptions);

                // 保存热键设置
                _startHotKey = (cbStartHotKey.SelectedItem as ComboBoxItem)?.Content.ToString();
                _stopHotKey = (cbStopHotKey.SelectedItem as ComboBoxItem)?.Content.ToString();
                var saveDialog = new Microsoft.Win32.SaveFileDialog
                {
                    Filter = "配置文件|*.json",
                    FileName = "skill_config.json",
                    DefaultExt = ".json"
                };

                if (saveDialog.ShowDialog() == true)
                {
                    var config = new
                    {
                        Rules,
                        StartHotKey = _startHotKey,
                        StopHotKey = _stopHotKey,
                        SelectedWindowTitle = (cbGameWindows.SelectedItem as WindowInfo)?.Title // 新增
                    };

                    var options = new JsonSerializerOptions { WriteIndented = true };
                    json = JsonSerializer.Serialize(config, options);
                    File.WriteAllText(saveDialog.FileName, json);
                    tbStatus.Text = $"配置已保存到 {saveDialog.FileName}";
                    // 添加文件刷新
                    //File.WriteAllText(saveDialog.FileName, string.Empty);
                    //File.WriteAllText(saveDialog.FileName, json);
                    //File.SetAttributes(saveDialog.FileName, FileAttributes.Normal);
                    // 重新注册热键
                    //UnregisterHotkeys();
                    RegisterHotkeys();
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"保存配置失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadConfig_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 创建序列化选项
                var jsonOptions = new JsonSerializerOptions
                {
                    Converters = { new ColorJsonConverter() }  // 添加自定义转换器
                };
                //UnregisterHotkeys();
                var openDialog = new Microsoft.Win32.OpenFileDialog
                {
                    Filter = "配置文件|*.json",
                    DefaultExt = ".json"
                };

                if (openDialog.ShowDialog() == true)
                {
                    string json = File.ReadAllText(openDialog.FileName);
                    var config = JsonSerializer.Deserialize<ConfigData>(json, jsonOptions);

                    Rules.Clear();
                    foreach (var rule in config.Rules)
                    {
                        Rules.Add(rule);
                    }

                    // 加载热键设置
                    _startHotKey = config.StartHotKey;
                    _stopHotKey = config.StopHotKey;

                    // 更新UI
                    SetComboBoxSelection(cbStartHotKey, _startHotKey);
                    SetComboBoxSelection(cbStopHotKey, _stopHotKey);
                    if (!string.IsNullOrEmpty(config.SelectedWindowTitle))
                    {
                        bool found = false;
                        foreach (var item in cbGameWindows.Items)
                        {
                            if (item is WindowInfo window &&
                                window.Title == config.SelectedWindowTitle)
                            {
                                cbGameWindows.SelectedItem = item;
                                _selectedWindowHandle = window.Handle;
                                tbStatus.Text = $"已选择窗口: {window.Title}";
                                found = true;
                                break;
                            }
                        }

                        if (!found)
                        {
                            // 尝试刷新窗口列表
                            RefreshWindows();
                            foreach (var item in cbGameWindows.Items)
                            {
                                if (item is WindowInfo window &&
                                    window.Title == config.SelectedWindowTitle)
                                {
                                    cbGameWindows.SelectedItem = item;
                                    _selectedWindowHandle = window.Handle;
                                    tbStatus.Text = $"已恢复窗口: {window.Title}";
                                    break;
                                }
                            }
                        }
                    }
                    // 延迟热键注册直到窗口就绪
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        try
                        {
                            //UnregisterHotkeys(); // 双重确保
                            RegisterHotkeys();
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"热键注册失败: {ex.Message}");
                        }
                    }), DispatcherPriority.ApplicationIdle);
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"加载配置失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        // 添加检查窗口句柄的方法
        private bool IsWindowHandleValid()
        {
            var helper = new WindowInteropHelper(this);
            return helper.Handle != IntPtr.Zero;
        }
        private void LoadConfig()
        {
            try
            {
                if (File.Exists(_configFilePath))
                {
                    // 创建带转换器的序列化选项
                    var jsonOptions = new JsonSerializerOptions
                    {
                        Converters = { new ColorJsonConverter() }
                    };
                    string json = File.ReadAllText(_configFilePath);
                    var config = JsonSerializer.Deserialize<ConfigData>(json, jsonOptions);

                    Rules.Clear();
                    foreach (var rule in config.Rules)
                    {
                        Rules.Add(rule);
                    }

                    // 加载热键设置
                    _startHotKey = config.StartHotKey;
                    _stopHotKey = config.StopHotKey;

                    // 更新UI
                    SetComboBoxSelection(cbStartHotKey, _startHotKey);
                    SetComboBoxSelection(cbStopHotKey, _stopHotKey);
                    if (!string.IsNullOrEmpty(config.SelectedWindowTitle))
                    {
                        // 等待窗口列表刷新完成
                        Dispatcher.Invoke(() =>
                        {
                            foreach (var item in cbGameWindows.Items)
                            {
                                if (item is WindowInfo window &&
                                    window.Title == config.SelectedWindowTitle)
                                {
                                    cbGameWindows.SelectedItem = item;
                                    _selectedWindowHandle = window.Handle;
                                    tbStatus.Text = $"已选择窗口: {window.Title}";
                                    break;
                                }
                            }

                            // 如果没有找到匹配窗口，尝试重新查找
                            if (cbGameWindows.SelectedItem == null)
                            {
                                RefreshWindows();
                                foreach (var item in cbGameWindows.Items)
                                {
                                    if (item is WindowInfo window &&
                                        window.Title == config.SelectedWindowTitle)
                                    {
                                        cbGameWindows.SelectedItem = item;
                                        _selectedWindowHandle = window.Handle;
                                        tbStatus.Text = $"已恢复窗口: {window.Title}";
                                        break;
                                    }
                                }
                            }
                        });
                    }
                    tbStatus.Text = "配置已自动加载";
                }
                //else
                //{
                //    InitializeRules();
                //}
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"加载配置失败: {ex.Message}");
                //InitializeRules();
            }
        }
        private void SetComboBoxSelection(System.Windows.Controls.ComboBox comboBox, string value)
        {
            foreach (ComboBoxItem item in comboBox.Items)
            {
                if (item.Content.ToString() == value)
                {
                    comboBox.SelectedItem = item;
                    return;
                }
            }
            // 如果找不到匹配项，默认选择波浪号
            comboBox.SelectedIndex = 7; // 波浪号在列表中的位置
        }
        private void btnTestConditions_Click(object sender, RoutedEventArgs e)
        {
            if (dgRules.SelectedItem is SkillRule selectedRule)
            {
                if (_selectedWindowHandle == IntPtr.Zero)
                {
                    System.Windows.MessageBox.Show("请先选择游戏窗口");
                    return;
                }

                // 获取窗口位置
                GetWindowRect(_selectedWindowHandle, out RECT windowRect);

                var result = new StringBuilder();
                result.AppendLine($"测试规则: {selectedRule.Key}");
                result.AppendLine($"条件数量: {selectedRule.Conditions.Count}");

                int index = 0;
                var clientOffset = Win32PointHelper.GetClientTopLeft(_selectedWindowHandle);
                foreach (var condition in selectedRule.Conditions)
                {
                    int absoluteX = windowRect.Left + clientOffset.X + condition.OffsetX;
                    int absoluteY = windowRect.Top + clientOffset.Y + condition.OffsetY;

                    // 获取颜色
                    var pixelColor = Win32ColorHelper.GetPixelColor(absoluteX, absoluteY);

                    // 检查匹配
                    bool isMatch = IsColorMatch(pixelColor, condition.TargetColor, condition.Tolerance);

                    result.AppendLine($"条件 #{index}:");
                    result.AppendLine($"  位置: ({absoluteX}, {absoluteY})");
                    result.AppendLine($"  目标颜色: R:{condition.TargetColor.R} G:{condition.TargetColor.G} B:{condition.TargetColor.B}");
                    result.AppendLine($"  实际颜色: R:{pixelColor.R} G:{pixelColor.G} B:{pixelColor.B}");
                    result.AppendLine($"  容差: {condition.Tolerance}");
                    result.AppendLine($"  跳过规则: {condition.IsSkipCondition}");
                    result.AppendLine($"  匹配: {isMatch}");
                    result.AppendLine();

                    index++;
                }

                System.Windows.MessageBox.Show(result.ToString(), "条件测试结果");
            }
        }

        protected override void OnClosed(EventArgs e)
        {
            //CleanupHwndSource();
            HotkeyHolder.UnregisterHotKey();
            StopPositionTracking();
            if (_trayIcon != null)
            {
                _trayIcon.Dispose();
            }

            // 保存配置
            try
            {
                var config = new ConfigData
                {
                    Rules = Rules,
                    StartHotKey = _startHotKey,
                    StopHotKey = _stopHotKey,
                    SelectedWindowTitle = (cbGameWindows.SelectedItem as WindowInfo)?.Title
                };
                // 确保目录存在
                Directory.CreateDirectory(System.IO.Path.GetDirectoryName(_configFilePath));
                var options = new JsonSerializerOptions { WriteIndented = true };
                string json = JsonSerializer.Serialize(config, options);
                File.WriteAllText(_configFilePath, json);
            }
            catch
            {
                // 忽略保存错误
            }
            base.OnClosed(e);
        }

        private void AddCondition_Click(object sender, RoutedEventArgs e)
        {
            if (dgRules.SelectedItem is SkillRule selectedRule)
            {
                var condition = new SkillCondition
                {
                    OffsetX = 0,
                    OffsetY = 0,
                    TargetColor = Color.Blue,
                    Tolerance = 10
                };

                selectedRule.Conditions.Add(condition);
            }
        }
        private void cbGameWindows_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cbGameWindows.SelectedItem is WindowInfo selectedWindow)
            {
                _selectedWindowHandle = selectedWindow.Handle;
                tbStatus.Text = $"已选择窗口: {selectedWindow.Title}";
            }
        }

        #region Win32 API
        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder text, int count);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsWindowVisible(IntPtr hWnd);
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        private static extern IntPtr GetWindowDC(IntPtr hWnd);

        [DllImport("gdi32.dll")]
        private static extern uint GetPixel(IntPtr hdc, int x, int y);

        [DllImport("user32.dll")]
        private static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);

        [DllImport("user32.dll")]
        private static extern IntPtr GetDC(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        private string GetProcessName(IntPtr hwnd)
        {
            GetWindowThreadProcessId(hwnd, out uint processId);
            try
            {
                Process process = Process.GetProcessById((int)processId);
                return process.ProcessName;
            }
            catch
            {
                return string.Empty;
            }
        }
        private void HotKeyComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateHotkey();
        }
        private void CheckBox_CheckedChanged(object sender, RoutedEventArgs e)
        {
            UpdateHotkey();
        }
        private void UpdateHotkey()
        {
            // 确保窗口已加载完成
            if (!IsLoaded) return;
            // 先注销当前所有热键
            HotkeyHolder.UnregisterHotKey();
            // 构建正确的热键格式
            _startHotKey = BuildHotkeyString(true);
            _stopHotKey = BuildHotkeyString(false);

            // 重新注册热键
            RegisterHotkeys();
        }
        private string BuildHotkeyString(bool isStart)
        {
            StringBuilder builder = new StringBuilder();

            // 添加空格确保正确解析
            if (isStart ? chkstartCtrl.IsChecked == true : chkstopCtrl.IsChecked == true)
                builder.Append("Ctrl + ");

            if (isStart ? chkstartShift.IsChecked == true : chkstopShift.IsChecked == true)
                builder.Append("Shift + ");

            if (isStart ? chkstartAlt.IsChecked == true : chkstopAlt.IsChecked == true)
                builder.Append("Alt + ");

            // 添加主键（使用修正后的格式）
            var combo = isStart ? cbStartHotKey : cbStopHotKey;
            if (combo.SelectedItem is ComboBoxItem item)
            {
                builder.Append(item.Content.ToString());
            }

            return builder.ToString().TrimEnd(' ', '+'); // 清理尾部多余字符
        }

        #endregion
        private void TestMouseSideButtons()
        {
            Debug.WriteLine("[TEST] 开始测试鼠标侧键支持");

            // 测试系统是否支持鼠标侧键
            int xButtonCount = SystemInformation.MouseButtons;
            Debug.WriteLine($"[TEST] 系统报告的鼠标按键数量: {xButtonCount}");

            // 测试鼠标侧键的虚拟键码
            bool xButton1Supported = User32.GetKeyState((int)Keys.XButton1) != -1;
            bool xButton2Supported = User32.GetKeyState((int)Keys.XButton2) != -1;

            Debug.WriteLine($"[TEST] XButton1 支持: {xButton1Supported}");
            Debug.WriteLine($"[TEST] XButton2 支持: {xButton2Supported}");
        }
        private void TestMouseSideKeyRegistration()
        {
            Debug.WriteLine("=== 开始测试鼠标侧键注册 ===");

            try
            {
                // 测试直接注册
                var testHook = new HotkeyHook();
                testHook.KeyPressed += (s, e) => {
                    Debug.WriteLine($"[TEST] 直接注册收到事件: {e.Key}");
                };

                testHook.RegisterHotKey(User32.HotKeyModifiers.MOD_NONE, Keys.XButton1);
                Debug.WriteLine("[TEST] 直接注册成功");

                Thread.Sleep(5000); // 等待5秒测试

                testHook.UnregisterHotKey();
                testHook.Dispose();
                Debug.WriteLine("[TEST] 直接注册测试完成");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[TEST] 直接注册失败: {ex.Message}");
            }
        }

        private void TestSingleMouseKey(string keyName, Keys key)
        {
            try
            {
                // 直接测试底层注册
                var testHook = new HotkeyHook();
                testHook.RegisterHotKey(User32.HotKeyModifiers.MOD_NONE, key);
                Debug.WriteLine($"[SUCCESS] {keyName} 底层注册成功");
                testHook.UnregisterHotKey();
                testHook.Dispose();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[FAILED] {keyName} 底层注册失败: {ex.Message}");
            }

            try
            {
                // 测试通过 HotkeyHolder 注册
                HotkeyHolder.RegisterHotKey(keyName, (s, e) => {
                    Debug.WriteLine($"[EVENT] {keyName} 被按下");
                });
                Debug.WriteLine($"[SUCCESS] {keyName} HotkeyHolder 注册成功");
                HotkeyHolder.UnregisterHotKey();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[FAILED] {keyName} HotkeyHolder 注册失败: {ex.Message}");
            }
        }

        private void TestCombinationKey(string hotkey)
        {
            try
            {
                HotkeyHolder.RegisterHotKey(hotkey, (s, e) => {
                    Debug.WriteLine($"[EVENT] {hotkey} 被按下");
                });
                Debug.WriteLine($"[SUCCESS] {hotkey} 注册成功");
                HotkeyHolder.UnregisterHotKey();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[FAILED] {hotkey} 注册失败: {ex.Message}");
            }
        }

        private void btnTestMouseKeys_Click(object sender, RoutedEventArgs e)
        {
            TestMouseSideButtons();
            TestMouseSideKeyRegistration();
        }
    }

    public class WindowInfo
    {
        public IntPtr Handle { get; set; }
        public string Title { get; set; }
    }

    public class SkillRule : INotifyPropertyChanged
    {
        // 添加属性标识是否是鼠标操作
        [JsonIgnore]
        public bool IsMouseAction => Key == "鼠标左键";
        private bool _isEnabled = true;

        public bool IsEnabled
        {
            get => _isEnabled;
            set
            {
                _isEnabled = value;
                OnPropertyChanged(nameof(IsEnabled));
            }
        }

        public string Key { get; set; }
        //public int CheckInterval { get; set; }
        public ObservableCollection<SkillCondition> Conditions { get; set; } = new ObservableCollection<SkillCondition>();

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public class SkillCondition : INotifyPropertyChanged
    {
        private bool _isSkipCondition;
        public bool IsSkipCondition
        {
            get => _isSkipCondition;
            set { _isSkipCondition = value; OnPropertyChanged(nameof(IsSkipCondition)); }
        }

        private int _offsetX;
        private int _offsetY;
        private Color _targetColor;
        private int _tolerance = 10;
        [JsonPropertyName("x")]  // 简化属性名
        public int OffsetX
        {
            get => _offsetX;
            set { _offsetX = value; OnPropertyChanged(nameof(OffsetX)); }
        }

        [JsonPropertyName("y")]  // 简化属性名
        public int OffsetY
        {
            get => _offsetY;
            set { _offsetY = value; OnPropertyChanged(nameof(OffsetY)); }
        }

        [JsonPropertyName("color")]  // 简化属性名
        public Color TargetColor
        {
            get => _targetColor;
            set
            {
                _targetColor = value;
                OnPropertyChanged(nameof(TargetColor));
                OnPropertyChanged(nameof(TargetColorHex));
                OnPropertyChanged(nameof(ColorBrush));
                OnPropertyChanged(nameof(TextColor));
            }
        }

        [JsonPropertyName("tol")]  // 简化属性名
        public int Tolerance
        {
            get => _tolerance;
            set { _tolerance = value; OnPropertyChanged(nameof(Tolerance)); }
        }

        [JsonIgnore]
        public string TargetColorHex => $"#{TargetColor.R:X2}{TargetColor.G:X2}{TargetColor.B:X2}";
        [JsonIgnore]
        public System.Windows.Media.Brush ColorBrush =>
            new SolidColorBrush(System.Windows.Media.Color.FromRgb(
                TargetColor.R, TargetColor.G, TargetColor.B));
        [JsonIgnore]
        public System.Windows.Media.Brush TextColor =>
            (TargetColor.R * 0.299 + TargetColor.G * 0.587 + TargetColor.B * 0.114) > 150 ?
                System.Windows.Media.Brushes.Black :
                System.Windows.Media.Brushes.White;
        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

    }
    public class ColorJsonConverter : JsonConverter<Color>
    {
        public override Color Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
        {
            // 先尝试读取字符串格式
            if (reader.TokenType == JsonTokenType.String)
            {
                string hex = reader.GetString();
                if (hex.StartsWith("#") && hex.Length == 7)
                {
                    return Color.FromArgb(
                        int.Parse(hex.Substring(1, 2), System.Globalization.NumberStyles.HexNumber),
                        int.Parse(hex.Substring(3, 2), System.Globalization.NumberStyles.HexNumber),
                        int.Parse(hex.Substring(5, 2), System.Globalization.NumberStyles.HexNumber));
                }
            }
            // 处理对象格式 { "R": 255, "G": 0, "B": 0 }
            else if (reader.TokenType == JsonTokenType.StartObject)
            {
                int r = 0, g = 0, b = 0;
                while (reader.Read())
                {
                    if (reader.TokenType == JsonTokenType.EndObject)
                        break;

                    if (reader.TokenType == JsonTokenType.PropertyName)
                    {
                        var propName = reader.GetString();
                        reader.Read();
                        switch (propName)
                        {
                            case "R": r = reader.GetInt32(); break;
                            case "G": g = reader.GetInt32(); break;
                            case "B": b = reader.GetInt32(); break;
                        }
                    }
                }
                return Color.FromArgb(r, g, b);
            }

            throw new JsonException("Invalid color format");
        }

        public override void Write(Utf8JsonWriter writer, Color value, JsonSerializerOptions options)
        {
            writer.WriteStringValue($"#{value.R:X2}{value.G:X2}{value.B:X2}");
        }
    }
}