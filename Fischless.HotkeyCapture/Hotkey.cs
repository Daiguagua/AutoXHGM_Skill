using System.Diagnostics;
using System.Windows.Forms;
using Vanara.PInvoke;

namespace Fischless.HotkeyCapture;

public sealed class Hotkey
{
    public bool Alt { get; set; }
    public bool Control { get; set; }
    public bool Shift { get; set; }
    public bool Windows { get; set; }

    private Keys key;

    public Keys Key
    {
        get => key;
        set
        {
            if (value != Keys.ControlKey && value != Keys.Alt && value != Keys.Menu && value != Keys.ShiftKey)
            {
                key = value;
            }
            else
            {
                key = Keys.None;
            }
        }
    }

    public User32.HotKeyModifiers ModifierKey =>
        (Windows ? User32.HotKeyModifiers.MOD_WIN : User32.HotKeyModifiers.MOD_NONE) |
        (Control ? User32.HotKeyModifiers.MOD_CONTROL : User32.HotKeyModifiers.MOD_NONE) |
        (Shift ? User32.HotKeyModifiers.MOD_SHIFT : User32.HotKeyModifiers.MOD_NONE) |
        (Alt ? User32.HotKeyModifiers.MOD_ALT : User32.HotKeyModifiers.MOD_NONE);

    public Hotkey()
    {
        Reset();
    }

    public Hotkey(string hotkeyStr)
    {
        Debug.WriteLine($"[DEBUG] 解析热键字符串: '{hotkeyStr}'");

        try
        {
            // 支持无修饰键的单个按键
            if (!hotkeyStr.Contains('+'))
            {
                string mappedKey = MapSpecialKey(hotkeyStr);
                Debug.WriteLine($"[DEBUG] 单键映射: '{hotkeyStr}' -> '{mappedKey}'");
                Key = (Keys)Enum.Parse(typeof(Keys), mappedKey);
                Debug.WriteLine($"[DEBUG] 最终Key值: {Key}");
                return;
            }

            string[] keyStrs = hotkeyStr.Replace(" ", string.Empty).Split('+');
            Debug.WriteLine($"[DEBUG] 分割后的按键: [{string.Join(", ", keyStrs)}]");

            bool mainKeyFound = false;
            foreach (string keyStr in keyStrs)
            {
                Debug.WriteLine($"[DEBUG] 处理按键部分: '{keyStr}'");

                if (keyStr.Equals("Win", StringComparison.OrdinalIgnoreCase))
                {
                    Windows = true;
                    Debug.WriteLine("[DEBUG] 设置 Windows 修饰键");
                }
                else if (keyStr.Equals("Ctrl", StringComparison.OrdinalIgnoreCase))
                {
                    Control = true;
                    Debug.WriteLine("[DEBUG] 设置 Control 修饰键");
                }
                else if (keyStr.Equals("Shift", StringComparison.OrdinalIgnoreCase))
                {
                    Shift = true;
                    Debug.WriteLine("[DEBUG] 设置 Shift 修饰键");
                }
                else if (keyStr.Equals("Alt", StringComparison.OrdinalIgnoreCase))
                {
                    Alt = true;
                    Debug.WriteLine("[DEBUG] 设置 Alt 修饰键");
                }
                else
                {
                    string mappedKey = MapSpecialKey(keyStr);
                    Debug.WriteLine($"[DEBUG] 主键映射: '{keyStr}' -> '{mappedKey}'");
                    Key = (Keys)Enum.Parse(typeof(Keys), mappedKey);
                    Debug.WriteLine($"[DEBUG] 最终Key值: {Key}");
                    mainKeyFound = true;
                }
            }

            if (!mainKeyFound)
            {
                throw new ArgumentException("No main key specified");
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[ERROR] 热键解析失败: {ex.Message}");
            throw new ArgumentException("Invalid Hotkey");
        }
    }


    public override string ToString()
    {
        string str = string.Empty;
        if (Key != Keys.None)
        {
            str = string.Format("{0}{1}{2}{3}{4}",
                Windows ? "Win + " : string.Empty,
                Control ? "Ctrl + " : string.Empty,
                Shift ? "Shift + " : string.Empty,
                Alt ? "Alt + " : string.Empty,
                Key);
        }
        return str;
    }
    public static string MapSpecialKey(string keyStr)
    {
        return keyStr switch
        {
            // 数字键
            "0" => "D0",
            "1" => "D1",
            "2" => "D2",
            "3" => "D3",
            "4" => "D4",
            "5" => "D5",
            "6" => "D6",
            "7" => "D7",
            "8" => "D8",
            "9" => "D9",

            // 符号键
            ";" => "Oem1",
            "?" => "Oem2",
            "~" => "Oem3",
            "[" => "Oem4",
            "\\" => "Oem5",
            "]" => "Oem6",
            "," => "OemComma",
            "." => "OemPeriod",
            "=" => "OemPlus",
            "-" => "OemMinus",

            // 鼠标键
            "鼠标中键" => "MButton",
            "鼠标侧键1" => "XButton1",
            "鼠标侧键2" => "XButton2",

            // 小键盘
            "小键盘/" => "Divide",
            "小键盘*" => "Multiply",
            "小键盘-" => "Subtract",
            "小键盘+" => "Add",
            "小键盘." => "Decimal",
            "小键盘0" => "NumPad0",
            "小键盘1" => "NumPad1",
            "小键盘2" => "NumPad2",
            "小键盘3" => "NumPad3",
            "小键盘4" => "NumPad4",
            "小键盘5" => "NumPad5",
            "小键盘6" => "NumPad6",
            "小键盘7" => "NumPad7",
            "小键盘8" => "NumPad8",
            "小键盘9" => "NumPad9",

            // 字母键
            "A" => "A",
            "B" => "B",
            "C" => "C",
            "D" => "D",
            "E" => "E",
            "F" => "F",
            "G" => "G",
            "H" => "H",
            "I" => "I",
            "J" => "J",
            "K" => "K",
            "L" => "L",
            "M" => "M",
            "N" => "N",
            "O" => "O",
            "P" => "P",
            "Q" => "Q",
            "R" => "R",
            "S" => "S",
            "T" => "T",
            "U" => "U",
            "V" => "V",
            "W" => "W",
            "X" => "X",
            "Y" => "Y",
            "Z" => "Z",

            _ => keyStr // 默认返回原值
        };
    }
    public void Reset()
    {
        Alt = false;
        Control = false;
        Shift = false;
        Windows = false;
        Key = Keys.None;
    }
}
