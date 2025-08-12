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
        try
        {
            string[] keyStrs = hotkeyStr.Replace(" ", string.Empty).Split('+');
            bool mainKeyFound = false;
            foreach (string keyStr in keyStrs)
            {
                if (keyStr.Equals("Win", StringComparison.OrdinalIgnoreCase))
                {
                    Windows = true;
                }
                else if (keyStr.Equals("Ctrl", StringComparison.OrdinalIgnoreCase))
                {
                    Control = true;
                }
                else if (keyStr.Equals("Shift", StringComparison.OrdinalIgnoreCase))
                {
                    Shift = true;
                }
                else if (keyStr.Equals("Alt", StringComparison.OrdinalIgnoreCase))
                {
                    Alt = true;
                }
                else
                {
                    // 添加特殊键映射
                    string mappedKey = MapSpecialKey(keyStr);
                    Key = (Keys)Enum.Parse(typeof(Keys), mappedKey);
                    mainKeyFound = true;
                }
            }
            // 确保有主按键
            if (!mainKeyFound)
            {
                throw new ArgumentException("No main key specified");
            }
        }
        catch
        {
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
            "~" => "Oem3",
            ";" => "Oem1",
            "?" => "Oem2",
            "[" => "Oem4",
            "\\" => "Oem5",
            "]" => "Oem6",
            "," => "OemComma",
            "." => "OemPeriod",
            "=" => "OemPlus",
            "-" => "OemMinus",
            "鼠标中键" => "MButton",
            "鼠标侧键1" => "XButton1",
            "鼠标侧键2" => "XButton2",
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
