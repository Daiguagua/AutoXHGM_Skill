namespace Fischless.HotkeyCapture;
using System.Windows.Forms;

public sealed class HotkeyHolder
{
    private static Hotkey? hotkey;
    private static HotkeyHook? hotkeyHook;
    private static Action<object?, KeyPressedEventArgs>? keyPressed;

    public static void RegisterHotKey(string hotkeyStr, Action<object?, KeyPressedEventArgs> keyPressed = null!)
    {
        if (string.IsNullOrEmpty(hotkeyStr))
        {
            UnregisterHotKey();
            return;
        }
        // 验证热键格式
        if (!IsValidHotkey(hotkeyStr))
        {
            throw new ArgumentException($"Invalid hotkey format: {hotkeyStr}");
        }
        hotkey = new Hotkey(hotkeyStr);

        hotkeyHook?.Dispose();
        hotkeyHook = new HotkeyHook();
        hotkeyHook.KeyPressed -= OnKeyPressed;
        hotkeyHook.KeyPressed += OnKeyPressed;
        HotkeyHolder.keyPressed = keyPressed;
        hotkeyHook.RegisterHotKey(hotkey.ModifierKey, hotkey.Key);
    }
    private static bool IsValidHotkey(string hotkeyStr)
    {
        string[] parts = hotkeyStr.Split('+');
        bool hasMainKey = false;

        foreach (string part in parts)
        {
            if (part.Equals("Win", StringComparison.OrdinalIgnoreCase) ||
                part.Equals("Ctrl", StringComparison.OrdinalIgnoreCase) ||
                part.Equals("Shift", StringComparison.OrdinalIgnoreCase) ||
                part.Equals("Alt", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            // 检查是否是有效主键
            try
            {
                string mappedKey = Hotkey.MapSpecialKey(part);
                if (Enum.IsDefined(typeof(Keys), mappedKey))
                {
                    hasMainKey = true;
                }
            }
            catch
            {
                return false;
            }
        }

        return hasMainKey;
    }
    public static void UnregisterHotKey()
    {
        if (hotkeyHook != null)
        {
            hotkeyHook.KeyPressed -= OnKeyPressed;
            hotkeyHook.UnregisterHotKey();
            hotkeyHook.Dispose();
        }
    }

    private static void OnKeyPressed(object? sender, KeyPressedEventArgs e)
    {
        keyPressed?.Invoke(sender, e);
    }
}
