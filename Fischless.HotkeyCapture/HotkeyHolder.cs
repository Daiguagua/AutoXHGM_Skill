namespace Fischless.HotkeyCapture;

using Skill_Loop.Fischless.HotkeyCapture;
using System.Diagnostics;
using System.Windows.Forms;
using Vanara.PInvoke;

public sealed class HotkeyHolder
{

    private static readonly Dictionary<string, (Hotkey hotkey, HotkeyHook hook, Action<object?, KeyPressedEventArgs> callback)>
        _registeredHotkeys = new Dictionary<string, (Hotkey, HotkeyHook, Action<object?, KeyPressedEventArgs>)>();

    private static readonly Dictionary<string, (MouseButton button, MouseHook hook, Action<object?, KeyPressedEventArgs> callback)>
        _registeredMouseHotkeys = new Dictionary<string, (MouseButton, MouseHook, Action<object?, KeyPressedEventArgs>)>();

    private static MouseHook? _globalMouseHook;
    public static void RegisterHotKey(string hotkeyStr, Action<object?, KeyPressedEventArgs> keyPressed = null!)
    {
        if (string.IsNullOrEmpty(hotkeyStr))
        {
            UnregisterHotKey();
            return;
        }
        if (IsMouseHotkey(hotkeyStr))
        {
            RegisterMouseHotkey(hotkeyStr, keyPressed);
            return;
        }
        // 如果已存在，先注销
        if (_registeredHotkeys.ContainsKey(hotkeyStr))
        {
            UnregisterHotKey(hotkeyStr);
        }

        try
        {
            Debug.WriteLine($"[HOTKEY] 开始注册键盘热键: {hotkeyStr}");

            var hotkey = new Hotkey(hotkeyStr);
            var hotkeyHook = new HotkeyHook();

            EventHandler<KeyPressedEventArgs> handler = (sender, e) =>
            {
                Debug.WriteLine($"[HOTKEY] 键盘热键触发: {e.Key}, Modifier: {e.Modifier}");
                keyPressed?.Invoke(sender, e);
            };

            hotkeyHook.KeyPressed += handler;
            hotkeyHook.RegisterHotKey(hotkey.ModifierKey, hotkey.Key);

            _registeredHotkeys[hotkeyStr] = (hotkey, hotkeyHook, keyPressed);
            Debug.WriteLine($"[HOTKEY] 键盘热键注册成功: {hotkeyStr}");
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[HOTKEY] 键盘热键注册失败: {hotkeyStr}, 错误: {ex.Message}");
            throw;
        }
    }
    private static void RegisterMouseHotkey(string hotkeyStr, Action<object?, KeyPressedEventArgs> keyPressed)
    {
        try
        {
            Debug.WriteLine($"[HOTKEY] 开始注册鼠标热键: {hotkeyStr}");

            var mouseButton = GetMouseButtonFromString(hotkeyStr);

            if (_globalMouseHook == null)
            {
                _globalMouseHook = new MouseHook();
            }

            EventHandler<MouseKeyPressedEventArgs> mouseHandler = (sender, e) =>
            {
                if (e.Button == mouseButton)
                {
                    Debug.WriteLine($"[HOTKEY] 鼠标热键触发: {e.Button}");
                    var keyPressedArgs = new KeyPressedEventArgs(
                        User32.HotKeyModifiers.MOD_NONE,
                        GetKeysFromMouseButton(e.Button));
                    keyPressed?.Invoke(sender, keyPressedArgs);
                }
            };

            _globalMouseHook.MouseKeyPressed += mouseHandler;
            _globalMouseHook.RegisterMouseButton(mouseButton);

            _registeredMouseHotkeys[hotkeyStr] = (mouseButton, _globalMouseHook, keyPressed);
            Debug.WriteLine($"[HOTKEY] 鼠标热键注册成功: {hotkeyStr}");
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[HOTKEY] 鼠标热键注册失败: {hotkeyStr}, 错误: {ex.Message}");
            throw;
        }
    }
    private static bool IsMouseHotkey(string hotkeyStr)
    {
        return hotkeyStr.Contains("鼠标中键") ||
               hotkeyStr.Contains("鼠标侧键1") ||
               hotkeyStr.Contains("鼠标侧键2");
    }
    private static MouseButton GetMouseButtonFromString(string hotkeyStr)
    {
        return hotkeyStr switch
        {
            var s when s.Contains("鼠标中键") => MouseButton.MiddleButton,
            var s when s.Contains("鼠标侧键1") => MouseButton.XButton1,
            var s when s.Contains("鼠标侧键2") => MouseButton.XButton2,
            _ => throw new ArgumentException($"不支持的鼠标热键: {hotkeyStr}")
        };
    }
    private static Keys GetKeysFromMouseButton(MouseButton button)
    {
        return button switch
        {
            MouseButton.MiddleButton => Keys.MButton,
            MouseButton.XButton1 => Keys.XButton1,
            MouseButton.XButton2 => Keys.XButton2,
            _ => Keys.None
        };
    }
    public static void UnregisterHotKey(string hotkeyStr = null)
    {
        if (string.IsNullOrEmpty(hotkeyStr))
        {
            Debug.WriteLine("[HOTKEY] 注销所有热键");

            foreach (var kvp in _registeredHotkeys.ToList())
            {
                try
                {
                    kvp.Value.hook.UnregisterHotKey();
                    kvp.Value.hook.Dispose();
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[HOTKEY] 键盘热键注销失败: {kvp.Key}, 错误: {ex.Message}");
                }
            }
            _registeredHotkeys.Clear();

            foreach (var kvp in _registeredMouseHotkeys.ToList())
            {
                try
                {
                    kvp.Value.hook.UnregisterMouseButton(kvp.Value.button);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[HOTKEY] 鼠标热键注销失败: {kvp.Key}, 错误: {ex.Message}");
                }
            }
            _registeredMouseHotkeys.Clear();

            if (_globalMouseHook != null)
            {
                _globalMouseHook.Dispose();
                _globalMouseHook = null;
            }
        }
        else
        {
            if (IsMouseHotkey(hotkeyStr) && _registeredMouseHotkeys.ContainsKey(hotkeyStr))
            {
                Debug.WriteLine($"[HOTKEY] 注销鼠标热键: {hotkeyStr}");
                try
                {
                    var (button, hook, callback) = _registeredMouseHotkeys[hotkeyStr];
                    hook.UnregisterMouseButton(button);
                    _registeredMouseHotkeys.Remove(hotkeyStr);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[HOTKEY] 鼠标热键注销失败: {hotkeyStr}, 错误: {ex.Message}");
                }
            }
            else if (_registeredHotkeys.ContainsKey(hotkeyStr))
            {
                Debug.WriteLine($"[HOTKEY] 注销键盘热键: {hotkeyStr}");
                try
                {
                    var (hotkey, hook, callback) = _registeredHotkeys[hotkeyStr];
                    hook.UnregisterHotKey();
                    hook.Dispose();
                    _registeredHotkeys.Remove(hotkeyStr);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[HOTKEY] 键盘热键注销失败: {hotkeyStr}, 错误: {ex.Message}");
                }
            }
        }
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
}
