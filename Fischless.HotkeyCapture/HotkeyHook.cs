using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Vanara.PInvoke;

namespace Fischless.HotkeyCapture;

public sealed class HotkeyHook : IDisposable
{
    public event EventHandler<KeyPressedEventArgs>? KeyPressed = null;

    private readonly Window window = new();
    private int currentId;

    private class Window : NativeWindow, IDisposable
    {
        public event EventHandler<KeyPressedEventArgs>? KeyPressed = null;

        public Window()
        {
            CreateHandle(new CreateParams());
        }

        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);

            if (m.Msg == (int)User32.WindowMessage.WM_HOTKEY)
            {
                Keys key = (Keys)(((int)m.LParam >> 16) & 0xFFFF);
                User32.HotKeyModifiers modifier = (User32.HotKeyModifiers)((int)m.LParam & 0xFFFF);

                KeyPressed?.Invoke(this, new KeyPressedEventArgs(modifier, key));
            }
        }

        public void Dispose()
        {
            DestroyHandle();
        }
    }

    public HotkeyHook()
    {
        window.KeyPressed += (sender, args) =>
        {
            KeyPressed?.Invoke(this, args);
        };
    }

    public void RegisterHotKey(User32.HotKeyModifiers modifier, Keys key)
    {
        currentId += 1;
        Debug.WriteLine($"[DEBUG] 尝试注册热键 ID:{currentId}, Modifier:{modifier}, Key:{key}");

        if (!User32.RegisterHotKey(window!.Handle, currentId, modifier, (uint)key))
        {
            int error = Marshal.GetLastWin32Error();
            Debug.WriteLine($"[ERROR] 热键注册失败，错误码: {error}");

            if (error == SystemErrorCodes.ERROR_HOTKEY_ALREADY_REGISTERED)
            {
                Debug.WriteLine("[ERROR] 热键已被其他程序注册");
                throw new InvalidOperationException("Hotkey already registered");
            }
            else
            {
                Debug.WriteLine($"[ERROR] 热键注册失败，原因: {GetErrorMessage(error)}");
                throw new InvalidOperationException($"Hotkey registration failed: {error}");
            }
        }
        else
        {
            Debug.WriteLine($"[SUCCESS] 热键注册成功 ID:{currentId}");
        }
    }

    private string GetErrorMessage(int errorCode)
    {
        return errorCode switch
        {
            0x581 => "热键已被注册",
            0x58B => "热键未注册",
            5 => "拒绝访问",
            87 => "参数错误",
            _ => $"未知错误 ({errorCode})"
        };
    }


    public void UnregisterHotKey()
    {
        for (int i = currentId; i > 0; i--)
        {
            User32.UnregisterHotKey(window!.Handle, i);
        }
    }

    public void Dispose()
    {
        UnregisterHotKey();
        window?.Dispose();
    }
}
