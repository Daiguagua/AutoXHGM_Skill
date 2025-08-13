using AutoXHGM_Skill.Fischless.HotkeyCapture;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Vanara.PInvoke;

namespace Fischless.HotkeyCapture;

public sealed class MouseHook : IDisposable
{
    public event EventHandler<MouseKeyPressedEventArgs>? MouseKeyPressed;

    private const int WH_MOUSE_LL = 14;
    private const int HC_ACTION = 0;
    private const int WM_MBUTTONDOWN = 0x0207;
    private const int WM_XBUTTONDOWN = 0x020B;

    private User32.HookProc _proc;
    private IntPtr _hookID = IntPtr.Zero;
    private readonly HashSet<MouseButton> _registeredButtons = new();

    public MouseHook()
    {
        _proc = HookCallback;
    }

    public void RegisterMouseButton(MouseButton button)
    {
        Debug.WriteLine($"[MOUSE_HOOK] 注册鼠标按键: {button}");

        if (!_registeredButtons.Contains(button))
        {
            _registeredButtons.Add(button);
        }

        if (_hookID == IntPtr.Zero)
        {
            _hookID = SetHook(_proc);
            Debug.WriteLine($"[MOUSE_HOOK] 钩子已安装: {_hookID}");
        }
    }

    public void UnregisterMouseButton(MouseButton button)
    {
        Debug.WriteLine($"[MOUSE_HOOK] 注销鼠标按键: {button}");

        _registeredButtons.Remove(button);

        if (_registeredButtons.Count == 0 && _hookID != IntPtr.Zero)
        {
            UnhookWindowsHookEx(_hookID);
            _hookID = IntPtr.Zero;
            Debug.WriteLine("[MOUSE_HOOK] 钩子已卸载");
        }
    }

    private IntPtr SetHook(User32.HookProc proc)
    {
        using var curProcess = Process.GetCurrentProcess();
        using var curModule = curProcess.MainModule;

        return SetWindowsHookEx(
            WH_MOUSE_LL,
            proc,
            GetModuleHandle(curModule?.ModuleName),
            0);
    }

    private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (nCode >= HC_ACTION)
        {
            var mouseButton = GetMouseButtonFromMessage((int)wParam, lParam);

            if (mouseButton.HasValue && _registeredButtons.Contains(mouseButton.Value))
            {
                Debug.WriteLine($"[MOUSE_HOOK] 检测到注册的鼠标按键: {mouseButton}");
                MouseKeyPressed?.Invoke(this, new MouseKeyPressedEventArgs(mouseButton.Value));
            }
        }

        return CallNextHookEx(_hookID, nCode, wParam, lParam);
    }

    private MouseButton? GetMouseButtonFromMessage(int wParam, IntPtr lParam)
    {
        return wParam switch
        {
            WM_MBUTTONDOWN => MouseButton.MiddleButton,
            WM_XBUTTONDOWN => GetXButtonType(lParam),
            _ => null
        };
    }

    private MouseButton? GetXButtonType(IntPtr lParam)
    {
        var hookStruct = Marshal.PtrToStructure<MSLLHOOKSTRUCT>(lParam);
        var xButton = (hookStruct.mouseData >> 16) & 0xFFFF;

        return xButton switch
        {
            0x0001 => MouseButton.XButton1,
            0x0002 => MouseButton.XButton2,
            _ => null
        };
    }

    public void Dispose()
    {
        foreach (var button in _registeredButtons.ToList())
        {
            UnregisterMouseButton(button);
        }
    }

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern IntPtr SetWindowsHookEx(int idHook, User32.HookProc lpfn, IntPtr hMod, uint dwThreadId);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool UnhookWindowsHookEx(IntPtr hhk);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern IntPtr GetModuleHandle(string? lpModuleName);

    [StructLayout(LayoutKind.Sequential)]
    private struct MSLLHOOKSTRUCT
    {
        public int x;
        public int y;
        public uint mouseData;
        public uint flags;
        public uint time;
        public IntPtr dwExtraInfo;
    }
}

public class MouseKeyPressedEventArgs : EventArgs
{
    public MouseButton Button { get; }

    public MouseKeyPressedEventArgs(MouseButton button)
    {
        Button = button;
    }
}
