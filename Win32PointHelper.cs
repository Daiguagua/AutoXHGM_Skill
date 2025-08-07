using System.Drawing;
using System.Runtime.InteropServices;

namespace AutoXHGM_Skill
{
    internal class Win32PointHelper
    {
        [DllImport("user32.dll")]
        public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        public static extern bool ClientToScreen(IntPtr hWnd, ref Point lpPoint);

        [DllImport("user32.dll")]
        public static extern bool ScreenToClient(IntPtr hWnd, ref Point lpPoint);

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        // 获取客户区左上角相对于窗口左上角的偏移量
        public static Point GetClientTopLeft(IntPtr hwnd)
        {
            GetWindowRect(hwnd, out RECT windowRect);
            Point clientPoint = new Point(0, 0); // 客户区(0,0)点
            ClientToScreen(hwnd, ref clientPoint);

            return new Point(
                clientPoint.X - windowRect.Left,
                clientPoint.Y - windowRect.Top
            );
        }

        // 将客户区坐标转换为屏幕坐标
        public static Point ClientToScreenPoint(IntPtr hwnd, Point clientPoint)
        {
            GetWindowRect(hwnd, out RECT windowRect);
            Point screenPoint = clientPoint;
            ClientToScreen(hwnd, ref screenPoint);
            return screenPoint;
        }
    }
}
