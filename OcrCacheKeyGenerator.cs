using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Skill_Loop
{
    // OcrCacheKeyGenerator.cs
    public static class OcrCacheKeyGenerator
    {
        public static string GenerateKey(IntPtr windowHandle, SkillCondition condition)
        {
            // 获取窗口位置和客户区偏移
            Win32PointHelper.GetWindowRect(windowHandle, out var windowRect);
            var clientOffset = Win32PointHelper.GetClientTopLeft(windowHandle);

            // 计算绝对坐标
            int absoluteX = windowRect.Left + clientOffset.X + condition.OcrRegionX;
            int absoluteY = windowRect.Top + clientOffset.Y + condition.OcrRegionY;

            // 使用区域坐标和大小生成键
            return $"{windowHandle}_{absoluteX}_{absoluteY}_{condition.OcrRegionWidth}_{condition.OcrRegionHeight}";
        }

        public static string GenerateKey(IntPtr windowHandle, System.Drawing.Rectangle region)
        {
            // 使用区域坐标和大小生成键
            return $"{windowHandle}_{region.X}_{region.Y}_{region.Width}_{region.Height}";
        }
    }
}
