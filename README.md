星痕共鸣技能循环助手
上班摸鱼产物，代码也大部分都是deepseek写的，因为太刀循环太麻烦了，所以写个小程序，开源给各位太刀虾一起用hh（以下README 由deepseek亲自完成x 因为本人也不太会用github hh 

项目概述
这是一个为游戏《星痕共鸣》设计的技能循环自动化工具。程序通过检测游戏窗口中特定位置的颜色条件，自动触发预设的技能按键序列，帮助玩家实现高效的技能循环。

主要功能
智能技能触发：根据屏幕特定位置的颜色条件自动释放技能

可视化调试：实时显示检测点位置，方便配置和调试

专业取色器：直观地选择和配置颜色条件

热键控制：支持自定义热键启动/停止技能循环

系统托盘运行：后台运行时最小化到系统托盘

配置管理：支持保存和加载技能配置方案

技术栈
.NET 8.0

WPF (Windows Presentation Foundation)

Fischless.WindowsInput (输入模拟库)

Hardcodet.Wpf.TaskbarNotification (系统托盘库)

特别感谢 Fischless.WindowsInput 库提供的高质量输入模拟功能。
https://github.com/GenshinMatrix/Fischless

使用说明
安装要求
.NET 8.0 桌面运行时

基本使用步骤
启动程序

双击运行 AutoXHGM_Skill.exe

选择游戏窗口

在顶部下拉框中选择游戏窗口

点击"刷新窗口"更新窗口列表

配置技能规则

在"技能规则"区域点击"添加规则"

设置技能按键和检测频率

点击"编辑条件"配置颜色检测点：

设置偏移坐标 (X,Y)

设置目标颜色和容差值

双击条件行使用专业取色器

设置全局选项

调整全局检测频率（毫秒）

设置启动/停止热键

保存配置

点击"保存配置"按钮保存当前设置

开始技能循环

点击"开始循环"按钮或使用预设热键启动

程序将自动最小化到系统托盘

停止技能循环

点击系统托盘图标恢复主界面

点击"停止循环"按钮或使用停止热键

专业取色器使用
在编辑条件时双击条件行打开专业取色器：

移动鼠标到游戏窗口内，会显示放大镜

左键点击选择位置和颜色

右键或ESC取消

构建说明
克隆仓库：

bash
git clone https://github.com/Daiguagua/AutoXHGM_Skill.git
使用Visual Studio 2022或更高版本打开解决方案

安装所需NuGet包：

Fischless.WindowsInput

Hardcodet.Wpf.TaskbarNotification

构建项目 (Ctrl+Shift+B)

注意事项
程序需要以管理员权限运行（如果游戏需要管理员权限）

确保游戏窗口不被其他窗口遮挡

首次使用时建议在非战斗场景测试规则

热键冲突时可在设置中更换热键

贡献指南
欢迎提交Issue和Pull Request！贡献前请阅读：

Fork项目并创建特性分支

遵循现有代码风格

提交清晰的commit信息

创建详细的Pull Request描述

许可证
本项目采用 GNU Affero General Public License v3.0 许可证 - 详情请见 LICENSE 文件。
