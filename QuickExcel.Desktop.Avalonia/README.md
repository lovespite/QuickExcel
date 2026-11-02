# QuickExcel Avalonia desktop

迁移中的新宿主，当前完成 A2 基础工作区。默认启动器仍指向 WPF，完整网格/工作区对齐、PDF.js/Monaco、平台发布依次在 A3–A6 实施。

从仓库根目录运行：

```powershell
dotnet run --project .\QuickExcel.Desktop.Avalonia\QuickExcel.Desktop.Avalonia.csproj -c Release --artifacts-path .\artifacts\build
.\scripts\test-avalonia-workspace.ps1
```

文件、历史与计算复用 Core/Infrastructure；编辑和会话复用 Presentation；网格仅为可见内容创建共享 Skia 像素，一格原生 TextBox 承接输入。新宿主不引用旧 Desktop 或 Grid.Wpf。合成输入、应用截图与自动 smoke 不代表真实 OS IME 或辅助技术验收。

Windows 组合根继续选择 ScriptWorkerRuntimeFactory/ScriptWorkerActionExecutor，随包复制完整 worker。其他平台在等效独立沙箱完成前拒绝脚本运行，没有进程内或同权限后备。

日常设置仍使用用户 QuickExcel 目录的既有格式；smoke 强制使用每次新建 artifacts 目录内的 settings，文件对话框和关闭决策仅在显式 smoke 入口注入。

版本与许可：Avalonia.Desktop/Themes.Fluent 12.1.4，精确 NuGet 锁文件随源码；SkiaSharp 4.154.0/HarfBuzz 14.2.1.400 沿用共享渲染层。Avalonia、ANGLE、MicroCom、Tmds.DBus 及构建服务声明保留在 Licenses；来源与不可变身份沿用 [A0 来源记录](../tools/QuickExcel.AvaloniaProbe/SOURCES.md)，渲染依赖见 [共享渲染来源](../QuickExcel.Rendering.Skia/SOURCES.md)。A2 未引入 WebView 运行依赖，A4 正式接入时再增加相应锁文件与许可。

阶段检查、失败修复与延期实机清单见 [实施日志](../QuickExcel.Desktop/docs/63-avalonia-implementation-log.md)。
