# QuickExcel Avalonia 0.6.0-preview.7

从仓库根目录运行 `Start-QuickExcel.cmd`，启动通过生产自检的活动 Windows 包；未选包时从源码构建。`Start-QuickExcel-Software.cmd` 可为一次启动强制软件绘制。发布使用 `scripts/build-quickexcel.ps1 -RunSmoke -CreateArchive`，生成自包含应用、独立隔离 worker、ZIP 和哈希清单。需要 WebView2 Evergreen 才能打开内嵌 Monaco 与 PDF.js；文件、网格和计算不依赖安装 Node。

界面使用菜单/Ribbon、公式栏、工作表标签和原生网格。脚本编辑器在保存、切换模块、执行和关闭前确认草稿版本；修改源码撤销旧授权。打印预览与导出共享冻结的 PDF 会话。Windows 实际打印使用系统对话框，Linux 使用 CUPS 查询纸张/硬边距并提交同一份 PDF。

Windows 脚本使用原 AppContainer 隔离；Linux 的受测 Ubuntu 候选使用 Landlock/seccomp/cgroup。没有平台隔离或系统保护服务时相应执行/持久信任能力明确失败，不降级。macOS worker 尚未实现，执行保持关闭。平台依赖、包与验证边界见 [67 号日志](docs/67-avalonia-platforms.md)。

物理输入法、跨屏 DPI、辅助技术、物理打印和干净设备验收按用户要求延期。Linux VM 的中文输入没有候选框，仍待统一验证。历史 M4/M5/M6 门槛没有因 UI 迁移而自动关闭。

当前目录保留产品文档、许可证与历史记录，已无 WPF 应用源码或构建入口。WPF 回退使用已归档的受测包，源码标签为 `avalonia-wpf-baseline-20261102`。详见 [Windows 发布](docs/66-avalonia-windows-release.md)、[A7 收尾](docs/68-avalonia-retirement.md)、[旧使用说明归档](docs/69-wpf-usage-archive.md)。
