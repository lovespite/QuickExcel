# A7 旧 WPF 宿主收尾与验收索引

2026-11-03。默认 Windows 交付范围完成工程切换和收尾；实机验收仍延期。A6 Linux 有受测 Ubuntu 候选，macOS worker 未实现且执行保持关闭，不将其计为完整支持。阶段提交互相独立，所有提交仅在本机，等待用户验收后才可推送。

## 收尾范围

- 从当前源码与解决方案移除 `QuickExcel.Desktop.csproj`、`QuickExcel.Grid.Wpf.csproj`、WPF 窗口/控件/诊断/绘制源码，以及 `build-quickexcel-wpf.ps1`。17 项目 Release 构建不需要 WPF 或 FluentIcons.Wpf；产品项目没有 WindowsDesktop/WPF 引用。
- 独立 ZIP/XML/v1 包夹具移至 `QuickExcel.Core.Test/Fixtures/WorkbookPackageFixture.cs`，测试直接编译自身文件，不再链接旧桌面源码。夹具内容与 v1 回归保留。
- `QuickExcel.Desktop/docs`、发布历史、Fluent Icons/旧网格/WriteableBitmapEx 许可证与原始来源清单保留；旧使用说明单独归档为 [69 号记录](69-wpf-usage-archive.md)。共享网格仍保留原有 MIT 来源。
- `Start-QuickExcel.cmd` 与默认发布脚本只选择 Avalonia，或显式选择经过哈希核验的 WPF 归档包。`Start-QuickExcel-Software.cmd` 为本次 Windows 启动强制 Avalonia 软件渲染；普通入口不继承这个设置。WPF 归档的相同诊断选项仍受支持。
- 保留 `test-m5-grid.ps1` 和 `test-script-signatures.ps1` 用于历史回退包，改为要求显式路径并核验归档类型；不再默认指向失效的 WPF 源码构建输出。Excel COM 对照脚本、历史部署工具和其证据保留。
- Windows/Unix 发布版本从唯一的 Avalonia 项目版本读取，当前 `0.6.0-preview.7`。发布时拒绝旧 UI、设计器、测试 probe、测试密钥、源码映射及符号；worker 拒绝 Web/UI/Skia 载荷。

## 工程验证

| 检查 | 证据 |
| --- | --- |
| 17 项目 Release 构建 | `artifacts/build`，0 警告、0 错误 |
| 迁移后的存储/v1/文档夹具 | **112/112**，`artifacts/avalonia/tests/a7-fixture-migration.trx` |
| A6 后 Windows worker/签名/隔离/恢复/故障/页范围 | **104/104**，`artifacts/avalonia/tests/a6-final-platform.trx` |
| Windows Monaco/PDF.js/签名/actions | **42**，`artifacts/avalonia/web-tests/20261102-215000-dd70e316` |
| Windows 网格与工作区 | **77 + 32**，`artifacts/avalonia/workspace/20261102-215016-3e684a01`、`20261102-215020-da328b2f` |
| 删除 WPF 源码后运行归档 | **157**，`artifacts/validation/m5-grid/20261102-215204-c32b25ea` |
| 实际 CMD → PowerShell 5 → 归档 EXE | **157**，`artifacts/avalonia/rollback/launcher-check-a7-01`，随后指回 Avalonia |
| Windows 正式发布自检 | **7**，`artifacts/avalonia/releases/20261102-215213-6e40b24e/evidence/self-check` |
| 正式包部署副本 | **4** 场景，`artifacts/avalonia/deployment/20261102-215330-1a6e619b`；中文/空格、只读安装通过，Web/worker 缺失明确失败 |
| 实际普通/软件渲染启动器 | 各 **7**，`artifacts/avalonia/releases/launcher-check-a7-02`、`software-check-a7-02`，请求状态分别 false/true |

首次实际 CMD 检查发现 PowerShell 7 向系统 PowerShell 5 传递了不兼容的 PSModulePath，导致 Get-FileHash 不可用，自检尚未启动。修复为在 PowerShell 5 中优先加载自身的 Utility/Management 模块；没有跳过哈希检查。实际两个 CMD 入口与归档切换均重测通过。

没有重复执行全量 Core 或整套性能基准。A7 不改变默认绘制路径、Web 模型或 PDF 缓存；A5 的固定负载、预算及完整原始样本继续见 [66 号日志](66-avalonia-windows-release.md)。软件渲染检查仅验证显式诊断入口，不作为默认性能数字。A6 Linux 的最终安全/信任/虚拟打印/包验证见 [67 号日志](67-avalonia-platforms.md)。

收尾候选 `artifacts/avalonia/releases/20261102-215213-6e40b24e/application`，ZIP SHA-256：`3959FB3504F2E639607985EC38AC4489E651217930A0C884F6A055E55F00966C`。阶段提交后将从干净 HEAD 再生成正式验收包；最终活动目录由 `artifacts/avalonia/active-windows.json` 指向，包内 release.json 记录源提交、dirty 状态、版本、逐文件哈希和自检，旁边 evidence 保存结果。该指针与生成包均在 Git 之外，不用文档中的旧目录冒充最新源码。

## 回退与待验收

WPF 源码标签：`avalonia-wpf-baseline-20261102`，提交 `37911cb04374b3dabd4413f8b7993c1c523c2eb0`。归档：`artifacts/avalonia/rollback/20261102-202406-c6898eb0/application`；ZIP SHA-256：`05A1BA89D1958428385F2D1331541FBD8CE7B73F488031B280D585300F442EA3`。回退仅切换完整包，不覆盖工作簿、信任、设置或浏览器 profile，也不依赖当前树内保留两套 UI。

用户要求延期的项目：真实中文 IME、WebView 物理焦点/弹窗、跨屏 DPI、辅助技术、物理打印及目标干净设备。Linux VM 输入中文时没有候选窗，仍是未解决项。macOS native/WKWebView/Keychain/打印/签名公证缺少设备证据，worker 还需独立实现与负向验证，因此 macOS 完整支持未完成。

原 M4/M5/M6、直接长路径启动及其他历史门槛保留原状态。当前结果为 Windows 工程迁移和 WPF 收尾完成、Linux 工程候选完成；不能表述为所有平台和实机门槛全部完成。
