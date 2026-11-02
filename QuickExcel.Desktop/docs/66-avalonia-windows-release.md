# A5：Windows 发布、切换与回退

2026-11-03。A5 工程切换完成；A4 提交 `ea396e5`。用户已要求实机项目记入日志后统一验收，Windows 干净设备、IME、物理打印、跨屏 DPI 与辅助技术仍延期，不据此宣称 Beta 或关闭既有 M4/M6 门槛。

## 发布形态

`build-avalonia-release.ps1` 生成全新 Windows x64 自包含目录，版本 `0.6.0-preview.5`。主程序和 AppContainer worker 各自包含 .NET 10；worker 不依赖开发机全局 .NET，目录扫描拒绝 UI/WebView/Skia/PDF 载荷。正式构建 `QuickExcelDiagnostics=false`，不编译 Diagnostics 目录、对话框/验证/关闭决策注入，也拒绝测试命令行。运行时依赖清单与文件均排除 Avalonia.DesignerSupport；不使用裁剪、忽略依赖错误或删除必要 BCL 来减小包。

构建先生成正式离线 Web 资源，再还原并锁定 RID 依赖。`Directory.Build.props` 仅在显式传入 `QuickExcelReleaseLockDirectory` 时把 RID 锁放在 artifacts；正式包保存这些完整锁与 npm 锁，普通源码锁不被发布重写。逐文件大小/SHA-256、Git HEAD/dirty、SDK、Web 清单与安装自检进入 `release.json`；ZIP 另有 SHA-256。每次新目录，旧包不覆盖。无 Node、测试密钥、probe、fixture、PDB、source map 或 PDF scripting 引擎随包发布。

WebView 选择开源控件的 WebView2 adapter，要求注册的 **Evergreen stable Runtime**：检查微软官方 Clients `pv` 及对应 ClientState `EBWebView` 安装文件，显式传入该目录，避免自动选用 Edge 预览通道或任意环境覆盖。缺少 Runtime 时给出安装说明，网格/文件工作区仍可使用。在线/离线安装由用户运行微软安装器，不在启动时下载安装软件。

依据：[微软 Runtime 分发与注册检测](https://learn.microsoft.com/microsoft-edge/webview2/concepts/distribution)、[Evergreen 离线安装器](https://developer.microsoft.com/microsoft-edge/webview2/)。实际受测 Runtime **156.0.4314.58**，.NET **10.0.12**，Windows **10.0.22621**。没有执行卸载本机 Runtime；缺少 Runtime 的干净设备验收仍延期。

## 设置与回退

普通首次启动仅复制旧 `settings.json`、`recent-files.json` 到独立 `LocalAppData/QuickExcel/Avalonia`；先备份原始字节，flush 后原子发布，`migration-v1.json` 记录版本及哈希。现有 Avalonia 文件不覆盖，独占迁移锁避免并发；失败保留旧文件并显示状态。信任 pin、证书、恢复状态、WebView profile 均不迁移。单次用户授权和持久信任仍由原 DPAPI 合同管理。

`Start-QuickExcel.cmd` 调用 `start-quickexcel.ps1`：优先读取 artifacts 下经过核验的活动包；没有活动包才构建正式 Web 并从 Avalonia 源码启动。`build-quickexcel.ps1` 默认发布并在正式自检通过后更新活动指针。WPF 专用旧脚本暂留 `build-quickexcel-wpf.ps1`，后续 A7 整理。

旧受测基线由 A0 冻结清单逐项核验后另存到 `artifacts/avalonia/rollback/20261102-202406-c6898eb0/application`，同层有 ZIP/SHA-256。保留原 metadata 的旧 HEAD/dirty，未伪造为新的干净构建。回退只切换程序与原设置，工作簿不搬移、不覆盖。

```powershell
.\scripts\build-quickexcel.ps1 -RunSmoke -CreateArchive
.\scripts\select-quickexcel-package.ps1 -PackageDirectory <artifacts 内的已验包>
# 包内安装诊断，目标必须为新目录：
.\QuickExcel.Avalonia.exe --self-check <新诊断目录>
```

`--self-check` 是正式包安装诊断，不开放故障 RPC：验证资源清单、运行环境、文件往返、真正 Monaco/语言 worker/草稿确认、AppContainer 原生 worker、PDF.js/生成/导出与关闭。设置和输出只写显式诊断目录；缺少 worker 时无同权限或进程内后备。`--user-data <目录> [文件]` 提供隔离人工验收配置。

## 工程检查

| 检查 | 证据与结果 |
| --- | --- |
| 设置复制/备份/幂等/损坏/既有配置保护 | 2 项，`artifacts/avalonia/tests/a5-settings.trx` |
| 协议、资源、worker 隔离/故障/恢复 | 77 项，`artifacts/avalonia/releases/20261102-201956-4687611a/evidence/release-contracts.trx`；未跑全量 Core |
| 正式候选安装自检 | `20261102-202216-8ca3e947`、限制 stable Runtime 后 `20261102-202535-28f75ac0` 各 7 项通过 |
| 本机部署副本 | `artifacts/avalonia/deployment/20261102-202406-993cfaae`；中文/空格、只读安装目录通过；缺少 Web 资源/worker 明确失败 |
| Avalonia 入口/网格 | 77 项，`artifacts/avalonia/workspace/20261102-202842-235f3507` |
| Avalonia 工作区 | 32 项，`artifacts/avalonia/workspace/20261102-202850-1a2fac6c` |
| 归档 WPF 基线 | 157 项，`artifacts/validation/m5-grid/20261102-202448-b05cbecf` |
| 真实启动器切换旧包/新包 | WPF 157 项与 Avalonia 7 项，`artifacts/avalonia/rollback/launcher-check-02`、`artifacts/avalonia/releases/launcher-check-a5-02`；最终指回 Avalonia |
| 完整 Web 性能复核 | `artifacts/avalonia/web-benchmarks/20261102-202948-686ffeef`，编辑、PDF、回收、并发及各 5 次独立冷启动全部通过；原始样本见 `runs.log` 指向的目录 |
| 最终包部署复核 | `artifacts/avalonia/deployment/20261102-203754-55a66331`，四种副本场景再次通过；正式 EXE 拒绝四种 smoke/benchmark 命令，退出码 2 |
| 最终 Release 解决方案 | 19 项目，0 警告、0 错误；`artifacts/build` |

最终 A5 包为 `artifacts/avalonia/releases/20261102-203655-64921c89/application`，安装自检 7 项通过，含独立 RID 锁及正式资源；源码锁未被重写。ZIP SHA-256：`12B29484F68891CB3A739BA6D2C8922FA86D67B454EC514362E7DD9FAE79539F`。版本源码及发布包均只保存在本机，未推送远程。本地 WPF 源码标签 `avalonia-wpf-baseline-20261102` 指向 `37911cb04374b3dabd4413f8b7993c1c523c2eb0`。

性能复核中 Monaco 输入/补全 p95 为 13.3/3.8 ms，十次开关后增长 22,638,592 bytes；5 页 E4 最大进程树 private commit 为 791,056,384 bytes，低于原 768 MiB 限额但余量有限。100 页 PDF 十次开关并空闲 30 秒后较基线减少 14,016,512 bytes，浏览器进程全部退出；并发补全/PDF 缩放 p95 为 2.2/42.4 ms。没有上调预算或过滤失败样本，CPU 完成观测不代表物理显示延迟。

失败样本全部保留：首轮发布包含 NuGet native PDB，扫描也误把 System.Diagnostics BCL 识别为开发工具；改为构建时排除符号并只拒绝 Avalonia.Diagnostics。第二轮安装自检误用 API v1 不接受的带下划线函数名，已改为合法 `INSTALLCHECK`，没有改变脚本语义。首次 PowerShell 5 启动器用 `&` 运行 GUI EXE 时未等待、没有 LASTEXITCODE；改为明确等待进程并按 Windows 参数规则引用，重测通过。

只读 ACL 是本机独立副本，不能代替实际 Program Files、不同账号或干净机器。历史直接长路径 EXE 启动失败、正式安装器/程序签名及物理项目仍按原门禁跟踪。A6 的非 Windows 原生桥、沙箱、信任与打印单独交付，不能从本阶段外推。
