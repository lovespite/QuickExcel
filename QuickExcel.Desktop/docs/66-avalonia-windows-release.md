# A5：Windows 发布、切换与回退

## Windows ClickOnce 发布（2026-11-20）

当前产品版本为 `3.13.0`。Windows 主可执行文件改为 `quickxl.exe`；
`QuickExcel.Avalonia.dll`、deps/runtimeconfig、产品显示名、隔离 worker 和
macOS/Linux 可执行文件名称均保持不变。构建、发布、`dotnet run`、文件关联、
ClickOnce launcher 及诊断脚本使用新入口；本地启动器仍支持历史已验证包的旧入口。
设置页 API 浏览器在首次选择成员、切换类型或筛选后会恢复显示调用类型标签；
设置冒烟检查覆盖宿主调用、本地成员和枚举值。

`scripts/new-code-signing-certificate.ps1` 创建 RSA 3072 / SHA-256 的代码签名
自签证书，有效期精确为创建时间起 10 年。私钥保留在当前用户 `My` 证书存储，
并将加密 PFX、公钥 CER、证书元数据和 DPAPI 保护的随机 PFX 密码写入新的
`artifacts/signing/` 目录。该目录仅当前用户可访问；不自动信任证书，不把私钥
或密码放入发布包。迁移机器前按目录内 README 导出使用自选密码的便携备份。

```powershell
pwsh -File scripts/new-code-signing-certificate.ps1
# 使用上一步返回的 thumbprint；后续版本继续使用同一证书：
pwsh -File scripts/build-clickonce.ps1 -CertificateThumbprint <thumbprint>
```

ClickOnce 脚本默认调用 `build-quickexcel.ps1 -SkipTests -RunSmoke`，在正式包
自检及逐文件身份记录前签名 QuickExcel 的 EXE/DLL，再更新本地启动指针。
`-UseActivePackage` 仅可复用同一内部版本、通过正式自检、文件哈希一致且已由
指定证书签名的活动包。所有输出都位于本仓库 artifacts；已有包不覆盖。

产品版本来自 Avalonia csproj 的 `Version`，不独立递增 ClickOnce revision。
例如内部 `3.13.0` 对应 ClickOnce / EXE 文件版本 `3.13.0.0`。
Microsoft.DotNet.Mage 固定为 10.0.0；现代 .NET 使用
其 Launcher.exe 启动 `quickxl.exe`，应用和部署清单使用 SHA-256 签名。安装包不带 provider URL
或 subscription 自动更新配置，可从后续完整包手动更新。

发布目录包含 `setup.exe`、`QuickExcel.application`、完整 Application Files、
微软签名的 WebView2 Evergreen **在线引导程序**、公钥证书和中文安装说明。
setup.exe 检测 HKLM/HKCU 的 WebView2 stable Runtime，缺少时通过
`MicrosoftEdgeWebview2Setup.exe /silent /install` 联网下载并安装对应架构的运行库；
已安装时跳过。此步骤需要连接 Internet，包内不再包含完整离线运行库。
微软分发说明：<https://learn.microsoft.com/en-us/microsoft-edge/webview2/concepts/distribution>；
引导程序来源：<https://go.microsoft.com/fwlink/p/?LinkId=2124703>。
发布与验证脚本检查 Microsoft 签名及引导程序体积，并拒绝混入旧的完整离线安装器。
主程序和隔离 worker 各自包含 .NET 10，worker 原生 VC++ 依赖随包携带。
受支持的 Windows 10/11 x64 提供 ClickOnce 所需的 .NET Framework；无需安装
.NET SDK。微软安装器原样保留，来源、SHA-256、安装器文件版本记录在
`clickonce-release.json`；相关声明及 launcher MIT 全文可在设置的许可证列表查看。

`test-clickonce.ps1` 独立验证两个 XML 签名、SHA-256 文件清单的完整覆盖、发布
和 EXE 版本、现代 .NET launcher 入口、程序签名、必要组件，并对 ClickOnce
实际载荷重新运行生产自检。WinVerifyTrust 只允许成功或自签证书的
CERT_E_UNTRUSTEDROOT；损坏文件的 TRUST_E_BAD_DIGEST 仍拒绝。不会通过修改
信任存储让校验通过。发布另生成完整 ZIP 和 SHA-256 文件。

自签证书不会自动获得 Windows/SmartScreen 信任。目标设备如要求受信任发布者，
需用户或管理员核对指纹后显式配置证书信任。自动化载荷自检不代表干净设备的
ClickOnce 安装、缺少 WebView2 时的安装分支或物理输入验收；既有延期门槛保留。

本次 `3.13.0` 验证：API 目录/能力清单/翻译共 10 项定向单元测试通过；
API 生成器 fixture 自测和正式生成物 `--check` 通过；中英文 Menus 模式分别
通过 136 项 UI 检查（含调用类型的首次显示、切换及文件关联入口），并检查截图。
中文证据：`artifacts/avalonia/workspace/20261120-020742-728ff60b`；
英文证据：`artifacts/avalonia/workspace/menus-en-20261120-020850`（通过 `dotnet run` 启动）。
Windows 发布已运行 `build-quickexcel.ps1 -SkipTests -RunSmoke` 并签名，
`active-windows.json` 指向 `artifacts/avalonia/releases/20261120-020817-5493b6da/application`；
正式包 32 项生产自检通过，`Start-QuickExcel.cmd` 启动同一包的 32 项自检也通过，
启动器证据：`artifacts/avalonia/launcher-self-check-20261120-020929`。
ClickOnce `3.13.0.0` 的 756 项文件哈希、两个清单签名、代码签名和实际载荷的
32 项生产自检通过，证据：`artifacts/clickonce/work/20261120-020930-437bb740/verification`。
新分发包：`artifacts/clickonce/QuickExcel-3.13.0-win-x64-20261120-020930-437bb740.zip`，
232,914,231 字节；ZIP 的 763 个条目逐项核对大小和 SHA-256 通过。
ZIP SHA-256：`4ED9FC1BD8E98A9455FD546D2AB15B61D45ECBA2C65797A19E1B5C556FC8EC6A`。
WebView2 分发组件由 294,099,168 字节的完整离线安装器改为 1,897,688 字节的
在线引导程序，Microsoft 签名校验通过，旧离线安装器被新的引导程序检查拒绝。
未运行全量回归；未执行干净设备 ClickOnce 安装或卸载本机 WebView2。

历史离线包交付记录：内部版本 `3.12.0` / ClickOnce `3.12.0.0`，证书有效期
2026-11-20 至 2036-11-20，指纹 `4D4195DD8C41C2F6A10AFFF8DE0F0A9674FC3C01`。
当时活动包为 `artifacts/avalonia/releases/20261120-015011-7ffdc1eb/application`；
ClickOnce ZIP 为 `artifacts/clickonce/QuickExcel-3.12.0-win-x64-20261120-015312-b0b1b23a.zip`。
756 项应用文件哈希、两个清单签名、EXE/DLL 签名、32 项生产自检通过；
ZIP 的 763 个条目逐项复核通过，PFX 可用 DPAPI 密码解密并匹配签名证书。
校验证据在 `artifacts/clickonce/work/20261120-015312-b0b1b23a/verification`；
修改 XML 或已签名 PE 的负向检查均拒绝，未运行全量回归。
ZIP SHA-256：`6E2B86D56271D42FB6375B691EC7D134EEFA9FF7BEFFE55A3F8075B22984E8CB`。

2026-11-03。A5 工程切换完成；A4 提交 `ea396e5`。用户已要求实机项目记入日志后统一验收，Windows 干净设备、IME、物理打印、跨屏 DPI 与辅助技术仍延期，不据此宣称 Beta 或关闭既有 M4/M6 门槛。

## 发布形态

`build-avalonia-release.ps1` 生成全新 Windows x64 自包含目录。A5 候选版本为 `0.6.0-preview.5`；当前版本由 Avalonia 项目统一提供，为 `0.6.0-preview.7`。主程序和 AppContainer worker 各自包含 .NET 10；worker 不依赖开发机全局 .NET，目录扫描拒绝 UI/WebView/Skia/PDF 载荷。正式构建 `QuickExcelDiagnostics=false`，不编译 Diagnostics 目录、对话框/验证/关闭决策注入，也拒绝测试命令行。运行时依赖清单与文件均排除 Avalonia.DesignerSupport；不使用裁剪、忽略依赖错误或删除必要 BCL 来减小包。

构建先生成正式离线 Web 资源，再还原并锁定 RID 依赖。`Directory.Build.props` 仅在显式传入 `QuickExcelReleaseLockDirectory` 时把 RID 锁放在 artifacts；正式包保存这些完整锁与 npm 锁，普通源码锁不被发布重写。逐文件大小/SHA-256、Git HEAD/dirty、SDK、Web 清单与安装自检进入 `release.json`；ZIP 另有 SHA-256。每次新目录，旧包不覆盖。无 Node、测试密钥、probe、fixture、PDB、source map 或 PDF scripting 引擎随包发布。

WebView 选择开源控件的 WebView2 adapter，要求注册的 **Evergreen stable Runtime**：检查微软官方 Clients `pv` 及对应 ClientState `EBWebView` 安装文件，显式传入该目录，避免自动选用 Edge 预览通道或任意环境覆盖。缺少 Runtime 时给出安装说明，网格/文件工作区仍可使用。在线/离线安装由用户运行微软安装器，不在启动时下载安装软件。

依据：[微软 Runtime 分发与注册检测](https://learn.microsoft.com/microsoft-edge/webview2/concepts/distribution)、[Evergreen 离线安装器](https://developer.microsoft.com/microsoft-edge/webview2/)。实际受测 Runtime **156.0.4314.58**，.NET **10.0.12**，Windows **10.0.22621**。没有执行卸载本机 Runtime；缺少 Runtime 的干净设备验收仍延期。

## 设置与回退

普通首次启动仅复制旧 `settings.json`、`recent-files.json` 到独立 `LocalAppData/QuickExcel/Avalonia`；先备份原始字节，flush 后原子发布，`migration-v1.json` 记录版本及哈希。现有 Avalonia 文件不覆盖，独占迁移锁避免并发；失败保留旧文件并显示状态。信任 pin、证书、恢复状态、WebView profile 均不迁移。单次用户授权和持久信任仍由原 DPAPI 合同管理。

`Start-QuickExcel.cmd` 调用 `start-quickexcel.ps1`：优先读取 artifacts 下经过核验的活动包；没有活动包才构建正式 Web 并从 Avalonia 源码启动。`build-quickexcel.ps1` 默认发布并在正式自检通过后更新活动指针。A7 已移除 WPF 专用旧脚本 `build-quickexcel-wpf.ps1`；回退仅选择已归档且通过哈希核验的 WPF 包，见 [68 号记录](68-avalonia-retirement.md)。

旧受测基线由 A0 冻结清单逐项核验后另存到 `artifacts/avalonia/rollback/20261102-202406-c6898eb0/application`，同层有 ZIP/SHA-256。保留原 metadata 的旧 HEAD/dirty，未伪造为新的干净构建。回退只切换程序与原设置，工作簿不搬移、不覆盖。

```powershell
.\scripts\build-quickexcel.ps1 -RunSmoke -CreateArchive
.\scripts\select-quickexcel-package.ps1 -PackageDirectory <artifacts 内的已验包>
# 包内安装诊断，目标必须为新目录：
.\quickxl.exe --self-check <新诊断目录>
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

最终 A5 包为 `artifacts/avalonia/releases/20261102-203655-64921c89/application`，安装自检 7 项通过，含独立 RID 锁及正式资源；源码锁未被重写。ZIP SHA-256：`12B29484F68891CB3A739BA6D2C8922FA86D67B454EC514362E7DD9FAE79539F`。这是 A5 历史候选，生成包仅保存在本机；阶段源码现已在迁移分支推送。当前活动包由 `artifacts/avalonia/active-windows.json` 指向，不能用 A5 旧包代表最新源码。本地 WPF 源码标签 `avalonia-wpf-baseline-20261102` 指向 `37911cb04374b3dabd4413f8b7993c1c523c2eb0`。

性能复核中 Monaco 输入/补全 p95 为 13.3/3.8 ms，十次开关后增长 22,638,592 bytes；5 页 E4 最大进程树 private commit 为 791,056,384 bytes，低于原 768 MiB 限额但余量有限。100 页 PDF 十次开关并空闲 30 秒后较基线减少 14,016,512 bytes，浏览器进程全部退出；并发补全/PDF 缩放 p95 为 2.2/42.4 ms。没有上调预算或过滤失败样本，CPU 完成观测不代表物理显示延迟。

失败样本全部保留：首轮发布包含 NuGet native PDB，扫描也误把 System.Diagnostics BCL 识别为开发工具；改为构建时排除符号并只拒绝 Avalonia.Diagnostics。第二轮安装自检误用 API v1 不接受的带下划线函数名，已改为合法 `INSTALLCHECK`，没有改变脚本语义。首次 PowerShell 5 启动器用 `&` 运行 GUI EXE 时未等待、没有 LASTEXITCODE；改为明确等待进程并按 Windows 参数规则引用，重测通过。

只读 ACL 是本机独立副本，不能代替实际 Program Files、不同账号或干净机器。历史直接长路径 EXE 启动失败、正式安装器/程序签名及物理项目仍按原门禁跟踪。A6 的非 Windows 原生桥、沙箱、信任与打印单独交付，不能从本阶段外推。
