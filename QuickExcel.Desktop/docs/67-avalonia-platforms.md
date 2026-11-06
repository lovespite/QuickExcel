# A6 平台实施、隔离 ADR 与交付边界

> 2026-11-03 引擎迁移更新：当前生产路径为固定 JSC、IPC 6，隔离 TS 编译与同步模块/中断契约见 [71 号实施记录](71-jsc-typescript-implementation-log.md)。执行仍为 256 MiB 进程 / 256 KiB 栈 / 1 秒操作 / 5 秒整轮；64 MiB 引擎硬限额 `JSC-MEM-001` 未实现，用户明确延期且不作为推进门禁。下文 QuickJS、旧协议、预算及构建入口均为当时版本的历史证据，不代表当前实现。

2026-11-03。A5 独立提交 `4b2d313`。Linux 工程实现与受测 Ubuntu 包完成，实机验收延期；macOS arm64 已完成脚本 worker 的工程实现与本机自检（Seatbelt、看门狗、登记恢复），签名/公证、持久信任与实机门禁仍独立记录，不把 macOS 完整支持标为完成。Windows、Linux、macOS 分别记录，不以 UI 能启动或设备验收延期替代工程安全证据。物理输入、打印、DPI、辅助技术和干净设备按用户要求统一延期。Linux 中文输入无候选框仍未解决。

## Linux worker 决策

Ubuntu 24.04.5 x64 / Linux 6.8.0-147 / systemd 用户会话 / cgroup v2 / Landlock ABI 4 是 A6 受测环境。原生桥仍为 QuickJS-NG 0.17.0、原 SHA-256，C ABI、1 秒中断、64 MiB 引擎预算与 16 MiB 桥接均保持。A6 原始证据使用 IPC 4；当前事件/虚拟文件功能已升级到 IPC 5，契约及新增 Windows 验证见 [69 号记录](69-script-events-and-virtual-files.md)，不能将历史 Linux 包当作最新源码的验证。Unix 使用单调时钟及 PIC；Linux 编译使用 Clang 18。GCC 13 在未修改的 QuickJS 优化阶段发生内部编译器错误，失败日志保留，未修改引擎规避。macOS arm64 原生构建入口已准备，尚无实际构建证据。

Bubblewrap 在当前普通用户环境受到 Ubuntu AppArmor 的 user namespace / network capability 策略限制，不能启动隔离网络；没有禁用该策略。最终 Linux 路径采用独立的 Landlock + seccomp + cgroup 组合，不依赖 Bubblewrap，也不提供同权限执行后备。

1. 宿主先建立私有 0700 目录、独占租约和持久所有权记录，复制完整自包含 worker，只允许平面常规文件，无链接路径。载荷文件只读，独立记录 SHA-256。
2. 原生 launcher 是宿主直接子进程，立即设置 parent-death SIGKILL 并核对父 PID；关闭所有额外继承描述符，只有 stdin/stdout/stderr。此时只等待启动字节，不加载 CLR 或工作簿源码。
3. 宿主调用 systemd 用户管理器创建唯一 scope，内核核验 PID 所属组、`memory.max=268435456`、`memory.swap.max=0`、`pids.max=32` 后才放行。scope 自带 5 秒最长生命周期和控制组终止；无用户总线/控制器或核验失败则拒绝。
4. launcher 应用 no_new_privs、Landlock 文件白名单和默认拒绝的 seccomp，再 exec worker。仅私有载荷、精确 ELF loader、系统库及必要的只读 CPU/进程信息可读；没有 home/tmp/工作簿目录、可写文件、网络 socket、外部 IPC、fork/子进程、namespace、ptrace 或进程内存访问权限。
5. worker 在第一条协议消息前再次关闭 dumpability，检查实际读文件/建 socket 被拒绝，再以 TSYNC 对所有 CLR 线程安装禁止 exec 的最终过滤器。宿主在发送源码前核对 NoNewPrivs、Seccomp 与 CapEff/CapPrm。clone 仅接受同进程线程标志；clone3 返回 ENOSYS，让 libc 使用可检查的 clone 参数。自身线程 affinity 是 CLR 启动所需，限定 pid=0。
6. 每轮结束等待实际进程退出再清理。恢复必须取得独占租约、验证完整记录和 PID 启动时间，不杀活跃进程、不按名称前缀清扫、不跟随链接。未登记目录保留。已终止 scope 由 systemd 收集。

Landlock 当前并不隐藏所有路径元数据（例如部分 stat 查询）；本策略承诺禁止未授权文件内容读取、写入和目录枚举，不声称完全隐藏宿主存在信息。系统只读库属于受信基础；已控制宿主登录用户的恶意程序不属于 workbook worker 沙箱能够隔离的对象。所有网络创建 syscall 均由 seccomp 拒绝，包含 IPv4/IPv6、UDP/DNS 与 Unix socket；唯一数据通道是继承 IPC。

依据：[内核 Landlock](https://docs.kernel.org/userspace-api/landlock.html)、[seccomp filter](https://www.kernel.org/doc/html/next/userspace-api/seccomp_filter.html)、[systemd scope 生命周期](https://www.freedesktop.org/software/systemd/man/latest/systemd.scope.html)、[用户管理器 StartTransientUnit](https://www.freedesktop.org/software/systemd/man/latest/org.freedesktop.systemd1.html)、[内存与任务限额](https://www.freedesktop.org/software/systemd/man/255/systemd.resource-control.html)。

## 已执行证据

Windows 最终定向平台/签名/worker/故障/隔离/恢复/页范围检查 **104/104**：`artifacts/avalonia/tests/a6-final-platform.trx`。前一轮 **78/78**：`artifacts/avalonia/tests/a6-windows-platform.trx`；原生 bridge smoke 10,000 调用通过。没有执行全量 Core。

Linux 使用普通 probe 用户 uid 1001，宿主由该用户的 systemd 服务启动。VM 为 A4 的独立 A6 overlay，未修改既有受测快照。原始结果在 VM `/home/probe/avalonia-a6`；以下小型证据已复制到仓库 `artifacts/avalonia/a6-linux/evidence-16`：

| 检查 | 结果 |
| --- | --- |
| `result-14` | API v1、v2 模块/批处理/日志、冻结动作/一次事务/撤销和无重新授权 redo、引擎时间/内存限额、失败后新轮次 UTF-8，6 组通过 |
| `sandbox-15` | 独立 native probe 32 项文件/网络/进程/系统调用负向检查；不依赖 JS 缺少宿主 API 的假设 |
| 同一 native probe | 第 33 个线程被拒绝；320 MiB 实际触页被 256 MiB cgroup OOM 杀死（kernel 日志 `a6-oom.log`）；5 秒 scope 超时终止 |
| `lifetime-16` | 只杀测试宿主，worker **2.02 ms** 内消失；活跃租约保留，崩溃后仅恢复 1 个登记轮次，未登记目录保留 |

失败均保留：初次跨 RID 图错误复制 Windows DLL，改为全图显式 `-p:QuickExcelNativeRid=linux-x64` 并核验包；systemd 异步 job 返回早于内核迁移，增加有限等待及内核核验；ELF interpreter 需要单独执行白名单；CLR 创建线程需要仅限自身的 sched_setaffinity；exec 后 dumpability 重置，最终 guard 再次关闭且禁止重新开启。没有提高资源预算、删除负向断言或放宽用户数据/网络权限。

## 信任保护

Windows 保留原 DPAPI v1 字节与用途绑定。Unix 使用平台密钥服务保护的随机 256 位密钥与 AES-GCM v2 密文，路径和 provider 进入认证用途；密钥不会写入普通文件或进程参数。Linux 通过系统 `secret-tool` 接入 Secret Service，查询和保存均有超时；缺失、锁定、错误与篡改均拒绝持久自动信任。签名验证与保护服务失败分开处理，仍可由用户明确选择仅本次授权。macOS 的 Data Protection Keychain 适配已编译，使用仅本机、解锁可访问、不跨设备同步的条目；尚未在 Mac 执行。

`test-linux-secret-service.py` 在独立 `dbus-run-session` 内建立一次性 GNOME Keyring，密码只经 stdin 传递；不使用正常桌面密钥环。`trust-20/trust-21` 与最终 `secret-28` 均通过：签名 pin 保存/重开、路径绑定、密文篡改拒绝、撤销以及服务不可用时不落明文、不自动授权。依据：[libsecret](https://gnome.pages.gitlab.gnome.org/libsecret/)、[secret-tool 实现](https://github.com/GNOME/libsecret/blob/main/tool/secret-tool.c)、[Apple Keychain](https://developer.apple.com/documentation/security/ksecclassgenericpassword)。

## CUPS 打印

Linux/macOS 通过系统 libcups 查询队列、纸张和硬边距，选择不匹配时禁止提交。打印范围与份数受界限检查；`lp` 使用参数列表传递同一个已完成的 PrintSession PDF，显式关闭缩放、合版和封面页，设置方向、纸张、份数与页范围。提交有 30 秒限时，PDF 租约保留至客户端结束；中断时明确提示可能已入队，不自动重试。Unix PDF 临时会话目录为 0700。

真实自动验证使用普通用户拥有的临时 CUPS 服务、独立端口 8631 和 IPP 虚拟打印机 8632，不访问系统队列或物理打印机。系统 cupsd 的 AppArmor profile 禁止读取用户私有配置，夹具复制系统二进制到自己目录运行；没有修改或禁用系统 profile，产品仍访问正常系统 CUPS 服务。

`cups-25` 与最终 `cups-28` 均通过 4 项应用检查和独立交接检查：调度器接收的 PDF SHA-256 与生成文件完全一致；cups-filters 将应用舍入后的 A4 尺寸规范化为毫米，打印机端文件因此并非逐字节相同。Poppler 独立验证页数、全部文字、中文及坐标，最大差异 **0.044436 点**，在工程几何比较采用的 0.1 点容差内（不是实机打印精度门槛）。保留原始两份 PDF/哈希；不把此结果当作物理打印通过。早期未安装 cups-filters 时明确拒绝 PDF，加入依赖声明，没有提交 raw 作业绕过过滤与纸张核对。

依据：[CUPS API](https://openprinting.github.io/cups/doc/cupspm.html)、[lp 提交语义](https://openprinting.github.io/cups/doc/man-lp.html)、[打印选项](https://www.cups.org/doc/options.html)。

## 最终工程复验与包

- `artifacts/avalonia/a6-linux/evidence-26/desktop-26`：普通 probe 用户、Xvfb 1440×900、WebKitGTK 2.52.6，工作区 **32**、网格 **77**、Web **42** 项通过。Linux 本轮实际包含独立 worker、actions、撤销/数据 redo、持久信任与撤销。WebKit 沙箱保持启用；退出后无宿主/浏览器残留。Xvfb 软件渲染警告保留，不声称 GPU、物理输入或辅助技术通过。
- `artifacts/avalonia/a6-linux/evidence-29/functional-28`：6 组 native/API/动作功能再次通过；`sandbox-28`：32 项 syscall/文件/网络负向探测、32 线程上限、256 MiB 实际 OOM 限额、5 秒 scope 生命周期再次通过。
- 同目录 `lifetime-28`：宿主被测试杀死后 worker **5.21 ms** 内消失；活跃租约和未登记目录保留，仅回收已终止的登记轮次。最终载荷检查增加 512 文件、256 MiB 总大小、单文件 128 MiB 限制，并拒绝悬空链接。
- `secret-28`：真实系统保护服务 4 项 + 服务不可用 1 项通过；`cups-28`：真实 CUPS 交接、纸张拒绝与独立 PDF 几何核对通过。
- `install-27`：从正式 tar 包提取后，以普通用户直接启动（不需要把宿主包装成 systemd 服务），**7** 项生产安装自检通过，包含真实脚本隔离、Monaco/PDF.js、保存重开及导出。
- `deployment-29`：生产包自检、测试参数拒绝、中文/空格/只读安装目录、Web 缺失、worker 缺失，**5** 项通过。只在本次夹具复制目录中修改精确文件；缺失时明确失败，没有替代编辑器或低权限后备。

Linux 自包含候选包：`artifacts/avalonia/platform-releases/linux-x64-20261102-213902-7f3851b0/application`；同目录 `QuickExcel-linux-x64.tar.gz` SHA-256：`FE0ABB3EDE94B421DAAE41027B511AC77FE8881447C0BC276C5F7BBB781DDD91`。包记录 native 编译器/源码/二进制哈希、NuGet 与 Web 锁文件、许可证、诊断关闭状态及源树身份。RID 恢复和普通构建锁文件分离；跨构建显式传递 RuntimeIdentifier 与 QuickExcelNativeRid，避免 .NET restore 引入宿主 RID。每次使用新 build 目录，不让旧 Windows DLL 混入 Linux worker。包内没有测试 probe、Node、WPF、设计器或符号文件；worker 没有 UI/Web/Skia 载荷。

复现入口：

```powershell
# 在目标平台先构建 native；Linux 使用 CC=clang-18。
# bash QuickExcel.Scripting/native/build.sh
# bash QuickExcel.ScriptWorker/native/build-linux.sh
./scripts/publish-avalonia-unix.ps1 -Rid linux-x64 -CreateArchive
```

Linux 依赖、普通用户启动、自检与失败条件在 `scripts/README-avalonia-unix.txt`。开发验证工具 `tools/QuickExcel.PlatformSmoke` 和其中 TestProbe 绝不可进入产品；脚本 `test-linux-worker-lifetime.py`、`test-linux-secret-service.py`、`test-cups-handoff.py`、`test-avalonia-linux-package.py` 均只处理显式的新夹具路径。

## macOS arm64 实现、证据与未完成项

2026-11-06。macOS 26.6 arm64 实现了脚本执行的独立路径，不改变 Windows/Linux 契约：

- **原生引擎**：以同一 WebKit 2.54.2 提交构建 JSCOnly arm64 的 JavaScriptCore。GTK 发布包缺失的 104 个 macOS 源文件由 `tools/QuickExcel.JscProbe/macos-sources.json` 固定哈希补入；`WebKitFS.cmake` 的 framework 头目录与 `PlatformJSCOnly.cmake` 的重复 TimeZoneCocoa 修正记录在 `source-adjustments.json` 并写明原因。WebKit Apple 端口按上游设计链接系统 `libicucore`（macOS 26.6 为 ICU 78.1），macOS 包不再分发 ICU 77 动态库。
- **worker**：宿主建立私有 0700 载荷目录、只允许扁平常规文件、记录 SHA-256 并持有独占租约；启动同目录 `quickexcel-sandbox`，launcher 等待单字节屏障后才应用默认拒绝的 Seatbelt 配置并 exec worker。worker 在读取任何协议字节前调用 `qx_macos_lock_worker` 验证 `/etc/passwd` 读取、载荷写入和网络连接均被拒绝，并启动 kqueue 父进程退出看门狗与物理内存/线程轮询看门狗；宿主通过 `libquickexcel-sandbox` 的固定参数包装器再次核对文件写入、网络和 fork 拒绝。没有同权限或 in-process 后备。
- **资源**：执行 GC 堆硬限 128 MiB（Windows/Linux 仍 64 MiB）；256 MiB 进程上限由 25 ms 轮询的 phys_footprint/线程看门狗执行，是有界超调的轮询上限，**不等于** cgroup 硬限。一次 768 MiB 快速分配的实测在 526,239,160 字节处被终止，超调真实存在。线程上限 32。
- **信任**：Data Protection Keychain 适配已在本机执行。未签名测试宿主没有 application-identifier，Keychain 返回 -34018，持久信任失败关闭，只允许用户明确的一次性授权；签名包上的持久信任、权限和公证仍未验证。
- **本机证据**：发布包 `--self-check` **27/27**，包含 JSC 隔离 worker、隔离 TypeScript 编译器、API v3 动作/撤销、P2 对象、WKWebView Monaco/PDF.js、只读与恢复；定向测试 **24/24**（`MacScriptWorkerTests` 5 项，策略/执行/信任 19 项）；独立协议往返 `41→42`；父进程被 SIGKILL 后 worker 在 0.6 s 内消失；看门狗在超限分配时杀死 worker。
- **未完成**：Developer ID 签名/公证的实际流程、多文件 `.app` 封签（需 single-file/Resources 布局）与签名包持久信任验证；物理打印、输入、DPI、辅助功能和干净设备验收；Windows AppContainer 故障 fixture 不适配 osx-arm64，不纳入 macOS 定向门禁。仅列 entitlement、普通子进程或 `sandbox-exec` 仍不能替代上述证据。

已准备 arm64 native 构建、Keychain 保护、CUPS 打印适配、RID 发布以及 `.app`/签名/公证工具 `package-avalonia-macos.py`。该工具仅允许 Mac arm64，在 repository artifacts 内建立新包并逐个签署 Mach-O 与托管 DLL；多文件 .NET bundle 的最终封签仍受 `codesign` 将 `Contents/MacOS` 中非代码文件视为未签名子组件的限制，当前未完成，需改为 single-file 或 Resources 布局。只有明确提供 Developer ID 与 `--notary-profile` 才会提交 Apple 服务；当前没有执行任何签名或上传。本机预览直接运行原始发布目录，不把 `.app` 签名/公证标为通过。签名后完整哈希清单在 bundle 外记录，避免修改已签名资源。仅声明 .NET 所需的 allow-jit，不开启调试或禁用库校验。依据：[.NET macOS 部署](https://learn.microsoft.com/en-us/dotnet/core/deploying/macos)、[Apple 公证](https://developer.apple.com/documentation/security/notarizing-macos-software-before-distribution)。

macOS worker 已实现并通过本机工程检查，但签名/公证、签名包持久信任和物理门禁仍未完成，不能把本机自检当作完整平台验收。旧缓存保持 blocked/stale；没有证据时不扩大执行范围。App Sandbox/XPC 仍是未来可选架构；当前的 Seatbelt 路径必须保留默认拒绝、宿主核对、父进程/资源看门狗与负向证据，不能用普通子进程或仅列 entitlement 替代。依据：[Apple App Sandbox](https://developer.apple.com/library/archive/documentation/Miscellaneous/Reference/EntitlementKeyReference/Chapters/EnablingAppSandbox.html)、[XPC](https://developer.apple.com/documentation/xpc)。不把 Mac 完整支持标为完成。

用户延期的是实机验证，不是降低隔离标准。当前工程包默认交付目标仍为 Windows，Linux 为受测 Ubuntu 候选；macOS 不在默认交付支持范围。A7 已清理 WPF 生产入口并保留归档回退，见 [68 号记录](68-avalonia-retirement.md)；Mac 设备/隔离工程和各平台实机门禁继续独立跟踪。Windows 原 M4/M5/M6 未关闭门槛保持原状态。


## macOS worker 复用修复（3.8.1）

2026-11-06。macOS API v3/v4 的授权工作簿接入与 Windows 相同的 `SharedScriptWorkers`：按 SessionId、授权生命周期 token 和角色分池，UDF 与 Actions/Events 各自惰性启动独立进程/JSC 上下文。成功的调用、批次、计算轮和空闲时间均保留上下文，避免每轮重新复制/校验完整载荷、启动进程和加载模块。API v4 初始化入口随 Actions/Events 上下文执行一次；失败、撤权、脚本修改和关闭沿用 Windows 的终止、事务回滚和禁止自动重放规则。控制台、无授权生命周期的显式调用、编译器和 Linux 路径仍保持一次性执行。

macOS launcher 新增宿主选择的 `persistent` 模式。原来的 `RLIMIT_CPU=30/31s` 是进程启动以来的累计 CPU 时间，会误杀完成多次正常轮次的持久进程，因此只对持久执行取消该累计期限；一次性执行和编译器仍保留。每次原生操作一秒、每个活动轮最多五秒的宿主截止不变，空闲不消耗活动轮期限。Seatbelt 默认拒绝、源码发送前宿主核验、父进程退出、256 MiB footprint/32 线程轮询看门狗、私有只读载荷与登记租约均保留，空闲期间也继续受保护。该内存限制仍不是 cgroup 等价硬限，JSC-MEM-001 仍 OPEN。

macOS 的实时动作在进入原生操作前准备独立 UDF 上下文（仅有 formula-functions 绑定、已授权且带实时模型的工作簿；不包括控制台）。私有载荷的持久复制可能超过一秒，不能留到动作中的模型回调才启动。准备仍计入动作原有五秒宿主期限，只加载模块、不调用 UDF 或发布计算结果。后续动作、事件和重算使用相同 UDF worker。初始化中的首次公式读取与动作内冷 UDF 写后读均覆盖这一时序。

持久 worker 的异步退出清理与新启动可能竞争租约根锁；macOS 现在针对 Darwin EWOULDBLOCK 做有截止、可取消的等待，不绕过独占所有权。源码修改先取消 worker 生命周期时，正在执行的旧计算按过期请求处理，避免与计算轮取消的先后时序将其误报为 worker 通信故障。

### 本次验证与交付

- 原生 JSC 固定来源重建、macOS sandbox 重建通过；来源校验与原有签名/内存限制边界保持。
- osx-arm64 / .NET 10 Release 定向测试 **147/147**，未运行全量回归。覆盖 API v3/v4 的 PID/闭包跨轮保留、空闲超过五秒、两个角色/不同工作簿隔离、源码失效/撤权/关闭清理、初始化不重放、冷 UDF 事务读取/回滚、控制台一次性状态、根锁竞争及取消、一次性/持久 worker 的超限分配，以及累计 **32 秒 CPU** 后仍复用同一 PID、随后无限循环仍触发超时终止。记录：`artifacts/test-results/macos-worker-reuse/worker-reuse-final.trx`；中间失败记录保留用于说明已修复的租约竞争、失效取消时序和冷启动计时问题。
- 独立于测试/构建的真实 worker 小样本：每轮一个批次、**64 次简单 UDF 调用**；一次性路径 10 轮、10 个 PID，中位 **1659.309 ms/轮**；持久路径首轮 **1661.039 ms**，随后 30 轮仅 1 个 PID，中位 **1.341 ms/轮**。包括创建/取得 runtime、批次调用与释放轮次；所有返回值逐个校验。数据为自包含测试载荷，原始样本 `artifacts/test-results/macos-worker-reuse/performance.json`，采样程序在 `artifacts/macos-worker-probe/`。这衡量重复调用的启动开销，不是复杂工作簿整体性能结论。
- macOS 生产发布成功，完整 worker、247 个离线 Web 资源、许可证及 macOS 对应源码齐全，无测试 fixture/诊断载荷。发布有两个既存 CS0649 警告（诊断关闭后的 WebScriptsOnly/WebModulesOnly），无错误；生产自检 **32/32**，新增覆盖原来只在 Windows 检查的 API v4 初始化/跨轮状态和 IStorage。候选与证据：`artifacts/avalonia/platform-releases/osx-arm64-worker-reuse-3.8.1/`。
- 通过自检并逐文件校验后，已更新本地原始启动包 `artifacts/publish-mac-app`；`release.json` 为 **3.8.1 / passed**，入口为 `Start-QuickExcel.sh` 或 `QuickExcel.Avalonia`。旧包保留在上述候选目录的 `previous-publish-mac-app`，切换记录为 `activation.json`。既有 `.app` 预览归档不作为这次已更新的启动入口，签名/公证状态未改变。
- 按仓库要求尝试 `scripts/build-quickexcel.ps1 -SkipTests -RunSmoke`，本机没有 `pwsh`，命令无法启动。Windows 包未发布或验证；本机原本没有 `artifacts/avalonia/active-windows.json`，本次也未创建或切换该指针。失败记录：`artifacts/test-results/macos-worker-reuse/windows-publish.log`。
