# ADR：M4 Windows AppContainer worker 隔离

2026-10-28。状态：无能力 AppContainer 已接入实际桌面公式及 actions worker；本机系统调用、IPv4/IPv6 回环、DNS 的正反对照及 WFP 拒绝原因取证已完成。部署矩阵与 M6 QA-OS-01 尚未完成，继续限定显式授权的可信脚本使用。

## 实际 worker 启动与清理

`ScriptWorkerRuntime` 的唯一启动路径是 `WorkerProcess`：每轮创建随机 AppContainer profile，能力列表为空；`WorkerSandbox` 在当前用户 LocalAppData 的 `QuickExcel.WorkerSandboxes.v1` 下原子创建私有 DACL 目录，将 worker 的平铺二进制依赖复制到本轮 `payload`，仅给予本轮 SID 读/执行权限。宿主用户保留完全控制。拒绝来源路径的 reparse point；不改安装目录、Windows、.NET 或整个用户目录 ACL，也不授予网络能力或回环豁免。部署测试与剩余环境见 [回收与部署验证](16-m4-recovery-and-deployment.md)。

用 `STARTUPINFOEX` 显式继承 stdin/stdout/stderr 三个句柄，并以挂起、无窗口、Unicode 环境块创建进程。先配置命名 Job（256 MiB、单活动进程、关闭即终止），通过 `PROC_THREAD_ATTRIBUTE_JOB_LIST` 在 CreateProcess 时加入，消除创建后再挂接 Job 的宿主崩溃窗口。在 CLR 或 worker 执行前，宿主读取 OS token，核验 `TokenIsAppContainer=1`、SID 与本轮一致、`TokenCapabilities` 数量为零，核验 Job 成员身份，最后恢复主线程。任何一步失败都结束启动，公式与 actions 均无普通同权限进程或进程内回退。

环境为白名单：SystemRoot/WINDIR、profile 内 LOCALAPPDATA/TEMP/TMP、关闭 .NET diagnostics 和 CLI telemetry；不继承 PATH、代理、用户秘密、startup hooks 或 profiler 配置。运行库由机器注册的 .NET 10 安装解析。脚本不接收工作簿文件句柄或路径；输入继续通过有界 IPC。AppContainer 自己的 profile 仍可写，不等于零磁盘访问，也尚无磁盘配额。

取消、超时及通信故障先关闭 Job/终止进程；Dispose 等待退出（1 秒），关闭管道，再删除私有 payload 与 profile。文件映像释放可能略晚于进程退出信号，删除遇到占用/拒绝时最多再重试 1 秒。清理失败作为 `WorkerCleanup` 报错并保留回收记录；启动阶段看门狗与用户取消分别映射为超时和取消。安装目录不产生临时 ACE。宿主异常退出时 Job 杀进程；桌面下次启动以及每轮 worker 启动前回收孤立资源，独占租约保护其他活动实例，Job 未清空时拒绝回收。完整协议和故障证据见 [回收与部署验证](16-m4-recovery-and-deployment.md)。

## 回环和 DNS：正反对照与拒绝取证

普通网络验证（无需管理员；profile 需要正常用户环境）：

```powershell
dotnet run --project .\tools\QuickExcel.SandboxProbe\QuickExcel.SandboxProbe.csproj -c Release --artifacts-path .\artifacts\build -- --network D:\Projects\QuickExcel
```

附加 WFP 证据（只读管理员 helper 会触发 UAC；探针本身保持原用户权限）：

```powershell
dotnet run --project .\tools\QuickExcel.SandboxProbe\QuickExcel.SandboxProbe.csproj -c Release --artifacts-path .\artifacts\build -- --network-wfp D:\Projects\QuickExcel
```

两种模式先运行真实 `QuickExcel.ScriptWorker.exe` 的 IPC/UDF，再以**生产同一个 `WorkerProcess`** 启动开发用系统调用载荷；网络 API 不加入发布 worker。正向对照为同用户 Job 进程；父进程分别启动 IPv4/IPv6 回环监听器并记录是否真正 accept。所有网络请求固定为 `example.com.` DNS 查询、DNS 服务器的 53 端口或 `1.1.1.1:443` 握手，不发送工作簿或用户数据。包含 .NET DNS、`DnsQuery_W` wire-only/bypass-cache/no-hosts/no-NetBT/no-multicast、cache-only、原始 UDP/TCP DNS。

本机证据：`artifacts/sandbox-probe/20261028-032304-ca1b855a/`；`network.json` 保存原始 API 结果，`wfp-summary.json` 独立关联事件时间、进程路径/监听端口、filterId 和实际 BLOCK 规则。不能把“API 超时”本身改写成拒绝。

| 检查 | 正向对照 | 生产启动器隔离进程 | WFP 证据 |
| --- | --- | --- | --- |
| IPv4 回环 | 连接且 accept 成功 | 11 秒超时，未 accept | 接收端 `AppContainer Loopback`，filterId 264373，4 条 drop |
| IPv6 回环 | 连接且 accept 成功 | 11 秒超时，未 accept | 接收端 `AppContainer Loopback`，filterId 264374，4 条 drop |
| 外网 TCP | 成功 | WSAEACCES / 10013 | `UWP Default Outbound Block Rule` |
| DNS TCP | 收到 DNS 响应 | WSAEACCES / 10013 | 同上，目标 53 端口 |
| 系统 DNS、原生 wire-only、直接 DNS UDP | 均成功 | 超时 | 同一隔离载荷的 DNS UDP 多次重试被出站规则拒绝（filterId 6340378），包括配置的两条 DNS 路径 |
| DNS cache-only | 本轮 9701，无记录；先前轮曾命中 | 9701，无记录 | 缓存无记录不等于 OS 拒绝，不用作门禁 |

filterId 是本机当次标识，不硬编码到验证条件。回环事件归属于接收端父进程，因此仅按隔离子进程 appId 查询会漏掉回环拒绝。helper 分别查询子进程和监听端口；不修改防火墙、事件缓冲设置或回环豁免。`NetworkIsolationDiagnoseConnectFailureAndGetInfo` 对 DNS/外网报告缺少 InternetClient；它只是补充能力诊断，回环返回 NONE，不能单靠该 API 证明回环可达。

`--network` 退出 0 仅表示流程完成；`--network-wfp` 还要求正向对照通过、隔离侧未观察到可达、两种回环及 DNS UDP 均存在可关联的 BLOCK 事件，否则退出 2（证据不足）。所有报告的 `releaseIsolationValidated` 保持 false。普通用户读取 WFP 返回 5；首轮管理员查询的 timewindow 参数在本机返回 FWP_E_INVALID_INTERVAL，最终不用该参数并在汇总中按时间与路径筛选。失败证据保留在 `20261028-031849-a8540655`。

离线复核已保存证据，无需管理员或联网：

```powershell
dotnet run --project .\tools\QuickExcel.SandboxProbe\QuickExcel.SandboxProbe.csproj -c Release --artifacts-path .\artifacts\build -- --analyze-wfp D:\Projects\QuickExcel\artifacts\sandbox-probe\20261028-032304-ca1b855a
```

实现依据：[AppContainer 启动](https://learn.microsoft.com/en-us/windows/win32/secauthz/implementing-an-appcontainer)、[DNS 原生查询](https://learn.microsoft.com/en-us/windows/win32/api/windns/nf-windns-dnsquery_w)、[网络能力诊断](https://learn.microsoft.com/en-us/windows/win32/api/netfw/nf-netfw-networkisolationdiagnoseconnectfailureandgetinfo)。这些 API 文档不代替本机 WFP 取证。

## 本轮回归和发布验证

新增 `M4IsolationTests` 核验真实 worker token/不同轮身份、文件/写入/最小继承句柄/子进程限制、环境白名单及正常/取消/强杀后的 profile/payload 清理。开发 fixture 直接调用托管文件 API 与 Win32 API；句柄正向对照保留在原探针中。现有模块、actions、合批、协议攻击、无限循环、内存、取消、过期和重算恢复用例全部改经 AppContainer 启动路径。

首次完整回归 608 项通过（后续再补两项清理回归），native smoke PASS。发布树 `artifacts/publish/20261028-032125-c4575fc1` 的七个 WPF smoke 结果均为新 PASS，实际停止按钮单次 23 ms；详情在对应 `artifacts/smoke/20261028-032125-c4575fc1`。

最终回归：Release solution 0 警告/0 错误，Core 全套 **610 项通过，0 跳过**；native smoke PASS；`scripts/build-quickexcel.ps1 -RunSmoke` 成功。最终发布目录 `artifacts/publish/20261028-033052-429141c3` 与默认 `artifacts/QuickExcel` 已更新；七个 WPF 结果文件见 `artifacts/smoke/20261028-033052-429141c3`。元数据记录 `isolationMode=appcontainer-no-capabilities-v1`，最终 OS 放行标记仍为 false。ExcelIO tracked diff 为空。本段在最终发布验证后补记；未创建 ZIP，也未提交当前工作树。

## 决策与依据

以下保留首轮开发探测的决策依据。Job 负责生命周期、提交内存和单进程约束。删除令牌特权、LUA token 不能单独提供文件/网络访问隔离；现已采用无能力 AppContainer、显式句柄列表和挂起创建，运行库授权改为每轮私有副本的只读 ACL。

Microsoft 的 [AppContainer 启动说明](https://learn.microsoft.com/en-us/windows/win32/secauthz/implementing-an-appcontainer)描述包 SID、能力和 ACL 的共同作用；[进程属性文档](https://learn.microsoft.com/en-us/windows/win32/api/processthreadsapi/nf-processthreadsapi-updateprocthreadattribute)规定显式继承句柄列表；[受限令牌文档](https://learn.microsoft.com/en-us/windows/win32/api/securitybaseapi/nf-securitybaseapi-createrestrictedtoken)区分删去特权和限制 SID。选择依据包括这些机制契约与下述本机实测，不把文档描述代替实测。

## 可复现开发探测

从仓库根运行（正常用户环境，Windows x64 / .NET 10 SDK）：

```powershell
dotnet run --project .\tools\QuickExcel.SandboxProbe\QuickExcel.SandboxProbe.csproj -c Release --artifacts-path .\artifacts\build -- D:\Projects\QuickExcel
```

最后一个参数应为实际仓库根目录。工具不进入主 solution、桌面项目依赖或发布包。它把自身及相同 QuickJS DLL 复制到 `artifacts/sandbox-probe/<唯一运行目录>/payload`，仅使用旁边的合成 canary 文件。创建随机临时 AppContainer profile，仅给该 SID 的 payload 读/执行权限；结束后恢复原 ACL、删除 profile，保留 JSON 证据。发生清理失败时报告并返回失败。未修改机器上的 .NET/Windows 目录 ACL，也不修改防火墙或回环豁免。

所有子进程均无窗口、挂起创建，先配置与实际 worker 相同的 Job，再恢复执行。三种策略使用 stdin/stdout/stderr 的显式继承列表；额外正向对照故意把 canary 句柄加入列表，证明确实能读到它。探测直接执行托管文件和 socket 系统调用，包含加载 QuickJS DLL、包内模块和 UDF 调用，不只检查 JS 是否缺少 API。

网络探测包括本机临时回环监听器、`example.com` DNS 查询及 `1.1.1.1:443` TCP 握手。不传输工作簿、源码或用户数据。超时单独标记 `inconclusive:timeout`，不能当成明确的访问拒绝。命令退出 0 只表示探测流程及清理完成，JSON 始终保留 `releaseIsolationValidated: false`。

## 本机结果

环境：Windows `10.0.22621.0`，.NET `10.0.12`。证据：`artifacts/sandbox-probe/20261028-024603-084bb5b9/probe.json`。

| 探测 | Job + 最小句柄 | 删减令牌 + Job | 无能力 AppContainer + Job |
| --- | --- | --- | --- |
| .NET、native DLL、ES 模块及 UDF | 成功 | 成功 | 成功 |
| 合成私有文件读取/写入 | 允许 | 允许 | UnauthorizedAccess / 0x80070005 |
| payload 写入 | 允许 | 允许 | UnauthorizedAccess / 0x80070005 |
| 未列入继承列表的 canary 句柄 | 不可读 | 不可读 | 不可读 |
| 子进程创建 | Win32 1816 拒绝 | Win32 1816 拒绝 | Win32 1816 拒绝 |
| 外网 TCP | 连接成功 | 连接成功 | Socket AccessDenied |
| 回环连接 | 成功 | 成功 | 超时，拒绝原因待取证 |
| DNS | 成功 | 成功 | 超时，拒绝原因待取证 |

正向句柄对照能读取 canary；每次探测前重置文件指针，避免把先前读取到文件末尾误判为隔离。实际 AppContainer token 标记为 true。基线 token 的 HasRestrictions 也为 true，因此不能仅凭该布尔值把删减令牌与基线区分。最终 profile 删除 HRESULT 为 0，ACL 清理成功。

受限执行环境的首次 profile 创建返回 0x80070002，证据保留在 `20261028-024309-23b08922`；正常用户环境完成探测。中间运行 `20261028-024324-742ee1b9` 保留了较早的超时分类，以上最终记录明确纠正为待确认。

## 尚未关闭的工作

1. 推进独立标准/管理员账户、真实 Program Files、跨版本更新和干净机器部署验证；本机路径/ACL/替换卸载结果见部署记录。宿主崩溃后资源回收已实现；临时存储磁盘额度仍未实现。
2. 扩大系统调用边界测试：注册表、设备、命名对象、其他进程对象、代理/系统服务、缓存和其他 DNS/网络拓扑。当前 AppContainer 不是 LPAC，系统向所有应用容器开放的资源仍可能可读，不宣称零系统资源访问。
3. 严格确定性执行策略仍未开放，保守执行的 Date/Math.random/eval/Function 语义不变。
4. M6 在实际发布目录和干净机器重新验收 QA-OS-01，包含被攻陷的 native worker 威胁模型。当前仍仅供显式授权的可信源码使用。
