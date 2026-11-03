# M4 首切片：独立脚本 worker

> 2026-11-03 引擎迁移更新：当前生产路径为固定 JSC、IPC 6，隔离 TS 编译与同步模块/中断契约见 [71 号实施记录](71-jsc-typescript-implementation-log.md)。执行仍为 256 MiB 进程 / 256 KiB 栈 / 1 秒操作 / 5 秒整轮；64 MiB 引擎硬限额 `JSC-MEM-001` 未实现，用户明确延期且不作为推进门禁。下文 QuickJS、旧协议、预算及构建入口均为当时版本的历史证据，不代表当前实现。

> 2026-10-28 最新更新：公式与 actions 的实际 worker 已接入无能力 AppContainer、创建时加入 Job、挂起启动与令牌核验、最小继承句柄、只读私有 payload、环境白名单和 profile 清理；宿主崩溃后由下次启动回收资源，详见 [回收与部署验证](16-m4-recovery-and-deployment.md)。IPv4/IPv6 回环及 DNS 的本机 WFP 拒绝证据见 [隔离 ADR](15-m4-os-isolation-adr.md)。以下 Job-only 安全边界与 250 ms 清理说明为首切片历史；当前等待退出 1 秒，文件删除占用重试最多再 1 秒。M6 最终门禁仍未完成。

> 最新更新：UDF 有界合批和 16 MiB 消息预算已实现，IPC 当前协议 4；见 [合批记录](14-m4-batching.md)与 [OS 原型 ADR](15-m4-os-isolation-adr.md)。下文保留历次切片的版本、预算及验收证据。

> 2026-10-28 更新：包内 ES 模块及 actions 事务已接入，IPC 当前协议 3；见 [actions 契约与验收](13-m4-actions.md)与 [模块契约](12-m4-modules.md)。下文保留 2026-10-27 首切片的协议 1 和发布证据，模块/动作待办以最新记录为准。合批及 OS 访问隔离仍待完成。

2026-10-27，基于已提交的 `2e749ff`（M3 及日期显示修复）。版本 `0.5.0-preview.1` 是 M4 进行中的可信脚本预览，不代表 M4 或 Beta 全部完成。

## 本轮范围

| 任务 | 状态 | 已实现及剩余范围 |
| --- | --- | --- |
| JS-007 | 进行中 | 独立 Contracts/Worker、握手、长度前缀协议、类型桥、批次请求、完整身份检查及反压已接入；计算计划合批与跨轮驻留 worker 待后续 |
| JS-008 | 首切片实现 | 宿主看门狗、取消即终止、会话失效取消、断管/强杀失败、显式重新计算恢复；本轮无自动重放 |
| FMT-008 / JS-009 | 待开始 | 包内模块加载、actions 必需能力、动作快照/patch/二次校验/一次撤销尚未实现 |
| JS-010 | 保持 M3 策略 | 每轮新进程/上下文、稳定顺序完整重算，不复用 UDF 和下游；严格确定性策略尚未开放 |
| SEC-001 / JS-014 | 部分原型 | Windows Job 生命周期/256 MiB 提交内存/单进程限制；受限令牌、AppContainer、文件与网络访问拒绝尚未验证 |

## 当前执行契约

桌面组合根只选择 `ScriptWorkerRuntimeFactory`。进程启动、握手、Job 配置或通信失败都会结束本轮；不会退回进程内 QuickJS。`QuickJsRuntimeFactory` 继续用于原生桥和历史测试。

当前为**每次授权重算一个专用进程**，而不是跨轮常驻进程池。一个进程绑定一个会话、输入修订、计算代次和脚本代次，加载一次，结束即销毁。这样先验证故障边界与上下文重建；启动开销和跨轮复用留到测量之后决定。重算不授权脚本；打开/迁移/保存/导出均不运行源码。

Core 的 `CalculationRequest` 带会话取消令牌。编辑、撤销/重做、源码变更、启停、再次重算及关闭使旧计算令牌失效，仍保留完整 stamp 的最终提交检查。旧请求返回 `false`；用户主动取消抛出取消异常，由桌面显示停止状态。关闭后旧错误不会改写新会话的授权或诊断。

桌面计算期间仍禁用编辑工作区，状态栏的「停止计算」保持可用。点击后终止 worker 并保留旧缓存，恢复编辑/保存能力。完整并发编辑工作流仍属于 M5。

## 协议、预算与故障

- 专用重定向 stdin/stdout 二进制管道；stderr 不承载协议，不记录源码、单元格值或任意本地路径；累计超过 16 KiB 终止进程。
- 四字节小端长度 + UTF-8 JSON。收发均限制 4 MiB；先读长度再分配，出站序列化也经过容量受限缓冲。深度 24，拒绝重复键、未知字段、非法类型、截断数据及未知版本。
- 握手匹配协议 1、API `[1,2]`、固定引擎声明 `QuickJS-NG/0.17.0` 和策略 `conservative-worker-v1`。加载校验源码 SHA-256，限 1 MiB UTF-8。引擎二进制来自固定源码构建；握手不是二进制远程证明。
- 每次响应严格匹配 SessionId / WorkerInstanceId / RequestId / Revision / CalculationGeneration / ScriptGeneration。此版本仅计算请求，计算代次不允许 null。重复 requestId 和跨轮请求使 worker 退出。
- 保留 Blank / Text / Number / Boolean 类型与 v1 一维、v2 显式二维区域形状。领域 Error 在调用前传播，不依据字符串外观猜错误类型。标量标签显式；集合元素采用 JSON 原生类型，禁止嵌套集合。
- 最多 256 次调用/批次、255 参数/调用、4096 注册函数、1,000,000 位置/范围，仍受帧字节限额。单进程同时只允许一个请求，**零等待队列**，并发请求立即拒绝；当前计算器发送单调用批次，尚未实现拓扑层合批优化。
- 每次引擎调用 1 秒，整个授权轮（含启动/握手/加载）最多 5 秒。取消直接关闭 Job/终止进程；无需等待 JS 合作。Dispose 最多等待 250 ms 清理。测试中的宽松超时只用于避免 CI 调度抖动，不能当作 p95 性能证据。
- 引擎分配 64 MiB、栈 256 KiB；Job 进程提交内存 256 MiB、活动进程数 1、关闭最后句柄终止进程。Job 必须在发送源码前配置成功。
- 日志最多 100 条，每条最多 4000 UTF-8 字节，按 Unicode rune 截断；结果日志与当前请求绑定。普通函数异常成为该格 Script 错误；协议故障、进程退出、引擎中断和资源限额结束整轮，拒绝部分缓存。
- 原生桥提供 `qx_was_interrupted`，包含异常对象转字符串期间的中断。worker 不靠用户可伪造的异常文本识别超时；已中断句柄不再调用。

## 安全与部署边界

Job 是资源和生命周期约束，**没有提供文件/网络访问隔离**。受限令牌、最小句柄集合与 AppContainer/访问策略的开发级探测仍待完成。当前只适合显式授权的可信源码，不能宣称可安全运行任意第三方脚本。Microsoft 的 [Job Objects](https://learn.microsoft.com/en-us/windows/win32/procthread/job-objects) 和 [限制结构](https://learn.microsoft.com/en-us/windows/win32/api/winnt/ns-winnt-jobobject_extended_limit_information) 描述的是进程/资源管理能力；M6 仍需单独通过 OS 访问策略门禁。

桌面与测试构建通过 `scripts/ScriptWorker.targets` 复制 worker 的完整输出到 `ScriptWorker/`，发布也保留该目录，包含 exe、deps/runtimeconfig、Contracts、托管桥、native DLL 与许可证。文件不得缺失；故障注入用 `WorkerFixture` 只进入测试输出，不进入发布包。所有托管构建继续使用仓库内 `--artifacts-path`，不改同级 ExcelIO 源码。

原生构建现在按 C/构建脚本时间戳自动重建，不再依赖可能缺失的 `OS` 环境变量；PowerShell 子进程恢复自身内置模块搜索路径，CMake 子进程剥离 dotnet 的 MSBuild 路径覆盖并显式传入 Windows x64 平台变量。固定上游版本、压缩包 SHA-256 和 MIT 许可证保持不变。

## 验收记录

新增 `M4ProtocolTests`、`M4WorkerTests`、`M4WorkerFaultTests` 及真实 WPF `MainWindow.WorkerSmoke`。覆盖实际子进程、类型/区域、保守重算、日志、无限循环、异常 toString、加载失败、源/输入/授权/会话失效、强杀后编辑保存、协议串话/重复/半帧/超长/非法值、无响应看门狗及引擎外内存拒绝。

2026-10-27 验证：Release solution 构建成功，0 警告/0 错误；Core 全套 **466 项通过，0 跳过**（原有 396 项 + 本轮 70 项）；native smoke PASS。重新编译固定版本原生引擎成功，上游 C 编译警告保留；不把随后增量托管构建的零警告描述为上游原生源码无警告。

发布树 `artifacts/publish/20261027-090739-d990126a` 的完整 WPF 验收六个结果文件均为新 PASS，其中实际停止按钮返回 23 ms（单次样本，非 p95）。旧版接口在 UI 线程同步调用曾使 IPC await 捕获上下文而阻塞，已补齐 ConfigureAwait(false)，增加单线程上下文回归，完整 WPF 重跑通过；首次失败证据保留在 `artifacts/smoke/20261027-090433-54b0408d`。

该发布树在更新默认安装目录时遇到文件访问拒绝，未作为最终 ZIP。随后重新运行完整发布流程成功。ExcelIO tracked diff 为空；其独立工作树已有的其他状态不属于本轮修改。

最终包：`artifacts/releases/QuickExcel-0.5.0-preview.1-win-x64-20261027-091025-e7dd6c3e.zip`，2,254,879 字节。

SHA-256：`6ABD20335EE963F25B01912FCFDBA4C015E43044EEB378AA7957C000B2E99D8D`。

最终发布轮 466 项测试通过、native smoke PASS、完整 WPF 六个结果文件 PASS；实际停止按钮单次返回 19 ms。ZIP 校验 sidecar 后在新目录解压、独立运行完整 WPF smoke，退出码 0，六个结果全部新 PASS；默认 `artifacts/QuickExcel` 的应用、Core、Infrastructure、Contracts、托管桥、worker 与 native DLL 哈希均与 ZIP 一致。

详细证据：`artifacts/smoke/20261027-091025-e7dd6c3e` 与 `artifacts/validation/release-0.5-preview/archive-validation.json`。本段在 ZIP 复验后补写，不改写已验证 ZIP；包内 BUILD-METADATA 记录真实源码 HEAD、dirty 状态与验证结果。本轮修改尚未提交，源码快照由 `record-baseline.ps1` 单独记录。

下一切片：FMT-008 包内模块与 actions 能力元数据 → JS-009 动作 snapshot/patch/CAS/Undo → worker 调度合批与 OS 访问策略原型。上述任务退出前，M4 保持进行中。
