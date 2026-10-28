# M4 合批与消息容量

2026-10-28，基于 `27f8264`（M4 stage3）。本切片接通计算计划到 worker 的有界 UDF 合批，将 IPC 与托管/原生桥消息上限从 4 MiB 提高到 16 MiB，并新增 Windows 隔离开发探测。后续同日切片已接入实际 AppContainer worker 并完成回环/DNS 的本机 WFP 取证；严格执行策略、部署矩阵及 M6 最终 OS 门禁仍未完成。OS 结果和部署选择见 [隔离 ADR](15-m4-os-isolation-adr.md)。

## 调度与语义

Core 新增可选 `IScriptBatchRuntime` / `IScriptInvocationBatch` 端口，不引用传输类型。Infrastructure 的 worker 适配器在加入批次时编码、测量并限制容量，加入批次不执行用户代码；每批仅执行一次，不能重复发送。原生进程内适配器继续按单次调用执行，作为回归参考。

计算器保持原先按工作表、行、列优先的拓扑出队顺序。只合并连续、互不依赖、整格为 UDF 且参数没有嵌套 UDF 的调用。遇到依赖尚未返回的结果、普通公式、嵌套 UDF 或 IF/IFERROR 时先提交当前批次到轮内结果集合，再继续原求值路径。这里的轮内结果尚未写入文档；只有整轮成功后才经会话 stamp 门禁提交。

不把拓扑层全部重排为宽度优先，不预先执行未选中的条件分支，不合并相同参数，不复用 UDF 或下游的旧缓存。模块实例和 worker 仍每轮重建，闭包计数器的同轮调用顺序保持。带错误的输入不调用 UDF；普通函数错误仍按格返回，传输故障、超时、取消和不完整回包丢弃整轮。

当前未合批的表达式仍可正常计算，只是保留单调用 IPC。例如 `=F(A1)+G(A2)`、`=IF(B1,F(A1),G(A2))` 和 `=F(G(A1))` 不做表达式续体调度。跨轮驻留进程及跨公式的更广泛合批仍需单独测量和设计。

## 16 MiB 容量与协议 4

4 MiB 是 QuickExcel 自己设置的防御性预算，不是 Windows 管道或 QuickJS 的固定限制。本轮按用户要求放宽为 **16 MiB UTF-8 JSON**，统一调整：

- `WorkerProtocol.MaxFrameBytes`：IPC 读端先检查长度再分配；写端继续通过有上限的缓冲序列化。
- `QuickJsSession.MaxMessageBytes`：模块及调用输入按 UTF-8 字节检查，替换旧调用路径的字符串字符数检查。
- 原生 `QX_RESULT_LIMIT`：返回 JSON 上限提高到 16 MiB；引擎来源版本、源码压缩包 SHA-256 和许可证未改变。

源码总量 1 MiB、单格文本 32767 字符、范围 1000000 位置、动作 patch 20000 格、QuickJS 堆 64 MiB、Job 提交内存 256 MiB、单次 1 秒与整轮 5 秒保持原值。16 MiB 只放宽消息大小，不保证所有接近该上限的复杂负载都能在内存/时间预算内完成。

协议调用批次仍最多 256 项。计算器的批次还为回复预留最坏大小：每个 UTF-16 单元最多转义为 6 个 JSON 字节，加上结果元数据及 envelope；因此当前计算批次最多 **85 次调用**，输入超过字节预算则提前分批。整个范围参数不会被拆分或改变形状；单次调用本身超限即拒绝本轮，不通过重试执行两次脚本。

IPC 升为 **协议 4**。`InvocationResult.valueTooLong` 明确表示原始返回文本超过单格限制，必须与 `value` / `error` 互斥。worker 不把超长字符串装入合批回包；计算器仍将该格标为 Value 错误。这个标记防止无效单格返回挤爆其他合法结果的回包。协议 1/2/3 配对均拒绝，更新必须同时替换宿主与整个 ScriptWorker 目录。XLSQ 文件格式和 scriptApiVersion 没有变化。

## 策略验证与未完成项

新增真实 worker 反例覆盖 Date.now、Math.random、eval、Function 及包内可变闭包。当前 `conservative-worker-v1` 仍允许这些语言能力，因此本轮不开放确定性声明或结果复用；`profile: strict` 继续明确拒绝。

JS-010 下一步需要单独冻结严格策略的语义：时间/随机来源如何禁用或显式注入，动态构造入口如何约束，共享状态如何隔离，以及策略身份如何绑定授权和输入。仅冻结注册表或证明重复样例相等不足以关闭任务。

SEC-001 / JS-014 的实际 worker 已接入 AppContainer、挂起启动、最小句柄、运行库副本 ACL/临时 profile/环境白名单；回环与 DNS 已取得本机拒绝事件。部署矩阵、宿主崩溃残留回收和更多系统调用测试见隔离 ADR。M6 仍需实际发布包、故障攻击与干净机器门禁。

## 验证

新增 `M4BatchTests`、`M4MessageBudgetTests`、`M4PolicyTests`，扩展协议及故障夹具。覆盖跨表依赖与优先顺序、懒分支、嵌套调用、区域/描述符、相同参数的状态性调用、字节分批、最坏文本回包、超长单格、坏回包、超时无部分提交及旧协议拒绝。

容量实测包括 5 MiB 原生输入/返回和真实 worker 输入、超过旧 4 MiB 的动作快照及一次提交/撤销、中文 JSON 转义后超过 4 MiB 的 IPC 帧。600 个独立闭包计数调用使用 8 个批次，原路径为 600 个；这只是 IPC 次数证据，不是 p95 性能结论。发布版 WPF 另验证 60 次状态性调用合为 1 批，保存 `worker-batches.xlsq`。

2026-10-28 最终验收：

| 检查 | 结果与证据 |
| --- | --- |
| Release solution 构建 | 0 警告、0 错误；`artifacts/m4-batching-build.log` |
| 全量 Core 测试 | **605 通过、0 失败、0 跳过**（原有 583 + 本轮 22）；`artifacts/m4-batching-tests.log`，发布流程再次通过 |
| 原生运行时 smoke | PASS；10,000 次调用 147 ms，仅单次 smoke 观测；`artifacts/m4-batching-validation.log` |
| 发布版 WPF | 七份新 PASS，包含 60 次闭包 UDF 合为 1 批；`artifacts/smoke/20261028-025316-6ff28a2c` |
| OS 开发探测 | AppContainer/Job/删减令牌及句柄对照完成，清理成功；具体通过项与待确认项见 ADR，不是发布隔离放行 |
| 候选程序 | `artifacts/m4-batching-candidate/QuickExcel.exe`；完整 worker/native/许可证，不含 WorkerFixture 或 SandboxProbe |
| 合批示例 | `artifacts/smoke/20261028-025316-6ff28a2c/worker-batches.xlsq` |

执行了标准 Release solution 构建、全量 Core 测试，以及 `scripts/build-quickexcel.ps1 -RunSmoke -OutputDirectory artifacts\m4-batching-candidate`；后者包含全量测试、native run、publish 与真实 WPF 验收，托管输出始终指定仓库内 artifacts。原生桥预算变化需要重新编译，首次因受限环境下 Visual Studio FileTracker 的系统路径访问拒绝而失败；正常构建权限下完成重编译及定向测试。没有改动固定上游版本、归档 SHA-256、许可证或同级 ExcelIO 源码；检查时同级库 tracked diff 为空。

候选仍使用 `0.5.0-preview.1`，BUILD-METADATA 记录源码基线 `27f8264` 和 dirty 状态、协议 4、16 MiB/85 次计算批次及 `osAccessIsolationValidated: false`。本节验收在候选生成后补写并同步到候选文档，未改变已验证的二进制。没有生成发布 ZIP 或提交。
