# M4：宿主崩溃资源回收与部署验证

2026-10-28。宿主异常退出后的 profile/payload 回收已实现，完整回归 625 项通过。本机部署矩阵完成五项通过、一项失败的验证；M4 和 M6 OS 安全门禁仍未关闭。

## 所有权与回收协议

生产代码：`WorkerSandboxLease`、`WorkerSandbox`、`WorkerJob`、`WorkerProcess`（Infrastructure/Scripting）。

1. 当前用户 LocalAppData 下的 `QuickExcel.WorkerSandboxes.v1` 是版本化所有权根目录。根和每轮记录目录使用受保护 DACL，只允许当前用户；读取时核验 owner、DACL 和整条路径无 reparse point。
2. 持有根目录 `.gate` 文件的独占句柄，先回收孤立记录，再创建 `QuickExcel.Worker.<32 位小写 GUID>` 目录及独占 `lease` 文件。目录本身就是记录，先于 profile、payload 和 Job 创建，不依赖可能截断的 JSON、PID 或时间戳。
3. profile 名、`payload` 子目录和 `Global\QuickExcel.Worker.<GUID>` Job 名都从记录名推导。worker 只能读取/执行 payload，不能读写记录目录、租约或根目录。Job DACL 只授权当前用户，拒绝复用已有同名对象。
4. 先配置 Job，再以 `PROC_THREAD_ATTRIBUTE_JOB_LIST` 创建挂起进程；核验零能力 AppContainer token 和 Job 成员身份后才恢复线程。租约、根目录锁和 Job 句柄均不继承给 worker；显式继承仍只有三个 IPC 句柄。
5. 正常释放先终止并等待 worker，然后清理 payload 和 profile，最后关闭租约并删除记录。宿主死亡会释放租约和 Job；桌面启动及每轮 worker 启动前都执行回收，**不会执行脚本、恢复授权或重放 action**。
6. 回收器在根锁内尝试独占租约；被占用就跳过，避免删除其他实例的资源。租约可用后还须确认对应 Job 不存在或活动进程数为零；仍有进程最多等 1 秒，失败保留记录。使用全局 Job 名覆盖同一用户的不同登录会话，但多会话实机矩阵尚未完成。
7. payload 必须是平铺文件集合，拒绝链接、子目录和未知记录内容，不递归删除 payload。只读文件属性在私有副本上清除后删除。进程退出信号可能先于文件映像释放，访问拒绝/占用最多重试 1 秒；失败报告并保留记录供下次回收，其他可回收条目仍继续处理。清理失败释放租约，避免本进程一直阻止重试。
8. `DeleteAppContainerProfile` 仅按推导的名称操作当前用户 profile。不存在的 profile 删除成功，因此目录已创建但 profile 尚未创建、复制中断、已部分删除等阶段都可重试。最后记录目录删除前崩溃也可恢复。

目录和租约协议针对进程崩溃；不宣称突然断电的跨文件系统事务持久性。Windows 正在关闭句柄时，租约仍可能暂时被占用，这一轮安全跳过，后续启动重试。根目录锁等待上限为 5 秒；回收并不是整个应用启动的统一 1 秒预算。损坏或不安全记录会报告失败并保留，不自动放宽权限。

旧实现的临时 payload 和 profile 使用互不关联的随机名字，没有所有权记录，**不按模糊前缀扫描或删除这些历史目录/profile**。本协议覆盖引入版本化记录后的创建与清理生命周期；旧遗留须依据单独证据处理。临时磁盘配额仍待实现。

Windows API 依据：[在创建进程时加入 Job，避免挂起进程孤立窗口](https://devblogs.microsoft.com/oldnewthing/20230209-00/?p=107812)、[Job 对象生命周期](https://learn.microsoft.com/en-us/windows/win32/procthread/job-objects)、[DeleteAppContainerProfile 的幂等与重试语义](https://learn.microsoft.com/en-us/windows/win32/api/userenv/nf-userenv-deleteappcontainerprofile)。

## 故障测试

`QuickExcel.Core.Test/M4RecoveryTests.cs` 新增 15 项案例，`SandboxHostFixture` 仅用于测试，和原 WorkerFixture/SandboxProbe 一样禁止进入发布包。

- 强杀独立宿主，覆盖只有记录、profile 创建后、复制中途、完整 payload、真实 worker 已创建、故障 worker 持续运行六个阶段。父测试只杀宿主，不杀进程树，验证 Job 终止忽略 IPC EOF 的 worker。
- 活动租约不被其他宿主回收；保留额外 Job 句柄制造仍运行的 worker，验证回收明确拒绝，payload 全部文件不变。
- 文件占用导致清理失败，记录保留；六个独立回收进程竞争时仅一个回收成功；重复回收幂等。
- Dispose 清理失败后释放租约，下一次回收成功；下一轮 worker 创建会自动回收旧记录。
- 创建租约前中断、来源不存在、预取消、非规范目录名，以及根/记录/payload junction 拒绝和外部 canary 保留。
- 测试宿主启动失败写 stderr 并退出；捕获异常并设置 `SEM_NOGPFAULTERRORBOX`，不弹 Windows 错误报告窗口。首轮受限环境曾触发该弹框，已修正并增加回归；当时残留的两组测试资源及只读部署副本均已清理。

现有 `M4IsolationTests` 继续覆盖真实 token、唯一身份、文件/句柄/进程限制、环境、取消及强杀清理，并新增 worker 无法读写所有权记录的断言。真实 AppContainer 测试必须在正常 Windows 用户环境运行；受限 runner 的 profile/LocalAppData 访问失败不等于通过，不改变生产隔离来适配 runner。

```powershell
dotnet test .\QuickExcel.Core.Test\QuickExcel.Core.Test.csproj -c Release --artifacts-path .\artifacts\build -v minimal --filter 'FullyQualifiedName~M4RecoveryTests|FullyQualifiedName~M4IsolationTests'
.\scripts\build-quickexcel.ps1 -RunSmoke -CreateArchive -OutputDirectory artifacts\validation\m4-recovery\final-package
```

最终发布轮：625 通过、0 失败、0 跳过；native smoke PASS；WPF 总结果及 storage/close/calculation/date-display/worker/actions 六个子结果均 PASS。Release solution 和开发探针构建均为 0 警告、0 错误。最后仅加强活动 Job 测试的断言，不改变已验证产品二进制。

证据：

- `artifacts/validation/m4-recovery/m4-recovery-final.trx`：15 个回收案例通过。
- `artifacts/publish/20261028-041335-edd6ac02/BUILD-METADATA.json`：完整发布轮测试/native/WPF 为 passed，OS 最终门禁仍为 false。
- `artifacts/smoke/20261028-041335-edd6ac02`：最终包 WPF 证据。
- 验证 ZIP：`artifacts/releases/QuickExcel-0.5.0-preview.1-win-x64-20261028-041335-edd6ac02.zip`。
- SHA-256：`EE632AC7CF99A371718A23006C8C35A6C13F4FACF254B7AD5FC2523D08012AAA`。

本记录在验证后补写，不修改已经记录哈希的 ZIP；包内文档是发布时快照，以仓库本记录为最新结果。本轮工作树包含此前未提交的 M4 工作，没有修改兄弟 ExcelIO 源码。

## 本机部署矩阵

脚本：`scripts/test-deployment-matrix.ps1`。只接受仓库 artifacts 下的发布树，在唯一测试目录复制；修改 ACL/属性仅限测试副本。结束前核验文件 SHA-256、目录/文件 DACL 和属性未被应用改动，恢复测试 DACL 后删除副本。所有成功案例都运行完整 WPF smoke（包括公式、模块、actions、停止和 Undo/Redo），并通过卸载删除验证。

```powershell
.\scripts\test-deployment-matrix.ps1 -PackageDirectory artifacts\validation\m4-recovery\final-package
# 单独补测一个场景：
.\scripts\test-deployment-matrix.ps1 -PackageDirectory artifacts\validation\m4-recovery\final-package -CasesToRun long-path-dotnet
```

环境：Windows 11 内核 `10.0.22621.0`，.NET 10.0.12 x64，当前账户非提权 token，`LongPathsEnabled=0`。不是独立标准用户账户、管理员 token 或干净机器验收。证据为 `artifacts/deployment-matrix/20261028-041459-13d4c7b1/matrix.json`，以及长路径托管启动对照 `20261028-041750-081bdd1a/matrix.json`。

| 场景 | 结果 | 范围 |
| --- | --- | --- |
| 中文和空格便携目录 | 通过 | 完整 WPF smoke；安装内容与 ACL 不变；删除成功 |
| 宿主只读/执行 ACL | 通过 | 写入探针确认拒绝；不依靠宿主安装目录写权限 |
| 只读文件属性 | 通过 | 副本保留只读属性；worker 私有 payload 可以正确清理 |
| 同版本替换及卸载 | 通过 | 删除安装树、重新复制同一发布包、再次运行完整 smoke、再次删除；不代表跨版本迁移 |
| 315 字符 EXE 路径直接启动 | **失败** | 短工作目录下 Process.Start/CreateProcess 返回 Win32 2，进程尚未启动；Shell 启动同样失败 |
| 330 字符安装路径，通过 `dotnet QuickExcel.dll` | 通过 | 短工作目录启动完整 WPF；真实 worker 从长路径读取源文件并复制到私有 payload；安装内容/ACL 不变，删除成功 |

长路径两种启动方式是独立结果，不把托管启动通过替代 EXE 启动失败。当前推荐使用短安装目录。未修改机器注册表或组策略；系统路径策略与程序声明的限制参见 [Windows 最大路径长度](https://learn.microsoft.com/en-us/windows/win32/fileio/maximum-file-path-limitation)。尚不宣称 Explorer/双击/仓库 cmd 启动器支持任意长路径。

首轮证据 `20261028-041018-95094f4a` 保留：只读 ACL 场景的数据提交成功，但 smoke 在 `_busy` 变 false 后、按钮异步操作更新日志前检查，造成日志断言竞态。已改为等待实际按钮的完整 Task；新包通过。测试脚本恢复 ACL 改用只写 DACL 的 .NET API，避免 PowerShell Set-Acl 额外请求审核权限；未提升账户或扩大生产 ACL。

## 后续验证边界

1. 使用同一 ZIP 和 SHA-256，在独立标准账户、提权管理员账户及多登录会话分别验证；记录账户 token 状态。不能把本机非提权管理员账户等同于独立标准账户。
2. 在真实 Program Files 安装树验证普通用户执行、拒绝写入、公式/actions 和崩溃回收；安装/卸载由授权安装流程执行。当前只读 ACL 测试不冒充实际 Program Files。
3. 使用两份不同版本的包进行更新/回滚，含旧宿主活动时的新宿主回收，不删除活跃租约；保留工作簿与授权语义。当前只测了同版本重新安装。
4. 干净机器缺失/正确安装 .NET 10 Desktop Runtime、EXE 长路径启动、杀毒软件锁文件、profile 删除失败、磁盘不足/额度，以及更多 OS 对象访问面继续验证。
5. 严格执行策略、最终 M6 QA-OS-01 和被攻陷 native worker 的威胁模型仍待完成；继续仅运行显式授权的可信脚本。
