# M4 包内模块加载与执行

> 后续更新：actions 能力与绑定、快照/patch/一次撤销已实现，IPC 当前协议 3，见 [M4 actions 记录](13-m4-actions.md)。下文保留模块切片的协议 2 与 514 项测试证据；其中 actions 待办以新记录为准。

2026-10-28，基于 `8d193a1` 的下一切片。FMT-008 的包内模块加载部分已接入，actions 能力元数据与 JS-009 动作事务仍未实现；M4 保持进行中。worker 生命周期与资源边界见 [11-m4-worker.md](11-m4-worker.md)。

## 文件版本与兼容

- XLSQ 保持 `formatVersion: 2`、`minReaderVersion: 2`。`scriptApiVersion: 1` 保留旧单源码函数包装语义；现存 API v1 多模块包仍可保存，但执行必须显式迁移到 API v2，不按源码或模块数量猜执行模式。
- `scriptApiVersion: 2` 明确表示 ES 模块加载，单模块包也可使用。`requiredFeatures` 仍为 `typed-workbook`、`stable-sheet-ids`、`script-modules`。旧读取器会拒绝 API v2；本次不接受 `script-actions` 或 `kind: actions`。
- 函数描述符的 `apiVersion` 独立控制参数桥。在 API v2 模块里仍可注册 v1 函数或 v2 `JS.*` 描述符；文件级 v2 不改变 Blank/区域参数语义。
- 保存、打开、迁移和导出不执行源码，不持久化授权。API v2 的源码集合摘要包含执行版本、入口顺序、模块 ID、路径和源码摘要；API v1 的既有摘要算法保持兼容。包中每个源码条目的 SHA-256 继续校验，未被导入的源码也校验。

## 加载契约

`import`/`export ... from` 只接受 manifest `scripts.modules` 中精确、区分大小写的逻辑 ID。例如 `import { discount } from 'price'`；`scripts/lib/price.js` 是存储路径，不能用作导入名。相对路径、绝对路径、URL、未声明 ID、文件/网络 loader、外部包和 import attributes 均不支持。

使用 QuickJS-NG 原生 ES 模块链接器，支持命名/default/namespace 导入、再导出、活绑定及同步循环依赖。宿主按 `entryModules` 的顺序生成内部根模块，依赖按 ES 模块规则先求值，共享依赖每轮只初始化一次；未从入口可达的模块不执行。循环中的暂时性死区错误仍是加载失败，不人为改写源码规避。

所有模块共享本轮只读的 `quickExcel.functions` 与 `console` API，顶层变量遵循模块作用域。只有加载期间可以注册函数，所有入口成功后冻结注册表。任何语法、链接、初始化、重复注册或资源错误都销毁本轮 runtime；不发布部分注册或计算缓存。SUM/IF 等内置名称冲突仍由计算器拒绝。

初始化必须同步完成。不驱动 Promise 任务循环；动态 `import()`、需要继续执行的顶层 `await`、初始化期间留下的异步任务均使加载失败。UDF 仍仅支持同步标量返回，不能通过异步任务延迟注册或触发模块执行。

## 传输、授权与容量

- IPC 升为协议 **2**，新增 `loadModules`；协议 1 的旧 worker/宿主配对明确拒绝。API v1 源码仍用 `load`，API v2 传完整模块集合与入口。整个集合摘要在 worker 端重新验证，绑定 ID、路径、源码和入口顺序，字典枚举顺序不影响摘要。
- 最多 64 个模块和 64 个不重复入口；总源码最多 1 MiB UTF-8，载荷仍受 4 MiB 帧限制。逻辑 ID 最多 128 个 ASCII 字母/数字/`_`/`-`/`.` 字符，排除 `.`/`..`；包内路径最多 512 字符，禁止控制字符、穿越、路径别名和 URI 字符。
- native 加载全图共享一次 1 秒期限，宿主全轮 5 秒期限与取消杀进程保持。每轮新 worker、新模块实例、完整重算，不复用上一轮模块状态或 UDF 结果。
- `UpdateScriptModulesCommand` 更新集合，统一递增源码修订；入口/ID/路径/API 变化也递增绑定修订。任何接受的源码或绑定变化撤销授权、取消旧计算并拒绝迟到结果。单元格撤销不会回退模块源码或恢复授权。
- 模块源码来自已经校验的内存快照；native loader 从该集合取源码，从不打开路径。Windows Job 仍仅约束资源与生命周期，OS 文件/网络隔离未完成，本版继续限定可信源码。

## 示例与桌面边界

模块 `price`（存储在 `scripts/lib/price.js`）：

```javascript
export const discount = (price, rate) => price * (1 - rate);
```

入口模块 `main`（`scripts/main.js`）：

```javascript
import { discount } from 'price';
quickExcel.functions.register('DISCOUNT', discount);
// 表格公式：=DISCOUNT(B2,D2)
```

通过 Core 构造/更新包并用 `TypedWorkbookStorage` 保存，会生成完整 manifest 与摘要，无需手改 ZIP：

```csharp
var modules = new ScriptModuleSet(new Dictionary<string, ScriptModule>
{
    ["main"] = new("scripts/main.js", mainSource),
    ["price"] = new("scripts/lib/price.js", priceSource)
}, ["main"], scriptApiVersion: 2);
document.Session.Execute(new UpdateScriptModulesCommand(modules));
document.Save(path);
```

桌面打开后仍需点击“应用脚本并重算”；编辑器显示首入口源码，日志注明模块数量与入口。编辑首入口、停用、停止、保存重开均接入现有流程。模块树、多文件编辑器及可视化 API 迁移入口尚未提供，属于后续桌面工作。本切片的真实 WPF smoke 会生成 `worker-modules.xlsq` 示例。

## 验证

新增 `M4ModuleTests` 覆盖真实 worker 与原生桥的模块语义、全轮失败保护、取消、授权/绑定失效、往返、容量边界、摘要及恶意协议消息。`TypedStorageTests` 增补 API 篡改、辅助源码摘要和未实现 actions 拒绝；保留 API v1 多模块只保存的回归。`MainWindow.WorkerSmoke` 增补实际桌面组合根的打开/应用/编辑/保存/重开场景。

2026-10-28 验证结果：

| 检查 | 结果与证据 |
| --- | --- |
| Release solution 构建 | 0 警告、0 错误；`artifacts/m4-modules-build-final.log` |
| 全套 Core 测试 | **514 通过、0 失败、0 跳过**，原有 466 + 新增 48；`artifacts/m4-modules-validation.log` |
| 原生运行时 smoke | PASS；由完整发布验证脚本运行，日志同上 |
| 发布版 WPF | 六个结果文件均为新 PASS，包含实际模块包打开/应用/编辑/保存/重开；`artifacts/smoke/20261028-012051-dcc9cbe9` |
| 本地候选程序 | `artifacts/m4-modules-candidate/QuickExcel.exe`；完整 worker 与许可证齐备，不含 WorkerFixture |
| 示例工作簿 | `artifacts/smoke/20261028-012051-dcc9cbe9/worker-modules.xlsq` |

执行了 `dotnet build .\QuickExcel.slnx -c Release --artifacts-path .\artifacts\build -v minimal`，以及 `scripts/build-quickexcel.ps1 -RunSmoke -OutputDirectory artifacts\m4-modules-candidate`；后者包含全量 `dotnet test`、native `dotnet run`、publish 和 WPF smoke，全部托管输出指定仓库内 artifacts 路径。最初 native 构建因受限环境下 Visual Studio FileTracker 系统路径访问失败，获得构建权限后重试成功；原生来源版本、归档 SHA-256 和许可证未变。

本轮没有生成新的发布 ZIP、提交或 Beta 放行；候选程序仍使用现有预览版版本号，BUILD-METADATA 记录基线 `8d193a1` 与 dirty 源码状态。上述验收记录在验证后补写，不据此声称旧 ZIP 包含模块能力。

## 后续

FMT-008 的 actions 必需能力与绑定 → JS-009 快照/patch/CAS/单次撤销 → worker 合批调度及 OS 访问策略原型。包内模块执行完成不等于完整 M4 或 Beta 完成。
