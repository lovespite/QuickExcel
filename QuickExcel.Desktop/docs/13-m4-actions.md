# M4 actions 操作脚本

> 后续更新：有界 UDF 合批已接入；IPC 当前协议 4，消息上限统一为 16 MiB。见 [合批与容量](14-m4-batching.md)及 [OS 原型 ADR](15-m4-os-isolation-adr.md)。下文协议 3 和 583 项结果为 actions 切片历史证据。

2026-10-28，基于 `abf5883`（M4 stage2）的下一切片。FMT-008 的动作能力元数据与绑定、JS-009 的工作簿快照、patch 二次校验、原子提交及一次撤销已实现。M4 仍在进行中：计算计划合批、严格运行策略和 OS 文件/网络访问隔离尚未完成。只运行可信源码。

## 执行与事务

1. 打开或保存只读写源码和绑定，不运行脚本、不恢复授权。用户显式“应用脚本并重算”后发现动作，动作体仍未运行；选择并点击“运行操作”才调用动作。
2. `DocumentSession.BeginAction` 在锁内捕获不可变工作簿、修订、会话/脚本/计算代次以及唯一 operationId。发现动作、运行动作与计算互斥；新请求取消旧请求。
3. `IScriptActionExecutor` 由 `ScriptWorkerActionExecutor` 实现，每次发现或运行均创建独立 worker 和新模块实例。按 actions 绑定选择入口，复用包内逻辑 ID loader。动作上下文忽略函数注册，公式上下文忽略动作注册；模块顶层初始化仍会运行，动作体不会在计算或发现过程中调用。
4. worker 仅接收快照副本，没有文档引用。动作只构造建议 patch，不能直接修改宿主。同步结束后，worker 校验一次，宿主协议层校验一次，Core 命令准备阶段再验证全部编辑。
5. 宿主在会话锁内检查授权、当前会话/修订/代次、operationId、动作名称和取消状态；验证目标稳定 SheetId、坐标、类型、公式语法、重复写入、容量和撤销预算。全部通过后提交一条 `SetCellsCommand` 事务，Revision 仅增加一次，产生一次文档变更通知及一个撤销条目。
6. 源码/绑定变化、编辑、撤销/重做、停用、新计算/动作、关闭或取消使旧结果失效。失败、非法 patch、过期结果、超时和 worker 退出均不提交部分数据或历史。不自动重试；成功、无变更或失败的请求均不能复用。冲突后须明确重新运行以捕获新快照。

一次撤销恢复全部修改；重做恢复记录的内容，不再次调用动作，即使已经停用脚本也可执行。无变更 patch 不递增修订、不占历史。动作保留目标现有/继承样式，公式缓存按正常编辑规则失效；动作提交本身不启动重算，桌面标记等待重算。Undo/Redo 的常规重算可能调用已授权 UDF，但不会重跑动作。

## API 与示例

可运行示例：[3×4 表行列转置](../../examples/actions/transpose-3x4/README.md)，包含 JavaScript 源码和 `.xlsq` 生成器，演示原地转置、清空旧区域及一次撤销。

模块源码（文件级 scriptApiVersion 2）：

```javascript
quickExcel.actions.register({name: '整理商品', description: '去空格并写入公式'}, ({snapshot, patch}) => {
  const id = snapshot.sheets[0].id; // 稳定 SheetId，名称仅供显示
  const rows = snapshot.getRange(id, 'A1:A2');
  patch.setValues(id, 'B1:B2', rows.map(row => row.map(value =>
    typeof value === 'string' ? value.trim() : value)));
  patch.setFormula(id, 'C1', '=1+2');
  console.log('整理完成');
});
```

`register('名称', callback)` 是描述为空的简写。名称区分大小写，支持中文，非空、最多 128 字符、无控制字符；说明最多 1000 字符，一轮最多 256 个不重复动作。只在模块初始化阶段注册。回调必须同步完成且返回 undefined/null；Promise、异步待执行任务或其他返回值均拒绝。

| 对象/API | 契约 |
| --- | --- |
| `snapshot.baseRevision` / `dateSystem` | 捕获修订与 `Windows1900` / `Mac1904`；日期保持数字序列，不猜成 JS Date |
| `snapshot.sheets` | 冻结的 `{id, name}` 数组，按工作簿顺序；不包含样式/图片或可变宿主对象 |
| `snapshot.getCell(id, 'A1')` | 冻结的 `{value, formula, state, hasValue}`；无缓存公式为 value=null、hasValue=false，空白位置为 null、hasValue=true；公式缓存可能 Dirty/Stale/Blocked，读取不触发重算 |
| `snapshot.getRange(id, 'A1:B2')` | 冻结的二维值数组，空洞为 null；需判断公式缓存有效性时用 getCell |
| `patch.setValues(id, range, rows)` / `setValue(id, cell, value)` | 数字必须有限，文本保留空文本/数字外观/前导 =，支持布尔和 null（清空内容）；显式错误为 `{kind:'error', code:'DivisionByZero'}`，代码使用 Core CellErrorCode 枚举名 |
| `patch.setFormulas(id, range, rows)` / `setFormula(id, cell, formula)` | 以 = 开始，2–8192 字符，宿主公式解析器验证语法；不在动作内求值，不接受动态数组溢出 |

范围仅支持不带表名的 A1 单格或矩形，不支持命名范围、整行/整列、反向范围和外部引用。二维数组形状必须与范围一致。patch 内包括跨调用的任何重叠写入都拒绝；即使动作捕获了无效写入的异常，最终也不能提交之前的部分编辑。快照在整轮内不随 patch 改变，读取刚写入的目标仍得到旧值。

当前限制：最多 64 表、快照 100000 个存储格、累计读取 100000 个位置、patch 20000 个目标、单元格文本 32767 字符；坐标范围 100000 行 × 1024 列，仍受全簿 1000000 格和会话历史预算限制。整个 IPC 帧及 native 桥消息最多 16 MiB（合批切片由 4 MiB 放宽），复杂快照可能先触及帧预算。日志最多 100 条、每条 4000 UTF-8 字节。worker 整轮 5 秒、native 每次加载/调用 1 秒、QuickJS 堆 64 MiB、Job 提交内存 256 MiB；取消终止进程。

本切片只支持写值与公式，不支持结构、样式、图片、合并、任意宿主调用、异步任务或文件/网络 API。写入含合并区域或单元格内图片的工作表整批拒绝；其他工作表仍可操作。原始内容、样式和图片不会被脚本隐式移除。

## 包格式与版本

XLSQ 保持 formatVersion/minReaderVersion 2。actions 要求 scriptApiVersion 2，manifest.requiredFeatures 在既有三个能力上增加 `script-actions`。`state/document.json` 的 scriptBindings 使用 `{module:'main', scope:'workbook', kind:'actions'}`；同一模块可同时有 formula-functions 和 actions 两项绑定。全部入口必须有绑定，绑定必须指向入口；重复项、未知 kind/scope、能力与绑定不一致、API v1 actions 均拒绝。

```csharp
var modules = new ScriptModuleSet(
    new Dictionary<string, ScriptModule> { ["main"] = new("scripts/main.js", source) },
    ["main"], scriptApiVersion: 2, bindings: [new("main", "actions")]);
document.Session.Execute(new UpdateScriptModulesCommand(modules));
document.Save(path);
```

actions 包的集合摘要包含有序入口、模块 ID/路径/源码摘要和按 module/kind 排序的绑定；原 API v1 与纯函数 API v2 摘要算法不变。保存保留同模块不同绑定各自的可选元数据。绑定变化与源码变化都会撤销授权。旧读取器通过必需能力拒绝新包，不可默默忽略动作。桌面暂不提供创建绑定或多文件编辑器；通过 Core 生成包后打开，编辑器仍编辑首入口。

IPC 升为 **协议 3**，新增 loadActions / runAction 及类型化快照/patch。动作消息的 CalculationGeneration 为 null，公式消息为非 null；宿主仍保留内部代次用于废弃旧请求。完整 envelope 身份和单次 operationId 双重匹配；跨模式消息或旧协议 1/2 配对均拒绝。Contract 仅依赖 BCL，worker 继续不引用 Core；桌面不使用进程内后备。

## 验证与进度

`M4ActionTests` 覆盖真实 native/worker 的不可变快照、类型、缓存、跨表事务、模块入口隔离、一次撤销/重做、重放、无变更、源码/绑定/授权/代次/修订失效、取消、异常/超时/Promise、容量/样式/不支持目标及伪造回包。`TypedStorageTests` 覆盖动作能力与绑定往返、摘要篡改、元数据保存；原 API v1、纯函数模块与存储回归继续执行。

`MainWindow.ActionSmoke` 使用发布目录的实际 worker，验证动作包打开、显式发现、真实运行按钮、日志、停用后一次撤销/重做、保存重开和停止按钮。验收会生成 `worker-actions.xlsq` 示例与 `actions-smoke-result.txt`；发布脚本要求该结果为新的 PASS，并将协议 3 与本文打包。

2026-10-28 最终验收：

| 检查 | 结果与证据 |
| --- | --- |
| Release solution 构建 | 0 警告、0 错误；`artifacts/m4-actions-build.log` |
| 全量 Core 测试 | **583 通过、0 失败、0 跳过**，原有 514 + 新增 69；`artifacts/m4-actions-validation.log` |
| 原生运行时 smoke | PASS，10,000 次调用 147 ms（单次 smoke 观测，不是性能分位基准）；日志同上 |
| 发布版 WPF | 七个结果文件均为新 PASS，包含真实动作运行/停止及一次撤销/重做；`artifacts/smoke/20261028-015253-60ca024d` |
| 候选程序 | `artifacts/m4-actions-candidate/QuickExcel.exe`；完整 ScriptWorker、原生库和许可证，不含 WorkerFixture |
| 示例动作包 | `artifacts/smoke/20261028-015253-60ca024d/worker-actions.xlsq`；打开后显式应用并选择“整理商品” |

执行了 `dotnet build .\QuickExcel.slnx -c Release --artifacts-path .\artifacts\build -v minimal` 和 `scripts/build-quickexcel.ps1 -RunSmoke -OutputDirectory artifacts\m4-actions-candidate`；发布脚本包含全量测试、native run、publish 和 WPF smoke，所有托管输出均位于仓库 artifacts。验收过程中修正了一处新增测试夹具漏重算 state 摘要的问题，最终全套通过。

验收记录在发布验证后补写；候选仍使用 0.5.0-preview.1，BUILD-METADATA 记录基线 abf5883 与 dirty 源码状态。没有修改原生来源版本、许可证或同级 ExcelIO 源码；没有生成发布 ZIP 或提交。后续 M4 继续计算计划合批与 OS 访问策略原型，完整 Beta 仍须后续门禁。
