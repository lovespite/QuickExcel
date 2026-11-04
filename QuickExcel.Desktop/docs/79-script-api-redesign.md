# 脚本 JavaScript API 重新设计（API v3）

2026-11-04。本文替换 [76](76-live-workbook-script-model.md) 中的对象 API 部分：旧 `$` / `ExcelScript` 子集整体废弃，不提供兼容层。文件 `scriptApiVersion` 目标为 3（本轮运行时切片仍写 2，存储升级随迁移切片完成），IPC 提升为 10；隔离、事务、限额和 worker 生命周期契约沿用 11/15/16/69/72/77 号文档，本文只改变脚本可见的对象模型、类型声明与宿主派发方式。

## 1. 现状与问题

当前脚本只能看到 5 个对象、14 个方法（Workbook / Worksheet / Range / RangeFormat / RangeFill），声明写在 `QuickExcel.Web/src/quickexcel.d.ts`，实现写在 `WorkbookModelBootstrap.cs` 的 C# 原始字符串里，宿主白名单写在 `ScriptModelAdapter.cs`。同一套接口在三个地方手工维护，任何新增成员都要同步修改，且无法覆盖工作表结构、表格、图表、条件格式等 Excel 文档模型。这正是本次重做的直接原因。

## 2. 已确认的决策

| 决策 | 取值 | 理由 |
| --- | --- | --- |
| 对象模型基线 | Office Scripts 的 `ExcelScript` 同步模型 | 与 JSC 同步 IPC、事务草稿、单轮执行天然一致；成员名与 Excel 用户认知一致 |
| 用户命名空间 | `QuickExcel`，成员名与 `ExcelScript` 完全一致 | 规避微软商标与官方实现暗示，同时保留可迁移性 |
| 首期类型范围 | 全量类型骨架，运行时按能力清单报未实现 | 类型层一次对齐，避免逐版本补声明；未实现能力必须显式失败 |
| 类型定义来源 | 官方 `office-scripts-docs.d.ts`，固定 commit 与 SHA-256 | 有权威、可持续更新的定义，不再手工维护 |
| API 版本 | `scriptApiVersion = 3`，IPC 10 | 契约整体变更，版本号必须可区分；正式版未发布，不存在外部文件负担 |

不做的事：不实现 Office.js 的 `Excel.run` / `context.sync` 异步模型；不暴露 DOM、网络、Node、动态 `import`、`eval`；不承诺与 Microsoft Office Scripts 运行时完全等价。

## 3. API 形态

### 3.1 入口

```ts
function main(workbook: QuickExcel.Workbook, host: QuickExcel.Host) {
    const sheet = workbook.getActiveWorksheet();
    sheet.getRange("A1:B2").setValues([[1, 2], [3, 4]]);
    host.files.writeText("//temp/report.txt", "done");
}
```

- `main(workbook, host)` 是唯一动作入口；每个声明 `actions` 绑定的入口模块的 `main` 就是一个动作，动作名取模块 ID。多动作必须拆成多个入口模块。新建模块时可选填单行描述（最多 200 字符），保存在 manifest 的 `scripts.modules[id].description`，仅显示在操作列表/提示中，不参与执行、缓存摘要或模块哈希；签名载荷覆盖描述，防止签名包被改写为误导性文案。
- 模块初始化阶段存在一个只用于注册的全局 `host`；它提供 `functions.register` 与 `events.on`，`files` 在该阶段访问会报 `InvalidReference`。运行阶段同一个 `host` 对象也会作为 `main` 的第二个参数传入。
- 不提供 `actions.register`、`QxEventNames` 或事件属性处理器。
- 控制台请求获得同样的 `workbook` 与 `host` 作用域。
- 事件处理器签名为 `(context: QuickExcel.EventContext) => void`，`context` 含 `workbook`、`host`、`event`。
- 删除 `$`、`quickExcel` 全局、`snapshot`、`patch`，不保留别名。旧示例与模板全部重写。

### 3.2 对象句柄

对象是 worker 内的不可变描述符，不持有宿主状态，每次方法调用都重新向宿主解析：

| 对象 | 句柄内容 |
| --- | --- |
| Workbook | 空（会话内唯一） |
| Worksheet | `sheetId` |
| Range / RangeAreas | `sheetId` + 一个或多个 A1 矩形 |
| Table / Chart / Shape / PivotTable 等 | `sheetId` + 名称或稳定 ID |

规则：

- 句柄可保存在模块变量中，作用域仅限当前授权工作簿会话；跨工作簿或失效会话使用报 `InvalidReference`。
- 集合方法返回新数组，不返回活代理；数组元素是句柄。
- 不支持对象身份比较（`a === b` 只比较 JS 对象实例），需要身份时比较 `getId()` 或地址。

### 3.3 值与类型

严格跟随 ExcelScript 的值语义，不再使用 QuickExcel 自定义的错误对象：

| 场景 | 行为 |
| --- | --- |
| `getValues()` | `(string \| number \| boolean)[][]`；空单元格为空字符串，错误单元格为 `#DIV/0!` 等显示字符串 |
| 精确类型 | 首期未实现 `getValueTypes()`（P1）；脚本可用 `getValue()` 的返回类型与 `getFormula()` 区分空白、文本、数值、布尔与错误 |
| 日期 | 按序列值（number）返回；需要日期语义时读取数字格式分类 |
| `null` / `undefined` | 不作为单元格值出现；参数校验失败抛出 `InvalidArgument` |
| 公式 | `getFormula()` / `getFormulas()` 只读文本，不触发计算；`getValues()` 在事务草稿内按需计算 |

这意味着旧 API 的 `{ kind: 'error', code }` 值被移除。脚本需要错误码时使用 `getValueTypes()` 与公式文本自行判断，或调用 QuickExcel 扩展 `host.values.getErrorCode(range)`（P1 提供）。

### 3.4 名称、地址与区域

- A1 地址支持单格与矩形；整行、整列与带工作表前缀的地址暂不支持，抛 `InvalidArgument`。
- 工作表可按名称或 ID 获取；重名时名称解析结果由宿主决定并保持稳定。
- 多区域选区只能通过 `getSelectedRanges()` 获取；`getSelectedRange()` 在多区域时抛错，与 ExcelScript 一致。

## 4. 能力清单与运行时

### 4.1 清单格式

运行时不再手写对象方法，而是由生成的能力清单驱动。清单只列**已实现**成员；未列出的成员由代理统一拒绝。

```json
{
  "apiVersion": 3,
  "source": { "file": "office-scripts-docs.d.ts", "commit": "<sha>", "sha256": "<hex>" },
  "types": {
    "Workbook": {
      "members": {
        "getWorksheet": { "kind": "host", "host": "sheet", "arity": 1, "cost": 1 },
        "getWorksheets": { "kind": "host", "host": "sheets", "arity": 0, "cost": 1 },
        "getSelectedRange": { "kind": "host", "host": "active", "arity": 0, "cost": 1 }
      }
    },
    "Worksheet": {
      "members": {
        "getId": { "kind": "local" },
        "getRange": { "kind": "local" }
      }
    }
  }
}
```

- `kind` 只有 `host` 与 `local` 两种。`local` 是纯 JS 句柄运算（地址解析、对象构造、`getFormat` 等），不产生 IPC，也没有 `host`/`arity`/`cost`。
- `host` 是 `ScriptModelAdapter` 的方法名，不是对象限定名；JS 与宿主双向校验，任何一侧缺失都视为构建错误。
- `arity` 是宿主调用参数个数，包含句柄上下文（例如读值传 `sheetId` 与地址，通常是 2 个参数），不是脚本侧可见的参数个数。
- `cost` 是该成员单次调用的单元格上限（读取 100000、写入 20000、标量 1），宿主再按实际单元格数累计 76 号文档的读取/写入预算，并另外限制 4096 次请求。

### 4.2 代理生成

每个类型首次实例化时构建原型，后续复用：

```js
function makeHandle(type, data) {
  const target = Object.create(prototypeFor(type));
  Object.defineProperty(target, "__handle", { value: Object.freeze(data) });
  return Object.freeze(new Proxy(target, { get(t, key) {
    if (key in t) return t[key];
    if (key === "then" || key === "toJSON" || key === Symbol.toPrimitive) return undefined;
    return unsupported(type, key);
  }}));
}
```

必须处理 `then`、`toJSON`、`Symbol.iterator`、`Symbol.toPrimitive`、`toString`、`constructor`，否则对象会被误判为 thenable 或干扰日志序列化。未实现成员返回一个函数，调用时抛 `UnsupportedApi`，错误消息包含类型、成员和建议的替代接口；属性读取同样抛错。

### 4.3 错误模型

| 代码 | 含义 | 是否使事务失败 |
| --- | --- | --- |
| `UnsupportedApi` | 成员未实现 | 否（未发生宿主写入） |
| `InvalidArgument` | 地址、矩阵尺寸、类型不合法 | 否，但如果此前已有宿主写入，事务仍按既有规则失败 |
| `ItemNotFound` | 工作表、表格、图表不存在 | 否 |
| `InvalidReference` | 句柄失效或跨工作簿 | 否 |
| `BudgetExceeded` | 超出请求数或单元格预算 | 是 |
| `Locked` | 只读、保护、活动请求失效 | 是 |
| `TransactionAborted` | 宿主拒绝提交 | 是 |

错误对象是 `Error` 子类，带只读 `code` 与 `member`，类型声明为 `QuickExcel.ScriptError`。计算错误仍然是值，不是异常。

### 4.4 与现有事务契约的关系

- 一次执行一个事务，读己写，成功后一次提交、一次撤销，失败整体回滚（76/77 号文档不变）。
- 被 `try/catch` 捕获的宿主写入失败仍使事务失败；`UnsupportedApi` 在 IPC 之前抛出，可安全捕获且不污染事务。
- 公式读取在事务草稿内计算；`getFormula` 不触发计算；不可用函数或受阻依赖仍抛错而不是返回旧缓存。
- 隔离、超时、栈、内存与 worker 复用边界全部不变；本文不放松任何 AppContainer、Landlock、seccomp 或 systemd 要求。

## 5. 类型声明生成

### 5.1 来源

- 文件：`OfficeDev/office-scripts-docs-reference` 的 `generate-docs/script-inputs/office-scripts-docs.d.ts`，约 487 KB，近 300 个类型。
- 固定 commit、git blob SHA-1 与文件 SHA-256，存放于 `third_party/office-scripts-docs/`，附 `NOTICE.md`、`source.lock.json` 与来源说明。第一次拉取用 `--update-lock` 写入 SHA-256，之后每次校验 blob 与哈希，任何不一致直接中止。
- 许可证为 CC-BY-4.0：允许商用与修改，但必须署名、附许可证链接并注明修改。本项目将命名空间改为 `QuickExcel`、删除 `OfficeScript` 命名空间并剥离文档注释，均属于修改，必须在 NOTICE 与根目录 `THIRD-PARTY-LICENSES.md` 中写明。许可证正文以链接方式提供，符合 CC-BY 对 URI 的要求。
- 文档注释一律不进入生成物。注释是整套材料中最具独创性的部分，剥离后既降低版权风险，也显著缩小声明体积；已实现成员的说明由我们自己的能力清单维护。

### 5.2 生成物

| 产物 | 位置 | 用途 |
| --- | --- | --- |
| `quickexcel.d.ts` | `QuickExcel.Web/src/` | Monaco 提示、段落导航、编译 worker 的唯一声明 |
| `capabilities.json` | `QuickExcel.Scripting/` 内嵌资源（`ScriptCapabilities`） | 运行时代理与宿主白名单的共同数据源 |
| 宿主白名单 | `QuickExcel.Infrastructure/Scripting/ScriptModelAdapter` 直接读取同一内嵌清单 | 方法名、参数个数与单次上限校验，避免生成第二份手写白名单 |
| 能力覆盖报告 | `artifacts/`，不入库 | 已实现/未实现成员统计与差异 |

生成器默认输出到 `artifacts/api-gen/`，确认差异后再用 `--out` / `--capabilities` 指向产品位置，避免误覆盖。`quickexcel.d.ts` 不再手工编辑，任何修改都通过生成器或手写的扩展声明片段合并。

### 5.3 生成流程

1. 校验固定文件的 git blob SHA-1 与 SHA-256；失败即中止，不回退到网络最新版本。
2. 用 TypeScript Compiler API 解析声明，读取每个接口、枚举、类型别名和成员签名。
3. 重写命名空间 `ExcelScript` → `QuickExcel`，删除 `OfficeScript`，剥离全部文档注释，只保留成员名、参数名与返回类型。
4. 合并手写的 `quickexcel-extensions.d.ts`（`Host`、`VirtualFiles`、`Events`、`ActionContext`、`Event` 等）。
5. 输出声明与能力清单；以 `--check` 模式比对已提交文件，CI 中不一致即失败。
6. 生成物带有头部注释：来源 commit、SHA-256、许可证与"已修改"声明。

### 5.4 编译预算

`quickexcel.d.ts` 从约 6 KB 增长到数百 KB。剥离注释后体积明显小于原始 487 KB，但仍需实测完整检查耗时，包括 `lib.es*` 与用户模块。若超出预算，依次采用：进一步压缩声明（只保留已实现类型）、提高 TS 编译截止时间。不得关闭完整检查或退化为仅转译。

### 5.5 当前状态

2026-11-04：`tools/QuickExcel.ApiGen` 已落地，包含来源拉取与校验、声明与清单生成、覆盖报告、`--check` 比对和基于合成 fixture 的自测；`third_party/office-scripts-docs` 已固定来源与 SHA-256，根目录新增 `THIRD-PARTY-LICENSES.md`。

真实声明的首次生成结果：315 个类型、1932 个成员，剥离注释并删除 `OfficeScript` 后约 120 KB；能力清单覆盖 6 个类型、30 个成员（含 `Range.clear` 与 `RangeAreas`）。完整类型检查（`lib.es2020` + 声明 + 示例模块）实测 0.86–0.96 秒，在 5 秒预算内。

P0 运行时已实施并发布声明：能力清单嵌入 `QuickExcel.Scripting`，JS 侧由 Proxy 按清单生成 local/host 成员，宿主 `ScriptModelAdapter` 用同一清单校验方法名与参数个数；`scriptApiVersion` 仍写 2、IPC 升至 10，存储格式与 1/2 迁移提示留待存储切片；`QuickExcel.Web/src/quickexcel.d.ts` 已替换为生成物。

## 6. 版本与迁移

- `scriptApiVersion = 3`。文件写入 3；读取 3 正常执行。当前运行时切片尚未切换文件版本，见 §5.5。
- 读取 1/2 时允许打开工作簿但禁用脚本执行，给出明确诊断"脚本使用已废弃的 API v2，需要重写"；保存时要求移除旧脚本或改写为 v3，不自动转换源码。
- IPC 版本提升为 10，桥标识同步更新；旧 host/worker/native 混用继续拒绝。
- 模块导入、签名、事件、虚拟文件、控制台的存储与授权规则不变；签名仍覆盖源码原文，不覆盖生成声明。
- 示例、模板、诊断页与文档中的脚本片段全部重写为新 API。

## 7. 分阶段实施

| 阶段 | 对象范围 | 退出条件 |
| --- | --- | --- |
| P0 | 生成器、命名空间、代理运行时、Workbook / Worksheet / Range / RangeAreas、v3 存储与 IPC | 生成物可复现；核心示例可编译执行；未实现成员稳定报 `UnsupportedApi`；旧 API 完全移除。2026-11-04：除 v3 存储/迁移外已完成，`RangeAreas` 已支持多区域选区的本地句柄 |
| P1 | RangeFormat / Font / Fill / Border / Alignment / NumberFormat / Protection、Table / Sort / AutoFilter | 格式化与表格脚本通过定向用例；预算映射验证 |
| P2 | Chart、ConditionalFormat、DataValidation、Comment | 对象句柄与宿主模型对齐；未实现成员清单收敛 |
| P3 | PivotTable、Shape、Slicer/Timeline、PageLayout、DocumentProperties、CustomXml、Query | 按实际需求推进；其余继续显式拒绝 |

每个阶段只声明已实现成员；类型骨架始终是全量的，两者通过能力清单区分。

## 8. 验证计划

| 检查 | 方法 | 通过标准 |
| --- | --- | --- |
| 声明来源 | SHA-256 与固定 commit 比对 | 完全一致 |
| 生成确定性 | 连续两次生成并比对，`--check` 模式 | 无差异 |
| 类型可用性 | 用固定 TS 5.9.3 编译新 API 示例 | 零错误，且提示包含成员文档 |
| 编译预算 | 记录完整检查耗时分布 | 在截止时间内；超出则按 5.4 处理 |
| 运行时 | JSC/隔离 worker 冒烟 | 支持成员可用；未实现成员抛 `UnsupportedApi`；事务回滚正确 |
| 预算 | 4096 次请求边界与单元格预算 | 与 76 号文档一致，代理开销可测 |
| 许可 | NOTICE、LICENSE、文档署名审查 | 满足 CC-BY-4.0 |
| 回归 | 只跑脚本、存储、worker 相关定向用例 | 不跑全量回归 |

## 9. 风险与未决

- **许可证性质**：CC-BY-4.0 是内容许可，不含专利授权条款；是否直接分发该声明需要一次正式法务确认。备选方案是仅作参考、自写同名声明，但工作量和漂移风险更高。
- **编译时限**：500 KB 声明与 5 秒截止时间可能冲突，P0 必须先测量再决定是否调整预算。
- **API 漂移**：官方定义随 Excel 更新；固定版本后每季度做一次差异报告，不自动跟随。
- **语义差异**：空单元格、错误字符串、区域设置、日期与浮点精度都可能与 Microsoft 运行时不同；这些差异必须在声明注释和用户文档中逐条列出，不能用"兼容 ExcelScript"一句话带过。
- **全量骨架的误导**：类型全量、实现少量，用户可能认为都能用。除了 `UnsupportedApi`，编辑器还需在未实现成员上显示能力标记，文档提供覆盖矩阵。官方声明中的枚举尚不作为运行时值注入，首期已实现成员不接受枚举参数；P1 实现枚举型成员时需同时提供 `QuickExcel` 枚举值对象。

## 10. 待确认

- `Host` 是作为 `main` 的第二参数，还是挂在 `workbook` 上；本文按第二参数设计。
- `host.values.getErrorCode()` 这类 QuickExcel 扩展是否纳入 P1。
- 未实现成员是否在 Monaco 中以删除线或警告标记，而不是仅靠运行时错误。
