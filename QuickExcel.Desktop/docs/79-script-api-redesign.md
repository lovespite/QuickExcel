# 脚本 JavaScript API 重新设计（API v3）

2026-11-04。本文替换 [76](76-live-workbook-script-model.md) 中的对象 API 部分：旧 `$` / `ExcelScript` 子集整体废弃，不提供兼容层。文件 `scriptApiVersion = 3`、IPC 10；隔离、事务、限额和 worker 生命周期契约沿用 11/15/16/69/72/77 号文档。P0、P1、P2 已实现；成员与参数支持边界见 §7.1 / §7.2。

## 1. 现状与问题

重做前脚本只能看到 5 个对象、14 个方法（Workbook / Worksheet / Range / RangeFormat / RangeFill），声明、JS 实现与宿主白名单分别手工维护。现在以固定声明与能力清单生成代理、类型和白名单；当前实现 **41 个类型的 353 个成员（353/1932，18.27%）**。

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
| 精确类型 | `getValueType()` / `getValueTypes()` 返回 `QuickExcel.RangeValueType`，区分空白、文本、整数、浮点、布尔与错误 |
| 日期 | 按序列值（number）返回；需要日期语义时读取数字格式分类 |
| `null` / `undefined` | 不作为单元格值出现；参数校验失败抛出 `InvalidArgument` |
| 公式 | `getFormula()` / `getFormulas()` 只读文本，不触发计算；`getValues()` 在事务草稿内按需计算 |

旧 API 的 `{ kind: 'error', code }` 值已移除。`host.values.getErrorCode(range)` 返回左上角单元格的错误显示代码或 `null`；这能区分错误值与内容相同的普通文本。

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

固定来源共 315 个类型、1932 个接口成员，剥离注释并删除 `OfficeScript` 后约 120 KB。P0 首次生成覆盖 6 个类型、30 个成员；P1 后覆盖 **16 个类型、174 个成员（9.01%）**。覆盖率按能力清单计数，不表示全部参数组合或 Microsoft 运行时语义均已实现；未支持的组合显式拒绝。`host` 扩展与枚举值不计入 174。完整类型检查（`lib.es2020` + 声明 + 示例模块）首次实测 0.86–0.96 秒，在 5 秒预算内。

能力清单嵌入 `QuickExcel.Scripting`，JS 侧由 Proxy 按清单生成 local/host 成员，宿主用同一清单校验方法名与参数个数。P1 成员通过 `getApi:类型.成员` / `setApi:类型.成员` 分发，传输参数为句柄和脚本参数数组。所用枚举从固定 AST 生成，注入不可变 `QuickExcel` 全局对象。几何边界与 Core 一致：200000 行、2048 列。生成器验证来源锁、SHA-256 与 git blob SHA-1；合成 fixture 使用独立覆盖输入。

## 6. 版本与迁移

- `scriptApiVersion = 3`。新建文件、示例、执行与编译 worker 均使用 3；worker 握手只宣告 3。函数注册描述符中的 `apiVersion: 1/2` 是独立的函数参数语义，继续保留。
- 文件版本 1/2 允许打开和查看源码，但禁用脚本执行与保存。诊断为“脚本使用已废弃的 API v1/v2，需要按 QuickExcel API v3 重写，或移除全部脚本后保存。”不自动转换源码。脚本窗口提供“确认使用 API v3”和“移除全部脚本”：前者由用户确认源码已改写，后者移除脚本；两者均走可撤销命令，清除旧签名并撤销既有授权。只编辑旧源码不会偷偷变更文件版本。
- IPC 版本提升为 10，桥标识同步更新；旧 host/worker/native 混用继续拒绝。
- 模块导入、签名、事件、虚拟文件、控制台的存储与授权规则不变；签名仍覆盖源码原文，不覆盖生成声明。
- 示例、模板、诊断页与文档中的脚本片段全部重写为新 API。

## 7. 分阶段实施

| 阶段 | 对象范围 | 退出条件 |
| --- | --- | --- |
| P0 | 生成器、命名空间、代理运行时、Workbook / Worksheet / Range / RangeAreas、v3 存储与 IPC | 已完成：确定性生成、JS/TS 示例、UnsupportedApi、v1/v2 查看与迁移保存门禁、v3 执行/存储与旧入口移除 |
| P1 | RangeFormat / Font / Fill / Border / Alignment / NumberFormat / Protection、Table / Sort / AutoFilter | 已完成 §7.1 范围：格式化、表操作、预算、原子提交/失败回滚、保存重开与撤销的定向用例；生产包自检覆盖真实隔离 worker |
| P2 | Chart、ConditionalFormat、DataValidation、Comment | 已完成 §7.2 范围：真实 Core 对象、稳定会话句柄、事务预算与保护、JS/TS 定向检查、XLSQ/XLSX 往返及生产包自检 |
| P3 | PivotTable、Shape、Slicer/Timeline、PageLayout、DocumentProperties、CustomXml、Query | 按实际需求推进；其余继续显式拒绝 |

每个阶段只声明已实现成员；类型骨架始终是全量的，两者通过能力清单区分。

### 7.1 P1 支持矩阵与限制

| 范围 | 已支持 | 明确未支持 |
| --- | --- | --- |
| 字体、填充 | 字体名/大小/颜色、粗体、斜体、删除线、下划线；纯色填充和清除；混合格式标量读取返回 null | 上下标、色调与图案填充 |
| 边框 | 四周与内部横/纵线、颜色、已建模的实线/虚线/点线/双线及可保存的粗细 | 对角线、点划线等未建模线型，以及无法无损保存的线型/粗细组合 |
| 区域格式 | 对齐、换行、缩小、缩进、旋转、行高列宽（API 点与 Core DIP 换算）；数字格式标量/矩阵读写和分类；文本/值类型/公式矩阵 | 自动适应、阅读顺序、自动缩进和区域设置转换；颜色参数使用 #RRGGBB |
| 保护 | 单元格 locked/formulaHidden；无密码工作表 protect/unprotect、选项与状态读取 | 密码、暂停保护、可编辑范围 |
| 表和列 | 按名称/ID 获取、稳定表/列句柄、创建、重命名、范围/行数、样式/条纹/首尾列标记、底边 resize、表尾 addRow/addRows、convertToRange/delete | 创建时必须有唯一普通文本表头和至少一行数据；暂不自动生成表头、中间插入、增删列或切换汇总/表头 |
| 表尾追加 | 空白目标区追加并保持已有行；显式值矩阵或单空行 | 含汇总行、覆盖目标数据/链接/批注；不隐式移动表外数据 |
| 排序 | 区域/表的稳定多键行排序，键为区域内列偏移；值/字体色/填充色、大小写、表排序状态清除/重用 | 按列、图标、textAsNumber、显式拼音/笔画 method；默认文本为序数比较 |
| 筛选 | 工作表与表 AutoFilter、列 Filter；值集合、自定义条件/AND/OR/通配符、状态/条件读取、清除/重新应用 | Top/Bottom、动态日期、颜色、图标等高级筛选 |
| 清除与错误值 | clear(all/formats/contents)、host.values.getErrorCode | all 遇到链接/批注拒绝，不能静默删除附件；其他 clear 模式未实现 |

格式写入沿用 Core 命令，保留富文本继承字体。表样式解析使用既有主题/样式定义，转换为区域时固化有效样式，并将结构化引用转换为 A1；删除表将其引用变为 #REF!。带活动筛选的表必须先清除筛选再转换。表 ID 在会话内预留，删除后新建不复用旧句柄身份。

所有写入在宿主私有草稿中验证，保护状态按命令顺序检查，成功一次提交/撤销，失败连同虚拟文件整体回滚；捕获的写入失败仍不可提交。公式排序/筛选使用草稿中的当前值，提交复用已校验的键/条件，不因提交重复调用 UDF。计算后的格式、保护与筛选更新同步到私有值视图。

单轮累计读 100000 格、写 20000 格（重复操作重复计数）、最多 4096 请求；大范围格式、排序、筛选和表结构操作也计费，不按单次标量请求绕过预算。原始模块源码总额仍为 1 MiB；原生加载器额外预留 32 KiB，仅用于宿主为最多 64 个动作入口追加的注册代码，用户输入限制不变。AppContainer、编译/执行进程和时间预算均未放宽，JSC-MEM-001 仍 OPEN。

定向证据位于 `artifacts/test-results/api-p0-p1/`。新增 `ScriptP1Tests` 覆盖 JS/TS 格式/枚举、表转换和有效格式、尾部追加、公式排序/筛选、保护与虚拟文件原子性、失效句柄、累计预算和 1 MiB 源码；`ScriptApiMigrationTests` 用独立 ZIP/JSON 构造 v1/v2 输入。发布自检增加 API v3 TypeScript 格式、表、排序、筛选和撤销检查。

2026-11-04 验证：解决方案 Release 构建零警告/错误，生成物 `--check` 与生成器 fixture 自测通过，JSC smoke 通过（10000 次调用 98 ms）。`verified-summary.json` 汇总定向运行与修复后局部复测的每个用例最新结果：701 通过、0 失败、1 跳过（已知 JSC-MEM-001）；不是全量回归，也不表示单次运行 701 项。旧批次故障注入工具已适配 v3 持久 worker，模块状态测试改为 Windows 跨轮保留/Linux 单轮重建的既有契约。

交付：`scripts/build-quickexcel.ps1 -SkipTests -RunSmoke` 已发布并激活 `artifacts/avalonia/releases/20261104-112512-75d6b7ff/application`，生产自检 10 项通过，`active-windows.json` 的目录和发布清单 SHA-256 已核对。首次自检曾因测试流程未持有正常 UI 执行门禁，被自动重算取消；补上同一执行门禁后重新发布通过，未放宽取消或隔离契约，失败包保留证据且从未激活。

### 7.2 P2 支持矩阵与限制

P2 新增 179 个成员，累计 353 个，涉及 41 个类型。枚举与声明仍从固定来源生成；类型骨架中的其他成员及未建模参数继续显式拒绝，不能将该统计解释为全部 Office Scripts 兼容。

| 范围 | 已支持 | 明确限制 |
| --- | --- | --- |
| Chart | 工作表图表创建/获取/删除，名称/ID/所属表，类型、位置尺寸、数据范围、可见数据选项；标题、图例、背景/系列填充和标题/图例字体；系列名称、筛选、平滑、标记大小和数据引用；数值轴上下限/数字格式 | 簇状柱形、折线、饼图、散点、瀑布、漏斗；首行标题、首列分类、按列分系列；不支持组合/3D/趋势线/导出图像或按行分系列。系列引用为同表单行/列且与分类长度相同；不隐式补齐数据。分类轴数值属性仅用于散点图 X 轴。自动轴界限读取为 null，不伪造布局结果 |
| ConditionalFormat | 区域规则集合、稳定 ID、优先级重排、StopIfTrue、范围、删除/清除；自定义公式、数值比较、文本比较、唯一/重复值；条件字体/填充/数字格式/四周边框 | 非易变已建模标准公式；脚本暂不创建/配置色阶、数据条、图标集、Top/Bottom、日期/平均值预设。桌面及文件层已支持三种视觉规则（见 docs/54）；脚本 getType 返回对应类型，范围/优先级等基础操作保留其完整元数据和身份，视觉规则拒绝 StopIfTrue=true。setRanges 当前接受单矩形，已导入多区域可用 getRanges 读取；部分覆盖清除拒绝，保留范围外规则 |
| DataValidation | 整数、小数、日期序列公式、文本长度、列表、自定义公式；提示/错误警告、忽略空白、规则类型、有效性、无效单元格、清除 | 不支持 time 和 JS Date 参数；日期使用序列值公式。列表为直接文本或同表单行/列 Range。混合规则显示 inconsistent，须分块读取；替换/清除不越界删除部分相交规则。getInvalidCells 最多返回 1024 个单格区域；禁用错误提示仍检查有效性 |
| Comment / CommentReply | 工作簿/工作表创建和查找真正的线程评论；文本、作者显示名、创建日期、位置、ID、回复、解决/重新打开、删除；批注与回复一起撤销、保存重开后继续编辑 | 纯文本、本地作者 QuickExcel，无账户查找/邮件/通知；不虚构邮箱身份，不支持 mentions/rich content。已解决线程需先重新打开才能回复。带旧版显示副本的导入讨论继续只读保留，不写成旧式 Note；未知扩展仍走原包保护路径 |

图表和条件格式的会话 ID 在属性修改及撤销/重做中保留；删除后重建、图表粘贴使用新身份，旧句柄返回 InvalidReference。重新打开文件建立新会话身份；评论 UUID 则保存在 OOXML threadedComments/person 部件中。原始线程未修改时沿用字节保留；编辑纯文本线程时保留作者 ID、账户元数据、父子关系、时间与解决状态。旧版副本和未知内容的保护未放宽，带评论的结构移动/复制仍禁止。

规则写入按实际区域格数累计收费（重复操作重复计数），通过 Core 命令检查保护；图表与批注也不能绕过对象保护。条件格式、验证规则、评论、单元格与虚拟文件属于同一原子动作，捕获写入失败仍整体回滚。新增文本计入历史和 XML 写出预算。嵌套参数中的 Range 通过句柄转换，循环/深层输入被拒绝。

P2 定向证据在 `artifacts/test-results/api-p2/`：`ScriptP2Tests` 覆盖 JS/TS、图表轴与系列、条件边框、线程回复/删除/重新打开/多工作表/XLSX 导出、部分范围保护、预算和失败回滚；现有图表、线程保留、条件格式、验证、存储及 P1 回归保持覆盖。生产自检增加 P2 对象事务及其 XLSQ 保存重开，仍验证整个动作一次撤销。JSC-MEM-001 继续按既有决定延期，P2 不改变隔离、进程/栈/时间限制。

2026-11-04 验证：`api-p2-regression.trx` 237 项、`api-p2-extended.trx` 53 项、`api-p2-final.trx` 49 项均通过；`verified-summary.json` 按用例保留最新结果，去重后 **265 通过、0 失败**，不是全量回归。生成物 `--check` 与生成器 fixture 自测通过；JSC smoke 通过（10000 次调用 91 ms）。发布构建成功，保留两个既有生产配置未赋值字段警告（`Program.WebScriptsOnly` / `WebModulesOnly`）。

交付：`scripts/build-quickexcel.ps1 -SkipTests -RunSmoke` 已发布并激活 `artifacts/avalonia/releases/20261104-120050-b8cbf778/application`。生产自检 **12 项通过**，包括真实隔离 worker 中 P2 对象事务、XLSQ 保存重开和动作撤销；`active-windows.json` 指向该包，发布清单 SHA-256 已核对，与 `Start-QuickExcel.cmd` 使用的启动选择一致。

## 8. 验证计划

| 检查 | 方法 | 通过标准 |
| --- | --- | --- |
| 声明来源 | SHA-256 与固定 commit 比对 | 完全一致 |
| 生成确定性 | 连续两次生成并比对，`--check` 模式 | 无差异 |
| 类型可用性 | 用固定 TS 5.9.3 编译新 API 示例 | 零错误，提示包含完整成员签名；上游文档注释不复制 |
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
- **全量骨架的误导**：类型全量、实现少量，用户可能认为都能用。运行时 `UnsupportedApi` 和本页覆盖矩阵已提供；Monaco 的逐成员能力标记仍未实施。P1 使用的枚举已注入运行时，其余类型骨架中的枚举不代表已有运行时能力。

## 10. 待确认

- `Host` 已确定为 `main` 的第二参数；`host.values.getErrorCode(range)` 已纳入 P1。
- 未实现成员是否在 Monaco 中以删除线或警告标记，而不是仅靠运行时错误。
