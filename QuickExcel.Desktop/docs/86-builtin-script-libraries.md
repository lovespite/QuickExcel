# 内置 JavaScript 数据处理库与兼容性验证

QuickExcel 7.0.0 内置以下九个固定版本的 JavaScript 库。脚本使用静态
`import` 即可；安装后的应用不需要 Node、npm 或联网下载。支持 JavaScript
与 TypeScript，编辑器使用同一组上游类型声明，定义可只读打开。

这里的“通过”指下表所列功能在实际隔离 JSC worker 中通过验证，不表示已经
运行各项目全部上游测试，也不表示提供完整浏览器或 Node.js 环境。

## 逐库结果

| 库 / 固定版本 | 验证通过的功能 | 发现的问题与已补齐内容 | 仍有的兼容性限制 |
| --- | --- | --- | --- |
| Lodash 4.18.1 | 分组、排序、去重、链式操作、Unicode 分词、模板、命名导出 | 将完整 CommonJS 构建封装成 ESM，保留默认 `_` 与命名导出 | `delay`、`defer`、`debounce`、`throttle` 等依赖后台定时器的调用不支持；没有伪造定时器 |
| decimal.js 10.6.0 | 十进制精度、舍入、超安全整数、独立精度配置 | 无需运行时补丁 | 大数对象不能直接作为单元格值，需转字符串或数值；转 Number 仍可能损失精度 |
| Day.js 1.11.23 | 闰年、无效日期、UTC、时区、严格格式解析、Duration | 预装 15 个常用插件与 `zh-cn`，修复仅基础包时缺少 UTC/tz/duration、严格解析无效的问题 | 只有打包插件可用，不支持任意插件子路径导入；默认 locale 仍为英文；无时区日期沿用运行环境本地时区 |
| Papa Parse 5.7.1 | 引号/逗号/跨行 CSV、表头、自动类型转换、导出 CSV、逐行回调 | 提供默认与命名导出；将不支持的浏览器模式改为明确报错，并收窄类型声明 | 不支持 `worker:true`、`download:true`、FileReader/Node stream；用 `workbook.files` 或 Action 中的授权 fetch 取得文本后再解析；Blob 可先 `await blob.text()` |
| simple-statistics 8.0.0 | 均值、标准差、分位数、回归、相关性、聚类 | 无需运行时补丁 | 统计前应清理空值、字符串和无效数值；未验证每个统计函数 |
| Zod 4.6.7 | 对象验证、失败结果、转换、JSON Schema、URL、File | 增加纯数据 URL 实现和内存 File，使合法 URL 不再被错误拒绝，File schema 可以执行 | File 只是内存 Blob，不代表磁盘权限；异步 refinement 应在支持 Promise 的 Action 中等待 |
| Fuse.js 7.5.0 | 模糊搜索、序列化索引重建、扩展搜索 | 无需运行时补丁 | 搜索索引和工作集仍受 worker 内存与操作时限约束 |
| Arquero 8.0.3 | 分组聚合、连接、透视、窗口、CSV、Unicode Arrow 往返、HTML 字符串、固定宽度/JSON/CSV/Arrow 流输入、授权 CSV 下载 | 使用浏览器依赖映射；四个远程加载器走宿主 fetch；四个流读取器兼容字符串/字节并缓冲 UTF-8；修复 Arquero/Flechette 两处上游声明错误 | 压缩文件的 gzip/deflate codec 未提供；流读取先缓冲后解析，输入上限 10 MiB；没有通用 ReadableStream/TransformStream 构造器；远程加载仅限获授权的 Actions |
| math.js 15.2.0 | 矩阵、单位、表达式、符号求导、大数、复数、线性方程、随机数 | 使用浏览器依赖映射，避开 Node 专有依赖 | 无 Node 文件/终端能力；不是完整 CAS；随机数不作安全随机数保证；大型计算仍受时限约束 |

许可证：Lodash、decimal.js、Day.js、Papa Parse、Zod 为 MIT；
simple-statistics 为 ISC；Fuse.js 与 math.js 为 Apache-2.0；Arquero 为
BSD-3-Clause。完整原文、传递依赖与类型包声明随安装包一起发布，可在设置的
开源信息中打开。版本、来源、npm integrity 与改动说明见
`tools/QuickExcel.ScriptLibraries/package-lock.json` 和 `NOTICE.md`。

Day.js 预装插件：UTC、Timezone、CustomParseFormat、Duration、IsBetween、
IsSameOrBefore、IsSameOrAfter、IsoWeek、WeekOfYear、QuarterOfYear、
AdvancedFormat、LocalizedFormat、LocaleData、MinMax、RelativeTime。
locale `zh-cn` 另外提供。

math.js 上游声明会把 `add(20,22)` 推断成过窄的字面量联合类型，直接与 `42`
比较会触发 TS2367。计算结果没有此问题；需要这种比较时可用
`Number(add(20,22))`，或在调用处提供普通 `number` 类型参数。本轮保留上游
泛型声明，未用 `any` 覆盖它。

后续可补：gzip/deflate 可通过额外的纯 JS codec 实现，但需要同时限制解压后
大小和耗时，本轮未加入。更多 Day.js 插件/语言可以在固定集合中增加。后台
定时器需要设计任务生命周期和取消语义，不适合用同步等待冒充；DOM、浏览器
Worker 和 Node 文件系统也不会因为内置数据处理库而开放。

## 使用

```ts
import { sumBy } from 'lodash';
import Decimal from 'decimal.js';
import dayjs from 'dayjs';
import Papa from 'papaparse';
import { mean } from 'simple-statistics';
import { z } from 'zod';
import Fuse from 'fuse.js';
import * as aq from 'arquero';
import * as math from 'mathjs';

export function main(workbook: QuickExcel.Workbook) {
  const rows = Papa.parse<{ item: string; amount: number }>(
    'item,amount\napple,20\nbanana,22',
    { header: true, dynamicTyping: true }
  ).data;
  const total = z.number().parse(sumBy(rows, 'amount'));
  workbook.getActiveCell().setValue(new Decimal(total).toFixed(2));
}
```

各库按模块导入，不注入 `_` 等全局变量。命名与大小写必须精确；例如
`lodash` 有效，`./lodash`、`Lodash`、`node:fs` 和网络 URL 不会被作为内置库
解析。Day.js 插件已预装，直接调用 `dayjs.utc()` / `dayjs.tz()` 即可。
同名工作簿模块优先，避免改变已有工作簿的自定义模块含义。

## 运行与安全边界

- 构建时按 lockfile 下载并校验，禁用 npm lifecycle scripts。生成代码、类型、
  缓存均留在 `artifacts/script-libraries`；运行时只读取已嵌入资源。
  发布包 `metadata` 同时保留库 lockfile 和带代码/类型哈希的 catalog。
- 原生模块加载器通过专用回调按需取得已验证库代码。没有磁盘、URL 或 npm
  模块搜索；内置库不成为 UDF/Actions/Events 入口。
- 用户源码原有 64 模块 / 1 MiB 上限保持；内置模块单独限制为最多 16 个 /
  4 MiB 源码。单库按需加载，未导入库不执行初始化代码。
- URL 使用 core-js-pure 3.51.0（MIT）的纯数据实现；内存 File 不访问磁盘。
  未添加操作系统网络/文件权限、DOM、后台任务或浏览器 Worker。
- 网络沿用初始化声明、会话同意、宿主代理与精确主机授权。支持普通 Arquero
  远程读取；响应体仍使用用户设置的网络限额，默认 8 MiB、最大 10 MiB。
- UDF 与 Actions/Events 的持久上下文保持隔离。数据修改仍经宿主事务，异常
  时回滚，成功后可撤销。不要依赖库的可变配置跨这些上下文共享。
- 原有 256 MiB 进程约束、JS 栈、1 秒原生操作看门狗、宿主时限及 16 MiB IPC
  预算不变。内置更多库不意味着允许更大工作集。流读取是受限缓冲，不是
  无限数据流处理。
- 库代码、类型与适配层哈希纳入 worker 握手、编译器缓存与执行同意策略。
  库升级随产品显式发布；旧的持久同意不能自动沿用到不同库集合。
- XLSQ 仍只保存用户源码，不复制 npm 包，也尚不保存可选的历史库版本锁。
  因此源码签名不是历史运行时版本承诺；需要旧版本重现时保留完整旧安装包。
  旧产品打开新导入时会拒绝不认识的模块，不会自动下载或改写源码。

## 验证与证据

`BuiltInLibraryTests` 包含九库共 51 个功能/不支持行为检查，以及九库真实
隔离编译器类型检查、错误类型/非法路径拒绝、全部库同一 worker 执行、重复
调用、用户源码配额边界、同名模块优先、事务提交/回滚/撤销、授权网络、UTF-8
跨块读取、流上限与压缩拒绝测试。功能预期在 fixture 中明确保存。

编辑器 `ScriptNavigationSmoke` 在真实 Monaco 中检查九库类型，并打开只读
Lodash 定义。发布自检通过真正的隔离 Native AOT worker 同时导入全部九库，
执行真实 TypeScript 编译并验证结果与 URL/File，不用测试夹具替代发布内容。

相关验证命令（仓库根目录、Windows PowerShell 7）：

```powershell
dotnet test QuickExcel.Core.Test/QuickExcel.Core.Test.csproj -c Release --artifacts-path artifacts/build --filter 'FullyQualifiedName~BuiltInLibraryTests|FullyQualifiedName~TypeScriptCompilerTests|FullyQualifiedName~M4ModuleTests|FullyQualifiedName~ScriptFetchRuntimeTests|FullyQualifiedName~WorkerJsonSourceGenerationTests'
./scripts/test-avalonia-web.ps1 -ModulesOnly
./scripts/build-quickexcel.ps1 -SkipTests -RunSmoke
```

另以新发布包的 `ScriptWorker` 目录作为 `QuickExcelPublishedWorkerDirectory`
重跑库集成测试，验证独立 Native AOT 产物。测试日志与 TRX 位于
`artifacts/script-libraries`，编辑器证据位于 `artifacts/avalonia/web-tests`，
生产自检与激活记录由常规 Windows 发布流程生成。

2026-11-20 本地结果：

| 验证 | 结果与证据 |
| --- | --- |
| Release 解决方案构建 | 通过，0 警告 / 0 错误；`artifacts/script-libraries/solution-build.log` |
| 相关脚本/模块/编译/网络/签名/授权测试 | 182/182；`library-targeted-final.trx` |
| 独立 Native AOT worker 库集成测试 | 27/27；`library-aot.trx`。九库组合用例改为 TypeScript 后另行通过 `library-aot-combined-types.trx` |
| 实际 Monaco 编辑器检查 | 52/52；`artifacts/avalonia/web-tests/20261120-112330-da20bd77` |
| 最终生产包安装自检 | 37/37，含九库联合 TypeScript 编译/执行与 URL/File；release `20261120-112821-199282a7` |
| 运行依赖 npm audit | 本次查询报告 0 已知漏洞；`audit-final.json`，不代表完整安全审计 |
| 激活 | `active-windows.json` 指向新包，manifest 哈希、许可证/库元数据及已测试 AOT worker 哈希核对通过；`delivery.json`。旧可用包保留 |

未展开目录的日志/TRX 均位于 `artifacts/script-libraries`。第一次候选包
`20261120-112357-91226b84` 因自检示例触发上述 math.js TS2367 被阻止激活；
修正比较表达式并在 AOT worker 中验证后重新发布，未跳过自检。最终生产裁剪
仍有既有的诊断字段与 ODBC/DataSet 裁剪警告；库 worker 的 Native AOT 构建
及新库功能检查通过，本次没有以关闭检查来消除警告。

本轮验证平台是 Windows x64。macOS/Linux 使用相同来源和适配代码，但本轮
没有在这两个系统执行验收，不能把 Windows 结果当作跨平台验证结果。
