# QuickExcel.ApiGen

从固定的 Office Scripts 声明文件生成 QuickExcel 脚本声明与能力清单。开发工具，不参与产品构建，也不在运行时加载。

## 前置

- Node.js 22 及以上
- 首次使用执行 `npm install`

## 命令

以下命令从仓库根目录执行。

```powershell
# 1. 下载并校验固定来源（首次带 --update-lock 写入 SHA-256）
node .\tools\QuickExcel.ApiGen\src\fetch-source.mjs --update-lock

# 2. 生成声明与能力清单
node .\tools\QuickExcel.ApiGen\src\generate.mjs --out .\QuickExcel.Web\src\quickexcel.d.ts --capabilities .\QuickExcel.Scripting\capabilities.json

# 3. 仅校验仓库中的生成物是否最新
node .\tools\QuickExcel.ApiGen\src\generate.mjs --check --out .\QuickExcel.Web\src\quickexcel.d.ts --capabilities .\QuickExcel.Scripting\capabilities.json

# 4. 生成器自测（不访问网络）
node .\tools\QuickExcel.ApiGen\test\generate.test.mjs
```

默认输出到 `artifacts/api-gen/`，避免误覆盖产品文件；确认差异后再用 `--out` / `--capabilities` 指向实际位置。

## 生成内容

- `quickexcel.d.ts`：全量类型骨架，命名空间改为 `QuickExcel`，移除 Microsoft 的注释与 `OfficeScript` 命名空间，追加手写扩展。
- `capabilities.json`：只列已实现成员，供 worker 代理、宿主白名单和设置页 API 浏览器使用；每个成员附带来自固定声明的 `signature`（不含 Microsoft 文档注释），同时生成所选枚举的数值与工作表边界。
- `coverage.json`：已实现类型/成员与全量类型/成员的覆盖统计，用于评估"类型全量、实现少量"的差距。

## 维护能力清单

`src/capabilities.overrides.json` 是唯一的手写输入：

- 键必须存在于官方声明中，否则生成失败；
- `kind` 为 `local` 时表示纯 JS 句柄运算，不产生 IPC，不写 `host`/`arity`/`cost`；
- `kind` 为 `host` 时，`host` 必须是 `ScriptModelAdapter` 的方法名；
- `arity` 是宿主调用参数个数，包含 `sheetId`、地址等句柄上下文；
- `cost` 映射到 76 号文档的单元格预算。
- 可选的 `summary` 是产品自己维护的成员说明；设置页优先显示它，未提供时显示调用方式。生成器不会写入 Microsoft 文档注释。
- `enums` 从固定声明 AST 生成运行时数值，注入只读 `QuickExcel` 对象；不手写第二套枚举。
- `limits` 必须与 Core 的 `WorkbookLimits` 一致。单次调用与累计预算仍由宿主强制验证。
- `--overrides` 只用于选择其他手写输入（例如合成 fixture），产品默认使用上述文件。固定产品来源每次生成和 `--check` 都校验 SHA-256 与 git blob SHA-1，缺失或不匹配即失败。

未列出的成员保留在类型骨架中，运行时抛 `UnsupportedApi`。

## 测试

生成器自测使用 `fixtures/mini.d.ts` 与独立 `fixtures/capabilities.json` 验证命名空间重写、注释剥离、成员保留、清单过滤、生成确定性和 TypeScript 完整编译。测试不下载真实声明。

## 许可

生成物包含 CC BY 4.0 材料。分发前必须保留 `third_party/office-scripts-docs/NOTICE.md` 中的署名与修改声明。
