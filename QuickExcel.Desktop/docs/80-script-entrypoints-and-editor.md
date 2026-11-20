# 初始化主入口、自动发现与脚本编辑器

2026-11-05，产品 2.0.0。新工作簿使用文件 `scriptApiVersion: 4`，IPC 11。对象模型及能力清单沿用 79 的实现边界；本章替换其入口与注册规则。API v3 文件保持原版本执行与保存，不自动改写源码；旧 v1/v2 的执行和保存门禁不变。v3 的注册方法只存在于版本隔离的兼容路径，v4 的 `host` 没有 `functions`、`events`、`actions` 或注册方法。

产品 3.8.2 起，设置页的脚本 API 标题与新建工作簿共用 `ScriptModuleSet.CurrentApiVersion`，升级当前 API 时无需另改设置页版本号。`ScriptCapabilities.ApiVersion` 是 v3/v4 共用的能力清单契约版本，不用于显示当前文件 API；既有文件仍保留自己的版本声明。

产品 6.1.0 当前使用 IPC 13，文件 API 仍为 v4。初始化与 Actions 支持 async / 返回 Promise。网络地址只能在初始化根模块顶层声明；用户授权后也只有 Actions 可以调用 fetch。设置中的网络总开关、正文 / Action 超时 / 网络面板限额，以及保持不变的 UDF / 初始化 / Events / 控制台执行边界，见 [85：脚本 fetch 与网络面板](85-script-fetch-and-network.md)。

## 角色与加载

产品 7.0.0 内置九个固定数据处理库，提供静态导入、完整上游类型声明与只读定义导航；文件 API v4 和 IPC 13 保持不变。具体清单、适配及限制见 [86：内置 JavaScript 数据处理库](86-builtin-script-libraries.md)。

### 环境级禁用（3.7.0）

设置 → 脚本与 API → **禁用脚本**保存为本机用户偏好。勾选立即撤销当前工作簿授权，取消正在运行的执行/编译并终止常驻 worker。取消勾选不会自动重新授权，须显式应用脚本。

启动参数 `--no-scripts`（可放在文件路径或其他参数前后）及 `QUICKEXCEL_NO_SCRIPTS=1` 也会禁用脚本。启动时独立读取进程环境与 Windows 用户、机器环境；任一范围为 `1` 均生效，其他范围的 `0` 不可覆盖。命令行及环境变量变更在下次启动生效，设置复选框在当前进程立即生效。

三种来源为 OR 关系，优先于有效签名、证书信任及已接受的引擎策略。状态栏独立显示 `No scripts - User setting`、`No scripts - Command line`、`No scripts - Env`，多来源按此顺序以 `/` 连接；其他操作的状态消息不覆盖此提示。手动应用、执行操作、控制台及需要编译的签名操作明确报告“当前环境已禁用脚本。”。

限制覆盖初始化、模块顶层、UDF、Actions、Events、控制台以及 TypeScript 编译；编译缓存与常驻上下文也不能绕过。仍可查看/编辑源码、读取签名信息及保存文件。普通公式照常计算，原 UDF 缓存按现有规则标为受阻，不作为新鲜结果。生产安装自检如遇禁用策略会失败，不以跳过脚本检查冒充完整通过。

定向验证：`ScriptExecutionPolicyTests`、`DesktopSettingsMigrationTests` 以及 `scripts/test-no-scripts.ps1`。界面检查分别覆盖用户即时设置、参数首尾位置、进程环境、组合来源及已保存设置启动；用户/机器环境的独立 OR 合并通过注入各作用域取值覆盖，不改动机器真实环境。

### 模块角色

| 角色 | 绑定 | 发现规则 | 执行时机 |
| --- | --- | --- | --- |
| 主入口 | `initialization`，最多一个 | 可选导出的 `main(workbook, host)`，支持 async / Promise | 每次应用脚本，或打开工作簿后获准启用脚本时一次 |
| UDF | `formula-functions` | 导出的同步函数 | 公式求值 |
| Actions | `actions` | 必须导出 `main(workbook, host)`，支持 async / Promise，模块 ID 为操作名 | 用户执行操作 |
| Events | `events` | 导出的约定名称处理器 | 对应事件派发 |
| Common | 无绑定、不是根入口 | 按逻辑 ID `import` | 被所在上下文的模块依赖加载 |

v4 每个入口只绑定一个角色。UDF、Actions、Events 都是各自上下文的加载根，不需要主入口导入。模块顶层用于声明和依赖初始化；依赖主入口状态的工作簿操作应写在动作或事件函数体内。模块按 ES 模块规则加载，再运行主入口函数，最后才允许动作和事件执行。

Windows 与 macOS 沿用独立的 UDF 与 Actions/Events 持久隔离 worker（macOS 复用实现见 [67](67-avalonia-platforms.md)）。主入口只属于 Actions/Events；UDF 不加载或导入主入口、Actions、Events，Actions/Events 也不能导入 UDF 根模块；双方可分别实例化 Common。全局对象和模块闭包跨成功的调用与重算保留，但不跨工作簿、授权周期或重新打开保留。UDF 只从公式参数获取输入，不能读取主入口设置的状态。

授权仍是前提：打开未受信任工作簿不会直接执行源码；受信任文件自动启用时、或用户点击应用时，经过同一个初始化流程。每次重新应用会替换旧上下文并再次初始化。主入口的单元格、样式和文件操作沿用一个事务与一次撤销；异常、取消和提交拒绝不会留下部分写入。初始化过的上下文一旦丢失，脚本停用，需要重新应用，不会在下一次动作或重算中偷偷重放初始化。初始化入口不出现在操作列表。

Linux 的 worker 仍是一次性上下文，当前明确拒绝 `initialization` 绑定，不宣称支持此持久初始化模型。macOS 支持相同的持久初始化模型，继续使用 Seatbelt 隔离及逐操作/逐轮期限；控制台和编译器仍是一次性 worker。

## 共享初始化状态示例

Common 模块 `state`：

```ts
export const settings = { label: "", count: 0 };
```

主入口 `main`：

```ts
import { settings } from "state";
export function main(workbook: QuickExcel.Workbook, host: QuickExcel.Host): void {
    settings.label = "Ready";
    settings.count = 0;
}
```

Actions 模块 `increment`：

```ts
import { settings } from "state";
export function main(workbook: QuickExcel.Workbook): void {
    workbook.getActiveCell().setValue(++settings.count);
}
```

也可在主入口设置 `globalThis` 属性；TypeScript 需要正常声明其类型。主入口自己的模块级 `let` 仍遵循 ES 模块作用域，不会自动变成其他模块的全局变量。

## UDF 导出

```ts
// 模块 ID: prices；单元格公式：=JS.prices.DOUBLE(21)
export function DOUBLE(value: unknown): number {
    if (typeof value !== "number") throw new TypeError("Expected a number");
    return value * 2;
}
```

名称为 `JS.<模块名编码>.<导出名>`，大小写不敏感。模块名使用既有模板编码（非小写 ASCII 字母/数字转 `_XX`、数字开头加 `_`、长名称保留前缀和 SHA-256 短摘要）；函数导出名只能使用 ASCII 字母、数字和下划线，不能以数字开头。默认描述符采用 apiVersion 2、`any` 参数与标量返回，参数数量取函数的 JavaScript `length`。应使用固定必需参数；默认值和 rest 参数不会自动扩展公式的参数数量。需要明确区域、空值或函数别名时，可导出声明式 `functionDescriptors` 对象，以导出名为键提供既有 `QuickExcel.FunctionDescriptor`，没有调用注册函数的入口。未导出的辅助函数不会成为公式函数。异步或生成器函数不接受。

## Events 签名与模板

事件 `cell.click` 对应唯一约定名称 `onCellClick`：将点分段首字母大写后加 `on`。其余如 `workbook.beforeSave` 对应 `onWorkbookBeforeSave`。

```ts
export function onCellClick(context: QuickExcel.EventContext): void {
    const { workbook, host, event } = context;
    console.log(event.name, event.ranges);
}
```

必须导出、同步、恰好一个 context 参数，并且运行时不返回非空值。保留的 `onXxx` 名称拼错、签名错误或异步声明都会使加载失败。Common 或 Actions 中同名导出不会挂载。多个 Events 模块可处理同一事件，按入口顺序在一个事务中执行，失败全部回滚。现有 18 种事件、队列上限、关闭和保存/计算完成通知的只读规则保持不变。

选择 Events 文件后显示“添加事件处理器”。弹窗可按事件名、函数名、中英文说明或时机搜索，选中后显示当前界面语言的说明与触发时机。确认时在打开弹窗前的光标处插入 JS JSDoc 或 TS 类型模板。插入一次可整体撤销；取消无改动；原草稿版本、所属模块变化或输入法组合输入未完成时拒绝插入。插入位置由用户指定，需位于模块顶层；重复导出仍由编辑器和编译器诊断。

## 编辑器

深色文件树保留主入口，按 Actions / Common / Events / UDF 分类，显示 TS/JS 图标、缩进线及紧凑展开箭头。只读定义文件仍可导航。文件、编辑、视图、运行和签名操作集中到菜单；补全、参数提示、类型提示、定义跳转、下一问题和代码格式化通过菜单触发 Monaco 命令。事件按钮只在 Events 文件显示。保留 Monaco 每个文件的草稿、撤销、光标及滚动位置，保留 F12 控制台与原有草稿确认协议。

## 验证

### 3.2.0 编辑器运行入口与本机签名记忆（2026-11-05）

编辑区域右上角新增绿色 ▶，与 F5、运行菜单的“应用并重算”调用同一草稿确认、TypeScript 检查、信任验证和隔离 worker 流程。原生窗口和 Monaco 内均支持 F5；浏览器捕获该键以防刷新页面，修饰键、自动重复和组合输入不发起应用。F12 仍在脚本工作区切换控制台，主表的 F5/F12 分别为定位/另存为。

签名密码窗口新增独立的“记住上次的私钥文件”和“记住密码”选项，初始关闭。成功签名后写入本机用户数据目录的 `script-signing-preferences.json`；再次打开默认可选择上次文件，密码只对 SHA-256 相同的证书文件回填。取消对话框和签名失败不更新记忆，取消勾选并完成签名会移除对应数据。私钥文件不复制，路径和密码均不进入 XLSQ；Windows 通过当前用户 DPAPI 保护整条记录，并绑定本地设置路径，其他平台复用已有系统密钥服务保护，无明文降级。保护服务不可用时仍可手动签名，但提示记忆保存失败。

签名/保护/本地化定向测试 46 项通过（`artifacts/tests/signing/signing.trx`），覆盖实际 DPAPI、篡改、复制到不同设置路径、两项独立选择、证书内容变化和保护失败不覆盖旧设置。Commands 20 项覆盖签名窗口、自动填充、取消记忆及逐项检查 XLSQ 不携带密码/私钥路径。实际 Monaco 模块专项 40 项通过（`artifacts/avalonia/web-tests/20261105-123749-5b960688`），含 ▶、F5 防刷新、组合输入/重复/修饰键门禁。未运行全量回归。

3.2.1 修正运行按钮图标偏上：用 12×14 的原生三角形路径替换文本字符，显式设置水平和垂直内容居中，填充跟随按钮前景色。保留按钮尺寸、提示、无障碍名称及执行入口。Release 构建无警告/错误，Monaco 模块专项 40 项通过（`artifacts/avalonia/web-tests/20261105-124524-c955d636`）；已复核该目录 `script-run-button.png` 的放大渲染，三角形边界中心与按钮中心一致。未运行全量回归。

3.2.1 已通过 `build-quickexcel.ps1 -SkipTests -RunSmoke -CreateArchive` 发布到 `artifacts/avalonia/releases/20261105-124618-946acb6d/application`，28 项生产自检通过。活动指针已切换，清单、宿主 EXE/DLL、Web 清单及 ZIP SHA-256 均核验一致，默认启动器使用新包；3.2.0 包保留。ZIP 为 `QuickExcel-Avalonia-3.2.1-win-x64-20261105-124618-946acb6d.zip`，SHA-256 为 `157439CDD0CF9EA8B823130D45500E6546296EB5AA4E831F39D4691875C2A17B`。

相关测试与界面检查记录在 `artifacts/script-v4-*.log` 和 `artifacts/avalonia/web-tests/`。覆盖初始化/重算/重新应用/重开、独立 UDF、声明式事件发现、错误签名、事务回滚、失败后停用、角色保存、JS/TS 模板、光标插入及失效防护、一次撤销、中英文搜索说明与菜单导航。只运行脚本、worker、编译器、模块存储和编辑器相关检查；未运行完整回归。界面检查不替代物理输入法或无障碍验证。

最终记录：运行时/会话/事件/控制台/签名定向组 164 项通过；编译器/模板/模块/隔离定向组 104 项通过（两组有重叠）。原生 JSC smoke、ApiGen 确定性生成检查通过。完整脚本编辑器检查 172 项通过，证据为 `artifacts/avalonia/web-tests/20261105-102433-c55ac6f2/`。

已执行 `scripts/build-quickexcel.ps1 -SkipTests -RunSmoke`。2.0.0 发布包为 `artifacts/avalonia/releases/20261105-102633-9586aa11/application`；生产自检 22 项全部通过，其中包含 v4 初始化次数、自动发现、状态隔离和角色保存重开。`active-windows.json` 已切换到此包，清单 SHA-256 与启动器验证值一致；生产 Web 包的 test/probe 标志均为 false。上一活动包保留。
