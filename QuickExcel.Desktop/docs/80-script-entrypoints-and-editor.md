# 初始化主入口、自动发现与脚本编辑器

2026-11-05，产品 2.0.0。新工作簿使用文件 `scriptApiVersion: 4`，IPC 11。对象模型及能力清单沿用 79 的实现边界；本章替换其入口与注册规则。API v3 文件保持原版本执行与保存，不自动改写源码；旧 v1/v2 的执行和保存门禁不变。v3 的注册方法只存在于版本隔离的兼容路径，v4 的 `host` 没有 `functions`、`events`、`actions` 或注册方法。

## 角色与加载

| 角色 | 绑定 | 发现规则 | 执行时机 |
| --- | --- | --- | --- |
| 主入口 | `initialization`，最多一个 | 可选导出的同步 `main(workbook, host)` | 每次应用脚本，或打开工作簿后获准启用脚本时一次 |
| UDF | `formula-functions` | 导出的同步函数 | 公式求值 |
| Actions | `actions` | 必须导出同步 `main(workbook, host)`，模块 ID 为操作名 | 用户执行操作 |
| Events | `events` | 导出的约定名称处理器 | 对应事件派发 |
| Common | 无绑定、不是根入口 | 按逻辑 ID `import` | 被所在上下文的模块依赖加载 |

v4 每个入口只绑定一个角色。UDF、Actions、Events 都是各自上下文的加载根，不需要主入口导入。模块顶层用于声明和依赖初始化；依赖主入口状态的工作簿操作应写在动作或事件函数体内。模块按 ES 模块规则加载，再运行主入口函数，最后才允许动作和事件执行。

Windows 沿用独立的 UDF 与 Actions/Events 持久隔离 worker。主入口只属于 Actions/Events；UDF 不加载或导入主入口、Actions、Events，Actions/Events 也不能导入 UDF 根模块；双方可分别实例化 Common。全局对象和模块闭包跨成功的调用与重算保留，但不跨工作簿、授权周期或重新打开保留。UDF 只从公式参数获取输入，不能读取主入口设置的状态。

授权仍是前提：打开未受信任工作簿不会直接执行源码；受信任文件自动启用时、或用户点击应用时，经过同一个初始化流程。每次重新应用会替换旧上下文并再次初始化。主入口的单元格、样式和文件操作沿用一个事务与一次撤销；异常、取消和提交拒绝不会留下部分写入。初始化过的上下文一旦丢失，脚本停用，需要重新应用，不会在下一次动作或重算中偷偷重放初始化。初始化入口不出现在操作列表。

Linux 的 worker 仍是一次性上下文，当前明确拒绝 `initialization` 绑定，不宣称支持此持久初始化模型；macOS 仍不执行脚本。沙箱、资源限额和原生桥保持不变。

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

相关测试与界面检查记录在 `artifacts/script-v4-*.log` 和 `artifacts/avalonia/web-tests/`。覆盖初始化/重算/重新应用/重开、独立 UDF、声明式事件发现、错误签名、事务回滚、失败后停用、角色保存、JS/TS 模板、光标插入及失效防护、一次撤销、中英文搜索说明与菜单导航。只运行脚本、worker、编译器、模块存储和编辑器相关检查；未运行完整回归。界面检查不替代物理输入法或无障碍验证。

最终记录：运行时/会话/事件/控制台/签名定向组 164 项通过；编译器/模板/模块/隔离定向组 104 项通过（两组有重叠）。原生 JSC smoke、ApiGen 确定性生成检查通过。完整脚本编辑器检查 172 项通过，证据为 `artifacts/avalonia/web-tests/20261105-102433-c55ac6f2/`。

已执行 `scripts/build-quickexcel.ps1 -SkipTests -RunSmoke`。2.0.0 发布包为 `artifacts/avalonia/releases/20261105-102633-9586aa11/application`；生产自检 22 项全部通过，其中包含 v4 初始化次数、自动发现、状态隔离和角色保存重开。`active-windows.json` 已切换到此包，清单 SHA-256 与启动器验证值一致；生产 Web 包的 test/probe 标志均为 false。上一活动包保留。
