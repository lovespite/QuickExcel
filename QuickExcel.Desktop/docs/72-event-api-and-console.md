# 事件名称、属性处理器与调试控制台

2026-11-03。在现有事件与隔离脚本执行基础上增加名称枚举、属性式处理器和交互控制台。只映射已有的 18 个事件，不新增保存事件。`$` 与 `quickExcel` 继续指向同一个 API 对象。

## 事件 API

`QxEventNames` 是 JS 可用的冻结名称对象，在 TypeScript 中声明为字符串枚举。原有字符串和 `.on(...)` 注册继续有效，枚举与属性均提供编辑器补全和类型检查。

```typescript
$.events.on(QxEventNames.ON_CELL_COMMIT, ({ event }) => {
  console.log(event.name, event.ranges);
});

$.events.onWorksheetActivate = (sheet, { snapshot, patch }) => {
  console.log('当前工作表', sheet.id, sheet.name);
  patch.setValue(sheet.id, 'A1', snapshot.baseRevision);
};

// 再次赋值替换该属性的处理器，null 移除；不影响 .on() 注册的监听器。
$.events.onWorksheetActivate = null;
```

| 枚举成员 | 原有事件 | 属性处理器 |
| --- | --- | --- |
| `ON_OPEN` | `workbook.open` | `onWorkbookOpen` / `onOpen` |
| `ON_BEFORE_CLOSE` | `workbook.beforeClose` | `onWorkbookBeforeClose` / `onBeforeClose` |
| `ON_CLOSE` | `workbook.close` | `onWorkbookClose` / `onClose` |
| `ON_ACTIVATE` | `workbook.activate` | `onWorkbookActivate` / `onActivate` |
| `ON_DEACTIVATE` | `workbook.deactivate` | `onWorkbookDeactivate` / `onDeactivate` |
| `ON_BEFORE_SAVE` | `workbook.beforeSave` | `onWorkbookBeforeSave` / `onBeforeSave` |
| `ON_SAVED` | `workbook.saved` | `onWorkbookSaved` / `onSaved` |
| `ON_CALCULATED` | `workbook.calculated` | `onWorkbookCalculated` / `onCalculated` |
| `ON_WORKSHEET_ADDED` | `sheet.added` | `onWorksheetAdded` |
| `ON_WORKSHEET_DELETED` | `sheet.deleted` | `onWorksheetDeleted` |
| `ON_WORKSHEET_RENAMED` | `sheet.renamed` | `onWorksheetRenamed` |
| `ON_WORKSHEET_ACTIVATE` | `sheet.activate` | `onWorksheetActivate` |
| `ON_WORKSHEET_DEACTIVATE` | `sheet.deactivate` | `onWorksheetDeactivate` |
| `ON_SELECTION_CHANGED` | `selection.changed` | `onSelectionChanged` |
| `ON_CELL_CLICK` | `cell.click` | `onCellClick` |
| `ON_CELL_EDIT` | `cell.edit` | `onCellEdit` |
| `ON_CELL_COMMIT` | `cell.commit` | `onCellCommit` |
| `ON_CELLS_CHANGED` | `cells.changed` | `onCellsChanged` |

工作表属性处理器接收 `(sheet, context)`，`sheet` 为冻结的 `{ id, name }`，删除通知保留删除前的身份。其他属性处理器接收 `(event, context)`；第二个参数始终为原有 `snapshot / patch / fs / event` 上下文。`.on(name, handler)` 的单一上下文参数保持不变。

同一事件的长短属性名称共享一个处理器槽。首次赋值按注册时序加入监听列表，替换不改变位置，移除后重新赋值加入末尾。属性处理器计入每事件最多 32 个处理器的限额，与 `.on` 在同一事务中按顺序执行。仅允许初始化阶段注册函数或移除为 null；非法值、运行中重新赋值和异步返回仍拒绝。仍需 API v2 的 actions 绑定及现有脚本授权；保存/计算完成及关闭通知的写入限制见 [69 号契约](69-script-events-and-virtual-files.md)。

## 调试控制台

入口为“脚本 → 调试控制台”，也可以在脚本编辑器底部选择“控制台”。

- `Ctrl+Enter`（macOS 为 Cmd+Enter）执行；普通 Enter 换行。
- `Alt+↑/↓` 浏览最多 50 条会话命令历史，返回最新位置恢复未执行输入。
- 输出采用 Chrome DevTools 风格的紧凑工具栏、等宽日志行、时间戳和底部 `›` 输入区。日志列表按需创建可见行，支持选择文字、复制筛选结果和清空。
- `console.log/info/warn/error/debug` 均接受多个参数和对象。`warn` 使用黄色背景与警告图标，`error` 和执行异常使用红色背景与错误图标；`console.error()` 只输出消息，不抛出异常或中止执行。编辑器补全与悬浮提示包含这些 API 的说明。
- 可按日志/信息、警告、错误、调试级别筛选，也可组合不区分大小写的文本过滤。工具栏显示可见条数/总条数及错误、警告总数；筛选和清空不改变工作簿。
- “停止”或已有 Esc 取消按钮终止当前工作，不提交部分修改。

控制台执行独立 JavaScript，可直接使用 `snapshot`、`patch`、`fs`、`event`（此处为 null）、`$`、`quickExcel` 和 `QxEventNames`：

```javascript
1 + 2
snapshot.sheets
snapshot.getCell(snapshot.sheets[0].id, 'A1')
console.log('值', { answer: 42 })
console.warn('缺少可选字段', { cell: 'A1' })
console.error('输入格式错误')
patch.setValue(snapshot.sheets[0].id, 'A1', 42)
fs.writeText('//mem/debug', '跨命令保留的文本')
```

每次命令使用新的隔离动作 worker，不自动加载或授权工作簿模块，不读取其局部变量；变量不跨命令保留，需要临时状态时使用 `//mem/`。控制台当前不接受 TypeScript 注解或 ESM 导入，不提供断点、单步或调用栈暂停。

显式执行当前输入是一轮独立请求，不改变工作簿的脚本授权。源码以 JSON 字符串编码后，在隔离 worker 的动作上下文内执行；宿主及浏览器不 eval 用户代码。快照、容量、AppContainer / Landlock、截止时间、Job/scope、代次、取消和提交校验沿用原动作路径，没有同权限或进程内回退。macOS 仍禁用执行。

单元格与持久文件 patch 完整校验后作为一个可撤销事务提交；失败、停止、超时或修订过期时整体丢弃。只读公式工作簿仅允许检查，不接受单元格或文件写入。控制台不允许改写脚本源码/绑定/签名，以避免绕过独立源码历史；这些内容仍使用脚本编辑器修改。系统文件/保留内容与保护、验证等原守卫继续生效。

日志按每轮最多 100 条、每条 4000 UTF-8 字节返回；对象以有界深度与条目数显示，循环引用标为 `[Circular]`，普通 getter 不求值。日志在轮次结束后交付，并非逐条实时流；普通动作异常保留之前输出，超时/无效上下文不再调用引擎取日志。单会话最多保留 500 条输入/输出，每条最多 4000 UTF-16 单元；控制台输入上限 64 KiB UTF-8。历史和输出不保存进 XLSQ，不标脏，切换工作簿会话不互相泄漏。

级别显示复用现有 worker 的 `[warn]` / `[error]` 等日志前缀，保留原始文本供复制和诊断，并识别公式日志的 `JavaScript：` 上下文。宿主拒绝及运行失败单独标为错误。本次 UI 修改不变更 IPC；由于既有传输格式为文本，普通 `console.log('[warn] ...')` 也会按该前缀显示。

## 协议与验证

为保留异常前输出，`WorkerFailure` 增加有界 `Logs`，IPC 升至 **7**；宿主和 worker 必须同包更新，拒绝混用 IPC 1–6。没有增加 worker 执行能力；即时代码复用动作消息。历史发布包及当前默认启动器指针不原地改写。

本机 Windows 定向证据：

- 真实 WebView2 / Avalonia 脚本专项 **71 项通过**：`artifacts/avalonia/web-tests/20261103-110232-27dcdf4c/result.json`。覆盖真实枚举/属性补全、类型推断、日志/结果、事务撤销、普通异常输出、停止、历史、清空、会话切换、只读检查以及合成 preedit 门禁；原有脚本编辑/保存/恢复流程同时通过。`-ScriptsOnly` 未运行无关 PDF 项。
- 事件、虚拟文件及 IPC 共 **80 项完成验证**：`artifacts/script-console-tests/events-and-protocol.trx` 首轮 79 项通过；一个新断言把 ImmutableArray 的身份比较误作序列比较，改为序列后 `all-event-mappings.trx` 单项通过，实际调用并验证全部 18 个事件。未修改产品规则来适配断言。
- 控制台、原动作、会话及 worker **114 项通过**：`artifacts/script-console-tests/console-action-session-worker.trx`。覆盖未授权模块不执行、快照读取、原子文件/单元格修改、撤销重做、普通异常、超时、过期/取消、原日志及输出限额。
- 新增错误日志边界 **3 项通过**：`artifacts/script-console-tests/console-log-limits.trx`。日志预算用尽后仍显示宿主拒绝原因；畸形 worker 的超条目数/超 UTF-8 容量失败日志在呈现前拒绝。

以上共 197 项不同的 Core 定向用例通过，未运行全仓回归或性能基准。正式 Web 资源已重建，测试辅助代码不进入正式 bundle。最终解决方案 Release 构建记录在 `artifacts/script-console-tests/final-build.log`。默认发布包未切换；实机 IME、DPI、辅助技术和跨平台认证范围保持既有状态，合成 preedit 检查不代表物理 IME 验收。

### 控制台样式增量验证（2026-11-03）

- 控制台定向 **19 项通过**：`artifacts/script-console-tests/console-styling.trx`，验证所有日志 API、级别与原始文本、公式上下文、异常、输出上限和原有撤销/隔离规则。
- Avalonia / WebView2 脚本专项 **78 项通过**：`artifacts/avalonia/web-tests/20261103-111940-33e5ba9a/result.json`。新增检查实际日志行的前景/背景色、文字选择、级别与文本过滤、复制按钮、计数、空状态、760 px 最小窗口布局；原执行、停止、历史、清空、只读与会话隔离检查继续通过。
- 同目录 `console-styled.png` / `console-styled-narrow.png` 为实际控件渲染截图，已检查常规与最小窗口宽度的布局。Release 解决方案构建 **0 警告、0 错误**：`artifacts/script-console-tests/console-styling-build.log`。正式 Web 资源已恢复；本次未运行全仓回归，未切换默认发布包。
