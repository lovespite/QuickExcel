# 实时工作簿对象与每工作簿脚本上下文

2026-11-04。此文更新 11/13/69/72 中关于逐次创建执行 worker、预先复制全量单元格与文件快照的运行机制。文件 `scriptApiVersion` 仍为 2；初版 IPC 为 8；当前为 9，UDF 独立上下文与事务内计算见 [77](77-udf-context-isolation-and-transaction-calculation.md)，桥标识为 `JavaScriptCore/2.54.2/b95c91b1/qx-sync-2`，旧 host/worker/native 混用会被拒绝。

## 对象 API

```typescript
function main(workbook: ExcelScript.Workbook) {
    const selectedCell = $.getActiveCell();
    const selectedSheet = $.getActiveWorksheet();
    const selectedRange = $.getActiveRange();
    selectedCell.getFormat().getFill().setColor("#ff0000");
    selectedRange.getFormat().getFill().setColor("#00ff00");
    console.info(selectedSheet.getCell("A1").getValue());
}
```

`workbook === $ === quickExcel`。具备 actions 绑定的入口模块中，顶层 `main` 自动成为可执行动作；单入口名为 `main`，多入口名为 `<moduleId>.main`。仍需显式授权脚本，仍支持 `$.actions.register` 和事件处理器。TypeScript 使用内置的 `ExcelScript` 子集声明，不表示已实现完整 Microsoft Office Scripts API。

| 对象 | 当前支持 |
| --- | --- |
| Workbook / `$` | `getActiveCell()`、`getActiveWorksheet()`、`getActiveRange()`、`getActiveRanges()`、`getWorksheet(nameOrId)`、`getWorksheets()` |
| Worksheet | `getId()`、`getName()`、`getCell("A1")`、`getCell(row, column)`（从 0 开始）、`getRange("A1:B2")` |
| Range | `getAddress()`（不含工作表名前缀）、`getWorksheet()`、`getValue()`、`getValues()`、`setValue()`、`setValues()`、`getFormula()`、`setFormula()`、`setFormulas()`、`getFormat()` |
| RangeFormat / RangeFill | `getFill()`；`getColor()`、`setColor("#RRGGBB")`、`clear()` |

选区来自本轮开始时的 Avalonia 实际选区；多区域选择必须用 `getActiveRanges()`，单区域方法会报错，避免无意覆盖中间单元格。`getValue()` 读取区域左上角；`setValue()` 填满区域；矩阵写入要求尺寸完全一致。地址支持矩形 A1 表示法，暂不支持名称、整行整列或带工作表前缀地址。

## 事务语义

用户确认的语义是：同一次执行内立即读到已写入的字面值、公式文本和填充样式；整个执行成功后一次提交，一次撤销；失败时所有单元格、样式和虚拟文件修改均不提交。重叠写入按调用顺序生效。非法写入即使被 JavaScript `catch` 捕获，也会使整次事务失败。

对象只是 worker 内的轻量句柄。读写通过同步 IPC 向宿主拥有的 `ScriptWorkbookModel` 请求；宿主在独立的命令草稿上操作。事务内部保留不可变领域模型以支持隔离、校验及撤销，不再向 worker 预先复制整本工作簿。提交继续检查授权、输入版本、活动请求、只读、保护、富文本、数组溢出、数据验证和现有附件限制；不解除合并区域/单元格内图片的值写入限制。脚本执行期间出现用户编辑或新请求，旧草稿失效。

当前 `getValue/getValues` 会按需对事务草稿计算并读取新结果，覆盖内置公式和已授权 UDF；计算结果仅保留在事务内，正式提交后仍走正常计算轮。不可用函数及受阻依赖的旧缓存不能冒充新结果。`getFormula` 只读取公式文本，不触发计算。详细契约与边界见 [77](77-udf-context-isolation-and-transaction-calculation.md)。

旧 `snapshot` 方法保留本轮开始时的只读语义，旧 `patch` 方法仍可用；其写入会进入同一个事务，新对象可立即读到。`fs` 首次访问时才生成逻辑文件视图，`//xlsq/`、`//temp/`、`//mem/` 仍不是 OS 路径。原始 `workbook.xlsx` 替换与单元格修改、源码/绑定/签名重写与数据修改继续禁止混合提交。完成通知的只读规则不变。

## 上下文及资源边界

Windows API v2 在同一授权工作簿会话内分别复用 UDF worker/context 与 actions/events worker/context；两者均为零能力 AppContainer。各自只初始化对应绑定入口的可达图，各自的模块变量和全局对象互不相通；同一模块作为两边的依赖时分别实例化。两种角色均跨调用、跨轮次复用，不逐 UDF 调用重建。不同工作簿完全隔离。注册许可依据原生模块来源和可达绑定角色检查；worker 角色在其整个生命周期内固定。每个 worker 内串行，动作读取公式时宿主可调用另一个 UDF worker，不重入动作的 JSC。

模块变量与闭包在成功执行之间保留。撤销只撤销工作簿事务，不倒退 JavaScript 堆。动作/事件异常、取消、超时、过期提交会丢弃上下文；下一次显式调用重新初始化模块，不重放失败事件，也不声称能够恢复任意 JavaScript 堆的失败前状态。动作仅修改自己的 JS 状态不再使公式缓存失效。宿主拒绝最终提交时保守丢弃执行上下文并使公式缓存失效，避免复用失败事务中的 UDF 状态。普通 UDF 异常仍遵守既有单元格错误语义。关闭、禁用、更改源码/绑定/签名、重新授权均释放或替换 worker。

控制台继续使用独立的一轮执行，不隐式加载已存储模块；可使用同样的对象 API。API v1 保留旧语义。Linux 保留强制五秒 OS 生命周期和逐轮 worker，本次没有放松 systemd/Landlock/seccomp 要求，也没有验证新的 Linux 原生构建。macOS 脚本仍禁用。

IPC 8 引入、IPC 9 延续每轮 `RoundId`、`loadWorkbook/beginRound/endRound` 与嵌套 `model` 请求。宿主核对会话、worker、轮次、请求、操作 ID 和严格递增序号；模型方法为白名单。单轮最多 4096 次模型请求、累计 100000 格读取、累计 20000 格写入（重复计数），模型请求与响应合计不超过 16 MiB。原生单操作一秒、JS 栈 256 KiB、执行进程 256 MiB 和活动轮五秒仍有效；Windows 空闲间隔不消耗活动轮时限。失败进程在 Job 中终止，所有权 lease 保持到验证退出与清理完成；不新增同权限或进程内回退。JSC-MEM-001 仍未实现，状态不变。

## 初版性能原因与验证（IPC 8 历史记录）

原发布包的空事件热样本中位数约 400 ms，主要开销是每轮创建隔离沙箱/复制约 147 MiB worker 载荷、CLR/JSC 启动、模块初始化和清理。之前仅文件视图生成，在 10 万数值单元格样本中就约 285 ms。新路径把启动成本移到上下文首次加载，把单元格和文件读取改为按需执行。

2026-11-04 Windows Release 开发构建，使用真实 `DocumentSession → ScriptWorkerActionExecutor → AppContainer worker → CommitAction` 路径，每组首次调用后再采集 30 次；仅单元格数为 0 的组保留 A1=0 用于校验读取。没有模拟 worker、关闭隔离或预先执行完整事件。

| 请求 | 工作簿 | 热执行中位数 | P95 |
| --- | --- | ---: | ---: |
| `onActivate` 打印工作表名 | 极小工作簿 | 0.93 ms | 1.79 ms |
| 同上 | 100000 个数值单元格 | 4.03 ms | 5.66 ms |
| 示例 `main`，活动单元格/区域填色并读 A1 | 极小工作簿 | 3.32 ms | 7.24 ms |
| 同上 | 100000 个数值单元格 | 77.06 ms | 282.70 ms |

事件首次调用分别约 432/230 ms，说明冷启动成本仍在。样例使用单格选区；大工作簿格式命令仍受宿主模型校验、不可变结构和历史成本影响。这些数字包含请求创建、同步 IPC、脚本执行、提交和结束轮次，不含 UI 的 100 ms 事件定时器等待、布局绘制或公式重算，也不是跨设备性能保证。原始数据为 `live-model-raw.jsonl`，汇总为 `live-model-summary.json`。

回归包含实际 JSC/隔离 worker 下的 JavaScript/TypeScript 示例、选区、顺序覆盖、读己写、一次撤销、失败回滚、非法写入被捕获、超时、过期请求、跨轮状态、五秒以上空闲、工作簿隔离、角色伪造、IPC 身份/顺序/非法值故障。桌面 Automation 检查包含真正的动作下拉框、事件派发和编辑器入口；它不代表物理输入、完整性能或跨平台质量门已关闭。

原始测量与构建/测试日志位于忽略的 `artifacts/diagnostics/event-latency-20261104/`；桌面 76 项通过证据为 `artifacts/avalonia/workspace/20261104-012902-3888e8f2/`。

验证结果：解决方案 Release 构建通过；动作/事件/文件/控制台/对象与共享上下文 144 项定向用例全部通过（其中新增未绑定模块测试的构造错误修正后单独重跑，共享上下文及跨角色调用 6 项通过）；相关 worker/IPC/批处理/原生桥/TypeScript 验证 80 项通过，JSC-MEM-001 按原计划跳过；原生 smoke 通过。没有运行全量回归。

补充验证：最终提交被宿主拒绝后的公式缓存失效、过期事务、事件文件与恶意模型 IPC 共 60 项通过，记录在 `live-model-commit-rejection-tests.log`。

最终 Windows 包为 `artifacts/avalonia/releases/20261104-013931-a5410f2f/application`，已通过发布自检并激活默认启动器，旧包保留。使用此包的宿主程序集及自包含 worker 再测 30 个热样本：简单事件中位数 **0.78 ms** / P95 **2.20 ms**；10 万单元格事件 **4.64 ms** / **8.99 ms**；示例填色与读取分别 **2.54 ms** / **4.17 ms** 和 **83.00 ms** / **261.31 ms**。首次事件约 545/336 ms，仍包含冷启动成本。范围与上述测量相同，不含 UI 排队、绘制或重算。数据见 `delivery-model-raw.jsonl` / `delivery-model-summary.json`；发布与自检日志见 `live-model-delivery.log` 及该 release 的 `evidence/`。
