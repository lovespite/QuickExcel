# QuickExcel 0.6 交互预览版

最新 `0.6.0-preview.2` 已完成 UI-004/UI-005 当前范围，测试与限制见 [结构与工作区](docs/18-m5-structure-and-workspace.md)。结构菜单操作选区所在整行/整列；排序选不含表头的矩形，按当前列；筛选选含表头的矩形。B2 冻结第一行/列，顶部最多 10 行为可点击定位的只读参照区。格式菜单修改样式，可一次撤销，原始输入不变。

文件菜单提供最近 10 文件、清空记录、关闭工作簿（Ctrl+W）和脚本面板开关。保存发现外部修改/删除时保留编辑并阻止覆盖，请另存新路径或重新打开。筛选/冻结仅本次工作区有效，重开复位；筛选下复制可见行、粘贴跨隐藏行拒绝、Delete 只清可见格，排序前先清除筛选。

2026-10-28 M5 首切片：UI-002 / UI-003 本切片完成，按需行对象、可见行刷新和统一键盘/选区/粘贴规则已接入；本机中文拼音、编辑器焦点及保存重开验收通过。具体实现、测试证据、性能数据与 IME 验收边界见 [M5 记录](docs/17-m5-grid-and-input.md)。

2026-10-28：已补齐宿主崩溃后的 profile/payload 回收，625 项测试通过。本机部署矩阵五项通过；直接长路径 EXE 启动失败仍待处理，详见 [回收与部署验证](docs/16-m4-recovery-and-deployment.md)。

本版保留 M2/M3 文件与计算及 M4 worker/模块/actions 能力，开始推进 M5 桌面交互。需要 Windows x64 与 .NET 10 Windows Desktop Runtime；保留完整发布目录后运行 QuickExcel.exe。

最新切片已把无能力 AppContainer 接入公式与 actions 的实际 worker，并补齐回环/DNS 的本机 WFP 拒绝证据。详见 [隔离 ADR](docs/15-m4-os-isolation-adr.md)和 [合批及容量](docs/14-m4-batching.md)；部署矩阵、严格策略及 M6 OS 门禁仍未完成，仅供可信脚本使用。

## 使用

1. 打开或新建工作簿。打开不执行源码，已有公式缓存可能过期。
2. 双击单元格或在公式栏编辑。前置单引号强制文本，`00123` 保留文本，单独输入单引号表示空文本。
3. Ctrl+S 保存 XLSQ v2；活动单元格与公式栏会先提交。输入无效阻止保存；重算失败仍可保存有效输入与旧缓存。
4. v1 文件迁移后显示未保存状态，首次保存选择其他文件名，旧文件保留。
5. 含未知部件的文件会显示兼容报告。未修改工作簿内容时可以原样保留；编辑后可取消，或明确另存移除这些内容的副本。重新打开副本可刷新兼容报告。
6. “导出值快照”将公式全部转为结果；“导出公式”冻结脚本及传递依赖结果，保留独立标准公式。需要冻结的结果缺失时拒绝，过期结果须明确选择。
7. 只在信任当前源码时点击“应用脚本并重算”（Ctrl+Enter）。源码改变会立即撤销授权。文件级 API v2 可执行包内模块；编辑器显示首入口，辅助模块保留。API v1 多模块包须显式迁移到 v2 才能执行。
8. 包含 actions 绑定的文件应用后显示动作列表。选择并运行，全部写值/公式通过校验后一次提交，Ctrl+Z 一次撤销，重做不再次运行动作。动作执行期间可停止；冲突或失败后须重新运行。动作提交后公式缓存等待重算。

## 编辑和快捷键

Ctrl+O 打开，Ctrl+N 新建，Ctrl+Shift+S 另存。首次单击单元格只选中，再次单击同一格进入编辑；选中时显示计算结果和绿色边框，编辑时显示原始输入与光标。也可按 F2 编辑，或直接键入替换。Enter/Shift+Enter 提交后向下/上，Tab/Shift+Tab 提交后向右/左，Escape 取消草稿；Shift+方向键扩展矩形选区。输入法组合期间请先确认或取消候选，再保存或切换。

Ctrl+C 复制矩形显示结果，非连续选区拒绝复制。Ctrl+V 粘贴值，`=` 开头保留为文本；Ctrl+Shift+V 显式按输入解析公式。单值填充选区，矩阵需选左上角或同尺寸矩形；支持引号包裹的多行/Tab 文本。Delete 清空实际选中位置。一次粘贴/清空可整体撤销。

Ctrl+Z/Y 在文本编辑中作用于当前文本框，退出编辑后作用于工作簿事务。默认历史预算 32 MiB，截断时显示提示。键盘扩选和 Ctrl+A 最多 20,000 格；复制最多 100,000 格、4 Mi 个字符。导入的隐藏行列属于逻辑矩形，复制/粘贴仍包含它们。

一次粘贴最多 20,000 格，任一目标非法则整批不写。界面最多显示 100,000 行 × 256 列；模型最多 64 表、1,024 列、全簿 1,000,000 存储格，文件还须满足解压与解析预算。远端空洞不展开为密集矩形。

## 文件与脚本

v2 保存唯一的 workbook.xlsx，以及身份/修订状态和脚本源码模块映射。支持数字、布尔、标准错误、文本、空白/空文本、公式缓存、1900/1904 日期系统及基础格式。历史 v1 中已经丢失的类型不猜测恢复。

```javascript
quickExcel.functions.register("DISCOUNT", (price, rate) => price * (1 - rate));
quickExcel.functions.register("DOUBLE", value => value * 2);
```

支持同步标量、v1 一维区域与 v2 显式二维区域参数；没有 Promise、任意对象或动态数组溢出支持。QuickJS 每轮运行于 AppContainer worker，执行期间可点击状态栏「停止计算」。文件级 API v2 支持按清单逻辑 ID 导入 ES 模块、多个入口和共享依赖；初始化必须同步完成。actions 使用独立上下文、只读快照和有界 patch，首版每次最多写 20,000 格；含合并或单元格内图片的工作表暂不接受动作写入。打开、迁移、另存均不会复制运行授权。OS 最终发布验收仍待完成，继续仅供可信脚本使用。发布包的 `ScriptWorker` 目录必须保留；见 [actions 示例与边界](docs/13-m4-actions.md)、[模块契约](docs/12-m4-modules.md)及 [worker 支持范围](docs/11-m4-worker.md)。

跨表计算已接入；未授权脚本和依赖仍保留 Blocked 缓存。主题、富文本、图表/透视表、复杂锚点、单元格内图片与数组公式等不在可编辑保真承诺内；完整矩阵见 [M2 文件与类型](docs/09-m2-files-and-types.md)。图片/合并完整界面呈现、排序筛选、恢复和无障碍等留待后续。

## 构建与验收

```powershell
dotnet build .\QuickExcel.slnx -c Release --artifacts-path .\artifacts\build -v minimal
dotnet test .\QuickExcel.Core.Test\QuickExcel.Core.Test.csproj -c Release --artifacts-path .\artifacts\build -v minimal
.\scripts\build-quickexcel.ps1 -RunSmoke -CreateArchive
.\scripts\test-m5-grid.ps1 -PackageDirectory artifacts\QuickExcel -Mode All
```

从仓库根目录执行。所有引用项目输出也留在本仓库 artifacts。发布脚本执行 Core/native/WPF 检查并生成 ZIP、SHA-256 和 BUILD-METADATA；未声称无 SDK 干净机器、签名或安装器验收已完成。详见[发布说明](RELEASE-NOTES.md)。

计算范围与函数列表见 [M3 支持矩阵](docs/10-m3-calculation.md)。新计算器拒绝脚本覆盖内置名称；请使用 JS 命名空间。结构命令 API 已提供，完整界面操作继续 M5。
