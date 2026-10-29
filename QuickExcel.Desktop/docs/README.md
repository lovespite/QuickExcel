# QuickExcel 架构与实施文档

最新 `0.6.0-preview.3` 请先读 [M5 正式界面与桌面流程](19-m5-product-workspace.md)：启动、导航、脚本专页、查找替换、基础图片/合并，以及验证和未关闭门禁。下述原型切片记录保留为历史证据。

设计基线：2026-10-22；最新阶段：2026-10-28 的 **M5 结构与工作区切片**。先读 [UI-004 / UI-005 实现与验收](18-m5-structure-and-workspace.md)、[UI-002 / UI-003](17-m5-grid-and-input.md)、[M4 回收与部署](16-m4-recovery-and-deployment.md)及[隔离 ADR](15-m4-os-isolation-adr.md)。M4 未关闭项、M5 后续交互和 M6 Beta 门禁继续按计划推进。

当前使用 XLSQ v2、类型化跨表计算及独立 AppContainer worker。Core 保持纯 .NET；ExcelIO 与脚本适配位于 Infrastructure。下文明确标记 0.2 的内容保留为历史。

## 阅读顺序

| 文档 | 解决的问题 |
| --- | --- |
| [M5 结构与工作区](18-m5-structure-and-workspace.md) | 结构、排序引用、筛选/冻结、基础样式、最近文件与外部冲突 |
| [M5 网格与输入](17-m5-grid-and-input.md) | 网格适配、虚拟化测量、键盘/选区/剪贴板、IME 验收边界 |
| [M4 回收与部署](16-m4-recovery-and-deployment.md) | 宿主崩溃回收、部署证据与未关闭项 |
| [M4 包内模块](12-m4-modules.md) | API v2、逻辑 ID 导入、同步初始化、版本兼容与模块验证 |
| [M4 actions 操作脚本](13-m4-actions.md) | 只读快照、二次校验、原子提交、一次撤销、桌面运行与验证 |
| [M4 独立脚本进程](11-m4-worker.md) | worker 首切片的取消、故障和资源边界；保留历史发布证据 |
| [产品范围](01-product-and-scope.md) | 长期 Beta 的产品边界和暂缓功能 |
| [系统架构](02-system-architecture.md) | 当前模块依赖及后续数据/计算设计 |
| [文件格式与兼容性](03-file-format-and-compatibility.md) | 当前 XLSQ v1、目标 v2、类型与保存协议 |
| [脚本与运行时](04-scripting-and-runtime.md) | 当前 QuickJS 适配、目标 worker/动作/权限边界 |
| [实施计划](05-implementation-plan.md) | M1 进展、M2–M6 依赖与退出条件 |
| [质量性能与交付](06-quality-performance-and-delivery.md) | 已有证据、候选包验收和未来质量门禁 |
| [架构决策](07-architecture-decisions.md) | 已选方向、取舍及待验证事项 |
| [基线与实施记录](08-baseline-and-progress.md) | 初始 102 例基线和 0.2 新验证分别记载 |
| [M3 跨表与增量计算](10-m3-calculation.md) | 0.4 语义矩阵、脚本桥、结构命令、导出与验收 |
| [M2 文件与类型](09-m2-files-and-types.md) | 0.3 支持矩阵、实现、验证与下一阶段 |

用户操作见[试用说明](../README.md)，公开限制和运行时要求见[发布说明](../RELEASE-NOTES.md)。开发时先读系统架构与实施计划，再进入专题。

历史修复版 0.2.1：关闭确认延后到 WPF Closing 事件返回后执行，并新增真实窗口关闭回归。0.2.0 的发布记录保留在实施记录第 6 节，关闭缺陷与 0.2.1 验证见第 7 节。

## 状态约定

- **当前实现**：已在源码中接入的行为；验证状态以实施记录为准。
- **目标设计**：尚待实施的格式、worker、完整桌面和 Beta 能力。
- **待验证决策**：有明确实验或里程碑，未完成取证前不视为通过。
- **后续范围**：本次 0.2 交付不包含的功能，不因已有基础类或局部测试而关闭整个任务。

DOM-001/002、APP-001/002/003、CMD-001/002、UI-001 的 M1 实现与交付验收已完成。Release solution 构建 0 警告、0 错误，Core 221 例通过，最终发布包 native/WPF smoke 和 ZIP 解压后 WPF 复验均通过；包与 SHA-256 见实施记录第 6 节。0.3 在此基线上推进 M2；性能数字仍是目标。

## 0.2 历史实现

| 项目 | 实现与边界 |
| --- | --- |
| 平台与发布 | Windows x64、WPF、.NET 10；framework-dependent，需要 .NET 10 Windows Desktop Runtime x64 |
| 领域与会话 | 类型值、稀疏不可变快照、稳定会话内 SheetId、单写入门禁、单调 Revision |
| 编辑与历史 | 单格/公式栏/粘贴/清空统一提交，新增表可撤销，默认 32 MiB 历史预算 |
| 计算 | 同表依赖和错误诊断；结果须匹配会话/输入修订/计算代次/脚本代次，跨表及依赖保持 Blocked 缓存 |
| 脚本 | QuickJS-NG 0.17.0，每轮新 runtime，显式启用；任何脚本编辑撤销授权；无 worker/OS 隔离 |
| 保存 | 不可变快照、串行写入和原子替换；待编辑输入先提交，重算失败时允许保存输入与旧缓存并提示 |
| 格式 | XLSQ v1 三条目；literal 保存为文本，重开保持 Text；XLSX 导出全部公式的当前值 |
| 兼容元数据 | 保留原型受支持的样式、尺寸、隐藏、合并与图片；界面呈现与完整类型保真尚不完整 |
| 验证与产物 | 本轮 build/Core 日志位于 `artifacts/validation/release-0.2`；最终包 smoke、ZIP/hash 记录见实施记录 |

源码入口：[桌面组合根](../DesktopComposition.cs)、[编辑控制器](../EditorController.cs)、[文档 facade](../../QuickExcel.Core/WorkbookDocument.cs)、[会话](../../QuickExcel.Core/DocumentSession.cs)、[存储适配](../../QuickExcel.Infrastructure/Storage/LegacyWorkbookStorage.cs)、[包封装](../../QuickExcel.Infrastructure/DocumentPackage.cs)、[原生桥](../../QuickExcel.Scripting/QuickJsSession.cs)。

## 文档维护

任务状态、契约和验证证据一起更新。保留初始 102 例与原型 smoke 的历史结论，新增记录不冒充历史测试。修改格式、授权或编译边界时同步 ADR 与专题。发布包元数据记录 Git HEAD 和工作树状态；本次实现尚未提交，不能只用旧 HEAD 表示实际构建源码。
