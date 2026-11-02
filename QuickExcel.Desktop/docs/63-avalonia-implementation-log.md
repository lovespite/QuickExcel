# Avalonia 迁移实施日志

2026-11-03。计划见 [61 号文档](61-avalonia-migration-plan.md)，A0 的依赖、运行与性能基线见 [62 号记录](62-avalonia-a0-validation.md)。每阶段独立本地提交，用户验收后才推送远程。

## 验收顺序调整

用户要求：“涉及实机验证的先跳过，在实施日志中记录即可，后期统一验证。”因此实机项延期，不阻塞工程实施，也不标记为通过。Linux 虚拟机用户实测无法输入中文、无候选框，保留为已知未解决问题。macOS、真实打印、跨屏 DPI、辅助技术、干净设备部署与真实设备性能继续待验收。不得用此调整降低脚本隔离、文件保真或草稿安全要求。

## A1 共享层（完成）

- 新建 `QuickExcel.Presentation`（net10.0）：迁入 11 个 Interaction 源文件、EditorController、IGridSurface；提供公共共享类型。编辑、验证确认、预期修订、撤销事务、公式粘贴、组合门禁与原单元格草稿所有权逻辑保持。
- 新建 `QuickExcel.Grid.Shared`（net10.0）：移动 SeriesSizes、FastGridCellAddress、ActiveSeries 和矩形/合并选区算法。保留历史命名空间、原始来源哈希与 MIT 声明；WPF 与测试均引用同一程序集，不链接重复源。许可发布到独立子目录，避免两个项目的同名输出冲突。
- 新建本地文件选择端口，明确取消返回 null、非本地存储 URI 拒绝。WPF 打开/保存使用适配器，现有注入对话框和 SaveChecked/打开指纹流程保留；没有把所有 URI 当成本地路径。
- Core、Infrastructure、计算、文件格式及 worker 实现未改动；共享程序集不引用 WPF、Avalonia、Infrastructure 或 native。

实际验证：

| 检查 | 结果/证据 |
| --- | --- |
| Release 解决方案构建 | 0 警告、0 错误；输出 `artifacts/build` |
| 相关规则与新契约测试 | **177/177**；`artifacts/avalonia/tests/a1-shared-final.trx` |
| WPF 发布与交互 | `artifacts/avalonia/a1-wpf`；**95** 项检查，`artifacts/validation/m5-grid/20261102-164457-15d1682c` |
| WPF 工作区 | **157** 项检查，`artifacts/validation/m5-grid/20261102-164503-b7c96f15`；包含关闭取消、保存与冲突等原入口 |

定向测试覆盖编辑原子历史、确认期间修订冲突、端口路径限制、共享程序集依赖、剪贴板/组合规则、尺寸/超大格/合并选区、筛选排序、查找与兼容报告。首次执行暴露两项旧测试仍硬编码 100,000 行和 4,096 样式（现有限额已是 200,000/8,192）；改用 `WorkbookLimits` 验证真实边界后通过，未调整产品限额。首次发布的同名许可证冲突已修复并重发。WPF smoke 为路由事件/合成交互证据，不冒充物理输入。没有运行全量 Core 回归或重复性能基准。

## A2 Avalonia 工作区（实施中）

接下来创建正式宿主，接入共享编辑/文件流程、Fluent 工作区和基础网格；正式默认入口继续保留 WPF，功能对齐与 Web 组件在后续阶段记录。
