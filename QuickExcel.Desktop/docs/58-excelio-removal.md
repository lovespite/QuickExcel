# 移除 ExcelIO 与 ExcelIO.Formula

2026-11-02：QuickExcel 的构建、测试与发布不再依赖同级 ExcelIO 仓库。本次仅修改 QuickExcel；同级库的源码未修改、复制或删除。

## 实现与兼容性

- 从 solution 和 Infrastructure 删除两个项目引用，删除 `DocumentPackage`、`LegacyWorkbookAdapter`、`LegacyWorkbookStorage`、`LegacyWorkbookCalculator`。
- 桌面继续使用 `TypedWorkbookStorage` 与 `WorkbookCalculator`；公式 AST、绑定及求值已由 QuickExcel 自有实现承担。脚本 worker、协议、原生引擎和隔离策略未改变。
- XLSQ v1 仍可导入，文本不会猜测为数字/布尔，文件缓存始终作为旧值处理，打开不执行脚本；升级必须另存为 v2，原文件不覆盖。测试与桌面 storage smoke 通过独立 ZIP/XML 夹具覆盖这些行为，无需运行旧序列化器。
- 发布脚本不再读取同级项目版本、Git 状态或生成 ExcelIO 许可条目；拒绝新发布树中的 ExcelIO 文件，更新已有输出目录时清理两个库的 DLL/PDB/XML。基线采集仅访问本仓库。

## 回归迁移

`WorkbookDocumentTests`、`CalculationAdapterTests` 和 `StorageAdapterTests` 改用当前实现，保留文档往返、稀疏坐标、文本身份、公式缓存、样式/图片、原子保存、脚本参数、释放及取消验证。

旧计算器的跨表拒算已替换为跨表计算/缺失表错误；脚本覆盖内置函数替换为冲突拒绝与资源释放；未知函数保持旧缓存并标记 Blocked；空白的直接公式引用得到数字 0。文档导出验证先拒绝旧缓存，再重算并导出。图片验证保留浮动图片往返和快照不变，并明确测试单元格内图片拒绝写入。

只验证旧对象拷贝/全局回调的测试随旧模型删除；旧稠密展开预算退出，当前稀疏保存仍由 `TypedStorageTests.RemoteCellsRemainSparseAndDoNotHitTheLegacyDenseBudget` 覆盖。旧阶段文档保留历史事实，当前依赖与构建说明已更新。

发布验证还修正了 worker smoke 的旧脚本函数名：`VALUE` 已是内置函数，夹具改为 API v1 支持的纯字母名称 `WORKERVALUE`，并在失败时记录实际输入、结果和界面诊断。生产函数注册及隔离规则未放宽。

## 验证

- 独立输出目录 `artifacts/build-no-excelio` 的 Release solution 构建：0 警告、0 错误。
- 上述三个测试类加 `TypedStorageTests`、`M3CalculationTests`：177 项通过，0 失败、0 跳过；TRX：`artifacts/validation/excelio-removal/dependency-removal.trx`。
- 补强非法 v1 夹具后，20 项文档测试再次通过（`document-regressions.trx`）；独立 worker 启动专项 1 项通过（`worker-launch.trx`）。QuickJS native smoke 通过。
- 新输出的 10 份 restore graph、18 份 deps 清单均无 ExcelIO 项目/库；新构建目录没有 ExcelIO DLL。
- `build-quickexcel.ps1 -SkipTests -OutputDirectory artifacts/validation/excelio-removal/package` 完成，发布元数据与许可清单不再读取/记录 ExcelIO。验证预置的 6 个旧 DLL/PDB/XML 在输出覆盖后已被清理。产物审计见 `artifacts/validation/excelio-removal/dependency-audit.json`。
- 首次依赖移除验证中，同一组 DLL 的桌面存储（含 v1）、计算、日期显示、worker、actions、关闭生命周期检查通过，结果位于 `artifacts/smoke/20261102-081949-332f37fc/`。综合 UI smoke 随后在 `MainWindow.CellLayoutSmoke.cs:60` 因 nullable 空值失败；当时该文件未修改，完整 UI smoke 未通过。该次发布元数据保留 `skipped`，后续修复与通过证据见下。
- 未运行完整 Core 回归套件，未改动 M4/M6 的未完成门禁。

## 后续修复：布局冒烟的视口前提（2026-11-02）

定向复现确认：缩放/选择结束后，网格 `FirstVisibleRowScrollIndex=1`，选中 C4，A1 已滚出可见区域。`GetModelCellBounds(0, 0)` 正确返回 null；测试直接访问 `.Value` 导致异常，并非单元格或存储数据丢失。

布局检查现在在测量前显式滚动目标单元格进入视口并等待绘制。填充采样还主动滚到远处，确认 A1 不可见，再恢复 A1 区域并检查选区仍停在填充区域外。相邻单元格、编辑器尺寸与像素填充断言全部保留；坐标仍缺失时报告单元格、滚动索引、网格尺寸和选区，不跳过检查。

- Desktop Release 构建：0 警告、0 错误。
- `scripts/test-m5-grid.ps1 -Mode Interaction -PackageDirectory artifacts/build/bin/QuickExcel.Desktop/release`：95 项 WPF 断言通过，含 100%/125%/150% 缩放。证据：`artifacts/validation/m5-grid/20261102-082624-e0b00e89/Interaction/`。
- `scripts/build-quickexcel.ps1 -SkipTests -RunSmoke -OutputDirectory artifacts/validation/cell-layout-fix/package`：发布、native smoke、完整桌面 smoke 均通过，10 个结果文件位于 `artifacts/smoke/20261102-082645-db18ccdf/`；新包元数据记录 `wpfSmoke=passed`。新发布包仍无 ExcelIO/ExcelIO.Formula 依赖。
- 本轮只改布局 smoke 的前提和诊断，未改变网格生产逻辑；WPF 输入是合成事件/模型选区，不代表物理 IME 或完整 M6 验收。未运行全量 Core 回归或基准，未生成 ZIP。
