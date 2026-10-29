# 主界面网格迁移（2026-10-29）

主工作区直接使用 `QuickExcel.Grid.Wpf` 的 `FastGridControl`，由 `FastGridSurface` 适配。旧 `WpfGridSurface`、DataGrid 单元格/行容器、冻结副本网格、旧尺寸事件和 Adorner 实现、旧 Grid.xaml 主题以及 `FastGridWindow` 试验入口已移除。此前的 [23 号记录](23-fast-grid-implementation.md) 保留历史对照证据。

## 行为与边界

- 文档仍由 `EditorController` / `DocumentSession` 写入。公式栏和原生 TextBox 保留原地址草稿、组合输入门控、显式公式粘贴和原子撤销。忙碌时被窗口拦截的组合事件不能激活网格草稿。
- 主界面使用已有导航、剪贴板、保存、查找、结构和 Ribbon 命令。绑定新文档重置地址；普通计算和格式刷新保留选区、滚动位置、模型和输入控件。过滤掉当前行时定位可见标题。
- 冻结行列由同一控件绘制。筛选按钮、编辑层和图片/合并层按冻结及滚动分区裁剪。字体、字号、对齐、换行和各边框由 `FastGridArtwork` 投影到内容层，选区在其上独立绘制。图片解码继续限制像素数和解码宽度。
- 行高和列宽拖动由主窗口处理，拖动时只更新视图，结束时提交一个尺寸命令，取消不产生历史。尺寸依旧以行高 DIP / 列宽字符单位保存，几何会按设备像素取整。
- 隐藏当前行列后选择最近可见位置；“取消隐藏”包含当前选区两侧相邻的手动隐藏行列。区域操作可包含手动隐藏记录，筛选造成的不连续选区仍拒绝；复制跳过手动隐藏和筛选隐藏行，粘贴仍不能跨筛选隐藏行。
- `FastGridBenchmark` 只测当前实现；旧 DataGrid 对照代码已删除，历史原始样本保留在 artifacts 和 23 号记录中。没有修改 ExcelIO 源码、Core 模型或脚本隔离策略。

## 验证

### 选框重叠修复（2026-10-29）

普通多格选区内的小框有两个来源：选区层重复绘制当前格边框，以及透明原生 TextBox 的 Windows 主题焦点边框。现在当前格已在选区内时只绘制选区外框；待输入的文本框保持键盘焦点，但整体透明，进入编辑时恢复显示。当前地址和公式栏保持原语义。

先用 2×2 数据 `7/8/11/12` 复现，旧实现的选框断言失败证据在 `artifacts/validation/m5-grid/20261029-081852-ffbf06af/GridIntegration/`。最终 Desktop Release 构建通过，`GridIntegration` 的 11 项选框/编辑器检查和 41 项合并选区断言通过；整网格像素对比确认待输入文本框没有焦点边框残留，截图在 `20261029-082111-f2489372/GridIntegration/selection-outline.png`。`Interaction` 92 项断言通过，证据在 `20261029-082125-1ed9c382/Interaction/`。未运行完整回归；构建有已知 NU1900 数据源警告。验证为 WPF 绘制及处理路径检查，不代表物理输入验收。

### 合并单元格选区修复（2026-10-29）

合并内容层与选区层此前使用不同的范围：选中 F1:H7 时仍绘制 F1 的小框。现在网格接收快照中的合并范围，点击、拖选、程序选区和 Ctrl 增删均覆盖整个合并区域；当前格边框也使用合并范围。编辑地址归属左上角，编辑框覆盖合并区域，方向键、Tab/Enter 从区域边缘移出。单个合并区域粘贴单值只写左上角，并保持一次撤销。

针对本次修复，Desktop Release 构建通过，`MergedSelectionTests|M5InteractionTests` 共 24 项通过；未运行完整回归。`GridIntegration` 新增 41 项合并选区断言，涵盖 F1:H7、内部命中路径、Ctrl 增删、拖选、编辑及粘贴归属、导航、冻结滚动裁剪、隐藏行、取消合并与撤销。证据与截图在 `artifacts/validation/m5-grid/20261029-080318-75280ca9/GridIntegration/`。`Interaction` 92 项断言通过，证据在 `20261029-080325-c2e99551/Interaction/`。点击检查调用真实命中及选区处理路径，不代表物理鼠标输入验证。构建仍有 NuGet 漏洞数据源不可达的 NU1900 警告。

### 主网格迁移验证

所有命令从仓库根目录执行，构建输出在 `artifacts/build`。没有运行完整回归套件。

```powershell
dotnet build .\QuickExcel.slnx -c Release --artifacts-path .\artifacts\build -v minimal
dotnet test .\QuickExcel.Core.Test\QuickExcel.Core.Test.csproj -c Release --artifacts-path .\artifacts\build --filter "FullyQualifiedName~M5InteractionTests|FullyQualifiedName~M5DimensionTests|FullyQualifiedName~M5WorkspaceTests|FullyQualifiedName~DataOperationsTests|FullyQualifiedName~PresentationCalculationTests" -v minimal
.\scripts\test-m5-grid.ps1 -PackageDirectory artifacts\build\bin\QuickExcel.Desktop\release -Mode Interaction
.\scripts\test-m5-grid.ps1 -PackageDirectory artifacts\build\bin\QuickExcel.Desktop\release -Mode Workspace
.\scripts\test-m5-grid.ps1 -PackageDirectory artifacts\build\bin\QuickExcel.Desktop\release -Mode Ribbon
.\scripts\test-m5-grid.ps1 -PackageDirectory artifacts\build\bin\QuickExcel.Desktop\release -Mode Product
.\scripts\test-m5-grid.ps1 -PackageDirectory artifacts\build\bin\QuickExcel.Desktop\release -Mode DataOperations
.\scripts\test-m5-grid.ps1 -PackageDirectory artifacts\build\bin\QuickExcel.Desktop\release -Mode GridIntegration
```

Release 编译通过，NuGet 漏洞数据源不可达产生 NU1900 警告。78 项相关 Core 测试通过。专项 WPF 断言包括 Interaction 92、Workspace 108、Ribbon 21、Product 41、DataOperations 81；新增 GridIntegration 独立覆盖双向编辑归属、无效草稿地址、编辑中保存、日期数值精度和真实 WPF 关闭生命周期，避免为这些检查运行完整脚本/存储回归。Product 的 AppContainer 部分在受限执行器失败，在正常 Windows 用户环境通过，未弱化隔离。

原始功能证据位于 `artifacts/validation/m5-grid/` 的各次独立运行目录；`grid-core-tests.log` 和 `grid-solution-build.log` 位于 artifacts 根目录。

| 模式 | 通过的证据目录（相对 artifacts/validation/m5-grid） |
| --- | --- |
| Interaction | `20261029-074600-c36ae3c1/Interaction` |
| Workspace | `20261029-074602-5fd07272/Workspace` |
| Ribbon | `20261029-074407-ec13959b/Ribbon` |
| Product | `20261029-074411-33ccd7f0/Product` |
| DataOperations | `20261029-074415-637f91c7/DataOperations` |
| GridIntegration | `20261029-074607-d16805b7/GridIntegration` |

性能单独执行，保留原始样本：

```powershell
.\scripts\test-m5-grid.ps1 -PackageDirectory artifacts\build\bin\QuickExcel.Desktop\release -Mode FastGridBenchmark
.\scripts\test-m5-grid.ps1 -PackageDirectory artifacts\build\bin\QuickExcel.Desktop\release -Mode FastGridDragBenchmark
```

`20261029-073638-5290aabe/FastGridBenchmark` 达到预设预算：导航 p95 30.23 ms，25.6M 格选区 p95 1.28 ms，100k 行筛选 p95 82.48 ms。绑定只有一个样本，不视作分位数。`20261029-073839-2eb188e6/FastGridDragBenchmark` 的绘制预算通过，但像素自动滚动节拍 p95 为 33.66 ms，高于 33.34 ms 门槛，整体预算未通过；不能把结果文件中的“PASS: benchmark completed”解释为所有性能门槛通过。

移除无格式工作簿的多余样式遍历、复用一次可见列列表后，`20261029-074054-0f6b3814/FastGridDragBenchmark` 整体预算通过：视口拖选绘制 p95 2.50 ms、0 模型读取；像素滚动节拍 p95 32.28 ms。两次原始结果均保留，调度节拍接近门槛且受本机环境影响，不能据单次通过关闭 M6。

移除无格式工作簿的多余样式遍历、复用一次可见列列表后，`20261029-074054-0f6b3814/FastGridDragBenchmark` 整体预算通过：视口拖选绘制 p95 2.50 ms、0 模型读取；像素滚动节拍 p95 32.28 ms。两次原始结果均保留，调度节拍接近门槛且受本机环境影响，不能据单次通过关闭 M6。

真实 OS IME 候选窗口、跨屏 DPI、完整 UI Automation 单元格导航和 M5/M6 性能/安全门禁仍需独立验收。当前路由事件、模型选区、像素快照和 Dispatcher 时序检查不代表物理输入到显示延迟。默认切换是本次用户明确要求的集成变更，不据此关闭上述质量门禁。
