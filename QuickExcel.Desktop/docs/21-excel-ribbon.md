# 类 Excel 功能区（2026-10-29）

命令栏改为浅灰选项卡、白色分组面板、组间分隔线和底部分组名。开始页之外提供文件菜单及开始、插入、数据、视图选项卡；工作表/脚本切换保持独立。标题栏提供保存、撤销、重做。

开始选项卡包含剪贴板、字体、对齐方式、数字、单元格和编辑六组；原有文件、结构、排序、筛选、冻结、查找与定位入口均保留。图标使用固定版本 `FluentIcons.Wpf 2.1.344` 的 Regular 图形；MIT 声明随构建和发布输出复制。字体名称优先使用中文，选中状态随当前单元格、撤销和重做更新。脚本页面禁用工作表功能区，文件及页面切换仍可用。

字体、字号、垂直对齐和自动换行补入 `CellStylePatch`，继续通过 `SetCellStyleCommand` / `DocumentSession` 原子提交和撤销。网格与冻结区域显示这些已有存储字段；字号按点转换为 WPF DIP。自动换行不自动调整行高，可在“单元格 → 格式”调整。保留原有颜色预设、合并语义和粘贴策略：普通粘贴为值，粘贴公式显式解析输入。不提供尚无命令支持的占位选项卡或按钮。

功能区使用原生 WPF TabControl、Button、ToggleButton、ComboBox 与 ContextMenu；主题位于 `Themes/Ribbon.xaml`。最小 900 DIP 窗口下分组保持单行，横向滚动可访问末尾命令。图标按钮含工具提示、辅助名称、键盘焦点和禁用状态。

验证仅运行相关区域：

```powershell
dotnet build .\QuickExcel.Desktop\QuickExcel.Desktop.csproj -c Release --artifacts-path .\artifacts\build -v minimal
dotnet test .\QuickExcel.Core.Test\QuickExcel.Core.Test.csproj -c Release --artifacts-path .\artifacts\build --filter FullyQualifiedName~M5WorkspaceTests -v minimal
.\scripts\test-m5-grid.ps1 -PackageDirectory artifacts\build\bin\QuickExcel.Desktop\release -Mode Ribbon
.\scripts\test-m5-grid.ps1 -PackageDirectory artifacts\build\bin\QuickExcel.Desktop\release -Mode Product
```

本机验证：Release 构建成功；29 项工作区测试通过；21 项功能区 WPF 断言通过；41 项现有产品 WPF 断言通过。功能区专项覆盖实际按钮与下拉框、选区样式、草稿归属、撤销/重做、页签、辅助名称及窄窗口可达性，并保留正常/窄窗口截图。Product 在受限执行环境的 worker 步骤失败后，已在正常用户环境复验通过；未降低隔离要求。截图和逐项结果保存在 `artifacts/validation/m5-grid` 下各独立运行目录。未运行完整回归或性能基准，不据此关闭 M5/M6 门禁。
