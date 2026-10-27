# QuickExcel 0.4 计算预览版

本版在 M2 文件支持之上交付 M3 计算子集，支持跨表引用和标准公式增量求值。需要 Windows x64 与 .NET 10 Windows Desktop Runtime；保留完整发布目录后运行 QuickExcel.exe。

## 使用

1. 打开或新建工作簿。打开不执行源码，已有公式缓存可能过期。
2. 双击单元格或在公式栏编辑。前置单引号强制文本，`00123` 保留文本，单独输入单引号表示空文本。
3. Ctrl+S 保存 XLSQ v2；活动单元格与公式栏会先提交。输入无效阻止保存；重算失败仍可保存有效输入与旧缓存。
4. v1 文件迁移后显示未保存状态，首次保存选择其他文件名，旧文件保留。
5. 含未知部件的文件会显示兼容报告。未修改工作簿内容时可以原样保留；编辑后可取消，或明确另存移除这些内容的副本。重新打开副本可刷新兼容报告。
6. “导出值快照”将公式全部转为结果；“导出公式”冻结脚本及传递依赖结果，保留独立标准公式。需要冻结的结果缺失时拒绝，过期结果须明确选择。
7. 只在信任当前源码时点击“应用脚本并重算”（Ctrl+Enter）。源码改变会立即撤销授权，多模块包当前只能保存、不能执行。

## 编辑和快捷键

Ctrl+O 打开，Ctrl+N 新建，Ctrl+Shift+S 另存；表格 Ctrl+C 复制显示值、Ctrl+V 粘贴矩形数据、Delete 清空。Ctrl+Z/Y 撤销/重做单元格事务及新增工作表；脚本编辑器使用自己的文本历史。默认历史预算 32 MiB，截断时显示提示。

一次粘贴最多 20,000 格，任一目标非法则整批不写。界面最多显示 100,000 行 × 256 列；模型最多 64 表、1,024 列、全簿 1,000,000 存储格，文件还须满足解压与解析预算。远端空洞不展开为密集矩形。

## 文件与脚本

v2 保存唯一的 workbook.xlsx，以及身份/修订状态和脚本源码模块映射。支持数字、布尔、标准错误、文本、空白/空文本、公式缓存、1900/1904 日期系统及基础格式。历史 v1 中已经丢失的类型不猜测恢复。

```javascript
quickExcel.functions.register("DISCOUNT", (price, rate) => price * (1 - rate));
quickExcel.functions.register("DOUBLE", value => value * 2);
```

支持同步标量与一维区域参数；没有 Promise、对象或动态数组溢出支持。QuickJS 仍在应用进程内，仅供可信脚本；没有独立 worker 或 OS 沙箱。打开、迁移、另存均不会复制运行授权。

跨表计算已接入；未授权脚本和依赖仍保留 Blocked 缓存。主题、富文本、图表/透视表、复杂锚点、单元格内图片与数组公式等不在可编辑保真承诺内；完整矩阵见 [M2 文件与类型](docs/09-m2-files-and-types.md)。图片/合并完整界面呈现、排序筛选、恢复和无障碍等留待后续。

## 构建与验收

```powershell
dotnet build .\QuickExcel.slnx -c Release --artifacts-path .\artifacts\build -v minimal
dotnet test .\QuickExcel.Core.Test\QuickExcel.Core.Test.csproj -c Release --artifacts-path .\artifacts\build -v minimal
.\scripts\build-quickexcel.ps1 -RunSmoke -CreateArchive
```

从仓库根目录执行。所有引用项目输出也留在本仓库 artifacts。发布脚本执行 Core/native/WPF 检查并生成 ZIP、SHA-256 和 BUILD-METADATA；未声称无 SDK 干净机器、签名或安装器验收已完成。详见[发布说明](RELEASE-NOTES.md)。

计算范围与函数列表见 [M3 支持矩阵](docs/10-m3-calculation.md)。新计算器拒绝脚本覆盖内置名称；请使用 JS 命名空间。结构命令 API 已提供，完整界面操作继续 M5。
