# QuickExcel 0.2 试用版

Windows x64 桌面表格，基于 WPF、ExcelIO 和 QuickJS-NG。工作簿扩展名为 **`.xlsq`**。本版交付 M1 的统一会话、稀疏类型模型、事务编辑、撤销/重做和保存整合；M2–M6 的完整 Beta 能力仍在计划中。

## 开始试用

运行完整发布目录中的 `QuickExcel.exe`，或在源码仓库根目录双击 `Start-QuickExcel.cmd`。需要 **.NET 10 Windows Desktop Runtime（x64）**；程序采用 framework-dependent 发布，复制单个 exe 无法替代完整目录。

1. 启动后展示示例工作簿与 JavaScript 源码。打开文件时脚本默认停用，单元格可能显示旧缓存。
2. 检查源码，只在信任它时点击 **应用脚本并重算**（Ctrl+Enter）。
3. 双击价格、数量等单元格编辑，也可在公式栏输入后按 Enter。数值与布尔输入在内存中有独立类型；前置单引号表示文本，`00123` 保留为文本。
4. 修改脚本后授权立即撤销；即使新源码暂时无效，也不会继续使用旧授权。确认编辑后再次点击应用。
5. **保存 .xlsq**（Ctrl+S）先提交正在编辑的单元格与公式栏，再保存源码、公式和结果缓存。无效输入阻止保存；若输入有效但重算失败，仍可保存输入和旧缓存，界面提示结果可能过期。
6. **导出 .xlsx** 将所有公式转换为当前缓存值，不带脚本。先检查缓存状态与运行记录；导出不修改原文档公式和已保存修订。

## 编辑与撤销

- Ctrl+O 打开，Ctrl+N 新建，Ctrl+S 保存，Ctrl+Shift+S 另存为。
- 表格中 Ctrl+C 复制显示值，Ctrl+V 粘贴制表符分隔的矩形数据；Delete 清空选择内容。
- 一次粘贴或清空是一个事务；任一目标非法，整批不提交。Ctrl+Z / Ctrl+Y 撤销或重做单元格编辑、粘贴、清空和新增工作表。
- 脚本编辑器使用自身文本撤销，不与表格命令历史混用。脚本授权或重算不会把已保存文档标为未保存。
- 撤销历史默认预算 32 MiB，释放较早记录时界面会提示；单次操作超过预算则在修改前拒绝。
- 底部可切换/新增工作表、扩展显示行列。保存前无需专门按 Enter 提交公式栏。

## 脚本 API

```javascript
quickExcel.functions.register("DISCOUNT", (price, rate) => price * (1 - rate));
quickExcel.functions.register("DOUBLE", value => value * 2);
```

单元格可使用 `=DISCOUNT(B2,D2)`、`=DOUBLE(A1)`、`=SUM(E2:E5)`。函数名不区分大小写，当前支持同步数字、字符串、布尔或空结果与一维区域参数；不支持 Promise、对象、动态数组溢出。

脚本不提供文件系统、网络、工作簿写入或 .NET 反射接口。QuickJS 有内存、栈和执行时间限制，但仍在应用进程内运行，**没有独立 worker 或操作系统沙箱，只应启用可信脚本**。每轮完整计算创建新运行时，不依赖上一轮的可变 JavaScript 状态。

## 文件与兼容性

XLSQ 仍为 `formatVersion: 1`，包含 `manifest.json`、`workbook.xlsx` 和 `scripts/functions.js`。只保存源码，不保存字节码、脚本授权或会话身份。保存从不可变快照开始，写同目录临时文件后替换；失败保留上次成功文件，只有成功保存才更新已保存修订。

**内存类型不等于文件类型保真。** v1 保存和 `.xlsx` 结果导出仍沿用旧文本格式；重新打开字面量为 Text，不猜测恢复数字、布尔或错误身份。空白与空文本可能合并，公式缓存重新载入时视为可能过期。完整日期系统和类型往返属于 M2；当前适配器不支持保存 Mac1904 快照。

原型已支持的样式、行列尺寸、隐藏状态、合并和图片在受支持格式范围内保留；合并/图片的完整界面呈现和编辑仍待后续。复杂 Excel 文件、未知部件、数字格式和图表不承诺无损往返。

## 当前边界

- 同表公式可重算；跨表公式及其依赖为 Blocked，保留缓存和诊断。没有完整增量计算、跨表引用变换或完整 Excel 函数兼容。
- 界面最多显示/编辑 100,000 行 × 256 列，单次粘贴 20,000 格。领域最多 64 张表、1,024 列和总计 1,000,000 个存储格；未显示的受支持内容仍会保存。
- IO 旧库转换另有 1,000,000 个占位行/单元格预算，跨度过大的稀疏数据保存会明确拒绝，不会无限扩展稠密列表。
- 计算/IO 期间暂时锁定界面。排序、筛选、冻结、完整 IME/无障碍验收、恢复、签名安装器和自动更新仍属后续范围。

## 构建与验证

从仓库根目录运行：

```powershell
dotnet build .\QuickExcel.slnx -c Release --artifacts-path .\artifacts\build -v minimal
dotnet test .\QuickExcel.Core.Test\QuickExcel.Core.Test.csproj -c Release --artifacts-path .\artifacts\build -v minimal
.\scripts\build-quickexcel.ps1 -RunSmoke -CreateArchive
```

脚本使用 `-r win-x64 --self-contained false`，发布前验证 Core 和 native，按参数在新目录运行带看门狗的 WPF smoke。ZIP、SHA-256 及包内 `BUILD-METADATA.json` 标识本次产物和实际验收结果；最终包复验进度见[实施记录](docs/08-baseline-and-progress.md)。没有通过无 SDK 干净机器验收的声明。

Core 现在是纯 .NET 程序集；ExcelIO、文件封装与 QuickJS 工厂由 `QuickExcel.Infrastructure` 实现，WPF 的组合根注入这些端口。仓库边界和原生构建要求见[仓库说明](../README.md)，完整公开范围见[发布说明](RELEASE-NOTES.md)。
