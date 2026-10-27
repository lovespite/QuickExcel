# QuickExcel 0.2

Windows x64 桌面表格，基于 WPF、[ExcelIO](https://github.com/lovespite/ExcelIO) 与 QuickJS-NG，工作簿扩展名为 `.xlsq`。

0.2 是实施计划 M1 范围的试用版：统一文档会话、不可变稀疏模型、事务编辑、撤销/重做，以及可靠的保存入口。XLSQ 仍为 v1，字面量保存后重新打开为文本；跨表计算、worker、完整类型保真和 M2–M6 Beta 能力留待后续。具体范围见[发布说明](QuickExcel.Desktop/RELEASE-NOTES.md)。

## 启动与试用

解压完整发布目录后运行 `QuickExcel.exe`。程序需要 **.NET 10 Windows Desktop Runtime（x64）**，采用 framework-dependent 发布，不包含运行时。源码仓库也可双击 `Start-QuickExcel.cmd`，优先启动 `artifacts/QuickExcel/QuickExcel.exe`。

打开工作簿不会执行脚本。只启用你信任的脚本；当前 QuickJS 在应用进程内运行，没有独立 worker 或操作系统沙箱。操作与格式限制见[试用说明](QuickExcel.Desktop/README.md)。

## 仓库与模块

QuickExcel 与同级 `../ExcelIO` 是独立 Git 仓库，分别维护提交、分支和远端。ExcelIO 是源码构建依赖，不复制到本仓库，也不作为普通 QuickExcel 修改的一部分。

| 目录 | 职责 |
| --- | --- |
| `QuickExcel.Core` | 纯 .NET 领域模型、DocumentSession、命令、历史、WorkbookDocument facade 和端口；不引用 ExcelIO、WPF 或原生运行时 |
| `QuickExcel.Infrastructure` | ExcelIO 存储/计算适配和 QuickJS runtime factory；依赖同级 ExcelIO 和本仓库 Scripting |
| `QuickExcel.Scripting` | 托管 QuickJS 桥、固定版本原生来源、运行时 smoke |
| `QuickExcel.Desktop` | WPF 界面、组合根、ViewModel、编辑控制器和设计文档 |
| `QuickExcel.Core.Test` | 领域、会话、适配器、存储和真实运行时集成回归 |

迁出基线为 ExcelIO 提交 `5ff1aa0`。本轮交付仍是未提交的 QuickExcel 工作树；具体 HEAD、验证日志和产物来源见[实施记录](QuickExcel.Desktop/docs/08-baseline-and-progress.md)。

## 构建与发布

从源码构建需要 Windows x64、.NET 10 SDK 和同级 ExcelIO 仓库。首次构建原生引擎还需要 CMake、Visual Studio C++ 工具及下载固定源码的网络连接，见[原生运行时来源](QuickExcel.Scripting/native/SOURCES.md)。

在仓库根目录运行：

```powershell
dotnet build .\QuickExcel.slnx -c Release --artifacts-path .\artifacts\build -v minimal
dotnet test .\QuickExcel.Core.Test\QuickExcel.Core.Test.csproj -c Release --artifacts-path .\artifacts\build -v minimal
dotnet run --project .\QuickExcel.Scripting\smoke\QuickExcel.Scripting.Smoke.csproj -c Release --artifacts-path .\artifacts\build
.\scripts\build-quickexcel.ps1 -RunSmoke -CreateArchive
```

发布脚本使用 `-r win-x64 --self-contained false`。默认运行 Core 测试和 native smoke，发布到 `artifacts/QuickExcel`；`-RunSmoke` 加入 WPF 验收，`-CreateArchive` 同时启用 WPF 验收并在 `artifacts/releases` 生成唯一命名 ZIP 与 `.sha256`。可用 `-OutputDirectory artifacts/QuickExcel-0.2` 选择本仓库 artifacts 内的另一目录。WPF smoke 使用新的独立目录、45 秒看门狗、退出码及新 PASS 文件判定。

包内包含 `RELEASE-NOTES.md`、`BUILD-METADATA.json`、原生许可证和依赖来源说明。元数据记录本次实际验证状态，不能用既有 smoke 结果代替新源码验证；无 SDK 干净机器、安装器和签名验收尚未完成。

所有托管构建、测试、运行和发布均须指定 `--artifacts-path`，让同级引用项目的输出也进入本仓库 `artifacts/build`。artifacts、原生下载和构建缓存不提交到 Git。

- [试用说明](QuickExcel.Desktop/README.md)
- [发布说明](QuickExcel.Desktop/RELEASE-NOTES.md)
- [设计与实施文档](QuickExcel.Desktop/docs/README.md)
