# QuickExcel

Windows x64 桌面表格，基于 WPF、ExcelIO 与 QuickJS-NG，文档扩展名为 `.xlsq`。

## 仓库边界

QuickExcel 与 ExcelIO 是同级的独立 Git 仓库，分别维护提交、分支与远端配置。

```text
D:\Projects\
├── ExcelIO\             # 文件读写库与公式引擎
└── QuickExcel\          # 本仓库：桌面应用、Core、Scripting、测试与文档
```

当前通过项目引用单向依赖 `../ExcelIO/ExcelIO/ExcelIO.csproj` 和 `../ExcelIO/ExcelIO.Formula/ExcelIO.Formula.csproj`；构建源码时需要同级 ExcelIO 仓库。ExcelIO 源码及其测试仍由 ExcelIO 仓库维护，不复制进本仓库，也不加入本仓库的 Git 跟踪。迁出基线为 ExcelIO 提交 `5ff1aa0`。

构建脚本和启动器使用 `--artifacts-path`，把依赖库在内的托管构建输出集中到本仓库的 `artifacts/build`。手动调用 dotnet 时也应使用下方参数，避免在同级 ExcelIO 中生成 QuickExcel 构建缓存。原生引擎输出位于本仓库的 `QuickExcel.Scripting/native`，均由 Git 忽略。

## 构建与运行

需要 Windows x64、.NET 10 SDK；首次构建原生引擎还需要 CMake、Visual Studio C++ 工具及下载源码的网络连接，详见[原生运行时来源](QuickExcel.Scripting/native/SOURCES.md)。

在本仓库根目录执行：

```powershell
dotnet build .\QuickExcel.slnx -c Release --artifacts-path .\artifacts\build -v minimal
dotnet test .\QuickExcel.Core.Test\QuickExcel.Core.Test.csproj -c Release --artifacts-path .\artifacts\build -v minimal
dotnet run --project .\QuickExcel.Scripting\smoke\QuickExcel.Scripting.Smoke.csproj -c Release --artifacts-path .\artifacts\build
.\scripts\build-quickexcel.ps1
.\Start-QuickExcel.cmd
```

发布目录为 `artifacts/QuickExcel`。已有发布包、smoke 输出与迁移前的构建缓存也保留在 `artifacts` 下，不提交到 Git。

- [原型使用说明](QuickExcel.Desktop/README.md)
- [设计与实施文档](QuickExcel.Desktop/docs/README.md)
