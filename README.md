# QuickExcel 0.3 文件预览版

Windows x64 桌面表格，基于 WPF、同级 ExcelIO 源码库与 QuickJS-NG。当前阶段在 0.2 的文档会话、事务编辑和撤销/重做基础上加入 **M2 文件与类型保真**。

## 试用

运行 `artifacts/QuickExcel/QuickExcel.exe`，或双击 `Start-QuickExcel.cmd`。发布包需要 **.NET 10 Windows Desktop Runtime x64**，采用 framework-dependent 发布。

- 新建及保存使用 **XLSQ v2**：数字、文本、布尔、标准错误、空文本、空白和公式缓存分别保存；保留 1900/1904 日期系统、数字格式及稳定工作表身份。
- v1 打开时迁移，不猜测恢复历史上已丢失的类型；首次保存必须选新路径，保留原文件。
- 不支持的 XLSX 部件和属性进入兼容报告。工作簿内容未变时保留原 XLSX 字节；编辑后普通保存会阻止不确定的损失，可明确选择另存受支持副本。
- 导出 XLSX 为类型化结果快照。缺失缓存阻止导出；旧缓存需要用户明确选择。
- 打开、迁移、保存不执行脚本，不恢复授权。模块源码和绑定可持久化；多模块执行留到 M4。

本版不等于完整 Beta。跨表/增量计算仍属 M3，独立 worker 属 M4，完整桌面交互与自动恢复属 M5/M6。进程内 QuickJS 仅供可信脚本，没有 OS 沙箱。见[支持范围与验收](QuickExcel.Desktop/docs/09-m2-files-and-types.md)及[发布说明](QuickExcel.Desktop/RELEASE-NOTES.md)。

## 仓库边界

QuickExcel 与同级 `../ExcelIO` 是独立 Git 仓库。ExcelIO 是源码构建依赖；普通 QuickExcel 任务不修改、复制或提交同级库的源码。

| 目录 | 职责 |
| --- | --- |
| `QuickExcel.Core` | 纯 .NET 领域快照、会话、命令、历史和端口 |
| `QuickExcel.Infrastructure` | 类型化 XLSX/XLSQ v2、保留的 v1 适配、计算与运行时适配 |
| `QuickExcel.Scripting` | QuickJS 托管桥和固定来源的原生引擎 |
| `QuickExcel.Desktop` | WPF、组合根、编辑控制器和集成 smoke |
| `QuickExcel.Core.Test` | 领域、会话、存储、计算和格式异常回归 |

## 构建

需要 Windows x64、.NET 10 SDK 和同级 ExcelIO。首次原生构建还需 CMake、Visual Studio C++ 工具及下载固定源码的网络连接，见[来源说明](QuickExcel.Scripting/native/SOURCES.md)。

```powershell
dotnet build .\QuickExcel.slnx -c Release --artifacts-path .\artifacts\build -v minimal
dotnet test .\QuickExcel.Core.Test\QuickExcel.Core.Test.csproj -c Release --artifacts-path .\artifacts\build -v minimal
dotnet run --project .\QuickExcel.Scripting\smoke\QuickExcel.Scripting.Smoke.csproj -c Release --artifacts-path .\artifacts\build
.\scripts\build-quickexcel.ps1 -RunSmoke -CreateArchive
```

所有托管命令均须指定 `--artifacts-path`，让引用项目的构建输出也进入本仓库。发布脚本生成独立新发布树、ZIP、SHA-256 和 BUILD-METADATA；不把旧 smoke 文件当成本次通过证据。无 SDK 干净机器、安装器和签名验收尚未完成。

- [试用操作](QuickExcel.Desktop/README.md)
- [实施计划](QuickExcel.Desktop/docs/05-implementation-plan.md)
- [M2 支持矩阵与证据](QuickExcel.Desktop/docs/09-m2-files-and-types.md)
