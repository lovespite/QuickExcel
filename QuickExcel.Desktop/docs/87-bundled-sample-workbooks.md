# 随应用发布的示例工作簿

产品 7.1.0 将九个内置数据处理库的调用示例放入仓库和安装包的 `samples/`。
开始页原“打开示例”入口改为“打开示例工作簿”，文件菜单也提供同名命令。
点击后显示九个工作簿的名称与用途，支持搜索、滚动、双击或选择后打开。

每个 XLSQ 使用文件脚本 API v4，包含“使用说明”“输入数据”“运行结果”
“预期结果”四张工作表，和 `main` 初始化、`example` Actions、`common`
公共模块。独立 `.ts` 源码也随包提供。初始化为空操作；执行 `example` 后
写入结果，支持撤销/重做。脚本不声明网络权限，全部示例可以离线运行。

| 库 | 演示内容 |
| --- | --- |
| Lodash | 订单去重、分组、求和、排序 |
| decimal.js | 文本精确单价、舍入、小计与合计 |
| Day.js | 严格解析、时区转换、日期运算 |
| Papa Parse | 引号/逗号/换行 CSV 解析与导出 |
| simple-statistics | 均值、中位数、标准差、回归、相关性、预测 |
| Zod | 按行校验、年龄转换、可读错误信息 |
| Fuse.js | 容忍拼写错误的商品搜索 |
| Arquero | 建表、分组、计数、求和、排序 |
| math.js | 表达式、单位、求导、矩阵乘法、大数 |

## 打开、保存与信任

入口遵守已有编辑器提交、输入法组合和未保存更改处理规则。取消列表或取消
替换工作簿均保留当前文档。正式读取器读取、核对示例哈希后，通过现有
`WorkbookWorkspace.Import` 打开一个无文件绑定的可编辑副本。副本从一开始
就需要保存，不向“最近文件”写入安装目录路径，也不持有模板文件租约。

脚本没有签名或预置授权。打开示例不调用自动信任路径、不执行脚本。用户从
正常脚本编辑器应用并授权，然后在“自动执行”中执行 `example`。
保存副本使用普通保存对话框，不默认覆盖安装包内的文件。

## 生成与交付

`tools/QuickExcel.SampleGen` 使用正式 `TypedWorkbookStorage` 生成 XLSQ；
输入数据与独立预期值在生成器中维护，源代码来自 `samples/scripts`。生成时
不执行脚本，不将测试 worker、测试入口或 Node/npm 放入安装包。

```powershell
dotnet run --project tools/QuickExcel.SampleGen/QuickExcel.SampleGen.csproj -c Release --artifacts-path artifacts/build -- samples
dotnet test QuickExcel.Core.Test/QuickExcel.Core.Test.csproj -c Release --artifacts-path artifacts/build --filter 'FullyQualifiedName~WorkbookSampleTests|FullyQualifiedName~WorkbookWorkspaceTests|FullyQualifiedName~LocalizationCatalogTests'
./scripts/test-avalonia-workspace.ps1 -Mode Samples
./scripts/build-quickexcel.ps1 -SkipTests -RunSmoke
```

修改脚本或数据后应重新生成并提交对应 XLSQ 与 `catalog.json`。桌面项目将
整个 `samples/` 复制到构建和发布目录，因此各平台使用相同示例来源。
目录由 JsonDocument 显式读取，不依赖生产反射序列化。文件名只接受 `.xlsq`
基本名称，拒绝路径穿越；读取期间使用只读租约核对 SHA-256。

## 验证范围

单元/集成检查使用正式隔离 worker 编译、执行全部九个工作簿中的 TypeScript，
逐格对照独立预期结果，验证初始化不修改结果、重复运行、撤销/重做、源码同步
和模板不被修改。另检查目录中的越界路径、未知 ID 与文件哈希变化。
桌面专项检查覆盖开始页按钮、九项列表、搜索/无结果、最小窗口、取消、未保存
更改、可编辑副本、另存和文件菜单。截图保存在对应 UI 检查目录。

发布脚本与生产自检检查全部九个文件及哈希，实际显示示例列表，并通过发布的
Native AOT worker 执行一个已分发的 Lodash 工作簿、验证结果和撤销。
本轮实际验证平台为 Windows x64；其他平台的打包复制入口随桌面项目共用，
没有据此宣称完成跨平台或物理输入法验收。

2026-11-20 Windows x64 验证记录：

- 相关工作簿/工作区检查 29 项通过；新增 CSV 从 65 行缩减到 1 行后，旧结果
  全部清除的回归检查通过。最终 11 项示例检查使用已发布 Native AOT worker
  全部通过，见 `artifacts/script-libraries/samples-final-aot.trx`。
- 本地化检查 5 项、桌面专项检查 15 项通过；界面证据位于
  `artifacts/avalonia/workspace/20261120-114952-ebb426a2`。
- Release 解决方案构建成功，0 警告 / 0 错误。生产裁剪保留现有诊断字段与
  ODBC/DataSet 警告；没有放宽发布检查。
- 最终包 `20261120-115527-7a374f18` 的生产自检 41 项通过并激活为本地启动包。
  已核对活动 manifest、九个工作簿、独立源码与已测试 AOT worker 的哈希。
  `artifacts/script-libraries/samples-delivery.json` 保存最终核对记录，旧包保留。
