# QuickExcel 内置库示例

主界面点击 **打开示例工作簿**，选择一个示例即可打开可编辑副本。每个文件包含
“使用说明”“输入数据”“运行结果”“预期结果”四张工作表，以及可查看、修改的
TypeScript 脚本。打开示例不会自动执行或授权脚本。

1. 查看输入和预期结果，或修改输入。
2. 打开“脚本 → 脚本编辑器”，选择 Actions 下的 `example` 查看代码。
3. 点击“应用并重算”（F5），按应用正常流程同意本次执行。
4. 在“自动执行”中选择 `example`，点击“执行”，切换到“运行结果”。
5. 运行结果可撤销；保存时选择自己的位置，不覆盖安装目录中的原始示例。

| 工作簿 | 调用内容 |
| --- | --- |
| 01-lodash.xlsq | 订单去重、分组、求和、排序 |
| 02-decimal.js.xlsq | 精确金额、舍入、合计 |
| 03-dayjs.xlsq | 严格日期解析、时区、日期运算 |
| 04-papaparse.xlsq | CSV 引号和换行、自动类型、导出 |
| 05-simple-statistics.xlsq | 均值、中位数、标准差、回归和预测 |
| 06-zod.xlsq | 按行校验、类型转换、可读错误 |
| 07-fuse.js.xlsq | 带拼写错误的商品搜索 |
| 08-arquero.xlsq | 表格分组、计数、聚合和排序 |
| 09-mathjs.xlsq | 表达式、单位、求导、矩阵和大数 |

所有例子均离线运行。对应库由 QuickExcel 7.0 及更新版本内置，不需要 npm
安装。示例选择界面从 7.1 开始提供。预期结果仅对应出厂输入，修改输入后不会
随之更新。`scripts/` 提供工作簿内脚本的独立源码；`common.ts` 是公共读写模块。
示例代码由 QuickExcel 编写；库自身许可证见应用设置中的第三方声明。

维护者在仓库根目录重新生成工作簿和目录：

```powershell
dotnet run --project tools/QuickExcel.SampleGen/QuickExcel.SampleGen.csproj -c Release --artifacts-path artifacts/build -- samples
```

生成器仅使用正式 XLSQ 序列化器，不执行样例脚本。每个工作簿的脚本由集成测试
在隔离 worker 中运行，并与独立写出的“预期结果”逐格对照。
