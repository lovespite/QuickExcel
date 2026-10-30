# XC-0001：真实 Excel 业务样本台账

2026-10-30 完成 XC-0001，提交为本记录所在的 `test: complete real Excel business sample ledger (XC-0001)`。当前 R/D/C/E/W 矩阵见 [27](27-xlsx-compatibility-baseline.md)；本记录补齐五类业务样本和 E1 各项的真实 Excel 证据索引。XC-0002 的完整预期差异判定、XC-0003 的剩余定位边界仍单独推进，不据此关闭整个 E0。

## 业务样本

这些文件由 `scripts/test-excel-business-baseline.ps1` 调用本机 Excel 独立创建，不经过 QuickExcel 写入器，也不来自用户工作簿。每次使用新的 artifacts 目录；来源文件哈希、Excel 版本/build/语言/计算模式、日期系统、原值/公式/显示、保存重开结果均写入 JSON。固定捕获每表 A1:D6（包含空白），不把遗漏的空白区误当作已验证。

| ID | 固定输入与关键期望 | 维度 |
| --- | --- | --- |
| `roster` | 文本编号 `00123`/`00007`、中文姓名、布尔 TRUE、换行文字；编号保持 Text | R/D/E/W 基线；没有公式 |
| `quotation` | 单价 1234.567×2，ROUND 后 2469.13；-42.5×3=-127.5；合计 2341.63；货币三位小数和负数红色格式 | R/D/C/W；业务修改路径归 XC-0002 |
| `date-ledger-1900` | 序列 60 显示 1900-02-29，45292.5 显示 2024-01-01；0.5/1.5 显示 12:00:00/36:00:00 | R/D/W；不将日期格式当作新数据类型 |
| `date-ledger-1904` | 同样的序列 60 显示 1904-03-01，45292.5 显示 2028-01-02；时长不受日期系统影响 | R/D/W；与上一文件独立 |
| `cross-sheet` | 数据表 A2=21、B2=3；汇总跨表乘二为 42、SUM 为 24；保留公式文本 | R/D/C/W；引用变换归 XC-0002 |
| `print-template` | 12×5=60；A4 横向、适应 1×1 页、打印区域 A1:D6、重复第 1 行、页眉“业务基线”、页脚页码 | R/D/C/W；打印元数据 R/W，实际打印仍属 E4 |

六份文件共七张表、168 个格位置。`business-baseline` 工具在导入、未修改 XLSX 保存重开、XLSQ 保存重开、QuickExcel 自身重算四个状态比较类型、值、公式、格式器显示文本、可确定的整格字体/字号/粗斜体；合计 **5384/5384 属性断言**。这是属性断言数，不是 5384 个独立场景。实际 Excel 在来源与 QuickExcel 原包输出中还对照捕获的数字格式、字体颜色、工作表可见性和打印属性。

Excel 对混合中西文字体可能返回空整格 Font.Name；该字段保留在原始 JSON，不当作字体相等断言。逐字符字体及网格像素证据复用 XC-0102/0107。显示文本对照不是跨 DPI 截图验收。公式缓存导入仍为 stale；只有工具实际调用内置计算器成功后检查 Fresh，未执行工作簿脚本。

## E1 最小样本与期望索引

所有路径均在 `artifacts/compatibility` 下，二进制样本不进入 Git。每个专题文档提供脚本、明确子集及期望比较方式；合成 XML 拒绝夹具继续由 `FixtureCatalog`/Core.Test 内联构造，来源不同，不能互换验收结论。

| E1 项 | 真实 Excel 来源 / 最终对照 | 期望与限制入口 |
| --- | --- | --- |
| XC-0101 数字格式 | `xc0101-excel-07` / `xc0101-reopen-02` | [30](30-excel-number-formats.md)：43 个固定原值、格式、显示/颜色，已存版本化 JSON |
| XC-0102 样式/主题 | `xc0102-excel-02` / `xc0102-excel-reopen-02` | [31](31-xlsx-named-styles-and-theme.md)：8 格 104 属性，指定修改以外不变 |
| XC-0103 属性/保存 | `xc0103-excel-07` / `xc0103-excel-reopen-02` | [32](32-xlsx-save-and-metadata.md)：10 项属性、公式与重算结果 |
| XC-0104 视图/可见性 | `xc0104-excel-01` / `xc0104-excel-reopen-01`、`xc0104-excel-insert-01` | [33](33-xlsx-workbook-views.md)：六表视图，插行按 Excel 自身操作对照 |
| XC-0105 打印元数据 | `xc0105-excel-reopen-final`、`xc0105-excel-insert-final` | [34](34-xlsx-page-and-print-metadata.md)：三表 144 属性/列表和结构变换 |
| XC-0106 公式保留 | `xc0106-excel-readonly-03`、`xc0106-excel-unknown-01` / `xc0106-excel-readonly-reopen-01`、`xc0106-excel-unknown-final` | [35](35-xlsx-preserved-formulas.md)：只读数组/结构化/溢出与未知函数；源 calcChain 的 4 条 schema 发现保留，不称其全通过 |
| XC-0107 富文本/链接/批注 | `xc0107-excel-source-final` / `xc0107-excel-reopen-final-02` | [36](36-xlsx-rich-text-links-and-notes.md)：17 字符、3 种链接、2 条批注；VML 几何 0.75 pt 容差，其余精确 |

后续 E2/E3 在此业务基线之上增加自己的边界样本和 Excel 固定期望；台账不会把未实现的 Table、验证或条件格式标为已支持。

## 本项证据与复现

Excel 16.0 build 20527、64 位、语言 2052，自动计算 -4105。隐藏独立 COM、宏禁用、不更新外部链接；以普通模式打开并保存重开，没有请求修复。只关闭本脚本创建的实例。

- `xc0001-excel-business-03/excel-business-baseline.json`：六份来源，Excel 自身重算、保存重开全部一致。
- `xc0001-business-probe-final/business-baseline-report.json`：5384 属性通过，六份来源和输出 schema 均 0 问题，原 XLSX 字节保持，来源哈希不变；所有六个业务公式实际重算为 Fresh。
- `xc0001-business-excel-reopen-01/excel-business-baseline.json`：六份 QuickExcel 保存输出由 Excel 打开、重算、保存重开，与来源逐项相等。输入为 `xc0001-business-probe-01`，与最终工具输出的六份 XLSX 均为相同来源原字节。
- `xc0001-business-tests`：新增工具专项 4/4，通过故意改公式、改值和格式导致显示变化验证定位失败；本项不改产品代码，未运行全量套件或 WPF。
- 前两次生成因本机 COM 数字格式颜色关键字使用本地语言而失败；证据目录保留。脚本按当前文化处理颜色词，最终来源生成通过。构建仅既有 NU1900 警告，未发布。

```powershell
.\scripts\test-excel-business-baseline.ps1 -OutputDirectory artifacts/compatibility/new-business
dotnet run --project tools/QuickExcel.CompatibilityProbe/QuickExcel.CompatibilityProbe.csproj -c Release --artifacts-path artifacts/build -- business-baseline artifacts/compatibility/new-business/excel-business-baseline.json artifacts/compatibility/new-business-probe
.\scripts\test-excel-business-baseline.ps1 -OutputDirectory artifacts/compatibility/new-business-reopen -InputDirectory artifacts/compatibility/new-business-probe -ReferencePath artifacts/compatibility/new-business/excel-business-baseline.json
```

此项完成样本/矩阵与原包基线；普通改值、改样式、插删行列、改名后的预期差异判定在 XC-0002 单独验收，不能用原包相等替代编辑后保真。
