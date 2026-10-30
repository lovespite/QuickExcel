# XC-0303 结构化引用与表编辑

2026-10-31 完成声明范围，接续数据验证。结构化引用进入 Core AST 和稳定依赖索引，计算列、表/列改名、插删、复制、扩表使用同一命令与历史路径。未关闭其他发布或 M4/M6 门禁。

## 范围与约定

支持表名、单列/连续多列、当前行、#Data/#Headers/#Totals/#All 及可连续组合的区域，支持带空格和特殊字符的列名、跨表和定义名称。绑定到实际 Table 身份和几何范围后加入稀疏依赖；源公式仍保存结构化 AST。改名和结构变化改写节点，不以字符串替换引用。无汇总行的 #Totals 返回 #REF!，表数据区之外的当前行返回 #VALUE!；空表的占位行保留 Excel 的零/空值语义。

OOXML 中列名开头的 @ 是普通字符，Excel UI 的 @ 则表示当前行；存储编解码分别处理，且函数别名转换跳过列名中的括号/函数样式文本。原始受保护公式仍保留。非连续列联合、无法绑定的对象和未建模数组语义沿用只读/兼容报告，不凭缓存冒充重算。

数据页提供表改名、列改名和计算列入口。设置计算列明确替换整列（包括例外格），作为一次可撤销事务；普通编辑只改变该格。插入新记录或显式扩展底边为新增空白格填充模板，不覆盖已有空白例外。收缩后留在表外的公式补充原表限定。没有隐式的“在表外输入即扩表”。

插删行列保持列身份、筛选、计算列模板与依赖；新增列用本次中文 Excel 对照名称“列1”等。删除引用的列端点产生 #REF!。复制工作表创建独立表名/编号，并重绑定内部引用；把局部结构化公式复制到表外补充表名。删除整张表所属工作表将引用变为 #REF!。表公式模板也参与定义名称改名。

删除无汇总表的最后一条记录保留 Excel 插入占位行和计算列模板；输入首条新记录时填充模板，与输入一起撤销。受限边界：不允许删除仍保留的表头/汇总行；含汇总表至少保留一条数据行；有汇总行时用数据区插删调整，底边调整暂拒绝；整行列搬移仍拒绝。未实现的变换在提交前整体拒绝，不留下部分修改。

## 验证

Excel 16.0 build 20527 x64，语言 2052。任务自有隐藏 COM 实例记录 PID/版本和 SHA-256，禁用宏及外链更新，普通打开、完全重算、保存重开；不使用修复模式。原生初始包 XML 和十阶段期望保存于 Fixtures/excel-structured-tables.json，二进制和中间失败证据只在 artifacts。

| 证据（artifacts/compatibility） | 结果 |
| --- | --- |
| xc0303-excel-04 / xc0303-probe-06 | 原生十阶段：初始、改名、插行、插列、删除新增行列、复制表、复制公式、扩表、首行插入、删除引用列；含特殊列名和空表。20,891 项导入、命令、绑定、撤销/重做及 XLSX/XLSQ 往返检查，零差异；输出 schema 零问题 |
| xc0303-output-01 / xc0303-output-verdict.json | 前版相同十阶段输出经 Excel 普通打开、重算、保存重开，2,226 项检查零差异 |
| xc0303-output-02 / xc0303-ui-output-01 / xc0303-output-02-verdict.json | 增加空表的最新初始/删除列文件，以及桌面创建文件，经 Excel 复验；514 项检查零差异，输入文件哈希不变 |
| xc0303-empty-01 | Excel 删除最后数据行的独立实验：占位行、insertRow 和 A1 计算列模板保留；相同行为及重新输入/撤销已有回归测试 |
| xc0303-wpf-02 | 14 项真实 WPF 处理器检查：创建/计算列、改名、扩表、实际插行入口、复制、冲突原子拒绝、普通保存重开、原生文件打开；已检查数据页截图 |
| xc0303-table-basics-01 / xc0303-workspace-01 | 12 项 Table 基础与 153 项工作区 WPF 回归通过 |
| 最终相关测试 | 170 项结构化引用、Table、保留公式、兼容报告、工作表/名称、计算、剪贴板和会话测试通过 |

原生包的 calcChain 在现有 Open XML 审计器中有已知样式索引误判；探针初始写出显式诊断副本，其他未知元数据触发保护时也记录副本路径。该证据不代表任意来源文件均可无损普通保存。自行构造的受支持模型、XLSX/XLSQ 及桌面文件另走普通保存。Release 全解决方案构建成功，仅现有 NU1900 漏洞源网络警告；未跑全量回归或发布包装。WPF 检查不等于物理键盘/IME 验收。

复现：scripts/test-excel-structured-tables.ps1 -OutputDirectory <新目录>；CompatibilityProbe 的 structured-tables <原生目录> <新目录>；桌面 --structured-tables-smoke <新目录> [原生文件]。所有托管命令附带 --artifacts-path artifacts/build。早期失败保留于 excel-01/02、probe-01/02/03/05：已分别修正 Excel 文件锁、自动汇总样本、@ 列名、空表 insertRow 和删除端点语义。

参考：[Excel 结构化引用](https://support.microsoft.com/en-us/excel/using-structured-references-with-excel-tables)、[Open XML calculatedColumnFormula](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.calculatedcolumnformula)。
