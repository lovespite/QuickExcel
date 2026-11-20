# 统一数据导入导出

QuickExcel 5.0 将 XLSX、CSV/TSV、通用 SQL 数据库和本地 SQLite 文件集中到“数据导入 / 导出”视图。原有 XLSX 值快照、可移植公式及文本导出入口保留为快捷方式。

## 工作簿与文本

- XLSX 可打开完整工作簿，或选择工作表作为数据值导入新工作簿/当前工作簿。完整工作簿沿用已有兼容性和只读保护；导入不授予脚本执行权限。
- 数据值导入允许修改目标工作表名称。同名策略为自动重命名、跳过或报错；不隐式覆盖现有表。确认前显示最终名称映射，写入当前工作簿作为一次可撤销命令，受现有 32 MiB 历史预算限制。
- 数据值交换不携带公式、脚本、图表、链接、批注、富文本分段及其他对象；保留单元格类型与数字格式。界面明确说明此范围。缺失的公式结果不能导入；旧缓存须明确选择，并且不会标记为新计算结果。
- XLSX 导出支持完整工作簿值快照、可移植公式，以及当前工作表/矩形选区的数据值副本。文本导出支持当前表/选区、原始值/显示文本、CSV/TSV 和 UTF-8 BOM。
- CSV/TSV 继续支持 UTF-8、UTF-16 LE/BE、GB18030、逗号/制表符/分号/竖线、逐列文本/数字/日期/跳过。默认文本保留前导零和以等号开头的原文。
- 导入预览绑定文档身份和修订；预览后文档发生修改，须重新预览。新导入的副本立即标记为未保存，关闭时询问保存。导出副本不改变当前文件身份、保存状态和撤销历史。

## 数据库

通用 SQL 使用系统 ODBC 驱动（64 位应用需要相应驱动），支持标准 SQL、SQL Server、PostgreSQL、MySQL 标识符/建表方言。连接串仅保存在当前对话流程内，密码输入隐藏，不写入工作簿或桌面配置。实际数据库权限由连接账户控制。

SQLite 使用随包发布的 Microsoft.Data.Sqlite / SQLitePCLRaw 原生引擎，无需安装数据库服务。读取现有文件时不自动创建缺失文件。表/视图列表与数据预览采用限定的表读取；本功能不执行任意用户 SQL 脚本。

数据库导出选择当前表或矩形选区，可将首行作为列名；支持新建表和按列名追加。新建模式遇到已有表拒绝写入，追加须显式选择。名称经过方言引用，数据使用参数，不拼接到 SQL。不存在隐式 DROP、TRUNCATE、替换或 Upsert。

写入使用事务，失败或取消回滚行数据。SQLite 的建表和插入处于同一事务；ODBC 数据库的 DDL 事务能力不同，新建失败可能留下空表，界面明确说明。MySQL 追加要求事务表。未实际连接的外部 SQL Server/PostgreSQL/MySQL 服务不计为已通过互操作验证。

导入/导出受工作表行列、单元格数和数据预算限制，支持取消、命令超时和错误反馈。数据库中的 NULL 对应空白；文本不解析为公式。无时区日期转换为日期格式的序列值；带时区日期和超出浮点精度的整数/小数保留为文本。二进制字段拒绝导入，须先在数据库中显式转换。数据库导出的混合类型列转换为文本，错误值和虚构日期拒绝写入。SQL 数据不是工作簿格式，工作簿对象和格式不随数据库表保存。

## 验证

定向契约测试：`WorkbookDataTransferTests`、`DatabaseTransferTests`、既有 `TextImportTests` / 文本交换测试；桌面流程通过 DataTransfer 与 TextExchange smoke 检查。生产自检另行检查发布包 SQLite 原生依赖的建表、Unicode 写入和读回。数据库账户、远端驱动、真实数据库服务器和真实物理输入不由这些本地检查替代。

Windows 部分裁剪必须保留 `System.Data` 类型转发程序集：ODBC 内置的 [元数据 XML](https://raw.githubusercontent.com/dotnet/runtime/v10.0.15/src/libraries/System.Data.Odbc/src/Resources/System.Data.Odbc.OdbcMetaData.xml) 通过此程序集名称引用 `GroupByBehavior`、`IdentifierCase` 和 `SupportedJoinOperators`，缺失时 `GetSchema("Tables")` 无法建立表列表。生产自检读取这一内置资源并检查三个类型，独立于本机是否安装驱动。JSON 反射回退仍禁用，`DataSet.XmlSerializationIsSupported` 保持原有关闭设置；没有通过启用宽泛的序列化开关绕过裁剪问题。

新增组件及上游许可证见仓库 `THIRD-PARTY-LICENSES.md` 和发布包 `Licenses`；设置页可离线查看。

2026-11-20 定向验证：Release 构建零警告/零错误，65 项相关契约测试通过（`artifacts/validation/data-transfer/data-transfer.trx`）；统一视图 24 项、文本交换 11 项检查通过（`artifacts/validation/data-transfer-final-9ae2df36`），菜单命令 29 项通过（`artifacts/validation/data-transfer-smoke-e8f1e805/commands`）。新增工作簿导入的未保存状态、取消关闭、完整 XLSX 副本往返、SQLite 类型保护和约束回滚均在这些检查中覆盖。

CSV/TSV 新建的数据工作簿也采用 API v4，未包含源脚本的导出投影使用 XLSX 的规范空模块表示；既有工作簿的源代码和 API 版本不改写。此补充通过 10 项文本导入检查（`data-transfer-text-api4.trx`）。发布裁剪仍报告上游 Markdown/ODBC 反射注解警告，未全局抑制；以新包的生产自检和真实 ODBC 路径检查为声明范围内的运行证据。

最终 Windows 5.0.0 包为 `artifacts/avalonia/releases/20261120-050000-719bfea4/application`，36 项生产自检通过，活动指针与清单 SHA-256 已核验，`Start-QuickExcel.cmd` 使用此包。SQLite 原生引擎、ODBC 元数据和新增许可证齐全，数据库组件未进入隔离脚本 worker。

发布后使用本机 64 位 Access Text ODBC 驱动读取独立测试 CSV：表枚举、直接 `GetSchema` 和两行两列数据导入通过，前导零 `001` 与数字 `12.5` 保持正确。证据为 `artifacts/data-transfer/odbc-probe/final-exact/result.txt` 和 `verification.json`；被测副本中的原发布文件逐一核对 SHA-256，没有注入程序集或修改依赖清单。远端 SQL Server/PostgreSQL/MySQL 的写入和驱动互操作仍需实际连接验证。
