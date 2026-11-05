# QuickExcel.Core

Core 只依赖 .NET，负责工作簿领域模型、文档会话、命令事务和外部服务契约。存储、计算执行、脚本 worker 和 UI 的具体实现留在各自项目中。

## 目录与源码入口

| 目录 | 职责与主要入口 |
| --- | --- |
| `Domain/` | 不可变快照、单元格类型及工作表对象的数据定义；入口为 `WorkbookSnapshot.cs`、`WorksheetData.cs` |
| `Documents/` | 文档 facade、会话生命周期、修订、撤销重做、保存与计算协调；入口为 `WorkbookDocument.cs`、`DocumentSession.cs` |
| `Ports/` | 存储、计算和脚本运行时的注入契约及请求/结果类型；入口为 `WorkbookPorts.cs` |
| `Localization/` | 宿主提供的只读消息翻译目录；保留 Core 独立调用的默认文案与格式参数 |
| `Commands/` | 命令基础类型、通用单元格编辑及跨功能的工作表/名称/结构/样式/视图/保护操作；入口为 `DocumentCommands.cs`、`StructureCommands.cs`、`PresentationCommands.cs` |
| `Clipboard/` | 内部、外部及图表剪贴板模型和粘贴命令 |
| `Data/` | 清洗、排序描述、筛选、分组、分类汇总、数据验证和选区统计；包含对应的功能命令 |
| `Data/Tables/` | 表格创建、编辑、结构引用维护与样式呈现 |
| `Data/Pivots/` | 透视表命令、汇总与报表生成 |
| `Formatting/` | 显示格式、小数位调整和条件格式；`CellDisplayFormatter` 的所有分部文件放在一起 |
| `Formulas/` | 公式语法、绑定、复制、函数语义与动态数组溢出规则；计算器实现仍在 Infrastructure |
| `Objects/` | 图表、图片、绘图对象及批注操作，和未完整支持对象的编辑保护 |
| `Printing/` | 打印设置命令、分页及结构变化时的打印引用维护 |
| `Scripting/` | 动作、事件、控制台、宿主事务模型、脚本专用命令与逻辑虚拟文件；不包含脚本引擎实现 |

## 新增与维护约定

- 根目录只保留项目文件和本说明；新增源码放入对应职责目录。
- 专属某项功能的命令与该功能放在一起；命令基础设施、通用编辑和跨功能操作放在 `Commands/`，避免形成新的命令文件堆积。
- 工作表对象的数据定义放在 `Domain/`，修改这些对象的命令放在相应功能目录。
- 目录用于组织源码，不改变现有公共命名空间：领域类型使用 `QuickExcel.Domain`，既有公式类型使用 `QuickExcel.Formulas`，其余类型保持 `QuickExcel.Core`。其中 `Formulas/FormulaSpills.cs` 和 `FormulaSpillSafety.cs` 也保留 `QuickExcel.Core`。
- `DocumentSession` 的基础与计算分部在 `Documents/`；动作事务、事件和控制台分部随功能留在 `Scripting/ScriptWorkbookModel.cs`、`ScriptEvents.cs`、`ScriptConsole.cs`。
- 文件移动时同步更新文档链接。项目使用 SDK 默认的递归源码包含规则，无需手工列出 `Compile` 项。

构建、定向验证及本地启动包更新要求见仓库 [AGENTS.md](../AGENTS.md)。
