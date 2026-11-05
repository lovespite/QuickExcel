# 脚本 IStorage

2026-11-05，产品 2.1.0。API v3/v4 提供 `localStorage` 与 `sessionStorage` 全局对象，也可使用同一对象的 `host.localStorage` / `host.sessionStorage`。类型名为 `QuickExcel.IStorage`，全局别名为 `IStorage`。

```ts
interface IStorage {
    get(key: string): string | undefined;
    set(key: string, value: string): void;
    clear(): void;
    remove(key: string): void;
    keys(): string[];
    values(): string[];
    entries(): Array<[string, string]>;
}
```

| 存储 | 生命周期 | 总容量 | 单条容量 |
| --- | --- | --- | --- |
| `localStorage` | 随 XLSQ 工作簿保存、重开恢复 | 无存储专属总配额 | 1 MiB |
| `sessionStorage` | 当前文档会话内存，关闭释放 | 100 MiB | 1 MiB |

容量按 **键和值的 UTF-8 字节数之和**计量，1 MiB = 1,048,576 字节。边界值允许写入，覆盖先扣除旧条目；删除和清空释放容量。键区分大小写，允许空键、空值、Unicode 和 `__proto__` 等普通字符串。仅接受字符串，不隐式转换其他类型；不接受无法编码为 UTF-8 的孤立代理项。缺失键返回 `undefined`，删除缺失键无操作。枚举按键的 ordinal 顺序返回独立数组，修改数组不改变存储。

```ts
export function main(workbook: QuickExcel.Workbook): void {
    const count = Number(localStorage.get("runs") ?? "0") + 1;
    localStorage.set("runs", String(count));
    sessionStorage.set("lastAction", "更新计数");
    workbook.getActiveCell().setValue(count);
}
```

初始化主入口、Actions、Events 和显式控制台事务共享当前工作簿的两种存储；重新应用、停用脚本或重建 worker 不会清空会话存储。各文档会话隔离，另一个工作簿/同一文件的新会话不可读取旧会话内存。UDF 仍只从公式参数取输入，模块顶层及 UDF 中调用存储方法会失败，不引入隐式公式依赖或 UDF 副作用。

读操作立即看见本事务的先前写入。单元格、持久文件、localStorage 和 sessionStorage 经同一次宿主提交原子发布；异常、超限、取消、过期或提交拒绝全部回滚，捕获无效写入异常也不能提交已产生的部分变更。localStorage 修改标记工作簿未保存，并随同一事务撤销/重做；sessionStorage 不标记未保存，不进入撤销历史。`workbook.saved` / `workbook.calculated` 只允许会话存储写入；`workbook.close` 两者都只读。文档及受保护公式的只读门禁保持有效。

持久格式使用 XLSQ v2 的必需能力 `script-storage`，manifest 的 `localStorage` 数组逐项引用 `state/local-storage/NNNNNN.bin` 与 SHA-256。每个部件以小端 Int32 版本 1、条目数开头，随后逐条写入键字节长度、UTF-8 键、值字节长度、UTF-8 值。部件按约 4 MiB 分块；读取校验版本、长度、UTF-8、单条配额、重复键、引用顺序与摘要。无存储的既有文件打开为空，清空后保存移除能力声明及部件；旧读者通过未知必需能力门禁拒绝打开，避免丢失内容。

这些部件独立于 `host.files`，不会占用虚拟文件的 8 MiB 配额，也不能通过该接口覆盖。既有工作表结构、样式、剪贴板、计算和保存流程保留存储。带 localStorage 的主文件须使用 XLSQ；明确导出的 XLSX 是工作表副本，不含脚本存储，且不会标记主文件已保存。

“无总配额”指 localStorage 数据模型不设总容量阈值。文件读取仍遵循既有 64 MiB 压缩文件、256 MiB 累计展开及 512 MiB 解析预算；脚本仍遵循 16 MiB 整轮模型通信、4096 次模型调用、进程内存与时间预算。枚举在桥接内部分页，但最终数组和整轮累计通信仍受这些预算约束；超大枚举应改为按已知键读取并分次执行。单次持久修改还须满足现有撤销历史预算，不会绕过沙箱、通信或文件安全限制。

验证涵盖真实隔离 worker、TypeScript 编译、全部方法、Unicode/UTF-8 边界、精确 100 MiB 会话限额、超过 100 MiB 的持久数据保存重开、会话隔离、事务回滚、撤销/重做、只读事件、分页及既有通信预算、损坏包拒绝。生产自检增加存储执行、撤销/重做和保存重开检查。测试证据保存在 `artifacts/test-results/`；不运行无关完整回归。

最终验证：259 项相关测试全部通过，结果为 `artifacts/test-results/script-storage-verified.trx`；ApiGen 确定性检查、生成器自测及 JSC 原生 smoke 通过。已执行 `scripts/build-quickexcel.ps1 -SkipTests -RunSmoke`，2.1.0 包位于 `artifacts/avalonia/releases/20261105-112320-c0c10999/application`，25 项生产自检全部通过，证据为相邻的 `evidence/self-check/result.json`。`active-windows.json` 已指向该包并核对清单摘要，`Start-QuickExcel.cmd` 使用此已验证包；上一活动包保留。
