# 脚本事件与会话虚拟文件

2026-11-03。文件仍使用 XLSQ v2 / scriptApiVersion 2；宿主与 ScriptWorker 必须同时升级到 IPC **5**。API v1 公式语义保持不变。编辑器的「操作与事件入口」对应现有 `workbook/actions` 绑定与 `script-actions` 能力。

## 使用

在脚本编辑器中把模块设为「操作与事件入口」，然后应用脚本。事件与动作共享已有的签名/显式授权流程；未授权打开文件不会运行模块或事件。受信任签名仍可通过已有策略启用。事件注册不会出现在手动操作列表中。

```javascript
quickExcel.events.on('workbook.open', ({ fs }) => {
  fs.mkdir('//xlsq/data');
  fs.writeJson('//mem/session', { edits: 0 });
});

quickExcel.events.on('cell.commit', ({ event, snapshot, patch, fs }) => {
  const state = fs.exists('//mem/session')
    ? fs.readJson('//mem/session') : { edits: 0 };
  state.edits++;
  fs.writeJson('//mem/session', state);
  fs.writeJson('//temp/last-edit.json', {
    sheet: event.sheetId, ranges: event.ranges, revision: event.revision
  });
  // 仍通过 patch 写单元格；同一个事件内的文件与单元格一次性提交。
  // patch.setValue(snapshot.sheets[0].id, 'B1', state.edits);
});

quickExcel.events.on('workbook.beforeSave', ({ fs }) => {
  if (fs.exists('//mem/session'))
    fs.writeJson('//xlsq/data/statistics.json', fs.readJson('//mem/session'));
});

quickExcel.actions.register('读取统计', ({ fs }) => {
  console.log(fs.readJson('//xlsq/data/statistics.json'));
  console.log(fs.list('//xlsq/', { recursive: true }));
});
```

## 事件约定

| 事件 | 时机 |
| --- | --- |
| `workbook.open` | 会话首次启用动作/事件脚本时补发一次；不回放授权前的用户操作 |
| `workbook.activate` / `workbook.deactivate` | 主工作簿窗口获得/失去激活，或工作簿视图与首页之间切换 |
| `workbook.beforeClose` | 关闭/切换尝试，已提交编辑后、未保存提示前；处理器的永久修改纳入保存决策；取消后会话保留，下次尝试可再触发 |
| `workbook.close` | 成功替换会话或正常退出时、清理前；只读完成通知，不能再写单元格或文件；异常不阻止清理 |
| `sheet.added` / `sheet.deleted` / `sheet.renamed` | 对应命令或撤销/重做完成；删除事件携带删除前的稳定 ID 和名称 |
| `sheet.activate` / `sheet.deactivate` | 活动表切换，包括删除当前表后的切换；首次启用时报告当前活动表 |
| `cell.click` | 网格指针点击后的单元格/选区（包括区域表头选择），再次点击同一格也通知 |
| `cell.edit` | 网格或公式栏进入草稿编辑 |
| `cell.commit` | 网格或公式栏成功提交，携带原草稿所属单元格；失败不通知 |
| `selection.changed` | 鼠标或键盘改变选区 |
| `cells.changed` | 用户命令/撤销/重做更改单元格内容；每张表给出受影响单元格的外包矩形，可能包含未更改格 |
| `workbook.beforeSave` | 桌面保存捕获之前；处理器的永久文件/单元格修改包含在此次保存中 |
| `workbook.saved` / `workbook.calculated` | 成功保存/计算后；可读内容和读写临时/内存卷，不允许改持久输入，避免保存后再变脏或计算循环 |

`event` 为冻结对象：`name`、`sessionId`、`sequence`、`revision`、`sheetId`、`sheetName`、`ranges`、`isUndoOrRedo`。范围为 A1 字符串，最多 64 个；工作簿事件无工作表/范围。`event.revision` 是通知时修订，`snapshot.baseRevision` 是本次实际执行快照修订。

处理器是**异步调度的同步通知**，不是拦截器：不提供取消用户操作的返回值。桌面在编辑、IME 组合、拖选结束后的空闲时刻按入队顺序运行，`cell.edit` 等事件可能在草稿完成后送达；事件不能直接访问未提交草稿。每个事件最多 32 个处理器，按注册顺序在同一事务中运行；任一个失败则整次事件回滚。单元格快照在整个事务中不变，`fs` 则可读到本事务之前的文件写入。不同事件/动作使用新的 JS 实例，需要跨次数据时使用 `mem` 或 `temp`。

脚本生成的变更不再生成 `cells.changed`/结构事件。队列仅保留已注册事件，最多 256 条；溢出清空并记录诊断。源码/绑定变更或停用会清空待处理事件；关闭放弃剩余交互通知，仅发送 close。失败、超时、授权撤销或用户新修改导致的过期结果不重试。异常显示在现有诊断/状态栏。

## 虚拟文件 API

动作和事件上下文均有 `fs`；公式函数和模块顶层没有文件访问对象。

| 方法 | 语义 |
| --- | --- |
| `exists(path)` / `stat(path)` | 查询存在性或 `{path, type: 'file'|'directory', size}`（字节） |
| `list(directory, {recursive?: boolean})` | 默认直接子条目，按路径排序；支持递归枚举 |
| `mkdir(path)` | 创建目录；父目录隐式存在，显式空目录也会保留 |
| `readText(path)` / `writeText(path, text)` | 严格 UTF-8 文本，支持中文及补充平面字符 |
| `readJson(path)` / `writeJson(path, value)` | JSON 文本便捷方法 |
| `readBytes(path)` / `writeBytes(path, bytes)` | 独立 `Uint8Array` 副本，可存任意二进制/内存块 |
| `copy(from, to)` / `move(from, to)` | 文件复制/移动，可跨虚拟卷；覆盖目标文件，不覆盖目录 |
| `remove(path)` | 删除文件或显式空目录；不隐式递归删除 |

路径严格使用 `/`，只接受 `//xlsq/`、`//temp/`、`//mem/`。除卷根外不带尾随 `/`。不解析相对路径、`.`、`..`、反斜杠、URL 编码、驱动器、UNC 主机、ADS、查询/片段或控制字符；拒绝大小写歧义、文件/目录冲突。路径区分大小写，最多 248 个 UTF-16 单元。没有真实 OS 路径映射、工作目录、环境变量展开、符号链接或原始文件句柄。

* `//xlsq/`：当前工作簿的规范包投影，层级与外层 ZIP 对应。包含 `manifest.json`、`workbook.xlsx`、`state/document.json`、全部声明模块、存在的签名以及附加文件。`workbook.xlsx` 是二进制文件，不把内层 XLSX 的 ZIP 自动展开为目录。附加文件和显式空目录使用新必需能力 `virtual-files`、manifest `files` 清单及文件 SHA-256；旧读者会拒绝，避免保存时丢失。
* `//temp/`：会话临时文件，使用内存后端，不写宿主磁盘、不进入包，关闭/退出即释放；崩溃也没有磁盘残留。
* `//mem/`：会话二进制块/交换数据/缓存，使用同样的有界内存后端；跨动作可见，跨会话不共享。

包内容由当前不可变工作簿生成；修改系统文件时重新解析完整包，自动重算引用 SHA-256 和脚本集合摘要。会话/内容修订身份不可改写，清单路径与必需条目仍须有效。已建模的元数据以模型的规范序列化为准，不保证 JSON/XML 格式、空白或 ZIP 字节布局不变。脚本源码/绑定修改递增脚本代次、移除旧签名并撤销授权，沿用源码不进入单元格撤销历史的约定。签名变动也撤销授权。原始 `workbook.xlsx` 替换不能和单元格 patch 混在同一事务，以免对不同布局/样式表套用旧快照坐标。非法系统文件和不能安全重写的未建模 workbook.xlsx 整体拒绝；不绕过只读公式、保护、验证、富文本/链接/批注保留规则。

每轮文件快照总计最多 **8 MiB、4096 条目**（含包系统文件及两个临时卷）；文件操作最多 4096 条，写入 IPC 的累计二进制数据也最多 8 MiB。已有 16 MiB IPC/native 桥接总帧预算、工作簿单元格快照预算和 5 秒 Worker 生命周期继续生效：文件与单元格合计编码超限时，整轮拒绝。大包不能通过此快照 API 执行动作/事件；不截断条目或降级为真实文件访问。

永久文件与单元格原子提交、标脏、保存并一起撤销/重做；`temp`/`mem` 仅在整轮成功后交换会话状态，不标脏、不进入撤销历史。读取失败可以由脚本处理；写入/删除失败即使被脚本 catch 也使整轮写入失效。源码改写沿用独立的源码历史规则；源码/绑定/签名修改不能与单元格或其他永久数据文件混在同一事务。

## 隔离与验证

Windows 仍是零 capability AppContainer + Job，Linux 仍是既有 Landlock/seccomp/systemd 限制；没有新增 ACL、真实目录授权或同权限回退。Worker 只收有界字节副本、只返回操作描述。宿主重新校验路径、类型、配额、快照身份、代次、取消、格式及撤销预算后才提交。唯一实际持久化入口仍是用户已有保存操作。平台隔离与物理 IME/DPI/辅助功能等既有未完成门槛不因本功能关闭。

针对性测试包含 `ScriptEventFileTests`、既有 M4 actions/Worker/模块与 typed storage/v1 兼容测试。`scripts/test-avalonia-workspace.ps1 -Mode Automation` 增加指针点击、编辑提交、工作表切换、保存/关闭事件与虚拟文件保存检查；这是 Avalonia handler 自动化，不是物理输入验证。类型提示位于 `QuickExcel.Web/src/quickexcel.d.ts`。

本次本地验证：Release 解决方案构建 0 警告/0 错误；新增事件/文件测试 35 项及相关模块、动作、存储、Worker 故障/会话回归分批通过。旧模块夹具的 VALUE 测试函数与已存在的内置函数冲突，已改为 MODULEVALUE，未放宽冲突检查。Avalonia Automation 72 项通过，证据为 `artifacts/avalonia/workspace/20261103-030234-177a0363/result.json`，包含打开编辑器时通过虚拟文件改写源码后的刷新与授权撤销。未运行全量回归，也未运行物理输入或跨平台隔离认证。
