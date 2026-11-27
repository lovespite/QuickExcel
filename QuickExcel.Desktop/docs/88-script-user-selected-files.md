# 工作簿虚拟文件与用户选择的本机文件

2026-11-27。产品 8.0.0 将工作簿虚拟文件移至 `workbook.files`。8.0.1 加入 `host.fs` 编辑器补全和编译期检查。9.0.0 在 Windows 的 v4 用户 Action 中接入本机选择器和文件读写，使用 IPC 14。

## 已实现的工作簿虚拟文件入口

`workbook.files` 属于当前工作簿。它提供同一套同步方法，例如 `readText`、`writeBytes`、`list` 和 `remove`。`//xlsq/` 是工作簿包投影，`//temp/` 和 `//mem/` 仅在该文档会话内存在。三个卷都参与脚本事务；临时卷可跨成功的 Action 共享，关闭文档即释放。[虚拟文件契约](69-script-events-and-virtual-files.md) · [脚本对象 API](79-script-api-redesign.md)

| 旧写法 | v3 / v4 的新写法 |
| --- | --- |
| `host.files.readText("//xlsq/data.txt")` | `workbook.files.readText("//xlsq/data.txt")` |
| `context.host.files.writeJson("//mem/state", value)` | `context.workbook.files.writeJson("//mem/state", value)` |

旧的 `host.files` 属性已移除。Action、v4 初始化入口、事件和调试控制台均使用相应事务里的 workbook。UDF 与模块顶层无法访问虚拟文件。打开旧工作簿保留源码和签名原文；原先调用旧入口的 JS 或 TS 需由作者更新，已签名脚本修改后需要重新签名。

Windows x64 定向验证通过 209 项，覆盖真实 Native AOT worker、类型声明、跨轮失效、事务回滚和 Papa Parse 的用法提示。Monaco / Avalonia 脚本检查通过 220 项。8.0.0 发布包通过 41 项生产自检，其中一项用真实 TypeScript Action 写入 `//xlsq/`、保存重开并撤销。活动发布目录为 `artifacts/avalonia/releases/20261127-110608-cee186c8/application`，定向测试使用的 worker 与发布包内文件的 SHA-256 一致。

## 本机文件系统

Windows 目录选择器使用 Shell 的现代文件夹对话框，提供地址栏、侧边导航、搜索和多选。宿主支持本地固定磁盘和可移动磁盘上的普通文件与目录。提交文件草稿需要支持事务写入的卷，已在本地 NTFS 验证。网络共享、重解析点和不受支持的文件系统操作返回错误。macOS、Linux 尚无本机文件授权适配器，调用选择器返回 `UnsupportedApi`。工作簿虚拟文件沿用原有平台支持。

在 v4 的脚本编辑器中输入 `host.fs.`，可补全两个选择器。TypeScript 入口的 `host` 参数使用 `QuickExcel.Host`。新建 JavaScript 入口已带类型注释；已有 JavaScript 若用未标注的局部参数 `host`，可在入口函数的上一行添加 `/** @param {QuickExcel.Host} host */` 以启用成员提示。

一次典型导入只需三步：用 `await host.fs.getUserSelectedFiles(...)` 让用户选择文件，用 `await file.open()` 打开它，再用 `stream.readLine("utf-8")` 逐行读取。导出时，脚本把内容写入私有草稿，`await stream.close()` 成功才发布文件。用户选中目录时，授权覆盖其下整棵普通文件和目录树。脚本可把读到的内容交给 `workbook.files`，例如将文本存入 `//xlsq/import.txt`。[Actions 与宿主代理](85-script-fetch-and-network.md)

## 选择器与授权

```typescript
declare namespace QuickExcel {
  type FileMode = "open" | "create" | "open_or_create";
  type FileAccess = "r" | "w" | "rw";
  type Encoding = "utf-8" | "utf-16le" | "utf-16be";

  interface FileFilter {
    readonly name: string;
    readonly extensions: readonly string[];  // 如 [".csv", "tsv"]
  }

  type FileSelectOptions = {
    title?: string;
    filters?: readonly FileFilter[];
  } & (
    | { mode: "open"; multiSelect?: boolean }
    | { mode: "save"; multiSelect?: false; suggestedName?: string }
  ) & (
    | { fileMode: "open" | "open_or_create"; fileAccess: FileAccess }
    | { fileMode: "create"; fileAccess: "w" | "rw" }
  );

  interface DirectorySelectOptions {
    title?: string;
    fileAccess?: FileAccess;  // 默认 "r"
    allowDelete?: boolean;    // 默认 false，与读写权限分别授予
    multiSelect?: boolean;    // 默认 false
  }

  interface Host { readonly fs: UserFileSystem; }
  interface UserFileSystem {
    getUserSelectedFiles(options: FileSelectOptions): Promise<readonly FileHandle[]>;
    getUserSelectedDirectory(options: DirectorySelectOptions & { multiSelect: true }):
      Promise<readonly DirectoryHandle[]>;
    getUserSelectedDirectory(options?: DirectorySelectOptions & { multiSelect?: false }):
      Promise<DirectoryHandle | null>;
    getUserSelectedDirectory(options: DirectorySelectOptions):
      Promise<DirectoryHandle | readonly DirectoryHandle[] | null>;
  }
}
```

文件选择始终返回数组，默认单选；保存成功时数组长度为 1。目录默认返回一个对象，多选返回数组。用户在选择器中点“取消”分别得到空数组或 `null`，Action 可正常继续。用户在 QuickExcel 中终止整个 Action 则取消执行。宿主在弹窗前校验选项，包括从 JavaScript 传入的非法组合。标题、过滤器和建议名只决定界面呈现，建议名只接收单个文件名。

从 9.0.1 起，`filters[].extensions` 的每一项可带一个前导点。例如 `extensions: [".xlsq"]` 和 `extensions: ["xlsq"]` 都匹配 `*.xlsq`。扩展名主体由 1 至 24 个 ASCII 字母或数字组成。传入通配符（如 `"*.xlsq"`）会在弹窗前得到 `InvalidArgument`。

| `fileMode` | 目标已存在 | 目标不存在 |
| --- | --- | --- |
| `open` | 保留原内容，从字节 0 开始 | `ItemNotFound` |
| `create` | 首次打开得到空草稿，须有写权限 | 首次打开得到空草稿，须有写权限 |
| `open_or_create` | 保留原内容，从字节 0 开始 | 有写权限时准备空草稿，只有读权限时报 `PermissionDenied` |

`mode` 决定显示哪个选择器：“打开”选择已有文件，“保存”还可以指定新文件名。`fileMode` 决定拿到句柄后如何处理内容。对已选文件申请 `create` 时，选择界面须清楚确认将被覆盖的目标，多选同样适用。选择、查看属性、打开和放弃草稿均不创建或截断目标。目录选择器选取已有根目录。宿主向用户展示工作簿、Action 名、读写范围、是否可覆盖文件及删除子项。脚本标题作为补充文字。一次选择授权只持续到本次 Action 结束。

## 文件对象与目录对象

```typescript
declare namespace QuickExcel {
  interface FsEntryInfo {
    readonly type: "file" | "directory";
    readonly name: string;                  // 单个显示名
    readonly exists: boolean;
    readonly size: number | null;           // 文件字节数
    readonly createdAt: string | null;      // 带 Z 的 ISO 8601 UTC 时间
    readonly modifiedAt: string | null;
  }

  interface FsHandle {
    readonly name: string;
    stat(): Promise<FsEntryInfo>;
    release(): Promise<void>;
  }

  interface FileHandle extends FsHandle {
    readonly type: "file";
    open(): Promise<BufferedFileStream>;
    open(options: { buffering: "stream" }): Promise<StreamingFileStream>;
    openWorkbook(): Promise<SelectedWorkbook>;
  }

  interface DirectoryHandle extends FsHandle {
    readonly type: "directory";
    entries(): AsyncIterable<FsEntryInfo>;
    getFile(name: string, options?: { fileMode?: FileMode; fileAccess?: FileAccess }):
      Promise<FileHandle>;
    getDirectory(name: string, options?: { create?: boolean }): Promise<DirectoryHandle>;
    removeEntry(name: string, options?: { recursive?: boolean }): Promise<void>;
  }
}
```

- `name` 和 `type` 可直接读取。`stat()` 查询当前已发布状态。目录大小和平台未提供的时间返回 `null`。尚未创建的保存目标返回 `exists: false`，大小、时间均为 `null`；已存在的空文件大小为 0。原有对象被调换后，旧句柄报告失效。流的 `length` 单独反映草稿。
- `entries()` 分页枚举直接子项，返回元数据。访问孙辈时在子目录句柄上继续调用相同方法，无需再次选择。仅接受一个直接子项名，不接收绝对路径、相对路径串或父目录引用。枚举不保证排序；读目录期间发生修改时允许报告冲突。
- `r` 允许列举和读取；`w` 允许按已知名称打开、创建和写入；`rw` 合并两者。基本元数据随取得的句柄可读。`getFile` 默认 `open` 并继承目录的读写权限，显式请求只能收窄。子文件沿用上表的打开规则；目录写授权涵盖其下文件覆盖。`getDirectory(name, { create: true })` 在已有同名目录时返回该目录，新建子目录须有写权限，名称被文件占用则报冲突。
- 删除只接受父目录句柄加子项名，须有 `allowDelete`。默认能删文件或空目录，非空目录须显式给出 `{ recursive: true }`。没有删除所选根或从单独选中的文件反查父目录的接口。`release()` 释放这个句柄，已打开的流和已取得的子句柄分别管理；所有授权随 Action 统一到期。

## 读写 XLSX 和 XLSQ 工作簿

从 9.1.0 起，Windows 上的 v4 Action 可对选择器返回的文件调用 `await file.openWorkbook()`。已有工作簿用 `fileMode: "open"`，导入申请 `"r"`，原地编辑申请 `"rw"`。导出新工作簿或覆盖用户确认的目标时用 `fileMode: "create"` 和 `fileAccess: "w"`。新工作簿有一张名为 `Sheet1` 的工作表，格式由文件名的 `.xlsx` 或 `.xlsq` 扩展名决定。保留已有内容的打开方式需要读权限。

在脚本编辑器中输入 `file.` 可补全 `openWorkbook`。写下 `const book = await file.openWorkbook()` 后，输入 `book.` 即可查看工作簿方法。JavaScript Action 模板和 TypeScript 都支持这条调用链，编辑器的 `quickexcel.d.ts` 定义文件可查看 `SelectedWorkbook` 的完整签名。

| 方法 | 结果 |
| --- | --- |
| `await book.getWorksheetNames()` | 返回按工作簿顺序排列的表名。 |
| `await book.getUsedRange(sheet)` | 返回包含已存储单元格（含格式）的最小 A1 矩形，空表返回 `null`。 |
| `await book.getValues(sheet, address)` | 返回二维数组。数值、文本和布尔值保持类型，空单元格返回 `null`，单元格错误返回其显示文本。内置公式可重算，不能取得当前结果的公式会报错。 |
| `await book.getFormulas(sheet, address)` | 返回同形状的二维数组，公式单元格返回以 `=` 开头的源文本，其余返回 `null`。 |
| `await book.setValues(sheet, address, values)` | 将矩阵写入同尺寸的区域。`null` 清空单元格内容，以 `=` 开头的文本保留为文本。 |
| `await book.setFormulas(sheet, address, formulas)` | 将等尺寸的公式字符串矩阵写入区域。 |
| `await book.addWorksheet(name)` | 在末尾新增工作表。 |
| `await book.saveAs(file, type)` | 将当前草稿保存到另一个可写文件，`type` 为 `"xlsx"` 或 `"xlsq"`。 |
| `await book.close()`、`await book.abort()` | 分别发布并关闭、放弃并关闭。只读工作簿关闭时释放资源。 |

例如，读取用户所选报表的单价和数量，把金额公式写回同一张表，并把本次金额带回当前工作簿：

```typescript
export async function main(workbook: QuickExcel.Workbook, host: QuickExcel.Host) {
  const [file] = await host.fs.getUserSelectedFiles({
    mode: "open", fileMode: "open", fileAccess: "rw",
    filters: [{ name: "工作簿", extensions: ["xlsx", "xlsq"] }]
  });
  if (!file) return;
  const book = await file.openWorkbook();
  try {
    const [sheet] = await book.getWorksheetNames();
    const [[price, quantity]] = await book.getValues(sheet, "A2:B2");
    if (typeof price !== "number" || typeof quantity !== "number") {
      throw new Error("A2 和 B2 需要填写数值。");
    }
    await book.setFormulas(sheet, "C2", [["=A2*B2"]]);
    const [[amount]] = await book.getValues(sheet, "C2");
    if (typeof amount !== "number") throw new Error("金额计算失败。");
    await book.close();
    workbook.getActiveCell().setValue(amount);
  } catch (error) {
    await book.abort();
    throw error;
  }
}
```

读取外部 XLSQ 时，其脚本保持禁用。保存沿用桌面的格式兼容性和工作表保护检查，保留支持的工作簿内容及 XLSQ 的脚本、虚拟文件和持久存储。宿主在发布前重新打开编码结果进行校验，再检查用户所选文件的身份和版本。错误写入、放弃及 Action 结束时仍打开的草稿均不会发布。成功关闭后的外部修改独立于当前工作簿的撤销历史。

区域操作使用单个工作表名称和 A1 矩形。每个 Action 对所选工作簿合计最多读 100,000 格、写 20,000 格。查找已用区域按扫描的已存储单元格计数。文件读写继续使用现有的包解析、草稿容量、IPC 和 Action 时限。

### 将工作簿另存为副本

从 9.2.0 起，Windows v4 Action 的 `workbook` 和 `file.openWorkbook()` 返回的工作簿都有 `saveAs(file: FileHandle, type: "xlsx" | "xlsq"): Promise<void>` 方法。目标须是本次 Action 获得的可写句柄。格式由 `type` 决定，建议在选择器中提供匹配的扩展名，方便重新打开。

```typescript
export async function main(workbook: QuickExcel.Workbook, host: QuickExcel.Host) {
  const [target] = await host.fs.getUserSelectedFiles({
    mode: "save", fileMode: "create", fileAccess: "w",
    suggestedName: "报表.xlsx", filters: [{ name: "Excel 工作簿", extensions: ["xlsx"] }]
  });
  if (target) await workbook.saveAs(target, "xlsx");
}
```

`xlsx` 保存工作表、公式、格式及其他受支持的 Excel 内容。依赖 QuickExcel 脚本的公式仍保留公式文本，需要运行这些脚本时请选择 `xlsq`。`xlsq` 还保存脚本模块、签名、`//xlsq/` 附件及 `localStorage`。副本取自调用 `saveAs` 时的草稿，即使稍后才 `await`，在两者之间对当前工作簿的修改也只留在原工作簿。`sessionStorage`、`//temp/` 和 `//mem/` 只属于当前会话。

成功返回时，副本已原子写入目标文件。之后的 Action 回滚、撤销或 `abort()` 都不会删除它。外部工作簿仍绑定原文件，例如对 `const book = await source.openWorkbook()` 完成编辑后，调用 `await book.saveAs(target, "xlsq")`，再调用 `await book.abort()` 即可只保留副本；`await book.close()` 则会将待提交的编辑写回原文件。只读源工作簿也可另存。目标冲突时，磁盘上的已有文件得以保留，可以重新选择目标后重试。

## 二进制和文本流

```typescript
declare namespace QuickExcel {
  interface FileStreamLifetime {
    readonly position: number;
    readonly length: number;
    close(): Promise<void>;
    abort(): Promise<void>;
  }

  interface BufferedFileStream extends FileStreamLifetime {
    read(buffer: Uint8Array, offset: number, len: number): number;
    write(buffer: Uint8Array, offset: number, len: number): number;
    readLine(encoding?: Encoding): string | null;
    readAllText(encoding?: Encoding): string;
    writeText(text: string, encoding?: Encoding): void;
    writeLine(text: string, encoding?: Encoding): void;
    seek(position: number): void;
  }

  interface StreamingFileStream extends FileStreamLifetime {
    read(buffer: Uint8Array, offset: number, len: number): Promise<number>;
    write(buffer: Uint8Array, offset: number, len: number): Promise<number>;
    readLine(encoding?: Encoding): Promise<string | null>;
    readAllText(encoding?: Encoding): Promise<string>;
    writeText(text: string, encoding?: Encoding): Promise<void>;
    writeLine(text: string, encoding?: Encoding): Promise<void>;
    seek(position: number): Promise<void>;
  }
}
```

默认 `await file.open()` 准备有界缓冲区，之后的同步读写只在 worker 内访问字节，因此 `readLine()` 可以直接返回字符串。只写模式下宿主仅向 worker 发送初始长度，worker 向宿主提交写入区段。每个同步方法检查所属轮次及流权限，轮次结束释放内部缓冲。检查大文件的文件头、跳到尾部或顺序处理较多数据时，用 `await file.open({ buffering: "stream" })`。此时宿主持有流，按需传输，脚本等待每次操作；两种流的游标和 EOF 语义相同。

| 操作 | 约定 |
| --- | --- |
| `read(buffer, offset, len)` | 从文件游标读，填入 **当前 Uint8Array 视图** 的指定范围，返回实际字节数。正长度请求读到 EOF 返回 0；零长度也返回 0。支持短读，只改动实际读到的部分。 |
| `write(buffer, offset, len)` | 从视图指定区段取值，成功返回完整 `len`，允许覆盖或在末尾追加。异步写在调用时复制传入区段。零长度不改变文件。 |
| `seek(position)` | 绝对字节位置，允许 0 到当前长度，含末尾。偏移、长度与游标均要求非负安全整数，缓冲范围在任何拷贝前校验。只在操作成功后推进游标。 |
| `readLine(encoding)` | 从当前字节开始，识别 CR、LF、CRLF，移除行尾。空行返回 `""`，最后一行无需换行。仅当已无内容时返回 `null`。 |
| `readAllText(encoding)` | 解码从当前游标到 EOF 的内容，空余量返回 `""`。超过单次文本限额时报错，游标停在原处。 |
| `writeText` / `writeLine` | 默认严格 UTF-8，亦支持 UTF-16LE/BE。`writeLine` 追加固定 LF，需要 CRLF 可显式 `writeText(text + "\r\n")`。写入不自动加 BOM。 |

文本读取仅在文件字节位置 0 消费匹配编码的 BOM。冲突 BOM、残缺或非法编码返回 `InvalidEncoding`；写文本时也拒绝孤立代理项。文本和二进制共用一个字节游标，预读不能使下次读跳过字节。跨块字符和 CRLF 由流保留解码状态，seek 或更换编码从逻辑位置重建解码状态。相同流最多一项未完成操作，重叠调用报 `Locked`。

## 写入、删除与撤销

一个写流拥有一个私有草稿，读取权限为 `rw` 时可读到自己的写入。一个目标同时最多一个写流。显式 `close()` 是文件发布点：宿主依据所选目标或上次成功发布时记录的身份和版本验证授权，满足条件后才以同目录原子操作发布，并更新该句柄的身份、版本和存在状态。同一句柄再次 `open()` 以已发布内容为基础。目标遭到外部替换、编辑或抢先创建时返回 `Conflict`，保留对方的版本。后端只有在能保证该操作语义时才提供对应能力。保留已有内容且未改动时，关闭只释放资源；`create` 或 `open_or_create` 为新目标准备的空草稿在显式关闭时也创建空文件。

写入异常后流进入失败状态，后续关闭不可发布半份草稿。`abort()` 可重复调用，清理尚未发布的数据。Action 正常结束也丢弃未关闭的写流；Esc、超时、撤权、窗口关闭和 worker 退出均触发宿主清理。流程以发布动作作为取消边界，已发布的文件保留；回复丢失时需核对目标文件，宿主不会自动重试。

创建子目录、删除、成功关闭写流都是对磁盘的外部操作。即使 Action 随后抛错或工作簿事务提交失败，已完成的磁盘操作仍然生效。撤销、重做只影响工作簿。多文件导出按文件发布；递归删除中途因占用或取消而停止时，已完成的删除保留；发生非取消错误时返回已删除数量。导入脚本可先完成读取校验再修改工作簿，导出脚本可把 `close()` 放在校验和内容生成之后。

## 示例

把 UTF-8 文本写入当前工作表的 A 列。点取消保持原表，空行占一行，文件末尾不多写一个单元格。

```typescript
export async function main(workbook: QuickExcel.Workbook, host: QuickExcel.Host) {
  const [file] = await host.fs.getUserSelectedFiles({
    mode: "open", fileMode: "open", fileAccess: "r",
    title: "导入文本到 A 列", filters: [{ name: "文本文件", extensions: ["txt"] }]
  });
  if (!file) return;

  const stream = await file.open();
  const sheet = workbook.getActiveWorksheet();
  try {
    let line: string | null;
    let row = 1;
    while ((line = stream.readLine("utf-8")) !== null) {
      sheet.getRange("A" + row++).setValue(line);
    }
  } finally {
    await stream.close();
  }
}
```

在用户选定目录的 `reports` 子目录发布文件。只申请写权限；保存编码、换行及是否覆盖在调用处可见。

```typescript
export async function main(workbook: QuickExcel.Workbook, host: QuickExcel.Host) {
  const root = await host.fs.getUserSelectedDirectory({
    title: "选择报表输出目录", fileAccess: "w"
  });
  if (!root) return;

  const reports = await root.getDirectory("reports", { create: true });
  const file = await reports.getFile("summary.txt", { fileMode: "create" });
  const stream = await file.open();
  try {
    stream.writeLine(workbook.getActiveWorksheet().getName(), "utf-8");
    await stream.close();
  } catch (error) {
    await stream.abort();
    throw error;
  }
}
```

## 宿主接入与验收

1. **选择与身份。** Core 定义脚本专用选择端口，Avalonia 使用所属窗口的 `OpenFilePickerAsync`、`SaveFilePickerAsync`、`OpenFolderPickerAsync`。应用侧在返回对应 Action 前复核会话、脚本代次、操作身份与取消状态。宿主表保存根对象、身份及权限；IPC 只传不可猜测的引用和允许公开的元数据。句柄不可序列化或从 JSON 恢复。运行时独立校验每个请求，只允许已授权、用户发起的 v4 Action；初始化、UDF、Events、控制台和模块顶层均不能弹窗或复用句柄。异常、日志和协议不得泄露实际路径或 OS 句柄。[桌面脚本选择器](../../QuickExcel.Desktop.Avalonia/Views/MainWindow.ScriptFiles.cs) · [Avalonia 选择器](https://docs.avaloniaui.net/docs/services/storage/storage-provider)
2. **目录边界。** 宿主以用户选中的目录对象为根，按直接子项名逐级定位并验证身份。拒绝分隔符、点段、设备名和备用数据流等特殊形式。拒绝穿越符号链接、重解析点、挂载边界及特殊文件；枚举跳过这些项，显式请求返回 `UnsupportedApi`。删除同样逐项校验。句柄不能因同名替换或祖先移出授权树而落到别的对象。读写与发布保留身份绑定，避免“检查路径再按路径重开”；写入多硬链接对象若无法隔离目录项则拒绝。保护活动工作簿及其租约的写删，覆盖由别名或目录访问的情形。单独的字符串前缀校验不足以证明归属；平台或存储提供者不能满足相应保证时返回 `UnsupportedApi`。[Windows 目录句柄相对打开](https://learn.microsoft.com/en-us/windows/win32/api/winternl/nf-winternl-ntcreatefile) · [Linux 路径解析](https://docs.kernel.org/filesystems/path-lookup.html)
3. **调度与时限。** 扩展现有 fetch 专用 Promise 调度器以区分文件与网络请求，共享有序的 ModelCall 管道。worker 在宿主选择和异步 I/O 期间让出 JSC；每次 JS 执行或续接仍有 1 秒原生看门狗。Windows 仅在宿主实际显示、等待用户操作的选择界面期间冻结 Action 的剩余时限，同步管理执行器与 runtime 的两个截止时间；回到宿主后恢复剩余时间，不重置整轮预算。文件 I/O 和网络照常计时。总是允许取消和撤权，迟到的选择结果只做释放。Linux 的 5 秒 OS 生命周期照常生效，超时结束原 Action，不重放脚本。[执行器](../../QuickExcel.Infrastructure/Scripting/ScriptWorkerActionExecutor.cs) · [轮次期限与 IPC](../../QuickExcel.Infrastructure/Scripting/ScriptWorkerRuntimeFactory.cs) · [JSC 调度器](../../QuickExcel.Scripting/JscSession.cs)
4. **资源与错误。** 同步缓冲区每流 4 MiB、每 Action 合计 8 MiB，流式二进制每块至多 64 KiB，单行 1 MiB、整段文本 4 MiB（按编码字节）。宿主暂存草稿总量上限为 16 MiB，含保留旧内容的临时副本；流式读取可定位到大文件任意可表示的字节位置。枚举分页并限制同时打开的句柄、流和待处理选择器。块用 base64url；宿主与 worker 按实际编码后的请求和响应共同执行现有每帧及每轮累计 16 MiB、每轮 4096 次模型调用预算。与工作簿和 fetch 共享，分块不增加总量，fetch 仍单独最多 16 次。权限、旧句柄、冲突、解码、预算分别用稳定错误码 `PermissionDenied`、`InvalidReference`、`Conflict`、`InvalidEncoding`、`BudgetExceeded`；复用现有 `InvalidArgument`、`ItemNotFound`、`Locked` 等码。宿主对用户可见的错误只附操作和安全显示名。
5. **生成与兼容。** 在 ApiGen 的手写宿主扩展源维护类型与新增错误码，重新生成产品声明并检验 Monaco、隔离 TypeScript 编译和 Native AOT worker 一致性。IPC 14 在 Action 请求中传递宿主授予的文件能力，经已有的 JSC 异步续接入口处理文件响应。IPC 13 的旧端在握手时被拒绝。句柄仅驻留内存，无需增加 XLSQ 部件或迁移 v3 源码。[声明源](../../tools/QuickExcel.ApiGen/src/quickexcel-extensions.d.ts) · [Worker 协议](../../QuickExcel.ScriptContracts/WorkerProtocol.cs)

| 验收场景 | 预期行为 |
| --- | --- |
| 选择器与权限 | 多选文件/目录、单选、取消、覆盖确认、无窗口和非法参数均有明确结果；宿主拒绝伪造、跨 Action、跨文档及越权 IPC，只读流不能发布写入，只写流在任一模式都无法取得已有字节。 |
| 文件内容 | 两类流覆盖空文件、短读、零长度、Uint8Array 子视图、随机访问、跨块中文/CRLF、空行、无末尾换行、BOM、非法编码和超限；同步与流式结果一致。 |
| 目录与竞争 | 多级目录只选择一次；测试相同前缀的邻居目录、链接、硬链接、祖先移动、同名替换、缺失目标被抢建、活动工作簿别名、递归删除遇到占用。 |
| 落盘与失败 | 选择或放弃保存、超时、撤权、空间不足均保留此前目标；发布后 Action 抛错以及撤销/重做保留已发布文件。时间点跨越发布时按实际结果记录。 |
| 隔离与集成 | 真正的隔离 worker 验证角色、帧/累计预算、异步续接、看门狗、过期选择和清理。桌面验证文件/目录选择器及权限说明；各平台独立记录适配器证据。 |

## Windows 验证

13 项文件运行时测试使用实际的隔离 Native AOT worker。测试覆盖了带点和不带点的扩展名过滤、文件和目录多选、初始化阶段的权限限制、取消与迟到结果、句柄失效、跨块文本解码、写入冲突、保护活动工作簿，以及目录联接和硬链接的处理。[运行时测试](../../QuickExcel.Core.Test/ScriptUserFileTests.cs)

桌面检查从 v4 Action 唤起原生文件夹选择器，通过 Windows UI Automation 识别 Shell 的地址栏和搜索框，再关闭对话框并检验 Action 收到取消结果。检查脚本使用诊断版桌面程序。[目录选择器检查](../../scripts/test-script-folder-picker.ps1)
