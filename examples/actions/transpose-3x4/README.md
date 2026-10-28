# 3×4 表的行列转置 action

[`main.js`](main.js) 注册“转置 3×4 表”，将**首张工作表**的 `A1:D3`（3 行 4 列）按值转置到 `A1:C4`（4 行 3 列），清空旧区域剩余的 `D1:D3`。所有写入作为一次操作提交，可一次撤销。

运行前：

```text
1   2   3   4
5   6   7   8
9  10  11  12
```

运行后：

```text
1   5   9
2   6  10
3   7  11
4   8  12
```

## 生成与试用

从仓库根目录运行（Windows x64、.NET 10 SDK，原生构建要求与主项目相同）：

```powershell
dotnet run --project .\examples\actions\transpose-3x4\Transpose3x4.csproj -c Release --artifacts-path .\artifacts\build
```

生成器通过 Core/Infrastructure 保存绑定正确的示例包，并用真实 worker 验证转置、清空旧内容、一次撤销和重做。输出 `PASS` 后，用支持 actions 的 QuickExcel 打开 `artifacts/examples/transpose-3x4.xlsq`，点击“应用脚本并重算”，选择“转置 3×4 表”，再点击“运行操作”。交付文件保留运行前的数据，打开不会自动执行脚本。

这是固定区域、按值转置的示例。`A4:C4` 会被结果覆盖；示例文件已将这一区域留空。公式读取当前缓存值，不搬移公式或调整引用，样式保持目标单元格原样。每次都读取 `A1:D3`，重复演示前请撤销或重新打开原文件。

桌面暂不提供创建 actions 绑定的界面，因此新建普通工作簿后只粘贴 JS 不足以启用动作。生成器使用文件级 `scriptApiVersion: 2` 和 `main / workbook / actions` 绑定；详细接口见 [actions 契约](../../../QuickExcel.Desktop/docs/13-m4-actions.md)。修改 `main.js` 后重新运行生成命令即可更新示例包（会覆盖同名生成文件）。
