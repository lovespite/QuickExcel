// 公共模块：只负责读写工作表，不加载网络数据。
export type Cell = string | number | boolean;

export function readRows(workbook: QuickExcel.Workbook, address: string): Cell[][] {
  return workbook.getWorksheet('输入数据')!.getRange(address).getValues();
}

export function writeResult(workbook: QuickExcel.Workbook, rows: Cell[][]): void {
  const sheet = workbook.getWorksheet('运行结果')!;
  const width = Math.max(...rows.map(row => row.length));
  // 先清空旧内容，重复运行时不会留下上次多出来的行。
  sheet.getUsedRange(true).clear(QuickExcel.ClearApplyTo.contents);
  sheet.getRangeByIndexes(0, 0, rows.length, width)
    .setValues(rows.map(row => [...row, ...Array<Cell>(width - row.length).fill('')]));
  sheet.getRangeByIndexes(0, 0, 1, width).getFormat().getFont().setBold(true);
}
