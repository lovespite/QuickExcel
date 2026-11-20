import Papa from 'papaparse';
import { readRows, writeResult } from 'common';

export function main(workbook: QuickExcel.Workbook): void {
  const csv = String(readRows(workbook, 'A2')[0][0]);
  const parsed = Papa.parse<{ 商品: string; 数量: number; 单价: number }>(csv, {
    header: true, dynamicTyping: true, skipEmptyLines: true
  });
  if (parsed.errors.length) throw new Error(parsed.errors.map(error => error.message).join('; '));
  writeResult(workbook, [['商品', '数量', '单价', '小计'],
    ...parsed.data.map(row => [row.商品, row.数量, row.单价, row.数量 * row.单价]),
    ['', '', '', ''], ['导出 CSV', Papa.unparse(parsed.data), '', '']]);
}
