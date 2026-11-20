import * as aq from 'arquero';
import { readRows, writeResult } from 'common';

export function main(workbook: QuickExcel.Workbook): void {
  const sales = readRows(workbook, 'A2:C6').map(([category, region, amount]) => ({
    category: String(category), region: String(region), amount: Number(amount)
  }));
  const summary = aq.from(sales).groupby('category')
    .rollup({ orders: aq.op.count(), total: aq.op.sum('amount') })
    .orderby(aq.desc('total')).objects() as { category: string; orders: number; total: number }[];
  writeResult(workbook, [['类别', '记录数', '销售合计'],
    ...summary.map(row => [String(row.category), Number(row.orders), Number(row.total)])]);
}
