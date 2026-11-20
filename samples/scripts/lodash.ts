import _ from 'lodash';
import { readRows, writeResult } from 'common';

// 修改“输入数据”中的订单，再应用脚本并运行 example。
export function main(workbook: QuickExcel.Workbook): void {
  const orders = readRows(workbook, 'A2:C6').map(([id, category, amount]) => ({
    id: String(id), category: String(category), amount: Number(amount)
  })).filter(order => order.id !== '');
  const unique = _.uniqBy(orders, 'id');
  const grouped = _.groupBy(unique, 'category');
  const totals = _.map(grouped, (items, category) => ({
    category, count: items.length, amount: _.sumBy(items, 'amount')
  }));
  const sorted = _.orderBy(totals, ['amount', 'category'], ['desc', 'asc']);
  writeResult(workbook, [['类别', '去重订单数', '金额合计'],
    ...sorted.map(row => [row.category, row.count, row.amount])]);
}
