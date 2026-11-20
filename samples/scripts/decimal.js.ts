import Decimal from 'decimal.js';
import { readRows, writeResult } from 'common';

export function main(workbook: QuickExcel.Workbook): void {
  // 独立配置不影响同一上下文里的其他脚本。以字符串保留输入精度。
  const Money = Decimal.clone({ precision: 40, rounding: Decimal.ROUND_HALF_UP });
  const rows = readRows(workbook, 'A2:C4');
  let total = new Money(0);
  const detail = rows.map(([name, price, quantity]) => {
    const subtotal = new Money(String(price)).times(Number(quantity)).toDecimalPlaces(2);
    total = total.plus(subtotal);
    return [String(name), String(price), Number(quantity), subtotal.toFixed(2)];
  });
  writeResult(workbook, [['商品', '精确单价', '数量', '小计（字符串）'],
    ...detail, ['合计', '', '', total.toFixed(2)]]);
}
