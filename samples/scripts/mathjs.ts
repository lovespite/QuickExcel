import * as math from 'mathjs';
import { readRows, writeResult } from 'common';

export function main(workbook: QuickExcel.Workbook): void {
  const parameters = readRows(workbook, 'B2:B4').map(row => String(row[0]));
  const a = readRows(workbook, 'D2:E3').map(row => row.map(Number));
  const b = readRows(workbook, 'G2:H3').map(row => row.map(Number));
  const product = math.multiply(a, b) as number[][];
  writeResult(workbook, [['运算', '结果'],
    ['表达式', Number(math.evaluate(parameters[0]))],
    ['英寸转厘米', math.unit(parameters[1]).toNumber('cm')],
    ['对 x 求导', math.derivative(parameters[2], 'x').toString()],
    ['矩阵乘积 (1,1)', product[0][0]], ['矩阵乘积 (1,2)', product[0][1]],
    ['矩阵乘积 (2,1)', product[1][0]], ['矩阵乘积 (2,2)', product[1][1]],
    ['大数 0.1 + 0.2', math.add(math.bignumber('0.1'), math.bignumber('0.2')).toString()]]);
}
