import { mean, median, standardDeviation, linearRegression, linearRegressionLine, sampleCorrelation } from 'simple-statistics';
import { readRows, writeResult } from 'common';

export function main(workbook: QuickExcel.Workbook): void {
  const points = readRows(workbook, 'A2:B6').map(([x, y]) => [Number(x), Number(y)]);
  const xs = points.map(point => point[0]), ys = points.map(point => point[1]);
  const regression = linearRegression(points);
  writeResult(workbook, [['指标', '数值'], ['样本数', ys.length], ['Y 均值', mean(ys)],
    ['Y 中位数', median(ys)], ['Y 总体标准差', Number(standardDeviation(ys).toFixed(6))],
    ['回归斜率', regression.m], ['回归截距', regression.b],
    ['相关系数', sampleCorrelation(xs, ys)], ['X=6 的预测值', linearRegressionLine(regression)(6)]]);
}
