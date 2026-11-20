import Fuse from 'fuse.js';
import { readRows, writeResult } from 'common';

export function main(workbook: QuickExcel.Workbook): void {
  const products = readRows(workbook, 'A2:C6').map(([name, category, price]) => ({
    name: String(name), category: String(category), price: Number(price)
  }));
  // 修改 E2 中的搜索词，例如 iphne、ipad、galaxy。
  const query = String(readRows(workbook, 'E2')[0][0]);
  const search = new Fuse(products, { keys: ['name', 'category'], threshold: 0.35, ignoreLocation: true });
  const matches = search.search(query);
  writeResult(workbook, [['商品', '类别', '价格'],
    ...matches.map(({ item }) => [item.name, item.category, item.price])]);
}
