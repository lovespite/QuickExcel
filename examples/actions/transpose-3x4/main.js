// 固定区域、按值转置：首张工作表的 A1:D3（3 行 4 列）→ A1:C4（4 行 3 列）。
// 请为结果预留 A4:C4；只读取当前计算值，格式保持目标单元格原样。
function main(workbook, host) {
  const sheet = workbook.getActiveWorksheet();
  const source = sheet.getRange('A1:D3').getValues();
  const transposed = Array.from({ length: 4 }, (_, column) =>
    source.map(row => row[column]));

  sheet.getRange('A1:C4').setValues(transposed);
  // 清空结果区域外的旧内容，避免同一事务重复写入同一单元格。
  sheet.getRange('D1:D3').clear();
  console.log('已将 3 行 4 列转置为 4 行 3 列。');
}
