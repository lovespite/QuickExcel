// 固定区域、按值转置：首张工作表的 A1:D3（3 行 4 列）→ A1:C4（4 行 3 列）。
// 请为结果预留 A4:C4；公式只读取当前缓存值，格式保持目标单元格原样。
quickExcel.actions.register({
  name: '转置 3×4 表',
  description: '将 A1:D3 按值转置到 A1:C4，并清空 D1:D3；可一次撤销。'
}, ({ snapshot, patch }) => {
  const id = snapshot.sheets[0].id;
  // 先读取完整快照，源区域与目标区域重叠也不会影响后续读取。
  const source = snapshot.getRange(id, 'A1:D3');
  const transposed = Array.from({ length: 4 }, (_, column) =>
    source.map(row => row[column]));

  patch.setValues(id, 'A1:C4', transposed);
  // 只清空结果区域外的旧内容，避免 patch 对同一单元格重复写入。
  patch.setValues(id, 'D1:D3', [[null], [null], [null]]);
  console.log('已将 3 行 4 列转置为 4 行 3 列。');
});
