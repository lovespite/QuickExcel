import { z } from 'zod';
import { readRows, writeResult } from 'common';

export function main(workbook: QuickExcel.Workbook): void {
  const person = z.object({
    name: z.string().min(1, '姓名不能为空'),
    email: z.email('邮箱格式无效'),
    age: z.coerce.number().int('年龄需要整数').min(1, '年龄至少为 1')
  });
  const checked = readRows(workbook, 'A2:C4').map(([name, email, age], index) => {
    const result = person.safeParse({ name: String(name), email: String(email), age });
    return [index + 2, String(name), result.success ? '通过' : '未通过', result.success
      ? `${result.data.age} 岁 · ${result.data.email}`
      : result.error.issues.map(issue => `${issue.path.join('.')}: ${issue.message}`).join('；')];
  });
  writeResult(workbook, [['原始行号', '姓名', '校验结果', '详情'], ...checked]);
}
