import dayjs from 'dayjs';
import { readRows, writeResult } from 'common';

export function main(workbook: QuickExcel.Workbook): void {
  // UTC、Timezone、CustomParseFormat 等插件已经内置，无需再次 extend。
  const result = readRows(workbook, 'A2:B4').map(([date, zone]) => {
    const text = String(date), timezone = String(zone);
    if (!dayjs(text, 'YYYY-MM-DD HH:mm', true).isValid())
      return [text, timezone, '无效日期', '—'];
    const local = dayjs.tz(text, timezone);
    return [text, timezone, local.utc().format('YYYY-MM-DD HH:mm[ UTC]'),
      local.add(7, 'day').format('YYYY-MM-DD')];
  });
  writeResult(workbook, [['输入时间', '时区', 'UTC 时间', '七天后日期'], ...result]);
}
