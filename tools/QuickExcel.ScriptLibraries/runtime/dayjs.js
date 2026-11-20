import dayjs from 'dayjs';
import utc from 'dayjs/plugin/utc.js';
import timezone from 'dayjs/plugin/timezone.js';
import customParseFormat from 'dayjs/plugin/customParseFormat.js';
import duration from 'dayjs/plugin/duration.js';
import isBetween from 'dayjs/plugin/isBetween.js';
import isSameOrBefore from 'dayjs/plugin/isSameOrBefore.js';
import isSameOrAfter from 'dayjs/plugin/isSameOrAfter.js';
import isoWeek from 'dayjs/plugin/isoWeek.js';
import weekOfYear from 'dayjs/plugin/weekOfYear.js';
import quarterOfYear from 'dayjs/plugin/quarterOfYear.js';
import advancedFormat from 'dayjs/plugin/advancedFormat.js';
import localizedFormat from 'dayjs/plugin/localizedFormat.js';
import localeData from 'dayjs/plugin/localeData.js';
import minMax from 'dayjs/plugin/minMax.js';
import relativeTime from 'dayjs/plugin/relativeTime.js';
import 'dayjs/locale/zh-cn.js';
for (const plugin of [utc, timezone, customParseFormat, duration, isBetween, isSameOrBefore,
  isSameOrAfter, isoWeek, weekOfYear, quarterOfYear, advancedFormat, localizedFormat,
  localeData, minMax, relativeTime]) dayjs.extend(plugin);
export default dayjs;
