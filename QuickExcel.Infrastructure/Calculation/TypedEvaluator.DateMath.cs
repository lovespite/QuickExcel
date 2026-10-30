using System.Globalization;
using System.Text.RegularExpressions;
using QuickExcel.Domain;
using QuickExcel.Formulas;

namespace QuickExcel.Calculation;

internal sealed partial class TypedEvaluator
{
    private int LastDate => snapshot.DateSystem == ExcelDateSystem.Windows1900 ? 2958465 : 2957003;
    private CellValue DateMath(FunctionNode function)
    {
        var name = function.Name;
        CellValue Arg(int i)
        {
            var value = LookupScalar(function.Arguments[i]);
            if (value.Kind == CellValueKind.Error) Fail(value.ErrorCode, "日期或数学参数包含错误。");
            return value;
        }
        double N(int i) => Number(Arg(i));
        double D(int i) => DateNumber(Arg(i));
        var epochOffset = snapshot.DateSystem == ExcelDateSystem.Mac1904 ? 1462 : 0;
        double AnalysisDate(CellValue value, bool reference)
        {
            if (value.Kind == CellValueKind.Boolean) throw new FormulaValueException(CellErrorCode.Value, "分析日期参数不接受布尔值。");
            var serial = ValidDate(DateNumber(value));
            // Excel's legacy analysis functions distinguish direct text from referenced text in 1904 books.
            return serial + (reference || value.Kind != CellValueKind.Text ? epochOffset : 0);
        }
        double AnalysisArg(int i) => AnalysisDate(Arg(i), function.Arguments[i] is ReferenceNode);
        double AnalysisNumber(int i)
        {
            var value = Arg(i);
            if (value.Kind == CellValueKind.Boolean) throw new FormulaValueException(CellErrorCode.Value, "分析函数数值参数不接受布尔值。");
            return Number(value);
        }
        switch (name)
        {
            case "DATE":
            {
                var year = Math.Floor(N(0)); var month = Math.Floor(N(1)); var day = Math.Floor(N(2));
                if (year < 0 || year > 9999 || month < -32767 || month > 32766) return Error(CellErrorCode.Number);
                if (year < 1900) year += 1900;
                // Excel's DATE day argument uses this legacy signed-short sentinel on overflow.
                if (day < -32768 || day > 32767) day = 32767;
                var months = year * 12 + month - 1; var y = (int)Math.Floor(months / 12); var m = (int)(months - y * 12) + 1;
                if (y < 1 || y > 9999) return Error(CellErrorCode.Number);
                var result = ExcelSerialDate.FromDateTime(new DateTime(y, m, 1), snapshot.DateSystem).Serial + day - 1;
                return DateResult(result);
            }
            case "TIME":
            {
                var h = Math.Floor(N(0)); var m = Math.Floor(N(1)); var s = Math.Floor(N(2));
                if (h < -32768 || h > 32767 || m < -32768 || m > 32767 || s < -32768 || s > 32767) return Error(CellErrorCode.Number);
                var total = h * 3600 + m * 60 + s;
                return total < 0 ? Error(CellErrorCode.Number) : CellValue.FromNumber(total / 86400 % 1);
            }
            case "YEAR": case "MONTH": case "DAY":
            {
                var serial = ValidDate(D(0));
                var parts = DateParts(Math.Floor(serial * 86400 + .5) / 86400); return CellValue.FromNumber(name == "YEAR" ? parts.Year : name == "MONTH" ? parts.Month : parts.Day);
            }
            case "HOUR": case "MINUTE": case "SECOND":
            {
                var serial = ValidDate(D(0)); var rounded = Math.Floor(serial * 86400 + .5);
                if (rounded >= (LastDate + 1d) * 86400) return Error(CellErrorCode.Number);
                var seconds = (int)(rounded % 86400);
                return CellValue.FromNumber(name == "HOUR" ? seconds / 3600 : name == "MINUTE" ? seconds / 60 % 60 : seconds % 60);
            }
            case "EDATE": case "EOMONTH":
            {
                var parts = DateParts(AnalysisArg(0), windows: true); var offset = Math.Truncate(AnalysisNumber(1)); var months = parts.Year * 12d + parts.Month - 1 + offset;
                if (months < 1900 * 12 || months >= 120000) return Error(CellErrorCode.Number);
                var year = (int)Math.Floor(months / 12); var month = (int)(months - year * 12) + 1;
                var day = name == "EOMONTH" ? DateTime.DaysInMonth(year, month) : Math.Min(parts.Day, DateTime.DaysInMonth(year, month));
                var date = new DateTime(year, month, 1).AddDays(day - 1);
                return CellValue.FromNumber(ExcelSerialDate.FromDateTime(date, ExcelDateSystem.Windows1900).Serial - epochOffset);
            }
            case "WEEKDAY":
            {
                var date = ValidDate(D(0)); var rounded = Math.Floor(date * 86400 + .5) / 86400;
                var serial = (int)Math.Floor(ValidDate(rounded)); var mode = function.Arguments.Count == 1 ? 1 : Math.Floor(N(1)); var weekday = (serial + (epochOffset == 0 ? 6 : 5)) % 7;
                return mode switch { 1 => CellValue.FromNumber(weekday + 1), 2 or 11 => CellValue.FromNumber((weekday + 6) % 7 + 1),
                    3 => CellValue.FromNumber((weekday + 6) % 7), >= 12 and <= 17 => CellValue.FromNumber((weekday - ((int)mode - 10) + 7) % 7 + 1), _ => Error(CellErrorCode.Number) };
            }
            case "WORKDAY": case "NETWORKDAYS":
            {
                var start = (int)Math.Floor(AnalysisArg(0)); var second = name == "WORKDAY" ? Math.Floor(AnalysisNumber(1)) : Math.Floor(AnalysisArg(1));
                if (name == "WORKDAY" && start > LastDate) return Error(CellErrorCode.Number);
                if (name == "WORKDAY" && second == 0) return CellValue.FromNumber(start - epochOffset); // Excel does not inspect holidays for zero days.
                if (name == "WORKDAY" && Math.Abs(second) > LastDate) return Error(CellErrorCode.Number);
                var holidays = new HashSet<int>();
                if (function.Arguments.Count == 3)
                {
                    var holidayArgument = Evaluate(function.Arguments[2]);
                    foreach (var value in Values(holidayArgument))
                    {
                        check(); if (value.Kind == CellValueKind.Blank) continue;
                        if (value.Kind == CellValueKind.Error) Fail(value.ErrorCode, "节假日包含错误。");
                        if (value.Kind == CellValueKind.Boolean) return Error(CellErrorCode.Value);
                        var reference = name == "WORKDAY" ? holidayArgument is not Range r || ConditionalRanges.Shape(r.Reference) == (1, 1) : function.Arguments[2] is ReferenceNode;
                        holidays.Add((int)Math.Floor(AnalysisDate(value, reference)));
                    }
                }
                int Count(int first, int last)
                {
                    if (first > last) return 0;
                    var days = last - first + 1; var total = days / 7 * 5;
                    for (var i = 0; i < days % 7; i++) if (Working(first + i)) total++;
                    foreach (var holiday in holidays) { check(); if (holiday >= first && holiday <= last && Working(holiday)) total--; }
                    return total;
                }
                if (name == "NETWORKDAYS") { var end = (int)second; return CellValue.FromNumber(end < start ? -Count(end, start) : Count(start, end)); }
                var forward = second > 0; var target = (int)Math.Abs(second);
                if ((forward ? Count(start + 1, LastDate) : Count(0, start - 1)) < target) return Error(CellErrorCode.Number);
                // Binary search workday counts avoids iterating millions of dates or requested days.
                var low = forward ? start + 1 : 0; var high = forward ? LastDate : start - 1;
                while (low < high)
                {
                    check(); var middle = forward ? low + (high - low) / 2 : low + (high - low + 1) / 2;
                    var count = forward ? Count(start + 1, middle) : Count(middle, start - 1);
                    if (forward) { if (count >= target) high = middle; else low = middle + 1; }
                    else { if (count >= target) low = middle; else high = middle - 1; }
                }
                return low < epochOffset ? Error(CellErrorCode.Number) : CellValue.FromNumber(low - epochOffset);
            }
            case "INT": return CellValue.FromNumber(Math.Floor(Significant15(N(0))));
            case "MOD":
            {
                var number = N(0); var divisor = N(1); if (divisor == 0) return Error(CellErrorCode.DivisionByZero);
                if (Math.Abs(number / divisor) >= 1125900000000) return Error(CellErrorCode.Number);
                var remainder = number % divisor;
                if (remainder != 0 && Math.Sign(remainder) != Math.Sign(divisor)) remainder += divisor;
                return CellValue.FromNumber(remainder);
            }
            case "ROUNDUP": case "ROUNDDOWN":
            {
                var number = Significant15(N(0)); var digits = Math.Truncate(N(1)); if (number == 0) return CellValue.FromNumber(0);
                var factor = Math.Pow(10, Math.Abs(digits));
                if (digits >= 0 && !double.IsFinite(Math.Abs(number) * factor)) return CellValue.FromNumber(number);
                var scaled = Significant15(digits >= 0 ? Math.Abs(number) * factor : Math.Abs(number) / factor);
                var rounded = name == "ROUNDUP" ? Math.Ceiling(scaled) : Math.Floor(scaled);
                if (digits < 0 && !double.IsFinite(factor)) return name == "ROUNDUP" ? Error(CellErrorCode.Number) : CellValue.FromNumber(0);
                return CellValue.FromNumber(Math.Sign(number) * (digits >= 0 ? rounded / factor : rounded * factor));
            }
        }
        return Error(CellErrorCode.Name);
    }

    private double ValidDate(double serial)
    {
        if (serial < 0 || serial >= LastDate + 1d) Fail(CellErrorCode.Number, "日期超出工作簿日期系统的有效范围。");
        return serial;
    }
    private static double Significant15(double value) => double.Parse(value.ToString("G15", CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);
    private CellValue DateResult(double serial) => serial < 0 || serial > LastDate ? Error(CellErrorCode.Number) : CellValue.FromNumber(serial);
    private (int Year, int Month, int Day) DateParts(double serial, bool windows = false)
    {
        var system = windows ? ExcelDateSystem.Windows1900 : snapshot.DateSystem;
        if (serial < 0 || serial >= (windows ? 2958466d : LastDate + 1d)) Fail(CellErrorCode.Number, "日期无效。");
        var day = Math.Floor(serial);
        if (system == ExcelDateSystem.Windows1900)
        {
            if (day == 0) return (1900, 1, 0);
            if (day == 60) return (1900, 2, 29);
        }
        if (!new ExcelSerialDate(day, system).TryToDateTime(out var date)) Fail(CellErrorCode.Number, "日期无效。");
        return (date.Year, date.Month, date.Day);
    }
    private static bool Working(int serial) => (serial + 6) % 7 is >= 1 and <= 5;

    private double DateNumber(CellValue value)
    {
        if (value.Kind != CellValueKind.Text || double.TryParse(value.TextValue, NumberStyles.Float, CultureInfo.InvariantCulture, out _)) return Number(value);
        var text = value.TextValue.Trim();
        var match = Regex.Match(text, @"^(?:(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:\s+|$))?(?:(\d{1,2}):(\d{1,2})(?::(\d{1,2})(?:\.(\d{1,7}))?)?)?$", RegexOptions.CultureInvariant, TimeSpan.FromMilliseconds(100));
        if (!match.Success || !match.Groups[1].Success && !match.Groups[4].Success) throw new FormulaValueException(CellErrorCode.Value, "日期文本须为年在前的公历日期或 24 小时时间。");
        int Part(int i) => match.Groups[i].Success ? int.Parse(match.Groups[i].Value, CultureInfo.InvariantCulture) : 0;
        var hour = Part(4); var minute = Part(5); var second = Part(6);
        if (hour > 23 || minute > 59 || second > 59) throw new FormulaValueException(CellErrorCode.Value, "时间文本无效。");
        var fraction = match.Groups[7].Success ? double.Parse("0." + match.Groups[7].Value, CultureInfo.InvariantCulture) : 0;
        double serial = 0;
        if (match.Groups[1].Success)
        {
            var year = Part(1); var month = Part(2); var day = Part(3);
            if (year == 1900 && month == 2 && day == 29 && snapshot.DateSystem == ExcelDateSystem.Windows1900) serial = 60;
            else
            {
                if (year < 1 || month < 1 || month > 12 || day < 1 || day > DateTime.DaysInMonth(year, month)) throw new FormulaValueException(CellErrorCode.Value, "日期文本无效。");
                serial = ExcelSerialDate.FromDateTime(new DateTime(year, month, day), snapshot.DateSystem).Serial;
            }
        }
        return serial + (hour * 3600 + minute * 60 + second + fraction) / 86400;
    }
}
