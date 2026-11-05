using QuickExcel.Domain;

namespace QuickExcel.Core;

// Values intentionally match Excel's XlTimePeriods, independently verified in native files.
public enum ConditionalDatePeriod { Today, Yesterday, Last7Days, ThisWeek, LastWeek, LastMonth, Tomorrow, NextWeek, NextMonth, ThisMonth }

public static class ConditionalDates
{
    public static string XmlName(ConditionalDatePeriod period) => period switch
    {
        ConditionalDatePeriod.Today => "today", ConditionalDatePeriod.Yesterday => "yesterday", ConditionalDatePeriod.Last7Days => "last7Days",
        ConditionalDatePeriod.ThisWeek => "thisWeek", ConditionalDatePeriod.LastWeek => "lastWeek", ConditionalDatePeriod.LastMonth => "lastMonth",
        ConditionalDatePeriod.Tomorrow => "tomorrow", ConditionalDatePeriod.NextWeek => "nextWeek", ConditionalDatePeriod.NextMonth => "nextMonth",
        ConditionalDatePeriod.ThisMonth => "thisMonth", _ => throw new ArgumentOutOfRangeException(nameof(period))
    };
    public static string Formula(ConditionalDatePeriod period, CellPosition anchor)
    {
        var cell = new CellRectangle(anchor.Row, anchor.Column, anchor.Row, anchor.Column).ToString().Split(':')[0];
        return "=" + (period switch
        {
            ConditionalDatePeriod.Today => $"FLOOR({cell},1)=TODAY()",
            ConditionalDatePeriod.Yesterday => $"FLOOR({cell},1)=TODAY()-1",
            ConditionalDatePeriod.Tomorrow => $"FLOOR({cell},1)=TODAY()+1",
            ConditionalDatePeriod.Last7Days => $"AND(TODAY()-FLOOR({cell},1)<=6,FLOOR({cell},1)<=TODAY())",
            ConditionalDatePeriod.ThisWeek => $"AND(TODAY()-ROUNDDOWN({cell},0)<=WEEKDAY(TODAY())-1,ROUNDDOWN({cell},0)-TODAY()<=7-WEEKDAY(TODAY()))",
            ConditionalDatePeriod.LastWeek => $"AND(TODAY()-ROUNDDOWN({cell},0)>=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN({cell},0)<(WEEKDAY(TODAY())+7))",
            ConditionalDatePeriod.NextWeek => $"AND(ROUNDDOWN({cell},0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN({cell},0)-TODAY()<(15-WEEKDAY(TODAY())))",
            ConditionalDatePeriod.LastMonth => $"AND(MONTH({cell})=MONTH(EDATE(TODAY(),0-1)),YEAR({cell})=YEAR(EDATE(TODAY(),0-1)))",
            ConditionalDatePeriod.NextMonth => $"AND(MONTH({cell})=MONTH(EDATE(TODAY(),0+1)),YEAR({cell})=YEAR(EDATE(TODAY(),0+1)))",
            ConditionalDatePeriod.ThisMonth => $"AND(MONTH({cell})=MONTH(TODAY()),YEAR({cell})=YEAR(TODAY()))",
            _ => throw new ArgumentOutOfRangeException(nameof(period))
        });
    }
    public static bool Matches(double serial, ExcelDateSystem system, ConditionalDatePeriod period, DateTime today)
    {
        if (!double.IsFinite(serial) || serial < 0) return false;
        today = DateTime.SpecifyKind(today.Date, DateTimeKind.Unspecified);
        var start = today; var end = today.AddDays(1);
        switch (period)
        {
            case ConditionalDatePeriod.Yesterday: start = today.AddDays(-1); end = today; break;
            case ConditionalDatePeriod.Tomorrow: start = today.AddDays(1); end = today.AddDays(2); break;
            case ConditionalDatePeriod.Last7Days: start = today.AddDays(-6); break;
            case ConditionalDatePeriod.ThisWeek: case ConditionalDatePeriod.LastWeek: case ConditionalDatePeriod.NextWeek:
                start = today.AddDays(-(int)today.DayOfWeek + (period == ConditionalDatePeriod.LastWeek ? -7 : period == ConditionalDatePeriod.NextWeek ? 7 : 0)); end = start.AddDays(7); break;
            case ConditionalDatePeriod.ThisMonth: case ConditionalDatePeriod.LastMonth: case ConditionalDatePeriod.NextMonth:
                start = new DateTime(today.Year, today.Month, 1).AddMonths(period == ConditionalDatePeriod.LastMonth ? -1 : period == ConditionalDatePeriod.NextMonth ? 1 : 0); end = start.AddMonths(1); break;
            case ConditionalDatePeriod.Today: break;
            default: throw new ArgumentOutOfRangeException(nameof(period));
        }
        return serial >= ExcelSerialDate.FromDateTime(start, system).Serial && serial < ExcelSerialDate.FromDateTime(end, system).Serial;
    }
}
