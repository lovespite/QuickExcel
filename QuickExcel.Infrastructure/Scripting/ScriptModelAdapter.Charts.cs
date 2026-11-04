using System.Collections.Immutable;
using System.Text.Json;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Runtime;

internal sealed partial class ScriptModelAdapter
{
    private static WorksheetChartKind ChartKind(JsonElement value) => EnumName("ChartType", value) switch
    {
        "columnClustered" => WorksheetChartKind.Column, "line" => WorksheetChartKind.Line, "pie" => WorksheetChartKind.Pie,
        "xyscatter" => WorksheetChartKind.Scatter, "waterfall" => WorksheetChartKind.Waterfall, "funnel" => WorksheetChartKind.Funnel,
        _ => throw Unsupported("此图表类型尚未建模。")
    };
    private static int ChartType(WorksheetChartKind kind) => EnumNumber("ChartType", kind switch
    { WorksheetChartKind.Column => "columnClustered", WorksheetChartKind.Line or WorksheetChartKind.SmoothLine => "line", WorksheetChartKind.Pie => "pie", WorksheetChartKind.Scatter => "xyscatter", WorksheetChartKind.Waterfall => "waterfall", _ => "funnel" });
    private static void SeriesBy(JsonElement[] args, int at)
    { if (args.Length > at && args[at].ValueKind != JsonValueKind.Null && EnumName("ChartSeriesBy", args[at]) == "rows") throw Unsupported("图表目前按列组织系列；首行标题、首列分类。"); }
    private object ChartHandle(WorksheetSnapshot sheet, WorksheetChart chart) => Handle("Chart", new { sheetId = sheet.Id.ToString(), chartId = chart.ScriptId.ToString() });
    private object? WorksheetCharts(WorksheetSnapshot sheet, string member, JsonElement[] args)
    {
        if (member == "getCharts") { Arity(args, 0); Spend(Math.Max(1, sheet.Data.Charts.Length), false); return sheet.Data.Charts.Select(c => ChartHandle(sheet, c)).ToArray(); }
        if (member == "getChart")
        {
            Arity(args, 1); Spend(1, false); var name = Text(args[0]);
            var chart = sheet.Data.Charts.FirstOrDefault(c => c.Name.Equals(name, StringComparison.OrdinalIgnoreCase) || c.ScriptId.ToString() == name);
            return chart is null ? null : ChartHandle(sheet, chart);
        }
        Arity(args, 2, 3); SeriesBy(args, 2); var range = Address(args[1], sheet.Id); Spend(checked((int)Math.Min(range.Count, MaxReadCells + 1)), false);
        var kind = ChartKind(args[0]); var number = 1; while (sheet.Data.Charts.Any(c => c.Name == "Chart " + number)) number++;
        var added = SetChartCommand.FromRange(sheet, range, kind, new() { Kind = DrawingAnchorKind.OneCell, Width = 5_334_000, Height = 3_048_000 }) with { Name = "Chart " + number };
        Write(sheet.Id, range, new SetChartCommand(sheet.Id, sheet.Data.Charts.Length, added)); return ChartHandle(sheet, added);
    }
    private object? ChartApi(WorksheetSnapshot sheet, JsonElement data, string name, JsonElement[] args)
    {
        var chart = sheet.Data.Charts.FirstOrDefault(c => c.ScriptId.ToString() == Text(data.GetProperty("chartId"))) ?? throw new ModelFailure("InvalidReference", "图表句柄已失效。");
        var index = sheet.Data.Charts.IndexOf(chart); var type = name.Split('.')[0]; var member = name.Split('.')[1];
        var seriesId = data.TryGetProperty("seriesId", out var sid) && sid.ValueKind != JsonValueKind.Null ? Text(sid) : null;
        var series = seriesId is null ? null : chart.Series.FirstOrDefault(s => s.ScriptId.ToString() == seriesId) ?? throw new ModelFailure("InvalidReference", "图表系列句柄已失效。");
        var part = data.TryGetProperty("part", out var p) ? Text(p) : "area";
        object Child(string t, string? section = null, ChartSeries? s = null) => Handle(t, new { sheetId = sheet.Id.ToString(), chartId = chart.ScriptId.ToString(), part = section ?? part, seriesId = (s ?? series)?.ScriptId.ToString() });
        void Save(WorksheetChart next) => Write(sheet.Id, ScalarRange, new SetChartCommand(sheet.Id, index, next));
        if (type == "Chart")
        {
            var bounds = chart.Anchor.Bounds(sheet);
            if (member == "setData")
            {
                Arity(args, 1, 2); SeriesBy(args, 1); var range = Address(args[0], sheet.Id); Spend(checked((int)range.Count), false);
                var replacement = SetChartCommand.FromRange(sheet, range, chart.Kind, chart.Anchor, chart.Title);
                Write(sheet.Id, range, new SetChartCommand(sheet.Id, index, chart with { Series = replacement.Series, OriginalXml = [], StyleParts = [] })); return null;
            }
            if (member == "setPosition")
            {
                Arity(args, 1, 2); var start = Address(args[0], sheet.Id); var origin = CellBounds(sheet, start);
                var next = bounds with { X = origin.X, Y = origin.Y };
                if (args.Length == 2)
                {
                    var end = Address(args[1], sheet.Id); var terminal = CellBounds(sheet, end);
                    next = next with { Width = terminal.X + terminal.Width - origin.X, Height = terminal.Y + terminal.Height - origin.Y };
                }
                Save(chart with { Anchor = DrawingGeometry.FromBounds(sheet, next, DrawingAnchorKind.TwoCell) }); return null;
            }
            if (member == "delete") { Arity(args, 0); Write(sheet.Id, ScalarRange, new SetChartCommand(sheet.Id, index, null)); return null; }
            Arity(args, member.StartsWith("set") ? 1 : 0); Spend(1, false);
            if (member.StartsWith("get")) return member switch
            {
                "getId" => chart.ScriptId.ToString(), "getName" => chart.Name, "getChartType" => ChartType(chart.Kind), "getPlotVisibleOnly" => chart.VisibleOnly,
                "getLeft" => bounds.X * .75, "getTop" => bounds.Y * .75, "getWidth" => bounds.Width * .75, "getHeight" => bounds.Height * .75,
                "getWorksheet" => Handle("Worksheet", new { sheetId = sheet.Id.ToString() }), "getTitle" => Child("ChartTitle", "title"), "getLegend" => Child("ChartLegend", "legend"),
                "getFormat" => Child("ChartAreaFormat", "area"), "getAxes" => Child("ChartAxes"), "getSeries" => chart.Series.Select(s => Child("ChartSeries", "series", s)).ToArray(), _ => throw Unsupported(member)
            };
            if (member == "setName" && (string.IsNullOrWhiteSpace(Text(args[0])) || sheet.Data.Charts.Any(c => c != chart && c.Name.Equals(Text(args[0]), StringComparison.OrdinalIgnoreCase)))) throw Invalid("图表名称为空或重复。");
            var updated = member switch
            {
                "setName" => chart with { Name = Text(args[0]) }, "setChartType" => SetChartCommand.ChangeKind(chart, ChartKind(args[0])),
                "setPlotVisibleOnly" => chart with { VisibleOnly = Bool(args[0]) },
                "setLeft" or "setTop" or "setWidth" or "setHeight" => chart with { Anchor = DrawingGeometry.FromBounds(sheet, member switch {
                    "setLeft" => bounds with { X = Number(args[0]) / .75 }, "setTop" => bounds with { Y = Number(args[0]) / .75 },
                    "setWidth" => bounds with { Width = Number(args[0]) / .75 }, _ => bounds with { Height = Number(args[0]) / .75 } }, chart.Anchor.Kind) },
                _ => throw Unsupported(member)
            };
            Save(updated); return null;
        }
        if (type == "ChartSeries")
        {
            if (series is null) throw Invalid("缺少系列身份。");
            if (member is "setValues" or "setXAxisValues")
            {
                Arity(args, 1); var source = Address(args[0], sheet.Id); Spend(checked((int)source.Count), false);
                if (source.Rows != 1 && source.Columns != 1) throw Invalid("系列数据需要单行或单列区域。");
                var updated = member == "setValues" ? series with { Values = Reference(sheet, source) } : series with { Categories = Reference(sheet, source) };
                WriteRules(sheet, [source], new SetChartCommand(sheet.Id, index, chart with { Series = chart.Series.Replace(series, updated) })); return null;
            }
            Arity(args, member.StartsWith("set") ? 1 : 0); Spend(1, false);
            if (member.StartsWith("get")) return member switch
            {
                "getName" => series.NameReference is null ? series.Name : SeriesName(sheet, series), "getFiltered" => series.IsFiltered,
                "getSmooth" => series.Smooth, "getMarkerSize" => series.MarkerSize, "getChartType" => ChartType(chart.Kind),
                "getFormat" => Child("ChartSeriesFormat"), _ => throw Unsupported(member)
            };
            var next = member switch
            {
                "setName" => series with { Name = Text(args[0]), NameReference = null }, "setFiltered" => series with { IsFiltered = Bool(args[0]) },
                "setSmooth" when chart.Kind is WorksheetChartKind.Line or WorksheetChartKind.SmoothLine or WorksheetChartKind.Scatter => series with { Smooth = Bool(args[0]) },
                "setMarkerSize" => series with { MarkerSize = Number(args[0]) }, _ => throw Unsupported(member)
            };
            Save(chart with { Series = chart.Series.Replace(series, next) }); return null;
        }
        Arity(args, member.StartsWith("set") ? 1 : 0); Spend(1, false);
        if (type is "ChartAxes" or "ChartAxis")
        {
            if (chart.Kind is WorksheetChartKind.Pie or WorksheetChartKind.Funnel) throw Unsupported("此图表类型没有坐标轴。");
            if (type == "ChartAxes") return member switch { "getCategoryAxis" => Child("ChartAxis", "categoryAxis"), "getValueAxis" => Child("ChartAxis", "valueAxis"), _ => throw Unsupported(member) };
            var category = part == "categoryAxis";
            if (category && chart.Kind != WorksheetChartKind.Scatter) throw Unsupported("分类轴刻度和数字格式仅用于散点图数值 X 轴。");
            if (member.StartsWith("get")) return member switch { "getMinimum" => category ? chart.XMinimum : chart.Minimum, "getMaximum" => category ? chart.XMaximum : chart.Maximum, "getNumberFormat" => category ? chart.XNumberFormat : chart.NumberFormat, _ => throw Unsupported(member) };
            Save(member switch {
                "setMinimum" => category ? chart with { XMinimum = Number(args[0]) } : chart with { Minimum = Number(args[0]) },
                "setMaximum" => category ? chart with { XMaximum = Number(args[0]) } : chart with { Maximum = Number(args[0]) },
                "setNumberFormat" => category ? chart with { XNumberFormat = FormatCode(args[0]) } : chart with { NumberFormat = FormatCode(args[0]) }, _ => throw Unsupported(member) }); return null;
        }
        if (member == "getFormat") return Child(type == "ChartTitle" ? "ChartTitleFormat" : "ChartLegendFormat");
        if (member == "getFont") return Child("ChartFont");
        if (member == "getFill") return Child("ChartFill");
        if (type == "ChartTitle")
        {
            if (member == "getText") return chart.Title;
            if (member == "getVisible") return chart.Title.Length != 0;
            if (member == "setText") { Save(chart with { Title = Text(args[0]) }); return null; }
            throw Unsupported(member);
        }
        if (type == "ChartLegend")
        {
            if (member == "getVisible") return chart.Legend is not null;
            if (member == "getPosition") return EnumNumber("ChartLegendPosition", chart.Legend switch { "t" => "top", "b" => "bottom", "l" => "left", "r" => "right", _ => "invalid" });
            if (member == "setVisible") { Save(chart with { Legend = Bool(args[0]) ? chart.Legend ?? "r" : null }); return null; }
            if (member == "setPosition") { Save(chart with { Legend = EnumName("ChartLegendPosition", args[0]) switch { "top" => "t", "bottom" => "b", "left" => "l", "right" => "r", _ => throw Unsupported("此图例位置未支持。") } }); return null; }
        }
        if (type == "ChartFill")
        {
            if (part is not ("area" or "series")) throw Unsupported("此对象填充尚未建模。");
            if (member == "getSolidColor") return Rgb(series?.Color ?? chart.Appearance.Background);
            if (member != "setSolidColor") throw Unsupported(member);
            var color = Argb(args[0]); Save(series is not null ? chart with { Series = chart.Series.Replace(series, series with { Color = color }) } : chart with { Appearance = chart.Appearance with { Background = color } }); return null;
        }
        if (type == "ChartFont")
        {
            var font = part == "title" ? chart.Appearance.Title : chart.Appearance.Legend;
            if (member.StartsWith("get")) return member switch { "getBold" => font.Bold, "getItalic" => font.Italic, "getName" => font.FontFamily ?? "Calibri", "getSize" => font.Size, "getColor" => Rgb(font.Color), _ => throw Unsupported(member) };
            var next = member switch { "setBold" => font with { Bold = Bool(args[0]) }, "setItalic" => font with { Italic = Bool(args[0]) }, "setName" => font with { FontFamily = Text(args[0]) }, "setSize" => font with { Size = Number(args[0]) }, "setColor" => font with { Color = Argb(args[0]) }, _ => throw Unsupported(member) };
            Save(chart with { Appearance = part == "title" ? chart.Appearance with { Title = next } : chart.Appearance with { Legend = next } }); return null;
        }
        throw Unsupported(name);
    }
    private static DrawingBounds CellBounds(WorksheetSnapshot sheet, CellRectangle range) => new(DrawingGeometry.ColumnStart(sheet, range.Left), DrawingGeometry.RowStart(sheet, range.Top),
        DrawingGeometry.ColumnStart(sheet, range.Right + 1) - DrawingGeometry.ColumnStart(sheet, range.Left), DrawingGeometry.RowStart(sheet, range.Bottom + 1) - DrawingGeometry.RowStart(sheet, range.Top));
    private string SeriesName(WorksheetSnapshot sheet, ChartSeries series)
    {
        var r = WorksheetChart.ParseReference(series.NameReference!); var source = Model.View.Sheets.Single(s => s.Name == r.Sheet);
        var area = new CellRectangle(r.Row, r.Column, r.Row, r.Column); Model.PrepareValueRead(source.Id, area, cancellation);
        var cell = Model.ReadCell(source.Id, new(r.Row, r.Column), calculated: true); EnsureReadable(cell);
        return CellDisplayFormatter.Format(cell?.Value ?? CellValue.Blank, Model.View.Styles[source.GetStyleId(r.Row, r.Column)], Model.View.DateSystem);
    }
}
