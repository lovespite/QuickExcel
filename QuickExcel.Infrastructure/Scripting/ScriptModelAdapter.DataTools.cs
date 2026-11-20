using System.Text.Json;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Runtime;

internal sealed partial class ScriptModelAdapter
{
    private object? RangeDataApi(WorksheetSnapshot sheet, CellRectangle range, string member, JsonElement[] args)
    {
        switch (member)
        {
            case "calculate":
                Arity(args, 0);
                Model.PrepareValueRead(sheet.Id, range, cancellation);
                Spend(1, false);
                return null;
            case "getHyperlink":
                return GetHyperlink(sheet, range, args);
            case "setHyperlink":
                SetHyperlink(sheet, range, args);
                return null;
            case "find":
                return FindRange(sheet, range, args);
            case "replaceAll":
                return ReplaceAll(sheet, range, args);
            case "removeDuplicates":
                return RemoveDuplicates(sheet, range, args);
            case "autoFill":
                AutoFill(sheet, range, args);
                return null;
            case "flashFill":
                FlashFillRange(sheet, range, args);
                return null;
            case "copyFrom":
                CopyFrom(sheet, range, args);
                return null;
            case "getSurroundingRegion":
                return GetSurroundingRegion(sheet, range, args);
            case "getRangeEdge":
                return GetRangeEdge(sheet, range, args);
            case "group":
            case "ungroup":
                GroupRange(sheet, range, member, args);
                return null;
            case "hideGroupDetails":
            case "showGroupDetails":
                GroupDetails(sheet, range, member, args);
                return null;
            case "clearOrResetContents":
                ClearOrResetContents(sheet, range, args);
                return null;
            default:
                throw Unsupported(member);
        }
    }

    private object? GetHyperlink(WorksheetSnapshot sheet, CellRectangle range, JsonElement[] args)
    {
        Arity(args, 0); Spend(1, false);
        var link = sheet.Annotations.Links.GetValueOrDefault(new CellPosition(range.Top, range.Left));
        if (link is null) return null;
        var result = new Dictionary<string, string>();
        if (link.Target is not null) result["address"] = link.Target;
        if (link.Location is not null) result["documentReference"] = link.Location;
        if (link.Display is not null) result["textToDisplay"] = link.Display;
        if (link.Tooltip is not null) result["screenTip"] = link.Tooltip;
        return result;
    }

    private void SetHyperlink(WorksheetSnapshot sheet, CellRectangle range, JsonElement[] args)
    {
        Arity(args, 1);
        if (range.Count != 1) throw Unsupported("setHyperlink 目前只支持单个单元格。");
        var input = args[0];
        Fields(input, "address", "documentReference", "screenTip", "textToDisplay");
        string? Optional(string name) =>
            input.TryGetProperty(name, out var value) && value.ValueKind != JsonValueKind.Null ? Text(value) : null;
        var target = Optional("address"); var location = Optional("documentReference");
        var display = Optional("textToDisplay"); var tooltip = Optional("screenTip");
        CellHyperlink? link = string.IsNullOrEmpty(target) && string.IsNullOrEmpty(location) ? null
            : new CellHyperlink(string.IsNullOrEmpty(target) ? null : target, string.IsNullOrEmpty(location) ? null : location,
                string.IsNullOrEmpty(display) ? null : display, string.IsNullOrEmpty(tooltip) ? null : tooltip);
        var position = new CellPosition(range.Top, range.Left);
        Write(sheet.Id, new CellRectangle(range.Top, range.Left, range.Top, range.Left),
            new SetHyperlinkCommand(sheet.Id, position, link));
    }

    private object? FindRange(WorksheetSnapshot sheet, CellRectangle range, JsonElement[] args)
    {
        Arity(args, 1, 2);
        var query = Text(args[0]);
        if (query.Length == 0 || query.Length > WorkbookLimits.MaxCellTextLength) throw Invalid("查找内容不能为空或超过单元格文本上限。");
        var (completeMatch, matchCase, backwards) = FindCriteria(args, 1);
        var wholeSheet = range.Count == 1;
        if (sheet.Cells.Count == 0) return null;
        var search = wholeSheet
            ? new CellRectangle(sheet.Cells.Keys.Min(p => p.Row), sheet.Cells.Keys.Min(p => p.Column),
                sheet.Cells.Keys.Max(p => p.Row), sheet.Cells.Keys.Max(p => p.Column))
            : range;
        if (search.Count > MaxReadCells) throw new ModelFailure("BudgetExceeded", "查找范围超过读取预算。");
        var positions = sheet.Cells.Keys.Where(p => search.Contains(p.Row, p.Column))
            .OrderBy(p => p.Row).ThenBy(p => p.Column).ToArray();
        if (positions.Length == 0) return null;
        Model.PrepareValueRead(sheet.Id, search, cancellation);
        Spend(positions.Length, false);
        var values = Model.ValueView; var valueSheet = values.GetSheet(sheet.Id);
        if (wholeSheet)
        {
            var start = new CellPosition(range.Top, range.Left);
            bool After(CellPosition p) => p.Row > start.Row || p.Row == start.Row && p.Column > start.Column;
            positions = backwards
                ? positions.Where(p => !After(p)).Reverse().Concat(positions.Where(After).Reverse()).ToArray()
                : positions.Where(After).Concat(positions.Where(p => !After(p))).ToArray();
        }
        else if (backwards) Array.Reverse(positions);
        var comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        foreach (var position in positions)
        {
            cancellation.ThrowIfCancellationRequested();
            var cell = valueSheet.GetCell(position.Row, position.Column);
            var style = TablePresentation.NumberStyle(valueSheet, values.Styles, position.Row, position.Column);
            var text = CellDisplayFormatter.Format(cell?.Value ?? CellValue.Blank, style, values.DateSystem);
            var matched = completeMatch ? text.Equals(query, comparison) : text.Contains(query, comparison);
            if (matched) return RangeHandle(sheet.Id, new(position.Row, position.Column, position.Row, position.Column));
        }
        return null;
    }

    private (bool CompleteMatch, bool MatchCase, bool Backwards) FindCriteria(JsonElement[] args, int index)
    {
        if (args.Length <= index || args[index].ValueKind == JsonValueKind.Null) return (false, false, false);
        var criteria = args[index];
        Fields(criteria, "completeMatch", "matchCase", "searchDirection");
        var complete = criteria.TryGetProperty("completeMatch", out var c) && c.ValueKind != JsonValueKind.Null && Bool(c);
        var matchCase = criteria.TryGetProperty("matchCase", out var m) && m.ValueKind != JsonValueKind.Null && Bool(m);
        var backwards = criteria.TryGetProperty("searchDirection", out var d) && d.ValueKind != JsonValueKind.Null &&
            EnumName("SearchDirection", d) == "backwards";
        return (complete, matchCase, backwards);
    }

    private (bool CompleteMatch, bool MatchCase) ReplaceCriteria(JsonElement[] args, int index)
    {
        if (args.Length <= index || args[index].ValueKind == JsonValueKind.Null) return (false, false);
        var criteria = args[index];
        Fields(criteria, "completeMatch", "matchCase");
        var complete = criteria.TryGetProperty("completeMatch", out var c) && c.ValueKind != JsonValueKind.Null && Bool(c);
        var matchCase = criteria.TryGetProperty("matchCase", out var m) && m.ValueKind != JsonValueKind.Null && Bool(m);
        return (complete, matchCase);
    }

    private int ReplaceAll(WorksheetSnapshot sheet, CellRectangle range, JsonElement[] args)
    {
        Arity(args, 2, 3);
        var query = Text(args[0]); var replacement = Text(args[1]);
        if (query.Length == 0 || query.Length > WorkbookLimits.MaxCellTextLength) throw Invalid("查找内容不能为空或超过单元格文本上限。");
        WorkbookLimits.ValidateText(replacement);
        var (completeMatch, matchCase) = ReplaceCriteria(args, 2);
        var candidates = sheet.Cells.Where(p => range.Contains(p.Key.Row, p.Key.Column)).Select(p => p.Key)
            .OrderBy(p => p.Row).ThenBy(p => p.Column).ToArray();
        if (candidates.Length == 0) return 0;
        var used = new CellRectangle(candidates.Min(p => p.Row), candidates.Min(p => p.Column),
            candidates.Max(p => p.Row), candidates.Max(p => p.Column));
        if (used.Count > MaxReadCells) throw new ModelFailure("BudgetExceeded", "替换范围超过读取预算。");
        Model.PrepareValueRead(sheet.Id, used, cancellation);
        Spend(candidates.Length, false);
        var values = Model.ValueView; var valueSheet = values.GetSheet(sheet.Id);
        var comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        var edits = new List<CellEdit>();
        var replacements = 0;
        foreach (var position in candidates)
        {
            cancellation.ThrowIfCancellationRequested();
            var original = sheet.GetCell(position.Row, position.Column);
            if (original is null) continue;
            var calculated = valueSheet.GetCell(position.Row, position.Column);
            var style = TablePresentation.NumberStyle(valueSheet, values.Styles, position.Row, position.Column);
            var text = CellDisplayFormatter.Format(calculated?.Value ?? CellValue.Blank, style, values.DateSystem);
            var matched = completeMatch ? text.Equals(query, comparison) : text.Contains(query, comparison);
            if (!matched) continue;
            if (original.Content.Kind == CellContentKind.Formula)
                throw Unsupported("replaceAll 不替换公式结果；请先在 Excel 中转为值或缩小范围。");
            if (original.Content.RichText is not null)
                throw Unsupported("replaceAll 不隐式展开富文本；请先显式转为纯文本。");
            replacements += completeMatch ? 1 : CountOccurrences(text, query, comparison);
            var next = completeMatch ? replacement : text.Replace(query, replacement, comparison);
            CellContent content;
            if (original.Content.LiteralValue.Kind == CellValueKind.Text)
                content = CellContent.FromLiteral(CellValue.FromText(next), original.Content.StyleId);
            else
            {
                content = CellInput.Parse(next, original.Content.StyleId);
                if (content.Kind == CellContentKind.Formula) throw Invalid("值替换不能创建公式。");
            }
            edits.Add(new(new(sheet.Id, position.Row, position.Column), content));
        }
        if (edits.Count == 0) return 0;
        ExecuteObject(sheet.Id, new SetCellsCommand(edits), Math.Max(1, edits.Count));
        return replacements;
    }

    private static int CountOccurrences(string text, string query, StringComparison comparison)
    {
        var count = 0;
        for (var index = 0; (index = text.IndexOf(query, index, comparison)) >= 0; index += query.Length) count++;
        return count;
    }

    private object RemoveDuplicates(WorksheetSnapshot sheet, CellRectangle range, JsonElement[] args)
    {
        Arity(args, 2);
        if (range.Count > DataCleaning.MaxCells) throw new ModelFailure("BudgetExceeded", "删除重复项最多处理 20000 格。");
        if (args[0].ValueKind != JsonValueKind.Array) throw Invalid("columns 需要列索引数组。");
        var columns = args[0].EnumerateArray().Select(Int).ToArray();
        if (columns.Length is 0 or > 1024 || columns.Any(c => c < 0 || c >= range.Columns) || columns.Distinct().Count() != columns.Length)
            throw Invalid("columns 必须是区域内不重复的零基列索引。");
        Model.PrepareValueRead(sheet.Id, range, cancellation);
        var preview = DataCleaning.RemoveDuplicates(Model.ValueView, sheet.Id, range,
            columns.Select(c => range.Left + c), Bool(args[1]), expandToWholeRows: false);
        var edits = preview.Changes.Select(change => new CellEdit(change.Address, change.After)).ToArray();
        if (edits.Length > 0) ExecuteObject(sheet.Id, new SetCellsCommand(edits), Math.Max(1, edits.Length));
        var removed = preview.RemovedRows.Length;
        return Handle("RemoveDuplicatesResult", new Dictionary<string, object?> { ["removed"] = removed, ["uniqueRemaining"] = range.Rows - (Bool(args[1]) ? 1 : 0) - removed });
    }

    private void AutoFill(WorksheetSnapshot sheet, CellRectangle range, JsonElement[] args)
    {
        Arity(args, 0, 2);
        var hasDestination = args.Length >= 1 && args[0].ValueKind != JsonValueKind.Null;
        var extent = hasDestination ? Bounding(range, Address(args[0], sheet.Id)) : AutoFillExtent(sheet, range);
        if (extent == range) return;
        var type = args.Length >= 2 && args[1].ValueKind != JsonValueKind.Null
            ? EnumName("AutoFillType", args[1]) : "fillDefault";
        switch (type)
        {
            case "fillDefault":
            case "fillSeries":
                Write(sheet.Id, extent, new FillCellsCommand(sheet.Id, range, extent, CellFillMode.Auto));
                return;
            case "fillCopy":
                Write(sheet.Id, extent, new FillCellsCommand(sheet.Id, range, extent, CellFillMode.Copy));
                return;
            case "fillValues":
                FillValues(sheet, range, extent);
                return;
            case "fillFormats":
                FillFormats(sheet, range, extent);
                return;
            case "fillDays":
                FillSeries(sheet, range, extent, CellSeriesKind.Days, SeriesStep(sheet, range, extent, CellSeriesKind.Days));
                return;
            case "fillMonths":
                FillSeries(sheet, range, extent, CellSeriesKind.Months, SeriesStep(sheet, range, extent, CellSeriesKind.Months));
                return;
            case "fillYears":
                FillSeries(sheet, range, extent, CellSeriesKind.Years, SeriesStep(sheet, range, extent, CellSeriesKind.Years));
                return;
            case "linearTrend":
                FillSeries(sheet, range, extent, CellSeriesKind.Numbers, SeriesStep(sheet, range, extent, CellSeriesKind.Numbers));
                return;
            case "growthTrend":
                GrowthTrend(sheet, range, extent);
                return;
            case "fillWeekdays":
                Weekdays(sheet, range, extent);
                return;
            default:
                throw Unsupported("该 AutoFillType 尚未实现。");
        }
    }

    private void FillValues(WorksheetSnapshot sheet, CellRectangle source, CellRectangle extent)
    {
        var edits = new List<CellEdit>();
        for (var row = extent.Top; row <= extent.Bottom; row++)
        for (var column = extent.Left; column <= extent.Right; column++)
        {
            if (source.Contains(row, column)) continue;
            var sr = source.Top + Mod(row - source.Top, source.Rows);
            var sc = source.Left + Mod(column - source.Left, source.Columns);
            edits.Add(new(new(sheet.Id, row, column),
                CellContent.FromLiteral(sheet.GetCell(sr, sc)?.Value ?? CellValue.Blank, sheet.GetStyleId(row, column))));
        }
        if (edits.Count > 0) Write(sheet.Id, extent, new SetCellsCommand(edits));
    }

    private void FillFormats(WorksheetSnapshot sheet, CellRectangle source, CellRectangle extent)
    {
        var edits = new List<CellEdit>();
        for (var row = extent.Top; row <= extent.Bottom; row++)
        for (var column = extent.Left; column <= extent.Right; column++)
        {
            if (source.Contains(row, column)) continue;
            var sr = source.Top + Mod(row - source.Top, source.Rows);
            var sc = source.Left + Mod(column - source.Left, source.Columns);
            var content = sheet.GetCell(row, column)?.Content ?? CellContent.FromLiteral(CellValue.Blank);
            edits.Add(new(new(sheet.Id, row, column), content.WithStyle(sheet.GetStyleId(sr, sc))));
        }
        if (edits.Count > 0) Write(sheet.Id, extent, new SetCellsCommand(edits));
    }

    private static int Mod(int value, int divisor) => (value % divisor + divisor) % divisor;

    private void FillSeries(WorksheetSnapshot sheet, CellRectangle source, CellRectangle extent, CellSeriesKind kind, double step)
    {
        RequireSeriesOrigin(source, extent);
        Write(sheet.Id, extent, new FillSeriesCommand(sheet.Id, extent, kind, step));
    }

    private (double First, double? Second)? NumericSeeds(WorksheetSnapshot sheet, CellRectangle source, bool vertical)
    {
        var first = sheet.GetCell(source.Top, source.Left);
        if (first?.Content.Kind != CellContentKind.Literal || first.Content.LiteralValue.Kind != CellValueKind.Number) return null;
        var value = first.Content.LiteralValue.NumberValue;
        CellSnapshot? second = null;
        if (vertical && source.Rows > 1) second = sheet.GetCell(source.Top + 1, source.Left);
        else if (!vertical && source.Columns > 1) second = sheet.GetCell(source.Top, source.Left + 1);
        var next = second?.Content.Kind == CellContentKind.Literal && second.Content.LiteralValue.Kind == CellValueKind.Number
            ? second.Content.LiteralValue.NumberValue : (double?)null;
        return (value, next);
    }

    private double SeriesStep(WorksheetSnapshot sheet, CellRectangle source, CellRectangle extent, CellSeriesKind kind)
    {
        var vertical = extent.Rows > source.Rows;
        var seeds = NumericSeeds(sheet, source, vertical);
        if (seeds is null) return 1;
        var (first, second) = seeds.Value;
        if (second is null) return 1;
        if (kind is CellSeriesKind.Numbers or CellSeriesKind.Days) return second.Value - first;
        if (!new ExcelSerialDate(first, Model.View.DateSystem).TryToDateTime(out var start) ||
            !new ExcelSerialDate(second.Value, Model.View.DateSystem).TryToDateTime(out var end))
            throw Invalid("日期序列的起始值无效。");
        return kind == CellSeriesKind.Years
            ? end.Year - start.Year
            : (end.Year * 12 + end.Month) - (start.Year * 12 + start.Month);
    }

    private static void RequireSeriesOrigin(CellRectangle source, CellRectangle extent)
    {
        if (extent.Top != source.Top || extent.Left != source.Left)
            throw Unsupported("序列填充目前只支持从源区域向下或向右扩展。");
        var vertical = extent.Rows > source.Rows;
        if (vertical && (extent.Left != source.Left || extent.Right != source.Right) ||
            !vertical && (extent.Top != source.Top || extent.Bottom != source.Bottom))
            throw Invalid("序列填充只能沿一个方向扩展。");
    }

    private void GrowthTrend(WorksheetSnapshot sheet, CellRectangle source, CellRectangle extent)
    {
        RequireSeriesOrigin(source, extent);
        var vertical = extent.Rows > source.Rows;
        var seeds = NumericSeeds(sheet, source, vertical);
        if (seeds is null || seeds.Value.Second is not { } second)
            throw Invalid("增长趋势需要两个非零数值种子。");
        var first = seeds.Value.First;
        if (first == 0 || !double.IsFinite(second / first))
            throw Invalid("增长趋势需要两个非零数值种子。");
        var ratio = second / first;
        FillManual(sheet, source, extent, vertical, index =>
            CellValue.FromNumber(first * Math.Pow(ratio, index)));
    }

    private void Weekdays(WorksheetSnapshot sheet, CellRectangle source, CellRectangle extent)
    {
        RequireSeriesOrigin(source, extent);
        var vertical = extent.Rows > source.Rows;
        var seeds = NumericSeeds(sheet, source, vertical) ?? throw Invalid("工作日填充需要数字日期种子。");
        if (!new ExcelSerialDate(seeds.First, Model.View.DateSystem).TryToDateTime(out var start))
            throw Invalid("工作日填充的起始日期无效。");
        FillManual(sheet, source, extent, vertical, index =>
        {
            var date = start;
            for (var i = 0; i < index; i++)
            {
                do { date = date.AddDays(1); } while (date.DayOfWeek is DayOfWeek.Saturday or DayOfWeek.Sunday);
            }
            return CellValue.FromNumber(ExcelSerialDate.FromDateTime(DateTime.SpecifyKind(date, DateTimeKind.Unspecified), Model.View.DateSystem).Serial);
        });
    }

    private void FillManual(WorksheetSnapshot sheet, CellRectangle source, CellRectangle extent, bool vertical, Func<int, CellValue> valueAt)
    {
        var edits = new List<CellEdit>();
        for (var row = extent.Top; row <= extent.Bottom; row++)
        for (var column = extent.Left; column <= extent.Right; column++)
        {
            if (source.Contains(row, column)) continue;
            var index = vertical ? row - extent.Top : column - extent.Left;
            edits.Add(new(new(sheet.Id, row, column),
                CellContent.FromLiteral(valueAt(index), sheet.GetStyleId(row, column))));
        }
        if (edits.Count > 0) Write(sheet.Id, extent, new SetCellsCommand(edits));
    }

    private static CellRectangle Bounding(CellRectangle a, CellRectangle b) => new(
        Math.Min(a.Top, b.Top), Math.Min(a.Left, b.Left), Math.Max(a.Bottom, b.Bottom), Math.Max(a.Right, b.Right));

    private CellRectangle AutoFillExtent(WorksheetSnapshot sheet, CellRectangle source)
    {
        bool HasContent(int row, int column)
        {
            var cell = sheet.GetCell(row, column);
            return cell is not null && (cell.Content.Kind == CellContentKind.Formula || cell.Content.LiteralValue.Kind != CellValueKind.Blank);
        }
        int ColumnBottom(int column)
        {
            if (column < 0 || column >= WorkbookLimits.MaxColumns) return source.Bottom;
            var row = source.Top;
            while (row + 1 < WorkbookLimits.MaxRows && HasContent(row + 1, column)) row++;
            return row;
        }
        var bottom = source.Bottom;
        for (var column = source.Left; column <= source.Right; column++) bottom = Math.Max(bottom, ColumnBottom(column));
        if (source.Left > 0) bottom = Math.Max(bottom, ColumnBottom(source.Left - 1));
        if (source.Right + 1 < WorkbookLimits.MaxColumns) bottom = Math.Max(bottom, ColumnBottom(source.Right + 1));
        if (bottom <= source.Bottom) throw Invalid("autoFill 未提供目标区域，且相邻区域没有可延伸的数据。");
        return source with { Bottom = bottom };
    }

    private void FlashFillRange(WorksheetSnapshot sheet, CellRectangle range, JsonElement[] args)
    {
        Arity(args, 0);
        if (range.Columns != 1) throw Invalid("flashFill 需要单列目标区域。");
        bool HasContent(int row, int column)
        {
            var cell = sheet.GetCell(row, column);
            return cell is not null && (cell.Content.Kind == CellContentKind.Formula || cell.Content.LiteralValue.Kind != CellValueKind.Blank);
        }
        var target = range.Left; var left = target; var right = target;
        while (left > 0 && Enumerable.Range(range.Top, range.Rows).Any(row => HasContent(row, left - 1))) left--;
        while (right + 1 < WorkbookLimits.MaxColumns && Enumerable.Range(range.Top, range.Rows).Any(row => HasContent(row, right + 1))) right++;
        var records = new CellRectangle(range.Top, left, range.Bottom, right);
        if (records.Columns < 2) throw Invalid("flashFill 需要目标列旁边有来源数据。");
        if (records.Count > DataCleaning.MaxCells) throw new ModelFailure("BudgetExceeded", "快速填充最多处理 20000 格。");
        Model.PrepareValueRead(sheet.Id, records, cancellation);
        var preview = FlashFill.Preview(Model.ValueView, sheet.Id, records, target);
        var edits = preview.Changes.Select(change => new CellEdit(change.Address, change.After)).ToArray();
        if (edits.Length > 0) ExecuteObject(sheet.Id, new SetCellsCommand(edits), Math.Max(1, edits.Length));
    }
}
