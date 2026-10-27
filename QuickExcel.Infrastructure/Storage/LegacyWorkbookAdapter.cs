using System.Globalization;
using System.Runtime.CompilerServices;
using ExcelIO;
using QuickExcel.Domain;

namespace QuickExcel.Storage;

/// <summary>Copies at the legacy boundary; no ExcelIO object is retained by a domain snapshot.</summary>
public static class LegacyWorkbookAdapter
{
    // XLSX's explicit coordinates distinguish actual empty text cells from dense padding introduced by ExcelIO.
    private static readonly ConditionalWeakTable<XlWorkbook, List<Dictionary<int, List<int>>>> ExplicitLayouts = new();
    private const long MaxImageBytes = 64L * 1024 * 1024;

    internal static void RegisterLayout(XlWorkbook workbook, List<Dictionary<int, List<int>>> layouts)
    {
        ExplicitLayouts.Remove(workbook);
        ExplicitLayouts.Add(workbook, layouts);
    }

    public static WorkbookSnapshot FromWorkbook(XlWorkbook workbook, string scriptSource = "")
    {
        ArgumentNullException.ThrowIfNull(workbook);
        ArgumentNullException.ThrowIfNull(scriptSource);
        ValidateLegacyBudget(workbook);
        var styles = new List<StyleDefinition> { new() };
        var styleIds = new Dictionary<StyleDefinition, int> { [styles[0]] = 0 };
        int StyleId(XlStyle? style)
        {
            if (style is null) return 0;
            var value = FromStyle(style);
            if (styleIds.TryGetValue(value, out var existing)) return existing;
            var id = styles.Count;
            styles.Add(value);
            styleIds.Add(value, id);
            return id;
        }

        ExplicitLayouts.TryGetValue(workbook, out var layouts);
        var sheets = new List<WorksheetSnapshot>();
        for (var sheetIndex = 0; sheetIndex < workbook.Count; sheetIndex++)
        {
            var sheet = workbook[sheetIndex];
            var cells = new Dictionary<CellPosition, CellSnapshot>();
            var rows = new Dictionary<int, RowMetadata>();
            var columns = new Dictionary<int, ColumnMetadata>();
            for (var rowIndex = 0; rowIndex < sheet.Rows.Count; rowIndex++)
            {
                var row = sheet.Rows[rowIndex];
                var rowStyle = StyleId(row.Style);
                if (row.Height is not null || row.Hidden || rowStyle != 0)
                    rows.Add(rowIndex, new RowMetadata(row.Height, row.Hidden, rowStyle));
                IEnumerable<int> indices = layouts is not null && sheetIndex < layouts.Count
                    ? layouts[sheetIndex].GetValueOrDefault(rowIndex) ?? []
                    : Enumerable.Range(0, row.Cells.Count);
                foreach (var columnIndex in indices)
                {
                    var cell = row.Cells[columnIndex];
                    var styleId = StyleId(cell.Style);
                    if (cell.HasFormula)
                    {
                        var formula = cell.Formula!.StartsWith('=') ? cell.Formula : "=" + cell.Formula;
                        cells.Add(new CellPosition(rowIndex, columnIndex), new CellSnapshot(
                            CellContent.FromFormula(formula, styleId),
                            new CellCalculation(ParseFormulaCache(cell.Value), CalculationState.Stale, 0, 0, 0)));
                    }
                    else
                    {
                        // Literal identity is never inferred from string spelling (00123, TRUE, #VALUE!, ...).
                        cells.Add(new CellPosition(rowIndex, columnIndex), new CellSnapshot(
                            CellContent.FromLiteral(CellValue.FromText(cell.Value), styleId)));
                    }
                }
            }
            foreach (var (index, column) in sheet.Columns)
                columns.Add(index, new ColumnMetadata(column.Width, column.Hidden, StyleId(column.Style)));
            var images = sheet.Images.Select(image => new SheetImage(image.Bytes, image.Extension,
                image.RowIndex, image.ColumnIndex, image.RowSpan, image.ColumnSpan, image.PlaceInCell));
            sheets.Add(new WorksheetSnapshot(SheetId.New(), sheet.Name, cells, rows, columns,
                sheet.MergedCells, images, sheet.Options.ShowGridLines, sheet.Options.DefaultRowHeight, sheet.Options.TabColor));
        }
        if (sheets.Count == 0) sheets.Add(new WorksheetSnapshot(SheetId.New(), "Sheet1"));
        return new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, sheets, scriptSource, styles);
    }

    public static XlWorkbook ToWorkbook(WorkbookSnapshot snapshot, bool valuesOnly = false)
    {
        ArgumentNullException.ThrowIfNull(snapshot);
        // The current library has no date-system field. Refuse a change of date meaning rather than silently remap it.
        if (snapshot.DateSystem != ExcelDateSystem.Windows1900)
            throw new NotSupportedException("0.2 存储适配器尚不支持保存 1904 日期系统。");
        var plans = PlanDenseConversion(snapshot);
        var workbook = new XlWorkbook();
        var layouts = new List<Dictionary<int, List<int>>>();
        for (var sheetIndex = 0; sheetIndex < snapshot.Sheets.Length; sheetIndex++)
        {
            var source = snapshot.Sheets[sheetIndex];
            var target = workbook.NewWorksheet(source.Name);
            var (rowCount, rowWidths) = plans[sheetIndex];
            var layout = new Dictionary<int, List<int>>();
            layouts.Add(layout);
            target.Options.ShowGridLines = source.ShowGridLines;
            target.Options.DefaultRowHeight = source.DefaultRowHeight;
            target.Options.TabColor = source.TabColor;
            foreach (var (index, column) in source.Columns)
                target.Columns.Add(index, new XlColumn { Width = column.Width, Hidden = column.Hidden, Style = ToStyle(snapshot, column.StyleId) });
            target.MergedCells.AddRange(source.MergedCells);
            for (var rowIndex = 0; rowIndex < rowCount; rowIndex++)
            {
                var row = target.NewRow();
                if (source.Rows.TryGetValue(rowIndex, out var metadata))
                {
                    row.Height = metadata.Height;
                    row.Hidden = metadata.Hidden;
                    row.Style = ToStyle(snapshot, metadata.StyleId);
                }
                var width = rowWidths.GetValueOrDefault(rowIndex);
                for (var columnIndex = 0; columnIndex < width; columnIndex++) row.Cells.Add(new XlCell(row));
            }
            foreach (var (position, cell) in source.Cells)
            {
                var row = target.Rows[position.Row];
                var targetCell = row.Cells[position.Column];
                targetCell.Style = ToStyle(snapshot, cell.Content.StyleId);
                if (!valuesOnly && cell.Content.Kind == CellContentKind.Formula)
                    targetCell.SetFormula(cell.Content.FormulaText, ToLegacyText(cell.Value));
                else targetCell.SetCalculatedValue(ToLegacyText(cell.Value));
                if (!layout.TryGetValue(position.Row, out var columns)) layout.Add(position.Row, columns = []);
                columns.Add(position.Column);
            }
            foreach (var image in source.Images)
                target.AddImage(image.Data.ToArray(), image.Extension, image.RowIndex, image.ColumnIndex,
                    image.RowSpan, image.ColumnSpan, image.PlaceInCell);
        }
        RegisterLayout(workbook, layouts);
        return workbook;
    }

    /// <summary>Interpret only formula cache text; never use this to recover literal types.</summary>
    public static CellValue ParseFormulaCache(string text)
    {
        ArgumentNullException.ThrowIfNull(text);
        var error = text switch
        {
            "#NULL!" => CellErrorCode.Null,
            "#DIV/0!" => CellErrorCode.DivisionByZero,
            "#VALUE!" => CellErrorCode.Value,
            "#REF!" => CellErrorCode.Reference,
            "#NAME?" => CellErrorCode.Name,
            "#NUM!" => CellErrorCode.Number,
            "#N/A" => CellErrorCode.NotAvailable,
            "#CYCLE!" => CellErrorCode.Cycle,
            "#SCRIPT!" => CellErrorCode.Script,
            _ => (CellErrorCode?)null
        };
        if (error is { } code) return CellValue.FromError(code);
        if (text.Equals("TRUE", StringComparison.OrdinalIgnoreCase)) return CellValue.FromBoolean(true);
        if (text.Equals("FALSE", StringComparison.OrdinalIgnoreCase)) return CellValue.FromBoolean(false);
        if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) && double.IsFinite(number))
            return CellValue.FromNumber(number);
        return CellValue.FromText(text);
    }

    public static string ToLegacyText(CellValue value) => value.Kind switch
    {
        CellValueKind.Blank => "",
        CellValueKind.Text => value.TextValue,
        CellValueKind.Number => value.NumberValue.ToString("R", CultureInfo.InvariantCulture),
        CellValueKind.Boolean => value.BooleanValue ? "TRUE" : "FALSE",
        CellValueKind.Error => value.ErrorCode switch
        {
            CellErrorCode.Null => "#NULL!", CellErrorCode.DivisionByZero => "#DIV/0!",
            CellErrorCode.Value => "#VALUE!", CellErrorCode.Reference => "#REF!",
            CellErrorCode.Name => "#NAME?", CellErrorCode.Number => "#NUM!",
            CellErrorCode.NotAvailable => "#N/A", CellErrorCode.Cycle => "#CYCLE!", CellErrorCode.Script => "#SCRIPT!",
            _ => throw new InvalidOperationException("Unknown cell error.")
        },
        _ => throw new InvalidOperationException("Unknown cell value.")
    };

    private static List<(int RowCount, Dictionary<int, int> Widths)> PlanDenseConversion(WorkbookSnapshot snapshot)
    {
        long totalCells = 0, totalRows = 0, imageBytes = 0;
        var plans = new List<(int, Dictionary<int, int>)>();
        foreach (var sheet in snapshot.Sheets)
        {
            var rowCount = sheet.UsedRowCount;
            if (sheet.Rows.Count != 0) rowCount = Math.Max(rowCount, sheet.Rows.Keys.Max() + 1);
            var widths = new Dictionary<int, int>();
            foreach (var position in sheet.Cells.Keys)
                widths[position.Row] = Math.Max(widths.GetValueOrDefault(position.Row), position.Column + 1);
            foreach (var image in sheet.Images)
            {
                imageBytes += image.Data.Length;
                if (!image.PlaceInCell) continue;
                rowCount = Math.Max(rowCount, image.RowIndex + 1);
                widths[image.RowIndex] = Math.Max(widths.GetValueOrDefault(image.RowIndex), image.ColumnIndex + 1);
            }
            totalRows += rowCount;
            totalCells += widths.Values.Sum(width => (long)width);
            CheckDenseBudget(totalRows, totalCells, imageBytes);
            plans.Add((rowCount, widths));
        }
        return plans;
    }

    internal static void ValidateLegacyBudget(XlWorkbook workbook)
    {
        if (workbook.Count > WorkbookLimits.MaxSheets) throw new InvalidDataException("工作表数量超过 0.2 容量限制。");
        long totalCells = 0, totalRows = 0, imageBytes = 0;
        foreach (var sheet in workbook)
        {
            if (sheet.Rows.Count > WorkbookLimits.MaxRows || sheet.Columns.Keys.Any(index => index < 0 || index >= WorkbookLimits.MaxColumns))
                throw new InvalidDataException("工作表尺寸超过 0.2 容量限制。");
            var rowCount = sheet.Rows.Count;
            var widths = new Dictionary<int, int>();
            for (var index = 0; index < sheet.Rows.Count; index++)
            {
                var row = sheet.Rows[index];
                if (row.Cells.Count > WorkbookLimits.MaxColumns) throw new InvalidDataException("工作表列数超过 0.2 容量限制。");
                widths[index] = row.Cells.Count;
            }
            foreach (var image in sheet.Images)
            {
                if (image.RowIndex >= WorkbookLimits.MaxRows || image.ColumnIndex >= WorkbookLimits.MaxColumns ||
                    image.RowSpan > WorkbookLimits.MaxRows - image.RowIndex || image.ColumnSpan > WorkbookLimits.MaxColumns - image.ColumnIndex)
                    throw new InvalidDataException("图片范围超过 0.2 容量限制。");
                imageBytes += image.Bytes.Length;
                if (!image.PlaceInCell) continue;
                rowCount = Math.Max(rowCount, image.RowIndex + 1);
                widths[image.RowIndex] = Math.Max(widths.GetValueOrDefault(image.RowIndex), image.ColumnIndex + 1);
            }
            totalRows += rowCount;
            totalCells += widths.Values.Sum(width => (long)width);
            CheckDenseBudget(totalRows, totalCells, imageBytes);
        }
    }

    private static void CheckDenseBudget(long rows, long cells, long imageBytes)
    {
        if (rows > WorkbookLimits.MaxCells || cells > WorkbookLimits.MaxCells)
            throw new InvalidDataException("旧版 XLSX 适配所需的占位行或单元格超过 1,000,000 个容量预算；请缩小稀疏数据的跨度。");
        if (imageBytes > MaxImageBytes) throw new InvalidDataException("图片总大小超过 64 MiB 容量限制。");
    }

    private static StyleDefinition FromStyle(XlStyle style) => new()
    {
        FontName = style.FontName, FontSize = style.FontSize, FontColor = style.FontColor,
        Bold = style.Bold, Italic = style.Italic, FillColor = style.FillColor,
        Horizontal = (int)(style.Alignment?.Horizontal ?? XlHorizontalAlignment.General),
        Vertical = (int)(style.Alignment?.Vertical ?? XlVerticalAlignment.Bottom), WrapText = style.Alignment?.WrapText ?? false,
        LeftBorder = (int)(style.Border?.Left ?? XlBorderStyle.None), RightBorder = (int)(style.Border?.Right ?? XlBorderStyle.None),
        TopBorder = (int)(style.Border?.Top ?? XlBorderStyle.None), BottomBorder = (int)(style.Border?.Bottom ?? XlBorderStyle.None),
        LeftBorderColor = style.Border?.LeftColor, RightBorderColor = style.Border?.RightColor,
        TopBorderColor = style.Border?.TopColor, BottomBorderColor = style.Border?.BottomColor
    };

    private static XlStyle? ToStyle(WorkbookSnapshot snapshot, int id) => id == 0 && snapshot.Styles[0] == new StyleDefinition()
        ? null : ToStyle(snapshot.Styles[id]);

    private static XlStyle ToStyle(StyleDefinition style) => new()
    {
        FontName = style.FontName, FontSize = style.FontSize, FontColor = style.FontColor,
        Bold = style.Bold, Italic = style.Italic, FillColor = style.FillColor,
        Alignment = new XlAlignment { Horizontal = (XlHorizontalAlignment)style.Horizontal, Vertical = (XlVerticalAlignment)style.Vertical, WrapText = style.WrapText },
        Border = new XlBorder
        {
            Left = (XlBorderStyle)style.LeftBorder, Right = (XlBorderStyle)style.RightBorder,
            Top = (XlBorderStyle)style.TopBorder, Bottom = (XlBorderStyle)style.BottomBorder,
            LeftColor = style.LeftBorderColor, RightColor = style.RightBorderColor,
            TopColor = style.TopBorderColor, BottomColor = style.BottomBorderColor
        }
    };

    internal static XlStyle? CloneStyle(XlStyle? style) => style is null ? null : new XlStyle
    {
        FontName = style.FontName, FontSize = style.FontSize, FontColor = style.FontColor,
        Bold = style.Bold, Italic = style.Italic, FillColor = style.FillColor,
        Alignment = style.Alignment is null ? null : new XlAlignment
        { Horizontal = style.Alignment.Horizontal, Vertical = style.Alignment.Vertical, WrapText = style.Alignment.WrapText },
        Border = style.Border is null ? null : new XlBorder
        {
            Left = style.Border.Left, Right = style.Border.Right, Top = style.Border.Top, Bottom = style.Border.Bottom,
            LeftColor = style.Border.LeftColor, RightColor = style.Border.RightColor,
            TopColor = style.Border.TopColor, BottomColor = style.Border.BottomColor
        }
    };
}

