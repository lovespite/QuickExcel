using System.Collections.Immutable;

namespace QuickExcel.Domain;

public readonly record struct SheetId(Guid Value)
{
    public static SheetId New() => new(Guid.NewGuid());
    public override string ToString() => Value.ToString("D");
}

public readonly record struct CellPosition(int Row, int Column);
public readonly record struct CellAddress(SheetId SheetId, int Row, int Column);

/// <summary>A cell's immutable input and optional calculation cache.</summary>
public sealed record CellSnapshot
{
    public CellSnapshot(CellContent content, CellCalculation? calculation = null)
    {
        ArgumentNullException.ThrowIfNull(content);
        if (content.Kind == CellContentKind.Formula) WorkbookLimits.ValidateText(content.FormulaText);
        else if (content.LiteralValue.Kind == CellValueKind.Text) WorkbookLimits.ValidateText(content.LiteralValue.TextValue);
        if (calculation?.Value is { Kind: CellValueKind.Text } value) WorkbookLimits.ValidateText(value.TextValue);
        Content = content;
        Calculation = calculation;
    }

    public CellContent Content { get; }
    public CellCalculation? Calculation { get; }
    public CellValue Value => Content.Kind == CellContentKind.Literal
        ? Content.LiteralValue : Calculation?.Value ?? CellValue.Blank;
}

/// <summary>Value-only style definition. Alignment and border integer values match the storage adapter's supported enums.</summary>
public sealed record StyleDefinition
{
    public int NumberFormatId { get; init; }
    public string? NumberFormatCode { get; init; }
    public string? FontName { get; init; }
    public string? EastAsianFontName { get; init; }
    public string? FontScheme { get; init; }
    public int NamedStyleId { get; init; }
    public double? FontSize { get; init; }
    public string? FontColor { get; init; }
    public bool Bold { get; init; }
    public bool Italic { get; init; }
    public string? FillColor { get; init; }
    public int Horizontal { get; init; }
    public int Vertical { get; init; } = 2;
    public bool WrapText { get; init; }
    public int TextRotation { get; init; }
    public int Indent { get; init; }
    public bool ShrinkToFit { get; init; }
    public int LeftBorder { get; init; }
    public int RightBorder { get; init; }
    public int TopBorder { get; init; }
    public int BottomBorder { get; init; }
    public string? LeftBorderColor { get; init; }
    public string? RightBorderColor { get; init; }
    public string? TopBorderColor { get; init; }
    public string? BottomBorderColor { get; init; }
}

/// <summary>Height is in device-independent pixels (1/96 inch); XLSX stores points (1/72 inch).</summary>
public sealed record RowMetadata
{
    public RowMetadata(double? Height = null, bool Hidden = false, int StyleId = 0, bool? ApplyStyle = null)
    {
        this.Height = Height; this.Hidden = Hidden; this.StyleId = StyleId;
        this.ApplyStyle = ApplyStyle ?? StyleId != 0;
    }
    public double? Height { get; init; }
    public bool Hidden { get; init; }
    public int StyleId { get; init; }
    /// <summary>Independent of the index: an active style 0 overrides column formatting.</summary>
    public bool ApplyStyle { get; init; }
    public void Deconstruct(out double? height, out bool hidden, out int styleId) => (height, hidden, styleId) = (Height, Hidden, StyleId);
}
public sealed record ColumnMetadata(double? Width = null, bool Hidden = false, int StyleId = 0);

public sealed class SheetImage
{
    public SheetImage(IEnumerable<byte> data, string extension, int rowIndex, int columnIndex,
        int rowSpan = 1, int columnSpan = 1, bool placeInCell = false)
    {
        ArgumentNullException.ThrowIfNull(data);
        ArgumentException.ThrowIfNullOrWhiteSpace(extension);
        WorkbookLimits.ValidatePosition(rowIndex, columnIndex);
        if (rowSpan <= 0 || rowSpan > WorkbookLimits.MaxRows - rowIndex) throw new ArgumentOutOfRangeException(nameof(rowSpan));
        if (columnSpan <= 0 || columnSpan > WorkbookLimits.MaxColumns - columnIndex) throw new ArgumentOutOfRangeException(nameof(columnSpan));
        var normalized = extension.Trim().TrimStart('.').ToLowerInvariant();
        if (normalized is not ("png" or "jpg" or "jpeg" or "gif" or "bmp" or "tif" or "tiff"))
            throw new ArgumentException("Unsupported image extension.", nameof(extension));
        Data = data.ToImmutableArray();
        if (Data.IsEmpty) throw new ArgumentException("Image data cannot be empty.", nameof(data));
        Extension = normalized;
        RowIndex = rowIndex;
        ColumnIndex = columnIndex;
        RowSpan = rowSpan;
        ColumnSpan = columnSpan;
        PlaceInCell = placeInCell;
    }

    public ImmutableArray<byte> Data { get; }
    public string Extension { get; }
    public int RowIndex { get; }
    public int ColumnIndex { get; }
    public int RowSpan { get; }
    public int ColumnSpan { get; }
    public bool PlaceInCell { get; }
}

/// <summary>A sparse immutable sheet. Empty positions do not allocate cells or rows.</summary>
public sealed class WorksheetSnapshot
{
    public WorksheetSnapshot(SheetId id, string name,
        IEnumerable<KeyValuePair<CellPosition, CellSnapshot>>? cells = null,
        IEnumerable<KeyValuePair<int, RowMetadata>>? rows = null,
        IEnumerable<KeyValuePair<int, ColumnMetadata>>? columns = null,
        IEnumerable<string>? mergedCells = null, IEnumerable<SheetImage>? images = null,
        bool showGridLines = true, double? defaultRowHeight = null, string? tabColor = null,
        WorksheetView? view = null, SheetVisibility visibility = SheetVisibility.Visible)
    {
        if (id.Value == Guid.Empty) throw new ArgumentException("A sheet requires a nonempty stable identity.", nameof(id));
        WorkbookLimits.ValidateSheetName(name);
        WorkbookLimits.ValidateDimension(defaultRowHeight, nameof(defaultRowHeight));
        Id = id;
        Name = name;
        Cells = cells?.ToImmutableDictionary() ?? ImmutableDictionary<CellPosition, CellSnapshot>.Empty;
        if (Cells.Count > WorkbookLimits.MaxCells) throw new ArgumentException("The sheet exceeds the cell limit.", nameof(cells));
        foreach (var (position, cell) in Cells)
        {
            WorkbookLimits.ValidatePosition(position.Row, position.Column);
            ArgumentNullException.ThrowIfNull(cell);
            UsedRowCount = Math.Max(UsedRowCount, position.Row + 1);
            UsedColumnCount = Math.Max(UsedColumnCount, position.Column + 1);
        }
        Rows = rows?.ToImmutableDictionary() ?? ImmutableDictionary<int, RowMetadata>.Empty;
        foreach (var (row, metadata) in Rows)
        {
            WorkbookLimits.ValidatePosition(row, 0);
            ArgumentNullException.ThrowIfNull(metadata);
            ArgumentOutOfRangeException.ThrowIfNegative(metadata.StyleId);
            WorkbookLimits.ValidateDimension(metadata.Height, nameof(rows));
        }
        Columns = columns?.ToImmutableDictionary() ?? ImmutableDictionary<int, ColumnMetadata>.Empty;
        foreach (var (column, metadata) in Columns)
        {
            WorkbookLimits.ValidatePosition(0, column);
            ArgumentNullException.ThrowIfNull(metadata);
            ArgumentOutOfRangeException.ThrowIfNegative(metadata.StyleId);
            WorkbookLimits.ValidateDimension(metadata.Width, nameof(columns));
        }
        MergedCells = mergedCells?.ToImmutableArray() ?? [];
        if (MergedCells.Any(string.IsNullOrWhiteSpace)) throw new ArgumentException("Merge references cannot be empty.", nameof(mergedCells));
        Images = images?.ToImmutableArray() ?? [];
        if (Images.Any(image => image is null)) throw new ArgumentException("Images cannot contain null entries.", nameof(images));
        ShowGridLines = showGridLines;
        DefaultRowHeight = defaultRowHeight;
        TabColor = tabColor;
        View = view ?? new(); View.Validate();
        if (!Enum.IsDefined(visibility)) throw new ArgumentException("工作表可见性无效。");
        Visibility = visibility;
    }

    public SheetId Id { get; }
    public string Name { get; }
    public ImmutableDictionary<CellPosition, CellSnapshot> Cells { get; }
    public ImmutableDictionary<int, RowMetadata> Rows { get; }
    public ImmutableDictionary<int, ColumnMetadata> Columns { get; }
    public ImmutableArray<string> MergedCells { get; }
    public ImmutableArray<SheetImage> Images { get; }
    public bool ShowGridLines { get; }
    public double? DefaultRowHeight { get; }
    public string? TabColor { get; }
    public WorksheetView View { get; }
    public SheetVisibility Visibility { get; }
    /// <summary>The last stored cell's zero-based row plus one; row metadata does not allocate cells.</summary>
    public int UsedRowCount { get; }
    public int UsedColumnCount { get; }

    public CellSnapshot? GetCell(int row, int column)
    {
        WorkbookLimits.ValidatePosition(row, column);
        return Cells.GetValueOrDefault(new CellPosition(row, column));
    }

    /// <summary>Stored cells have a resolved style, including an explicit General (0) override.</summary>
    public int GetStyleId(int row, int column)
    {
        if (GetCell(row, column) is { } cell) return cell.Content.StyleId;
        return GetInheritedStyleId(row, column);
    }
    public int GetInheritedStyleId(int row, int column)
    {
        WorkbookLimits.ValidatePosition(row, column);
        if (Rows.TryGetValue(row, out var metadata) && metadata.ApplyStyle) return metadata.StyleId;
        return Columns.GetValueOrDefault(column)?.StyleId ?? 0;
    }
    public WorksheetSnapshot WithCells(ImmutableDictionary<CellPosition, CellSnapshot> cells)
    {
        ArgumentNullException.ThrowIfNull(cells);
        return new(Id, Name, cells, Rows, Columns, MergedCells, Images, ShowGridLines, DefaultRowHeight, TabColor, View, Visibility);
    }

    public WorksheetSnapshot WithName(string name) =>
        new(Id, name, Cells, Rows, Columns, MergedCells, Images, ShowGridLines, DefaultRowHeight, TabColor, View, Visibility);
    public WorksheetSnapshot WithView(WorksheetView view) =>
        new(Id, Name, Cells, Rows, Columns, MergedCells, Images, ShowGridLines, DefaultRowHeight, TabColor, view, Visibility);
    public WorksheetSnapshot WithVisibility(SheetVisibility visibility) =>
        new(Id, Name, Cells, Rows, Columns, MergedCells, Images, ShowGridLines, DefaultRowHeight, TabColor, View, visibility);
}

/// <summary>A shareable committed document state, independent of IO libraries and mutable sessions.</summary>
public sealed class WorkbookSnapshot
{
    public WorkbookSnapshot(Guid documentId, long revision, ExcelDateSystem dateSystem,
        IEnumerable<WorksheetSnapshot> sheets, string scriptSource = "", IEnumerable<StyleDefinition>? styles = null, ScriptModuleSet? scriptModules = null, WorkbookFileState? fileState = null, IEnumerable<QuickExcel.Formulas.DefinedName>? definedNames = null,
        SheetId? activeSheetId = null, WorkbookWindow? window = null)
    {
        if (documentId == Guid.Empty) throw new ArgumentException("A document requires a nonempty identity.", nameof(documentId));
        ArgumentOutOfRangeException.ThrowIfNegative(revision);
        if (!Enum.IsDefined(dateSystem)) throw new ArgumentOutOfRangeException(nameof(dateSystem));
        ArgumentNullException.ThrowIfNull(sheets);
        ArgumentNullException.ThrowIfNull(scriptSource);
        Sheets = sheets.ToImmutableArray();
        if (Sheets.Length is < 1 or > WorkbookLimits.MaxSheets)
            throw new ArgumentException($"A workbook requires 1 to {WorkbookLimits.MaxSheets} sheets.", nameof(sheets));
        if (Sheets.Any(sheet => sheet is null)) throw new ArgumentException("Sheets cannot contain null entries.", nameof(sheets));
        if (Sheets.Select(sheet => sheet.Id).Distinct().Count() != Sheets.Length)
            throw new ArgumentException("Sheet identities must be unique.", nameof(sheets));
        if (Sheets.Select(sheet => sheet.Name).Distinct(StringComparer.OrdinalIgnoreCase).Count() != Sheets.Length)
            throw new ArgumentException("Sheet names must be unique ignoring case.", nameof(sheets));
        if (!Sheets.Any(s => s.Visibility == SheetVisibility.Visible)) throw new ArgumentException("至少保留一张可见工作表。");
        ActiveSheetId = activeSheetId ?? Sheets.First(s => s.Visibility == SheetVisibility.Visible).Id;
        if (!Sheets.Any(s => s.Id == ActiveSheetId && s.Visibility == SheetVisibility.Visible)) throw new ArgumentException("活动工作表必须存在且可见。");
        Window = window ?? new();
        if (Sheets.Sum(sheet => (long)sheet.Cells.Count) > WorkbookLimits.MaxCells)
            throw new ArgumentException("The workbook exceeds the cell limit.", nameof(sheets));
        Styles = styles?.ToImmutableArray() ?? [new StyleDefinition()];
        if (Styles.IsEmpty || Styles.Any(style => style is null))
            throw new ArgumentException("A workbook requires a nonempty style table without null entries.", nameof(styles));
        foreach (var style in Styles)
        {
            WorkbookLimits.ValidateDimension(style.FontSize, nameof(styles));
            if (style.Horizontal is < 0 or > 4 || style.Vertical is < 0 or > 3 || style.NamedStyleId is < 0 or > 4095 ||
                style.Indent is < 0 or > 250 || style.TextRotation is < 0 or > 180 && style.TextRotation != 255 ||
                style.FontScheme is not (null or "major" or "minor") ||
                style.LeftBorder is < 0 or > 7 || style.RightBorder is < 0 or > 7 ||
                style.TopBorder is < 0 or > 7 || style.BottomBorder is < 0 or > 7)
                throw new ArgumentException("The style contains an unsupported alignment or border.", nameof(styles));
        }
        foreach (var sheet in Sheets)
        {
            if (sheet.Cells.Values.Any(cell => cell.Content.StyleId >= Styles.Length) ||
                sheet.Rows.Values.Any(row => row.StyleId >= Styles.Length) ||
                sheet.Columns.Values.Any(column => column.StyleId >= Styles.Length))
                throw new ArgumentException("A style reference is outside the workbook style table.", nameof(styles));
        }
        DocumentId = documentId;
        Revision = revision;
        DateSystem = dateSystem;
        ScriptModules = scriptModules ?? ScriptModuleSet.Single(scriptSource);
        if (ScriptModules.PrimarySource != scriptSource) throw new ArgumentException("Script source disagrees with the primary module.", nameof(scriptModules));
        ScriptSource = scriptSource;
        FileState = fileState;
        DefinedNames = definedNames?.ToImmutableArray() ?? [];
        if (DefinedNames.Length > 4096) throw new ArgumentException("名称数量超过 4096。", nameof(definedNames));
        var seenNames = new HashSet<(SheetId?, string)>();
        foreach (var name in DefinedNames)
        {
            if (name is null || name.Formula is null || !name.Formula.StartsWith('=') || string.IsNullOrWhiteSpace(name.Name) || name.Name.Length > 255 ||
                !System.Text.RegularExpressions.Regex.IsMatch(name.Name, @"^[\p{L}_][\p{L}\p{N}_.]*$") ||
                QuickExcel.Formulas.FormulaSyntax.IsCellName(name.Name) ||
                name.Name.Equals("TRUE", StringComparison.OrdinalIgnoreCase) || name.Name.Equals("FALSE", StringComparison.OrdinalIgnoreCase) ||
                name.Scope is { } scope && !Sheets.Any(s => s.Id == scope) ||
                !seenNames.Add((name.Scope, name.Name.ToUpperInvariant())))
                throw new ArgumentException("名称或名称作用域无效。", nameof(definedNames));
            var expression = QuickExcel.Formulas.FormulaSyntax.Parse(name.Formula);
            if (QuickExcel.Formulas.FormulaSyntax.Walk(expression).OfType<QuickExcel.Formulas.ReferenceNode>().Any(r =>
                r.Sheet is null || !r.ColumnAbsolute || !r.WholeColumns && !r.RowAbsolute ||
                r.IsRange && (!r.EndColumnAbsolute || !r.WholeColumns && !r.EndRowAbsolute)))
                throw new ArgumentException("名称中的单元格/区域必须使用带表名的绝对引用。", nameof(definedNames));
        }
    }

    public Guid DocumentId { get; }
    public long Revision { get; }
    public ExcelDateSystem DateSystem { get; }
    public ImmutableArray<WorksheetSnapshot> Sheets { get; }
    public ImmutableArray<StyleDefinition> Styles { get; }
    public string ScriptSource { get; }
    public ScriptModuleSet ScriptModules { get; }
    public WorkbookFileState? FileState { get; }
    public SheetId ActiveSheetId { get; }
    public WorkbookWindow Window { get; }
    public ImmutableArray<QuickExcel.Formulas.DefinedName> DefinedNames { get; }

    public WorksheetSnapshot GetSheet(SheetId id) => Sheets.FirstOrDefault(sheet => sheet.Id == id)
        ?? throw new KeyNotFoundException($"Sheet {id} does not exist.");

    public WorkbookSnapshot WithDefinedNames(IEnumerable<QuickExcel.Formulas.DefinedName> names) =>
        new(DocumentId, Revision, DateSystem, Sheets, ScriptSource, Styles, ScriptModules, FileState, names, ActiveSheetId, Window);

    public WorkbookSnapshot WithSheets(IEnumerable<WorksheetSnapshot> sheets)
    {
        var next = sheets.ToImmutableArray();
        return new(DocumentId, Revision, DateSystem, next, ScriptSource, Styles, ScriptModules, FileState, DefinedNames,
            next.Any(s => s.Id == ActiveSheetId && s.Visibility == SheetVisibility.Visible) ? ActiveSheetId : null, Window);
    }

    public WorkbookSnapshot WithRevision(long revision) =>
        new(DocumentId, revision, DateSystem, Sheets, ScriptSource, Styles, ScriptModules, FileState, DefinedNames, ActiveSheetId, Window);

    public WorkbookSnapshot WithScriptModules(ScriptModuleSet modules) =>
        new(DocumentId, Revision, DateSystem, Sheets, modules.PrimarySource, Styles, modules, FileState, DefinedNames, ActiveSheetId, Window);

    public WorkbookSnapshot WithScriptSource(string scriptSource) =>
        new(DocumentId, Revision, DateSystem, Sheets, scriptSource, Styles, ScriptModules.WithPrimarySource(scriptSource), FileState, DefinedNames, ActiveSheetId, Window);

    public WorkbookSnapshot WithActiveSheet(SheetId id) =>
        new(DocumentId, Revision, DateSystem, Sheets, ScriptSource, Styles, ScriptModules, FileState, DefinedNames, id, Window);
}
