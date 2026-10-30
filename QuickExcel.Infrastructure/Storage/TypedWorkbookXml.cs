using System.Globalization;
using System.Xml.Linq;
using QuickExcel.Domain;
using static QuickExcel.Storage.PackageIO;
using static QuickExcel.Storage.TypedStyles;
using static QuickExcel.Storage.XlsxPackage;

namespace QuickExcel.Storage;

internal sealed record TypedWorkbookRead(WorkbookSnapshot Snapshot, IReadOnlyList<WorkbookPartBinding> Bindings,
    IReadOnlyList<string> Issues, IReadOnlyList<CompatibilityDiagnostic> Diagnostics, bool HasWorkbookView, IReadOnlyList<PreservedFormula> PreservedFormulas);

internal static class TypedWorkbookXml
{
    internal static readonly XNamespace Quick = "urn:quickexcel:cell-error:v2";
    private const string ErrorExtension = "urn:quickexcel:cell-error:v2";
    private const string SheetType = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";

    public static TypedWorkbookRead Read(byte[] bytes, FileReadBudget budget)
    {
        var package = new XlsxPackage(bytes, budget);
        try { return Read(package, budget); }
        catch (Exception exception) when (exception is InvalidDataException or FormatException or OverflowException or ArgumentException or NotSupportedException)
        { package.AttachRejection(exception); throw; }
    }
    private static TypedWorkbookRead Read(XlsxPackage package, FileReadBudget budget)
    {
        var workbookPart = package.RelatedByType("", "officeDocument", true)!;
        package.RequireType(workbookPart, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
        var root = package.XmlPart(workbookPart).Root!;
        if (root.Name != Main + "workbook") throw new InvalidDataException("无效工作簿。");
        var textCompatibility = XlsxTextCompatibility.Read(package, root, workbookPart);
        package.Shape(XlsxTextCompatibility.Shape(root), "workbookPr bookViews sheets definedNames calcPr", "", workbookPart);
        if (root.Element(Main + "workbookPr") is { } pr) package.Shape(pr, "", "date1904", workbookPart);
        if (root.Element(Main + "calcPr") is { } calc)
        {
            package.Shape(calc, "", "calcId fullCalcOnLoad forceFullCalc", workbookPart);
            foreach (var flag in new[] { "fullCalcOnLoad", "forceFullCalc" })
                if (calc.Attribute(flag) is { } value) _ = Bool(value.Value);
            if (calc.Attribute("calcId") is { } id && !uint.TryParse(id.Value, NumberStyles.None, CultureInfo.InvariantCulture, out _))
                throw new InvalidDataException("计算引擎版本编号无效。");
        }
        var dateSystem = Bool((string?)root.Element(Main + "workbookPr")?.Attribute("date1904")) ? ExcelDateSystem.Mac1904 : ExcelDateSystem.Windows1900;
        var styles = TypedStyles.Read(package, workbookPart);
        var textReader = new XlsxCellText(package, workbookPart);
        var richErrors = XlsxRichErrors.Read(package, workbookPart);
        var strings = new List<XlsxTextValue>();
        var stylesPart = package.RelatedByType(workbookPart, "styles");
        var stringsPart = package.RelatedByType(workbookPart, "sharedStrings");
        if (stringsPart is not null)
        {
            package.RequireType(stringsPart, "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml");
            var sst = package.XmlPart(stringsPart).Root!;
            if (sst.Name != Main + "sst") throw new InvalidDataException("无效共享字符串。");
            foreach (var item in sst.Elements(Main + "si"))
            {
                using var sourceRange = package.SourceRange(stringsPart, "sharedStrings[" + strings.Count + "]");
                if (strings.Count >= WorkbookLimits.MaxCells) throw new InvalidDataException("共享字符串过多。");
                strings.Add(textReader.Read(item, stringsPart));
            }
        }
        package.PrepareSharedLocations();
        var sheets = new List<WorksheetSnapshot>();
        var bindings = new List<WorkbookPartBinding>();
        var sheetIds = new HashSet<string>();
        var sheetPaths = new HashSet<string>();
        long cellCount = 0;
        var selectedTabs = 0;
        var preservedFormulas = new List<PreservedFormula>();
        long inspectedArrayCells = 0;
        foreach (var definition in root.Element(Main + "sheets")?.Elements(Main + "sheet") ?? [])
        {
            package.SetReadContext(workbookPart, "sheet:" + (string?)definition.Attribute("name"), "workbook/sheet");
            budget.Check();
            if (sheets.Count >= WorkbookLimits.MaxSheets) throw new InvalidDataException("工作表数量超过限制。");
            foreach (var attribute in definition.Attributes().Where(a => !a.IsNamespaceDeclaration && a.Name != Rel + "id" && a.Name.LocalName is not ("name" or "sheetId" or "state")))
                package.Loss($"{workbookPart}: 工作表属性 {attribute.Name} 未建模", "sheet/@" + attribute.Name.LocalName, workbookPart,
                    sheetName: Decode(Attr(definition, "name")));
            var sheetId = Attr(definition, "sheetId");
            if (Int(sheetId, int.MaxValue) == 0 || !sheetIds.Add(Int(sheetId, int.MaxValue).ToString(CultureInfo.InvariantCulture))) throw new InvalidDataException("工作表编号无效或重复。");
            var path = package.Related(workbookPart, Attr(definition, Rel + "id"), "worksheet");
            if (!sheetPaths.Add(path)) throw new InvalidDataException("多张工作表指向同一部件。");
            package.RequireType(path, SheetType);
            package.RegisterSheet(path, Decode(Attr(definition, "name")));
            var visibility = XlsxViews.Visibility(definition);
            var sheet = package.XmlPart(path).Root!;
            if (sheet.Name != Main + "worksheet") throw new InvalidDataException("无效工作表部件。");
            package.Shape(sheet, "dimension sheetPr sheetViews sheetFormatPr cols sheetData mergeCells hyperlinks printOptions pageMargins pageSetup headerFooter rowBreaks colBreaks drawing legacyDrawing", "", path);
            var sharedFormulas = new SharedFormulaExpander(sheet);
            var cells = new Dictionary<CellPosition, CellSnapshot>();
            var cachedCells = new HashSet<CellPosition>();
            var firstPreserved = preservedFormulas.Count;
            var rows = new Dictionary<int, RowMetadata>();
            var columns = new Dictionary<int, ColumnMetadata>();
            var seenRows = new HashSet<int>();
            foreach (var column in sheet.Element(Main + "cols")?.Elements(Main + "col") ?? [])
            {
                package.SetReadContext(path, "columns:" + (string?)column.Attribute("min") + ":" + (string?)column.Attribute("max"), "worksheet/column");
                package.Shape(column, "", "min max width hidden style customWidth", path);
                var min = Int(Attr(column, "min"), WorkbookLimits.MaxColumns);
                var max = Int(Attr(column, "max"), WorkbookLimits.MaxColumns);
                if (min == 0 || max < min) throw new InvalidDataException("列范围无效。");
                var metadata = new ColumnMetadata(OptionalDouble(column.Attribute("width")), Bool((string?)column.Attribute("hidden")), Style(column, "style", styles.Count));
                if (column.Attribute("style") is not null) package.RegisterSharedUsage(stylesPart, "cellXfs[" + metadata.StyleId + "]", path, $"columns:{min}:{max}");
                for (var col = min - 1; col < max; col++)
                {
                    budget.Parse(96);
                    if (!columns.TryAdd(col, metadata)) throw new InvalidDataException("列定义重叠。");
                }
            }
            foreach (var row in sheet.Element(Main + "sheetData")?.Elements(Main + "row") ?? [])
            {
                package.SetReadContext(path, "row:" + (string?)row.Attribute("r"), "worksheet/row");
                budget.Parse(128);
                var rowIndex = Int(Attr(row, "r"), WorkbookLimits.MaxRows) - 1;
                if (rowIndex < 0 || !seenRows.Add(rowIndex)) throw new InvalidDataException("工作表行重复或无效。");
                package.Shape(row, "c", "r ht hidden s customHeight customFormat spans", path);
                var rowMetadata = new RowMetadata(OptionalDouble(row.Attribute("ht")) / 0.75, Bool((string?)row.Attribute("hidden")),
                    Style(row, "s", styles.Count), Bool((string?)row.Attribute("customFormat")));
                if (rowMetadata != new RowMetadata() || row.Attribute("s") is not null || row.Attribute("customFormat") is not null) rows.Add(rowIndex, rowMetadata);
                if (rowMetadata.ApplyStyle) package.RegisterSharedUsage(stylesPart, "cellXfs[" + rowMetadata.StyleId + "]", path, "row:" + (rowIndex + 1));
                var nextColumn = 0;
                foreach (var element in row.Elements(Main + "c"))
                {
                    package.SetReadContext(path, (string?)element.Attribute("r") ?? Reference(new(rowIndex, nextColumn)), "worksheet/cell");
                    budget.Parse(320);
                    if (++cellCount > WorkbookLimits.MaxCells) throw new InvalidDataException("单元格累计数量超过限制。");
                    var position = element.Attribute("r") is { } address ? Position(address.Value) : new CellPosition(rowIndex, nextColumn);
                    if (position.Row != rowIndex || position.Column >= WorkbookLimits.MaxColumns) throw new InvalidDataException("单元格地址与所在行不一致。");
                    nextColumn = position.Column + 1;
                    package.Shape(element, "f v is extLst", "r s t vm", path);
                    var style = element.Attribute("s") is not null ? Style(element, "s", styles.Count)
                        : rowMetadata.ApplyStyle ? rowMetadata.StyleId : columns.GetValueOrDefault(position.Column)?.StyleId ?? 0;
                    var value = Value(element, strings, package, path, textReader, out var richText);
                    value = XlsxRichErrors.Apply(element, value, richErrors, package, path);
                    if (package.HasSharedLocations)
                    {
                        var cellReference = Reference(position);
                        package.RegisterSharedUsage(stylesPart, "cellXfs[" + style + "]", path, cellReference);
                        if ((string?)element.Attribute("t") == "s")
                            package.RegisterSharedUsage(stringsPart, "sharedStrings[" + Int(element.Element(Main + "v")!.Value, strings.Count - 1) + "]", path, cellReference);
                    }
                    if (element.Element(Main + "v") is not null) cachedCells.Add(position);
                    CellSnapshot cell;
                    if (element.Element(Main + "f") is { } formula)
                    {
                        package.Shape(formula, "", "t si ref", path);
                        string expression;
                        var preserved = false;
                        try
                        {
                            if ((string?)formula.Attribute("t") == "array")
                            {
                                var range = Attr(formula, "ref"); ValidateRange(range);
                                if (!CellRectangle.Parse(range).Contains(position.Row, position.Column)) throw new InvalidDataException("数组公式主格不在结果区域内。");
                                expression = formula.Value;
                                preservedFormulas.Add(new(element.Attribute("cm") is null ? PreservedFormulaKind.Array : PreservedFormulaKind.DynamicArray,
                                    path, Decode(Attr(definition, "name")), Reference(position), range, expression));
                                preserved = true;
                            }
                            else expression = sharedFormulas.Expand(formula, position);
                        }
                        catch (NotSupportedException ex)
                        {
                            if ((string?)formula.Attribute("t") == "shared")
                            {
                                expression = sharedFormulas.Source(formula, position);
                                preservedFormulas.Add(new(PreservedFormulaKind.Shared, path, Decode(Attr(definition, "name")), Reference(position), (string?)formula.Attribute("ref"), expression));
                                preserved = true;
                            }
                            else
                            throw new WorkbookFeatureRejectedException(new("xlsx.formula-rejected", CompatibilityKind.Rejected,
                                "formula/" + ((string?)formula.Attribute("t") ?? "normal"), ex.Message, path,
                                Decode(Attr(definition, "name")), (string?)formula.Attribute("ref") ?? Reference(position)), ex);
                        }
                        if (expression.Length == 0) throw new InvalidDataException("公式内容缺失。");
                        if (!preserved)
                        {
                            try { _ = QuickExcel.Formulas.FormulaSyntax.Parse("=" + expression.TrimStart('=')); }
                            catch (Exception ex) when (ex is FormatException or OverflowException)
                            {
                                var external = System.Text.RegularExpressions.Regex.IsMatch(expression, @"(?:'[^']*\[[^\]]+\][^']*'|\[[^\]]+\][^!]+)!");
                                var kind = external ? PreservedFormulaKind.ExternalReference : expression.Contains('[') ? PreservedFormulaKind.StructuredReference : PreservedFormulaKind.UnsupportedSyntax;
                                preservedFormulas.Add(new(kind, path, Decode(Attr(definition, "name")), Reference(position), null, expression));
                            }
                        }
                        cell = new(CellContent.FromFormula("=" + expression.TrimStart('='), style),
                            new CellCalculation(element.Element(Main + "v") is null ? null : value, CalculationState.Stale, 0, 0, 0));
                    }
                    else cell = new(richText is null ? CellContent.FromLiteral(value, style) : CellContent.FromRichText(richText, style));
                    if (!cells.TryAdd(position, cell)) throw new InvalidDataException("单元格地址重复。");
                }
            }
            package.SetReadContext(path, null, "worksheet/layout");
            var view = sheet.Element(Main + "sheetViews")?.Element(Main + "sheetView");
            for (var p = firstPreserved; p < preservedFormulas.Count; p++)
            {
                var info = preservedFormulas[p];
                if (info.Kind is not (PreservedFormulaKind.Array or PreservedFormulaKind.DynamicArray) || info.Range is null) continue;
                var range = CellRectangle.Parse(info.Range); var complete = true;
                for (var row = range.Top; row <= range.Bottom && complete; row++)
                for (var column = range.Left; column <= range.Right; column++)
                {
                    if ((inspectedArrayCells & 1023) == 0) budget.Check();
                    if (++inspectedArrayCells > WorkbookLimits.MaxCells || !cachedCells.Contains(new(row, column))) { complete = false; break; }
                }
                preservedFormulas[p] = info with { HasCompleteCache = complete };
            }
            var viewState = XlsxViews.ReadSheet(package, path, sheet);
            if (Bool((string?)view?.Attribute("tabSelected"))) selectedTabs++;
            var format = sheet.Element(Main + "sheetFormatPr");
            if (format is not null) package.Shape(format, "", "defaultRowHeight", path);
            var properties = sheet.Element(Main + "sheetPr");
            if (properties is not null) package.Shape(properties, "tabColor pageSetUpPr", "", path);
            var tabColor = properties?.Element(Main + "tabColor");
            if (tabColor is not null) package.Shape(tabColor, "", "rgb", path);
            var merges = sheet.Element(Main + "mergeCells")?.Elements(Main + "mergeCell").Select(e => Attr(e, "ref")).ToArray() ?? [];
            foreach (var merge in merges) ValidateRange(merge);
            var images = TypedImages.Read(package, path, sheet);
            sheets.Add(new WorksheetSnapshot(SheetId.New(), Decode(Attr(definition, "name")), cells, rows, columns, merges, images,
                view?.Attribute("showGridLines") is null || Bool((string?)view.Attribute("showGridLines")), OptionalDouble(format?.Attribute("defaultRowHeight")) / 0.75, (string?)tabColor?.Attribute("rgb"),
                viewState, visibility, XlsxPrinting.Read(package, path, sheet), XlsxAnnotations.Read(package, path, sheet, textReader)));
            bindings.Add(new(sheetId, path));
        }
        package.SetReadContext(workbookPart, null, "workbook/model");
        if (sheets.Count == 0) throw new InvalidDataException("工作簿没有工作表。");
        if (selectedTabs > 1) package.Loss(workbookPart + ": 成组选择工作表尚不支持编辑。", "sheetViews/grouped-tabs", workbookPart);
        var window = XlsxViews.ReadWorkbook(package, workbookPart, root, sheets);
        var names = new List<QuickExcel.Formulas.DefinedName>();
        var printNames = new HashSet<(int, string)>();
        var rejectedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var inspectedNames = 0;
        foreach (var element in root.Element(Main + "definedNames")?.Elements() ?? [])
        {
            package.SetReadContext(workbookPart, (string?)element.Attribute("name"), "definedName");
            budget.Check();
            if (++inspectedNames > 4096) throw new InvalidDataException("名称数量超过 4096。");
            package.Shape(element, "", "name localSheetId", workbookPart);
            var name = Attr(element, "name");
            var scopeIndex = element.Attribute("localSheetId") is { } local ? Int(local.Value, sheets.Count - 1) : -1;
            if (name is "_xlnm.Print_Area" or "_xlnm.Print_Titles")
            {
                if (!printNames.Add((scopeIndex, name))) throw new InvalidDataException("重复的打印名称。");
                try
                {
                    if (scopeIndex < 0) throw new NotSupportedException("打印名称必须为工作表局部名称。");
                    var target = sheets[scopeIndex];
                    sheets[scopeIndex] = target.WithPrinting(XlsxPrinting.ReadName(target, Decode(element.Value), name == "_xlnm.Print_Titles"));
                }
                catch (Exception e) when (e is NotSupportedException or ArgumentException or FormatException or OverflowException)
                { package.Loss($"{workbookPart}: 打印名称 {name} 未支持：{e.Message}", "definedName/" + name, workbookPart, range: element.Value,
                    sheetName: scopeIndex < 0 ? null : sheets[scopeIndex].Name); }
                continue;
            }
            var definition = new QuickExcel.Formulas.DefinedName(name, "=" + element.Value.TrimStart('='), scopeIndex < 0 ? null : sheets[scopeIndex].Id);
            try
            {
                _ = new WorkbookSnapshot(Guid.NewGuid(), 0, dateSystem, sheets, styles: styles, definedNames: [definition]);
                if (name.StartsWith("_xlnm.", StringComparison.OrdinalIgnoreCase)) throw new ArgumentException("内置名称尚未建模。");
                names.Add(definition);
            }
            catch (Exception e) when (e is ArgumentException or FormatException or OverflowException)
            {
                package.Loss($"{workbookPart}: 名称 {name} 未支持：{e.Message}", "definedName/" + name, workbookPart,
                    range: element.Value, sheetName: scopeIndex < 0 ? null : sheets[scopeIndex].Name);
                rejectedNames.Add(name);
                if (!name.StartsWith("_xlnm.", StringComparison.OrdinalIgnoreCase))
                    preservedFormulas.Add(new(PreservedFormulaKind.DefinedName, workbookPart,
                        scopeIndex < 0 ? "(工作簿)" : sheets[scopeIndex].Name, name, null, element.Value));
            }
        }
        package.SetReadContext(workbookPart, null, "workbook/model");
        var snapshot = new WorkbookSnapshot(Guid.NewGuid(), 0, dateSystem, sheets, styles: styles, definedNames: names, activeSheetId: window.Active, window: window.Window, textCompatibilityVersion: textCompatibility);
        var preservedLocations = preservedFormulas.Select(f => (f.Part, f.Cell)).ToHashSet();
        // Names can hide unsupported syntax or circular/oversized binding behind an otherwise simple formula.
        foreach (var sheet in sheets)
        foreach (var (position, cell) in sheet.Cells)
        {
            budget.Check();
            if (cell.Content.Kind != CellContentKind.Formula) continue;
            var part = bindings[sheets.IndexOf(sheet)].Part;
            var address = Reference(position);
            if (preservedLocations.Contains((part, address))) continue;
            try
            {
                var parsed = QuickExcel.Formulas.FormulaSyntax.Parse(cell.Content.FormulaText);
                if (QuickExcel.Formulas.FormulaSyntax.Walk(parsed).OfType<QuickExcel.Formulas.NameNode>().Any(n => rejectedNames.Contains(n.Name)))
                    throw new FormatException("公式引用了未建模的名称。");
                _ = QuickExcel.Formulas.FormulaBinding.Bind(parsed, snapshot, sheet.Id);
            }
            catch (Exception e) when (e is FormatException or OverflowException)
            {
                preservedFormulas.Add(new(PreservedFormulaKind.UnsupportedSyntax, part, sheet.Name, address, null, cell.Content.FormulaText));
            }
        }
        XlsxMetadata.Inspect(package, sheets.Select(s => s.Name).ToArray());
        package.FinishInventory();
        return new(snapshot, bindings,
            package.Issues.ToArray(), package.Diagnostics, root.Element(Main + "bookViews") is not null, preservedFormulas);
    }

    private static int Style(XElement element, string name, int count) => Int((string?)element.Attribute(name) ?? "0", count - 1);
    private static CellValue Value(XElement element, List<XlsxTextValue> strings, XlsxPackage package, string path, XlsxCellText textReader, out RichCellText? richText)
    {
        foreach (var name in new[] { "f", "v", "is", "extLst" })
            if (element.Elements(Main + name).Count() > 1) throw new InvalidDataException("单元格内容存在重复定义。");
        richText = null;
        var type = (string?)element.Attribute("t");
        if (type == "inlineStr" && (element.Element(Main + "v") is not null || element.Element(Main + "f") is not null) ||
            type != "inlineStr" && element.Element(Main + "is") is not null) throw new InvalidDataException("单元格类型与内容冲突。");
        var value = element.Element(Main + "v");
        CellValue result;
        switch (type)
        {
            case "inlineStr":
                if (element.Element(Main + "is") is not { } inline) throw new InvalidDataException("内联文本内容缺失。");
                var parsed = textReader.Read(inline, path); richText = parsed.Rich; result = CellValue.FromText(parsed.Text); break;
            case "str": result = value is null ? CellValue.Blank : CellValue.FromText(Decode(value.Value)); break;
            case "s": var shared = strings[Int(value?.Value ?? "", strings.Count - 1)]; richText = shared.Rich; result = CellValue.FromText(shared.Text); break;
            case "b": result = value is null ? CellValue.Blank : CellValue.FromBoolean(Bool(value.Value)); break;
            case "e":
                if (value is null || !CellInput.TryFromErrorText(value.Value, out result) || result.ErrorCode == CellErrorCode.Calculation)
                    throw new InvalidDataException("错误单元格代码无效；#CALC! 必须使用富值元数据和 #VALUE! 回退值。");
                break;
            case null or "n":
                if (value is null || value.Value.Length == 0) result = CellValue.Blank;
                else if (double.TryParse(value.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) && double.IsFinite(number)) result = CellValue.FromNumber(number);
                else throw new InvalidDataException("数值单元格不是有限数值。");
                break;
            default: throw new NotSupportedException($"不支持的单元格类型：{type}");
        }
        foreach (var ext in element.Element(Main + "extLst")?.Elements() ?? [])
        {
            if ((string?)ext.Attribute("uri") == ErrorExtension && ext.Element(Quick + "error") is { } error)
            {
                if (type != "e" || value?.Value != "#VALUE!" || !Enum.TryParse<CellErrorCode>((string?)error.Attribute("code"), out var code) || code is not (CellErrorCode.Script or CellErrorCode.Cycle))
                    throw new InvalidDataException("宿主错误扩展无效。");
                result = CellValue.FromError(code);
            }
            else package.Loss($"{path}: 单元格扩展未建模", "cell/extLst", path, element);
        }
        return result;
    }

    public static byte[] Write(WorkbookSnapshot snapshot, bool valuesOnly, CancellationToken cancellationToken)
    {
        if (!valuesOnly && snapshot.FileState?.FormulaReadOnly == true) throw new InvalidOperationException(WorkbookFileState.FormulaReadOnlyMessage);
        long allocation = snapshot.Styles.Length * 2048L, imageBytes = 0;
        foreach (var sheet in snapshot.Sheets)
        {
            cancellationToken.ThrowIfCancellationRequested();
            allocation += sheet.Cells.Count * 640L + (sheet.Rows.Count + sheet.Columns.Count) * 256L;
            foreach (var cell in sheet.Cells.Values)
            {
                if (cell.Content.Kind == CellContentKind.Formula) allocation += cell.Content.FormulaText.Length * 4L;
                if (cell.Value.Kind == CellValueKind.Text) allocation += cell.Value.TextValue.Length * 8L;
                if (cell.Content.RichText is { } rich) allocation += RichAllocation(rich);
            }
            foreach (var link in sheet.Annotations.Links.Values)
                allocation += 512L + 8L * (link.Target?.Length ?? 0) + 8L * (link.Location?.Length ?? 0) + 8L * (link.Display?.Length ?? 0) + 8L * (link.Tooltip?.Length ?? 0);
            foreach (var note in sheet.Annotations.Notes.Values) allocation += 512L + note.Author.Length * 8L + note.Text.Text.Length * 8L + RichAllocation(note.Text);
            imageBytes += sheet.Images.Sum(image => (long)image.Data.Length);
            imageBytes += sheet.Printing.PrinterSettings.Length;
            imageBytes += sheet.Annotations.NoteDrawing.Length;
        }
        if (allocation > FileReadBudget.MaxExpandedBytes || imageBytes > FileReadBudget.MaxFileBytes)
            throw new InvalidDataException("写出 XML/图片预算超限，请缩小工作簿。");
        var parts = new Dictionary<string, byte[]>();
        var workbook = new XElement(Main + "workbook", new XAttribute(XNamespace.Xmlns + "r", Rel),
            new XElement(Main + "workbookPr", new XAttribute("date1904", snapshot.DateSystem == ExcelDateSystem.Mac1904 ? 1 : 0)));
        workbook.Add(XlsxViews.WriteWorkbook(snapshot));
        var sheets = new XElement(Main + "sheets");
        var rels = new XElement(PackageRel + "Relationships");
        var types = new XElement(ContentTypes + "Types",
            new XElement(ContentTypes + "Default", new XAttribute("Extension", "rels"), new XAttribute("ContentType", "application/vnd.openxmlformats-package.relationships+xml")),
            new XElement(ContentTypes + "Default", new XAttribute("Extension", "xml"), new XAttribute("ContentType", "application/xml")));
        void Override(string path, string type) => types.Add(new XElement(ContentTypes + "Override", new XAttribute("PartName", "/" + path), new XAttribute("ContentType", type)));
        Override("xl/workbook.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
        Override("xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
        rels.Add(Relationship("styles", "styles", "styles.xml"));
        XlsxRichErrors.Write(snapshot, parts, rels, types);
        var rootRels = new XElement(PackageRel + "Relationships", Relationship("workbook", "officeDocument", "xl/workbook.xml"));
        XlsxMetadata.Write(snapshot, parts, rootRels, types, cancellationToken);
        parts["_rels/.rels"] = XmlBytes(new XDocument(rootRels));
        var stylePackage = XlsxStylePreservation.Write(snapshot, cancellationToken);
        parts["xl/styles.xml"] = XmlBytes(stylePackage.Styles);
        if (stylePackage.Theme is { } theme)
        {
            parts["xl/theme/theme1.xml"] = theme;
            Override("xl/theme/theme1.xml", "application/vnd.openxmlformats-officedocument.theme+xml");
            rels.Add(Relationship("theme", "theme", "theme/theme1.xml"));
        }
        for (var i = 0; i < snapshot.Sheets.Length; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var sheet = snapshot.Sheets[i];
            var path = $"xl/worksheets/sheet{i + 1}.xml";
            sheets.Add(new XElement(Main + "sheet", new XAttribute("name", Encode(sheet.Name)), new XAttribute("sheetId", i + 1), new XAttribute(Rel + "id", $"sheet{i + 1}"),
                sheet.Visibility == SheetVisibility.Visible ? null : new XAttribute("state", sheet.Visibility == SheetVisibility.Hidden ? "hidden" : "veryHidden")));
            rels.Add(Relationship($"sheet{i + 1}", "worksheet", $"worksheets/sheet{i + 1}.xml"));
            Override(path, SheetType);
            var root = new XElement(Main + "worksheet", new XAttribute(XNamespace.Xmlns + "r", Rel));
            if (sheet.TabColor is not null || sheet.Printing.Properties is not null)
                root.Add(new XElement(Main + "sheetPr", sheet.TabColor is not null ? new XElement(Main + "tabColor", new XAttribute("rgb", sheet.TabColor)) : null,
                    XlsxPrinting.WriteProperties(sheet.Printing)));
            root.Add(XlsxViews.WriteSheet(sheet, sheet.Id == snapshot.ActiveSheetId));
            if (sheet.DefaultRowHeight is { } height) root.Add(new XElement(Main + "sheetFormatPr", new XAttribute("defaultRowHeight", height * 0.75)));
            if (sheet.Columns.Count > 0)
                root.Add(new XElement(Main + "cols", sheet.Columns.OrderBy(p => p.Key).Select(p => new XElement(Main + "col",
                    new XAttribute("min", p.Key + 1), new XAttribute("max", p.Key + 1), p.Value.Width is { } width ? new XAttribute("width", width) : null,
                    p.Value.Width.HasValue ? new XAttribute("customWidth", 1) : null, new XAttribute("hidden", p.Value.Hidden ? 1 : 0), new XAttribute("style", p.Value.StyleId)))));
            var data = new XElement(Main + "sheetData");
            var grouped = sheet.Cells.GroupBy(c => c.Key.Row).ToDictionary(g => g.Key, g => g.OrderBy(c => c.Key.Column));
            foreach (var rowIndex in grouped.Keys.Union(sheet.Rows.Keys).Order())
            {
                cancellationToken.ThrowIfCancellationRequested();
                var row = new XElement(Main + "row", new XAttribute("r", rowIndex + 1));
                if (sheet.Rows.TryGetValue(rowIndex, out var metadata))
                {
                    if (metadata.Height.HasValue) row.Add(new XAttribute("ht", metadata.Height.Value * 0.75), new XAttribute("customHeight", 1));
                    row.Add(new XAttribute("hidden", metadata.Hidden ? 1 : 0), new XAttribute("s", metadata.StyleId),
                        new XAttribute("customFormat", metadata.ApplyStyle ? 1 : 0));
                }
                if (grouped.TryGetValue(rowIndex, out var rowCells))
                    foreach (var (position, cell) in rowCells) row.Add(WriteCell(position, cell, valuesOnly));
                data.Add(row);
            }
            root.Add(data);
            if (sheet.MergedCells.Length > 0)
            {
                foreach (var merge in sheet.MergedCells) ValidateRange(merge);
                root.Add(new XElement(Main + "mergeCells", new XAttribute("count", sheet.MergedCells.Length), sheet.MergedCells.Select(m => new XElement(Main + "mergeCell", new XAttribute("ref", m)))));
            }
            XlsxAnnotations.WriteLinks(sheet.Annotations, root);
            XlsxPrinting.WriteSheet(sheet.Printing, root);
            TypedImages.Write(sheet, i + 1, parts, root, types);
            XlsxPrinting.WritePrinter(sheet.Printing, i + 1, parts, types);
            XlsxAnnotations.Write(sheet.Annotations, i + 1, parts, root, types);
            parts[path] = XmlBytes(new XDocument(root));
        }
        workbook.Add(sheets);
        if (!snapshot.DefinedNames.IsEmpty || snapshot.Sheets.Any(s => s.Printing.HasReferences))
        {
            var names = new XElement(Main + "definedNames");
            foreach (var name in snapshot.DefinedNames)
            {
                var element = new XElement(Main + "definedName", new XAttribute("name", name.Name), QuickExcel.Formulas.ExcelFormulaNames.Encode(name.Formula).TrimStart('='));
                if (name.Scope is { } scope) element.SetAttributeValue("localSheetId", snapshot.Sheets.IndexOf(snapshot.GetSheet(scope)));
                names.Add(element);
            }
            for (var i = 0; i < snapshot.Sheets.Length; i++) names.Add(XlsxPrinting.WriteNames(snapshot.Sheets[i], i));
            // Excel resolves same-spelling local/global names incorrectly when appended copies split
            // their groups. Match its serialization: name groups, local scopes before the global name.
            names.ReplaceNodes(names.Elements().OrderBy(n => (string)n.Attribute("name")!, StringComparer.OrdinalIgnoreCase)
                .ThenByDescending(n => (int?)n.Attribute("localSheetId") ?? -1).ToArray());
            workbook.Add(names);
        }
        workbook.Add(new XElement(Main + "calcPr", new XAttribute("fullCalcOnLoad", 1)));
        workbook.Add(XlsxTextCompatibility.Write(snapshot));
        parts["xl/workbook.xml"] = XmlBytes(new XDocument(workbook));
        parts["xl/_rels/workbook.xml.rels"] = XmlBytes(new XDocument(rels));
        parts["[Content_Types].xml"] = XmlBytes(new XDocument(types));
        return Zip(parts);
    }

    private static long RichAllocation(RichCellText rich) => rich.Runs.Sum(run => 512L +
        ((run.Font?.Name?.Length ?? 0) + (run.Font?.EastAsianName?.Length ?? 0)) * 8L);

    private static XElement WriteCell(CellPosition position, CellSnapshot cell, bool valuesOnly)
    {
        var result = new XElement(Main + "c", new XAttribute("r", Reference(position)), new XAttribute("s", cell.Content.StyleId));
        var formula = !valuesOnly && cell.Content.Kind == CellContentKind.Formula;
        if (formula) result.Add(new XElement(Main + "f", QuickExcel.Formulas.ExcelFormulaNames.Encode(cell.Content.FormulaText)[1..]));
        if (formula && cell.Calculation?.Value is null) return result;
        var value = cell.Value;
        switch (value.Kind)
        {
            case CellValueKind.Blank: if (formula) result.Add(new XElement(Main + "v")); break;
            case CellValueKind.Text:
                result.SetAttributeValue("t", formula ? "str" : "inlineStr");
                result.Add(formula ? new XElement(Main + "v", Encode(value.TextValue)) : cell.Content.RichText is { } rich ? XlsxCellText.Write(Main + "is", rich) : new XElement(Main + "is", XlsxCellText.Text(value.TextValue))); break;
            case CellValueKind.Number: result.Add(new XElement(Main + "v", value.NumberValue.ToString("R", CultureInfo.InvariantCulture))); break;
            case CellValueKind.Boolean: result.SetAttributeValue("t", "b"); result.Add(new XElement(Main + "v", value.BooleanValue ? "1" : "0")); break;
            case CellValueKind.Error:
                result.SetAttributeValue("t", "e");
                var hostError = value.ErrorCode is CellErrorCode.Script or CellErrorCode.Cycle;
                var richError = value.ErrorCode == CellErrorCode.Calculation;
                result.Add(new XElement(Main + "v", hostError || richError ? "#VALUE!" : CellInput.DisplayText(value)));
                if (richError) result.SetAttributeValue("vm", "1");
                if (hostError && !valuesOnly) result.Add(new XElement(Main + "extLst", new XElement(Main + "ext", new XAttribute("uri", ErrorExtension), new XElement(Quick + "error", new XAttribute("code", value.ErrorCode)))));
                break;
        }
        return result;
    }
    public static XElement Relationship(string id, string type, string target) => new(PackageRel + "Relationship",
        new XAttribute("Id", id), new XAttribute("Type", Rel.NamespaceName + "/" + type), new XAttribute("Target", target));
    public static CellPosition Position(string reference)
    {
        var index = 0; var column = 0;
        while (index < reference.Length && reference[index] is >= 'A' and <= 'Z')
        {
            column = checked(column * 26 + reference[index++] - 'A' + 1);
            if (column > WorkbookLimits.MaxColumns) throw new InvalidDataException("列地址超过限制。");
        }
        var row = Int(reference[index..], WorkbookLimits.MaxRows);
        if (column == 0 || row == 0) throw new InvalidDataException("无效单元格地址。");
        return new(row - 1, column - 1);
    }
    public static string Reference(CellPosition position)
    {
        var column = position.Column + 1; var text = "";
        while (column > 0) { text = (char)('A' + (column - 1) % 26) + text; column = (column - 1) / 26; }
        return text + (position.Row + 1).ToString(CultureInfo.InvariantCulture);
    }
    private static void ValidateRange(string range)
    {
        var pieces = range.Split(':');
        if (pieces.Length is < 1 or > 2) throw new InvalidDataException("合并范围无效。");
        var start = Position(pieces[0]); var end = Position(pieces[^1]);
        if (end.Row < start.Row || end.Column < start.Column) throw new InvalidDataException("合并范围倒置。");
    }
}
