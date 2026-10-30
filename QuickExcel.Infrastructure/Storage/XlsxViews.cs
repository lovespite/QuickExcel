using System.Globalization;
using System.Xml.Linq;
using QuickExcel.Domain;
using static QuickExcel.Storage.PackageIO;
using static QuickExcel.Storage.TypedStyles;
using static QuickExcel.Storage.TypedWorkbookXml;

namespace QuickExcel.Storage;

internal static class XlsxViews
{
    public static SheetVisibility Visibility(XElement definition) => (string?)definition.Attribute("state") switch
    { null or "visible" => SheetVisibility.Visible, "hidden" => SheetVisibility.Hidden, "veryHidden" => SheetVisibility.VeryHidden, _ => throw new InvalidDataException("工作表可见性无效。") };

    public static WorksheetView ReadSheet(XlsxPackage package, string path, XElement sheet)
    {
        if (sheet.Element(Main + "sheetViews") is not { } views) return new();
        package.Shape(views, "sheetView", "", path);
        if (views.Elements(Main + "sheetView").Count() > 1) package.Loss(path + ": 多窗口工作表视图尚不支持编辑。", "sheetViews/multiple", path);
        if (views.Element(Main + "sheetView") is not { } view) return new();
        package.Shape(view, "pane selection", "workbookViewId showGridLines tabSelected zoomScale zoomScaleNormal topLeftCell view", path);
        if ((string?)view.Attribute("workbookViewId") is not (null or "0")) package.Loss(path + ": 非主窗口视图未建模。", "sheetView/workbookViewId", path);
        if ((string?)view.Attribute("view") is not (null or "normal")) package.Loss(path + ": 页面布局/分页预览未映射到网格。", "sheetView/view", path);
        if (view.Attribute("tabSelected") is { } selected) _ = Bool(selected.Value);
        try
        {
            var result = new WorksheetView { Zoom = Zoom(view, "zoomScale") ?? 100, NormalZoom = Zoom(view, "zoomScaleNormal"), TopLeftCell = Cell((string?)view.Attribute("topLeftCell")) };
            if (view.Elements(Main + "pane").Count() > 1) throw new InvalidDataException("重复的窗格定义。");
            if (view.Element(Main + "pane") is { } pane)
            {
                package.Shape(pane, "", "xSplit ySplit topLeftCell activePane state", path);
                if ((string?)pane.Attribute("state") is not ("frozen" or "frozenSplit")) throw new NotSupportedException("非冻结的拆分窗格尚未支持。");
                var rows = Split(pane, "ySplit", WorkbookLimits.MaxRows);
                var columns = Split(pane, "xSplit", WorkbookLimits.MaxColumns);
                result = result with { FrozenRows = rows, FrozenColumns = columns, FrozenSplit = (string?)pane.Attribute("state") == "frozenSplit",
                    PaneTopLeftCell = Cell((string?)pane.Attribute("topLeftCell") ?? Reference(new(rows, columns))), ActivePane = Pane((string?)pane.Attribute("activePane")) };
            }
            var seen = new HashSet<SheetPane>();
            foreach (var selection in view.Elements(Main + "selection"))
            {
                package.Shape(selection, "", "pane activeCell sqref activeCellId", path);
                var selectionPane = Pane((string?)selection.Attribute("pane"));
                if (!seen.Add(selectionPane)) throw new InvalidDataException("同一窗格有重复选区。");
                if ((string?)selection.Attribute("activeCellId") is not (null or "0")) throw new NotSupportedException("非连续选区的活动范围索引尚未支持。");
                var cell = Cell((string?)selection.Attribute("activeCell"));
                var reference = (string?)selection.Attribute("sqref") ?? Reference(cell);
                var ranges = reference.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
                if (ranges.Length != 1) throw new NotSupportedException("非连续视图选区尚未支持。");
                var endpoints = ranges[0].Split(':');
                if (endpoints.Length is < 1 or > 2) throw new InvalidDataException("视图选区无效。");
                var first = Cell(endpoints[0]); var last = Cell(endpoints[^1]);
                var value = new SheetSelection(selection.Attribute("activeCell") is null ? first : cell, new(first.Row, first.Column, last.Row, last.Column),
                    selection.Attribute("activeCell") is null, selection.Attribute("sqref") is null);
                result = selectionPane switch { SheetPane.TopLeft => result with { TopLeft = value }, SheetPane.TopRight => result with { TopRight = value },
                    SheetPane.BottomLeft => result with { BottomLeft = value }, _ => result with { BottomRight = value } };
            }
            _ = new WorksheetSnapshot(SheetId.New(), "View", view: result);
            return result;
        }
        catch (Exception e) when (e is NotSupportedException or ArgumentException)
        {
            package.Loss(path + ": 视图仅保留，显示使用默认视图：" + e.Message, "sheetView/unsupported", path, view);
            return new();
        }
    }

    private static int? Zoom(XElement view, string name)
    {
        if (view.Attribute(name) is not { } value) return null;
        if (!int.TryParse(value.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var zoom)) throw new InvalidDataException("缩放格式无效。");
        if (zoom is < 10 or > 400) throw new NotSupportedException("缩放超出 10%–400%。");
        return zoom;
    }
    private static int Split(XElement pane, string name, int limit)
    {
        var text = (string?)pane.Attribute(name) ?? "0";
        if (!double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var count) || !double.IsFinite(count) || count < 0)
            throw new InvalidDataException("冻结位置无效。");
        if (count != Math.Truncate(count) || count >= limit) throw new NotSupportedException("冻结位置超出支持范围。");
        return (int)count;
    }
    private static CellPosition Cell(string? text)
    {
        text ??= "A1";
        var match = System.Text.RegularExpressions.Regex.Match(text, "^([A-Z]{1,3})([1-9][0-9]{0,6})$");
        if (!match.Success) throw new InvalidDataException("视图单元格地址无效。");
        var column = 0; foreach (var c in match.Groups[1].Value) column = column * 26 + c - 'A' + 1;
        var row = int.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture);
        if (column > 16384 || row > 1048576) throw new InvalidDataException("视图地址超出 XLSX 上限。");
        if (column > WorkbookLimits.MaxColumns || row > WorkbookLimits.MaxRows) throw new NotSupportedException("视图地址超出当前工作簿容量。");
        return new(row - 1, column - 1);
    }
    private static SheetPane Pane(string? value) => value switch
    { null or "topLeft" => SheetPane.TopLeft, "topRight" => SheetPane.TopRight, "bottomLeft" => SheetPane.BottomLeft, "bottomRight" => SheetPane.BottomRight, _ => throw new InvalidDataException("活动窗格无效。") };
    private static string PaneName(SheetPane pane) => pane switch
    { SheetPane.TopRight => "topRight", SheetPane.BottomLeft => "bottomLeft", SheetPane.BottomRight => "bottomRight", _ => "topLeft" };

    public static (SheetId Active, WorkbookWindow Window) ReadWorkbook(XlsxPackage package, string path, XElement workbook, IReadOnlyList<WorksheetSnapshot> sheets)
    {
        var views = workbook.Element(Main + "bookViews");
        if (views is not null) package.Shape(views, "workbookView", "", path);
        if (views?.Elements(Main + "workbookView").Count() > 1) package.Loss(path + ": 多工作簿窗口尚未支持编辑。", "bookViews/multiple", path);
        var view = views?.Element(Main + "workbookView");
        var active = sheets.FirstOrDefault(s => s.Visibility == SheetVisibility.Visible)?.Id ?? throw new InvalidDataException("工作簿没有可见工作表。");
        if (view is null) return (active, new());
        package.Shape(view, "", "activeTab xWindow yWindow windowWidth windowHeight firstSheet tabRatio visibility minimized showHorizontalScroll showVerticalScroll showSheetTabs autoFilterDateGrouping", path);
        foreach (var attribute in view.Attributes().Where(a => a.Name.LocalName is "firstSheet" or "tabRatio" or "visibility" or "minimized" or "showHorizontalScroll" or "showVerticalScroll" or "showSheetTabs" or "autoFilterDateGrouping"))
        {
            var normal = attribute.Name.LocalName switch { "firstSheet" => "0", "tabRatio" => "600", "visibility" => "visible", "minimized" => "0", _ => "1" };
            var value = attribute.Value is "true" ? "1" : attribute.Value is "false" ? "0" : attribute.Value;
            if (value != normal) package.Loss(path + ": 非默认窗口选项尚未建模：" + attribute.Name.LocalName, "workbookView/@" + attribute.Name.LocalName, path);
        }
        if (view.Attribute("activeTab") is { } index)
        {
            var tab = Int(index.Value, sheets.Count - 1);
            if (sheets[tab].Visibility != SheetVisibility.Visible) throw new InvalidDataException("活动表指向隐藏工作表。");
            active = sheets[tab].Id;
        }
        int? Signed(string name) => view.Attribute(name) is { } attr ? int.Parse(attr.Value, CultureInfo.InvariantCulture) : null;
        uint? Size(string name) => view.Attribute(name) is { } attr ? uint.Parse(attr.Value, CultureInfo.InvariantCulture) : null;
        return (active, new(Signed("xWindow"), Signed("yWindow"), Size("windowWidth"), Size("windowHeight")));
    }

    public static XElement WriteSheet(WorksheetSnapshot sheet, bool active)
    {
        var view = sheet.View;
        var root = new XElement(Main + "sheetView", new XAttribute("workbookViewId", 0), new XAttribute("showGridLines", sheet.ShowGridLines ? 1 : 0));
        if (active) root.SetAttributeValue("tabSelected", 1);
        if (view.Zoom != 100) root.SetAttributeValue("zoomScale", view.Zoom);
        if (view.NormalZoom is { } normal) root.SetAttributeValue("zoomScaleNormal", normal);
        if (view.TopLeftCell != default) root.SetAttributeValue("topLeftCell", Reference(view.TopLeftCell));
        if (view.HasFreeze) root.Add(new XElement(Main + "pane", view.FrozenColumns > 0 ? new XAttribute("xSplit", view.FrozenColumns) : null,
            view.FrozenRows > 0 ? new XAttribute("ySplit", view.FrozenRows) : null, new XAttribute("topLeftCell", Reference(view.PaneTopLeftCell)),
            new XAttribute("activePane", PaneName(view.ActivePane)), new XAttribute("state", view.FrozenSplit ? "frozenSplit" : "frozen")));
        foreach (var (pane, selection) in new[] { (SheetPane.TopLeft, view.TopLeft), (SheetPane.TopRight, view.TopRight), (SheetPane.BottomLeft, view.BottomLeft), (SheetPane.BottomRight, view.BottomRight) })
            if (selection is { } selected) root.Add(new XElement(Main + "selection", pane == SheetPane.TopLeft ? null : new XAttribute("pane", PaneName(pane)),
                selected.OmitActiveCell ? null : new XAttribute("activeCell", Reference(selected.ActiveCell)), selected.OmitRange ? null : new XAttribute("sqref", selected.Range.ToString())));
        return new(Main + "sheetViews", root);
    }
    public static XElement WriteWorkbook(WorkbookSnapshot snapshot)
    {
        var window = snapshot.Window;
        return new(Main + "bookViews", new XElement(Main + "workbookView", new XAttribute("activeTab", snapshot.Sheets.IndexOf(snapshot.GetSheet(snapshot.ActiveSheetId))),
            window.X is { } x ? new XAttribute("xWindow", x) : null, window.Y is { } y ? new XAttribute("yWindow", y) : null,
            window.Width is { } w ? new XAttribute("windowWidth", w) : null, window.Height is { } h ? new XAttribute("windowHeight", h) : null));
    }
}
