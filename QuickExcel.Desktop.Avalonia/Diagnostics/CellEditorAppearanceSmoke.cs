using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Presenters;
using Avalonia.Input;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.Media.TextFormatting;
using Avalonia.Threading;
using Avalonia.VisualTree;
using QuickExcel.Calculation;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Rendering;
using QuickExcel.Rendering.Skia;
using QuickExcel.Storage;
using SkiaSharp;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task RunCellEditorAppearanceChecksAsync(string directory, Action<bool, string> check)
    {
        static Task Layout() => Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background).GetTask();
        const string wrappingText = "MMMM MMMM MMMM";
        var styled = new StyleDefinition { FontName = "Consolas", FontSize = 18, Bold = true, Italic = true,
            FontColor = "FF1450CC", FillColor = "FFEDF2FF", WrapText = true, Vertical = 0 };
        StyleDefinition[] styles = [new(), styled, styled with { WrapText = false },
            new() { FontName = "Times New Roman", FontSize = 26, FontColor = "FF008000", Vertical = 0 },
            new() { FontName = "Microsoft YaHei UI", FontSize = 16, FontColor = "FFCC0000", WrapText = true, Vertical = 0 },
            new() { FontColor = "FF0000FF", NumberFormatCode = "0;[Red]-0", FontSize = 18 },
            styled with { FontName = "Arial", FontSize = 16, FontColor = "FFFFFFFF", FillColor = "FF202020", Horizontal = 2, Vertical = 1 }];
        CellSnapshot Text(string text, int style = 0) => new(CellContent.FromLiteral(CellValue.FromText(text), style));
        var sheet = new WorksheetSnapshot(SheetId.New(), "Editor appearance", new Dictionary<CellPosition, CellSnapshot>
        {
            [new(0, 0)] = Text(wrappingText, 1), [new(0, 1)] = Text(wrappingText, 2), [new(0, 2)] = Text("Default"),
            [new(1, 0)] = Text("Wide MMM", 3), [new(1, 1)] = Text("中文自动换行中文自动换行", 4),
            [new(1, 2)] = new(CellContent.FromFormula("=-2", 5)), [new(2, 0)] = Text("Merged white text wraps across this cell", 6),
            [new(2, 2)] = Text("Conditional")
        }, rows: Enumerable.Range(0, 5).ToDictionary(r => r, r => new RowMetadata(Height: 140, StyleId: r == 4 ? 1 : 0)),
            columns: new Dictionary<int, ColumnMetadata> { [3] = new(StyleId: 2) }, mergedCells: ["A3:B4"]);
        var path = Path.Combine(directory, "editor-appearance.xlsq");
        new TypedWorkbookStorage().Save(path, new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet], styles: styles));
        await OpenPathAsync(path); await CalculateAsync(); await Layout();
        var editor = _grid.CellEditor;
        TextPresenter Presenter() => editor.GetVisualDescendants().OfType<TextPresenter>().Single();
        using var renderer = new SkiaRenderer();

        async Task CheckAppearance(CellPosition position, string name, string? draft = null)
        {
            _grid.Select(position); check(_grid.BeginEdit(), name + ": starts editing");
            if (draft is not null) editor.Text = draft;
            editor.SelectionStart = editor.SelectionEnd = editor.CaretIndex = 0; _grid.Focus(); await Layout();
            var anchor = _grid.Current;
            var style = ConditionalPresentation.GetStyle(Document.Snapshot, _grid.Sheet, anchor.Row, anchor.Column);
            var display = CellDisplayFormatter.FormatDisplay(_grid.Sheet.GetCell(anchor.Row, anchor.Column)?.Value ?? CellValue.Blank, style, Document.Snapshot.DateSystem);
            var foreground = SkiaRenderer.Color(display.Color ?? style.FontColor, SKColors.Black);
            check(editor.CaretBrush is ISolidColorBrush caret && caret.Color == Color.FromArgb(foreground.Alpha, foreground.Red, foreground.Green, foreground.Blue),
                name + ": caret remains visible in the cell text color");
            var fontSize = (style.FontSize is { } points ? points * 96 / 72 : 14) * _grid.Zoom;
            var font = renderer.ResolveFont(new(style.FontName ?? "Microsoft YaHei UI", style.EastAsianFontName,
                style.FontSize ?? 14 * 72d / 96, style.Bold, style.Italic));
            check(editor.FontFamily.FamilyNames[0] == font.Family && Math.Abs(editor.FontSize - fontSize) < .01 &&
                editor.FontWeight == (font.Bold ? FontWeight.Bold : FontWeight.Normal) && editor.FontStyle == (font.Italic ? FontStyle.Italic : FontStyle.Normal),
                name + ": font family, point size, zoom, weight and slant match effective cell style");
            var presenter = Presenter();
            var runs = presenter.TextLayout.TextLines.SelectMany(l => l.TextRuns).OfType<ShapedTextRun>().ToArray();
            check(runs.Length > 0 && runs.All(r => Math.Abs(r.ShapedBuffer.FontRenderingEmSize - fontSize) < .01),
                name + ": native glyphs use the requested rendering size");
            using var layout = renderer.Layout(new(editor.Text!, style, Math.Max(1, editor.Width / _grid.Zoom - 4)));
            check(runs.All(r => layout.FontIdentities.Any(f => f.StartsWith(r.ShapedBuffer.GlyphTypeface.FamilyName + ";", StringComparison.Ordinal))),
                name + ": native glyphs use the same resolved fonts as the grid; native=" + string.Join(",", runs.Select(r => r.ShapedBuffer.GlyphTypeface.FamilyName).Distinct()) + "; grid=" + string.Join(",", layout.FontIdentities));
            check(presenter.TextLayout.TextLines.Count == layout.Metrics.Lines,
                $"{name}: native and grid layouts have the same line count ({presenter.TextLayout.TextLines.Count})");
            var file = Path.Combine(directory, "editor-appearance-" + name + ".png");
            using (var image = new RenderTargetBitmap(new PixelSize((int)Math.Ceiling(editor.Bounds.Width), (int)Math.Ceiling(editor.Bounds.Height)), new Vector(96, 96)))
            { image.Render(editor); image.Save(file, PngBitmapEncoderOptions.Default); }
            using var pixels = SKBitmap.Decode(file);
            var background = SkiaRenderer.Color(style.FillColor, SKColors.White);
            check(pixels.GetPixel(pixels.Width - 3, 3) == background, name + ": cell background is preserved, including behind white text");
            var ink = 0;
            for (var y = 2; y < pixels.Height - 2; y++)
            for (var x = 2; x < pixels.Width - 2; x++)
            {
                var color = pixels.GetPixel(x, y);
                if (Math.Abs(color.Red - foreground.Red) < 15 && Math.Abs(color.Green - foreground.Green) < 15 && Math.Abs(color.Blue - foreground.Blue) < 15) ink++;
            }
            check(ink > 10, name + ": rendered text pixels have the effective foreground color");
            _grid.Cancel();
        }

        var revision = Document.Snapshot.Revision;
        await CheckAppearance(new(0, 0), "wrapped");
        await CheckAppearance(new(0, 1), "unwrapped");
        await CheckAppearance(new(0, 2), "default-reset");
        await CheckAppearance(new(1, 0), "serif");
        await CheckAppearance(new(1, 1), "chinese");
        await CheckAppearance(new(1, 2), "number-format-color");
        await CheckAppearance(new(3, 1), "merged-white");
        await CheckAppearance(new(4, 1), "row-inherited", wrappingText);
        await CheckAppearance(new(0, 3), "column-inherited", wrappingText);
        check(Document.Snapshot.Revision == revision && Document.GetCellInput(0, 1, 2) == "=-2",
            "Applying appearance and cancelling drafts preserves content, formulas and workbook styles");
        await ExecuteAsync(new SetConditionalFormatsCommand(sheet.Id, [new([new(2, 2, 2, 2)], ConditionalFormatKind.Expression, 1,
            new(FontColor: "FF800080", Bold: true, Italic: true), "=TRUE")]));
        await CheckAppearance(new(2, 2), "conditional");
        _grid.Select(new(0, 0)); await ApplyStyleAsync(new(FontName: "Arial", FontSize: 14, Bold: false, Italic: false, FontColor: "FF008000", WrapText: false));
        await CheckAppearance(new(0, 0), "format-refresh");
        await History(false); await CheckAppearance(new(0, 0), "format-undo");
        await ZoomAsync(150); await CheckAppearance(new(0, 0), "zoom-wrap");
        _grid.Select(new(0, 0)); _grid.BeginEdit(); await Layout();
        var composing = Presenter(); composing.PreeditText = "zhong";
        check(_grid.IsComposing && !CommitEditors() && editor.TextWrapping == TextWrapping.Wrap,
            "Styled wrapping editor keeps the native IME composition gate");
        composing.PreeditText = null;
        editor.CaretIndex = editor.SelectionStart = editor.SelectionEnd = editor.Text!.Length;
        editor.RaiseEvent(new TextInputEventArgs { RoutedEvent = InputElement.TextInputEvent, Text = " added" });
        check(editor.Text == wrappingText + " added" && await _grid.CommitAsync(), "Styled wrapping editor types and commits through the original cell draft");
        await History(false); check(Document.GetCellInput(0, 0, 0) == wrappingText, "Styled edit is undone in one history step");
    }
}
