using QuickExcel.Domain;
using QuickExcel.Rendering.Skia;
using QuickExcel.Rendering.Smoke;
using SkiaSharp;
using UglyToad.PdfPig;

namespace QuickExcel.Rendering.Test;

public sealed class RenderingTests
{
    private static string[] Fonts()
    {
        var directory = Environment.GetEnvironmentVariable("QUICKEXCEL_TEST_FONTS");
        if (directory is null)
        {
            var root = new DirectoryInfo(AppContext.BaseDirectory);
            while (root is not null && !File.Exists(Path.Combine(root.FullName, "QuickExcel.slnx"))) root = root.Parent;
            directory = Path.Combine(root?.FullName ?? throw new InvalidOperationException("Set QUICKEXCEL_TEST_FONTS."), "artifacts", "skia-p0", "fonts");
        }
        Assert.True(Directory.Exists(directory), "Run scripts/get-rendering-test-fonts.ps1 first.");
        return Directory.GetFiles(directory).Where(f => f.EndsWith(".ttf") || f.EndsWith(".otf")).ToArray();
    }
    private static SkiaRenderer Renderer(RenderBudget? budget = null) => new(Fonts(), false, budget);
    private static SkiaReport Report(RenderBudget? budget = null)
    { var book = ReportFixture.Create(); return new(book, book.Sheets[0], "fixture.xlsx", fontFiles: Fonts(), systemFonts: false, budget: budget); }
    private static readonly StyleDefinition Default = new() { FontName = "Noto Sans", EastAsianFontName = "Noto Sans CJK SC", FontSize = 11 };

    [Fact]
    public void LayoutUsesPointsAndSharesWrappingRotationAndShrinkMetrics()
    {
        using var renderer = Renderer();
        using var small = renderer.Layout(new("QuickExcel 中文", Default));
        using var large = renderer.Layout(new("QuickExcel 中文", Default with { FontSize = 22 }));
        Assert.InRange(large.Metrics.Width / small.Metrics.Width, 1.98, 2.02);
        using var wrapped = renderer.Layout(new("e\u0301 word 中文 wrapping text", Default with { WrapText = true }, 60));
        Assert.True(wrapped.Metrics.Lines >= 3); Assert.True(wrapped.Metrics.Width <= 61);
        using var rotated = renderer.Layout(new("QuickExcel 中文", Default with { TextRotation = 90 }));
        Assert.Equal(rotated.Metrics.Width, rotated.Metrics.RotatedHeight, 4);
        using var shrunk = renderer.Layout(new("QuickExcel 中文", Default with { ShrinkToFit = true }, 40));
        Assert.InRange(shrunk.Metrics.Width, 35, 41);
        Assert.DoesNotContain(renderer.Diagnostics, d => d.Code == "missing-glyph");
        using var explicitLines = renderer.Layout(new("first\nsecond\n第三行", Default, 500));
        Assert.Equal(3, explicitLines.Metrics.Lines);
        Assert.True(explicitLines.Metrics.Height > small.Metrics.Height * 2.5);
    }

    [Fact]
    public void RichRunsInheritBaseFontWithoutFlattening()
    {
        using var renderer = Renderer();
        var rich = new RichCellText([new("Inherited "), new("bold", new(Bold: true, Color: "FFFF0000")), new(" small", new(Size: 8))]);
        using var text = renderer.Layout(new(rich.Text, Default, RichText: rich));
        Assert.Contains(text.FontIdentities, f => f.Contains("weight=700"));
        Assert.Contains(text.FontIdentities, f => f.Contains("weight=400"));
        Assert.Equal(3, rich.Runs.Length); Assert.Null(rich.Runs[0].Font);
    }

    [Fact]
    public void MissingFontsAndScalarsHaveBoundedLocatedDiagnostics()
    {
        using var renderer = Renderer(new() { MaxDiagnostics = 4 });
        for (var i = 0; i < 20; i++)
        { using var text = renderer.Layout(new("Missing \U0010FFFF", Default with { FontName = "QuickExcel Missing Font" }), $"A{i}"); }
        Assert.InRange(renderer.Diagnostics.Count, 1, 4);
        Assert.Contains(renderer.Diagnostics, d => d.Code == "missing-glyph" && d.Location == "A0");
        Assert.Contains(renderer.Diagnostics, d => d.Code == "font-substitution");
    }

    [Fact]
    public void ReportIsFrozenRepeatsTitlesAndHonorsHiddenRows()
    {
        using var report = Report();
        var before = report.Job.Book;
        Assert.Equal(3, report.Job.Layout.Pages.Length);
        Assert.All(report.Job.Layout.Pages, p => { Assert.Equal(0, p.Rows[0]); Assert.Equal(1, p.Rows[1]); Assert.DoesNotContain(6, p.Rows); });
        Assert.DoesNotContain(24, report.Job.Layout.Pages[0].Rows);
        Assert.Contains(24, report.Job.Layout.Pages[1].Rows);
        using var bitmap = report.RenderPage(0, 144);
        Assert.Equal((int)Math.Ceiling(report.Job.Layout.PaperWidth * 1.5), bitmap.Width);
        Assert.Same(before, report.Job.Book); Assert.Empty(report.Job.Warnings);
    }

    [Fact]
    public void PdfIsMultipageTextAndVectorAndHasPhysicalPaperSize()
    {
        using var report = Report(); using var folder = new TestFolder();
        var path = Path.Combine(folder.Path, "report.pdf"); report.WritePdf(path);
        using var pdf = PdfDocument.Open(path);
        Assert.Equal(3, pdf.NumberOfPages);
        foreach (var page in pdf.GetPages())
        {
            Assert.InRange(Math.Abs(page.Width - report.Job.Layout.PaperWidth * .75), 0, .2);
            Assert.Contains("QuickExcel", page.Text); Assert.Contains("Record", page.Text);
            Assert.True(page.Paths.Count > 30); Assert.InRange(page.NumberOfImages, 0, 2);
        }
        Assert.Empty(Directory.GetFiles(folder.Path, "*.tmp"));
    }

    [Fact]
    public void CancellationAndOutputBudgetPreserveExistingDestinationAndCleanScratch()
    {
        using var report = Report(new() { MaxPdfBytes = 10 }); using var folder = new TestFolder();
        var path = Path.Combine(folder.Path, "existing.pdf"); File.WriteAllText(path, "original");
        Assert.Throws<IOException>(() => report.WritePdf(path)); Assert.Equal("original", File.ReadAllText(path));
        using var canceled = new CancellationTokenSource(); canceled.Cancel();
        Assert.Throws<OperationCanceledException>(() => report.WritePdf(path, canceled.Token));
        Assert.Equal("original", File.ReadAllText(path)); Assert.Single(Directory.GetFiles(folder.Path));
    }

    [Fact]
    public void DestinationReplacementFailureKeepsOriginalAndCleansScratch()
    {
        using var report = Report(); using var folder = new TestFolder();
        var target = Path.Combine(folder.Path, "destination");
        Directory.CreateDirectory(target); File.WriteAllText(Path.Combine(target, "original"), "keep");
        Assert.True(Record.Exception(() => report.WritePdf(target)) is IOException or UnauthorizedAccessException);
        Assert.Equal("keep", File.ReadAllText(Path.Combine(target, "original")));
        Assert.Empty(Directory.GetFiles(folder.Path, "*.tmp"));
    }

    [Fact]
    public void CancellationDuringPageRenderingCleansPartiallyWrittenPdf()
    {
        using var folder = new TestFolder(); using var canceled = new CancellationTokenSource();
        var book = ReportFixture.Create(); var cells = 0;
        using var report = new SkiaReport(book, book.Sheets[0], "cancel.xlsx", (b, s, row, column) =>
        {
            if (++cells == 10) canceled.Cancel();
            return b.Styles[s.GetStyleId(row, column)];
        }, Fonts(), false);
        var path = Path.Combine(folder.Path, "existing.pdf"); File.WriteAllText(path, "original");
        Assert.Throws<OperationCanceledException>(() => report.WritePdf(path, canceled.Token));
        Assert.True(cells >= 10); Assert.Equal("original", File.ReadAllText(path)); Assert.Single(Directory.GetFiles(folder.Path));
    }

    [Fact]
    public void PreviewAndSinglePageWorkAreBudgeted()
    {
        using var report = Report(new() { MaxPreviewBytes = 1024 });
        Assert.Throws<NotSupportedException>(() => report.RenderPage(0));
        Assert.Throws<NotSupportedException>(() => Report(new() { MaxPageCells = 5 }));
        Assert.Throws<ArgumentOutOfRangeException>(() => report.RenderPage(0, double.NaN));
    }

    [Fact]
    public void UnsupportedPrintContentOnlyChangesOutputProjection()
    {
        var book = ReportFixture.Create(); var original = book.Sheets[0];
        var sheet = original.WithImages(original.Images.Add(new([1, 2, 3], "png", 0, 0)))
            .WithPrinting(original.Printing with { Setup = original.Printing.Setup! with { CellComments = "atEnd" }, HeaderFooter = new() { OddHeader = "&G", OddFooter = "Keep &P" } });
        book = book.WithSheets([sheet]);
        using var report = new SkiaReport(book, sheet, "report.xlsx", fontFiles: Fonts(), systemFonts: false);
        Assert.Equal(3, report.Job.Warnings.Count); Assert.Equal(2, sheet.Images.Length);
        Assert.Single(report.Job.Sheet.Images); Assert.Null(report.Job.Sheet.Printing.HeaderFooter!.OddHeader);
        Assert.Equal("&G", sheet.Printing.HeaderFooter!.OddHeader); Assert.Equal("atEnd", sheet.Printing.Setup!.CellComments);
    }

    private sealed class TestFolder : IDisposable
    {
        public string Path { get; } = System.IO.Path.Combine(FindArtifacts(), "rendering-tests", Guid.NewGuid().ToString("N"));
        public TestFolder() => Directory.CreateDirectory(Path);
        private static string FindArtifacts()
        {
            var root = new DirectoryInfo(AppContext.BaseDirectory);
            while (root is not null && root.Name != "artifacts") root = root.Parent;
            return root?.FullName ?? throw new InvalidOperationException("Use --artifacts-path under this repository.");
        }
        public void Dispose() => Directory.Delete(Path, true);
    }
}
