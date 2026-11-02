using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.Json;
using QuickExcel.Domain;
using QuickExcel.Rendering;
using QuickExcel.Rendering.Skia;
using QuickExcel.Rendering.Smoke;
using SkiaSharp;
using UglyToad.PdfPig;

// Run on an isolated full filesystem to exercise a real OS write failure.
if (args.ElementAtOrDefault(0) == "--pdf-storage-failure")
{
    var destination = Path.GetFullPath(args[1]); var original = File.ReadAllBytes(destination);
    var files = Directory.GetFiles(args[2]).Where(f => f.EndsWith(".ttf") || f.EndsWith(".otf")).ToArray();
    var fixture = ReportFixture.Create();
    using var task = new SkiaReport(fixture, fixture.Sheets[0], "storage-failure.xlsx", fontFiles: files, systemFonts: false);
    IOException? failure = null;
    try { task.WritePdf(destination); }
    catch (IOException ex) { failure = ex; }
    if (failure is null || !original.SequenceEqual(File.ReadAllBytes(destination)) || Directory.GetFiles(Path.GetDirectoryName(destination)!, "*.tmp").Length != 0)
        throw new InvalidDataException("Storage failure did not preserve the existing PDF and clean scratch files.");
    Console.WriteLine(JsonSerializer.Serialize(new { storageFailure = failure.Message, originalPreserved = true, temporaryFilesCleaned = true }));
    return;
}

var output = Path.GetFullPath(args.ElementAtOrDefault(0) ?? "artifacts/skia-p0/text");
Directory.CreateDirectory(output);
var fontDirectory = args.ElementAtOrDefault(1);
var fontFiles = fontDirectory is null ? [] : Directory.GetFiles(fontDirectory).Where(f => f.EndsWith(".ttf") || f.EndsWith(".otf")).ToArray();
foreach (var path in fontFiles) { using var face = SKTypeface.FromFile(path); Console.WriteLine($"FONT {Path.GetFileName(path)}: {face.FamilyName}; weight={face.FontWeight}; italic={face.IsItalic}"); }
using var renderer = new SkiaRenderer(fontFiles, systemFonts: fontDirectory is null);
var watch = Stopwatch.StartNew();
using var surface = SKSurface.Create(new SKImageInfo(900, 420));
void Draw(SKCanvas canvas)
{
    canvas.Clear(SKColors.White);
    string[] samples = ["QuickExcel 中文报表：收入、成本和利润", "组合字符 e\u0301 A\u030a — office ffi", "Mixed direction: العربية 123 עברית", "Supplementary: 𝄞 𠮷", "换行和富文本：共享布局，共享测量。"];
    for (var i = 0; i < samples.Length; i++)
    {
        var style = new StyleDefinition { FontName = "Noto Sans", EastAsianFontName = "Noto Sans CJK SC", FontSize = 18, Vertical = 1 };
        renderer.DrawCell(canvas, new(new(10, 10 + i * 65, 870, 60), new(samples[i], style), true), location: $"A{i + 1}");
    }
}
Draw(surface.Canvas);
using (var png = surface.Snapshot().Encode(SKEncodedImageFormat.Png, 100))
using (var file = File.Create(Path.Combine(output, "text.png"))) png.SaveTo(file);
using (var stream = File.Create(Path.Combine(output, "text.pdf")))
using (var pdf = SKDocument.CreatePdf(stream))
{
    var canvas = pdf.BeginPage(675, 315); canvas.Scale(.75f); Draw(canvas); pdf.EndPage(); pdf.Close();
}
using var parsed = PdfDocument.Open(Path.Combine(output, "text.pdf"));
var extracted = parsed.GetPage(1).Text;
File.WriteAllText(Path.Combine(output, "text.txt"), extracted);
// PdfPig's Page.Text omits ActualText substitutions. Semantic extraction is separately
// verified by scripts/verify-skia-output.py using MuPDF, including combining clusters.
if (!extracted.Contains("QuickExcel")) throw new InvalidDataException("PDF text objects are missing.");
File.WriteAllText(Path.Combine(output, "result.json"), JsonSerializer.Serialize(new { platform = RuntimeInformation.OSDescription, architecture = RuntimeInformation.ProcessArchitecture.ToString(),
    runtime = RuntimeInformation.FrameworkDescription, elapsedMs = watch.Elapsed.TotalMilliseconds, privateBytes = Process.GetCurrentProcess().PrivateMemorySize64,
    fonts = fontFiles.Select(f => new { file = Path.GetFileName(f), sha256 = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(f))) }), diagnostics = renderer.Diagnostics,
    extracted }, new JsonSerializerOptions { WriteIndented = true }));
Console.WriteLine($"PASS: Skia text/Unicode PDF on {RuntimeInformation.RuntimeIdentifier}: {output}");
var book = ReportFixture.Create();
using var report = new SkiaReport(book, book.Sheets[0], "fixture.xlsx", fontFiles: fontFiles, systemFonts: fontDirectory is null, timestamp: new DateTime(2026, 11, 2, 12, 0, 0));
report.WritePdf(Path.Combine(output, "report.pdf"));
for (var page = 0; page < report.Job.Layout.Pages.Length; page++)
{
    using var bitmap = report.RenderPage(page);
    using var data = bitmap.Encode(SKEncodedImageFormat.Png, 100);
    using var file = File.Create(Path.Combine(output, $"report-{page + 1}.png")); data.SaveTo(file);
}
File.WriteAllText(Path.Combine(output, "report-layout.json"), JsonSerializer.Serialize(report.Job.Layout, new JsonSerializerOptions { WriteIndented = true }));
File.WriteAllText(Path.Combine(output, "report-diagnostics.json"), JsonSerializer.Serialize(new { report.Job.Warnings, report.Renderer.Diagnostics }, new JsonSerializerOptions { WriteIndented = true }));
Console.WriteLine($"PASS: {report.Job.Layout.Pages.Length} report pages, PDF and PNG.");
var formats = new List<object>();
foreach (var extension in new[] { "png", "jpg", "gif", "bmp", "tiff" })
{
    var bytes = ImageFixtures.Bytes(extension); var image = new SheetImage(bytes, extension, 0, 0);
    using var decoded = renderer.DecodeImage(image); var pixel = decoded.GetPixel(0, 0);
    if (decoded.Width != 2 || decoded.Height != 1 || pixel.Red < 180 || pixel.Green > 80 || pixel.Blue > 80 || !bytes.SequenceEqual(image.Data))
        throw new InvalidDataException($"Image format/first-frame failure: {extension}");
    formats.Add(new { format = extension, decoded.Width, decoded.Height, firstPixel = pixel.ToString() });
}
var sentinel = Path.Combine(output, "atomic-target.pdf"); File.WriteAllText(sentinel, "existing");
try { report.WritePdf(sentinel, new CancellationToken(true)); throw new InvalidOperationException("Canceled export succeeded."); }
catch (OperationCanceledException) { }
using (var limited = new SkiaReport(book, book.Sheets[0], "fixture.xlsx", fontFiles: fontFiles, systemFonts: fontDirectory is null, budget: new() { MaxPdfBytes = 100 }))
{
    try { limited.WritePdf(sentinel); throw new InvalidOperationException("Oversized export succeeded."); }
    catch (IOException) { }
}
if (File.ReadAllText(sentinel) != "existing" || Directory.GetFiles(output, "*.tmp").Length != 0) throw new InvalidDataException("Atomic failure cleanup changed the destination.");
report.WritePdf(sentinel); using (var replacement = PdfDocument.Open(sentinel)) if (replacement.NumberOfPages != 3) throw new InvalidDataException("Atomic replacement failed.");
File.Delete(sentinel);
File.WriteAllText(Path.Combine(output, "platform-checks.json"), JsonSerializer.Serialize(new { formats, atomicReplace = true, canceledPreservesTarget = true, budgetFailurePreservesTarget = true }, new JsonSerializerOptions { WriteIndented = true }));
Console.WriteLine("PASS: five image formats, first frames, atomic PDF replacement, cancellation and budget-failure cleanup.");
