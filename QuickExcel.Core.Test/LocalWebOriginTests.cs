using System.Net;
using System.Net.Http.Headers;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using QuickExcel.Infrastructure.Web;

namespace QuickExcel.Core.Test;

public sealed class LocalWebOriginTests : IDisposable
{
    private readonly string _directory = Path.Combine(AppContext.BaseDirectory, "test-data", "web-origin-" + Guid.NewGuid().ToString("N"));
    public LocalWebOriginTests()
    {
        Directory.CreateDirectory(_directory);
        var bytes = Encoding.UTF8.GetBytes("<!doctype html><title>QuickExcel</title>");
        File.WriteAllBytes(Path.Combine(_directory, "script.html"), bytes);
        File.WriteAllText(Path.Combine(_directory, "manifest.json"), JsonSerializer.Serialize(new { schemaVersion = 1, packages = new { test = "1" },
            files = new[] { new { name = "script.html", size = bytes.Length, sha256 = Convert.ToHexString(SHA256.HashData(bytes)) } } }));
    }
    public void Dispose() => Directory.Delete(_directory, true);
    private LocalWebOrigin Create() => new(new(_directory), "script.html", Guid.NewGuid().ToString("N"), 1);
    [Theory]
    [InlineData("testBuild")]
    [InlineData("probeBuild")]
    public void ProductCatalogRejectsExplicitTestAndProbeBundles(string flag)
    {
        var path = Path.Combine(_directory, "manifest.json"); File.WriteAllText(path, File.ReadAllText(path).Insert(1, $"\"{flag}\":true,"));
        Assert.Throws<InvalidDataException>(() => new WebAssetCatalog(_directory));
        Assert.NotNull(new WebAssetCatalog(_directory, allowTestResources: true));
    }
    [Fact]
    public void CatalogRejectsChangedOrEscapingAssetsBeforeStartingBrowser()
    {
        File.AppendAllText(Path.Combine(_directory, "script.html"), "tamper");
        Assert.Throws<InvalidDataException>(() => new WebAssetCatalog(_directory));
        var manifest = File.ReadAllText(Path.Combine(_directory, "manifest.json")).Replace("script.html", "../script.html"); File.WriteAllText(Path.Combine(_directory, "manifest.json"), manifest);
        Assert.Throws<InvalidDataException>(() => new WebAssetCatalog(_directory));
    }
    [Fact]
    public async Task OriginRequiresExactHostTokenPathAndSameOriginRequests()
    {
        await using var origin = Create(); using var client = new HttpClient();
        using var valid = await client.GetAsync(origin.Page); Assert.Equal(HttpStatusCode.OK, valid.StatusCode); Assert.True(valid.Headers.Contains("Content-Security-Policy"));
        foreach (var name in new[] { "/script.html", $"/{origin.Instance}/script.html?x=1", $"/{origin.Instance}/%2e%2e/script.html", $"/{origin.Instance}/documents/unknown.pdf" })
            Assert.Equal(HttpStatusCode.NotFound, (await client.GetAsync(origin.Page.GetLeftPart(UriPartial.Authority) + name)).StatusCode);
        foreach (var (header, value) in new[] { ("Host", "attacker.invalid"), ("Origin", "https://attacker.invalid"), ("Sec-Fetch-Site", "cross-site") })
        {
            using var request = new HttpRequestMessage(HttpMethod.Get, origin.Page); request.Headers.TryAddWithoutValidation(header, value);
            using var result = await client.SendAsync(request); Assert.Equal(HttpStatusCode.NotFound, result.StatusCode);
        }
        using var post = await client.PostAsync(origin.Page, new StringContent("")); Assert.Equal(HttpStatusCode.NotFound, post.StatusCode);
    }
    [Fact]
    public async Task PdfRangeLeaseServesExactBytesThenRevokesWithoutChangingSource()
    {
        await using var origin = Create(); var path = Path.Combine(_directory, "fixture.pdf"); var bytes = Encoding.ASCII.GetBytes("%PDF-test-resource-boundary"); File.WriteAllBytes(path, bytes);
        await using var lease = await origin.AddPdfAsync(path); using var client = new HttpClient();
        using var request = new HttpRequestMessage(HttpMethod.Get, lease.Uri); request.Headers.Range = new RangeHeaderValue(2, 8);
        using var partial = await client.SendAsync(request); Assert.Equal(HttpStatusCode.PartialContent, partial.StatusCode);
        Assert.Equal(bytes[2..9], await partial.Content.ReadAsByteArrayAsync()); Assert.Equal(bytes.Length, partial.Content.Headers.ContentRange!.Length);
        using var invalid = new HttpRequestMessage(HttpMethod.Get, lease.Uri); invalid.Headers.Range = new RangeHeaderValue(1000, null);
        using var rejected = await client.SendAsync(invalid); Assert.Equal(HttpStatusCode.RequestedRangeNotSatisfiable, rejected.StatusCode);
        await lease.DisposeAsync(); using var expired = await client.GetAsync(lease.Uri); Assert.Equal(HttpStatusCode.NotFound, expired.StatusCode);
        Assert.Equal(bytes, File.ReadAllBytes(path));
    }
    [Fact]
    public async Task ClosingOriginStopsPendingClientsAndReleasesFiles()
    {
        var origin = Create(); var path = Path.Combine(_directory, "fixture.pdf"); File.WriteAllText(path, "%PDF-test");
        await using var lease = await origin.AddPdfAsync(path);
        using var slow = new System.Net.Sockets.TcpClient(); await slow.ConnectAsync(IPAddress.Loopback, origin.Page.Port);
        await slow.GetStream().WriteAsync("GET /"u8.ToArray());
        await origin.DisposeAsync().AsTask().WaitAsync(TimeSpan.FromSeconds(2)); await origin.DisposeAsync();
        using var exclusive = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        await Assert.ThrowsAsync<ObjectDisposedException>(() => origin.AddPdfAsync(path));
    }
}
