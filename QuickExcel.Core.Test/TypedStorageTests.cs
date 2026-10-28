using System.Collections.Immutable;
using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Core.Test;

public sealed class TypedStorageTests : IDisposable
{
    private readonly string _folder = Path.Combine(Path.GetTempPath(), "QuickExcel-v2-" + Guid.NewGuid().ToString("N"));
    private readonly TypedWorkbookStorage _storage = new();
    private static readonly XNamespace Ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly byte[] Pixel = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aA1sAAAAASUVORK5CYII=");
    public TypedStorageTests() => Directory.CreateDirectory(_folder);
    private string PathFor(string name) => Path.Combine(_folder, name);
    private static CellSnapshot Literal(CellValue value, int style = 0) => new(CellContent.FromLiteral(value, style));
    private static CellSnapshot Formula(CellValue? value, CalculationState state = CalculationState.Stale, long revision = 0) =>
        new(CellContent.FromFormula("=1+1"), new CellCalculation(value, state, revision, 0, 0));
    private static WorkbookSnapshot Model(IEnumerable<CellSnapshot>? cells = null, ExcelDateSystem dates = ExcelDateSystem.Windows1900) =>
        new(Guid.NewGuid(), 0, dates, [new WorksheetSnapshot(SheetId.New(), "类型测试", (cells ?? []).Select((c, i) => new KeyValuePair<CellPosition, CellSnapshot>(new(0, i), c)))], "throw new Error('opening must not execute');");

    [Theory]
    [InlineData(ExcelDateSystem.Windows1900)]
    [InlineData(ExcelDateSystem.Mac1904)]
    public void EveryScalarAndCacheTypeRoundTripsWithStableIdentity(ExcelDateSystem dates)
    {
        CellValue[] values = [CellValue.Blank, CellValue.FromText(""), CellValue.FromText("00123"), CellValue.FromText("TRUE"), CellValue.FromText("#VALUE!"),
            CellValue.FromText("_x0041_\r\n中文\0"), CellValue.FromNumber(-3.125), CellValue.FromBoolean(false), CellValue.FromBoolean(true),
            .. Enum.GetValues<CellErrorCode>().Select(CellValue.FromError)];
        var snapshot = Model(values.Select(LiteralValue).Concat(values.Select(v => Formula(v))).Append(Formula(null)), dates);
        var path = PathFor("types.xlsq");
        _storage.Save(path, snapshot);
        var result = _storage.Open(path);
        Assert.Equal(snapshot.DocumentId, result.DocumentId);
        Assert.Equal(snapshot.Sheets[0].Id, result.Sheets[0].Id);
        Assert.Equal(dates, result.DateSystem);
        Assert.Equal(snapshot.ScriptSource, result.ScriptSource);
        Assert.Empty(result.FileState!.CompatibilityIssues);
        for (var i = 0; i < values.Length; i++)
        {
            Assert.Equal(values[i], result.Sheets[0].GetCell(0, i)!.Value);
            var cache = result.Sheets[0].GetCell(0, values.Length + i)!.Calculation!;
            Assert.Equal(values[i], cache.Value);
            Assert.Equal(CalculationState.Stale, cache.State);
        }
        Assert.Null(result.Sheets[0].GetCell(0, values.Length * 2)!.Calculation!.Value);
        var outer = ReadZip(File.ReadAllBytes(path));
        Assert.Equal(new[] { "manifest.json", "scripts/main.js", "state/document.json", "workbook.xlsx" }, outer.Keys.Order());
        var state = Encoding.UTF8.GetString(outer["state/document.json"]);
        Assert.DoesNotContain("formula", state.Replace("formula-functions", "binding"));
        var inner = ReadZip(outer["workbook.xlsx"]);
        var xml = XDocument.Parse(Encoding.UTF8.GetString(inner["xl/worksheets/sheet1.xml"]));
        var cells = xml.Descendants(Ns + "c").ToArray();
        Assert.Equal("inlineStr", (string?)cells[2].Attribute("t"));
        Assert.Equal("00123", cells[2].Element(Ns + "is")!.Value);
        Assert.Null(cells[6].Attribute("t"));
        Assert.Equal("b", (string?)cells[7].Attribute("t"));
        Assert.Equal("str", (string?)cells[values.Length + 4].Attribute("t"));
        static CellSnapshot LiteralValue(CellValue value) => Literal(value);
    }

    [Theory]
    [InlineData(ExcelDateSystem.Windows1900)]
    [InlineData(ExcelDateSystem.Mac1904)]
    public void DatesFormatsStylesAndFloatingImagesKeepRawSemantics(ExcelDateSystem dateSystem)
    {
        var cells = new Dictionary<CellPosition, CellSnapshot>();
        double[] numbers = [59, 60, 60.5, 61, -1.25, 0.75, 100000.125];
        for (var i = 0; i < numbers.Length; i++) cells[new(i, 0)] = Literal(CellValue.FromNumber(numbers[i]), 1);
        var sheet = new WorksheetSnapshot(SheetId.New(), "日期", cells,
            rows: [new KeyValuePair<int, RowMetadata>(0, new(31, true, 1))], columns: [new KeyValuePair<int, ColumnMetadata>(0, new(25, true, 1))],
            mergedCells: ["B1:C1"], images: [new SheetImage(Pixel, "png", 5, 4, 2, 3)], showGridLines: false, defaultRowHeight: 21, tabColor: "FF123456");
        var styles = new[] { new StyleDefinition(), new StyleDefinition { NumberFormatCode = "[h]:mm:ss", FontName = "Calibri", FontSize = 12,
            Bold = true, Italic = true, FontColor = "FF112233", FillColor = "FFEEFFEE", Horizontal = 2, WrapText = true, BottomBorder = 3, BottomBorderColor = "FF000000" } };
        var snapshot = new WorkbookSnapshot(Guid.NewGuid(), 42, dateSystem, [sheet], styles: styles);
        var path = PathFor("dates.xlsq");
        var previous = CultureInfo.CurrentCulture;
        try { CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("de-DE"); _storage.Save(path, snapshot); }
        finally { CultureInfo.CurrentCulture = previous; }
        var loaded = _storage.Open(path);
        Assert.Equal(42, loaded.Revision);
        Assert.Equal(styles, loaded.Styles);
        Assert.Equal(dateSystem, loaded.DateSystem);
        Assert.Equal(sheet.Rows, loaded.Sheets[0].Rows);
        Assert.Equal(sheet.Columns, loaded.Sheets[0].Columns);
        Assert.Equal(sheet.MergedCells.ToArray(), loaded.Sheets[0].MergedCells.ToArray());
        Assert.False(loaded.Sheets[0].ShowGridLines);
        Assert.Equal(Pixel, Assert.Single(loaded.Sheets[0].Images).Data.ToArray());
        for (var i = 0; i < numbers.Length; i++) Assert.Equal(numbers[i], loaded.Sheets[0].GetCell(i, 0)!.Value.NumberValue);
        Assert.Empty(loaded.FileState!.CompatibilityIssues);
        var xml = XDocument.Parse(Encoding.UTF8.GetString(ReadZip(ReadZip(File.ReadAllBytes(path))["workbook.xlsx"])["xl/styles.xml"]));
        Assert.Contains(xml.Descendants(Ns + "bottom"), e => (string?)e.Attribute("style") == "dashed");
    }

    [Fact]
    public void RemoteCellsRemainSparseAndDoNotHitTheLegacyDenseBudget()
    {
        var cells = Enumerable.Range(0, 1100).Select(row => new KeyValuePair<CellPosition, CellSnapshot>(new(row, 1023), Literal(CellValue.FromNumber(row))))
            .Append(new(new(99999, 1023), Literal(CellValue.FromText("last"))));
        var snapshot = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [new(SheetId.New(), "Sparse", cells)]);
        var path = PathFor("sparse.xlsq");
        _storage.Save(path, snapshot);
        var loaded = _storage.Open(path);
        Assert.Equal(1101, loaded.Sheets[0].Cells.Count);
        Assert.Equal("last", loaded.Sheets[0].GetCell(99999, 1023)!.Value.TextValue);
        Assert.True(new FileInfo(path).Length < 100_000);
    }

    [Fact]
    public void V1MigrationDoesNotGuessTypesExecuteScriptsOrOverwriteOriginal()
    {
        var oldPath = PathFor("old.xlsq");
        new LegacyWorkbookStorage().Save(oldPath, Model([Literal(CellValue.FromNumber(123)), Literal(CellValue.FromBoolean(true)), Formula(CellValue.FromNumber(2))]));
        var original = File.ReadAllBytes(oldPath);
        var migrated = _storage.Open(oldPath);
        Assert.Equal(CellValue.FromText("123"), migrated.Sheets[0].GetCell(0, 0)!.Value);
        Assert.Equal(CellValue.FromText("TRUE"), migrated.Sheets[0].GetCell(0, 1)!.Value);
        using var session = new DocumentSession(migrated);
        Assert.True(session.IsDirty);
        Assert.False(session.ScriptsEnabled);
        Assert.Throws<InvalidOperationException>(() => _storage.Save(oldPath, migrated));
        var newPath = PathFor("migrated.xlsq");
        _storage.Save(newPath, migrated);
        Assert.Equal(original, File.ReadAllBytes(oldPath));
        Assert.Equal(migrated.Sheets[0].Id, _storage.Open(newPath).Sheets[0].Id);
        Assert.Equal(2, _storage.Open(newPath).FileState!.SourceFormatVersion);
    }

    [Theory]
    [InlineData("formatVersion", "3")]
    [InlineData("formatVersion", "\"2\"")]
    [InlineData("minReaderVersion", "3")]
    [InlineData("scriptApiVersion", "3")]
    [InlineData("requiredFeatures", "[\"typed-workbook\",\"stable-sheet-ids\",\"script-actions\"]")]
    [InlineData("workbook", "{\"path\":\"../workbook.xlsx\",\"sha256\":\"bad\"}")]
    public void UnsupportedAndMalformedManifestFieldsAreRejected(string field, string json)
    {
        var path = PathFor("bad.xlsq"); _storage.Save(path, Model());
        EditJson(path, "manifest.json", root => root[field] = JsonNode.Parse(json));
        Assert.ThrowsAny<InvalidDataException>(() => _storage.Open(path));
    }

    [Theory]
    [InlineData("hash")]
    [InlineData("identity")]
    [InlineData("binding")]
    [InlineData("duplicate-property")]
    [InlineData("module-alias")]
    [InlineData("unlisted-entry")]
    public void PackageAmbiguityAndTamperingAreRejected(string scenario)
    {
        var path = PathFor("tamper.xlsq"); _storage.Save(path, Model([Literal(CellValue.FromNumber(1))]));
        var parts = ReadZip(File.ReadAllBytes(path));
        if (scenario == "hash") parts["scripts/main.js"] = Encoding.UTF8.GetBytes("// replaced");
        if (scenario == "identity" || scenario == "binding")
        {
            var state = JsonNode.Parse(parts["state/document.json"])!.AsObject();
            state["sheets"]![0]![scenario == "identity" ? "id" : "part"] = scenario == "identity" ? Guid.Empty.ToString() : "xl/worksheets/other.xml";
            parts["state/document.json"] = Encoding.UTF8.GetBytes(state.ToJsonString());
            Rehash(parts, "state", "state/document.json");
        }
        if (scenario == "duplicate-property") parts["manifest.json"] = Encoding.UTF8.GetBytes("{\"formatVersion\":2,\"formatVersion\":2}");
        if (scenario == "module-alias")
        {
            var manifest = JsonNode.Parse(parts["manifest.json"])!.AsObject();
            manifest["scripts"]!["modules"]!["alias"] = manifest["scripts"]!["modules"]!["main"]!.DeepClone();
            parts["manifest.json"] = Encoding.UTF8.GetBytes(manifest.ToJsonString());
        }
        if (scenario == "unlisted-entry") parts["scripts/hidden.js"] = Encoding.UTF8.GetBytes("throw 1");
        File.WriteAllBytes(path, Zip(parts));
        Assert.ThrowsAny<InvalidDataException>(() => _storage.Open(path));
    }

    [Fact]
    public void MultiModuleSourceAndBindingsRoundTripButCannotRunInLegacyRuntime()
    {
        var modules = new ScriptModuleSet(new Dictionary<string, ScriptModule> { ["main"] = new("scripts/main.js", "import 'price';"), ["price"] = new("scripts/lib/price.js", "export const rate = 0.2;") }, ["main"], 4, 2);
        var basic = Model();
        var snapshot = new WorkbookSnapshot(basic.DocumentId, 12, basic.DateSystem, basic.Sheets, modules.PrimarySource, basic.Styles, modules);
        var path = PathFor("modules.xlsq"); _storage.Save(path, snapshot);
        var result = _storage.Open(path);
        Assert.Equal(modules.Modules, result.ScriptModules.Modules);
        Assert.Equal(4, result.ScriptModules.Revision);
        using var document = new WorkbookDocument(result, new(_storage, new NoCalculation()));
        Assert.Throws<NotSupportedException>(document.EnableScriptsAndCalculate);
        Assert.False(document.ScriptsEnabled);
        var changed = result.WithScriptSource("// edited main").WithRevision(13);
        _storage.Save(path, changed);
        Assert.Equal(modules.Modules["price"], _storage.Open(path).ScriptModules.Modules["price"]);
    }

    [Theory]
    [InlineData("version")] [InlineData("source")] [InlineData("actions-feature")] [InlineData("actions-binding")]
    public void ModuleModeRejectsTamperingAndMismatchedActionCapability(string change)
    {
        var modules = new ScriptModuleSet(new Dictionary<string, ScriptModule>
        {
            ["main"] = new("scripts/main.js", "import 'price';"),
            ["price"] = new("scripts/lib/price.js", "throw new Error('opening never executes');")
        }, ["main"], scriptApiVersion: 2);
        var path = PathFor("module-mode.xlsq"); _storage.Save(path, Model().WithScriptModules(modules));
        Assert.Equal(2, _storage.Open(path).ScriptModules.ScriptApiVersion);
        var parts = ReadZip(File.ReadAllBytes(path));
        var manifest = JsonNode.Parse(parts["manifest.json"])!;
        if (change == "version") manifest["scriptApiVersion"] = 1;
        if (change == "source") parts["scripts/lib/price.js"] = Encoding.UTF8.GetBytes("// tampered helper");
        if (change == "actions-feature") manifest["requiredFeatures"]!.AsArray().Add("script-actions");
        parts["manifest.json"] = Encoding.UTF8.GetBytes(manifest.ToJsonString());
        if (change == "actions-binding")
        {
            var state = JsonNode.Parse(parts["state/document.json"])!;
            state["scriptBindings"]![0]!["kind"] = "actions";
            parts["state/document.json"] = Encoding.UTF8.GetBytes(state.ToJsonString());
            Rehash(parts, "state", "state/document.json");
        }
        File.WriteAllBytes(path, Zip(parts));
        Assert.Throws<InvalidDataException>(() => _storage.Open(path));
    }

    [Theory]
    [InlineData("feature")] [InlineData("api")] [InlineData("kind")] [InlineData("scope")]
    [InlineData("duplicate")] [InlineData("module")] [InlineData("removed-action")]
    public void ActionPackageRejectsInconsistentCapabilitiesBindingsAndDigest(string fault)
    {
        var path = PathFor("action-tamper.xlsq");
        _storage.Save(path, Model().WithScriptModules(M4ActionTests.Actions("throw new Error('never execute on open');", true)));
        var parts = ReadZip(File.ReadAllBytes(path));
        var manifest = JsonNode.Parse(parts["manifest.json"])!;
        if (fault == "feature") manifest["requiredFeatures"]!.AsArray().RemoveAt(3);
        if (fault == "api") manifest["scriptApiVersion"] = 1;
        parts["manifest.json"] = Encoding.UTF8.GetBytes(manifest.ToJsonString());
        if (fault is not ("feature" or "api"))
        {
            var state = JsonNode.Parse(parts["state/document.json"])!;
            var bindings = state["scriptBindings"]!.AsArray();
            if (fault == "kind") bindings[1]!["kind"] = "unknown";
            if (fault == "scope") bindings[1]!["scope"] = "worksheet";
            if (fault == "duplicate") bindings.Add(bindings[1]!.DeepClone());
            if (fault == "module") bindings[1]!["module"] = "missing";
            if (fault == "removed-action") bindings.RemoveAt(1);
            parts["state/document.json"] = Encoding.UTF8.GetBytes(state.ToJsonString());
            Rehash(parts, "state", "state/document.json");
        }
        File.WriteAllBytes(path, Zip(parts));
        Assert.Throws<InvalidDataException>(() => _storage.Open(path));
    }

    [Fact]
    public void TwoKindsOnOneModulePreserveTheirOwnOptionalMetadata()
    {
        var path = PathFor("action-metadata.xlsq");
        _storage.Save(path, Model().WithScriptModules(M4ActionTests.Actions("", true)));
        EditJson(path, "state/document.json", root =>
        {
            root["scriptBindings"]![0]!["note"] = "functions";
            root["scriptBindings"]![1]!["note"] = "actions";
        });
        var parts = ReadZip(File.ReadAllBytes(path));
        Rehash(parts, "state", "state/document.json"); File.WriteAllBytes(path, Zip(parts));
        _storage.Save(path, _storage.Open(path).WithScriptSource("// changed"));
        var state = JsonNode.Parse(ReadZip(File.ReadAllBytes(path))["state/document.json"])!;
        Assert.Equal("functions", state["scriptBindings"]![0]!["note"]!.GetValue<string>());
        Assert.Equal("actions", state["scriptBindings"]![1]!["note"]!.GetValue<string>());
    }

    [Fact]
    public void UnknownWorksheetContentIsPreservedUntilAnExplicitLossyCopy()
    {
        var path = PathFor("unknown.xlsx");
        _storage.Export(path, Model([Literal(CellValue.FromNumber(5))]));
        var parts = ReadZip(File.ReadAllBytes(path));
        var sheet = XDocument.Parse(Encoding.UTF8.GetString(parts["xl/worksheets/sheet1.xml"]));
        sheet.Root!.Add(new XElement(Ns + "dataValidations", new XAttribute("count", 1), new XElement(Ns + "dataValidation", new XAttribute("sqref", "A1"), new XElement(Ns + "formula1", "10"))));
        parts["xl/worksheets/sheet1.xml"] = Encoding.UTF8.GetBytes(sheet.ToString());
        File.WriteAllBytes(path, Zip(parts));
        var original = File.ReadAllBytes(path);
        var loaded = _storage.Open(path);
        Assert.Contains(loaded.FileState!.CompatibilityIssues, issue => issue.Contains("dataValidations"));
        var saved = PathFor("preserved.xlsq");
        _storage.Save(saved, loaded.WithScriptSource("// source only"));
        Assert.Equal(original, ReadZip(File.ReadAllBytes(saved))["workbook.xlsx"]);
        using var session = new DocumentSession(loaded);
        session.Execute(new SetCellsCommand([new(new(loaded.Sheets[0].Id, 0, 0), CellContent.FromLiteral(CellValue.FromNumber(8)))]));
        var before = File.ReadAllBytes(saved);
        Assert.Throws<WorkbookCompatibilityException>(() => _storage.Save(saved, session.Snapshot));
        Assert.Equal(before, File.ReadAllBytes(saved));
        _storage.Save(PathFor("subset.xlsq"), session.Snapshot, true);
        Assert.Empty(_storage.Open(PathFor("subset.xlsq")).FileState!.CompatibilityIssues);
        Assert.Equal(original, File.ReadAllBytes(path));
    }

    [Fact]
    public void ExportRequiresAvailableCachesAndExplicitStaleChoice()
    {
        var path = PathFor("export.xlsx"); File.WriteAllText(path, "previous");
        Assert.Throws<InvalidOperationException>(() => _storage.Export(path, Model([Formula(null)]), true, false));
        Assert.Equal("previous", File.ReadAllText(path));
        var source = Model([Formula(CellValue.FromText("00123")), Formula(CellValue.FromError(CellErrorCode.Script))]);
        Assert.Throws<InvalidOperationException>(() => _storage.Export(path, source));
        _storage.Export(path, source, true, false);
        var result = _storage.Open(path);
        Assert.Equal(CellValue.FromText("00123"), result.Sheets[0].GetCell(0, 0)!.Value);
        Assert.Equal(CellValue.FromError(CellErrorCode.Value), result.Sheets[0].GetCell(0, 1)!.Value);
        Assert.All(result.Sheets[0].Cells.Values, c => Assert.Equal(CellContentKind.Literal, c.Content.Kind));
        Assert.Equal(CellContentKind.Formula, source.Sheets[0].GetCell(0, 0)!.Content.Kind);
    }

    [Theory]
    [InlineData("../evil")]
    [InlineData("/absolute")]
    [InlineData("scripts\\evil")]
    [InlineData("scripts/%2e%2e/evil")]
    [InlineData("MANIFEST.JSON")]
    public void AmbiguousZipPathsAreRejectedBeforeParsing(string name)
    {
        var path = PathFor("paths.xlsq"); _storage.Save(path, Model());
        using (var zip = ZipFile.Open(path, ZipArchiveMode.Update)) using (var stream = zip.CreateEntry(name).Open()) stream.WriteByte(1);
        Assert.ThrowsAny<InvalidDataException>(() => _storage.Open(path));
    }

    [Theory]
    [InlineData("missing")]
    [InlineData("escape")]
    [InlineData("external")]
    [InlineData("cycle")]
    [InlineData("duplicate-id")]
    public void InvalidWorkbookRelationshipsNeverResolveToFilesOrNetwork(string mode)
    {
        var path = PathFor("relationship.xlsx"); _storage.Export(path, Model([Literal(CellValue.FromText("safe"))]));
        var parts = ReadZip(File.ReadAllBytes(path));
        var rels = XDocument.Parse(Encoding.UTF8.GetString(parts["xl/_rels/workbook.xml.rels"]));
        var relation = rels.Root!.Elements().Last();
        if (mode == "missing") relation.SetAttributeValue("Target", "worksheets/missing.xml");
        if (mode == "escape") relation.SetAttributeValue("Target", "../../outside.xml");
        if (mode == "external") { relation.SetAttributeValue("TargetMode", "External"); relation.SetAttributeValue("Target", "https://example.invalid/sheet.xml"); }
        if (mode == "cycle") relation.SetAttributeValue("Target", "workbook.xml");
        if (mode == "duplicate-id") rels.Root.Add(new XElement(relation));
        parts["xl/_rels/workbook.xml.rels"] = Encoding.UTF8.GetBytes(rels.ToString()); File.WriteAllBytes(path, Zip(parts));
        Assert.ThrowsAny<InvalidDataException>(() => _storage.Open(path));
    }

    [Theory]
    [InlineData("dtd")]
    [InlineData("deep")]
    [InlineData("nan")]
    [InlineData("duplicate-cell")]
    public void MalformedWorksheetXmlIsRejected(string kind)
    {
        var path = PathFor("xml.xlsx"); _storage.Export(path, Model([Literal(CellValue.FromNumber(2))]));
        var parts = ReadZip(File.ReadAllBytes(path));
        var text = Encoding.UTF8.GetString(parts["xl/worksheets/sheet1.xml"]);
        if (kind == "dtd") text = text.Replace("<worksheet", "<!DOCTYPE worksheet [<!ENTITY secret SYSTEM 'file:///C:/not-read'>]><worksheet");
        if (kind == "deep") text = text.Replace("</worksheet>", string.Concat(Enumerable.Repeat("<deep>", 70)) + string.Concat(Enumerable.Repeat("</deep>", 70)) + "</worksheet>");
        if (kind == "nan") text = text.Replace("<v>2</v>", "<v>NaN</v>");
        if (kind == "duplicate-cell") text = text.Replace("</row>", "<c r=\"A1\"><v>3</v></c></row>");
        parts["xl/worksheets/sheet1.xml"] = Encoding.UTF8.GetBytes(text); File.WriteAllBytes(path, Zip(parts));
        Assert.ThrowsAny<Exception>(() => _storage.Open(path));
    }

    [Fact]
    public void CancellationAndReplacementFailureKeepExistingFileAndCleanTemporaryFiles()
    {
        var path = PathFor("safe.xlsq"); var snapshot = Model([Literal(CellValue.FromNumber(1))]); _storage.Save(path, snapshot);
        var previous = File.ReadAllBytes(path);
        using var cts = new CancellationTokenSource(); cts.Cancel();
        Assert.Throws<OperationCanceledException>(() => _storage.Save(path, snapshot, false, cts.Token));
        Assert.Throws<OperationCanceledException>(() => _storage.Open(path, cts.Token));
        if (OperatingSystem.IsWindows())
        {
            using var held = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            Assert.ThrowsAny<IOException>(() => _storage.Save(path, snapshot.WithRevision(2)));
        }
        Assert.Equal(previous, File.ReadAllBytes(path));
        Assert.Empty(Directory.GetFiles(_folder, "*.tmp"));
    }

    [Fact]
    public void OptionalTopLevelMetadataSurvivesEditingWithoutGrantingTrust()
    {
        var path = PathFor("optional.xlsq"); _storage.Save(path, Model());
        EditJson(path, "manifest.json", root => root["vendorComment"] = "keep this");
        var snapshot = _storage.Open(path);
        _storage.Save(path, snapshot.WithScriptSource("// modified"));
        Assert.Equal("keep this", JsonNode.Parse(ReadZip(File.ReadAllBytes(path))["manifest.json"])!["vendorComment"]!.GetValue<string>());
        using var session = new DocumentSession(_storage.Open(path)); Assert.False(session.ScriptsEnabled);
    }


    [Fact]
    public void SharedFormulasExpandMixedReferencesWithoutChangingQuotedText()
    {
        var path = PathFor("shared.xlsx");
        _storage.Export(path, Model([Literal(CellValue.FromNumber(1))]));
        var parts = ReadZip(File.ReadAllBytes(path));
        var sheet = XDocument.Parse(Encoding.UTF8.GetString(parts["xl/worksheets/sheet1.xml"]));
        sheet.Root!.Element(Ns + "sheetData")!.ReplaceWith(new XElement(Ns + "sheetData",
            new XElement(Ns + "row", new XAttribute("r", 1), new XElement(Ns + "c", new XAttribute("r", "B1"),
                new XElement(Ns + "f", new XAttribute("t", "shared"), new XAttribute("si", 0), new XAttribute("ref", "B1:B2"), "A1+$A1+A$1+$A$1+IF(A1=1,\"A1\",'A1 sheet'!A1)+LOG10(A1)"), new XElement(Ns + "v", 7))),
            new XElement(Ns + "row", new XAttribute("r", 2), new XElement(Ns + "c", new XAttribute("r", "B2"),
                new XElement(Ns + "f", new XAttribute("t", "shared"), new XAttribute("si", 0)), new XElement(Ns + "v", 8)))));
        parts["xl/worksheets/sheet1.xml"] = Encoding.UTF8.GetBytes(sheet.ToString()); File.WriteAllBytes(path, Zip(parts));
        var loaded = _storage.Open(path);
        Assert.Equal("=A2+$A2+A$1+$A$1+IF(A2=1,\"A1\",'A1 sheet'!A2)+LOG10(A2)", loaded.Sheets[0].GetCell(1, 1)!.Content.FormulaText);
        Assert.Equal(8, loaded.Sheets[0].GetCell(1, 1)!.Value.NumberValue);
        using var session = new DocumentSession(loaded);
        session.Execute(new SetCellsCommand([new(new(loaded.Sheets[0].Id, 0, 0), CellContent.FromLiteral(CellValue.FromNumber(5)))]));
        _storage.Save(PathFor("expanded.xlsq"), session.Snapshot);
        Assert.Equal(loaded.Sheets[0].GetCell(1, 1)!.Content, _storage.Open(PathFor("expanded.xlsq")).Sheets[0].GetCell(1, 1)!.Content);
    }

    [Fact]
    public void NonstandardPartLocationsAndSharedStringTypesAreResolvedFromRelationships()
    {
        var path = PathFor("relocated.xlsx"); _storage.Export(path, Model());
        var parts = ReadZip(File.ReadAllBytes(path));
        var rels = XDocument.Parse(Encoding.UTF8.GetString(parts["xl/_rels/workbook.xml.rels"]));
        var relationshipNamespace = rels.Root!.Name.Namespace;
        rels.Root.Elements().Last().SetAttributeValue("Target", "../sheets/data.xml");
        rels.Root.Add(new XElement(relationshipNamespace + "Relationship", new XAttribute("Id", "strings"), new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"), new XAttribute("Target", "sharedStrings.xml")));
        parts["xl/_rels/workbook.xml.rels"] = Encoding.UTF8.GetBytes(rels.ToString());
        var sheet = XDocument.Parse(Encoding.UTF8.GetString(parts["xl/worksheets/sheet1.xml"]));
        sheet.Root!.Element(Ns + "sheetData")!.Add(new XElement(Ns + "row", new XAttribute("r", 9), new XElement(Ns + "c", new XAttribute("r", "D9"), new XAttribute("t", "s"), new XElement(Ns + "v", 0))));
        parts.Remove("xl/worksheets/sheet1.xml"); parts["sheets/data.xml"] = Encoding.UTF8.GetBytes(sheet.ToString());
        parts["xl/sharedStrings.xml"] = Encoding.UTF8.GetBytes(new XElement(Ns + "sst", new XElement(Ns + "si", new XElement(Ns + "t", "00123"))).ToString());
        var types = XDocument.Parse(Encoding.UTF8.GetString(parts["[Content_Types].xml"]));
        types.Root!.Elements().Single(e => (string?)e.Attribute("PartName") == "/xl/worksheets/sheet1.xml").SetAttributeValue("PartName", "/sheets/data.xml");
        types.Root.Add(new XElement(types.Root.Name.Namespace + "Override", new XAttribute("PartName", "/xl/sharedStrings.xml"), new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml")));
        parts["[Content_Types].xml"] = Encoding.UTF8.GetBytes(types.ToString()); File.WriteAllBytes(path, Zip(parts));
        var loaded = _storage.Open(path);
        Assert.Equal(CellValue.FromText("00123"), loaded.Sheets[0].GetCell(8, 3)!.Value);
        _storage.Save(PathFor("relocated.xlsq"), loaded);
        var reopened = _storage.Open(PathFor("relocated.xlsq"));
        Assert.Equal(loaded.Sheets[0].Id, reopened.Sheets[0].Id);
        Assert.Equal("sheets/data.xml", reopened.FileState!.SheetParts[0].Part);
    }

    [Fact]
    public void ModuleRevisionDoesNotGoBackwardsAcrossCellUndoRedo()
    {
        using var session = new DocumentSession(Model());
        var id = session.Snapshot.Sheets[0].Id;
        session.Execute(new SetCellsCommand([new(new(id, 0, 0), CellContent.FromLiteral(CellValue.FromNumber(1)))]));
        session.Execute(new UpdateScriptSourceCommand("// version one"));
        session.Execute(new SetCellsCommand([new(new(id, 0, 0), CellContent.FromLiteral(CellValue.FromNumber(2)))]));
        session.Execute(new UpdateScriptSourceCommand("// version two"));
        var revision = session.Snapshot.ScriptModules.Revision;
        Assert.True(session.Undo()); Assert.True(session.Undo());
        Assert.Equal(revision, session.Snapshot.ScriptModules.Revision);
        Assert.Equal("// version two", session.Snapshot.ScriptSource);
        Assert.True(session.Redo()); Assert.Equal(revision, session.Snapshot.ScriptModules.Revision);
        session.Execute(new UpdateScriptSourceCommand("// version three"));
        Assert.Equal(revision + 1, session.Snapshot.ScriptModules.Revision);
    }

    [Fact]
    public void NestedOptionalMetadataAndSheetAnnotationsSurviveRewrites()
    {
        var path = PathFor("annotations.xlsq"); _storage.Save(path, Model([Literal(CellValue.FromNumber(1))]));
        var parts = ReadZip(File.ReadAllBytes(path));
        var state = JsonNode.Parse(parts["state/document.json"])!;
        state["sheets"]![0]!["comment"] = "sheet note";
        state["calculation"]!["vendorNote"] = "cache note";
        parts["state/document.json"] = Encoding.UTF8.GetBytes(state.ToJsonString()); Rehash(parts, "state", "state/document.json");
        var manifest = JsonNode.Parse(parts["manifest.json"])!;
        manifest["scripts"]!["modules"]!["main"]!["description"] = "source note";
        parts["manifest.json"] = Encoding.UTF8.GetBytes(manifest.ToJsonString()); File.WriteAllBytes(path, Zip(parts));
        var loaded = _storage.Open(path); _storage.Save(path, loaded.WithScriptSource("// updated"));
        var after = ReadZip(File.ReadAllBytes(path));
        Assert.Equal("sheet note", JsonNode.Parse(after["state/document.json"])!["sheets"]![0]!["comment"]!.GetValue<string>());
        Assert.Equal("cache note", JsonNode.Parse(after["state/document.json"])!["calculation"]!["vendorNote"]!.GetValue<string>());
        Assert.Equal("source note", JsonNode.Parse(after["manifest.json"])!["scripts"]!["modules"]!["main"]!["description"]!.GetValue<string>());
    }

    [Theory]
    [InlineData("central-count")]
    [InlineData("entry-size")]
    [InlineData("manifest-size")]
    [InlineData("script-total")]
    public void CapacityChecksRejectBeforeEntryOrJsonMaterialization(string mode)
    {
        var path = PathFor("capacity.xlsq"); _storage.Save(path, Model());
        var bytes = File.ReadAllBytes(path);
        if (mode == "central-count")
            System.Buffers.Binary.BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(bytes.Length - 12, 2), 4000);
        else if (mode == "entry-size")
        {
            var offset = (int)System.Buffers.Binary.BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(bytes.Length - 6, 4));
            System.Buffers.Binary.BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(offset + 24, 4), 65 * 1024 * 1024);
        }
        else
        {
            var parts = ReadZip(bytes);
            if (mode == "manifest-size") parts["manifest.json"] = Encoding.UTF8.GetBytes(new string(' ', 64 * 1024 + 1));
            else { parts["scripts/main.js"] = new byte[600_000]; parts["scripts/extra.js"] = new byte[600_000]; }
            bytes = Zip(parts);
        }
        File.WriteAllBytes(path, bytes);
        Assert.Throws<InvalidDataException>(() => _storage.Open(path));
    }

    private void EditJson(string path, string part, Action<JsonObject> edit)
    {
        var parts = ReadZip(File.ReadAllBytes(path)); var root = JsonNode.Parse(parts[part])!.AsObject(); edit(root);
        parts[part] = Encoding.UTF8.GetBytes(root.ToJsonString()); File.WriteAllBytes(path, Zip(parts));
    }
    private static void Rehash(Dictionary<string, byte[]> parts, string field, string part)
    {
        var manifest = JsonNode.Parse(parts["manifest.json"])!.AsObject(); manifest[field]!["sha256"] = Convert.ToHexStringLower(SHA256.HashData(parts[part]));
        parts["manifest.json"] = Encoding.UTF8.GetBytes(manifest.ToJsonString());
    }
    private static Dictionary<string, byte[]> ReadZip(byte[] bytes)
    {
        using var memory = new MemoryStream(bytes); using var zip = new ZipArchive(memory);
        return zip.Entries.ToDictionary(e => e.FullName, e => { using var input = e.Open(); using var output = new MemoryStream(); input.CopyTo(output); return output.ToArray(); });
    }
    private static byte[] Zip(Dictionary<string, byte[]> parts)
    {
        using var memory = new MemoryStream();
        using (var zip = new ZipArchive(memory, ZipArchiveMode.Create, true))
            foreach (var (name, bytes) in parts) { using var entry = zip.CreateEntry(name).Open(); entry.Write(bytes); }
        return memory.ToArray();
    }
    private sealed class NoCalculation : IWorkbookCalculator
    { public CalculationResult Calculate(CalculationRequest request, CancellationToken cancellationToken = default) => throw new Exception("Calculation must not run."); }
    public void Dispose() => Directory.Delete(_folder, true);
}
