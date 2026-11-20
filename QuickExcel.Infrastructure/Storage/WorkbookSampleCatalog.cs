using System.Text.Json;
using QuickExcel.Domain;

namespace QuickExcel.Storage;

public sealed record WorkbookSample(string Id, string File, string Title, string TitleEn, string Description, string DescriptionEn, string Sha256);

/// <summary>Offline examples distributed alongside the application. Opening only reads a snapshot.</summary>
public sealed class WorkbookSampleCatalog
{
    private readonly string _directory;
    public IReadOnlyList<WorkbookSample> Samples { get; }

    public WorkbookSampleCatalog(string directory)
    {
        _directory=Path.GetFullPath(directory);
        using var document=JsonDocument.Parse(File.ReadAllText(Path.Combine(_directory,"catalog.json")));
        if(document.RootElement.GetProperty("schemaVersion").GetInt32()!=1)throw new InvalidDataException("Unsupported sample catalog.");
        var samples=new List<WorkbookSample>();
        foreach(var item in document.RootElement.GetProperty("samples").EnumerateArray())
        {
            string Read(string name) => item.GetProperty(name).GetString() ?? throw new InvalidDataException("Incomplete sample catalog.");
            var sample=new WorkbookSample(Read("id"),Read("file"),Read("title"),Read("titleEn"),Read("description"),Read("descriptionEn"),Read("sha256"));
            if(sample.Id.Length is <1 or >128 || sample.File.Length is <6 or >160 || !sample.File.EndsWith(".xlsq",StringComparison.Ordinal) ||
                sample.File.Any(c=>!char.IsAsciiLetterOrDigit(c)&&c is not '-' and not '_' and not '.') || sample.File.Contains("..",StringComparison.Ordinal) ||
                sample.Sha256.Length!=64 || sample.Sha256.Any(c=>!char.IsAsciiHexDigitUpper(c)) ||
                samples.Any(s=>s.Id==sample.Id||s.File.Equals(sample.File,StringComparison.OrdinalIgnoreCase)))
                throw new InvalidDataException("Invalid sample catalog entry.");
            samples.Add(sample);
        }
        if(samples.Count is <1 or >100)throw new InvalidDataException("Invalid sample catalog size.");
        Samples=samples.AsReadOnly();
    }

    public WorkbookSnapshot Read(string id)
    {
        var sample=Samples.SingleOrDefault(s=>s.Id==id) ?? throw new ArgumentException("Unknown workbook sample.",nameof(id));
        var path=Path.Combine(_directory,sample.File);
        var storage=new TypedWorkbookStorage();
        using var lease=storage.AcquireReadOnlyFile(path);
        if(storage.ReadVersion(path).Sha256!=sample.Sha256)throw new InvalidDataException("The sample workbook changed. Reinstall the matching sample files.");
        return storage.Open(path);
    }
}
