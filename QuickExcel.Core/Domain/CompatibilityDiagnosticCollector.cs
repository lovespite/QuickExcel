using System.Globalization;

namespace QuickExcel.Domain;

/// <summary>One inclusive row interval in one imported/current worksheet. Row numbers are one-based.</summary>
public sealed record CompatibilityRowSpan(string Sheet, int FirstRow, int LastRow);

/// <summary>Exact row coverage, independent of the bounded diagnostic location samples.</summary>
public sealed class CompatibilityRowCoverage : IEquatable<CompatibilityRowCoverage>
{
    public CompatibilityRowCoverage(IEnumerable<CompatibilityRowSpan> spans) => Spans = Array.AsReadOnly(spans.ToArray());
    public IReadOnlyList<CompatibilityRowSpan> Spans { get; }
    public long Count => Spans.Sum(span => (long)span.LastRow - span.FirstRow + 1);
    public bool Equals(CompatibilityRowCoverage? other) => other is not null && Spans.SequenceEqual(other.Spans);
    public override bool Equals(object? obj) => obj is CompatibilityRowCoverage other && Equals(other);
    public override int GetHashCode()
    {
        var hash = new HashCode();
        foreach (var span in Spans) hash.Add(span);
        return hash.ToHashCode();
    }
    public static CompatibilityRowCoverage Union(IEnumerable<CompatibilityRowCoverage> coverages)
    {
        var rows = new CompatibilityRows(new());
        foreach (var coverage in coverages) rows.Add(coverage);
        return rows.Snapshot();
    }
}

public sealed record CompatibilityDiagnosticSummary(CompatibilityRowCoverage Rows, bool HasNonRowLocations,
    bool LocationsTruncated);

/// <summary>
/// Retains every semantic feature and its exact row coverage while bounding detailed location samples.
/// Resource limits reject an inspection rather than silently presenting incomplete type/row totals.
/// </summary>
public sealed class CompatibilityDiagnosticCollector(int maximumSamples = 1024, int maximumTypes = 1024)
{
    private sealed class Feature(CompatibilityDiagnostic first, CompatibilityRowBudget budget)
    {
        public CompatibilityDiagnostic First { get; } = first;
        public CompatibilityRows Rows { get; } = new(budget);
        public int FirstSample { get; set; } = -1;
        public bool HasNonRowLocations { get; set; }
        public bool LocationsTruncated { get; set; }
    }

    private readonly Dictionary<(CompatibilityKind Kind, string Code, string Feature, string? Namespace), Feature> _features = [];
    private readonly CompatibilityRowBudget _budget = new();
    private readonly List<CompatibilityDiagnostic> _samples = [];
    private readonly HashSet<CompatibilityDiagnostic> _sampleSet = [];
    private readonly List<CompatibilityDiagnostic> _notices = [];

    public void Add(CompatibilityDiagnostic diagnostic)
    {
        if (diagnostic.IsReportNotice)
        {
            if (!_notices.Contains(diagnostic) && _notices.Count < 32) _notices.Add(diagnostic);
            return;
        }
        var key = (diagnostic.Kind, diagnostic.TypeCode, diagnostic.TypeFeature, diagnostic.TypeNamespace);
        if (!_features.TryGetValue(key, out var feature))
        {
            if (_features.Count >= maximumTypes) throw new InvalidDataException("兼容性问题类型超过报告容量，无法可靠统计此文件。");
            _features.Add(key, feature = new(diagnostic with { Summary = null }, _budget));
        }
        if (diagnostic.Summary is { } summary)
        {
            feature.Rows.Add(summary.Rows);
            feature.HasNonRowLocations |= summary.HasNonRowLocations;
            feature.LocationsTruncated |= summary.LocationsTruncated;
        }
        else feature.HasNonRowLocations |= !feature.Rows.Add(diagnostic);
        var sample = diagnostic with { Summary = null };
        if (_sampleSet.Contains(sample)) return;
        if (_samples.Count < maximumSamples)
        {
            _sampleSet.Add(sample);
            if (feature.FirstSample < 0) feature.FirstSample = _samples.Count;
            _samples.Add(sample);
        }
        else feature.LocationsTruncated = true;
    }

    public IReadOnlyList<CompatibilityDiagnostic> Diagnostics
    {
        get
        {
            var result = new List<CompatibilityDiagnostic>(_samples);
            foreach (var feature in _features.Values)
            {
                var summary = new CompatibilityDiagnosticSummary(feature.Rows.Snapshot(), feature.HasNonRowLocations,
                    feature.LocationsTruncated);
                if (feature.FirstSample >= 0) result[feature.FirstSample] = result[feature.FirstSample] with { Summary = summary };
                else result.Add(feature.First with { Part = null, SheetName = null, Range = null, Summary = summary });
            }
            result.AddRange(_notices);
            if (_features.Values.Any(feature => feature.LocationsTruncated) && !result.Any(d => d.Code == "report.locations-truncated"))
                result.Add(new("report.locations-truncated", CompatibilityKind.PreservedOnly, "报告容量",
                    "详细位置仅保留前 1024 条样本；问题类型和受影响行数仍完整统计，保存保护覆盖全部未支持内容。"));
            return result;
        }
    }

    /// <summary>Creates complete row coverage from either aggregated diagnostics or an unaggregated report.</summary>
    public static CompatibilityDiagnosticSummary Summarize(IEnumerable<CompatibilityDiagnostic> diagnostics)
    {
        var rows = new CompatibilityRows(new());
        var nonRow = false;
        var truncated = false;
        foreach (var diagnostic in diagnostics)
        {
            if (diagnostic.Summary is { } summary)
            {
                rows.Add(summary.Rows);
                nonRow |= summary.HasNonRowLocations;
                truncated |= summary.LocationsTruncated;
            }
            else nonRow |= !rows.Add(diagnostic);
        }
        return new(rows.Snapshot(), nonRow, truncated);
    }
}

internal sealed class CompatibilityRowBudget
{
    // An interval is compact even for an entire range; adversarial fragmented ranges have a fixed bound.
    private int _intervals;
    public void Change(int delta)
    {
        if (delta > 262144 - _intervals) throw new InvalidDataException("兼容性受影响行统计超过容量，无法可靠统计此文件。");
        _intervals += delta;
    }
}

internal sealed class CompatibilityRows(CompatibilityRowBudget budget)
{
    private readonly Dictionary<string, SortedSet<(int First, int Last)>> _sheets = new(StringComparer.Ordinal);

    public bool Add(CompatibilityDiagnostic diagnostic)
    {
        var sheet = diagnostic.SheetName;
        if (sheet is null && diagnostic.Part?.StartsWith("xl/worksheets/", StringComparison.Ordinal) == true) sheet = diagnostic.Part;
        if (sheet is null || diagnostic.Range is not { Length: > 0 } range) return false;
        var found = false;
        var complete = true;
        foreach (var area in range.Split([' ', ','], StringSplitOptions.RemoveEmptyEntries))
        {
            if (area.StartsWith("row:", StringComparison.Ordinal))
            {
                if (Row(area.AsSpan(4), out var row)) { Add(sheet, row, row); found = true; }
                else complete = false;
                continue;
            }
            var endpoints = area.Split(':');
            if (endpoints.Length is < 1 or > 2 || !CellRow(endpoints[0], out var first) ||
                !CellRow(endpoints[^1], out var last)) { complete = false; continue; }
            Add(sheet, Math.Min(first, last), Math.Max(first, last)); found = true;
        }
        return found && complete;
    }

    public void Add(CompatibilityRowCoverage coverage)
    {
        foreach (var span in coverage.Spans) Add(span.Sheet, span.FirstRow, span.LastRow);
    }

    private void Add(string sheet, int first, int last)
    {
        if (first < 1 || last < first || last > WorkbookLimits.MaxRows)
            throw new InvalidDataException("兼容性诊断行范围无效。");
        if (!_sheets.TryGetValue(sheet, out var intervals))
        {
            // A current report can contain both imported names and renamed current sheets.
            if (_sheets.Count >= WorkbookLimits.MaxSheets * 2) throw new InvalidDataException("兼容性诊断工作表数量超过限制。");
            _sheets.Add(sheet, intervals = []);
        }
        var previous = intervals.GetViewBetween((0, 0), (first, int.MaxValue)).Max;
        if (previous.First > 0 && previous.Last >= first - 1)
        {
            first = previous.First; last = Math.Max(last, previous.Last);
            intervals.Remove(previous); budget.Change(-1);
        }
        var overlapping = intervals.GetViewBetween((first, 0), (last + 1, int.MaxValue)).ToArray();
        foreach (var current in overlapping)
        {
            last = Math.Max(last, current.Last);
            intervals.Remove(current); budget.Change(-1);
        }
        budget.Change(1); intervals.Add((first, last));
    }

    public CompatibilityRowCoverage Snapshot() => new(_sheets.OrderBy(pair => pair.Key, StringComparer.Ordinal)
        .SelectMany(pair => pair.Value.Select(span => new CompatibilityRowSpan(pair.Key, span.First, span.Last))));

    private static bool CellRow(string reference, out int row)
    {
        var text = reference.AsSpan();
        if (!text.IsEmpty && text[0] == '$') text = text[1..];
        var letters = 0;
        while (letters < text.Length && (text[letters] is >= 'A' and <= 'Z' or >= 'a' and <= 'z')) letters++;
        if (letters is < 1 or > 3) { row = 0; return false; }
        text = text[letters..];
        if (!text.IsEmpty && text[0] == '$') text = text[1..];
        return Row(text, out row);
    }

    private static bool Row(ReadOnlySpan<char> text, out int row) =>
        int.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out row) && row is >= 1 and <= WorkbookLimits.MaxRows;
}
