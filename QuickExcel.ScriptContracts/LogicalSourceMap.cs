namespace QuickExcel.ScriptContracts;

/// <summary>Bounded v3 VLQ decoder. Coordinates exposed to UI are one-based UTF-16 columns.</summary>
public sealed class LogicalSourceMap
{
    private readonly Dictionary<int, List<(int Column, int Line, int SourceColumn)>> _lines = new();
    public LogicalSourceMap(string mappings, string source, int nameCount)
    {
        var sourceLines = source.Split('\n');
        var line = 0; var column = 0; var originalLine = 0; var originalColumn = 0; var file = 0; var name = 0; var index = 0;
        int Read()
        {
            var value = 0; var shift = 0;
            while (index < mappings.Length && shift <= 25)
            {
                var c = mappings[index++];
                var digit = c switch { >= 'A' and <= 'Z' => c - 'A', >= 'a' and <= 'z' => c - 'a' + 26,
                    >= '0' and <= '9' => c - '0' + 52, '+' => 62, '/' => 63, _ => -1 };
                if (digit < 0) break;
                value |= (digit & 31) << shift;
                if ((digit & 32) == 0) return (value & 1) != 0 ? -(value >> 1) : value >> 1;
                shift += 5;
            }
            throw new ProtocolException("Invalid source map VLQ.");
        }
        var segments = 0;
        while (index < mappings.Length)
        {
            if (mappings[index] == ';') { index++; line++; column = 0; continue; }
            if (mappings[index] == ',') throw new ProtocolException("Empty source map segment.");
            if (++segments > 1_000_000 || line > 1_048_576) throw new ProtocolException("Source map segment budget exceeded.");
            var delta = Read();
            if (delta < 0 || column > int.MaxValue - delta) throw new ProtocolException("Source map columns must be ordered.");
            column += delta;
            if (!_lines.TryGetValue(line, out var points)) _lines[line] = points = [];
            if (index < mappings.Length && mappings[index] is not (',' or ';'))
            {
                file = checked(file + Read()); originalLine = checked(originalLine + Read()); originalColumn = checked(originalColumn + Read());
                if (file != 0 || originalLine < 0 || originalLine >= sourceLines.Length || originalColumn < 0 || originalColumn > sourceLines[originalLine].Length)
                    throw new ProtocolException("Source map points outside the original source.");
                if (index < mappings.Length && mappings[index] is not (',' or ';'))
                {
                    name = checked(name + Read());
                    if (name < 0 || name >= nameCount) throw new ProtocolException("Source map name is invalid.");
                }
                points.Add((column, originalLine, originalColumn));
            }
            else points.Add((column, -1, -1)); // Unmapped generated code must not inherit an earlier mapping.
            if (index < mappings.Length && mappings[index] == ',')
            {
                index++;
                if (index == mappings.Length || mappings[index] is ',' or ';') throw new ProtocolException("Missing source map segment.");
            }
            else if (index < mappings.Length && mappings[index] != ';') throw new ProtocolException("Unexpected source map field.");
        }
    }
    public (int Line, int Column)? OriginalPosition(int line, int column)
    {
        if (line < 1 || column < 1 || !_lines.TryGetValue(line - 1, out var points)) return null;
        (int Column, int Line, int SourceColumn)? chosen = null;
        foreach (var point in points) { if (point.Column > column - 1) break; chosen = point; }
        return chosen is { Line: >= 0 } value ? (value.Line + 1, value.SourceColumn + 1) : null;
    }
}
