using System.Collections.Immutable;
using System.Text;

namespace QuickExcel.Core;

public enum ScriptConsoleLevel { Input, Result, Log, Info, Warning, Error, Debug }

public sealed record ScriptConsoleEntry(long Sequence, DateTimeOffset Time, string Kind, string Text)
{
    // Keep the existing bounded worker log wire format; presentation consumes its level prefixes.
    public ScriptConsoleLevel Level { get; init; } = Describe(Kind, Text).Level;
    public string DisplayText => Describe(Kind, Text).Text;

    private static (ScriptConsoleLevel Level, string Text) Describe(string kind, string text)
    {
        if (kind == "input") return (ScriptConsoleLevel.Input, text);
        const string calculationPrefix = "JavaScript：";
        var context = text.StartsWith(calculationPrefix, StringComparison.Ordinal) ? calculationPrefix : "";
        var message = text.AsSpan(context.Length);
        var (level, length) = message switch
        {
            _ when message.StartsWith("[warn] ", StringComparison.Ordinal) => (ScriptConsoleLevel.Warning, 7),
            _ when message.StartsWith("[error] ", StringComparison.Ordinal) => (ScriptConsoleLevel.Error, 8),
            _ when message.StartsWith("[info] ", StringComparison.Ordinal) => (ScriptConsoleLevel.Info, 7),
            _ when message.StartsWith("[debug] ", StringComparison.Ordinal) => (ScriptConsoleLevel.Debug, 8),
            _ when message.StartsWith("=> ", StringComparison.Ordinal) && context.Length == 0 => (ScriptConsoleLevel.Result, 3),
            _ => (ScriptConsoleLevel.Log, 0)
        };
        return (level, length == 0 ? text : context + message[length..].ToString());
    }
}

public sealed partial class DocumentSession
{
    public const int MaxConsoleSourceBytes = 64 * 1024;
    public const int MaxConsoleEntries = 500;
    private readonly Queue<ScriptConsoleEntry> _consoleEntries = new();
    private long _consoleSequence;
    private string? _activeConsoleSource;
    public ImmutableArray<ScriptConsoleEntry> ConsoleEntries { get { lock (_gate) return _consoleEntries.ToImmutableArray(); } }
    public void ClearConsole() { lock (_gate) { ThrowIfClosed(); _consoleEntries.Clear(); } }
    private void AppendConsole(IEnumerable<string> messages, string kind = "output", ScriptConsoleLevel? level = null)
    {
        foreach (var message in messages.Take(100))
        {
            if (message is null) continue;
            if (_consoleEntries.Count == MaxConsoleEntries) _consoleEntries.Dequeue();
            var entry = new ScriptConsoleEntry(++_consoleSequence, DateTimeOffset.UtcNow, kind, message.Length > 4000 ? message[..4000] : message);
            _consoleEntries.Enqueue(level is { } severity ? entry with { Level = severity } : entry);
        }
    }
    /// <summary>One user-entered command, never a persistent grant to workbook scripts.</summary>
    public ScriptActionRequest BeginConsole(string source)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(source);
        if (Encoding.UTF8.GetByteCount(source) > MaxConsoleSourceBytes) throw new ArgumentException("控制台代码最多 64 KiB。");
        return BeginActionCore("__console__", null, source);
    }
}
