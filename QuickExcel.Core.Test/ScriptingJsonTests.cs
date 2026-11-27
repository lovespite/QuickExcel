using System.Text;
using System.Text.Json;
using QuickExcel.Scripting;

namespace QuickExcel.Core.Test;

public sealed class ScriptingJsonTests
{
    [Fact]
    public void ExplicitBridgeValuesPreserveJsonBytesAndRangeShape()
    {
        using var source = JsonDocument.Parse("{\"name\":\"中文 🧮\",\"n\":1.25,\"empty\":null}");
        object?[] values = [null, true, false, "中文 🧮 <>&\\\"", 12, 1.25d, 123.456m,
            Guid.Parse("00112233-4455-6677-8899-aabbccddeeff"), new byte[] { 0, 127, 255 }, source.RootElement,
            new Dictionary<string, object?> { ["kind"] = "range", ["rows"] = 1, ["columns"] = 3,
                ["values"] = new List<object?> { 1d, "x", null } }];

        Assert.Equal(JsonSerializer.Serialize(values), ScriptJsonValues.Serialize(values));
        Assert.Equal(JsonSerializer.Serialize(values), ScriptJsonValues.ToElement(values).GetRawText());
    }

    [Fact]
    public void Utf8BudgetIncludesEscapingAndArrayDelimiters()
    {
        object?[] values = ["中", "\""];
        var expected = JsonSerializer.Serialize(values);
        var length = Encoding.UTF8.GetByteCount(expected);
        Assert.Equal(expected, ScriptJsonValues.Serialize(values, length));
        Assert.Throws<InvalidOperationException>(() => ScriptJsonValues.Serialize(values, length - 1));
        Assert.Throws<InvalidOperationException>(() => ScriptJsonValues.ToElement(values, length - 1));
    }

    [Fact]
    public void OversizedRangeResultStopsBeforeSerializingTheRemainingCells()
    {
        var cell = new string('a', 32_767);
        var visited = 0;
        var values = Enumerable.Range(0, 1_024).Select(_ => { visited++; return cell; });
        Assert.Throws<ScriptMessageLimitException>(() => ScriptJsonValues.ToElement(values));
        Assert.InRange(visited, 1, 1_023);
    }

    [Fact]
    public void UnknownObjectsCannotFallBackToReflection()
    {
        Assert.Throws<InvalidOperationException>(() => ScriptJsonValues.Serialize(new { hidden = 1 }));
        Assert.Throws<InvalidOperationException>(() => ScriptJsonValues.ToElement(new { hidden = 1 }));
    }
}
