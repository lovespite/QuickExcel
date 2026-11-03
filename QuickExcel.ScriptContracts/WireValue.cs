using System.Text.Json;

namespace QuickExcel.ScriptContracts;

/// <summary>Explicit scalar tags; arrays/ranges contain only scalar values. Errors are never inferred from text.</summary>
public sealed record WireValue(string Kind, JsonElement Value)
{
    public static WireValue FromObject(object? value) => FromJson(WorkerProtocol.EncodePayload(value));

    private static WireValue FromJson(JsonElement value) => value.ValueKind switch
    {
        JsonValueKind.Null => new("blank", value),
        JsonValueKind.String => new("text", value),
        JsonValueKind.Number when value.TryGetDouble(out var number) && double.IsFinite(number) => new("number", value),
        JsonValueKind.True or JsonValueKind.False => new("boolean", value),
        JsonValueKind.Array => Collection("array", value),
        JsonValueKind.Object => Collection("range", value),
        _ => throw new ProtocolException("Unsupported script value.")
    };

    private static WireValue Collection(string kind, JsonElement value)
    {
        var result = new WireValue(kind, value);
        result.Validate();
        return result;
    }

    public void Validate(bool scalarOnly = false)
    {
        var valid = Kind switch
        {
            "blank" => Value.ValueKind == JsonValueKind.Null,
            "text" => Value.ValueKind == JsonValueKind.String,
            "boolean" => Value.ValueKind is JsonValueKind.True or JsonValueKind.False,
            "number" => Value.ValueKind == JsonValueKind.Number && Value.TryGetDouble(out var number) && double.IsFinite(number),
            "array" when !scalarOnly => ValidateArray(Value),
            "range" when !scalarOnly => ValidateRange(Value),
            _ => false
        };
        if (!valid) throw new ProtocolException("Invalid typed script value.");
    }

    private static bool ValidateArray(JsonElement value)
    {
        if (value.ValueKind != JsonValueKind.Array || value.GetArrayLength() > 1_000_000) return false;
        foreach (var item in value.EnumerateArray())
        {
            if (item.ValueKind is not (JsonValueKind.Null or JsonValueKind.String or JsonValueKind.Number or JsonValueKind.True or JsonValueKind.False)) return false;
            FromJson(item).Validate(scalarOnly: true);
        }
        return true;
    }

    private static bool ValidateRange(JsonElement value) => value.ValueKind == JsonValueKind.Object &&
        value.EnumerateObject().Count() == 4 && value.TryGetProperty("kind", out var kind) && kind.ValueKind == JsonValueKind.String && kind.GetString() == "range" &&
        value.TryGetProperty("rows", out var rows) && rows.ValueKind == JsonValueKind.Number && rows.TryGetInt32(out var r) && r >= 0 &&
        value.TryGetProperty("columns", out var columns) && columns.ValueKind == JsonValueKind.Number && columns.TryGetInt32(out var c) && c >= 0 &&
        value.TryGetProperty("values", out var values) && ValidateArray(values) && (long)r * c == values.GetArrayLength();

    public object? ToObject(bool scalarOnly = false)
    {
        Validate(scalarOnly);
        return Kind switch
        {
            "blank" => null, "text" => Value.GetString(), "number" => Value.GetDouble(), "boolean" => Value.GetBoolean(),
            // JsonElement preserves the exact existing v1 array / v2 range bridge shape in JscSession.
            _ => Value
        };
    }
}
