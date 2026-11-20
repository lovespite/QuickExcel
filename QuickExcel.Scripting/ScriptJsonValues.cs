using System.Collections;
using System.Text;
using System.Text.Json;

namespace QuickExcel.Scripting;

/// <summary>Writes the scalar, range and protocol values accepted by the native bridge without reflection.</summary>
public static class ScriptJsonValues
{
    public static string Serialize(object? value, int maxBytes = JscSession.MaxMessageBytes)
    {
        using var stream = Encode(value, maxBytes);
        return Encoding.UTF8.GetString(stream.GetBuffer(), 0, checked((int)stream.Length));
    }

    public static JsonElement ToElement(object? value, int maxBytes = JscSession.MaxMessageBytes)
    {
        using var stream = Encode(value, maxBytes);
        stream.Position = 0;
        using var document = JsonDocument.Parse(stream);
        return document.RootElement.Clone();
    }

    private static MemoryStream Encode(object? value, int maxBytes)
    {
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(maxBytes);
        var stream = new BoundedStream(maxBytes);
        try
        {
            using (var writer = new Utf8JsonWriter(stream, new JsonWriterOptions { MaxDepth = 64 }))
                Write(writer, value);
            return stream;
        }
        catch { stream.Dispose(); throw; }
    }

    private static void Write(Utf8JsonWriter writer, object? value)
    {
        switch (value)
        {
            case null: writer.WriteNullValue(); break;
            case JsonElement element: element.WriteTo(writer); break;
            case string text: writer.WriteStringValue(text); break;
            case char character: writer.WriteStringValue(character.ToString()); break;
            case bool boolean: writer.WriteBooleanValue(boolean); break;
            case Guid identifier: writer.WriteStringValue(identifier); break;
            case DateTime date: writer.WriteStringValue(date); break;
            case DateTimeOffset date: writer.WriteStringValue(date); break;
            case byte[] bytes: writer.WriteBase64StringValue(bytes); break;
            case byte number: writer.WriteNumberValue(number); break;
            case sbyte number: writer.WriteNumberValue(number); break;
            case short number: writer.WriteNumberValue(number); break;
            case ushort number: writer.WriteNumberValue(number); break;
            case int number: writer.WriteNumberValue(number); break;
            case uint number: writer.WriteNumberValue(number); break;
            case long number: writer.WriteNumberValue(number); break;
            case ulong number: writer.WriteNumberValue(number); break;
            case float number: writer.WriteNumberValue(number); break;
            case double number: writer.WriteNumberValue(number); break;
            case decimal number: writer.WriteNumberValue(number); break;
            case PrepareModulesOptions options:
                JsonSerializer.Serialize(writer, options, ScriptingJsonContext.Default.PrepareModulesOptions);
                break;
            case IDictionary dictionary:
                writer.WriteStartObject();
                foreach (DictionaryEntry entry in dictionary)
                {
                    if (entry.Key is not string key) throw new InvalidOperationException("JavaScript object keys must be strings.");
                    writer.WritePropertyName(key);
                    Write(writer, entry.Value);
                }
                writer.WriteEndObject();
                break;
            case IEnumerable sequence:
                writer.WriteStartArray();
                foreach (var item in sequence) Write(writer, item);
                writer.WriteEndArray();
                break;
            default:
                throw new InvalidOperationException("Unsupported JavaScript argument; use a scalar, array or explicit JSON value.");
        }
    }

    private sealed class BoundedStream(int maxBytes) : MemoryStream
    {
        private void Check(int count)
        {
            if (Position + count > maxBytes)
                throw new InvalidOperationException(maxBytes == JscSession.MaxMessageBytes
                    ? "JavaScript JSON exceeds the 16 MiB UTF-8 input limit."
                    : "JavaScript JSON exceeds its UTF-8 byte limit.");
        }

        public override void Write(byte[] buffer, int offset, int count) { Check(count); base.Write(buffer, offset, count); }
        public override void Write(ReadOnlySpan<byte> buffer) { Check(buffer.Length); base.Write(buffer); }
        public override void WriteByte(byte value) { Check(1); base.WriteByte(value); }
    }
}
