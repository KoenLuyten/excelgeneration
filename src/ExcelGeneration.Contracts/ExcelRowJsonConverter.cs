using System.Globalization;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace ExcelGeneration.Contracts;

/// <summary>
/// Reads and writes an <see cref="ExcelRow"/> as a JSON array of primitives.
/// </summary>
/// <remarks>
/// Reading yields plain CLR primitives instead of <see cref="JsonElement"/> per cell, which keeps memory low for
/// large payloads. Writing normalizes CLR types (dates, enums, guids, ...) to invariant JSON representations.
/// </remarks>
public sealed class ExcelRowJsonConverter : JsonConverter<ExcelRow>
{
    public override ExcelRow Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
    {
        if (reader.TokenType != JsonTokenType.StartArray)
        {
            throw new JsonException("A row must be a JSON array of cell values.");
        }

        var values = new List<object?>();
        while (reader.Read())
        {
            switch (reader.TokenType)
            {
                case JsonTokenType.EndArray:
                    return new ExcelRow(values.ToArray());
                case JsonTokenType.Null:
                    values.Add(null);
                    break;
                case JsonTokenType.True:
                    values.Add(true);
                    break;
                case JsonTokenType.False:
                    values.Add(false);
                    break;
                case JsonTokenType.String:
                    values.Add(reader.GetString());
                    break;
                case JsonTokenType.Number:
                    if (reader.TryGetInt64(out var l))
                        values.Add(l);
                    else if (reader.TryGetDecimal(out var m))
                        values.Add(m);
                    else
                        values.Add(reader.GetDouble());
                    break;
                default:
                    throw new JsonException($"Cell values must be primitives (string, number, boolean or null), found {reader.TokenType}.");
            }
        }

        throw new JsonException("Unexpected end of JSON while reading a row.");
    }

    public override void Write(Utf8JsonWriter writer, ExcelRow value, JsonSerializerOptions options)
    {
        writer.WriteStartArray();
        foreach (var cell in value.Values)
        {
            WriteCell(writer, cell);
        }
        writer.WriteEndArray();
    }

    private static void WriteCell(Utf8JsonWriter writer, object? value)
    {
        switch (value)
        {
            case null:
            case DBNull:
                writer.WriteNullValue();
                break;
            case string s:
                writer.WriteStringValue(s);
                break;
            case bool b:
                writer.WriteBooleanValue(b);
                break;
            case byte or sbyte or short or ushort or int or long:
                writer.WriteNumberValue(Convert.ToInt64(value, CultureInfo.InvariantCulture));
                break;
            case uint ui:
                writer.WriteNumberValue(ui);
                break;
            case ulong ul:
                writer.WriteNumberValue(ul);
                break;
            case decimal m:
                writer.WriteNumberValue(m);
                break;
            case double d:
                if (double.IsFinite(d)) writer.WriteNumberValue(d); else writer.WriteNullValue();
                break;
            case float f:
                if (float.IsFinite(f)) writer.WriteNumberValue(f); else writer.WriteNullValue();
                break;
            case DateTime dt:
                writer.WriteStringValue(dt.ToString("yyyy-MM-dd'T'HH:mm:ss.FFFFFFF", CultureInfo.InvariantCulture));
                break;
            case DateTimeOffset dto:
                // Excel has no time zones: keep the clock time as observed in the value's own offset.
                writer.WriteStringValue(dto.DateTime.ToString("yyyy-MM-dd'T'HH:mm:ss.FFFFFFF", CultureInfo.InvariantCulture));
                break;
            case DateOnly d:
                writer.WriteStringValue(d.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture));
                break;
            case TimeOnly t:
                writer.WriteStringValue(t.ToString("HH:mm:ss.FFFFFFF", CultureInfo.InvariantCulture));
                break;
            case TimeSpan ts:
                writer.WriteStringValue(ts.ToString("c", CultureInfo.InvariantCulture));
                break;
            case Uri uri:
                writer.WriteStringValue(uri.ToString());
                break;
            case Enum e:
                writer.WriteStringValue(e.ToString());
                break;
            case IFormattable formattable:
                writer.WriteStringValue(formattable.ToString(null, CultureInfo.InvariantCulture));
                break;
            default:
                writer.WriteStringValue(value.ToString());
                break;
        }
    }
}
