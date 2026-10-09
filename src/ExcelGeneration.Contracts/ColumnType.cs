using System.Text.Json.Serialization;

namespace ExcelGeneration.Contracts;

/// <summary>Data type of a column. Serialized as a string.</summary>
[JsonConverter(typeof(JsonStringEnumConverter<ColumnType>))]
public enum ColumnType
{
    /// <summary>Text. Any primitive value is written as text.</summary>
    String,
    /// <summary>Whole number.</summary>
    Integer,
    /// <summary>Floating point / decimal number.</summary>
    Decimal,
    /// <summary><c>true</c>/<c>false</c>.</summary>
    Boolean,
    /// <summary>Date and time. ISO 8601 string (or OLE automation date number).</summary>
    DateTime,
    /// <summary>Date only. ISO 8601 string (<c>yyyy-MM-dd</c>).</summary>
    Date,
    /// <summary>Time of day / duration. <c>HH:mm:ss</c> or <c>d.HH:mm:ss</c>.</summary>
    Time,
    /// <summary>URL, written as a clickable hyperlink.</summary>
    Hyperlink,
    /// <summary>Excel A1 formula, with or without the leading <c>=</c>.</summary>
    Formula
}
