using System.Text.Json.Serialization;

namespace ExcelGeneration.Contracts;

/// <summary>
/// A data row: one primitive value per column, serialized as a positional JSON array (e.g. <c>[1,"abc",true,null]</c>).
/// </summary>
/// <remarks>
/// Rows are positional instead of objects so property names are not repeated for every row.
/// Values are interpreted on the server according to <see cref="ColumnDefinition.Type"/>.
/// </remarks>
[JsonConverter(typeof(ExcelRowJsonConverter))]
public readonly struct ExcelRow
{
    private readonly object?[]? _values;

    public ExcelRow(params object?[] values) => _values = values;

    /// <summary>The cell values. After deserialization each value is a <see cref="string"/>, <see cref="long"/>,
    /// <see cref="decimal"/>, <see cref="double"/>, <see cref="bool"/> or <c>null</c>.</summary>
    public IReadOnlyList<object?> Values => _values ?? [];

    public static implicit operator ExcelRow(object?[] values) => new(values);
}
