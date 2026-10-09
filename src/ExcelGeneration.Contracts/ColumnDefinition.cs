namespace ExcelGeneration.Contracts;

/// <summary>Definition of a single column.</summary>
public sealed record ColumnDefinition
{
    public ColumnDefinition()
    {
    }

    [System.Diagnostics.CodeAnalysis.SetsRequiredMembers]
    public ColumnDefinition(string name, ColumnType type = ColumnType.String, string? format = null)
    {
        Name = name;
        Type = type;
        Format = format;
    }

    /// <summary>Column name, written in the header row.</summary>
    public required string Name { get; init; }

    /// <summary>Data type used to interpret the row values. Defaults to <see cref="ColumnType.String"/>.</summary>
    public ColumnType Type { get; init; } = ColumnType.String;

    /// <summary>Optional Excel number format, e.g. <c>#,##0.00</c> or <c>dd/mm/yyyy</c>.</summary>
    public string? Format { get; init; }

    /// <summary>Optional fixed column width (in Excel character units).</summary>
    public double? Width { get; init; }
}
