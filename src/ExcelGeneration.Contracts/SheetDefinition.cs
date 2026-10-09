namespace ExcelGeneration.Contracts;

/// <summary>A single worksheet (tab).</summary>
public sealed record SheetDefinition
{
    /// <summary>Sheet name. Invalid characters are replaced and the name is truncated to 31 characters.</summary>
    public required string Name { get; init; }

    /// <summary>Column definitions. Row values are matched to columns by position.</summary>
    public required IReadOnlyList<ColumnDefinition> Columns { get; init; }

    /// <summary>Data rows. Each row holds one value per column, in column order.</summary>
    public required IEnumerable<ExcelRow> Rows { get; init; }

    /// <summary>Write the column names as the first row. Defaults to <c>true</c>.</summary>
    public bool ShowHeader { get; init; } = true;

    /// <summary>Optional layout of the header row. When omitted a plain header row is written.</summary>
    public HeaderStyle? Header { get; init; }

    /// <summary>Add an autofilter on the header row. Requires <see cref="ShowHeader"/>.</summary>
    public bool AutoFilter { get; init; }

    /// <summary>Adjust column widths to their content (columns with an explicit width are left alone).</summary>
    public bool AutoFitColumns { get; init; } = true;
}
