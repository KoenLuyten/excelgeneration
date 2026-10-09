using System.Data;
using ExcelGeneration.Contracts;

namespace ExcelGeneration.Client.Mapping;

/// <summary>Maps a <see cref="DataTable"/> to a <see cref="SheetDefinition"/>.</summary>
internal static class DataTableSheetMapper
{
    public static SheetDefinition Map(string name, DataTable table, SheetOptions options)
    {
        ArgumentNullException.ThrowIfNull(table);

        var columns = table.Columns
            .Cast<DataColumn>()
            .Where(c => !options.Ignored.Contains(c.ColumnName))
            .Select(c =>
            {
                var o = options.GetColumn(c.ColumnName);
                return (Source: c, o?.Order, Definition: new ColumnDefinition
                {
                    Name = o?.Name ?? (string.IsNullOrEmpty(c.Caption) ? c.ColumnName : c.Caption),
                    Type = o?.Type ?? ClrTypeMap.Map(c.DataType),
                    Format = o?.Format,
                    Width = o?.Width
                });
            })
            .OrderBy(c => c.Order ?? int.MaxValue)
            .ThenBy(c => c.Source.Ordinal)
            .ToArray();

        if (columns.Length == 0)
            throw new InvalidOperationException($"DataTable '{table.TableName}' has no columns to map.");

        var ordinals = columns.Select(c => c.Source.Ordinal).ToArray();

        return new SheetDefinition
        {
            Name = name,
            Columns = columns.Select(c => c.Definition).ToArray(),
            Rows = table.Rows.Cast<DataRow>()
                .Where(r => r.RowState != DataRowState.Deleted)
                .Select(r => ToRow(r, ordinals)),
            ShowHeader = options.ShowHeader,
            Header = options.Header,
            AutoFilter = options.AutoFilter,
            AutoFitColumns = options.AutoFitColumns
        };
    }

    private static ExcelRow ToRow(DataRow row, int[] ordinals)
    {
        var values = new object?[ordinals.Length];
        for (var i = 0; i < ordinals.Length; i++)
        {
            var value = row[ordinals[i]];
            values[i] = value is DBNull ? null : value;
        }
        return new ExcelRow(values);
    }
}
