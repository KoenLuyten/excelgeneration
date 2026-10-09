using System.Data;
using ExcelGeneration.Client.Mapping;
using ExcelGeneration.Contracts;

namespace ExcelGeneration.Client;

/// <summary>
/// Fluent builder that turns collections and <see cref="DataTable"/>s into a <see cref="GenerateExcelRequest"/>.
/// </summary>
/// <example>
/// <code>
/// var workbook = new WorkbookBuilder("orders.xlsx")
///     .AddSheet("Orders", orders, s =>
///     {
///         s.AutoFilter = true;
///         s.Header = new HeaderStyle { BackgroundColor = "#1F4E78", FontColor = "#FFFFFF" };
///         s.Column(o => o.Amount, c => c.Format = "#,##0.00");
///     })
///     .AddSheet(customersTable);
/// var file = await client.GenerateAsync(workbook);
/// </code>
/// </example>
public sealed class WorkbookBuilder(string? fileName = null)
{
    private readonly List<SheetDefinition> _sheets = [];

    /// <summary>File name of the generated workbook.</summary>
    public string? FileName { get; set; } = fileName;

    /// <summary>Adds a sheet with one row per item and one column per public property.</summary>
    public WorkbookBuilder AddSheet<T>(string name, IEnumerable<T> items, Action<SheetOptions<T>>? configure = null)
    {
        var options = new SheetOptions<T>();
        configure?.Invoke(options);
        _sheets.Add(EnumerableSheetMapper.Map(name, items, options));
        return this;
    }

    /// <summary>Adds a sheet named after <see cref="DataTable.TableName"/>.</summary>
    public WorkbookBuilder AddSheet(DataTable table, Action<SheetOptions>? configure = null)
    {
        ArgumentNullException.ThrowIfNull(table);
        var name = string.IsNullOrWhiteSpace(table.TableName) ? $"Sheet{_sheets.Count + 1}" : table.TableName;
        return AddSheet(name, table, configure);
    }

    /// <summary>Adds a sheet from a <see cref="DataTable"/>.</summary>
    public WorkbookBuilder AddSheet(string name, DataTable table, Action<SheetOptions>? configure = null)
    {
        var options = new SheetOptions();
        configure?.Invoke(options);
        _sheets.Add(DataTableSheetMapper.Map(name, table, options));
        return this;
    }

    /// <summary>Adds one sheet per table of the <see cref="DataSet"/>.</summary>
    public WorkbookBuilder AddSheets(DataSet dataSet, Action<DataTable, SheetOptions>? configure = null)
    {
        ArgumentNullException.ThrowIfNull(dataSet);
        foreach (DataTable table in dataSet.Tables)
        {
            AddSheet(table, configure is null ? null : o => configure(table, o));
        }
        return this;
    }

    /// <summary>Adds a sheet that was defined directly against the contract.</summary>
    public WorkbookBuilder AddSheet(SheetDefinition sheet)
    {
        ArgumentNullException.ThrowIfNull(sheet);
        _sheets.Add(sheet);
        return this;
    }

    public GenerateExcelRequest Build()
    {
        if (_sheets.Count == 0)
            throw new InvalidOperationException("Add at least one sheet.");

        return new GenerateExcelRequest { FileName = FileName, Sheets = _sheets.ToArray() };
    }
}
